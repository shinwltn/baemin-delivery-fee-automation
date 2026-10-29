"""배달의민족 외식배달비 자동조사 엔진.

2026-10-28에 전달받은 1028배민조사자동화.ipynb의 핵심 함수 정의를
가능한 한 원문 그대로 옮겼습니다. BlueStacks와 실조사 실행은 검증되지 않았습니다.
"""

import ast
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
import xml.etree.ElementTree as ET
from collections import defaultdict
from contextlib import redirect_stdout
from copy import copy
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from pprint import pprint
from shutil import copy2
import uiautomator2 as u2
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string


def get_month_cols(month):
    start = (column_index_from_string("BC") + (month - 8) * 6)
    return {
        "store_min": get_column_letter(start),
        "store_fee": get_column_letter(start + 1),
        "saver_min": get_column_letter(start + 2),
        "saver_fee": get_column_letter(start + 3),
        "one_min": get_column_letter(start + 4),
        "one_fee": get_column_letter(start + 5)
    }

def get_result_cols(month):
    start = (column_index_from_string("BC") + (month - 8) * 6 + 6)
    return {
        "change": get_column_letter(start),
        "store_status": get_column_letter(start + 1),
        "saver_status": get_column_letter(start + 2),
        "one_status": get_column_letter(start + 3),
        "notes": get_column_letter(start + 4)
    }


def go_home():
    package = d.app_current()["package"]
    for _ in range(4):
        if "현재 위치를 변경할 수 있습니다"in d.dump_hierarchy():
            return True
        d.press("back")
        time.sleep(1)
        if d.app_current()["package"] != package:
            d.app_start(package)
            time.sleep(3)
    return "현재 위치를 변경할 수 있습니다." in d.dump_hierarchy()


def set_address(address):

    # 1. 현재 주소 선택
    current_address = d.xpath('//*[contains(@content-desc,"현재 위치를 변경할 수 있습니다")]')
    if not current_address.wait(timeout=5):
        print("현재 주소 버튼 못 찾음")
        return False
    current_address.click()
    address_page = d.xpath('//*[contains(@text,"도로명") or contains(@content-desc,"도로명")]')
    if not address_page.wait(timeout=5):
        print("주소 설정 화면 진입 실패")
        return False

    # 2. 주소 검색창 클릭
    root = ET.fromstring(d.dump_hierarchy())
    for node in root.iter():
        text = node.attrib.get("text", "")
        desc = node.attrib.get("content-desc", "")
        if "도로명" in text or "도로명" in desc:
            x1, y1, x2, y2 = map(int, re.findall(r"\d+", node.attrib["bounds"]))
            d.click((x1+x2)//2, (y1+y2)//2)
            search_bottom = y2
            break

    # 3. 주소 입력
    box = d(focused=True)
    if not box.wait(timeout=5):
        print("주소 입력창 못 찾음")
        return False
    box.set_text(address)
    d.press("enter")

    # 4. 검색 결과 확인
    parts = [re.sub(r"[,]", "", part).strip() for part in address.split()]
    road = next((part for part in parts if part.endswith(("대로", "로", "길"))), None)
    if not road:
        print("도로명 추출 실패:", address)
        return False
    for _ in range(30):
        root = ET.fromstring(d.dump_hierarchy())
        results = []
        for node in root.iter():
            text = node.attrib.get("text", "")
            desc = node.attrib.get("content-desc", "")
            if road in text or road in desc:
                x1, y1, x2, y2 = map(int, re.findall(r"\d+", node.attrib["bounds"]))
                if y1 > search_bottom + 100:
                    results.append((y1, x1, x2, y2))
        if results:
            y1, x1, x2, y2 = min(results)
            d.click((x1+x2)//2, (y1+y2)//2)
            break
        time.sleep(0.3)
    else:
        print("주소 검색결과 못 찾음")
        return False

    # 5. 상세주소 입력
    box = d(focused=True)
    if not box.wait(timeout=5):
        print("상세주소 입력창 못 찾음")
        return False
    box.set_text("정문")

    # 6. 주소 등록
    root = ET.fromstring(d.dump_hierarchy())
    for node in root.iter():
        if node.attrib.get("text", "") == "주소 등록":
            x1, y1, x2, y2 = map(int, re.findall(r"\d+", node.attrib["bounds"]))
            d.click((x1+x2)//2, (y1+y2)//2)
            break

    # 7. 팝업 저장
    for _ in range(30):
        xml = d.dump_hierarchy()
        if "주소는 10개까지" in xml:
            d.click(575,923)
            print("저장 클릭")
            break
        time.sleep(0.3)
    else:
        print("주소 10개 팝업 못 찾음")
        return False

    # 8. 홈 화면 확인
    home = d.xpath('//*[@clickable="true" and' '(contains(@text,"검색")or contains(@content-desc,"검색"))]')
    if not home.wait(timeout=20):
        print("배민 첫 화면 로딩 실패")
        return False
    print("주소 설정 완료:", address)
    return True


def clean_result_name(name):
    name = str(name).strip()
    name = re.sub(r"^민트스타\s*가게,\s*", "", name)
    name = re.sub(r",\s*준비중\s*$", "", name)
    name = re.sub(r",\s*배달팁\s*(?:[\d,]+\s*원|할인)\s*$", "", name)
    name = name.replace("＆", "&")
    return name.strip()


def split_store_name(name):
    generic_tails = {"본점", "직영점", "가맹점", "직영"}
    name = clean_result_name(name)
    tokens = name.split()
    while (tokens and tokens[-1] in generic_tails):
        tokens.pop()
    branch = None
    if tokens:
        last = tokens[-1]
        if (last.endswith("점") and len(last) >= 2):
            branch = (last .replace("직영점", "점") .replace("본점", "점"))
            tokens = tokens[:-1]
    core = "".join(tokens).lower()
    core = re.sub(r"[^가-힣a-z0-9]", "", core)
    return core, branch


def name_match(target_name, candidate_name):
    if "검색결과 더보기" in candidate_name:
        return False
    candidate_name = clean_result_name(candidate_name)
    candidate_clean = (candidate_name .replace(" ", "") .replace("직영", "") .lower())
    target_clean = (target_name .replace(" ", "") .replace("직영", "") .lower())
    if (target_clean in candidate_clean or candidate_clean in target_clean):
        return True
    words = target_name.split()
    for word in words:
        word = (word .replace(" ", "") .replace("직영", "") .lower())
        if not word:
            continue
        if word not in candidate_clean:
            return False
    return True


def same_store_candidate(target_name, candidate_name, menu_name):
    def clean_brand_token(token):
        token = str(token or "").replace("직영", "").lower()
        return re.sub(r"[^가-힣a-z0-9]", "", token)
    target_name = clean_result_name(target_name)
    candidate_name = clean_result_name(candidate_name)
    target_tokens = target_name.split()
    candidate_tokens = candidate_name.split()
    if not target_tokens or not candidate_tokens:
        return False
    target_brand = clean_brand_token(target_tokens[0])
    candidate_brand = clean_brand_token(candidate_tokens[0])
    if not target_brand or not candidate_brand:
        return False
    if target_brand == candidate_brand:
        return True
    if min(len(target_brand), len(candidate_brand)) < 3:
        return False
    if target_brand in candidate_brand or candidate_brand in target_brand:
        return True
    common_prefix_len = 0
    for target_char, candidate_char in zip(target_brand, candidate_brand):
        if target_char != candidate_char:
            break
        common_prefix_len += 1
    if common_prefix_len >= 3:
        prefix_ratio = common_prefix_len / min(len(target_brand), len(candidate_brand))
        if prefix_ratio >= 0.4:
            return True
    similarity = SequenceMatcher(None, target_brand, candidate_brand).ratio()
    if similarity >= 0.80:
        return True
    return False


def open_store(store_name, item_name=None):
    search = d.xpath('//*[@clickable="true" and (contains(@text,"검색") or contains(@content-desc,"검색"))]')
    if not search.wait(timeout=5):
        print("검색창 못 찾음")
        return {"status": "TECHNICAL_FAIL", "reason": "검색창 못 찾음", "candidates": []}
    search.click()
    search_names = [store_name]
    alt_name = store_name.replace("직영", "")
    if alt_name != store_name:
        search_names.append(alt_name)
    same_store_candidates = []
    saw_store_card = False
    saw_no_result_message = False

    # 1. 대상처 검색
    for search_name in search_names:
        print("검색 시도:", search_name)
        box = d(className="android.widget.EditText")
        if not box.wait(timeout=5):
            print("검색 입력창 못 찾음")
            return {"status": "TECHNICAL_FAIL", "reason": "검색 입력창 못 찾음", "candidates": []}
        box.set_text(search_name)
        time.sleep(0.5)
        d.press("enter")
        time.sleep(0.5)

        # 2. 검색결과 확인
        for _ in range(20):
            xml = d.dump_hierarchy()
            root = ET.fromstring(xml)
            if any(phrase in xml for phrase in [
                "검색 결과가 없어요",
                "검색결과가 없어요",
                "검색 결과가 없습니다",
                "검색결과가 없습니다",
                "검색 결과 없음"
            ]):
                saw_no_result_message = True
            results = []
            screen_width, screen_height = d.window_size()
            delivery_end_y = screen_height + 1

            # 3. 배달 영역 끝 위치 확인
            for node in root.iter():
                text = node.attrib.get("text", "").strip()
                desc = node.attrib.get("content-desc", "").strip()
                value = text or desc
                bounds = get_bounds(node)
                if not value or not bounds:
                    continue
                x1, y1, x2, y2 = bounds
                if y1 > 250 and x1 == 0 and x2 >= screen_width * 0.9 and (value == "픽업" or "픽업 이런 가게는 어때요?" in value):
                    delivery_end_y = min(delivery_end_y, y1)

            # 4. 배달 영역 가게 수집
            for node in root.iter():
                text = node.attrib.get("text", "").strip()
                desc = node.attrib.get("content-desc", "").strip()
                value = text or desc
                if not value:
                    continue
                bounds = get_bounds(node)
                if not bounds:
                    continue
                x1, y1, x2, y2 = bounds
                if not (y1 > 300
                        and y1 < delivery_end_y
                        and x1 < 100
                        and "검색결과 더보기" not in value
                        and (y2 - y1) > 100):
                    continue
                saw_store_card = True
                candidate = clean_result_name(value)
                search_key = re.sub(r"[^가-힣a-z0-9]", "", str(search_name).lower())
                candidate_key = re.sub(r"[^가-힣a-z0-9]", "", str(candidate).lower())
                exact_match = name_match(search_name, value)
                punctuation_match = bool(search_key and candidate_key and search_key == candidate_key)
                if exact_match or punctuation_match:
                    results.append({"name": value,
                                    "clean_name": candidate,
                                    "bounds": (x1, y1, x2, y2),
                                    "name_typo": punctuation_match and not exact_match,
                                    "searched_name": search_name})
                elif same_store_candidate(store_name, candidate, item_name):
                    if candidate not in same_store_candidates:
                        same_store_candidates.append(candidate)

            # 5. 자동매칭 결과 처리
            if results:
                unique_results = []
                seen = set()
                for r in results:
                    name_key = r["clean_name"].replace(" ", "").replace("직영", "").lower()
                    key = (name_key, r["bounds"])
                    if key not in seen:
                        seen.add(key)
                        unique_results.append(r)
                print("검색결과 개수:", len(unique_results))
                for r in unique_results:
                    print(r["clean_name"], r["bounds"])
                if len(unique_results) >= 2:
                    print("사람확인 필요: 대상처 매칭 후보 여러 개")
                    return {"status": "MULTIPLE_MATCH_REVIEW",
                            "reason": "대상처 매칭 후보 여러 개",
                            "candidates": [r["clean_name"] for r in unique_results]}
                selected = unique_results[0]
                x1, y1, x2, y2 = selected["bounds"]
                if selected.get("name_typo"):
                    print("※ 대상처명 표기 차이 자동통과")
                    print("엑셀:", selected["searched_name"])
                    print("배민:", selected["clean_name"])
                d.click((x1 + x2) // 2, (y1 + y2) // 2)
                print("대상처 클릭:", selected["clean_name"])

                # 6. 상세화면 진입 확인
                for _ in range(30):
                    detail_xml = d.dump_hierarchy()
                    if "함께주문" in detail_xml:
                        print("가게 상세화면 진입 완료")
                        return {"status": "FOUND",
                                "reason": "",
                                "selected": selected["clean_name"],
                                "name_typo": selected.get("name_typo", False),
                                "searched_name": selected.get("searched_name"),
                                "candidates": []}
                    time.sleep(0.3)
                print("가게 상세화면 진입 실패")
                return {"status": "TECHNICAL_FAIL", "reason": "가게 상세화면 진입 실패", "candidates": [selected["clean_name"]]}
            time.sleep(0.3)

    # 7. 자동매칭 없음
    if same_store_candidates:
        print("동일 대상처 가능 후보:")
        for name in same_store_candidates:
            print("-", name)
        return {"status": "SAME_STORE_REVIEW", "reason": "동일 대상처 가능 후보", "candidates": same_store_candidates}
    if saw_store_card or saw_no_result_message:
        print("기존 대상처 없음 -> 신규 대상처 탐색 필요")
        return {"status": "REPLACEMENT_NEEDED", "reason": "신규 대상처 탐색 필요", "candidates": []}
    print("검색결과 화면 확인 실패")
    return {"status": "TECHNICAL_FAIL", "reason": "검색결과 화면 확인 실패", "candidates": []}


def to_number(value):
    if not value:
        return None
    nums = re.findall(r"[\d,]+", str(value))
    if nums:
        return int(nums[-1].replace(",", ""))
    return None


def get_delivery():
    root = ET.fromstring(d.dump_hierarchy())

    # 1. 배달정보 펼치기
    for node in root.iter():
        desc = node.attrib.get("content-desc", "")
        if "배달유형별 배달팁" in desc:
            if "펼치기" in desc:
                x1, y1, x2, y2 = map(int, re.findall(r"\d+", node.attrib["bounds"]))
                d.click((x1 + x2) // 2, (y1 + y2) // 2)
                time.sleep(0.5)
                root = ET.fromstring(d.dump_hierarchy())
            break

    # 2. 가게 배달정보 수집
    items = []
    reading = False
    for node in root.iter():
        text = node.attrib.get("text", "").strip()
        desc = node.attrib.get("content-desc", "").strip()
        if text == "최소주문":
            reading = True
        if text == "지금 배달되는 비슷한 가게":
            break
        if reading and (text or desc):
            items.append((text, desc))

    # 3. 최소주문금액
    min_order = None
    for i, (text, desc) in enumerate(items):
        if text == "최소주문":
            for text2, desc2 in items[i+1:i+5]:
                if text2.endswith("원"):
                    min_order = to_number(text2)
                    break
            break

    # 4. 배달종류별 배달비
    def find_fee(delivery_type):
        for i, (text, desc) in enumerate(items):
            if text == delivery_type or desc == delivery_type:
                for text2, desc2 in items[i+1:i+6]:
                    if (
                        text2 in ["가게배달", "알뜰배달", "한집배달"]
                        or desc2 in ["가게배달", "알뜰배달", "한집배달"]
                    ):
                        break
                    if text2.startswith("배달팁"):
                        return to_number(text2)
                    if text2.endswith("원") and "정가" not in desc2:
                        return to_number(text2)
                return None
        return None
    result = {
        "min_order": min_order,
        "store_fee": find_fee("가게배달"),
        "saver_fee": find_fee("알뜰배달"),
        "one_fee": find_fee("한집배달")
    }
    print(result)
    return result


def get_information():
    print("가게정보 클릭")
    d.click(771, 480)
    time.sleep(2)

    # 1. 가게 주소 수집
    root = ET.fromstring(d.dump_hierarchy())
    top_texts = []
    for node in root.iter():
        text = node.attrib.get("text", "").strip()
        if text:
            top_texts.append(text)
    store_address = None
    for i, text in enumerate(top_texts):
        if text == "주소" and i + 1 < len(top_texts):
            store_address = top_texts[i + 1]
            break
    print("가게 주소:", store_address)

    # 2. 사업자정보 수집
    for i in range(5):
        xml = d.dump_hierarchy()
        if "상호명" in xml and "사업자주소" in xml:
            print("사업자정보 찾음")
            break
        print("스크롤:", i + 1)
        d.swipe(450, 1200, 450, 450, duration=0.4)
        time.sleep(0.4)
    else:
        print("사업자정보 못 찾음")
        return None
    root = ET.fromstring(d.dump_hierarchy())
    texts = []
    for node in root.iter():
        text = node.attrib.get("text", "").strip()
        if text:
            texts.append(text)
    def get_value(label):
        for i, text in enumerate(texts):
            if text == label and i + 1 < len(texts):
                return texts[i + 1]
        return None
    owner = get_value("대표자명")
    business_name = get_value("상호명")
    business_address = get_value("사업자주소")
    if business_address:
        business_address = business_address.replace(
            "위 주소는 사업자등록증에 표기된 정보입니다.",
            ""
        ).strip()
    result = {"owner": owner, "business_name": business_name, "address": business_address, "store_address": store_address}
    print(result)
    return result


def get_texts():
    xml = d.dump_hierarchy()
    root = ET.fromstring(xml)
    texts = []
    for node in root.iter():
        text = node.attrib.get("text","").strip()
        if text:
            texts.append(text)
    return texts


def open_menu_search(store_name):
    for attempt in range(2):
        xml = d.dump_hierarchy()
        root = ET.fromstring(xml)
        nodes = []
        for node in root.iter():
            text = node.attrib.get("text", "").strip()
            bounds = node.attrib.get("bounds", "")
            nums = list(map(int, re.findall(r"\d+", bounds)))
            if len(nums) != 4:
                continue
            x1, y1, x2, y2 = nums
            nodes.append({"text": text, "bounds": (x1, y1, x2, y2)})
        target = None

        # 1. 대상처명 검색 메뉴
        search_titles = {f"{store_name}로 검색한 메뉴", f"{store_name}으로 검색한 메뉴"}
        for item in nodes:
            if item["text"] in search_titles:
                target = item
                print("대상처 검색 메뉴 기준 찾음:", item["text"])
                break

        # 2. 주요 메뉴 섹션
        if target is None:
            for item in nodes:
                if item["text"] == "배민Only 메뉴":
                    target = item
                    print("메뉴 기준 찾음:", item["text"])
                    break
        if target is None:
            for item in nodes:
                if item["text"] == "가장 인기 있는 메뉴":
                    target = item
                    print("메뉴 기준 찾음:", item["text"])
                    break
        if target is None:
            for item in nodes:
                if item["text"] == "메뉴소개":
                    target = item
                    print("메뉴 기준 찾음:", item["text"])
                    break

        # 3. 가게별 메뉴 섹션 제목
        if target is None:
            for item in nodes:
                text = item["text"]
                if not text:
                    continue
                if text in ("최근 주문 많은 조합", "추천 조합", "인기 조합"):
                    continue
                x1, y1, x2, y2 = item["bounds"]
                if 28 <= x1 <= 35 and y1 >= 450 and 35 <= (y2 - y1) <= 55 and not re.match(r"^(인기\s*\d+위|사장님 추천|리뷰\s*\d+|품절됐어요)$", text):
                    target = item
                    print("메뉴 섹션 기준 찾음:", text)
                    break

        # 4. [[...]] 메뉴 섹션
        if target is None:
            for item in nodes:
                text = item["text"]
                if text.startswith("[[") and text.endswith("]]"):
                    target = item
                    print("메뉴 섹션 기준 찾음:", text)
                    break

        # 5. 검색 돋보기 클릭
        if target is not None:
            x1, y1, x2, y2 = target["bounds"]
            search_x = x1 + 26
            search_y = y1 - 64
            print("메뉴 기준 위치:", target["bounds"])
            print("돋보기 클릭:", search_x, search_y)
            d.click(search_x, search_y)
            time.sleep(0.5)
            return True

        # 6. 한 번 스크롤 후 재탐색
        if attempt == 0:
            print("메뉴 영역 안 보임 → 한 번 스크롤")
            d.swipe(450, 1200, 450, 700, duration=0.4)
            time.sleep(0.5)
    print("메뉴 영역 못 찾음")
    print("\n현재 화면 메뉴 관련 텍스트:")
    for item in nodes:
        text = item["text"]
        if not text:
            continue
        x1, y1, x2, y2 = item["bounds"]
        if y1 >= 400:
            print(text, item["bounds"])
    return False


def search_menu(menu_name, menu_category=None):
    d.send_keys(menu_name)
    time.sleep(0.5)
    d.send_action("search")
    time.sleep(1)
    texts = get_texts()
    target = (str(menu_name) .replace(" ", "") .lower())
    for text in texts:
        if "\n" in str(text):
            continue
        value = (str(text) .replace(" ", "") .lower())
        if target in value or value in target:
            print("기존 메뉴 있음:", text)
            return True, []
    print("기존 메뉴 검색 안 됨")
    return False, []


def clean_menu_text(menu):
    text = unicodedata.normalize("NFKC", str(menu or "")).lower()
    return re.sub(r"[^가-힣a-z0-9]", "", text)


def get_popular_menu_candidates(old_menu):
    texts = get_texts()
    candidates = []
    seen = set()
    old_clean = clean_menu_text(old_menu)
    in_menu_area = False
    expect_menu = False
    for text in texts:
        value = str(text or "").strip()
        value_check = re.sub(r"\s+", " ", value).strip()
        if not value_check:
            continue
        if value_check == "가장 인기 있는 메뉴" or value_check.endswith("로 검색한 메뉴"):
            in_menu_area = True
            expect_menu = False
            continue
        if not in_menu_area:
            continue
        if "한 달간 주문수가" in value_check:
            continue
        if value_check == "사장님 추천" or re.fullmatch(r"인기\s*\d+\s*위", value_check):
            expect_menu = True
            continue
        if not expect_menu:
            continue
        if re.fullmatch(r"\d{1,2}:\d{2}", value_check):
            continue
        if value_check.startswith("리뷰 "):
            continue
        if "품절" in value_check:
            continue
        if value_check.startswith("오늘 ") and "오픈" in value_check:
            continue
        if len(value_check) > 80:
            continue
        menu_name = value
        menu_clean = clean_menu_text(menu_name)
        if not menu_clean:
            continue
        if menu_name in seen:
            expect_menu = False
            continue
        seen.add(menu_name)
        if min(len(old_clean), len(menu_clean)) >= 3 and (old_clean in menu_clean or menu_clean in old_clean):
            similarity = 1.0
        else:
            similarity = SequenceMatcher(None, old_clean, menu_clean).ratio()
        candidates.append({"rank": len(candidates) + 1, "menu": menu_name, "similarity": similarity})
        expect_menu = False
        if len(candidates) == 5:
            break
    recommendation = None
    if candidates:
        best = max(candidates, key=lambda x: x["similarity"])
        if best["similarity"] > 0:
            recommendation = best
    return candidates, recommendation


def get_bounds(node):
    nums = re.findall(r"\d+", node.attrib.get("bounds", ""))
    if len(nums) == 4:
        return list(map(int, nums))
    return None


def brand_name(name):
    name = str(name or "").strip().lower()
    name = name.replace("민트스타 가게,", "").strip()
    name = re.sub(r",?\s*준비중$", "", name).strip()
    parts = name.split()
    if len(parts) >= 2 and parts[-1].endswith("점"):
        parts = parts[:-1]
    return " ".join(parts)


def find_replacement_candidates(store):
    old_store = str(store.get("store") or "").strip()
    old_menu = str(store.get("menu") or "").strip()
    if not old_menu:
        return {"status": "TECHNICAL_FAIL", "reason": "기존 메뉴 없음", "candidates": [], "recommendation": None}
    print("\n===== 신규 대상처 탐색 =====")
    print("기존 대상처:", old_store)
    print("검색 메뉴:", old_menu)

    # 1. 기존 메뉴 검색
    search = d.xpath('//*[@clickable="true" and ' '(contains(@text,"검색") or contains(@content-desc,"검색"))]')
    if not search.wait(timeout=5):
        return {"status": "TECHNICAL_FAIL", "reason": "검색창 못 찾음", "candidates": [], "recommendation": None}
    search.click()
    time.sleep(0.5)
    box = d(className="android.widget.EditText")
    if not box.wait(timeout=5):
        return {"status": "TECHNICAL_FAIL", "reason": "검색 입력창 못 찾음", "candidates": [], "recommendation": None}
    box.set_text(old_menu)
    d.press("enter")
    time.sleep(1)

    # 2. 배달 탭
    delivery_ok = False
    delivery_tab_bounds = None
    for attempt in range(3):
        root = ET.fromstring(d.dump_hierarchy())
        delivery_node = None
        for node in root.iter():
            if node.attrib.get("content-desc", "").strip() == "배달 탭":
                delivery_node = node
                break
        if delivery_node is None:
            time.sleep(0.5)
            continue
        bounds = get_bounds(delivery_node)
        if not bounds:
            time.sleep(0.5)
            continue
        delivery_tab_bounds = bounds
        if delivery_node.attrib.get("selected", "false") == "true":
            delivery_ok = True
            break
        x1, y1, x2, y2 = bounds
        print(f"배달 탭 클릭 시도 {attempt + 1}")
        d.click((x1 + x2) // 2, (y1 + y2) // 2)
        time.sleep(1)
        root = ET.fromstring(d.dump_hierarchy())
        for node in root.iter():
            if node.attrib.get("content-desc", "").strip() != "배달 탭":
                continue
            checked_bounds = get_bounds(node)
            if checked_bounds:
                delivery_tab_bounds = checked_bounds
            if node.attrib.get("selected", "false") == "true":
                delivery_ok = True
            break
        if delivery_ok:
            break
    if not delivery_ok or not delivery_tab_bounds:
        print("배달 탭 진입 실패")
        return {"status": "TECHNICAL_FAIL", "reason": "배달 탭 진입 실패", "candidates": [], "recommendation": None}
    print("배달 탭 진입 확인 완료")

    # 3. 검색결과 확인
    root = ET.fromstring(d.dump_hierarchy())
    screen_texts = [node.attrib.get("text", "").strip() for node in root.iter() if node.attrib.get("text", "").strip()]
    no_direct_result = ("검색어와 비슷한 결과를 찾아봤어요." in screen_texts and "이런 메뉴나 가게는 어때요?" in screen_texts)
    if no_direct_result:
        print("대체 대상처 검색결과 없음")
        return {"status": "NO_REPLACEMENT_CANDIDATE", "reason": "대체 대상처 없음", "candidates": [], "recommendation": None}

    # 4. 검색어 보정 안내문 확인
    correction_hint_bounds = None
    correction_hint_text = None
    for node in root.iter():
        text = node.attrib.get("text", "").strip()
        if text.endswith("찾으시는 것 같아요."):
            correction_hint_bounds = get_bounds(node)
            correction_hint_text = text
            break
    if correction_hint_bounds:
        print("검색어 보정 안내문 감지:", correction_hint_text)

    # 5. 주문 많은 순
    sort_bounds = None
    sort_state = None
    for _ in range(20):
        root = ET.fromstring(d.dump_hierarchy())
        for node in root.iter():
            text = node.attrib.get("text", "").strip()
            desc = node.attrib.get("content-desc", "").strip()
            bounds = get_bounds(node)
            if not bounds:
                continue
            if "주문 많은 순" in text or "주문 많은 순" in desc:
                sort_bounds = bounds
                sort_state = "ORDER"
                break
            if text == "기본순" or desc == "기본순":
                sort_bounds = bounds
                sort_state = "DEFAULT"
                break
        if sort_bounds:
            break
        if (any(re.fullmatch(r"배달\s*1", n.attrib.get("text", "").strip()) for n in root.iter())
                and any(n.attrib.get("text", "").strip().startswith("최소주문")
                        for n in root.iter())):
            break
        time.sleep(0.3)
    if not sort_bounds:
        root = ET.fromstring(d.dump_hierarchy())
        single_cards = []
        for node in root.iter():
            desc = node.attrib.get("content-desc", "").strip()
            bounds = get_bounds(node)
            if not desc or not bounds:
                continue
            x1, y1, x2, y2 = bounds
            if x1 >= 100 or not (200 <= y1 < 1500) or y2 - y1 < 90:
                continue
            if any(n.attrib.get("text", "").strip().startswith("최소주문")
                   and (lambda b: bool(b) and y1 <= b[1] <= y2 + 15)(get_bounds(n))
                   for n in root.iter()):
                single_cards.append((desc, bounds))
        if len(single_cards) == 1:
            print("검색 결과 1개 확인 → 정렬 생략")
            sort_state = "SINGLE"
        else:
            print("정렬 기준 및 단일 가게 카드 확인 실패")
            return {"status": "TECHNICAL_FAIL",
                    "reason": "정렬 UI 없음 - 검색 결과 화면 확인 필요",
                    "candidates": [name for name, _ in single_cards],
                    "recommendation": None}
    if sort_state == "SINGLE":
        pass
    elif sort_state == "ORDER":
        print("이미 주문 많은 순 상태")
    else:
        x1, y1, x2, y2 = sort_bounds
        sort_x = (x1 + x2) // 2
        if correction_hint_bounds:
            hx1, hy1, hx2, hy2 = correction_hint_bounds
            sort_y = hy1 - 60
            print("기본순 기준 위치:", sort_bounds)
            print("검색어 보정 안내문 위치:", correction_hint_bounds)
            print("보정 안내문 기준 정렬 버튼 클릭:", (sort_x, sort_y))
        else:
            sort_y = y1 - 22
            print("기본순 기준 위치:", sort_bounds)
            print("정렬 버튼 상대좌표 클릭:", (sort_x, sort_y))
        d.click(sort_x, sort_y)
        order_bounds = None
        for attempt in range(20):
            root = ET.fromstring(d.dump_hierarchy())
            for node in root.iter():
                text = node.attrib.get("text", "").strip()
                desc = node.attrib.get("content-desc", "").strip()
                if "주문 많은 순" in text or "주문 많은 순" in desc:
                    order_bounds = get_bounds(node)
                    if order_bounds:
                        break
            if order_bounds:
                break
            if attempt == 6:
                d.click(sort_x, sort_y)
            time.sleep(0.3)
        if not order_bounds:
            print("주문 많은 순 못 찾음")
            return {"status": "TECHNICAL_FAIL", "reason": "주문 많은 순 못 찾음", "candidates": [], "recommendation": None}
        x1, y1, x2, y2 = order_bounds
        d.click((x1 + x2) // 2, (y1 + y2) // 2)
        time.sleep(1)
        print("주문 많은 순 선택 완료")

    # 6. 후보 수집
    root = ET.fromstring(d.dump_hierarchy())
    old_clean = old_menu.replace(" ", "").lower()
    candidates = []
    for node in root.iter():
        if node.attrib.get("clickable") != "true":
            continue
        bounds = get_bounds(node)
        if not bounds:
            continue
        x1, y1, x2, y2 = bounds
        if not (x1 == 0
                and x2 >= 850
                and 300 <= y1 < 1600
                and (y2 - y1) >= 150):
            continue
        texts = []
        descs = []
        values = []
        for child in node.iter():
            text = child.attrib.get("text", "").strip()
            desc = child.attrib.get("content-desc", "").strip()
            if text:
                if text not in texts:
                    texts.append(text)
                if text not in values:
                    values.append(text)
            if desc:
                if desc not in descs:
                    descs.append(desc)
                if desc not in values:
                    values.append(desc)
        if not any(value.startswith("최소주문") for value in values):
            continue
        raw_store_name = descs[0] if descs else ""
        if not raw_store_name:
            continue
        store_name = clean_result_name(raw_store_name)
        menu_name = ""
        for text in texts:
            if text == raw_store_name or text == store_name:
                continue
            if re.fullmatch(r"\d+\.\d", text):
                continue
            if re.fullmatch(r"\([\d,]+\)", text):
                continue
            if text in ["광고", "가게배달", "알뜰배달", "한집배달",
                        "픽업가능", "예약가능", ","]:
                continue
            if text.startswith("최소주문"):
                continue
            menu_name = text
            break
        menu_clean = menu_name.replace(" ", "").lower()
        similarity = SequenceMatcher(None, old_clean, menu_clean).ratio()
        review_count = 0
        for value in values:
            match = re.fullmatch(r"([\d,]+)개", value)
            if match:
                review_count = int(match.group(1).replace(",", ""))
                break
        delivery_fee = None
        fee_match = re.search(r"배달팁\s*([\d,]+)\s*원", raw_store_name)
        if fee_match and "할인" not in " ".join(values) and "할인" not in raw_store_name:
            delivery_fee = int(fee_match.group(1).replace(",", ""))
        if delivery_fee is None and "할인" not in " ".join(values) and "할인" not in raw_store_name:
            for i, value in enumerate(values):
                if value not in ["가게배달", "알뜰배달", "한집배달"]:
                    continue
                for fee_text in values[i + 1:i + 5]:
                    match = re.fullmatch(r"([\d,]+)(?:~([\d,]+))?원", fee_text.replace(" ", ""))
                    if match:
                        fee = match.group(2) or match.group(1)
                        delivery_fee = int(fee.replace(",", ""))
                        break
                if delivery_fee is not None:
                    break

        # 7. 점수
        old_brand = brand_name(old_store)
        new_brand = brand_name(store_name)
        old_brand_tokens = [re.sub(r"[^가-힣a-z0-9]", "", token.lower()) for token in old_brand.split()]
        new_brand_tokens = [re.sub(r"[^가-힣a-z0-9]", "", token.lower()) for token in new_brand.split()]
        old_brand_tokens = [token for token in old_brand_tokens if len(token) >= 3]
        new_brand_tokens = [token for token in new_brand_tokens if len(token) >= 3]
        brand_match = (old_brand == new_brand or any(token in new_brand_tokens for token in old_brand_tokens))
        brand_score = 10 if brand_match else 0
        menu_score = similarity * 30
        if delivery_fee is None:
            fee_score = 0
        elif delivery_fee <= 1000:
            fee_score = 20
        elif delivery_fee <= 2000:
            fee_score = 16
        elif delivery_fee <= 3000:
            fee_score = 12
        elif delivery_fee <= 4000:
            fee_score = 8
        elif delivery_fee <= 5000:
            fee_score = 4
        else:
            fee_score = 0
        if review_count >= 1000:
            review_score = 40
        elif review_count >= 500:
            review_score = 32
        elif review_count >= 200:
            review_score = 24
        elif review_count >= 100:
            review_score = 16
        elif review_count >= 50:
            review_score = 8
        elif review_count > 0:
            review_score = 4
        else:
            review_score = 0
        total_score = round(brand_score + menu_score + fee_score + review_score, 1)
        candidates.append({"store": store_name,
                           "menu": menu_name,
                           "brand_score": brand_score,
                           "menu_similarity": similarity,
                           "delivery_fee": delivery_fee,
                           "review_count": review_count,
                           "score": total_score})
    if sort_state == "SINGLE" and not candidates:
        raw_name, (cx1, cy1, cx2, cy2) = single_cards[0]
        store_name = clean_result_name(raw_name)
        texts = []
        values = []
        for node in root.iter():
            bounds = get_bounds(node)
            if not bounds:
                continue
            x1, y1, x2, y2 = bounds
            if cy1 <= y1 <= cy2 + 15 and x1 >= 180:
                t = node.attrib.get("text", "").strip()
                desc = node.attrib.get("content-desc", "").strip()
                if t:
                    texts.append((y1, t))
                    values.append(t)
                if desc:
                    values.append(desc)
        texts.sort()
        menu_name = next((t for _, t in texts
                          if not (t.startswith("최소주문") or t.startswith("배달팁")
                                  or t in ("할인", "픽업가능")
                                  or re.fullmatch(r"[\d.,()]+", t))), "")
        review_count = 0
        for value in values:
            m = re.fullmatch(r"([\d,]+)개", value)
            if m:
                review_count = int(m.group(1).replace(",", ""))
                break
        fee_match = next((re.search(r"배달팁\s*([\d,]+)원", v) for v in values if re.search(r"배달팁\s*([\d,]+)원", v)), None)
        delivery_fee = (int(fee_match.group(1).replace(",", "")) if fee_match and "할인" not in " ".join(values) else None)
        similarity = SequenceMatcher(None, old_menu.replace(" ", "").lower(),
                                     menu_name.replace(" ", "").lower()).ratio()
        old_brand = brand_name(old_store)
        new_brand = brand_name(store_name)
        old_tokens = [t for t in re.findall(r"[가-힣a-z0-9]+", old_brand) if len(t) >= 3]
        new_tokens = re.findall(r"[가-힣a-z0-9]+", new_brand)
        brand_match = old_brand == new_brand or any(t in new_tokens for t in old_tokens)
        brand_score = 10 if brand_match else 0
        menu_score = similarity * 30
        fee_score = (0 if delivery_fee is None else
                     20 if delivery_fee <= 1000 else 16 if delivery_fee <= 2000 else
                     12 if delivery_fee <= 3000 else 8 if delivery_fee <= 4000 else
                     4 if delivery_fee <= 5000 else 0)
        review_score = (40 if review_count >= 1000 else 32 if review_count >= 500 else
                        24 if review_count >= 200 else 16 if review_count >= 100 else
                        8 if review_count >= 50 else 4 if review_count > 0 else 0)
        candidates.append({"store": store_name, "menu": menu_name,
                           "brand_score": brand_score, "menu_similarity": similarity,
                           "delivery_fee": delivery_fee, "review_count": review_count,
                           "is_preparing": "준비중" in raw_name,
                           "score": round(brand_score + menu_score + fee_score + review_score, 1)})

    # 8. 추천
    candidates.sort(key=lambda x: x["score"], reverse=True)
    print("\n===== 신규 대상처 후보 =====")
    for i, item in enumerate(candidates, 1):
        print(f'\n{i}위: {item["store"]}')
        print("메뉴:", item["menu"])
        print("상호점수:", item["brand_score"])
        print("메뉴 유사도:", f'{item["menu_similarity"] * 100:.1f}%')
        print("배달비:", item["delivery_fee"])
        print("리뷰수:", item["review_count"])
        print("총점:", item["score"])
    food_terms = ("족발", "보쌈", "파스타", "필라프", "초밥", "쌀국수", "칼국수",
                  "짬뽕", "김밥", "피자", "치킨", "마라탕", "떡볶이", "햄버거",
                  "샐러드", "돈가스", "우동", "회", "커피")
    reference = " ".join([old_store, str(store.get("item") or ""), old_menu])
    expected_terms = [term for term in food_terms if term in reference]
    bakery_terms = ("파리바게뜨", "파리바게트", "뚜레쥬르", "베이커리", "제과점", "빵집")
    recommendable = []
    for item in candidates:
        if item.get("is_preparing"):
            continue
        name = item["store"]
        same_genre = bool(expected_terms and any(term in name for term in expected_terms))
        if "양식" in reference and any(word in name for word in bakery_terms):
            continue
        if item["brand_score"] > 0 or same_genre:
            recommendable.append(item)
    recommendation = recommendable[0] if recommendable else None
    if recommendation:
        print("\n★ 최종 추천:",
              recommendation["store"],
              "|",
              recommendation["score"],
              "점")
        return {"status": "REPLACEMENT_CANDIDATES",
                "candidates": candidates,
                "recommendation": recommendation,
                "order_sorted": sort_state != "SINGLE"}
    print("\n★ 자동 추천 가능한 같은 업종 대상처 없음 → 사람 확인")
    return {"status": "REPLACEMENT_REVIEW" if candidates else "NO_REPLACEMENT_CANDIDATE",
            "reason": "대체 대상처 업종·영업상태 확인 필요" if candidates else "대체 대상처 없음",
            "candidates": candidates,
            "recommendation": None,
            "order_sorted": sort_state != "SINGLE"}


def money(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return int(value)
    value = str(value).replace(",", "").replace("원", "").strip()
    if value.isdigit():
        return int(value)
    return value


def normalize_region_name(value):
    if value is None:
        return None
    text = str(value).strip()
    region_aliases = {
        "서울특별시": "서울특별시", "서울시": "서울특별시", "서울": "서울특별시",
        "부산광역시": "부산광역시", "부산시": "부산광역시", "부산": "부산광역시",
        "대구광역시": "대구광역시", "대구시": "대구광역시", "대구": "대구광역시",
        "인천광역시": "인천광역시", "인천시": "인천광역시", "인천": "인천광역시",
        "광주광역시": "광주광역시", "광주시": "광주광역시", "광주": "광주광역시",
        "대전광역시": "대전광역시", "대전시": "대전광역시", "대전": "대전광역시",
        "울산광역시": "울산광역시", "울산시": "울산광역시", "울산": "울산광역시",
        "세종특별자치시": "세종특별자치시", "세종시": "세종특별자치시", "세종": "세종특별자치시",
        "경기도": "경기도", "경기": "경기도",
        "강원특별자치도": "강원특별자치도", "강원도": "강원특별자치도", "강원": "강원특별자치도",
        "충청북도": "충청북도", "충북": "충청북도",
        "충청남도": "충청남도", "충남": "충청남도",
        "전북특별자치도": "전북특별자치도", "전라북도": "전북특별자치도", "전북": "전북특별자치도",
        "전라남도": "전라남도", "전남": "전라남도",
        "경상북도": "경상북도", "경북": "경상북도",
        "경상남도": "경상남도", "경남": "경상남도",
        "제주특별자치도": "제주특별자치도", "제주도": "제주특별자치도", "제주": "제주특별자치도"
    }
    for alias in sorted(region_aliases, key=len, reverse=True):
        if text.startswith(alias):
            text = region_aliases[alias] + text[len(alias):]
            break
    return text


def clean_address(value):
    if value is None:
        return None
    value = normalize_region_name(value)
    value = re.sub(r"\s+", "", value)
    value = value.replace(",", "")
    return value


def has_address_unit_conflict(old_address, new_address):
    if not old_address or not new_address:
        return False
    old_text = normalize_region_name(old_address)
    new_text = normalize_region_name(new_address)
    for unit in ("동", "층", "호"):
        old_units = set(re.findall(rf"(\d+)\s*{unit}(?![가-힣])", old_text))
        new_units = set(re.findall(rf"(\d+)\s*{unit}(?![가-힣])", new_text))
        if old_units and new_units and old_units != new_units:
            return True
    return False

def same_address(old_address, new_address):
    if old_address is None or new_address is None:
        return False
    if has_address_unit_conflict(old_address, new_address):
        return False
    old_text = normalize_region_name(old_address)
    new_text = normalize_region_name(new_address)
    old_text = re.sub(r"\s+", " ", str(old_text).replace(",", " ").strip())
    new_text = re.sub(r"\s+", " ", str(new_text).replace(",", " ").strip())
    if old_text == new_text:
        return True
    old_tokens = old_text.split()
    new_tokens = new_text.split()
    if len(new_tokens) >= len(old_tokens) and new_tokens[:len(old_tokens)] == old_tokens:
        return True
    if len(old_tokens) >= 5 and len(new_tokens) >= 5 and old_tokens[:5] == new_tokens[:5]:
        unit_conflict = False
        for unit in ["동", "층", "호"]:
            old_units = set(re.findall(rf"(\d+)\s*{unit}", old_text))
            new_units = set(re.findall(rf"(\d+)\s*{unit}", new_text))
            if old_units and new_units and old_units != new_units:
                unit_conflict = True
                break
        if not unit_conflict:
            return True
    old = clean_address(old_address)
    new = clean_address(new_address)
    if old is None or new is None:
        return False
    old_simple = re.sub(r"\([^,()]*?(?:동|읍|면|리|가)\)$", "", old)
    new_simple = re.sub(r"\([^,()]*?(?:동|읍|면|리|가)\)$", "", new)
    if old_simple == new_simple:
        return True
    old_base = re.sub(r"\([^()]*\)$", "", old)
    new_base = re.sub(r"\([^()]*\)$", "", new)
    if (old_base != old) != (new_base != new) and old_base == new_base:
        return True
    old_no_gu = re.sub(r"^((?:서울특별시|[가-힣]+광역시|세종특별자치시))[가-힣]+구", r"\1", old, count=1)
    new_no_gu = re.sub(r"^((?:서울특별시|[가-힣]+광역시|세종특별자치시))[가-힣]+구", r"\1", new, count=1)
    if old_no_gu != old and old_no_gu == new:
        return True
    if new_no_gu != new and new_no_gu == old:
        return True
    old_numbers = re.findall(r"\d+(?:-\d+)?", old)
    new_numbers = re.findall(r"\d+(?:-\d+)?", new)
    if old_numbers == new_numbers and (old in new or new in old):
        return True
    return False


def address_similarity(old_address, new_address):
    old = clean_address(old_address)
    new = clean_address(new_address)
    if old is None or new is None:
        return 0.0
    return SequenceMatcher(None, old, new).ratio()


def is_jibun_address_change(old_address, new_address):
    old_text = str(old_address or "")
    new_text = str(new_address or "")
    pattern = r"([가-힣]+동\d*가|[가-힣]+(?:동|리|읍|면))\s+(\d+(?:-\d+)?)"
    old_match = re.search(pattern, old_text)
    new_match = re.search(pattern, new_text)
    if not old_match or not new_match:
        return False
    old_area, old_lot = old_match.groups()
    new_area, new_lot = new_match.groups()
    return old_area == new_area and old_lot != new_lot


def check_address(old_address, business_address, store_address, target_changed):
    if target_changed:
        new_address = store_address or business_address
        if new_address is None:
            return {"new_address": None, "review": True, "changed": False, "similarity": 0.0, "reason": "신규 대상처 주소 확인 불가"}
        return {"new_address": new_address,
                "review": False,
                "changed": not same_address(old_address, new_address),
                "similarity": address_similarity(old_address, new_address),
                "reason": ""}
    business_same = same_address(old_address, business_address)
    store_same = same_address(old_address, store_address)
    business_sim = address_similarity(old_address, business_address)
    store_sim = address_similarity(old_address, store_address)
    if not business_same and not store_same:
        if not has_address_unit_conflict(old_address, business_address) and is_jibun_address_change(old_address, business_address):
            return {"new_address": business_address,
                    "review": False,
                    "changed": True,
                    "similarity": business_sim,
                    "reason": ""}
        if not has_address_unit_conflict(old_address, store_address) and is_jibun_address_change(old_address, store_address):
            return {"new_address": store_address, "review": False, "changed": True, "similarity": store_sim, "reason": ""}
    if business_same:
        new_address = business_address
    elif store_same:
        new_address = store_address
    elif business_sim >= store_sim:
        new_address = business_address
    else:
        new_address = store_address
    similarity = max(business_sim, store_sim)
    unit_conflict = has_address_unit_conflict(old_address, new_address)
    review = unit_conflict or not (business_same or store_same or similarity >= 0.85)
    if unit_conflict:
        reason = "상세주소 동·층·호 불일치"
    elif not review:
        reason = ""
    elif similarity >= 0.80:
        reason = "동일주소 여부 확인"
    else:
        reason = "주소 불일치 확인"
    return {"new_address": new_address, "review": review, "changed": False, "similarity": similarity, "reason": reason}


def find_address_typo(old_address, new_address):
    old_text = str(old_address or "")
    new_text = str(new_address or "")
    old_tokens = re.findall(r"[가-힣]+(?:동|읍|면|리|가)", old_text)
    new_tokens = re.findall(r"[가-힣]+(?:동|읍|면|리|가)", new_text)
    old_numbers = re.findall(r"\d+(?:-\d+)?", old_text)
    new_numbers = re.findall(r"\d+(?:-\d+)?", new_text)
    if old_numbers != new_numbers:
        return None
    for old_token in old_tokens:
        for new_token in new_tokens:
            if old_token == new_token:
                continue
            if len(old_token) != len(new_token):
                continue
            if old_token[-1] != new_token[-1]:
                continue
            diff_count = sum(a != b for a, b in zip(old_token, new_token))
            if diff_count == 1:
                return {"old": old_token, "new": new_token}
    return None


def delivery_status(old_fee, new_fee, target_changed):
    if target_changed and old_fee != new_fee:
        return "대상처변동"
    if old_fee is None and new_fee is not None:
        return "신규"
    if old_fee is not None and new_fee is None:
        return "탈락"
    return ""


def mark_red(cell):
    font = copy(cell.font)
    font.color = "FFFF0000"
    cell.font = font


def write_review_note(ws, row, result):
    review_types = []
    if result is None:
        review_types.append("기술 실패")
    else:
        status = result.get("status")
        reason = result.get("reason", "")
        if status == "SAME_STORE_REVIEW":
            review_types.append("유사 대상처 확인")
        elif status == "MULTIPLE_MATCH_REVIEW":
            review_types.append("복수 대상처 확인")
        elif (
            result.get("target_changed")
            or status in ["REPLACEMENT_SEARCH_FAIL", "NO_REPLACEMENT_CANDIDATE"]):
            review_types.append("신규 대상처 확인")
        if result.get("menu_exists") is False:
            review_types.append("메뉴 확인")
        if result.get("address_review"):
            review_types.append("주소 확인")
        if result.get("review") and not review_types:
            review_types.append("기술 실패")
    review_types = list(dict.fromkeys(review_types))
    review_col = get_column_letter(column_index_from_string(RESULT_COLS["notes"]) + 1)
    ws[f"{review_col}{row}"] = ", ".join(review_types)


def normalize_owner_name(value):
    if value is None:
        return ""
    text = unicodedata.normalize(
        "NFKC",
        str(value)
    ).casefold()
    return "".join(ch for ch in text if ch.isalnum())


def compare_and_write(ws, store, delivery=None, information=None, menu_exists=None, new_store=None, review_types=None):
    row = store["row"]
    review_types = list(review_types or [])

    # 1. 전월값
    old_store = store["store"]
    old_owner = store["owner"]
    old_address = ws[f"K{row}"].value
    prev_store_min = money(ws[f'{PREV_COLS["store_min"]}{row}'].value)
    prev_store_fee = money(ws[f'{PREV_COLS["store_fee"]}{row}'].value)
    prev_saver_min = money(ws[f'{PREV_COLS["saver_min"]}{row}'].value)
    prev_saver_fee = money(ws[f'{PREV_COLS["saver_fee"]}{row}'].value)
    prev_one_min = money(ws[f'{PREV_COLS["one_min"]}{row}'].value)
    prev_one_fee = money(ws[f'{PREV_COLS["one_fee"]}{row}'].value)
    if new_store is None:
        new_store = old_store
    target_changed = old_store != new_store

    # 2. 배달정보
    delivery_ok = delivery is not None
    curr_store_min = None
    curr_store_fee = None
    curr_saver_min = None
    curr_saver_fee = None
    curr_one_min = None
    curr_one_fee = None
    if delivery_ok:
        curr_min = money(delivery["min_order"])
        curr_store_fee = money(delivery["store_fee"])
        curr_saver_fee = money(delivery["saver_fee"])
        curr_one_fee = money(delivery["one_fee"])
        curr_store_min = curr_min if curr_store_fee is not None else None
        curr_saver_min = curr_min if curr_saver_fee is not None else None
        curr_one_min = curr_min if curr_one_fee is not None else None

    # 3. 사업자정보
    information_ok = information is not None
    new_owner = None
    owner_changed = False
    new_address = None
    address_review = False
    address_changed = False
    address_similarity_value = 0.0
    address_review_reason = ""
    if information_ok:
        new_owner = information.get("owner")
        old_owner_clean = unicodedata.normalize("NFKC", str(old_owner or "")).casefold()
        new_owner_clean = unicodedata.normalize("NFKC", str(new_owner or "")).casefold()
        old_owner_clean = "".join(ch for ch in old_owner_clean if ch.isalnum())
        new_owner_clean = "".join(ch for ch in new_owner_clean if ch.isalnum())
        if not new_owner_clean:
            owner_changed = False
        elif not old_owner_clean:
            owner_changed = True
        elif old_owner_clean == new_owner_clean:
            owner_changed = False
        elif old_owner_clean in new_owner_clean or new_owner_clean in old_owner_clean:
            owner_changed = False
        else:
            owner_changed = True
        address_result = check_address(old_address,
                                       information.get("address"),
                                       information.get("store_address"),
                                       target_changed)
        new_address = address_result["new_address"]
        address_review = address_result["review"]
        address_changed = address_result["changed"]
        address_similarity_value = address_result["similarity"]
        address_review_reason = address_result["reason"]
        if address_review:
            review_types.append("주소 확인")

    # 4. 메뉴 확인
    if menu_exists is False:
        review_types.append("메뉴 확인")

    # 5. 대상처 변경
    if target_changed:
        review_types.append("신규 대상처 확인")
    review_types = list(dict.fromkeys(review_types))

    # 6. 배달정보 비교
    min_changed = False
    fee_changed = False
    if delivery_ok:
        min_changed = (
            (prev_store_fee is not None and curr_store_fee is not None and prev_store_min != curr_store_min)
            or (prev_saver_fee is not None and curr_saver_fee is not None and prev_saver_min != curr_saver_min)
            or (prev_one_fee is not None and curr_one_fee is not None and prev_one_min != curr_one_min)
        )
        fee_changed = (
            (prev_store_fee is not None and curr_store_fee is not None and prev_store_fee != curr_store_fee)
            or (prev_saver_fee is not None and curr_saver_fee is not None and prev_saver_fee != curr_saver_fee)
            or (prev_one_fee is not None and curr_one_fee is not None and prev_one_fee != curr_one_fee)
        )
    store_status = ""
    saver_status = ""
    one_status = ""
    if delivery_ok:
        store_status = delivery_status(prev_store_fee, curr_store_fee, target_changed)
        saver_status = delivery_status(prev_saver_fee, curr_saver_fee, target_changed)
        one_status = delivery_status(prev_one_fee, curr_one_fee, target_changed)

    # 7. 변동사항
    comparison_complete = delivery_ok and information_ok and not address_review
    change_type = ""
    if comparison_complete and fee_changed:
        if address_changed:
            change_type = "주소"
        elif min_changed:
            change_type = "품질규격"
        elif owner_changed:
            change_type = "기타"
        else:
            change_type = "가격"
    notes = []
    if not target_changed:
        if delivery_ok and min_changed:
            notes.append("배달최소주문금액 변경")
        if information_ok and owner_changed:
            notes.append(f"대표자 변경({old_owner or '-'} -> {new_owner or '-'})")
        if information_ok and not address_review and address_changed:
            notes.append("주소 변경")
    note_text = ", ".join(notes)

    # 8. Excel 부분입력
    written_parts = []
    if not target_changed:
        if delivery_ok:
            ws[f'{CURR_COLS["store_min"]}{row}'] = curr_store_min
            ws[f'{CURR_COLS["store_fee"]}{row}'] = curr_store_fee
            ws[f'{CURR_COLS["saver_min"]}{row}'] = curr_saver_min
            ws[f'{CURR_COLS["saver_fee"]}{row}'] = curr_saver_fee
            ws[f'{CURR_COLS["one_min"]}{row}'] = curr_one_min
            ws[f'{CURR_COLS["one_fee"]}{row}'] = curr_one_fee
            ws[f'{RESULT_COLS["store_status"]}{row}'] = store_status
            ws[f'{RESULT_COLS["saver_status"]}{row}'] = saver_status
            ws[f'{RESULT_COLS["one_status"]}{row}'] = one_status
            written_parts.append("배달정보")
            if prev_store_min != curr_store_min:
                mark_red(ws[f'{CURR_COLS["store_min"]}{row}'])
            if prev_store_fee != curr_store_fee:
                mark_red(ws[f'{CURR_COLS["store_fee"]}{row}'])
            if prev_saver_min != curr_saver_min:
                mark_red(ws[f'{CURR_COLS["saver_min"]}{row}'])
            if prev_saver_fee != curr_saver_fee:
                mark_red(ws[f'{CURR_COLS["saver_fee"]}{row}'])
            if prev_one_min != curr_one_min:
                mark_red(ws[f'{CURR_COLS["one_min"]}{row}'])
            if prev_one_fee != curr_one_fee:
                mark_red(ws[f'{CURR_COLS["one_fee"]}{row}'])
        if information_ok:
            if new_owner is not None:
                ws[f"J{row}"] = new_owner
                written_parts.append("대표자")
                if owner_changed:
                    mark_red(ws[f"J{row}"])
            if new_address is not None and not address_review:
                ws[f"K{row}"] = new_address
                written_parts.append("주소")
                if address_changed:
                    mark_red(ws[f"K{row}"])
        if comparison_complete:
            ws[f'{RESULT_COLS["change"]}{row}'] = change_type
            written_parts.append("변동사항")
        ws[f'{RESULT_COLS["notes"]}{row}'] = note_text
    if target_changed:
        print("엑셀 입력 보류: 신규 대상처 확인")
    elif review_types:
        print("엑셀 부분입력 완료:", ", ".join(written_parts) if written_parts else "입력값 없음")
        print("검토 필요:", ", ".join(review_types))
    else:
        print("엑셀 입력 완료")
    print("행:", row)
    if delivery_ok:
        print("가게배달:", curr_store_min, curr_store_fee, store_status)
        print("알뜰배달:", curr_saver_min, curr_saver_fee, saver_status)
        print("한집배달:", curr_one_min, curr_one_fee, one_status)
    print("최소주문금액 변동:", min_changed)
    print("배달비 가격변동:", fee_changed)
    print("변동사항:", change_type)
    print("비고:", note_text)
    if address_review:
        print("사람확인 필요:", address_review_reason)
        print("기존 주소:", old_address)
        print("현재 주소:", new_address)
        print("주소 유사도:", f"{address_similarity_value * 100:.1f}%")
    return {"notes": notes,
            "min_changed": min_changed,
            "fee_changed": fee_changed,
            "change_type": change_type,
            "address_review": address_review,
            "old_address": old_address,
            "new_address": new_address,
            "address_similarity": address_similarity_value,
            "address_review_reason": address_review_reason,
            "review_types": review_types,
            "excel_written_parts": written_parts,
            "comparison_complete": comparison_complete}


def survey_one(store, ws):
    print("\n===== 조사 시작 =====")
    print("행:", store["row"])
    print("대상처:", store["store"])
    old_store = store["store"]
    new_store = old_store
    target_changed = False
    replacement_result = None

    # 1. 기존 대상처 검색
    open_result = open_store(old_store, store["item"])
    open_status = open_result["status"]
    store_name_typo = open_result.get("name_typo", False)
    store_name_excel = open_result.get("searched_name") or old_store
    store_name_baemin = open_result.get("selected")

    # 2. 기존 대상처 미발견 및 대체 탐색
    if open_status == "REPLACEMENT_NEEDED":
        print("기존 대상처 없음 → 신규 대상처 탐색")
        if not go_home():
            return {"status": "TECHNICAL_FAIL",
                    "review": True,
                    "row": store["row"],
                    "store": old_store,
                    "reason": "신규 대상처 탐색 전 첫 화면 복귀 실패",
                    "candidates": []}
        replacement_result = find_replacement_candidates(store)
        recommendation = replacement_result.get("recommendation")
        if recommendation is None:
            print("신규 대상처 추천 실패")
            return {"status": replacement_result.get("status", "REPLACEMENT_SEARCH_FAIL"),
                    "review": True,
                    "row": store["row"],
                    "store": old_store,
                    "reason": replacement_result.get("reason", "신규 대상처 추천 실패"),
                    "candidates": replacement_result.get("candidates", [])}
        new_store = recommendation["store"]
        print("★ 신규 대상처 추천:", new_store, "| 점수:", recommendation["score"])
        if not go_home():
            return {"status": "TECHNICAL_FAIL",
                    "review": True,
                    "row": store["row"],
                    "store": old_store,
                    "reason": "신규 대상처 진입 전 첫 화면 복귀 실패",
                    "candidates": []}
        new_open_result = open_store(new_store, store["item"])
        if new_open_result["status"] != "FOUND":
            return {"status": new_open_result["status"],
                    "review": True,
                    "row": store["row"],
                    "store": old_store,
                    "reason": "추천 신규 대상처 진입 실패: " + new_open_result["reason"],
                    "candidates": new_open_result.get("candidates", [])}
        new_store = new_open_result.get("selected") or new_store
        target_changed = True
        print("대상처 변경:", old_store, "→", new_store)

    # 3. 사람 확인 및 기술 오류 분기
    elif open_status != "FOUND":
        print("조사 보류:", open_result["reason"])
        return {"status": open_status,
                "review": True,
                "row": store["row"],
                "store": old_store,
                "reason": open_result["reason"],
                "candidates": open_result.get("candidates", [])}

    # 4. 배달정보
    delivery = get_delivery()
    if delivery is None:
        print("배달정보 수집 실패")
        return {"status": "PARTIAL_REVIEW",
                "review": True,
                "row": store["row"],
                "store": old_store,
                "reason": "배달정보 확인",
                "review_types": ["배달정보 확인"],
                "store_name_typo": store_name_typo,
                "store_name_excel": store_name_excel,
                "store_name_baemin": store_name_baemin,
                "candidates": []}
    if delivery["store_fee"] is None and delivery["saver_fee"] is None and delivery["one_fee"] is None:
        print("조사 보류: 배달비 3종 모두 확인 불가")
        return {"status": "PARTIAL_REVIEW",
                "review": True,
                "row": store["row"],
                "store": old_store,
                "reason": "배달정보 확인",
                "review_types": ["배달정보 확인"],
                "store_name_typo": store_name_typo,
                "store_name_excel": store_name_excel,
                "store_name_baemin": store_name_baemin,
                "candidates": []}

    # 5. 사업자정보
    information = get_information()
    if information is None:
        print("사업자정보 수집 실패")
        compare_result = compare_and_write(ws, store, delivery, None, None, new_store, ["사업자정보 확인"])
        return {"status": "PARTIAL_REVIEW",
                "review": True,
                "row": store["row"],
                "store": old_store,
                "target_changed": target_changed,
                "old_store": old_store,
                "new_store": new_store,
                "delivery": delivery,
                "information": None,
                "menu_exists": None,
                "reason": "사업자정보 확인",
                "review_types": compare_result["review_types"],
                "excel_written_parts": compare_result["excel_written_parts"],
                "store_name_typo": store_name_typo,
                "store_name_excel": store_name_excel,
                "store_name_baemin": store_name_baemin,
                "candidates": []}
    d.press("back")
    time.sleep(0.5)

    # 6. 기존 메뉴 확인
    if not open_menu_search(new_store):
        print("메뉴 검색창 진입 실패")
        compare_result = compare_and_write(ws, store, delivery, information, None, new_store, ["메뉴 확인"])
        return {"status": "PARTIAL_REVIEW",
                "review": True,
                "row": store["row"],
                "store": old_store,
                "target_changed": target_changed,
                "old_store": old_store,
                "new_store": new_store,
                "delivery": delivery,
                "information": information,
                "menu_exists": None,
                "reason": "메뉴 확인",
                "review_types": compare_result["review_types"],
                "excel_written_parts": compare_result["excel_written_parts"],
                "address_review": compare_result["address_review"],
                "old_address": compare_result["old_address"],
                "new_address": compare_result["new_address"],
                "address_similarity": compare_result["address_similarity"],
                "address_review_reason": compare_result["address_review_reason"],
                "store_name_typo": store_name_typo,
                "store_name_excel": store_name_excel,
                "store_name_baemin": store_name_baemin,
                "candidates": []}
    menu_exists, _ = search_menu(store["menu"], store["item"])
    d.press("back")
    time.sleep(0.5)

    # 7. 기존 메뉴 없으면 후보 추천
    menu_auto_match = False
    menu_auto_name = None
    menu_auto_similarity = 0.0
    menu_candidates = []
    menu_recommendation = None
    if menu_exists is False:
        menu_candidates, menu_recommendation = get_popular_menu_candidates(store["menu"])
        print("인기메뉴 후보:")
        for item in menu_candidates:
            print(f'{item["rank"]}위:', item["menu"], "| 유사도:", f'{item["similarity"] * 100:.1f}%')
        if menu_recommendation:
            print("★ 대체 추천:", menu_recommendation["menu"], "| 유사도:", f'{menu_recommendation["similarity"] * 100:.1f}%')
        else:
            print("★ 대체 추천 없음")
        if menu_recommendation and menu_recommendation["similarity"] >= 0.80:
            menu_auto_match = True
            menu_auto_name = menu_recommendation["menu"]
            menu_auto_similarity = menu_recommendation["similarity"]
            menu_exists = True
            print("※ 메뉴명 표기 차이 자동통과")
            print("기존:", store["menu"])
            print("현재:", menu_auto_name)
            print("유사도:", f"{menu_auto_similarity * 100:.1f}%")

    # 8. 엑셀 비교 / 입력
    compare_result = compare_and_write(ws, store, delivery, information, menu_exists, new_store)

    # 9. 주소 표기 비교
    address_auto_match = False
    if not target_changed and not compare_result["address_review"]:
        old_address_text = str(compare_result.get("old_address") or "").strip()
        new_address_text = str(compare_result.get("new_address") or "").strip()
        if old_address_text and new_address_text and old_address_text != new_address_text:
            address_auto_match = True
            print("※ 주소 표기 차이 자동통과")
            print("기존:", old_address_text)
            print("현재:", new_address_text)
            print("유사도:", f'{compare_result.get("address_similarity", 0.0) * 100:.1f}%')

    # 10. 결과
    result = {"status": "FOUND",
              "row": store["row"],
              "store": old_store,
              "target_changed": target_changed,
              "old_store": old_store,
              "new_store": new_store,
              "replacement_result": replacement_result,
              "delivery": delivery,
              "menu_exists": menu_exists,
              "menu_candidates": menu_candidates,
              "menu_recommendation": menu_recommendation,
              "menu_auto_match": menu_auto_match,
              "menu_auto_name": menu_auto_name,
              "menu_auto_similarity": menu_auto_similarity,
              "store_name_typo": store_name_typo,
              "store_name_excel": store_name_excel,
              "store_name_baemin": store_name_baemin,
              "address_auto_match": address_auto_match,
              "information": information,
              "notes": compare_result["notes"],
              "address_review": compare_result["address_review"],
              "old_address": compare_result["old_address"],
              "new_address": compare_result["new_address"],
              "address_similarity": compare_result["address_similarity"],
              "address_review_reason": compare_result["address_review_reason"],
              "review_types": compare_result["review_types"],
              "excel_written_parts": compare_result["excel_written_parts"],
              "comparison_complete": compare_result["comparison_complete"]}
    print("===== 조사 완료 =====")
    return result


def survey_region(SURVEY_REGION, show_header=True, overall_offset=0, overall_total=None):
    if SURVEY_REGION not in region_groups:
        raise ValueError(f"조사할 수 없는 지역입니다: {SURVEY_REGION}\n"
                         f"가능 지역: {sorted(region_groups.keys())}")
    region_addresses = region_groups[SURVEY_REGION]
    region_address_total = len(region_addresses)
    region_total = sum(len(stores) for stores in region_addresses.values())
    if overall_total is None:
        overall_total = region_total
    success = 0
    failed = 0
    processed = 0
    review_needed = []
    auto_passed = []
    replacement_queue = []
    survey_records = []
    if show_header:
        print("\n==============================")
        print("조사 지역:", SURVEY_REGION)
        print("주소지:", region_address_total, "개")
        print("조사 대상:", region_total, "개")
        print("==============================")
    for address_index, (address, stores) in enumerate(region_addresses.items(), 1):
        print("\n\n################################")
        print(f"[주소 {address_index}/{region_address_total}]", address)
        print("대상처:", len(stores), "개")
        print("################################")

        # 1. 주소 설정
        if not go_home():
            print("첫 화면 복귀 실패:", address)
            for store in stores:
                processed += 1
                failed += 1
                review_needed.append({"row": store["row"], "store": store["store"], "reason": "주소 설정 전 첫 화면 복귀 실패"})
            overall_processed = overall_offset + processed
            progress_pct = overall_processed / overall_total * 100 if overall_total else 0
            print(f"[전체 {overall_processed}/{overall_total} | {progress_pct:.1f}%] 지역 {SURVEY_REGION} {processed}/{region_total} | 주소 {address_index}/{region_address_total} | 성공 {success} | 확인/실패 {failed}")
            continue
        if not set_address(address):
            print("주소 설정 실패:", address)
            for store in stores:
                processed += 1
                failed += 1
                review_needed.append({"row": store["row"], "store": store["store"], "reason": "조사 주소 설정 실패"})
            overall_processed = overall_offset + processed
            progress_pct = overall_processed / overall_total * 100 if overall_total else 0
            print(f"[전체 {overall_processed}/{overall_total} | {progress_pct:.1f}%] 지역 {SURVEY_REGION} {processed}/{region_total} | 주소 {address_index}/{region_address_total} | 성공 {success} | 확인/실패 {failed}")
            continue

        # 2. 대상처 순차 조사
        previous_result = None
        for i, store in enumerate(stores):
            if i > 0:
                if previous_result is not None and previous_result.get("status") == "FOUND":
                    d.press("back")
                    search = d.xpath('//*[@clickable="true" and (contains(@text,"검색") or contains(@content-desc,"검색"))]')
                    if not search.wait(timeout=2):
                        print("검색창 확인 실패 → 기존 홈 복귀")
                        if not go_home():
                            print("첫 화면 복귀 실패")
                            break
                else:
                    print("직전 조사 미진입/보류 → back 생략, 홈에서 재검색")
                    if not go_home():
                        print("첫 화면 복귀 실패")
                        break
            result = survey_one(store, ws)
            previous_result = result
            write_review_note(ws, store["row"], result)
            survey_records.append({
                "log_id": f"{SURVEY_MONTH}_{SURVEY_REGION}_{store['row']}_{datetime.now():%Y%m%d%H%M%S%f}",
                "time": datetime.now(),
                "region": SURVEY_REGION,
                "address_index": address_index,
                "address": address,
                "store": store.copy(),
                "result": result,
                "progress": round((overall_offset + processed + 1) / overall_total * 100, 1)})
            processed += 1
            if result is not None:
                if result.get("store_name_typo"):
                    auto_passed.append({"row": store["row"],
                                        "store": store["store"],
                                        "reason": "대상처명 표기 차이 자동통과",
                                        "old_store_name": result.get("store_name_excel"),
                                        "new_store_name": result.get("store_name_baemin")})
                if result.get("menu_auto_match"):
                    auto_passed.append({"row": store["row"],
                                        "store": store["store"],
                                        "reason": "메뉴명 표기 차이 자동통과",
                                        "old_menu_name": store["menu"],
                                        "new_menu_name": result.get("menu_auto_name"),
                                        "menu_auto_similarity": result.get("menu_auto_similarity", 0.0)})
                if result.get("address_auto_match"):
                    auto_passed.append({"row": store["row"],
                                        "store": store["store"],
                                        "reason": "주소 표기 차이 자동통과",
                                        "old_address_text": result.get("old_address"),
                                        "new_address_text": result.get("new_address"),
                                        "address_auto_similarity": result.get("address_similarity", 0.0)})

            # 3. 조사 결과 판정
            if result is None:
                failed += 1
                review_needed.append({"row": store["row"], "store": store["store"], "reason": "기술적 조사 실패"})
                print("조사 실패:", store["store"])
            elif result.get("status") == "REPLACEMENT_NEEDED":
                failed += 1
                replacement_queue.append({"row": store["row"],
                                          "store": store["store"],
                                          "category": store["category"],
                                          "item": store["item"],
                                          "old_menu": store["menu"]})
                print("신규 대상처 탐색 필요:", store["store"])
            elif result.get("review"):
                failed += 1
                review_needed.append({"row": store["row"],
                                      "store": store["store"],
                                      "reason": result["reason"],
                                      "candidates": result.get("candidates", [])})
            else:
                needs_review = False
                if result.get("target_changed"):
                    needs_review = True
                    review_needed.append({"row": store["row"],
                                          "store": store["store"],
                                          "reason": "신규 대상처 확인",
                                          "change_reason": "기존 대상처 없음",
                                          "old_store": result["old_store"],
                                          "new_store": result["new_store"],
                                          "replacement_candidates": (result.get("replacement_result") or {}).get("candidates", []),
                                          "replacement_recommendation": (result.get("replacement_result") or {}).get("recommendation"),
                                          "menu_exists": result.get("menu_exists"),
                                          "old_menu": store["menu"],
                                          "menu_candidates": result.get("menu_candidates", []),
                                          "menu_recommendation": result.get("menu_recommendation"),
                                          "delivery": result.get("delivery", {}),
                                          "information": result.get("information", {}),
                                          "new_address": result.get("new_address")})
                if result.get("menu_exists") is False and not result.get("target_changed"):
                    needs_review = True
                    review_needed.append({"row": store["row"],
                                          "store": store["store"],
                                          "reason": "기존 메뉴 검색 안 됨",
                                          "category": ws[f'D{store["row"]}'].value,
                                          "menu_category": ws[f'E{store["row"]}'].value,
                                          "old_menu": store["menu"],
                                          "menu_candidates": result.get("menu_candidates", []),
                                          "menu_recommendation": result.get("menu_recommendation")})
                    print("조사 실패(메뉴 확인 필요):", store["store"])
                if result.get("address_review"):
                    needs_review = True
                    review_needed.append({"row": store["row"],
                                          "store": store["store"],
                                          "reason": "사업자주소 불일치 - " + result.get("address_review_reason", ""),
                                          "old_address": result["old_address"],
                                          "new_address": result["new_address"],
                                          "address_similarity": result.get("address_similarity", 0.0)})
                if needs_review:
                    failed += 1
                else:
                    success += 1
            overall_processed = overall_offset + processed
            progress_pct = overall_processed / overall_total * 100 if overall_total else 0
            print(f"[전체 {overall_processed}/{overall_total} | {progress_pct:.1f}%] 지역 {SURVEY_REGION} {processed}/{region_total} | 주소 {address_index}/{region_address_total} | 성공 {success} | 확인/실패 {failed}")

    # 4. 지역 조사 종료
    auto_passed_rows = len({item["row"] for item in auto_passed})
    print("\n==============================")
    print("지역 조사 종료")
    print("==============================")
    print("조사 지역:", SURVEY_REGION)
    print("주소지:", region_address_total)
    print("전체 대상:", region_total)
    print("처리:", processed)
    print("성공:", success)
    print("사람 확인/실패:", failed)
    print("------------------------------")
    print("자동통과:", auto_passed_rows)
    return {"region": SURVEY_REGION,
            "total": region_total,
            "processed": processed,
            "success": success,
            "failed": failed,
            "review_needed": review_needed,
            "auto_passed": auto_passed,
            "replacement_queue": replacement_queue,
            "records":survey_records}


def format_candidates(candidates):
    lines = []
    for i, item in enumerate(candidates or [], 1):
        if isinstance(item, str):
            lines.append(f"{i}위: {item}")
            continue
        name = item.get("store") or item.get("menu") or item.get("name") or ""
        details = [f"{i}위: {name}"]
        for label, key in [
            ("메뉴", "menu"),
            ("점수", "score"),
            ("메뉴유사도", "menu_similarity"),
            ("배달비", "delivery_fee"),
            ("리뷰", "review_count"),
            ("유사도", "similarity")
        ]:
            value = item.get(key)
            if value is None or (key == "menu" and name == value):
                continue
            if key in ("similarity", "menu_similarity"):
                value = f"{value * 100:.1f}%"
            details.append(f"{label}: {value}")
        lines.append(" | ".join(details))
    return "\n".join(lines)[:32000]


LOG_COLUMN_ORDER = ['지역', '주소번호', '엑셀행', '대상처', '조사판정', '확인사유', '검토 처리', '검토 메모', '상태 코드', '기존 메뉴', '메뉴 확인/대체', '신규 대상처 추천', '메뉴 추천', '자동통과 항목', '가게 최소주문금액', '가게배달비', '알뜰 최소주문금액', '알뜰배달비', '한집 최소주문금액', '한집배달비', '현재 대표자', '현재 사업자주소', '현재 가게주소', '조사주소', '대상처 검색 후보', '신규 대상처 후보 상세', '메뉴 후보 상세', '비고', '조사월', '기록일', '전체진행률(%)', '로그ID']
LOG_COLUMN_WIDTHS = {
    "지역": 10, "주소번호": 9, "엑셀행": 9, "대상처": 30,
    "조사판정": 19, "확인사유": 38,
    "검토 처리": 13, "검토 메모": 38, "상태 코드": 25,
    "기존 메뉴": 28, "메뉴 확인/대체": 30,
    "신규 대상처 추천": 31, "메뉴 추천": 29,
    "자동통과 항목": 24,
    "가게 최소주문금액": 18, "가게배달비": 14,
    "알뜰 최소주문금액": 18, "알뜰배달비": 14,
    "한집 최소주문금액": 18, "한집배달비": 14,
    "현재 대표자": 16, "현재 사업자주소": 44,
    "현재 가게주소": 44, "조사주소": 44,
    "대상처 검색 후보": 48, "신규 대상처 후보 상세": 52,
    "메뉴 후보 상세": 52, "비고": 38,
    "조사월": 10, "기록일": 21,
    "전체진행률(%)": 17, "로그ID": 38,
}
LOG_FONT = Font(name="맑은 고딕", size=8)
LOG_HEADER_FONT = Font(name="맑은 고딕", size=8, bold=True, color="FFFFFF")
LOG_HEADER_FILL = PatternFill(fill_type="solid", fgColor="30445F")

def format_log_row(sheet, row, headers):
    """새로 추가된 조사로그 행의 표시 형식만 설정."""
    sheet.row_dimensions[row].height = 19
    for col, title in enumerate(headers, 1):
        cell = sheet.cell(row, col)
        cell.font = LOG_FONT
        cell.alignment = Alignment(vertical="center")
        if title == "기록일" and cell.value is not None:
            cell.number_format = "yyyy-mm-dd hh:mm:ss"

def prepare_log_sheet(log_ws):
    """기존 로그의 열 이름을 기준으로 보존·재배치하고, 8pt·필터를 적용."""
    if getattr(log_ws, "_log_layout_ready", False):
        return {
            str(log_ws.cell(1, col).value): col
            for col in range(1, log_ws.max_column + 1)
            if log_ws.cell(1, col).value is not None
        }
    old_columns = {}
    for col in range(1, log_ws.max_column + 1):
        name = log_ws.cell(1, col).value
        if name is None or not str(name).strip():
            continue
        name = str(name).strip()
        if name in old_columns:
            raise ValueError(f"조사로그 열 제목 중복: {name}")
        old_columns[name] = col
    target_headers = LOG_COLUMN_ORDER + [name for name in old_columns if name not in LOG_COLUMN_ORDER]
    existing_headers = [
        str(log_ws.cell(1, col).value).strip()
        for col in range(1, log_ws.max_column + 1)
        if log_ws.cell(1, col).value is not None
        and str(log_ws.cell(1, col).value).strip()
    ]
    has_blank_header = len(old_columns) != log_ws.max_column
    needs_reorder = has_blank_header or existing_headers != target_headers
    if needs_reorder:
        existing_rows = [
            [log_ws.cell(row, old_columns[name]).value if name in old_columns else None for name in target_headers]
            for row in range(2, log_ws.max_row + 1)
        ]
        log_ws.delete_cols(1, log_ws.max_column)
        for col, title in enumerate(target_headers, 1):
            log_ws.cell(1, col, title)
        for row_idx, values in enumerate(existing_rows, 2):
            for col_idx, value in enumerate(values, 1):
                if value is not None:
                    log_ws.cell(row_idx, col_idx, value)
    for col, title in enumerate(target_headers, 1):
        cell = log_ws.cell(1, col)
        cell.font = LOG_HEADER_FONT
        cell.fill = LOG_HEADER_FILL
        cell.alignment = Alignment(vertical="center", horizontal="center")
        log_ws.column_dimensions[get_column_letter(col)].width = LOG_COLUMN_WIDTHS.get(title, 24)
    log_ws.row_dimensions[1].height = 24
    for row in range(2, log_ws.max_row + 1):
        format_log_row(log_ws, row, target_headers)
    log_ws.freeze_panes = "E2"
    log_ws.auto_filter.ref = f"A1:{get_column_letter(len(target_headers))}{max(log_ws.max_row, 1)}"
    log_ws.sheet_view.showGridLines = False
    log_ws._log_layout_ready = True
    return {title: col for col, title in enumerate(target_headers, 1)}


def append_log_to_sheet(book=None,records=None):
    if book is None:
        book = wb
    if "조사로그" not in book.sheetnames:
        log_ws = book.create_sheet("조사로그")
    else:
        log_ws = book["조사로그"]
    if records is None :
        records = [record for region_result in all_results.values() for record in region_result.get("records", [])]
    if not records:
        prepare_log_sheet(log_ws)
        print("새 조사 기록 없음")
        return 0
    columns = prepare_log_sheet(log_ws)
    existing_ids = {log_ws.cell(row, columns["로그ID"]).value for row in range(2, log_ws.max_row + 1)}
    added = 0
    for record in records:
        log_id = record["log_id"]
        if log_id in existing_ids:
            continue
        store = record["store"]
        result = record.get("result") or {}
        replacement = result.get("replacement_result") or {}
        recommendation = replacement.get("recommendation") or {}
        delivery = result.get("delivery") or {}
        information = result.get("information") or {}
        menu_recommendation = result.get("menu_recommendation") or {}
        status = result.get("status", "NOT_PROCESSED")
        reasons = list(result.get("review_types") or [])
        if result.get("reason"):
            reasons.insert(0, result["reason"])
        if result.get("target_changed"):
            reasons.append("신규 대상처 확인")
        if result.get("menu_exists") is False:
            reasons.append("기존 메뉴 확인")
        if result.get("address_review"):
            reasons.append("주소 확인")
        reasons = list(dict.fromkeys(reasons))
        if status in ("TECHNICAL_FAIL", "NOT_PROCESSED"):
            judgement = "기술 실패/미조사"
        elif result.get("target_changed"):
            judgement = "신규 대상처 확인"
        elif status in (
            "MULTIPLE_MATCH_REVIEW",
            "SAME_STORE_REVIEW",
            "REPLACEMENT_NEEDED",
            "NO_REPLACEMENT_CANDIDATE",
            "REPLACEMENT_SEARCH_FAIL"
        ):
            judgement = "대상처 확인"
        elif result.get("review") or reasons:
            judgement = "확인 필요"
        else:
            judgement = "성공"
        auto = []
        if result.get("store_name_typo"):
            auto.append("대상처명 표기")
        if result.get("menu_auto_match"):
            auto.append("메뉴명 표기")
        if result.get("address_auto_match"):
            auto.append("주소 표기")
        min_order = delivery.get("min_order")
        candidates = replacement.get("candidates") or []
        notes = result.get("notes") or ""
        if isinstance(notes, list):
            notes = ", ".join(notes)
        data = {
            "기록일": record["time"],
            "로그ID": log_id,
            "지역": record["region"],
            "주소번호": record["address_index"],
            "조사주소": record["address"],
            "엑셀행": store["row"],
            "대상처": store["store"],
            "조사판정": judgement,
            "확인사유": ", ".join(reasons),
            "자동통과 항목": ", ".join(auto),
            "메뉴 확인/대체": (
                result.get("menu_auto_name")
                or menu_recommendation.get("menu")
                or store.get("menu")
            ),
            "가게 최소주문금액": (
                min_order if delivery.get("store_fee") is not None else None
            ),
            "가게배달비": delivery.get("store_fee"),
            "알뜰 최소주문금액": (
                min_order if delivery.get("saver_fee") is not None else None
            ),
            "알뜰배달비": delivery.get("saver_fee"),
            "한집 최소주문금액": (
                min_order if delivery.get("one_fee") is not None else None
            ),
            "한집배달비": delivery.get("one_fee"),
            "비고": notes,
            "현재 대표자": information.get("owner"),
            "현재 사업자주소": information.get("address"),
            "현재 가게주소": information.get("store_address"),
            "전체진행률(%)": record.get("progress"),
            "기존 메뉴": store.get("menu"),
            "대상처 검색 후보": format_candidates(result.get("candidates")),
            "신규 대상처 추천": (
                recommendation.get("store")
                or (result.get("new_store") if result.get("target_changed") else "")
            ),
            "신규 대상처 후보 상세": format_candidates(candidates),
            "메뉴 추천": (
                menu_recommendation.get("menu")
                or result.get("menu_auto_name")
            ),
            "메뉴 후보 상세": format_candidates(result.get("menu_candidates")),
            "상태 코드": status,
            "조사월": SURVEY_MONTH
        }
        for title in data:
            if title not in columns:
                col = log_ws.max_column + 1
                log_ws.cell(1, col, title)
                columns[title] = col
        row_col = columns["엑셀행"]
        current_row = str(store["row"])
        for old_row in range(log_ws.max_row, 1, -1):
            if str(log_ws.cell(old_row, row_col).value) == current_row:
                existing_ids.discard(log_ws.cell(old_row, columns["로그ID"]).value)
                log_ws.delete_rows(old_row)
        target_row = log_ws.max_row + 1
        for title, value in data.items():
            log_ws.cell(target_row, columns[title], value)
        existing_ids.add(log_id)
        added += 1
        format_log_row(log_ws, target_row, [log_ws.cell(1, col).value for col in range(1, log_ws.max_column + 1)])
    log_ws.auto_filter.ref = (f"A1:{get_column_letter(log_ws.max_column)}{max(log_ws.max_row, 1)}")
    return added


BLUESTACKS_DIR = Path(r"C:\Program Files\BlueStacks_nxt")

def restart_bluestacks():
    global d
    player = BLUESTACKS_DIR / "HD-Player.exe"
    if os.name != "nt" or not player.exists():
        raise RuntimeError("Windows 또는 블루스택 실행 경로 확인 필요")
    command = (
        "$p = Get-CimInstance Win32_Process "
        "-Filter \"Name = 'HD-Player.exe'\"; "
        "$p | Select-Object ProcessId,CommandLine "
        "| ConvertTo-Json -Compress"
    )
    result = subprocess.run(["powershell", "-NoProfile", "-Command", command], capture_output=True, text=True, timeout=15)
    if result.returncode != 0:
        raise RuntimeError("블루스택 프로세스 확인 실패")
    processes = json.loads(result.stdout) if result.stdout.strip() else []
    if isinstance(processes, dict):
        processes = [processes]
    if len(processes) != 1:
        raise RuntimeError("블루스택 인스턴스를 정확히 1개 실행해야 합니다.")
    process = processes[0]
    instance = "Pie64"
    pid = process["ProcessId"]
    print("블루스택 인스턴스:", instance)
    print("프로세스 ID:", pid)
    serial = d.serial
    package = d.app_current()["package"]
    print("블루스택 종료:", instance)
    stopped = subprocess.run(["taskkill", "/F", "/PID", str(pid), "/T"], capture_output=True, text=True, timeout=20)
    if stopped.returncode != 0:
        raise RuntimeError("블루스택 종료 실패")
    time.sleep(3)
    print("블루스택 재실행")
    subprocess.Popen([str(player), "--instance", instance], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    deadline = time.monotonic() + 150
    while time.monotonic() < deadline:
        try:
            new_d = u2.connect(serial)
            new_d.info
            new_d.app_start(package)
            if new_d.app_current().get("package") == package:
                new_d.dump_hierarchy()
                d = new_d
                if go_home():
                    print("블루스택 재연결 완료")
                    return True
        except Exception:
            pass
        time.sleep(3)
    raise RuntimeError("블루스택 재연결 실패 - 다음 조사 중단")


class LogTee:
    def __init__(self, screen, buffer):
        self.screen = screen
        self.buffer = buffer
    def write(self, text):
        self.screen.write(text)
        self.buffer.write(text)
    def flush(self):
        self.screen.flush()
        self.buffer.flush()


def remember_region(region):
    for stores in region_groups[region].values():
        for store in stores:
            r = store["row"]
            SURVEY_SNAPSHOT[r] = {col: (ws[f"{col}{r}"].value, copy(ws[f"{col}{r}"].font)) for col in snapshot_cols}


def save_survey_checkpoint(book, path):
    temp_file = path.with_name(path.stem + "_저장중.xlsx")
    book.save(temp_file)
    os.replace(temp_file, path)


SURVEY_GROUPS = [
    ["경기"], ["서울", "강원", "인천"], ["경남", "전남", "제주"],
    ["경북", "충북", "세종"], ["부산", "광주", "대구", "울산"],
    ["전북", "충남", "대전"]
]
PUBLIC_REGIONS = set().union(*map(set, SURVEY_GROUPS))


def _read_targets(mode, target, workbook):
    """Read requested survey targets without modifying Excel or connecting to Android."""
    sheet = workbook["조사자료(rawdata)"]
    available = {}
    for row, values in enumerate(sheet.iter_rows(min_row=8, min_col=6, max_col=9, values_only=True), 8):
        reg = str(values[0] or "").strip()
        addr = str(values[2] or "").strip()
        name = str(values[3] or "").strip()
        if reg in PUBLIC_REGIONS and addr and name:
            available[row] = (reg, addr, name)
    if mode in ("전체", "조합"):
        if mode == "전체":
            numbers = list(range(1, 7))
        else:
            raw = target if isinstance(target, (list, tuple)) else [target]
            if not raw or any(type(n) is not int or n not in range(1, 7) for n in raw):
                raise ValueError("조합 TARGET에는 1~6의 번호를 입력하세요. 예: [2, 4]")
            numbers = list(dict.fromkeys(raw))
        scope = {r for n in numbers for r in SURVEY_GROUPS[n-1]}
        chosen = {r: data for r, data in available.items() if data[0] in scope}
        if len(available) != 745:
            raise ValueError(f"원본 전체조사 조건은 745개입니다. 현재 {len(available)}개이므로 조합/전체조사를 중단합니다.")
        return chosen, numbers
    if mode == "행":
        raw = target if isinstance(target, (list, tuple)) else [target]
        if not raw or any(type(n) is not int or n < 8 for n in raw):
            raise ValueError("행 TARGET에는 8 이상의 Excel 행 번호를 입력하세요. 예: [486]")
        requested = list(dict.fromkeys(raw))
        missing = [r for r in requested if r not in available]
        if missing:
            raise ValueError(f"조사표에서 대상처가 확인되지 않는 행: {missing}")
        chosen = {r: available[r] for r in requested}
    elif mode == "대상처":
        name = str(target or "").strip()
        found = {r: data for r, data in available.items() if data[2] == name}
        if not found:
            raise ValueError(f"조사표에 대상처 이름이 없습니다: {name}")
        if len(found) != 1:
            raise ValueError(f"같은 이름이 {len(found)}개입니다. Excel 행을 지정하세요: {list(found)}")
        chosen = found
    elif mode == "주소":
        address = str(target or "").strip()
        chosen = {r: data for r, data in available.items() if data[1] == address}
    elif mode in ("지역", "여러지역"):
        names = [target] if mode == "지역" else target
        if not isinstance(names, (list, tuple)):
            raise ValueError("여러지역 TARGET은 ['대전', '세종'] 형태로 입력하세요.") if mode == "여러지역" else ValueError("지역명을 입력하세요.")
        names = [str(n).strip() for n in names]
        bad = [n for n in names if n not in PUBLIC_REGIONS]
        if bad:
            raise ValueError(f"조사 가능 지역이 아닙니다: {bad}")
        chosen = {r: data for r, data in available.items() if data[0] in names}
    else:
        raise ValueError("MODE는 행, 대상처, 주소, 지역, 여러지역, 조합, 전체 중 하나를 입력하세요.")
    if not chosen:
        raise ValueError(f"조건에 맞는 조사 대상처가 없습니다: MODE={mode}, TARGET={target}")
    return chosen, None


def _selection_groups(chosen):
    """Make region → address → store subset from the initialized original groups."""
    requested = set(chosen)
    chosen_groups = {}
    for region in SURVEY_GROUPS:
        for name in region:
            if name not in region_groups:
                continue
            for address, stores in region_groups[name].items():
                subset = [store for store in stores if store["row"] in requested]
                if subset:
                    chosen_groups.setdefault(name, {})[address] = subset
    actual = {store["row"] for addresses in chosen_groups.values() for stores in addresses.values() for store in stores}
    if actual != requested:
        raise ValueError(f"선택된 조사 행이 초기화한 Excel과 다릅니다: {sorted(requested - actual)}")
    return chosen_groups


def _run_selection(chosen, result_file):
    """Call unchanged survey_region and append_log_to_sheet per selected address."""
    global wb, ws
    subset_groups = _selection_groups(chosen)
    total = len(chosen)
    processed = success = failed = 0
    output = Path(result_file)
    for region, addresses in subset_groups.items():
        for address, stores in addresses.items():
            original = region_groups[region]
            region_groups[region] = {address: stores}
            try:
                result = survey_region(region, overall_offset=processed, overall_total=total)
            finally:
                region_groups[region] = original
            records = result["records"]
            if result["processed"] != len(stores) or len(records) != len(stores):
                raise RuntimeError(f"완료되지 않은 주소지는 저장하지 않습니다: {region} / {address} (대상 {len(stores)}, 로그 {len(records)})")
            if any(rec.get("result") is None or "TECHNICAL_FAIL" in str((rec.get("result") or {}).get("status", "")) for rec in records):
                raise RuntimeError(f"기술적 실패로 주소지 저장 보류: {region} / {address}")
            added = append_log_to_sheet(book=wb, records=records)
            if added != len(stores):
                raise RuntimeError(f"조사로그 개수 불일치: {region}/{address}: {added}/{len(stores)}")
            save_survey_checkpoint(wb, output)
            processed += result["processed"]
            success += result["success"]
            failed += result["failed"]
            print(f"주소지 저장 완료: {region} / {address} ({processed}/{total})")
    print(f"조사 종료: 총 {total}, 처리 {processed}, 자동처리 {success}, 사람 확인/실패 {failed}")
    print("결과 파일:", output)
    return {"output_file": str(output), "processed": processed, "success": success, "review_or_failed": failed}


def _setup_data(month, file, output_file, new_survey):
    global SURVEY_MONTH, FILE, OUTPUT_FILE, NEW_SURVEY
    global SHEET_NAME, PREV_COLS, CURR_COLS, RESULT_COLS, wb, ws, groups, region_groups, EXCLUDED_REGIONS
    SURVEY_MONTH = month
    FILE = str(file)
    OUTPUT_FILE = str(output_file)
    NEW_SURVEY = new_survey
    SHEET_NAME = "조사자료(rawdata)"
    PREV_COLS = get_month_cols(SURVEY_MONTH - 1)
    CURR_COLS = get_month_cols(SURVEY_MONTH)
    RESULT_COLS = get_result_cols(SURVEY_MONTH)
    if Path(FILE).resolve() == Path(OUTPUT_FILE).resolve():
        raise ValueError("FILE과 OUTPUT_FILE은 서로 다른 파일이어야 합니다.")
    output_exists = Path(OUTPUT_FILE).exists()
    if NEW_SURVEY or not output_exists:
        if output_exists:
            answer = input("새 조사 시작: 기존 OUTPUT_FILE의 조사내용도 FILE 기준으로 바뀝니다. " "기존 결과는 자동 백업됩니다. 진행하려면 '초기화' 입력: ")
            if answer.strip() != "초기화":
                NEW_SURVEY = False
                raise RuntimeError("새 조사 초기화를 취소했습니다. 기존 파일은 유지됩니다.")
            backup_path = Path(OUTPUT_FILE).with_name(
                Path(OUTPUT_FILE).stem
                + "_초기화전백업_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                + Path(OUTPUT_FILE).suffix
            )
            copy2(OUTPUT_FILE, backup_path)
            print("기존 결과 파일 백업:", backup_path)
        wb = load_workbook(FILE)
        if "조사로그" in wb.sheetnames:
            log_ws = wb["조사로그"]
            if log_ws.max_row > 1:
                log_ws.delete_rows(2, log_ws.max_row - 1)
            if log_ws.auto_filter.ref:
                log_ws.auto_filter.ref = f"A1:{get_column_letter(log_ws.max_column)}1"
        else:
            log_ws = wb.create_sheet("조사로그")
            log_ws["A1"] = "기록일"
        wb.save(OUTPUT_FILE)
        NEW_SURVEY = False
        print("새 조사 시작 - 조사로그 0건:", OUTPUT_FILE)
    else:
        wb = load_workbook(OUTPUT_FILE)
        print("기존 조사 이어하기 - 조사로그 유지:", OUTPUT_FILE)
    ws = wb[SHEET_NAME]
    print("조사월:", SURVEY_MONTH)
    print("전월 열:", PREV_COLS)
    print("당월 열:", CURR_COLS)
    print("결과 열:", RESULT_COLS)
    print("파일 불러오기 완료:", OUTPUT_FILE)

    groups = defaultdict(list)
    region_groups = defaultdict(lambda: defaultdict(list))
    EXCLUDED_REGIONS = {"전국", "중앙"}
    for row in range(8, ws.max_row + 1):
        region = ws[f"F{row}"].value
        address = ws[f"H{row}"].value
        store_name = ws[f"I{row}"].value
        if not address or not store_name:
            continue
        store_data = {
            "row": row,
            "category": ws[f"D{row}"].value,
            "item": ws[f"E{row}"].value,
            "store": store_name,
            "owner": ws[f"J{row}"].value,
            "business_address": ws[f"K{row}"].value,
            "menu": ws[f"L{row}"].value
        }
        groups[address].append(store_data)
        if region:
            region = str(region).strip()
            if region not in EXCLUDED_REGIONS:
                region_groups[region][address].append(store_data)
    gyeonggi = list(region_groups["경기"].items())
    mid = (len(gyeonggi) + 1) // 2
    region_groups["경기1"] = dict(gyeonggi[:mid])
    region_groups["경기2"] = dict(gyeonggi[mid:])
    for region in ["경기1", "경기2"]:
        addresses = region_groups[region]
        count = sum(len(stores) for stores in addresses.values())
        print(f"{region}: 주소지 {len(addresses)}개 / 대상처 {count}개")


def _connect(bluestacks_dir):
    """Start BlueStacks' bundled ADB and connect automatically (no terminal needed).

    Do not guess an ADB port or silently choose among multiple devices.
    BlueStacks must already be running, with ADB enabled in its settings.
    """
    global d, BLUESTACKS_DIR
    BLUESTACKS_DIR = Path(bluestacks_dir).expanduser()
    print("BlueStacks 연결 준비:", BLUESTACKS_DIR)

    if os.name != "nt":
        raise RuntimeError("BlueStacks 자동 연결은 Windows 환경에서 실행해야 합니다.")

    adb_exe = BLUESTACKS_DIR / "HD-Adb.exe"
    if not adb_exe.is_file():
        raise FileNotFoundError(
            f"BlueStacks ADB 실행 파일을 찾지 못했습니다: {adb_exe}\n"
            "BLUESTACKS_DIR을 실제 BlueStacks 설치 폴더로 설정하세요."
        )

    def run_adb(command):
        try:
            completed = subprocess.run(
                [str(adb_exe), command],
                capture_output=True, text=True, timeout=25,
                errors="replace",
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise RuntimeError(f"BlueStacks ADB '{command}' 실행 실패: {exc}") from exc
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise RuntimeError(
                f"BlueStacks ADB '{command}' 명령 실패 (종료 코드 {completed.returncode}): {detail}"
            )
        return completed.stdout

    print("BlueStacks ADB 서버 자동 시작 중...")
    run_adb("start-server")
    device_output = run_adb("devices")
    devices = []
    unavailable = []
    for line in device_output.splitlines():
        parts = line.split()
        if len(parts) < 2 or parts[0] == "List" or parts[0].startswith("*"):
            continue
        serial, state = parts[0], parts[1]
        if state == "device":
            devices.append(serial)
        elif state in ("offline", "unauthorized"):
            unavailable.append((serial, state))

    if not devices:
        details = ", ".join(f"{s}({state})" for s, state in unavailable)
        suffix = f"\n현재 표시된 기기 상태: {details}" if details else ""
        raise RuntimeError(
            "BlueStacks ADB에서 연결된 Android 기기를 찾지 못했습니다.\n"
            "BlueStacks 5 인스턴스가 실행 중인지, 설정에서 ADB(안드로이드 디버그 브리지)가 "
            "활성화되어 있는지 확인하세요.\n"
            "연결 준비는 코드가 자동으로 수행하므로 터미널 명령을 직접 입력할 필요는 없습니다."
            + suffix
        )
    if len(devices) > 1:
        raise RuntimeError(
            "ADB 기기가 여러 개 연결되어 있어 안전하게 대상을 결정할 수 없습니다: "
            + ", ".join(devices)
            + "\n조사할 BlueStacks 인스턴스 하나만 실행한 뒤 다시 시도하세요."
        )

    serial = devices[0]
    print("ADB 기기 발견:", serial)
    try:
        connection = u2.connect(serial)
        screen_size = connection.window_size()
    except Exception as exc:
        raise RuntimeError(
            f"ADB에서 {serial} 기기를 찾았지만 uiautomator2 연결에 실패했습니다: {exc}"
        ) from exc
    d = connection
    print("BlueStacks 연결 완료 / 화면 크기:", screen_size)


def _run_group_survey(group_numbers):
    """Original 5-3 flow; only group loop accepts a list, not a contiguous range."""
    global wb, ws
    START_GROUP = min(group_numbers)
    END_GROUP = max(group_numbers)
    SURVEY_GROUPS = [
        ["경기"],
        ["서울", "강원", "인천"],
        ["경남", "전남", "제주"],
        ["경북", "충북", "세종"],
        ["부산", "광주", "대구", "울산"],
        ["전북", "충남", "대전"]
    ]
    PROGRESS_SHEET = "자동조사진행"
    TEMP_REGION_KEY = "__자동전체조사__"

    # 1. 실행 전 검사
    source = Path(OUTPUT_FILE)
    RUN_FILE = source.with_name(source.stem + "_자동전체조사.xlsx")
    if not source.exists():
        raise FileNotFoundError(f"원본 조사 파일이 없습니다: {source}")
    if not 1 <= START_GROUP <= END_GROUP <= len(SURVEY_GROUPS):
        raise ValueError("START_GROUP / END_GROUP 설정 오류")
    if TEMP_REGION_KEY in region_groups:
        raise RuntimeError("임시 조사 지역명이 이미 존재합니다. " "region_groups를 확인하세요.")
    for group in SURVEY_GROUPS:
        for region in group:
            if region not in region_groups:
                raise ValueError(f"조사 지역 정보가 없습니다: {region}")
    total = sum(len(stores) for group in SURVEY_GROUPS for region in group for stores in region_groups[region].values())
    print("전체 자동 조사 준비")
    print("전체 대상처:", total)
    print("조사 조합:", START_GROUP, "~", END_GROUP)
    print("결과 파일:", RUN_FILE)
    if total != 745:
        raise RuntimeError(f"예상 대상처 745개와 실제 {total}개가 다릅니다. " "region_groups를 확인하세요.")

    # 2. 결과 파일 준비
    if not RUN_FILE.exists():
        copy2(source, RUN_FILE)
        print("\n새 전체 조사 파일 생성")
    else:
        print("\n기존 전체 조사 파일에서 이어하기")
    run_wb = load_workbook(RUN_FILE)
    run_ws = run_wb[SHEET_NAME]
    if PROGRESS_SHEET not in run_wb.sheetnames:
        progress_ws = run_wb.create_sheet(PROGRESS_SHEET)
        progress_ws.append(["지역", "주소지", "대상처수", "상태", "완료시각"])
    else:
        progress_ws = run_wb[PROGRESS_SHEET]

    # 3. 기존 완료 주소지 확인
    done = {}
    for values in progress_ws.iter_rows(
        min_row=2,
        values_only=True
    ):
        if not values or not values[0]:
            continue
        region, address, count, status = values[:4]
        if status == "DONE":
            key = (region, address)
            if region not in region_groups:
                raise RuntimeError(f"완료 기록에 없는 지역이 있습니다: {region}")
            if address not in region_groups[region]:
                raise RuntimeError(f"완료 기록에 없는 주소지가 있습니다: " f"{region} / {address}")
            expected = len(region_groups[region][address])
            if count != expected:
                raise RuntimeError(f"완료 기록의 대상처 수 불일치: " f"{region} / {address} " f"(기록 {count}, 현재 {expected})")
            if key in done:
                raise RuntimeError(f"완료 기록 중복: {region} / {address}")
            done[key] = count
    completed = sum(done.values())
    print("이전 완료 대상처:", completed)
    print("남은 대상처:", total - completed)
    original_wb = wb
    original_ws = ws
    wb = run_wb
    ws = run_ws
    start_time = time.monotonic()
    newly_completed = 0
    restart_count = 0
    group_results = {}
    try:

        # 4. 지역 조합 순회
        for group_no in group_numbers:
            regions = SURVEY_GROUPS[group_no - 1]
            pending = []
            for region in regions:
                for address, stores in region_groups[region].items():
                    if (region, address) not in done:
                        pending.append((region, address, stores))
            if not pending:
                print(
                    f"\n[조합 {group_no}/6] "
                    "전체 완료된 조합 → 건너뜀"
                )
                continue
            group_total = sum(len(stores) for _, _, stores in pending)
            print(f"[조합 {group_no}/6] 조사 시작")
            print("지역:", ", ".join(regions))
            print("남은 주소지:", len(pending))
            print("남은 대상처:", group_total)
            print("전체 진행:", f"{completed}/{total}")

            # 5. 블루스택 재시작
            if group_no > 1:
                print("\n===== 블루스택 자동 재시작 =====")
                restart_start = time.monotonic()
                restart_bluestacks()
                restart_elapsed = (time.monotonic() - restart_start)
                restart_count += 1
                print(f"블루스택 재시작 완료: " f"{restart_elapsed:.1f}초")
            group_start = time.monotonic()
            group_processed = 0
            group_success = 0
            group_failed = 0

            # 6. 지역별 주소지 조사
            for region in regions:
                addresses = list(region_groups[region].items())
                for address_index, (address, stores) in enumerate(
                    addresses,
                    1
                ):
                    if (region, address) in done:
                        print(
                            f"\n완료 주소지 건너뜀: "
                            f"{region} / {address}"
                        )
                        continue
                    print("지역:", region)
                    print("주소지:", f"{address_index}/{len(addresses)}")
                    print("주소:", address)
                    print("대상처:", len(stores))
                    region_groups[TEMP_REGION_KEY] = {address: stores}
                    try:
                        result = survey_region(TEMP_REGION_KEY, overall_offset=completed, overall_total=total)
                    finally:
                        region_groups.pop(TEMP_REGION_KEY, None)

                    # 7. 조사 결과 검증
                    processed = result["processed"]
                    records = result["records"]
                    if (
                        processed != len(stores)
                        or len(records) != len(stores)
                    ):
                        raise RuntimeError(
                            f"조사 건수 불일치: "
                            f"{region} / {address} "
                            f"(대상 {len(stores)}, "
                            f"처리 {processed}, "
                            f"로그 {len(records)})"
                        )
                    technical_failures = [rec for rec in records if "TECHNICAL_FAIL" in str(rec)]
                    if technical_failures:
                        print("\n기술 오류 발견")
                        print("지역:", region)
                        print("주소지:", address)
                        print("기술 오류 건수:", len(technical_failures))
                        raise RuntimeError(f"기술 오류 발생 - " f"주소지 미완료: {region} / {address}")

                    # 8. 조사로그 기록
                    for rec in records:
                        rec["region"] = region
                        rec["address_index"] = address_index
                    added = append_log_to_sheet(book=run_wb, records=records)
                    if added != len(stores):
                        raise RuntimeError(f"조사로그 기록 건수 불일치: " f"{region} / {address} " f"(예상 {len(stores)}, 기록 {added})")
                    progress_ws.append([region, address, len(stores), "DONE", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

                    # 9. 주소지별 엑셀 저장
                    save_survey_checkpoint(run_wb, RUN_FILE)
                    done[(region, address)] = len(stores)
                    completed += len(stores)
                    newly_completed += len(stores)
                    group_processed += processed
                    group_success += result["success"]
                    group_failed += result["failed"]
                    elapsed = (time.monotonic() - start_time)
                    print("\n===== 주소지 저장 완료 =====")
                    print("지역:", region)
                    print("주소지:", address)
                    print("처리:", processed)
                    print("성공:", result["success"])
                    print("확인/실패:", result["failed"])
                    print("전체 진행:", f"{completed}/{total}")
                    print("진행률:", f"{completed / total * 100:.1f}%")
                    print(f"누적 소요: {elapsed / 60:.1f}분")
                    print("엑셀 저장 완료")
            group_elapsed = (time.monotonic() - group_start)
            group_results[group_no] = {
                "regions": regions,
                "processed": group_processed,
                "success": group_success,
                "failed": group_failed,
                "elapsed": group_elapsed
            }
            print(f"조합 {group_no}/6 완료")
            print("지역:", ", ".join(regions))
            print("처리:", group_processed)
            print("성공:", group_success)
            print("확인/실패:", group_failed)
            print(f"조합 소요: {group_elapsed / 60:.1f}분")
            print("전체 진행:", f"{completed}/{total}")

        # 10. 전체 조사 종료
        total_elapsed = (time.monotonic() - start_time)
        print("전체 자동 조사 종료")
        print("전체 대상처:", total)
        print("누적 완료:", completed)
        print("이번 실행 조사:", newly_completed)
        print("남은 대상처:", total - completed)
        print("블루스택 재시작:", restart_count)
        print(f"이번 실행 소요: " f"{total_elapsed / 60:.1f}분")
        print("최종 결과 파일:", RUN_FILE)
        if completed == total:
            print("\n★ 745개 전체 조사 완료")
        else:
            print("\n※ 아직 미완료 대상처가 있습니다.")
    finally:
        region_groups.pop(TEMP_REGION_KEY, None)
        wb = original_wb
        ws = original_ws
        run_wb.close()


def run_survey(month, input_file, output_file, mode, target=None,
               new_survey=False, bluestacks_dir=r"C:\Program Files\BlueStacks_nxt",
               preview_only=True):
    """Run survey with a single configuration cell. Preview mode never writes or connects."""
    if type(month) is not int or month not in (9, 10, 11, 12):
        raise ValueError("기존 Excel 열 배치에서 9~12월만 검증 가능한 범위입니다. 다른 월은 열 구조 확인 후 지원해야 합니다.")
    if type(preview_only) is not bool or type(new_survey) is not bool:
        raise ValueError("PREVIEW_ONLY와 NEW_SURVEY는 True 또는 False로 지정하세요.")
    input_path = Path(input_file).expanduser().resolve()
    output_path = Path(output_file).expanduser().resolve()
    if input_path == output_path:
        raise ValueError("INPUT_FILE과 OUTPUT_FILE은 다른 파일이어야 합니다.")
    if not input_path.is_file():
        raise FileNotFoundError(f"원본 Excel 파일을 찾을 수 없습니다: {input_path}")
    if not output_path.parent.is_dir():
        raise FileNotFoundError(f"결과 저장 폴더가 존재하지 않습니다: {output_path.parent}")
    if input_path.suffix.lower() != '.xlsx' or output_path.suffix.lower() != '.xlsx':
        raise ValueError("원본·결과 파일은 .xlsx 형식이어야 합니다.")
    is_full = mode in ("전체", "조합")
    full_path = output_path.with_name(output_path.stem + "_자동전체조사.xlsx")
    resurvey_path = full_path if mode in ("행", "대상처") and full_path.exists() and not new_survey else output_path
    result_path = full_path if is_full else resurvey_path
    if result_path.resolve() == input_path:
        raise ValueError("실제 결과 파일이 원본 Excel과 같습니다.")
    review_path = input_path if new_survey or not output_path.exists() else output_path
    if mode in ("행", "대상처") and resurvey_path.exists() and not new_survey:
        review_path = resurvey_path
    sample = load_workbook(review_path, read_only=True, data_only=False)
    try:
        if "조사자료(rawdata)" not in sample.sheetnames:
            raise ValueError("'조사자료(rawdata)' 시트를 찾을 수 없습니다.")
        chosen, numbers = _read_targets(mode, target, sample)
    finally:
        sample.close()
    print(f"조사 월: {month}월 | 모드: {mode} | 선택 대상처: {len(chosen)}개")
    print("원본 파일:", input_path)
    print("결과 파일:", result_path)
    if is_full:
        print("선택 조합:", numbers)
    elif len(chosen) <= 12:
        for row, (region, address, name) in chosen.items():
            print(f"  {row}행 | {region} | {address} | {name}")
    if not new_survey and result_path.exists():
        current_book = load_workbook(result_path, read_only=True, data_only=True)
        try:
            if "조사로그" in current_book.sheetnames:
                log = current_book["조사로그"]
                headers = next(log.iter_rows(min_row=1, max_row=1, values_only=True), ())
                if "조사월" in headers:
                    index = headers.index("조사월")
                    recorded = {str(row[index]).strip() for row in log.iter_rows(min_row=2, values_only=True) if len(row) > index and row[index] is not None}
                    if recorded and recorded != {str(month)}:
                        raise ValueError(f"기존 결과 파일의 조사월 {sorted(recorded)}과 이번 조사월 {month}이 다릅니다. 결과 파일명을 바꾸거나 NEW_SURVEY를 확인하세요.")
        finally:
            current_book.close()
    print("새 조사 시작:", new_survey, "| 미리보기:", preview_only)
    if preview_only:
        print("미리보기 완료. 실제 조사를 시작하려면 PREVIEW_ONLY=False로 변경하세요.")
        return {"output_file": str(result_path), "selected_rows": list(chosen), "preview": True}
    if is_full and not Path(bluestacks_dir).exists():
        raise FileNotFoundError(f"BlueStacks 설치 경로를 확인하세요: {bluestacks_dir}")
    reset_full_file = new_survey and is_full and full_path.exists()
    if reset_full_file:
        answer = input(f"기존 전체조사 결과를 백업하고 새로 시작합니다: {full_path}  계속하려면 '초기화' 입력: ")
        if answer.strip() != "초기화":
            raise RuntimeError("새 전체조사 초기화가 취소됐습니다.")
        backup = full_path.with_name(full_path.stem + "_초기화전백업_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx")
        copy2(full_path, backup)
        print("이전 전체조사 결과 백업:", backup)
    # 연결에 실패한 경우, 결과 Excel을 새로 만들거나 갱신하기 전에 중단합니다.
    _connect(bluestacks_dir)
    if mode in ("행", "대상처") and resurvey_path == full_path and not new_survey:
        # Reuse the existing full result workbook for direct row re-survey.
        _setup_data(month, input_path, full_path, False)
    else:
        _setup_data(month, input_path, output_path, new_survey)
    actual = _selection_groups(chosen)
    print("실행 지역:", ", ".join(actual))
    if is_full:
        if reset_full_file:
            full_path.unlink()
        _run_group_survey(numbers)
        return {"output_file": str(full_path), "selected_rows": list(chosen), "preview": False}
    return _run_selection(chosen, result_path)
