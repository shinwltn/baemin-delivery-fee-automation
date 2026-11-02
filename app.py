"""배달의민족 외식배달비 자동조사 Streamlit UI (3차 개선판).

새 조사: 고유 작업 ID와 결과 Excel을 자동 생성해 이전 결과를 덮어쓰지 않는다.
이어하기: 저장된 조사 파일을 재사용한다. 전체/조합은 완료 주소지 건너뛰기를 지원한다.
핵심 자동화 함수 baemin_survey.py는 수정하지 않는다.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import uuid
from contextlib import redirect_stdout
from datetime import datetime, timedelta
from pathlib import Path
import base64
from html import escape

import streamlit as st
from streamlit.components.v1 import html as component_html
from openpyxl import load_workbook
from baemin_survey import run_survey

BASE = Path(__file__).resolve().parent

ASSET_DIRS = (BASE.parent / "배민_ui_png", BASE / "배민_ui_png", BASE / "assets")
ASSETS = next((p for p in ASSET_DIRS if p.is_dir()), BASE / "assets")

IMAGE_FILES = {
    "hero_mascot.webp": "썸네일 배민.png",
    "idle_pout.webp": "후드뒤집어쓴 배민.png",
    "start_glasses.webp": "안경척 배민.png",
    "finish_happy.webp": "기쁨.png",
    "angry.png": "고민 배민.png",
    "sad.png": "고민 배민.png",
    "download.png": "확성기 배민.png",
    "log_analysis.webp": "분석중 배민.png",
    "log_worry.webp": "고민 배민.png",
    "sidebar_logo.png":"배민로고.png",

    "scooter_start_right.webp": "정면 라이더.png",
    "scooter_side_right.webp": "측면 라이더.png",
    "scooter_left_right.webp": "좌측 라이더.png",
    "scooter_basic_right.webp": "기본 라이더.png",
    "scooter_cat_right.webp": "고양이 라이더.png",
    "scooter_sprint_right.webp": "빨간 두건 라이더.png",
}

RIGHT_FACING = {
    "scooter_side_right.webp",
    "scooter_left_right.webp",
    "scooter_basic_right.webp",
    "scooter_cat_right.webp",
    "scooter_sprint_right.webp",
}

DATA_DIR = BASE / "data" / "streamlit_inputs"
RUNS_DIR = BASE / "data" / "streamlit_runs"
OUTPUT_DIR = BASE / "output" / "streamlit_runs"
JOBS_DIR = BASE / "logs" / "streamlit_jobs"
ACTIVE_FILE = JOBS_DIR / "active.json"
REGIONS = [
    "경기", "서울", "강원", "인천", "경남", "전남", "제주",
    "경북", "충북", "세종", "부산", "광주", "대구",
    "울산", "전북", "충남", "대전",
]
GROUPS = {
    1: "경기", 2: "서울·강원·인천", 3: "경남·전남·제주",
    4: "경북·충북·세종", 5: "부산·광주·대구·울산", 6: "전북·충남·대전",
}
for folder in (DATA_DIR, RUNS_DIR, OUTPUT_DIR, JOBS_DIR):
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="배민 외식배달비 자동조사", page_icon="📋", layout="wide",initial_sidebar_state="expanded")


def read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, ValueError):
        return {}


def save_json(path: Path, data: dict) -> None:
    """Atomic write with short retry for transient Windows file locks."""
    temporary = path.with_name(f"{path.stem}.{uuid.uuid4().hex}.writing")
    try:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        for attempt in range(8):
            try:
                os.replace(temporary, path)
                return
            except PermissionError:
                if attempt == 7:
                    raise
                time.sleep(.12 * (attempt + 1))
    finally:
        temporary.unlink(missing_ok=True)


def process_alive(pid: int) -> bool:
    """Windows에서 자식 프로세스 생존 확인."""
    if pid <= 0 or os.name != "nt":
        return False
    try:
        import ctypes
        kernel = ctypes.windll.kernel32
        kernel.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
        kernel.OpenProcess.restype = ctypes.c_void_p
        kernel.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
        kernel.GetExitCodeProcess.restype = ctypes.c_int
        kernel.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel.CloseHandle.restype = ctypes.c_int
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            return bool(kernel.GetExitCodeProcess(handle, ctypes.byref(code))) and code.value == 259
        finally:
            kernel.CloseHandle(handle)
    except Exception:
        return False


def active_job() -> tuple[Path | None, dict]:
    """전역 중복 실행 방지. 종료된 작업의 잠금은 안전하게 해제."""
    meta = read_json(ACTIVE_FILE)
    job_id = meta.get("job_id", "")
    if not re.fullmatch(r"job-[a-f0-9]{12}", job_id):
        # 예전 버전에서 남은 비정상 마커는 지우지 않는다.
        return None, {}
    job_dir = JOBS_DIR / job_id
    status = read_json(job_dir / "status.json")
    pid = int(status.get("pid") or meta.get("pid") or 0)
    if process_alive(pid):
        return job_dir, status
    # 프로세스가 시작될 시간을 잠깐 보장한다.
    age = datetime.now().timestamp() - ACTIVE_FILE.stat().st_mtime if ACTIVE_FILE.exists() else 0
    if age < 30 and status.get("state") not in ("completed", "failed", "interrupted"):
        return job_dir, status
    if status.get("state") not in ("completed", "failed", "interrupted"):
        status = status or {}
        status.update(state="interrupted", last_message="실행 프로세스가 종료되었습니다.",
                      error="정상 종료 확인 불가. 마지막 Excel 저장 상태를 확인한 후 이어하기를 사용하세요.")
        save_json(job_dir / "status.json", status)
    ACTIVE_FILE.unlink(missing_ok=True)
    return None, {}


def validated_target(mode: str, raw):
    if mode == "행":
        parts = [s.strip() for s in str(raw).split(",") if s.strip()]
        if not parts or any(not re.fullmatch(r"\d+", s) for s in parts):
            raise ValueError("Excel 행 번호를 입력해 주세요. 예: 8, 9, 10")
        return [int(s) for s in parts]
    if mode in ("여러지역", "조합") and not raw:
        raise ValueError("지역 또는 조합을 1개 이상 선택해 주세요.")
    if mode in ("대상처", "주소") and not str(raw).strip():
        raise ValueError("대상처명 또는 조사 주소를 입력해 주세요.")
    return raw


def preview_signature(*fields) -> str:
    return hashlib.sha256(json.dumps(fields, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()


def display_rows(input_path: Path, selected: list[int]) -> list[dict]:
    """Excel을 한 번만 순차 스캔 (읽기전용 random cell 조회 성능 문제 방지)."""
    if not selected:
        return []
    book = load_workbook(input_path, read_only=True, data_only=True)
    try:
        sheet = book["조사자료(rawdata)"]
        wanted = set(selected)
        rows = {}
        first = min(selected)
        for index, data in enumerate(
            sheet.iter_rows(min_row=first, max_row=max(selected), min_col=6, max_col=9, values_only=True),
            start=first,
        ):
            if index not in wanted:
                continue
            rows[index] = {"Excel 행": index, "지역": data[0], "조사 주소": data[2], "대상처": data[3]}
            if len(rows) == len(wanted):
                break
        return [rows[row] for row in selected]
    finally:
        book.close()


def expected_result_file(output_file: Path, mode: str) -> Path:
    if mode in ("전체", "조합"):
        return output_file.with_name(output_file.stem + "_자동전체조사.xlsx")
    # 새 작업마다 고유 폴더이므로 다른 모드가 남긴 결과와 충돌하지 않는다.
    return output_file


def all_surveys() -> list[dict]:
    results = []
    for path in RUNS_DIR.glob("survey-*.json"):
        item = read_json(path)
        if re.fullmatch(r"survey-[a-f0-9]{12}", item.get("survey_id", "")):
            results.append(item)
    return sorted(results, key=lambda x: x.get("created_at", ""), reverse=True)


def survey_label(record: dict) -> str:
    when = record.get("created_at", "").replace("T", " ")[:16]
    return f"{when} · {record['month']}월 {record['mode']} · {record['selected_total']}개 · {record['survey_id'][-6:]}"


def latest_attempt(survey_id: str) -> Path | None:
    attempts = []
    for job_dir in JOBS_DIR.glob("job-*"):
        config = read_json(job_dir / "config.json")
        if config.get("survey_id") == survey_id:
            attempts.append(job_dir)
    return max(attempts, key=lambda p: p.stat().st_mtime) if attempts else None


def tail_log(path: Path, max_bytes: int = 60000, max_lines: int = 140) -> str:
    if not path.is_file():
        return "(로그 대기 중)"
    with path.open("rb") as f:
        f.seek(0, 2)
        f.seek(max(0, f.tell() - max_bytes))
        content = f.read()
    return "\n".join(content.decode("utf-8", errors="replace").splitlines()[-max_lines:]) or "(로그 대기 중)"


def review_rows(file: Path) -> list[dict]:
    book = load_workbook(file, read_only=True, data_only=True)
    try:
        if "조사로그" not in book.sheetnames:
            return []
        iterator = book["조사로그"].iter_rows(values_only=True)
        headers = next(iterator, ())
        rows = []
        for values in iterator:
            r = dict(zip(headers, values))
            judgement = str(r.get("조사판정") or "")
            status = str(r.get("상태 코드") or "")
            if (any(t in judgement for t in ("확인", "실패", "미조사")) or "REVIEW" in status
                    or status in ("TECHNICAL_FAIL", "REPLACEMENT_NEEDED", "NO_REPLACEMENT_CANDIDATE", "REPLACEMENT_SEARCH_FAIL")):
                rows.append({
                    "엑셀 행": r.get("엑셀행"), "지역": r.get("지역"), "대상처": r.get("대상처"),
                    "조사판정": judgement, "상태": status, "확인사유": r.get("확인사유"),
                    "신규 대상처 추천": r.get("신규 대상처 추천"),
                })
        return rows
    finally:
        book.close()


def launch_job(survey: dict) -> Path:
    if active_job()[0] is not None or ACTIVE_FILE.exists():
        raise RuntimeError("이미 조사 중입니다. 현재 작업이 끝난 후 시작해 주세요.")
    job_id = "job-" + uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    config = {
        **survey,
        "selected_total": survey["selected_total"],
        "expected_result_file": survey["result_file"],
        "bluestacks_dir": survey["bluestacks_dir"],
    }
    save_json(job_dir / "config.json", config)
    marker = {"job_id": job_id, "created_at": datetime.now().isoformat()}
    try:
        with ACTIVE_FILE.open("x", encoding="utf-8") as file:
            json.dump(marker, file, ensure_ascii=False)
    except FileExistsError as exc:
        raise RuntimeError("동시에 다른 조사가 시작되었습니다.") from exc
    try:
        with (job_dir / "launcher.log").open("ab") as launcher:
            process = subprocess.Popen(
                [sys.executable, "-u", str(BASE / "streamlit_worker.py"), str(job_dir / "config.json")],
                cwd=BASE,
                stdin=subprocess.DEVNULL,
                stdout=launcher,
                stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0,
            )
        marker["pid"] = process.pid
        save_json(ACTIVE_FILE, marker)
        return job_dir
    except Exception:
        ACTIVE_FILE.unlink(missing_ok=True)
        raise


def art_path(name: str) -> Path | None:
    filename = IMAGE_FILES.get(name, name)
    return next((p / filename for p in ASSET_DIRS if (p / filename).is_file()), None)

@st.cache_data(show_spinner=False)
def image_uri(name: str, path_string: str, modified_ns: int) -> str:
    """Use supplied local PNGs; bundled fallbacks only when the external image is absent."""
    from PIL import Image, ImageOps
    with Image.open(path_string) as source:
        pic = source.convert('RGBA'); pic.thumbnail((850, 850), Image.Resampling.LANCZOS)
        if name in RIGHT_FACING: pic = ImageOps.mirror(pic)
        if name == 'sidebar_logo.png':
            from PIL import ImageDraw
            for point in ((0, 0), (pic.width-1, 0), (0, pic.height-1), (pic.width-1, pic.height-1)):
                r, g, b, a = pic.getpixel(point)
                if a and min(r, g, b) >= 235: ImageDraw.floodfill(pic, point, (255, 255, 255, 0), thresh=28)
        buffer = io.BytesIO(); pic.save(buffer, format='PNG', optimize=True)
    return 'data:image/png;base64,' + base64.b64encode(buffer.getvalue()).decode('ascii')


def art(name: str) -> str:
    path = art_path(name)
    return image_uri(name, str(path), path.stat().st_mtime_ns) if path else ''


def art_img(name: str, css: str = "", alt: str = "배민 캐릭터") -> str:
    uri = art(name)
    if not uri:
        return ""
    return f'<img src="{uri}" alt="{escape(alt, quote=True)}" style="{css}">'


def render_header(home: bool = True) -> None:
    """Large Home banner, compact brand strip elsewhere; valid HTML keeps cutout PNGs visible."""
    mascot = art_img('hero_mascot.webp','height:142px;max-width:42%;width:auto;object-fit:contain;object-position:right bottom;align-self:end;','배민과 고양이') if home else ''
    inner = (f'<div><div class="bm-kicker">BAEMIN DELIVERY FEE SURVEY</div><h1>배민 외식배달비 자동조사</h1><p>귀엽게, 정확하게! 배달비 조사 업무 도우미</p></div>{mascot}' if home else '<span class="bm-kicker">BAEMIN DELIVERY FEE SURVEY</span>')
    mode = 'bm-home' if home else 'bm-compact'
    header_html=f'<section class="bm-hero {mode}">{inner}</section>' if home else ''
    st.html(f"""<style>
      @import url('https://fonts.googleapis.com/css2?family=Jua&display=swap');
      [data-testid="stAppViewContainer"]{{background:linear-gradient(180deg,#e3f8f6,#f6fcfb 310px,#fff 690px)}}
      [data-testid="stHeader"]{{background:transparent}}
      .block-container{{max-width:1480px!important;padding-top:1.05rem!important;padding-bottom:1.4rem!important}}
      .bm-font,.bm-hero h1{{font-family:'BM JUA','Jua','Malgun Gothic',sans-serif;font-weight:400}}
      .bm-hero{{box-sizing:border-box;display:flex;align-items:center;justify-content:space-between;overflow:hidden;border-radius:18px;background:linear-gradient(110deg,#068e8b,#20b9b0 66%,#80ddd6);box-shadow:0 10px 23px rgba(12,117,115,.12);margin:0 0 15px;color:#fff}}
      .bm-hero.bm-home{{height:159px;padding:14px 25px 0;gap:12px}}
      .bm-hero.bm-compact{{height:48px;padding:0 18px;border-radius:13px;margin-bottom:12px}}
      .bm-hero h1{{font-size:clamp(28px,3.1vw,40px);line-height:1.25;margin:10px 0 5px;color:#fff}}
      .bm-hero p{{font-size:12px;color:#e6fffd;margin:0 0 15px}}
      .bm-kicker{{font-size:11px;font-weight:750;letter-spacing:.12em;color:#e5fffc}}
      .bm-caption{{font-size:12px;color:#718988}}
      .bm-panel{{border:0;background:#fff;border-radius:17px;box-shadow:0 5px 17px rgba(13,108,110,.075);padding:16px 20px}}
      .bm-mini-container{{background:#fff;border:0;border-radius:15px;box-shadow:0 4px 15px rgba(13,108,110,.07);padding:12px 20px 18px;margin:2px 0 12px}}
      .bm-mini-head{{display:flex;align-items:center;justify-content:space-between;gap:12px;font-size:12px;color:#335f5c}}
      .bm-mini-track{{position:relative;height:12px;background:#e2edec;border-radius:20px;margin:51px 17px 28px}}
      .bm-mini-fill{{height:100%;border-radius:20px;background:linear-gradient(90deg,#91e5df,#2ac1bc)}}
      .bm-mini-pin{{position:absolute;top:50%;left:var(--pin-x);transform:translate(-50%,-50%);width:24px;height:24px;background:#fff;border:3px solid #10b5ab;border-radius:50%;box-shadow:0 2px 7px #298f8840;z-index:2}}
      .bm-mini-img{{position:absolute;left:var(--pin-x);bottom:20px;transform:translateX(-50%);width:36px;height:36px;z-index:2}}
      .bm-mini-pct{{position:absolute;left:var(--pin-x);top:18px;transform:translateX(-50%);padding:3px 7px;background:#128f85;border-radius:6px;color:white;font-size:11px;font-weight:750;white-space:nowrap;z-index:2}}
      @media(max-width:1250px){{
    .st-key-bm_home_grid [data-testid="stHorizontalBlock"]{{flex-direction:column!important}}
    .st-key-bm_home_grid [data-testid="stColumn"]{{min-width:0!important;width:100%!important;max-width:100%!important;flex:none!important}}
}}
@media(max-width:1050px){{
    .st-key-bm_settings_layout [data-testid="stHorizontalBlock"]{{flex-direction:column!important}}
    .st-key-bm_settings_layout [data-testid="stColumn"]{{min-width:0!important;width:100%!important;max-width:100%!important;flex:none!important}}
}}
.st-key-start_actual_survey button{{
    background:linear-gradient(110deg,#009a91,#10c8b8)!important;
    color:white!important;font-weight:800!important;min-height:55px!important;
    box-shadow:0 6px 17px #10a79b3d!important
}}
      .bm-big-number{{font-size:clamp(42px,5vw,65px);color:#098e88;font-weight:800}}
      @media(max-width:720px){{.bm-hero.bm-home{{height:124px;padding:10px 15px 0}}.bm-hero img{{height:115px!important}}.bm-hero h1{{font-size:26px}}.bm-panel{{padding:13px}}}}
    </style>{header_html}""")


def format_elapsed(status: dict) -> str:
    """00m 00s until one hour, then 1h 00m 00s; stop at end time."""
    started = status.get('started_at')
    if not started:
        return '00m 00s'
    try:
        start = datetime.fromisoformat(started)
        if status.get('state') in ('completed', 'interrupted', 'failed'):
            end_str = status.get('ended_at') or status.get('updated_at')
            end = datetime.fromisoformat(end_str) if end_str else datetime.now()
        else:
            end = datetime.now()
        seconds = max(0, int((end-start).total_seconds()))
    except (ValueError, TypeError):
        return '00m 00s'
    hours, rem = divmod(seconds, 3600)
    minutes, seconds = divmod(rem, 60)
    return f'{hours}h {minutes:02d}m {seconds:02d}s' if hours else f'{minutes:02d}m {seconds:02d}s'


def job_progress(status: dict, config: dict) -> tuple[int,int,float]:
    """Prefer the selected scope; never present nationwide counts for partial group runs."""
    total = int(config.get('selected_total') or status.get('selected_total') or 0)
    processed = max(0, int(status.get('processed') or 0))
    if status.get('state') == 'completed':
        processed = total
    processed = min(processed, total) if total else processed
    pct = processed/total * 100 if total else 0.0
    # 100% means the worker has actually returned after the final Excel save.
    if status.get('state') in ('starting','running'):
        pct = min(pct, 99.9)
    return total, processed, max(0.0, min(100.0, pct))


def progress_stage(pct: float, state: str) -> tuple[str,str,str]:
    """User-defined art sequence. The head-on scooter is stationary at 0%."""
    if state == 'completed':
        return '조사 완료', '모든 대상처 조사 종료!', 'finish_happy.webp'
    if state in ('failed','interrupted'):
        return '조사 중단', '저장된 파일과 로그를 확인해 주세요.', 'log_worry.webp'
    if pct < 15:
        return '출발 준비', '정면 라이더 · 0~14%', 'scooter_start_right.webp'
    if pct < 30:
        return '조사 진행 중', '측면 라이더 · 15~29%', 'scooter_side_right.webp'
    if pct < 45:
        return '차근차근 조사 중', '좌측 라이더 · 30~44%', 'scooter_left_right.webp'
    if pct < 60:
        return '배달비 수집 중', '기본 라이더 · 45~59%', 'scooter_basic_right.webp'
    if pct < 75:
        return '고양이와 조사 중', '고양이 라이더 · 60~74%', 'scooter_cat_right.webp'
    return '마지막 질주!', '빨간 두건 라이더 · 75~99%', 'scooter_sprint_right.webp'


def render_mini_progress(status: dict,config: dict) -> None:
    _,_,pct=job_progress(status,config)
    pin=max(0.,min(100.,pct))
    happy=art_img('finish_happy.webp','width:36px;height:36px;object-fit:contain;','기쁨 캐릭터')
    st.html(f"""<div style="padding:34px 28px 29px">
      <div style="position:relative;height:11px;background:#203e4b;border-radius:20px">
        <div style="width:{pct:.2f}%;height:100%;background:#2ac1bc;border-radius:20px"></div>
        <div style="position:absolute;left:{pin:.2f}%;top:-27px;transform:translateX(-50%);width:36px;height:36px">{happy}</div>
        <div style="position:absolute;left:{pin:.2f}%;top:18px;transform:translateX(-50%);font-size:12px;font-weight:800;color:#087d79">{pct:.1f}%</div>
      </div>
    </div>""")

def render_progress(total: int, processed: int, saved: int, progress: float, state: str, elapsed: str='00m 00s', started_at: str='', ended_at: str='') -> None:
    pct = max(0.,min(100.,progress*100)); title, sub, name = progress_stage(pct,state)
    rider = art_img(name,'width:70px;height:75px;object-fit:contain;','진행률 배민 라이더')
    badge = '완료' if state=='completed' else ('중단' if state in ('failed','interrupted') else '조사 중')
    start = json.dumps(str(started_at).replace(' ','T')); end = json.dumps(str(ended_at).replace(' ','T'))
    active = 'true' if state in ('starting','running') else 'false'
    page = f"""<!doctype html><html lang="ko"><meta charset="UTF-8"><style>
      *{{box-sizing:border-box}}body{{margin:0;font-family:'Malgun Gothic',sans-serif;color:#193b38}}
      .card{{height:339px;border:0;border-radius:18px;background:#fff;padding:21px 23px;box-shadow:0 5px 17px rgba(13,108,110,.075)}}
      .top,.labels,.foot{{display:flex;align-items:center;justify-content:space-between;gap:12px}}
      .top{{font-size:19px;font-weight:750}}.badge{{font-size:12px;color:#0c8f88;background:#e5f7f4;padding:6px 12px;border-radius:30px}}
      .percent{{font-size:65px;font-weight:850;color:#098e88;text-align:center;letter-spacing:-3px;line-height:1.13;margin-top:14px}}
      .time{{font-size:13px;color:#899c9b;text-align:center;margin-top:10px}}.labels{{margin-top:14px}}
      .lane{{position:relative;height:83px;border-bottom:1px dashed #d5e9e6;margin-top:6px}}
      .track{{height:12px;background:#e1efed;border-radius:20px;overflow:hidden}}.fill{{height:100%;background:linear-gradient(90deg,#62dfd6,#11b6ad);border-radius:20px}}
      .foot{{font-size:12px;color:#718c88;margin-top:12px}}.stage{{font-weight:700}}.sub{{font-size:11px;color:#6a8784}}
    </style><div class="card"><div class="top"><span>실시간 조사 현황</span><span class="badge">{escape(badge)}</span></div>
    <div class="percent">{pct:.1f}<span style="font-size:.43em">%</span></div><div class="time" id="bm-timer">{escape(elapsed)}</div>
    <div class="labels"><span class="stage">{escape(title)}</span><span class="sub">{escape(sub)}</span></div>
    <div class="lane"><div style="position:absolute;bottom:0;left:clamp(0px,calc({pct:.2f}% - 35px),calc(100% - 70px));width:70px;height:75px">{rider}</div></div>
    <div class="track"><div class="fill" style="width:{pct:.2f}%"></div></div>
    <div class="foot"><span>완료 {processed:,}</span><span>남은 대상 {max(0,total-processed):,}</span><span>전체 {total:,}</span></div></div>
    <script>const t0={start},t1={end},isRun={active};function tick(){{if(!t0)return;const s=Date.parse(t0),e=isRun?Date.now():Date.parse(t1||t0);if(!Number.isFinite(s)||!Number.isFinite(e))return;let n=Math.max(0,Math.floor((e-s)/1000)),h=Math.floor(n/3600),m=Math.floor(n%3600/60),sec=n%60,p=x=>String(x).padStart(2,'0');document.getElementById('bm-timer').textContent=h?`${{h}}h ${{p(m)}}m ${{p(sec)}}s`:`${{p(m)}}m ${{p(sec)}}s`}}tick();setInterval(tick,1000)</script></html>"""
    component_html(page,height=346,scrolling=False)


def render_log_box(log_text: str, height: int=420, scroll_key: str='latest') -> None:
    """Fixed watermark + full scrollable log with saved scroll position."""
    picture = art('log_worry.webp'); key=json.dumps('bm_log_scroll_'+scroll_key)
    page=f"""<!doctype html><html lang="ko"><meta charset="UTF-8"><style>
      *{{box-sizing:border-box}}body{{margin:0;background:transparent;font-family:Consolas,'Malgun Gothic',monospace}}
      #frame{{height:{height-5}px;position:relative;background:#142333;border-radius:13px;overflow:hidden}}
      #watermark{{position:absolute;right:15px;bottom:10px;width:290px;max-width:52%;opacity:.16;pointer-events:none;z-index:0}}
      #scroll{{position:absolute;inset:0;overflow-y:auto;z-index:1;padding:14px 16px;color:#e0ebed}}
      pre{{white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px;line-height:1.58;margin:0;text-shadow:0 0 2px #142333}}
      #latest{{position:absolute;right:18px;top:10px;z-index:3;border:0;border-radius:8px;padding:5px 9px;cursor:pointer;color:#fff;background:#147f79;font-size:11px}}
    </style><div id="frame"><img id="watermark" src="{picture}" alt=""><div id="scroll"><pre>{escape(log_text)}</pre></div><button id="latest">최신으로 ↓</button></div>
    <script>const s=document.getElementById('scroll'),k={key};let v=null;try{{v=JSON.parse(sessionStorage.getItem(k)||'null')}}catch(e){{}}if(v)s.scrollTop=v.follow?s.scrollHeight:v.top;function remember(){{const follow=s.scrollHeight-s.clientHeight-s.scrollTop<35;try{{sessionStorage.setItem(k,JSON.stringify({{top:s.scrollTop,follow}}))}}catch(e){{}}}}s.addEventListener('scroll',remember);document.getElementById('latest').onclick=()=>{{s.scrollTop=s.scrollHeight;remember()}}</script></html>"""
    component_html(page,height=height,scrolling=False)


def render_review_heading(count: int):
    pic = art_img('log_analysis.webp', 'height:166px;width:auto;max-width:38%;object-fit:contain;object-position:right;', '분석중 배민')
    st.html(f"""
      <div class="bm-panel" style="display:flex;justify-content:space-between;align-items:center;gap:8px;padding-top:7px;padding-bottom:7px;">
        <div><div class="bm-font" style="font-size:25px;color:#154c49">👀 사람 확인 대상</div>
          <div style="font-size:29px;font-weight:800;color:#e17d53;margin-top:4px;">{count:,}<span style="font-size:15px">건</span></div>
          <div class="bm-caption">자동으로 확정하지 않은 항목을 검토해 주세요.</div></div>
        {pic}
      </div>
    """)


def render_download_heading():
    image = art_img('download.png','height:136px;max-width:185px;object-fit:contain;', '확성기 배민')
    st.html(f"""
      <div class="bm-panel" style="display:flex;justify-content:space-between;align-items:center;padding-top:8px;padding-bottom:8px;gap:12px;">
        <div><div class="bm-font" style="font-size:21px;color:#185c57">Excel 결과가 준비됐어요!</div>
        <div class="bm-caption">아래에서 Excel 파일을 내려받을 수 있어요.</div></div>{image}
      </div>
    """)


def survey_result_path(job_dir: Path | None, survey: dict | None=None) -> Path | None:
    """Resolve existing Excel from status/config or selected survey; do not invent fallback path."""
    if survey is not None:
        path_str = survey.get('result_file')
    elif job_dir is not None:
        status = read_json(job_dir / 'status.json')
        config = read_json(job_dir / 'config.json')
        path_str = status.get('output_file') or config.get('result_file')
    else:
        path_str = None
    if not path_str:
        return None
    path = Path(path_str)
    return path if path.is_file() else None


@st.cache_data(show_spinner=False)
def cached_review_rows(file_name: str, modified_ns: int) -> list[dict]:
    return review_rows(Path(file_name))


@st.cache_data(show_spinner=False)
def cached_all_result_rows(file_name: str, modified_ns: int) -> list[dict]:
    result = []
    book = load_workbook(file_name, read_only=True, data_only=True)
    try:
        if '조사로그' not in book.sheetnames:
            return []
        iterator = book['조사로그'].iter_rows(values_only=True)
        columns = list(next(iterator, ()))
        for values in iterator:
            if not any(value is not None for value in values):
                continue
            result.append({str(col or f'열{i+1}'): value for i,(col,value) in enumerate(zip(columns,values))})
        return result
    finally:
        book.close()


def review_category(item: dict) -> str:
    code = str(item.get('상태') or '').upper()
    reason = str(item.get('확인사유') or '')
    if 'TECHNICAL' in code or 'TECHNICAL' in reason.upper(): return '기술 실패'
    if 'MULTIPLE_MATCH' in code: return '유사·다중 대상처'
    if 'SAME_STORE' in code: return '중복 대상처'
    if 'REPLACEMENT' in code or '대체' in reason: return '신규·대체 대상처'
    if 'MENU' in code or '메뉴' in reason: return '메뉴 확인'
    if 'ADDRESS' in code or '주소' in reason: return '주소 확인'
    return '기타 확인'


def choose_result_source(job_dir: Path | None, suffix: str) -> tuple[Path | None,str]:
    """Choose latest in-progress run or an archived survey from the sidebar/page."""
    runs = all_surveys()
    choices = ['현재 작업'] if job_dir and job_dir.is_dir() else []
    choices += [r['survey_id'] for r in runs]
    if not choices:
        return None, ''
    lookup = {r['survey_id']: r for r in runs}
    selected = st.selectbox('조회할 조사 선택', choices,
        format_func=lambda x: '현재 작업 (진행 중 포함)' if x=='현재 작업' else survey_label(lookup[x]),
        key=f'result_source_{suffix}', label_visibility='collapsed')
    if selected == '현재 작업':
        return survey_result_path(job_dir), job_dir.name
    return survey_result_path(None, lookup[selected]), selected


def render_stop_control(job_dir: Path | None,state: str):
    if not job_dir or state not in ('running','starting'):
        st.button('전체 조사 안전 중단',icon=':material/stop_circle:',disabled=True,width='stretch',key='stop_disabled')
        return
    stop_file=job_dir/'stop.request'
    if stop_file.exists():
        st.button('중단 요청됨',icon=':material/stop_circle:',disabled=True,width='stretch',key=f'stop_sent_{job_dir.name}')
        st.caption('다음 Excel 저장 완료 후 중단됩니다.')
        return
    if st.button('전체 조사 안전 중단',icon=':material/stop_circle:',width='stretch',key=f'stop_{job_dir.name}'):
        stop_file.write_text('stop',encoding='utf-8')
        st.rerun()



def render_home(job_dir: Path | None):
    status=read_json(job_dir/'status.json') if job_dir else {}
    config=read_json(job_dir/'config.json') if job_dir else {}
    total,processed,pct=job_progress(status,config)
    state=status.get('state','idle')
    result_file=survey_result_path(job_dir) if job_dir else None

    with st.container(key='bm_home_grid'):
        left,right=st.columns([1.7,1],gap='medium')

        with left:
            render_progress(total,processed,int(status.get('saved') or 0),pct/100,state,format_elapsed(status),status.get('started_at',''),status.get('ended_at',''))
            render_stop_control(job_dir,state)
            if not job_dir: st.caption('조사를 시작하려면 Settings에서 조사표와 조건을 선택하세요.')
            if state=='failed': st.error('조사 오류가 발생했어요. 상세 로그를 확인해 주세요.')
            if state=='interrupted': st.warning('조사가 중단됐어요. Settings에서 이어할 수 있어요.')

        with right:
            review_count=0
            if result_file:
                try: review_count=len(cached_review_rows(str(result_file),result_file.stat().st_mtime_ns))
                except Exception: pass

            render_review_heading(review_count)
            if st.button('사람 확인 상세',icon=':material/fact_check:',width='stretch',key='open_review_home'):
                st.session_state['bm_nav_pending']='사람 확인'
                st.rerun()

            if result_file and state in ('completed','failed','interrupted'):
                render_download_heading()
                st.download_button('Excel 결과 다운로드',data=result_file.read_bytes(),file_name=result_file.name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',icon=':material/download:',width='stretch',key='excel_home')
            else:
                st.html('<div class="bm-panel" style="padding:15px;font-size:13px;color:#6a8985">조사가 끝나면 Excel 결과를 받을 수 있어요.</div>')
                st.download_button('Excel 결과 다운로드',data=b'',file_name='result.xlsx',disabled=True,icon=':material/download:',width='stretch',key='excel_home_disabled')

    st.markdown('**최근 실행 로그**')
    log=tail_log(job_dir/'survey.log',max_lines=65,max_bytes=25000) if job_dir else '조사를 시작하면 실행 기록이 표시됩니다.'
    st.html(f'<div style="height:clamp(170px,25vh,310px);overflow-y:auto;border-radius:14px;background:#142333;color:#e0ebed;padding:14px 18px"><pre style="white-space:pre-wrap;overflow-wrap:anywhere;font:12px/1.6 Consolas,monospace;margin:0">{escape(log)}</pre></div>')



def render_review_page(job_dir: Path | None):
    file, _ = choose_result_source(job_dir,'review')
    if file is None:
        render_review_heading(0)
        st.info('아직 저장된 결과 Excel이 없어요. 조사 중이라면 다음 Excel 저장 이후 확인할 수 있어요.')
        return
    try:
        rows = cached_review_rows(str(file),file.stat().st_mtime_ns)
    except Exception as e:
        st.warning(f'Excel 저장 중일 수 있어요. 잠시 후 다시 확인해 주세요: {e}')
        return
    render_review_heading(len(rows))
    if not rows:
        st.success('현재 조사로그에 사람 확인 대상이 없어요.')
        return
    options = ['전체'] + sorted({review_category(r) for r in rows})
    left,right = st.columns([1,2])
    with left: category = st.selectbox('유형별 보기',options)
    with right: search = st.text_input('대상처 검색',placeholder='업체명·지역·확인사유 검색')
    filtered = []
    for row in rows:
        if category != '전체' and review_category(row) != category: continue
        if search and search.casefold() not in ' '.join(str(x) for x in row.values()).casefold(): continue
        filtered.append({'유형':review_category(row),**row})
    st.caption(f'{len(filtered)} / {len(rows)}건 표시')
    st.dataframe(filtered,hide_index=True,height=365,width='stretch')


def render_result_page(job_dir: Path | None):
    st.subheader('📊 조사 결과')
    file, _ = choose_result_source(job_dir,'all')
    if not file:
        st.info('조사 결과 Excel이 저장되면 여기서 바로 조회할 수 있어요.')
        return
    try:
        data = cached_all_result_rows(str(file),file.stat().st_mtime_ns)
    except Exception as exc:
        st.warning(f'현재 저장 중인 파일은 잠시 뒤 다시 열어 주세요: {exc}')
        return
    q = st.text_input('결과 검색',placeholder='대상처명·상태코드·지역 등 검색')
    if q: data = [r for r in data if q.casefold() in ' '.join(str(v) for v in r.values()).casefold()]
    st.caption(f'조사로그 {len(data):,}건 · {file.name}')
    if data: st.dataframe(data,hide_index=True,height=350,width='stretch')
    else: st.info('조회할 조사로그가 아직 없어요.')
    render_download_heading()
    st.html("""
    <style>
      div.st-key-excel_result button {
          background:linear-gradient(115deg,#079E94,#20C7B6)!important;
          color:white!important;
          border:0!important;
          border-radius:12px!important;
          min-height:57px!important;
          font-weight:800!important;
          box-shadow:0 6px 15px rgba(12,147,136,.25)!important;
      }
      div.st-key-excel_result button:hover {
          filter:brightness(1.06);
          box-shadow:0 9px 19px rgba(12,147,136,.32)!important;
      }
    </style>
    """)

    st.download_button(
        'Excel 결과 다운로드',
        data=file.read_bytes(),
        file_name=file.name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary',
        icon=':material/download:',
        width='stretch',
        key='excel_result'
    )


def render_logs_page(job_dir: Path | None):

    st.subheader('상세 실행 기록')

    choices = sorted(
        JOBS_DIR.glob('job-*'),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )

    if not choices:
        st.info('아직 실행 기록이 없어요.')
        return

    ids = [p.name for p in choices]

    default = (
        ids.index(job_dir.name)
        if job_dir and job_dir.name in ids
        else 0
    )

    selected = st.selectbox(
        '실행 기록 선택',
        ids,
        index=default
    )

    folder = JOBS_DIR / selected
    log_file = folder / 'survey.log'

    if st.button(
        '로그 새로고침',
        icon=':material/refresh:'
    ):
        st.rerun()

    full_log = (
        log_file.read_text(
            encoding='utf-8',
            errors='replace'
        )
        if log_file.is_file()
        else '(로그 없음)'
    )

    st.caption(
        f'전체 {len(full_log.splitlines()):,}줄 · '
        '로그창 안에서 스크롤'
    )

    render_log_box(
        full_log,
        height=435,
        scroll_key=selected
    )

    state = read_json(folder / 'status.json')

    if state.get('error'):
        with st.expander('오류 상세 (Traceback)'):
            st.code(
                str(state['error'])[-6500:],
                language='text'
            )


def render_idle():
    img = art_img('idle_pout.webp','height:145px;width:auto;max-width:40%;object-fit:contain;','대기 캐릭터')
    st.html(f"""<section class="bm-panel" style="display:flex;align-items:center;justify-content:space-between;min-height:175px;">
      <div><div class="bm-font" style="font-size:25px;color:#18514e">조사 준비 중이에요!</div>
      <div class="bm-caption">왼쪽 메뉴에서 새 조사나 이어하기를 시작해 주세요.</div></div>{img}</section>""")


def render_settings(running: bool) -> None:
    """Original upload, preview, resume, and start functions; layout only."""

    active_dir, _ = active_job()
    running = active_dir is not None

    # st.form이 아닌 일반 위젯을 사용해야 '지역 → 행' 변경 시 입력창이 즉시 나타난다.
    start_type = st.radio("조사 시작 방식", ["새 조사 시작", "기존 작업 이어하기"], horizontal=True, disabled=running)

    incoming = None
    hash_of_excel = None
    selected_survey = None
    month = None
    mode = None
    target = None
    preview_clicked = False
    source = None
    blue_dir = r"C:\Program Files\BlueStacks_nxt"


    if start_type == "새 조사 시작":
        with st.container(key='bm_settings_layout'):
            upload_col,conditions_col=st.columns([1,1.25],gap='medium')

            with upload_col:
                st.subheader('1. 원본 조사표')
                uploaded=st.file_uploader('원본 Excel 업로드 (.xlsx)',type=['xlsx'],disabled=running)
                if uploaded is not None:
                    incoming=uploaded.getvalue()
                    hash_of_excel=hashlib.sha256(incoming).hexdigest()
                st.caption('결과는 작업별로 저장되며 기존 파일을 덮어쓰지 않습니다.')

            with conditions_col:
                st.subheader('2. 조사 조건')
                year_col,month_col=st.columns(2)
                with year_col: st.text_input('조사 연도',value='2026',disabled=True)
                with month_col: month=st.selectbox('조사 월',[9,10,11,12],index=1,format_func=lambda m:f'{m}월',disabled=running)

                mode=st.selectbox('조사 범위',['지역','행','대상처','주소','여러지역','조합','전체'],disabled=running)
                if mode=='지역': raw_target=st.selectbox('조사 지역',REGIONS,index=9,disabled=running)
                elif mode=='행': raw_target=st.text_input('Excel 행 번호 (쉼표 구분)',value='8',disabled=running)
                elif mode=='대상처': raw_target=st.text_input('대상처명 (정확하게)',disabled=running)
                elif mode=='주소': raw_target=st.text_input('조사표 H열 주소',disabled=running)
                elif mode=='여러지역': raw_target=st.multiselect('지역 여러 개',REGIONS,default=['대전','세종'],disabled=running)
                elif mode=='조합': raw_target=st.multiselect('조사 조합',list(GROUPS),default=[1],format_func=lambda n:f'{n}조합 ({GROUPS[n]})',disabled=running)
                else:
                    raw_target=None
                    st.warning('전체 조사는 745개 대상처 원본에서만 실행됩니다.')

                with st.expander('고급 설정 · BlueStacks 경로'):
                    blue_dir=st.text_input('BlueStacks 설치 폴더',value=blue_dir,disabled=running)

        try: target=validated_target(mode,raw_target)
        except ValueError: target=raw_target
        signature=preview_signature('new',hash_of_excel,month,mode,target,blue_dir)
        preview_clicked=st.button('조사 대상 미리보기',icon=':material/preview:',type='secondary',width='stretch',disabled=running)
    else:
        st.subheader("1. 저장된 작업 선택")
        surveys = all_surveys()
        if not surveys:
            st.info("아직 이 버전에서 실행한 조사가 없습니다. 먼저 '새 조사 시작'으로 작업을 만들어 주세요.")
            selected_survey = None
        else:
            ids = [r["survey_id"] for r in surveys]
            chosen_id = st.selectbox("이어할 작업", ids,
                                     format_func=lambda i: survey_label(next(r for r in surveys if r["survey_id"] == i)),
                                     disabled=running)
            selected_survey = next(r for r in surveys if r["survey_id"] == chosen_id)
            st.caption(f"원본 조사표와 저장된 Excel을 그대로 재사용 · {selected_survey['month']}월 / {selected_survey['mode']}")
            saved_file = Path(selected_survey["result_file"])
            if saved_file.is_file() and not running:
                st.download_button(
                    "📥 기존 조사 결과 다운로드",
                    data=saved_file.read_bytes(),
                    file_name=saved_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
            if selected_survey["mode"] not in ("전체", "조합"):
                st.warning("이 작업은 주소지별 자동 건너뛰기 기능이 없습니다. 다시 실행하면 선택된 업체를 재조사합니다.")
            source = Path(selected_survey["input_file"])
            month = selected_survey["month"]
            mode = selected_survey["mode"]
            target = selected_survey["target"]
            blue_dir = selected_survey["bluestacks_dir"]
        signature = preview_signature("resume", selected_survey["survey_id"] if selected_survey else None)
        preview_clicked = st.button("🔍 이어할 조사 미리보기", type="primary", width="stretch",
                                    disabled=running or selected_survey is None)

    if preview_clicked:
        st.session_state.pop("preview", None)
        try:
            if not (BASE / "streamlit_worker.py").is_file():
                raise FileNotFoundError("streamlit_worker.py가 없습니다. app.py와 같은 폴더에 놓아주세요.")
            if start_type == "새 조사 시작":
                target = validated_target(mode, raw_target)
                if incoming is None:
                    raise ValueError("원본 Excel을 먼저 업로드해 주세요.")
                if len(incoming) > 40 * 1024 * 1024:
                    raise ValueError("Excel 업로드 크기는 최대 40MB입니다.")
                source = DATA_DIR / f"preview_{hash_of_excel}.xlsx"
                if not source.exists():
                    source.write_bytes(incoming)
                # 고유 임시 경로: preview_only=True는 파일을 생성하지 않음.
                temporary_result = OUTPUT_DIR / f"preview_{uuid.uuid4().hex}.xlsx"
                survey = None
                output_file = temporary_result
            else:
                if selected_survey is None:
                    raise ValueError("이어할 작업을 선택해 주세요.")
                if not source.is_file():
                    raise FileNotFoundError("저장된 원본 조사표를 찾을 수 없습니다.")
                survey = selected_survey
                output_file = Path(survey["output_file"])

            with st.spinner("조사 대상 확인 중..."):
                # 엔진 자체의 대상 검증을 사용하되 실조사는 실행하지 않음.
                with redirect_stdout(io.StringIO()):
                    result = run_survey(
                        month=month, input_file=str(source), output_file=str(output_file),
                        mode=mode, target=target, new_survey=False, preview_only=True,
                    )
                display_file = source
                # 행/대상처 이어하기는 엔진이 기존 결과를 우선 읽는 경우가 있어 표시 데이터도 일치시킴.
                if survey:
                    full_path = Path(survey["result_file"])
                    if mode in ("행", "대상처") and full_path.exists():
                        display_file = full_path
                    elif output_file.exists():
                        display_file = output_file
                rows = display_rows(display_file, result["selected_rows"])

            st.session_state["preview"] = {
                "signature": signature,
                "rows": rows,
                "month": month,
                "mode": mode,
                "target": target,
                "selected_total": len(rows),
                "input_file": str(source),
                "blue_dir": blue_dir,
                "resume_survey": survey,
            }
        except Exception as exc:
            st.error(f"미리보기 실패: {exc}")

    preview = st.session_state.get("preview")
    if preview and preview["signature"] == signature and not running:
        st.divider()
        st.subheader("3. 미리보기 결과")
        p1, p2 = st.columns(2)
        p1.metric("선택한 대상처", f"{preview['selected_total']:,}개")
        p2.metric("조사 월", f"{preview['month']}월")
        st.dataframe(preview["rows"], hide_index=True, width="stretch", height=310)
        st.caption("결과 파일은 작업별 폴더에 자동 저장됩니다. 조사 완료 후 다운로드할 수 있어요.")

        if start_type == "기존 작업 이어하기":
            survey = preview["resume_survey"]
            if Path(survey["result_file"]).exists():
                st.info("기존 Excel에 저장된 결과를 유지하고 이어서 실행합니다.")
            elif latest_attempt(survey["survey_id"]):
                st.warning("이전 작업에서 저장된 결과 파일이 없습니다. 오류 로그를 먼저 확인하세요.")

        confirmation_key = "confirm_" + preview["signature"][:16]
        with st.container():
            start_info, start_art = st.columns([5, 1], vertical_alignment="center")
            with start_info:
                agreed = st.checkbox("조사 대상과 실행 방식을 확인했습니다. 실제 BlueStacks 조사를 시작합니다.", key=confirmation_key)
                run_clicked=st.button("실제 조사 시작",icon=":material/play_arrow:",disabled=not agreed,type="primary",width="stretch",key="start_actual_survey")
            with start_art:
                start_img = art_path("start_glasses.webp")
                if start_img and start_img.exists():
                    st.image(str(start_img), width=140)
        if run_clicked:
            try:
                if preview["signature"] != signature:
                    raise RuntimeError("조사 조건이 변경됐습니다. 미리보기를 다시 해주세요.")
                if not (Path(preview["blue_dir"]) / "HD-Adb.exe").is_file():
                    raise FileNotFoundError("BlueStacks HD-Adb.exe를 찾지 못했습니다. 고급 설정을 확인해 주세요.")
                if start_type == "새 조사 시작":
                    survey_id = "survey-" + uuid.uuid4().hex[:12]
                    run_folder = OUTPUT_DIR / survey_id
                    run_folder.mkdir(parents=True, exist_ok=False)
                    persisted_input = DATA_DIR / f"{survey_id}.xlsx"
                    shutil.copy2(Path(preview["input_file"]), persisted_input)
                    base_output = run_folder / f"2026_{preview['month']:02d}_배달비조사.xlsx"
                    survey = {
                        "survey_id": survey_id,
                        "created_at": datetime.now().isoformat(timespec="seconds"),
                        "month": preview["month"],
                        "mode": preview["mode"],
                        "target": preview["target"],
                        "selected_total": preview["selected_total"],
                        "input_file": str(persisted_input),
                        "output_file": str(base_output),
                        "result_file": str(expected_result_file(base_output, preview["mode"])),
                        "bluestacks_dir": preview["blue_dir"],
                    }
                    save_json(RUNS_DIR / f"{survey_id}.json", survey)
                else:
                    survey = preview["resume_survey"]
                    if not Path(survey["input_file"]).is_file():
                        raise FileNotFoundError("저장된 원본 Excel이 없습니다.")
                current_job = launch_job(survey)
                st.session_state["job_dir"] = str(current_job)
                st.session_state["bm_nav_pending"] = "🛵 조사 현황"
                st.session_state.pop("preview", None)
                st.rerun()
            except Exception as exc:
                st.error(f"조사 시작 실패: {exc}")
    elif preview and preview["signature"] != signature and not running:
        st.info("조사 조건이 바뀌었어. 미리보기를 다시 누르면 실제 조사 버튼이 활성화돼.")


def recent_job_dir(active_dir: Path | None) -> Path | None:
    if active_dir is not None:
        return active_dir
    value = st.session_state.get('job_dir')
    if value and Path(value).is_dir():
        return Path(value)
    paths = list(JOBS_DIR.glob('job-*'))
    return max(paths, key=lambda p:p.stat().st_mtime) if paths else None


# UI navigation. Stored survey/job states and engine calls above remain unchanged.
NAV_MAP={'🛵 조사 현황':'Home','⚙️ 조사 설정':'Settings','📊 조사 결과':'Results','🔎 사람 확인':'사람 확인','📝 상세 로그':'상세 로그'}
if 'bm_nav_pending' in st.session_state:
    pending=st.session_state.pop('bm_nav_pending'); st.session_state['bm_nav']=NAV_MAP.get(pending,pending)
if st.session_state.get('bm_nav') not in ('Home','Settings','Results','사람 확인','상세 로그'): st.session_state['bm_nav']='Home'
page=st.session_state['bm_nav']; CONTACT_EMAIL=''  # Add your email here to enable its link.

if 'bm_side_compact' not in st.session_state: st.session_state['bm_side_compact']=False
small=st.session_state['bm_side_compact']
side_width=74 if small else 244

st.html(f"""<style>
section[data-testid="stSidebar"]{{width:{side_width}px!important;min-width:{side_width}px!important;max-width:{side_width}px!important}}
section[data-testid="stSidebar"],section[data-testid="stSidebar"]>div{{background:#0b2030!important;color:#dce8ee!important}}
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"]{{padding:8px 11px!important}}
section[data-testid="stSidebar"] .stButton button{{display:flex!important;justify-content:flex-start!important;align-items:center!important;gap:11px!important;min-height:41px!important;padding:0 12px!important;border:0!important;border-radius:8px!important;box-shadow:none!important;font-size:13px!important;color:#bbced8!important}}
section[data-testid="stSidebar"] .stButton button[kind="tertiary"]{{background:transparent!important}}
section[data-testid="stSidebar"] .stButton button[kind="primary"]{{background:{'#123244' if small else '#24495f'}!important;color:white!important;box-shadow:inset 3px 0 #36d4c6!important}}
section[data-testid="stSidebar"] .stButton button:hover{{background:#1a3b50!important;color:white!important}}
section[data-testid="stSidebar"] .stButton button p{{text-align:left!important;margin:0!important}}
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"]{{display:none!important}}
.st-key-bm_results_sub{{border-left:1px solid #375267;margin-left:{'0' if small else '15px'};padding-left:{'0' if small else '9px'}}}
.bm-side-logo{{padding:3px 1px 10px;display:flex;align-items:center;gap:10px;justify-content:{'center' if small else 'flex-start'};min-height:{'42px' if small else '64px'}}}
.bm-side-logo img{{max-width:100%;width:{'38px' if small else '54px'};max-height:{'42px' if small else '58px'};object-fit:contain}}.bm-side-brand{{font-size:18px;font-weight:850;letter-spacing:.05em;color:#f3fffe}}
.bm-side-category{{font-size:10px;font-weight:700;letter-spacing:.14em;color:#8ba5b1;margin:13px 8px 7px}}
.bm-side-footer{{position:fixed;bottom:22px;left:{'17px' if small else '24px'};display:flex;gap:19px}}
.bm-side-footer a{{color:#aebfcc;text-decoration:none;font-size:12px;display:flex;align-items:center;gap:6px}}
.bm-side-footer svg{{width:16px;height:16px}}
{'section[data-testid="stSidebar"] .stButton button p,.bm-side-category,.bm-side-footer span{display:none!important}' if small else ''}
</style>""")

with st.sidebar:
    if st.button('펼치기' if small else '접기',icon=':material/keyboard_double_arrow_right:' if small else ':material/keyboard_double_arrow_left:',key='sidebar_toggle',type='tertiary',width='stretch'):
        st.session_state['bm_side_compact']=not small
        st.rerun()

    logo=art('sidebar_logo.png')
    if logo:
        brand_text = '' if small else '<span class="bm-side-brand">BAEMIN</span>'
        st.html(f'<div class="bm-side-logo"><img src="{logo}" alt="배민 로고">{brand_text}</div>')
    else: st.html('<div class="bm-side-logo" style="font-size:18px;font-weight:900;color:#2ac1bc">BAEMIN</div>')

    def nav(label,icon,key,target):
        selected = page==target or (target=='Results' and page in ('상세 로그','사람 확인'))
        if st.button(label,icon=icon,key=key,type='primary' if selected else 'tertiary',width='stretch'):
            if small and target=='Results': st.session_state['bm_side_compact']=False
            else: st.session_state['bm_nav']=target
            st.rerun()

    nav('Home',':material/home:','nav_home','Home')
    nav('Settings',':material/settings:','nav_settings','Settings')
    if not small: st.html('<div class="bm-side-category">RESULTS</div>')
    nav('Results',':material/assessment:','nav_results','Results')

    if not small:
        with st.container(key='bm_results_sub'):
            nav('상세 로그',':material/description:','nav_logs','상세 로그')
            nav('사람 확인',':material/fact_check:','nav_review','사람 확인')

    github='https://github.com/shinwltn/baemin-delivery-fee-automation'
    email=f'<a href="mailto:{escape(CONTACT_EMAIL,quote=True)}" title="Email"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m2 7 10 6 10-6"/></svg><span>Email</span></a>' if CONTACT_EMAIL else ''
    st.html(f'<div class="bm-side-footer"><a href="{github}" target="_blank" title="GitHub"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m8 8-4 4 4 4m8-8 4 4-4 4M14 4l-4 16"/></svg><span>GitHub</span></a>{email}</div>')


render_header(home=page=='Home')
active_dir,_=active_job(); job_dir=recent_job_dir(active_dir); running=active_dir is not None
if page!='Home':
    @st.fragment(run_every='2s')
    def mini_live_progress():
        active,_=active_job(); current=recent_job_dir(active)
        if not current or not current.is_dir(): return
        status=read_json(current/'status.json'); config=read_json(current/'config.json')
        render_mini_progress(status,config)
        if active is None and status.get('state') in ('completed','failed','interrupted') and st.session_state.get('finished_job_rerun')!=current.name:
            st.session_state['finished_job_rerun']=current.name; st.rerun()
    mini_live_progress()
if page=='Home':
    @st.fragment(run_every='2s')
    def home_live():
        active,_=active_job(); current=recent_job_dir(active)
        if current and current.is_dir():
            status=read_json(current/'status.json')
            if active is None and status.get('state') in ('completed','failed','interrupted') and st.session_state.get('finished_job_rerun')!=current.name:
                st.session_state['finished_job_rerun']=current.name; st.rerun()
        render_home(current)
    home_live()
elif page=='Settings': render_settings(running)
else:
    if page=='상세 로그': st.toggle('로그 실시간 자동 갱신',value=False,key='bm_log_live',help='끄면 로그를 읽는 동안 화면이 자동 갱신되지 않습니다.')
    refresh='2s' if page!='상세 로그' or st.session_state.get('bm_log_live',False) else None
    @st.fragment(run_every=refresh)
    def detail_live():
        active,_=active_job(); current=recent_job_dir(active)
        if page=='Results': render_result_page(current)
        elif page=='사람 확인': render_review_page(current)
        elif page=='상세 로그': render_logs_page(current)
    detail_live()

if page!='Home': st.html('<div style="text-align:center;margin:20px 0 8px;color:#87a49f;font-size:10px;letter-spacing:.17em">BAEMIN DELIVERY FEE SURVEY</div>')