"""Streamlit 조사 실행기.

별도 Python 프로세스에서 기존 baemin_survey.run_survey()를 호출한다.
결과와 진행 상태는 로컬 파일에 저장하며 기존 엔진 코드는 수정하지 않는다.
"""
from __future__ import annotations

import json
import os
import re
import sys
import time
import uuid
import traceback
from datetime import datetime
from pathlib import Path


def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def save_json(path: Path, obj: dict) -> None:
    """Write status atomically; tolerate transient Windows file-scanner locks."""
    content = json.dumps(obj, ensure_ascii=False, indent=2)
    temporary = path.with_name(f"{path.stem}.{uuid.uuid4().hex}.writing")
    try:
        temporary.write_text(content, encoding="utf-8")
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


GROUP_REGIONS = {
    1: {"경기"}, 2: {"서울", "강원", "인천"},
    3: {"경남", "전남", "제주"}, 4: {"경북", "충북", "세종"},
    5: {"부산", "광주", "대구", "울산"}, 6: {"전북", "충남", "대전"},
}


def previous_group_counts(config: dict) -> tuple[int, int]:
    """Read DONE records: (selected completed, completed outside chosen groups)."""
    if config.get("mode") not in ("조합", "전체"):
        return 0, 0
    output = Path(config.get("expected_result_file") or "")
    if not output.is_file():
        return 0, 0
    from openpyxl import load_workbook
    book = load_workbook(output, read_only=True, data_only=True)
    try:
        if "자동조사진행" not in book.sheetnames:
            return 0, 0
        selected_groups = config.get("target") or []
        selected = (set().union(*(GROUP_REGIONS.get(int(g), set()) for g in selected_groups))
                    if config.get("mode") == "조합" else set().union(*GROUP_REGIONS.values()))
        selected_done = outside_done = 0
        sheet = book["자동조사진행"]
        for record in sheet.iter_rows(min_row=2, max_col=4, values_only=True):
            if len(record) < 4 or record[3] != "DONE":
                continue
            count = int(record[2] or 0)
            if str(record[0]) in selected:
                selected_done += count
            else:
                outside_done += count
        return selected_done, outside_done
    finally:
        book.close()


PROGRESS = re.compile(r"\[전체\s+(\d+)/(\d+)\s*\|\s*([\d.]+)%\]")
CHECKPOINT = re.compile(r"전체 진행:\s*(\d+)/(\d+)")
PARTIAL_SAVE = re.compile(r"주소지 저장 완료:.*?\((\d+)/(\d+)\)")
PREVIOUS_DONE = re.compile(r"이전 완료 대상처:\s*(\d+)")


class StopRequested(BaseException):
    """Excel 주소지 저장 직후에만 발생시키는 안전 중단 신호."""


class LiveLog:
    """print()를 실시간 파일 기록하고, 엔진이 출력하는 진행 수치를 추출."""

    def __init__(self, job_dir: Path, status: dict) -> None:
        self.job_dir = job_dir
        self.status = status
        self.buffer = ""
        self.out = (job_dir / "survey.log").open("a", encoding="utf-8", buffering=1)
        self.outside_done = int(status.get("outside_done", 0))

    def _set_progress(self, global_count: int, engine_total: int) -> None:
        selected_total = int(self.status.get("selected_total") or 0)
        if self.status.get("mode") in ("조합", "전체"):
            count = max(0, global_count - self.outside_done)
            total = selected_total
        else:
            count = global_count
            total = selected_total or engine_total
        count = min(count, total) if total else count
        self.status["processed"] = count
        self.status["total"] = total
        self.status["progress"] = min(1.0, count / max(total, 1))

    def write(self, data: str) -> int:
        data = str(data)
        self.buffer += data
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            self.record(line)
        return len(data)

    def record(self, line: str) -> None:
        self.out.write(line + "\n")
        changed = False
        match = PROGRESS.search(line)
        if match:
            count, total = int(match.group(1)), int(match.group(2))
            self._set_progress(count, total)
            changed = True

        match = CHECKPOINT.search(line)
        if match:
            count, total = int(match.group(1)), int(match.group(2))
            self._set_progress(count, total)
            changed = True

        match = PARTIAL_SAVE.search(line)
        if match:
            self.status["saved"] = int(match.group(1))
            changed = True

        match = PREVIOUS_DONE.search(line)
        if match:
            self.status["saved"] = int(match.group(1))
            changed = True

        # 전체·조합 조사: 이전에 출력된 '전체 진행' 수치가 주소지 저장 직후 확정됨.
        if line.strip() == "엑셀 저장 완료":
            self.status["saved"] = self.status["processed"]
            changed = True

        if line.startswith(("[전체 ", "지역:", "주소:", "조사 종료:", "전체 자동 조사 종료")):
            self.status["last_message"] = line[:250]
            changed = True

        if changed:
            self.status["updated_at"] = timestamp()
            try:
                save_json(self.job_dir / "status.json", self.status)
            except PermissionError as exc:
                # Progress UI may temporarily be stale; never abort a real survey over JSON.
                self.out.write(f"[상태 파일 잠금] 진행률 저장은 다음 갱신에서 재시도: {exc}\n")

        # 실제 조사 엔진이 Excel 저장을 완료한 이후 출력하는 로그만 안전 중단 지점으로 사용.
        # 1) 행/지역/대상처/주소/여러지역: '주소지 저장 완료: ...'
        # 2) 전체/조합: 저장이 끝난 뒤 '엑셀 저장 완료'
        saved_checkpoint = (
            line.startswith("주소지 저장 완료:")
            or line.strip() == "엑셀 저장 완료"
        )
        if saved_checkpoint and (self.job_dir / "stop.request").is_file():
            raise StopRequested()

    def flush(self) -> None:
        self.out.flush()

    def isatty(self) -> bool:
        return False

    @property
    def encoding(self) -> str:
        return "utf-8"

    def close(self) -> None:
        if self.buffer:
            self.record(self.buffer)
            self.buffer = ""
        self.out.close()


def main(config_path: Path) -> None:
    job_dir = config_path.parent
    config = json.loads(config_path.read_text(encoding="utf-8"))
    try:
        selected_done, outside_done = previous_group_counts(config)
    except Exception:
        selected_done = outside_done = 0
    selected_total = int(config["selected_total"])
    status = {
        "state": "running",
        "mode": config.get("mode"),
        "outside_done": outside_done,
        "pid": os.getpid(),
        "started_at": timestamp(),
        "updated_at": timestamp(),
        "selected_total": config["selected_total"],
        "total": config["selected_total"],
        "processed": min(selected_total, selected_done),
        "saved": min(selected_total, selected_done),
        "progress": min(1.0, selected_done / max(selected_total, 1)),
        "last_message": "BlueStacks 연결 준비 중...",
        "output_file": config["expected_result_file"],
        "error": "",
    }
    save_json(job_dir / "status.json", status)
    log = LiveLog(job_dir, status)
    standard_out, standard_err = sys.stdout, sys.stderr
    sys.stdout, sys.stderr = log, log

    try:
        from baemin_survey import run_survey

        print("===== Streamlit 자동조사 시작 =====")
        print(f"조사월: {config['month']}월 / 범위: {config['mode']}")
        print(f"선택 대상처: {config['selected_total']}개")
        print("실행 방식: 기존 결과 유지 / 새 조사에서는 고유 파일 생성")
        result = run_survey(
            month=config["month"],
            input_file=config["input_file"],
            output_file=config["output_file"],
            mode=config["mode"],
            target=config["target"],
            new_survey=False,  # 초기화 입력을 요구하지 않도록 고유 파일을 사용.
            bluestacks_dir=config["bluestacks_dir"],
            preview_only=False,
        )
        status["state"] = "completed"
        status["processed"] = selected_total
        status["total"] = selected_total
        status["progress"] = 1.0
        status["output_file"] = result.get("output_file", config["expected_result_file"])
        status["last_message"] = "조사 정상 종료. 결과 Excel을 확인하세요."
        print("===== 조사 정상 종료 =====")
    except StopRequested:
        status["state"] = "interrupted"
        status["error"] = ""
        status["last_message"] = "사용자 요청으로 안전 중단했습니다. 저장된 Excel을 확인하세요."
        print("===== 사용자 요청으로 안전 중단 =====")
    except BaseException:
        status["state"] = "failed"
        status["error"] = traceback.format_exc()
        status["last_message"] = "조사 중 오류 발생. 저장 파일과 상세 로그를 확인하세요."
        print("===== 조사 오류 =====")
        traceback.print_exc()
    finally:
        status["updated_at"] = timestamp()
        status["ended_at"] = timestamp()
        # Closing the log after restoring stdout prevents an unhandled stop from
        # being raised during the final log flush.
        sys.stdout, sys.stderr = standard_out, standard_err
        try:
            log.close()
        finally:
            save_json(job_dir / "status.json", status)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("사용법: python streamlit_worker.py <config.json>")
    main(Path(sys.argv[1]).resolve())
