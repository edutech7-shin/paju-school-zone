"""
와석초 명렬표 일괄 처리 후 결과 엑셀 4종 저장.

사용:
  프로젝트 루트에서:
    python scripts/generate_waseok_outputs.py

환경변수 ADDRESS_CONVERT_API_KEY(.env는 자동 미로드 아님)는 터미널에서 설정하거나
PowerShell에서 .env를 읽어 export 한 뒤 실행하세요.
"""
from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_REPO_ROOT))
os.chdir(_REPO_ROOT)

from app.services import clear_context_cache, process_request


def main() -> None:
    repo_root = _REPO_ROOT

    roster_dir = Path(r"c:\Users\user\Downloads\학생 주소 명렬표")
    files = [roster_dir / f"학생명렬표({g}학년).xlsx" for g in range(1, 7)]
    for p in files:
        if not p.is_file():
            raise FileNotFoundError(f"명렬표 없음: {p}")

    payloads = [(p.name, p.read_bytes()) for p in files]

    clear_context_cache()
    result = process_request(payloads, school_name="와석초")

    out_dir = repo_root / "generated_outputs" / datetime.now().strftime("result_%Y%m%d_%H%M%S")
    out_dir.mkdir(parents=True, exist_ok=True)

    outputs = {
        "통학구역 보고서(와석초).xlsx": result["report_workbook"],
        "우리학교 통학구역별 재학생 현황(와석초).xlsx": result["requested_school_workbook"],
        "학생_처리내역(와석초).xlsx": result["detail_workbook"],
        "학구불일치_미분류_명단(와석초).xlsx": result["issue_workbook"],
    }
    for name, data in outputs.items():
        (out_dir / name).write_bytes(data)

    print("OUT_DIR=" + str(out_dir))
    for name in outputs:
        print("FILE=" + str(out_dir / name))


if __name__ == "__main__":
    main()
