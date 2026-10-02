"""에이전트 최종 보고(세 문서 묶음)를 구분선 기준으로 나눠 파일로 저장한다.

사용법:
    python3 tools/split_report.py <보고서 원문.md> <저장 폴더>

구분선 형식 (에이전트 절차서의 '최종 보고 양식'):
    ===== 문서 1: 검토 의견서 =====
    ===== 문서 2: 감사 조서 메모 =====   (법령 에이전트는 '내부 검토 메모')
    ===== 문서 3: 이메일 축약 의견 =====

저장 결과:
    01-opinion.md   검토 의견서
    02-memo.md      감사 조서 메모 / 내부 검토 메모
    03-email.txt    이메일 축약 의견 (메일 본문에 붙여 넣는 용도)
"""
import re
import sys
from pathlib import Path

MARKER = re.compile(r"^=====\s*문서\s*([123])\s*:\s*(.+?)\s*=====\s*$", re.MULTILINE)
OUTPUT_NAMES = {"1": "01-opinion.md", "2": "02-memo.md", "3": "03-email.txt"}


def split(text: str) -> dict:
    matches = list(MARKER.finditer(text))
    found = [m.group(1) for m in matches]
    if sorted(found) != ["1", "2", "3"]:
        raise SystemExit(f"구분선 3개(문서 1·2·3)를 찾지 못했습니다. 찾은 문서: {found or '없음'}")
    parts = {}
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        parts[m.group(1)] = text[m.end():end].strip() + "\n"
    return parts


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    src, out_dir = Path(sys.argv[1]), Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)
    for key, body in split(src.read_text(encoding="utf-8")).items():
        if key == "3":
            # 메일 본문에 붙여 넣을 수 있도록 남은 강조 표시를 걷어 낸다.
            body = body.replace("**", "")
        path = out_dir / OUTPUT_NAMES[key]
        path.write_text(body, encoding="utf-8")
        print(path)


if __name__ == "__main__":
    main()
