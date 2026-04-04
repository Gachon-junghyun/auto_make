"""
TCL 정산 폴더 일괄 PDF 변환 스크립트
구조: TCL_3월 정산 / 지사 / 점 / TCL_xxx.pptx + TCL_xxx.xlsx

출력 구조:
  TCL_3월 정산_PDF / 지사 / 점 / TCL_xxx_사진.pdf    ← pptx 변환
                                  TCL_xxx_견적서.pdf  ← xlsx 변환

사용법:
    python convert_to_pdf.py                         # 기본: 자동 탐지
    python convert_to_pdf.py --folder "TCL_3월 정산" # 폴더 직접 지정
    python convert_to_pdf.py --type pptx             # pptx만 변환
    python convert_to_pdf.py --type xlsx             # xlsx만 변환
    python convert_to_pdf.py --dry-run               # 대상 목록만 출력 (실제 변환 없음)
"""

import subprocess
import sys
import argparse
import shutil
import platform
from pathlib import Path


# ──────────────────────────────────────────────
# 1. LibreOffice 실행 경로 확인 (Windows/Mac/Linux 공통)
# ──────────────────────────────────────────────

def find_libreoffice() -> str | None:
    for cmd in ["libreoffice", "soffice"]:
        if shutil.which(cmd):
            return cmd

    if platform.system() == "Windows":
        for path in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]:
            if Path(path).exists():
                return path

    elif platform.system() == "Darwin":
        mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if Path(mac).exists():
            return mac

    return None


def check_libreoffice() -> str:
    lo = find_libreoffice()
    if lo:
        print(f"✅ LibreOffice 발견: {lo}")
        return lo

    print("❌ LibreOffice를 찾을 수 없습니다.")
    if platform.system() == "Windows":
        print("👉 https://www.libreoffice.org/download/libreoffice/ 에서 설치 후 다시 실행하세요.")
    elif platform.system() == "Darwin":
        print("👉 brew install --cask libreoffice")
    else:
        print("👉 sudo apt-get install -y libreoffice")
    sys.exit(1)


# ──────────────────────────────────────────────
# 2. 파일 유형별 PDF 이름 접미사
# ──────────────────────────────────────────────

SUFFIX_MAP = {
    "pptx": "_사진",      # TCL_가야점.pptx  →  TCL_가야점_사진.pdf
    "xlsx": "_견적서",    # TCL_가야점.xlsx  →  TCL_가야점_견적서.pdf
}


# ──────────────────────────────────────────────
# 3. 폴더 구조 탐색: 지사 > 점 > 파일
# ──────────────────────────────────────────────

def collect_files(base_folder: Path, file_types: list) -> list[dict]:
    """
    반환: [{'지사': str, '점': str, 'file': Path, 'type': 'pptx'|'xlsx'}, ...]
    루트에 있는 파일도 포함 (지사='(루트)', 점='(루트)')
    """
    targets = []

    for item in sorted(base_folder.iterdir()):
        # 루트 직속 파일
        if item.is_file() and item.suffix.lower().lstrip('.') in file_types:
            targets.append({'지사': '(루트)', '점': '(루트)', 'file': item,
                            'type': item.suffix.lower().lstrip('.')})

        # 지사 폴더
        elif item.is_dir():
            for 점_dir in sorted(item.iterdir()):
                if 점_dir.is_dir():
                    for f in sorted(점_dir.iterdir()):
                        ext = f.suffix.lower().lstrip('.')
                        if f.is_file() and ext in file_types:
                            targets.append({'지사': item.name, '점': 점_dir.name,
                                            'file': f, 'type': ext})

    return targets


# ──────────────────────────────────────────────
# 4. 출력 경로 계산
#    원본:  TCL_3월 정산 / 지사 / 점 / TCL_xxx.pptx
#    출력:  TCL_3월 정산_PDF / 지사 / 점 / TCL_xxx_사진.pdf
# ──────────────────────────────────────────────

def get_output_path(file_path: Path, base_folder: Path, pdf_root: Path) -> Path:
    """원본 경로 기준으로 PDF 출력 경로 계산"""
    rel = file_path.relative_to(base_folder)       # 지사/점/TCL_xxx.pptx
    suffix = SUFFIX_MAP.get(file_path.suffix.lower().lstrip('.'), "")
    pdf_name = file_path.stem + suffix + ".pdf"    # TCL_xxx_사진.pdf
    return pdf_root / rel.parent / pdf_name        # PDF루트/지사/점/TCL_xxx_사진.pdf


# ──────────────────────────────────────────────
# 5. 단일 파일 변환
# ──────────────────────────────────────────────

def convert_file(file_path: Path, out_pdf_path: Path) -> bool:
    """
    file_path → out_pdf_path 로 변환.
    LibreOffice는 --outdir 에 stem.pdf 로 저장하므로,
    변환 후 원하는 이름으로 rename 처리.
    """
    out_dir = out_pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice가 만들어낼 임시 PDF 이름 (stem 그대로)
    temp_pdf = out_dir / (file_path.stem + ".pdf")

    lo_exec = find_libreoffice() or "libreoffice"
    cmd = [lo_exec, "--headless", "--norestore",
           "--convert-to", "pdf",
           "--outdir", str(out_dir),
           str(file_path)]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    # 임시 파일 정리
    for tmp in out_dir.glob("lu*.tmp"):
        try:
            tmp.unlink()
        except Exception:
            pass

    # 변환 성공 여부: temp_pdf 존재 확인
    if temp_pdf.exists() and temp_pdf.stat().st_size > 0:
        # _사진 / _견적서 이름으로 rename
        if temp_pdf != out_pdf_path:
            if out_pdf_path.exists():
                out_pdf_path.unlink()
            temp_pdf.rename(out_pdf_path)

        size_kb = out_pdf_path.stat().st_size // 1024
        print(f"✅ {file_path.name}  →  {out_pdf_path.name}  ({size_kb} KB)")
        return True
    else:
        print(f"❌ 실패: {file_path.name}  (exit={result.returncode})")
        if result.stderr.strip():
            print(f"   {result.stderr.strip()[:200]}")
        return False


# ──────────────────────────────────────────────
# 6. 메인
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="TCL 정산 PPTX/XLSX → PDF 일괄 변환")
    parser.add_argument("--folder", default=None, help="원본 루트 폴더 경로")
    parser.add_argument("--type", choices=["pptx", "xlsx", "all"], default="all")
    parser.add_argument("--dry-run", action="store_true",
                        help="변환 없이 대상 목록만 출력")
    args = parser.parse_args()

    # ── 원본 폴더 결정
    if args.folder:
        base = Path(args.folder)
    else:
        script_dir = Path(__file__).parent
        candidates = [d for d in script_dir.iterdir()
                      if d.is_dir() and ("TCL" in d.name or "정산" in d.name)
                      and not d.name.endswith("_PDF")]
        if not candidates:
            print("❌ 변환할 폴더를 찾지 못했습니다. --folder 옵션으로 지정하세요.")
            sys.exit(1)
        base = candidates[0]
        print(f"📂 자동 탐지: {base}")

    if not base.exists():
        print(f"❌ 폴더가 존재하지 않습니다: {base}")
        sys.exit(1)

    # ── PDF 출력 루트 (원본 폴더명 + "_PDF")
    pdf_root = base.parent / (base.name + "_PDF")

    # ── 파일 유형
    file_types = ["pptx", "xlsx"] if args.type == "all" else [args.type]

    # ── LibreOffice 확인
    if not args.dry_run:
        check_libreoffice()
        print()

    # ── 파일 수집
    targets = collect_files(base, file_types)
    total = len(targets)

    if total == 0:
        print("⚠️  변환할 파일이 없습니다.")
        sys.exit(0)

    # ── 요약
    tag = "[DRY-RUN] " if args.dry_run else ""
    print(f"{tag}📋 변환 대상: 총 {total}개  →  출력 폴더: {pdf_root.name}")
    print("─" * 55)
    from collections import Counter
    cnt = Counter(t['지사'] for t in targets)
    for 지사, n in cnt.items():
        print(f"  {지사}: {n}개")
    print("─" * 55)
    print()

    if args.dry_run:
        print("파일 목록:")
        for t in targets:
            out = get_output_path(t['file'], base, pdf_root)
            print(f"  [{t['type'].upper()}] {t['지사']} / {t['점']} / {t['file'].name}")
            print(f"         → {out.relative_to(pdf_root)}")
        return

    # ── 변환 실행
    success, fail = 0, 0
    current_지사 = None

    for idx, t in enumerate(targets, 1):
        if t['지사'] != current_지사:
            current_지사 = t['지사']
            print(f"\n📁 {current_지사}")

        out_pdf = get_output_path(t['file'], base, pdf_root)
        print(f"  [{idx:>3}/{total}] {t['점']} / ", end="", flush=True)

        if convert_file(t['file'], out_pdf):
            success += 1
        else:
            fail += 1

    print(f"\n{'=' * 55}")
    print(f"✅ 성공: {success}개   ❌ 실패: {fail}개   (총 {total}개)")
    print(f"📂 저장 위치: {pdf_root}")
    print(f"{'=' * 55}")

    if fail > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
