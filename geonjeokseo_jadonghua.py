# -*- coding: utf-8 -*-
"""
수산나디자인 견적서 자동화 프로그램

흐름:
  makeList.xlsx 지점 목록 → 양식 선택(조명/비조명) → 견적서 생성
  견적서/관저점/TCL_관저점.xlsx  ←  지점별 하위 폴더
  견적서/관저점/TCL_관저점.pptx  ←  사진폼 기반 PPT 자동 생성

주요 규칙:
  - A6 (고객명)  : 양식 그대로 유지 (TCL일렉트로닉스코리아)
  - B9 (건명)    : "TCL 로고 제작설치 작업-{지점명}점" 자동 입력
  - A4/B4/C4     : makeList 날짜 있으면 입력, 없으면 양식 그대로
  - 파일명       : TCL_{지점명}점.xlsx / TCL_{지점명}점.pptx  (날짜 없음)
"""

import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import sys
import threading
from datetime import datetime

# ─────────────────────────────────────────────
# 경로 설정
# ─────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

YANGSIK_DIR    = os.path.join(BASE_DIR, "양식")
SAJIN_DIR      = os.path.join(BASE_DIR, "사진")
OUTPUT_DIR     = os.path.join(BASE_DIR, "견적서")
MAKE_LIST_PATH = os.path.join(BASE_DIR, "makeList.xlsx")

# 타입별 양식 파일명 (양식/ 폴더 안)
TEMPLATE_FILES = {
    '조명':  '견적서양식_조명.xlsx',
    '비조명': '견적서양식_비조명.xlsx',
}

# PPT 템플릿 파일명 (양식/ 폴더 안)
PPT_TEMPLATE = '사진폼.pptx'

# 사진 파일 확장자
SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.bmp')


# ─────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────
def branch_full(branch):
    """지점명 끝에 '점' 없으면 붙임 (관저 → 관저점, 관저점 → 관저점)"""
    return branch if branch.endswith('점') else branch + '점'


def get_branches():
    """사진 폴더 안의 하위 폴더를 지점 목록으로 인식 (리스트 업데이트용)"""
    if not os.path.exists(SAJIN_DIR):
        return []
    return sorted([
        d for d in os.listdir(SAJIN_DIR)
        if os.path.isdir(os.path.join(SAJIN_DIR, d)) and not d.startswith('.')
    ])


def get_branch_photo_dir(branch):
    """
    사진 폴더에서 지점명에 맞는 폴더 경로 반환
    makeList 지점명과 폴더명이 '점' 유무로 다를 수 있어 양쪽 모두 시도
    반환: (폴더경로, 실제폴더명) 또는 (None, None)
    """
    candidates = [branch]
    if branch.endswith('점'):
        candidates.append(branch[:-1])      # 관저점 → 관저
    else:
        candidates.append(branch + '점')    # 관저  → 관저점

    for name in candidates:
        path = os.path.join(SAJIN_DIR, name)
        if os.path.isdir(path):
            return path, name
    return None, None


def get_template_path(light_type):
    """타입(조명/비조명)에 맞는 양식 파일 경로 반환"""
    fname = TEMPLATE_FILES.get(light_type, TEMPLATE_FILES['조명'])
    path  = os.path.join(YANGSIK_DIR, fname)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"양식 파일을 찾을 수 없습니다: {fname}\n"
            f"경로: {path}\n\n"
            f"양식 폴더에 '{fname}' 파일을 넣어주세요."
        )
    return path


# ─────────────────────────────────────────────
# Excel ZIP 직접 수정 (border 보존용)
# ─────────────────────────────────────────────
def _save_excel_zip(template_path, out_path, updates):
    """
    openpyxl 대신 ZIP 직접 조작으로 xlsx 저장.
    styles.xml을 건드리지 않아 border가 보존됨.

    updates: {'A4': '2025년', 'B4': '1월', 'C4': '15일', 'B9': 'TCL 로고 제작설치 작업-관저점'}
             None 값인 키는 건너뜀 (양식 그대로 유지)
    """
    import zipfile, io
    from lxml import etree

    SS_NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    updates = {k: v for k, v in updates.items() if v is not None}

    buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as z_in:
        names = z_in.namelist()

        # ── 워크북에서 "견적서" 시트 → XML 파일 경로 결정 ──
        with z_in.open("xl/workbook.xml") as f:
            wb_tree = etree.parse(f)
        sheet_el = next(
            s for s in wb_tree.findall(f'.//{{{SS_NS}}}sheet')
            if s.get('name') == '견적서'
        )
        r_id = sheet_el.get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        )
        with z_in.open("xl/_rels/workbook.xml.rels") as f:
            rels_tree = etree.parse(f)
        rel = next(
            r for r in rels_tree.findall(f'.//{{{REL_NS}}}Relationship')
            if r.get('Id') == r_id
        )
        sheet_xml_path = "xl/" + rel.get('Target').lstrip('/')

        # ── sharedStrings.xml 읽기 / 업데이트 ──
        with z_in.open("xl/sharedStrings.xml") as f:
            ss_tree = etree.parse(f).getroot()

        def _get_si_text(si):
            return ''.join((t.text or '') for t in si.findall(f'.//{{{SS_NS}}}t'))

        sis      = ss_tree.findall(f'{{{SS_NS}}}si')
        ss_texts = [_get_si_text(si) for si in sis]

        def _get_or_add(value):
            if value in ss_texts:
                return ss_texts.index(value)
            new_si  = etree.SubElement(ss_tree, f'{{{SS_NS}}}si')
            t_elem  = etree.SubElement(new_si, f'{{{SS_NS}}}t')
            t_elem.text = value
            ss_texts.append(value)
            return len(ss_texts) - 1

        cell_to_ss = {ref: _get_or_add(val) for ref, val in updates.items()}

        # sharedStrings count 갱신
        n = len(ss_tree.findall(f'{{{SS_NS}}}si'))
        ss_tree.set('count', str(n))
        ss_tree.set('uniqueCount', str(n))
        new_ss_bytes = etree.tostring(
            ss_tree, xml_declaration=True, encoding='UTF-8', standalone=True
        )

        # ── sheet XML 읽기 / 셀 값 업데이트 ──
        with z_in.open(sheet_xml_path) as f:
            sh_tree = etree.parse(f).getroot()

        for ref, ss_idx in cell_to_ss.items():
            for c in sh_tree.findall(f'.//{{{SS_NS}}}c[@r="{ref}"]'):
                c.set('t', 's')
                v = c.find(f'{{{SS_NS}}}v')
                if v is None:
                    v = etree.SubElement(c, f'{{{SS_NS}}}v')
                v.text = str(ss_idx)

        new_sh_bytes = etree.tostring(
            sh_tree, xml_declaration=True, encoding='UTF-8', standalone=True
        )

        # ── ZIP 재조립 ──
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for name in names:
                if name == "xl/sharedStrings.xml":
                    z_out.writestr(name, new_ss_bytes)
                elif name == sheet_xml_path:
                    z_out.writestr(name, new_sh_bytes)
                else:
                    z_out.writestr(name, z_in.read(name))

    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())


# ─────────────────────────────────────────────
# makeList 관련
# ─────────────────────────────────────────────
def read_make_list():
    """
    makeList.xlsx에서 지점별 정보 읽기
    반환: {지점명: {'light': '조명'/'비조명', 'date': datetime 또는 None}}
    """
    result = {}
    if not os.path.exists(MAKE_LIST_PATH):
        return result
    try:
        from openpyxl import load_workbook
        wb = load_workbook(MAKE_LIST_PATH, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            name = str(row[0]).strip()

            # 조명 유무 (C열, index 2)
            raw   = row[2] if len(row) > 2 else None
            light = str(raw).strip() if raw else '조명'
            if light not in ('조명', '비조명'):
                light = '조명'

            # 날짜 (B열, index 1) — datetime / 문자열 모두 처리
            date_val    = row[1] if len(row) > 1 else None
            parsed_date = None
            if date_val:
                try:
                    if isinstance(date_val, datetime):
                        parsed_date = date_val
                    else:
                        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y',
                                    '%m/%d/%Y', '%Y.%m.%d', '%d.%m.%Y'):
                            try:
                                parsed_date = datetime.strptime(str(date_val).strip(), fmt)
                                break
                            except ValueError:
                                continue
                except Exception:
                    parsed_date = None

            result[name] = {'light': light, 'date': parsed_date}
    except Exception:
        pass
    return result


def apply_list(log_fn):
    """
    리스트 적용: makeList.xlsx에 있는 지점을 기준으로
    사진/{지점명}/시공전·시공중·시공후 폴더를 생성 (이미 있는 건 건드리지 않음)
    """
    make_list = read_make_list()
    if not make_list:
        log_fn("⚠  makeList.xlsx에 지점이 없습니다.")
        log_fn("    → makeList.xlsx에 지점명을 입력한 뒤 다시 시도해주세요.")
        return

    log_fn(f"📌 makeList 지점 수: {len(make_list)}개")
    created = 0

    for branch in make_list:
        bf = branch_full(branch)

        # '점' 유무 양쪽 모두 확인 → 이미 있으면 건너뜀
        existing_dir, existing_name = get_branch_photo_dir(branch)
        if existing_dir is not None:
            log_fn(f"  ↩  사진/{existing_name}/ 이미 있음 → 건너뜀")
            continue

        # 없으면 bf(점 포함 정규 이름) 기준으로 생성
        for phase in ('시공전', '시공중', '시공후'):
            os.makedirs(os.path.join(SAJIN_DIR, bf, phase), exist_ok=True)
        log_fn(f"  ✅ 사진/{bf}/시공전·시공중·시공후 폴더 생성")
        created += 1

    if created == 0:
        log_fn("ℹ  모든 지점 폴더가 이미 존재합니다.")
    else:
        log_fn(f"\n✅ 리스트 적용 완료 ({created}개 지점 폴더 생성)")


# ─────────────────────────────────────────────
# PPT 관련
# ─────────────────────────────────────────────
def calc_photo_positions(n, box_x, box_y, box_cx, box_cy):
    """
    n장 사진의 레이아웃 좌표 계산 (EMU 단위)
    - 1장: 박스 전체
    - 2장: 좌/우 2열
    - 3장 이상: 2열 그리드
    반환: [(x, y, cx, cy), ...]
    """
    if n <= 0:
        return []
    if n == 1:
        return [(box_x, box_y, box_cx, box_cy)]
    if n == 2:
        w = box_cx // 2
        return [
            (box_x,     box_y, w,          box_cy),
            (box_x + w, box_y, box_cx - w, box_cy),
        ]
    # 3장 이상: 2열 그리드
    cols = 2
    rows = (n + cols - 1) // cols
    w = box_cx // cols
    h = box_cy // rows
    result = []
    for i in range(n):
        row = i // cols
        col = i % cols
        # 마지막 열/행은 나머지 픽셀까지 포함
        cell_w = (box_cx - w) if col == cols - 1 else w
        cell_h = (box_cy - h * (rows - 1)) if row == rows - 1 else h
        result.append((
            box_x + col * w,
            box_y + row * h,
            cell_w,
            cell_h,
        ))
    return result


def get_image_for_ppt(photo_path):
    """
    EXIF 방향 보정 후 이미지 소스와 실제 픽셀 크기 반환
    반환: (img_src, width, height)
      - img_src : BytesIO (Pillow 있는 경우) 또는 파일 경로
      - width, height : 보정 후 픽셀 크기, 알 수 없으면 None
    """
    try:
        from PIL import Image, ImageOps
        import io
        img = Image.open(photo_path)
        img = ImageOps.exif_transpose(img)
        w, h = img.size
        buf = io.BytesIO()
        fmt = 'JPEG' if photo_path.lower().endswith(('.jpg', '.jpeg')) else 'PNG'
        img.save(buf, format=fmt, quality=92)
        buf.seek(0)
        return buf, w, h
    except ImportError:
        return photo_path, None, None   # Pillow 없으면 크기 알 수 없음
    except Exception:
        return photo_path, None, None


# 슬롯 가장자리 여백 비율 (각 면의 %)
PHOTO_MARGIN = 0.025


def fit_photo_in_slot(slot_x, slot_y, slot_cx, slot_cy, img_w, img_h):
    """
    슬롯 안에 이미지를 비율 유지 + 여백 적용하여 배치 (contain 방식)
    img_w / img_h 가 None이면 여백만 적용하고 비율 보정 생략
    반환: (x, y, cx, cy) EMU
    """
    # 여백 계산 (슬롯 크기 대비 비율)
    mg_x = int(slot_cx * PHOTO_MARGIN)
    mg_y = int(slot_cy * PHOTO_MARGIN)
    avail_cx = slot_cx - 2 * mg_x
    avail_cy = slot_cy - 2 * mg_y

    if img_w and img_h:
        # 비율 유지 contain
        scale  = min(avail_cx / img_w, avail_cy / img_h)
        fit_cx = int(img_w * scale)
        fit_cy = int(img_h * scale)
        # 남은 공간 중앙 정렬
        x = slot_x + mg_x + (avail_cx - fit_cx) // 2
        y = slot_y + mg_y + (avail_cy - fit_cy) // 2
    else:
        # 크기 불명 → 여백만 적용, 슬롯 꽉 채움
        fit_cx = avail_cx
        fit_cy = avail_cy
        x      = slot_x + mg_x
        y      = slot_y + mg_y

    return x, y, fit_cx, fit_cy


def run_make_ppt(branch, bf, branch_dir, log_fn, allow_overwrite=False):
    """
    PPT 생성: 사진폼.pptx 템플릿 기반
    - TextBox 4의 '관저점' → bf 교체
    - 기존 p:pic 제거
    - 시공전/시공중/시공후 사진을 각 직사각형 박스 안에 배치
    반환: 저장 경로 또는 None
    """
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.oxml.ns import qn

    out_path = os.path.join(branch_dir, f"TCL_{bf}.pptx")

    # 중복수정 비활성화 시 기존 파일 건너뜀
    if not allow_overwrite and os.path.exists(out_path):
        log_fn(f"⏭  [{bf}] PPT 이미 있음 → 건너뜀")
        return None

    template_path = os.path.join(YANGSIK_DIR, PPT_TEMPLATE)
    if not os.path.exists(template_path):
        log_fn(f"   ⚠ {PPT_TEMPLATE} 없음 → PPT 건너뜀 (양식 폴더에 넣어주세요)")
        return None

    prs  = Presentation(template_path)
    slide = prs.slides[0]

    # ── 1. 지점명 교체 ─────────────────────────────────
    # TextBox 4: 여러 run으로 분리된 "충청지사 - 관저점"
    # '관저점' 텍스트가 있는 run을 bf로 교체
    for shape in slide.shapes:
        if shape.name == 'TextBox 4' and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if '관저점' in run.text:
                        run.text = bf
            break

    # ── 2. 기존 p:pic 제거 (샘플 사진 삭제) ────────────
    spTree = slide.shapes._spTree
    NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    for pic_elem in list(spTree.findall(qn('p:pic'))):
        spTree.remove(pic_elem)

    # 이미지 관계(relationship) XML에서 직접 제거 → 고아 미디어 방지
    # _Relationships 객체는 del을 지원하지 않으므로 내부 XML을 직접 수정
    try:
        rels_elm = slide.part._rels._element   # lxml Element
        IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        for rel in list(rels_elm):
            if rel.get('Type') == IMAGE_TYPE:
                rels_elm.remove(rel)
    except Exception:
        pass   # 관계 정리 실패해도 PPT 파일은 정상 열림

    # ── 3. 사진 삽입 ────────────────────────────────────
    BOX_X  = 1590474
    BOX_CX = 4358806
    BOX_CY = 2403101
    PHASES = [
        ('시공전', 1325762),
        ('시공중', 4278091),
        ('시공후', 7230419),
    ]

    total_photos = 0
    branch_photo_dir, _ = get_branch_photo_dir(branch)

    if branch_photo_dir is None:
        log_fn(f"   ⚠ [{bf}] 사진 폴더 없음 (사진/{branch} 또는 사진/{bf})")
    else:
        for phase_name, box_y in PHASES:
            phase_dir = os.path.join(branch_photo_dir, phase_name)
            if not os.path.exists(phase_dir):
                continue

            photos = sorted([
                os.path.join(phase_dir, f)
                for f in os.listdir(phase_dir)
                if f.lower().endswith(SUPPORTED_EXTS)
            ])

            if not photos:
                continue

            slots = calc_photo_positions(
                len(photos), BOX_X, box_y, BOX_CX, BOX_CY
            )

            for photo_path, slot in zip(photos, slots):
                img_src, img_w, img_h = get_image_for_ppt(photo_path)
                x, y, cx, cy = fit_photo_in_slot(*slot, img_w, img_h)
                slide.shapes.add_picture(
                    img_src,
                    Emu(x), Emu(y), Emu(cx), Emu(cy)
                )
                total_photos += 1

            log_fn(f"   └ {phase_name}: {len(photos)}장 삽입")

    # ── 4. 저장 ─────────────────────────────────────────
    try:
        prs.save(out_path)
        log_fn(f"✅ [{bf}] TCL_{bf}.pptx 저장 완료! (사진 {total_photos}장)")
        return out_path
    except PermissionError:
        log_fn(f"❌ [{bf}] PPT 파일이 열려 있어 저장 실패 → 파일을 닫고 다시 시도하세요.")
        return None


# ─────────────────────────────────────────────
# PPT → 이미지 변환
# ─────────────────────────────────────────────

def _powerpoint_com_registered():
    """
    PowerPoint COM 서버 등록 여부 확인
    여러 레지스트리 위치를 순서대로 탐색 (Office 365 Click-to-Run 대응)

    Office 365 Click-to-Run은 HKEY_CLASSES_ROOT 최상위에
    'PowerPoint.Application' 키를 등록하지 않는 경우가 있어
    HKLM\\SOFTWARE\\Classes 및 버전별 키(.16/.15)도 함께 확인
    """
    try:
        import winreg
        checks = [
            # 버전 중립 키 — 3곳
            (winreg.HKEY_CLASSES_ROOT,   'PowerPoint.Application'),
            (winreg.HKEY_LOCAL_MACHINE,  r'SOFTWARE\Classes\PowerPoint.Application'),
            (winreg.HKEY_CURRENT_USER,   r'SOFTWARE\Classes\PowerPoint.Application'),
            # 버전별 키 — Office 365/2016/2019/2021 = .16, Office 2013 = .15
            (winreg.HKEY_CLASSES_ROOT,   'PowerPoint.Application.16'),
            (winreg.HKEY_LOCAL_MACHINE,  r'SOFTWARE\Classes\PowerPoint.Application.16'),
            (winreg.HKEY_CLASSES_ROOT,   'PowerPoint.Application.15'),
            (winreg.HKEY_LOCAL_MACHINE,  r'SOFTWARE\Classes\PowerPoint.Application.15'),
        ]
        for hive, subkey in checks:
            try:
                winreg.OpenKey(hive, subkey)
                return True          # 하나라도 열리면 설치 확인
            except OSError:
                continue
    except ImportError:
        pass
    return False


def _get_ppt_engine():
    """
    사용 가능한 PPT→이미지 변환 엔진 함수 반환
    우선순위: ① win32com (PowerPoint COM 등록 확인) → ② LibreOffice CLI
    없으면 None 반환
    """
    # ① pywin32 설치 여부 + PowerPoint COM 등록 여부 확인
    try:
        import win32com.client  # noqa
        import pythoncom        # noqa
        if _powerpoint_com_registered():
            return _convert_with_powerpoint
    except ImportError:
        pass

    # ② LibreOffice 헤드리스
    import shutil
    for cmd in ('soffice', 'libreoffice'):
        if shutil.which(cmd):
            return _convert_with_libreoffice

    return None


_RPC_E_CALL_REJECTED = -2147418111   # 0x80010001  피호출자가 호출을 거부


def _com_retry(func, max_retries=10, base_delay=0.4):
    """
    COM 'RPC_E_CALL_REJECTED (피호출자가 호출을 거부했습니다)' 오류 시 재시도
    PowerPoint가 내부 작업 중일 때 잠시 기다렸다가 다시 호출
    """
    import time
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            is_rejected = (
                hasattr(e, 'args') and e.args
                and e.args[0] == _RPC_E_CALL_REJECTED
            )
            if is_rejected and attempt < max_retries - 1:
                wait = min(base_delay * (attempt + 1), 3.0)
                time.sleep(wait)
                continue
            raise


def _convert_with_powerpoint(tasks, allow_overwrite, log_fn):
    """
    win32com / PowerPoint 를 이용한 슬라이드 → PNG 변환
    PowerPoint 인스턴스를 한 번만 열고 모든 지점 파일을 처리 후 닫음
    tasks: [(pptx_path, out_dir, bf), ...]
    """
    import win32com.client
    import pythoncom
    import time

    pythoncom.CoInitialize()
    app = None
    all_paths = []

    try:
        # Dispatch: COM 런타임이 설치된 버전으로 자동 연결
        # DispatchEx(버전별 ProgID) 방식은 구버전 잔여 레지스트리와 충돌 위험
        app = win32com.client.Dispatch('PowerPoint.Application')
        app.Visible = 1   # 0이면 일부 환경에서 Export 실패
        time.sleep(1.0)   # PowerPoint 초기화 대기

        for pptx_path, out_dir, bf in tasks:
            log_fn(f"🖼  [{bf}] 이미지 변환 중...")
            prs = None
            try:
                abs_pptx = os.path.abspath(pptx_path)
                # Open 도 거부될 수 있으므로 retry 적용
                prs = _com_retry(
                    lambda p=abs_pptx: app.Presentations.Open(
                        p, ReadOnly=True, Untitled=True, WithWindow=False
                    )
                )

                # 슬라이드 비율 기반 출력 해상도 (1920px 기준)
                px_w = 1920
                px_h = int(px_w * prs.PageSetup.SlideHeight
                           / prs.PageSetup.SlideWidth)

                n = prs.Slides.Count
                for i in range(1, n + 1):
                    img_name = (f"TCL_{bf}.png" if n == 1
                                else f"TCL_{bf}_{i:02d}.png")
                    img_path = os.path.join(out_dir, img_name)

                    if not allow_overwrite and os.path.exists(img_path):
                        log_fn(f"   ⏭  {img_name} 이미 있음 → 건너뜀")
                        continue

                    # Export 도 retry 적용
                    abs_img = os.path.abspath(img_path)
                    _com_retry(
                        lambda s=prs.Slides(i), p=abs_img: s.Export(p, 'PNG', px_w, px_h)
                    )
                    log_fn(f"   ✅ {img_name}")
                    all_paths.append(img_path)

            finally:
                if prs:
                    try:
                        prs.Close()
                    except Exception:
                        pass
                    time.sleep(0.8)   # 파일 간 PowerPoint 정리 대기

    finally:
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return all_paths


def _convert_with_libreoffice(tasks, allow_overwrite, log_fn):
    """
    LibreOffice 헤드리스를 이용한 슬라이드 → PNG 변환
    tasks: [(pptx_path, out_dir, bf), ...]
    """
    import subprocess
    import shutil
    import tempfile

    cmd = 'soffice' if shutil.which('soffice') else 'libreoffice'
    all_paths = []

    for pptx_path, out_dir, bf in tasks:
        log_fn(f"🖼  [{bf}] 이미지 변환 중...")

        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                [cmd, '--headless', '--convert-to', 'png',
                 '--outdir', tmpdir, os.path.abspath(pptx_path)],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode != 0:
                log_fn(f"   ❌ LibreOffice 변환 실패: {result.stderr[:300]}")
                continue

            base = os.path.splitext(os.path.basename(pptx_path))[0]

            # 단일 슬라이드: base.png / 복수: base1.png, base2.png ...
            single = os.path.join(tmpdir, base + '.png')
            if os.path.exists(single):
                img_name = f"TCL_{bf}.png"
                img_path = os.path.join(out_dir, img_name)
                if allow_overwrite or not os.path.exists(img_path):
                    shutil.move(single, img_path)
                    log_fn(f"   ✅ {img_name}")
                    all_paths.append(img_path)
                else:
                    log_fn(f"   ⏭  {img_name} 이미 있음 → 건너뜀")
            else:
                idx = 1
                while True:
                    src = os.path.join(tmpdir, f"{base}{idx}.png")
                    if not os.path.exists(src):
                        break
                    img_name = f"TCL_{bf}_{idx:02d}.png"
                    img_path = os.path.join(out_dir, img_name)
                    if allow_overwrite or not os.path.exists(img_path):
                        shutil.move(src, img_path)
                        log_fn(f"   ✅ {img_name}")
                        all_paths.append(img_path)
                    else:
                        log_fn(f"   ⏭  {img_name} 이미 있음 → 건너뜀")
                    idx += 1

    return all_paths


def run_ppt_to_image(log_fn, allow_overwrite=False):
    """
    각 지점의 TCL_{bf}.pptx → TCL_{bf}.png 이미지 변환
    엔진 우선순위: PowerPoint(win32com) → LibreOffice
    PowerPoint 인스턴스는 한 번만 열고 모든 지점을 처리
    """
    engine = _get_ppt_engine()
    if engine is None:
        raise RuntimeError(
            "PPT → 이미지 변환에 필요한 프로그램을 찾지 못했습니다.\n\n"
            "다음 중 하나를 준비해주세요:\n\n"
            "① Microsoft PowerPoint 설치\n"
            "   (설치 후 '실행하기.bat' 로 pywin32 패키지도 설치)\n\n"
            "② LibreOffice 설치  (무료)\n"
            "   https://www.libreoffice.org/\n\n"
            "※ pywin32는 설치됐지만 PowerPoint가 없거나\n"
            "   COM 등록이 안 된 경우에도 이 메시지가 뜹니다."
        )

    make_list = read_make_list()
    if not make_list:
        raise ValueError("makeList.xlsx에 지점이 없습니다.")

    # 변환할 파일 목록 수집 (PPT 없는 지점 미리 건너뜀)
    tasks = []
    for branch in make_list:
        bf         = branch_full(branch)
        branch_dir = os.path.join(OUTPUT_DIR, bf)
        pptx_path  = os.path.join(branch_dir, f"TCL_{bf}.pptx")
        if os.path.exists(pptx_path):
            tasks.append((pptx_path, branch_dir, bf))
        else:
            log_fn(f"⏭  [{bf}] PPT 파일 없음 → 건너뜀")

    if not tasks:
        return []

    engine_name = ('PowerPoint' if engine is _convert_with_powerpoint
                   else 'LibreOffice')
    log_fn(f"🔧 변환 엔진: {engine_name}  ({len(tasks)}개 지점)")

    # 엔진에 전체 목록을 한 번에 전달 → PowerPoint 인스턴스 1회만 생성
    return engine(tasks, allow_overwrite, log_fn)


# ─────────────────────────────────────────────
# 견적서 생성
# ─────────────────────────────────────────────
def run_make(log_fn, allow_overwrite=False):
    """견적서(Excel + PPT) 생성 핵심 함수 — makeList.xlsx 기준으로 지점별 생성"""
    make_list = read_make_list()
    if not make_list:
        raise ValueError(
            "makeList.xlsx에 지점이 없습니다.\n\n"
            "'리스트 업데이트' 버튼으로 지점을 추가하거나\n"
            "makeList.xlsx에 직접 지점명을 입력해주세요."
        )

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    log_fn(f"📌 처리할 지점: {', '.join(make_list.keys())}")

    created = []

    for branch, info in make_list.items():
        bf = branch_full(branch)           # 예: 관저 → 관저점

        # ── 출력 경로: 견적서/관저점/ ──
        branch_dir = os.path.join(OUTPUT_DIR, bf)
        os.makedirs(branch_dir, exist_ok=True)

        # ── 사진 폴더 자동 생성 (없으면 만들기) ──
        # get_branch_photo_dir 로 '점' 유무 양쪽 체크 → 없을 때만 bf 기준 생성
        existing_dir, _ = get_branch_photo_dir(branch)
        if existing_dir is None:
            for phase in ('시공전', '시공중', '시공후'):
                os.makedirs(os.path.join(SAJIN_DIR, bf, phase), exist_ok=True)

        # ━━ Excel 생성 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        fname    = f"TCL_{bf}.xlsx"
        out_path = os.path.join(branch_dir, fname)

        if not allow_overwrite and os.path.exists(out_path):
            log_fn(f"⏭  [{bf}] Excel 이미 있음 → 건너뜀 (중복수정 비활성화 상태)")
        else:
            light_type    = info.get('light', '조명')
            template_path = get_template_path(light_type)

            log_fn(f"📋 [{bf}] ({light_type}) Excel 생성 중...")

            # 날짜: makeList에 있으면 입력, 없으면 양식 그대로 유지 (None → ZIP에서 건너뜀)
            list_date = info.get('date', None)
            if list_date:
                log_fn(f"   └ 날짜: {list_date.year}년 {list_date.month}월 {list_date.day}일")
            else:
                log_fn(f"   └ 날짜: makeList 없음 → 양식 날짜 유지")

            # A6 (고객명): 양식 그대로 유지 — 건드리지 않음

            # B9 (건명): "TCL 로고 제작설치 작업-관저점" 형식
            updates = {
                'A4': f"{list_date.year}년"  if list_date else None,
                'B4': f"{list_date.month}월" if list_date else None,
                'C4': f"{list_date.day}일"   if list_date else None,
                'B9': f"TCL 로고 제작설치 작업-{bf}",
            }

            try:
                _save_excel_zip(template_path, out_path, updates)
                log_fn(f"✅ [{bf}] {fname} 저장 완료!")
                created.append(out_path)
            except PermissionError:
                log_fn(f"❌ [{bf}] 파일이 열려 있어 저장 실패 → 엑셀을 닫고 다시 시도하세요.")
            except Exception as e:
                log_fn(f"❌ [{bf}] Excel 저장 오류: {e}")

        # ━━ PPT 생성 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        log_fn(f"🎞  [{bf}] PPT 생성 중...")
        ppt_path = run_make_ppt(
            branch, bf, branch_dir, log_fn,
            allow_overwrite=allow_overwrite
        )
        if ppt_path:
            created.append(ppt_path)

    return created


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("수산나디자인  견적서 자동화")
        self.geometry("480x620")
        self.resizable(False, False)
        self.configure(bg="#F7F9FC")
        self._build_ui()

    def _build_ui(self):
        # ── 헤더 ──────────────────────────────
        header = tk.Frame(self, bg="#1565C0", height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="수산나디자인  견적서 자동화",
            font=("맑은 고딕", 16, "bold"),
            bg="#1565C0", fg="white"
        ).pack(expand=True)

        # ── 메인 버튼 ─────────────────────────
        mid = tk.Frame(self, bg="#F7F9FC")
        mid.pack(pady=(20, 4))

        self.make_btn = tk.Button(
            mid,
            text="📋   견적서 만들기",
            font=("맑은 고딕", 15, "bold"),
            bg="#1976D2", fg="white",
            activebackground="#0D47A1",
            relief="flat",
            padx=36, pady=14,
            cursor="hand2",
            command=self._on_make
        )
        self.make_btn.pack()

        # ── 중복수정 체크박스 ──────────────────
        self.overwrite_var = tk.IntVar(value=0)
        tk.Checkbutton(
            mid,
            text="중복수정  (이미 있는 파일 덮어쓰기 허용)",
            variable=self.overwrite_var,
            font=("맑은 고딕", 10),
            bg="#F7F9FC", fg="#555",
            activebackground="#F7F9FC",
            selectcolor="#F7F9FC",
        ).pack(pady=(8, 0))

        # ── 폴더 버튼 ─────────────────────────
        btn_frame = tk.Frame(self, bg="#F7F9FC")
        btn_frame.pack(pady=6)

        for label, path in [
            ("📁  사진 폴더", SAJIN_DIR),
            ("📁  양식 폴더", YANGSIK_DIR),
            ("📂  결과 폴더", OUTPUT_DIR),
        ]:
            tk.Button(
                btn_frame,
                text=label,
                font=("맑은 고딕", 10),
                bg="#E3F2FD", fg="#1565C0",
                activebackground="#BBDEFB",
                relief="flat",
                padx=10, pady=6,
                cursor="hand2",
                command=lambda p=path: self._open_folder(p)
            ).pack(side="left", padx=5)

        # ── 리스트 적용 + 이미지 변환 버튼 ──────
        list_frame = tk.Frame(self, bg="#F7F9FC")
        list_frame.pack(pady=2)

        tk.Button(
            list_frame,
            text="🗂   리스트 적용  (makeList.xlsx → 사진 폴더 생성)",
            font=("맑은 고딕", 10),
            bg="#E8F5E9", fg="#1B5E20",
            activebackground="#C8E6C9",
            relief="flat",
            padx=10, pady=6,
            cursor="hand2",
            command=self._on_apply_list
        ).pack()

        img_frame = tk.Frame(self, bg="#F7F9FC")
        img_frame.pack(pady=2)

        self.img_btn = tk.Button(
            img_frame,
            text="🖼   PPT → 이미지 변환",
            font=("맑은 고딕", 10),
            bg="#FFF8E1", fg="#E65100",
            activebackground="#FFE0B2",
            relief="flat",
            padx=10, pady=6,
            cursor="hand2",
            command=self._on_ppt_to_image
        )
        self.img_btn.pack()

        # ── 진행 로그 ─────────────────────────
        log_frame = tk.Frame(self, bg="#F7F9FC")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 0))

        tk.Label(
            log_frame, text="진행 상황",
            font=("맑은 고딕", 10, "bold"),
            bg="#F7F9FC", fg="#555"
        ).pack(anchor="w")

        self.log_box = tk.Text(
            log_frame,
            height=10,
            font=("맑은 고딕", 10),
            bg="#F0F4F8", fg="#222",
            relief="solid", bd=1,
            state="disabled",
            padx=8, pady=6
        )
        self.log_box.pack(fill="both", expand=True)

        # ── 상태바 ────────────────────────────
        self.status_var = tk.StringVar(value="  준비 완료")
        tk.Label(
            self,
            textvariable=self.status_var,
            font=("맑은 고딕", 9),
            bg="#CFD8DC", fg="#37474F",
            anchor="w", padx=10
        ).pack(fill="x", side="bottom")

    # ── 이벤트 핸들러 ─────────────────────────
    def _log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        self.update_idletasks()

    def _open_folder(self, path):
        os.makedirs(path, exist_ok=True)
        subprocess.Popen(["explorer", os.path.normpath(path)])

    def _on_make(self):
        self.make_btn.config(state="disabled", text="⏳   처리 중...")
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.status_var.set("  처리 중...")

        allow_overwrite = bool(self.overwrite_var.get())

        def worker():
            try:
                results = run_make(self._log, allow_overwrite=allow_overwrite)
                if results:
                    self._log(f"\n🎉  완료!  총 {len(results)}개 파일 생성됨")
                    self.status_var.set(f"  완료 — {len(results)}개 파일 생성")
                    self._open_folder(OUTPUT_DIR)
                else:
                    self._log("\n⚠  생성된 파일이 없습니다.")
                    self._log("    → makeList.xlsx에 지점을 추가하거나 중복수정을 체크해주세요.")
                    self.status_var.set("  완료 (생성 파일 없음)")
            except Exception as e:
                self._log(f"\n❌  오류 발생: {e}")
                self.status_var.set("  오류 발생")
                messagebox.showerror("오류", str(e))
            finally:
                self.make_btn.config(state="normal", text="📋   견적서 만들기")

        threading.Thread(target=worker, daemon=True).start()

    def _on_ppt_to_image(self):
        self.img_btn.config(state="disabled", text="⏳   변환 중...")
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.status_var.set("  PPT → 이미지 변환 중...")

        allow_overwrite = bool(self.overwrite_var.get())

        def worker():
            try:
                results = run_ppt_to_image(self._log, allow_overwrite=allow_overwrite)
                if results:
                    self._log(f"\n🎉  완료!  총 {len(results)}개 이미지 저장됨")
                    self.status_var.set(f"  완료 — {len(results)}개 이미지 저장")
                    self._open_folder(OUTPUT_DIR)
                else:
                    self._log("\n⚠  저장된 이미지가 없습니다.")
                    self._log("    → 먼저 '견적서 만들기'로 PPT를 생성하거나 중복수정을 체크해주세요.")
                    self.status_var.set("  완료 (저장 이미지 없음)")
            except Exception as e:
                self._log(f"\n❌  오류 발생: {e}")
                self.status_var.set("  오류 발생")
                messagebox.showerror("오류", str(e))
            finally:
                self.img_btn.config(state="normal", text="🖼   PPT → 이미지 변환")

        threading.Thread(target=worker, daemon=True).start()

    def _on_apply_list(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.status_var.set("  리스트 적용 중...")

        def worker():
            try:
                apply_list(self._log)
                self.status_var.set("  리스트 적용 완료")
            except Exception as e:
                self._log(f"❌  오류: {e}")
                self.status_var.set("  오류 발생")

        threading.Thread(target=worker, daemon=True).start()


# ─────────────────────────────────────────────
# 시작
# ─────────────────────────────────────────────
if __name__ == "__main__":
    missing = []
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")
    try:
        import pptx
    except ImportError:
        missing.append("python-pptx")

    if missing:
        root = tk.Tk()
        root.withdraw()
        msg = (
            "필수 패키지가 설치되어 있지 않습니다.\n\n"
            f"설치 필요: {', '.join(missing)}\n\n"
            "「실행하기.bat」 파일을 더블클릭해서 먼저 설치해주세요."
        )
        messagebox.showerror("패키지 오류", msg)
        sys.exit(1)

    app = App()
    app.mainloop()
