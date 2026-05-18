# python-docx 실전 레시피 — 프로페셔널 워드 문서 생성

> 이 문서의 코드를 조합하면 **표지 → 목차 → 본문(표+그래프+하이라이트) → 마무리**까지 완성된 워드 문서를 자동 생성할 수 있습니다.

---

## 전체 구조

```python
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['axes.unicode_minus'] = False

doc = Document()
# 1. 초기 설정
# 2. 표지
# 3. 목차
# 4. 본문 섹션들 (표 + 그래프 + 하이라이트)
# 5. 마무리
# 6. 저장
doc.save('output.docx')
```

---

## 레시피 1. 문서 초기 설정 (폰트 + 여백 + 줄간격)

```python
def init_document(doc):
    """문서 기본 스타일 초기화 (HD현대마린솔루션테크 CI)"""
    # Normal 스타일
    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x2C, 0x2C, 0x2C)  # HD Dark Gray
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(4)

    # Heading 스타일 (HD현대 초록 기반)
    for level, (size, color) in enumerate([
        (16, '008233'),  # Heading 1 — HD Prosperity Green
        (13, '008233'),  # Heading 2 — HD Prosperity Green
        (11, '00AD1D'),  # Heading 3 — HD Heritage Green
    ], 1):
        h = doc.styles[f'Heading {level}']
        h.font.name = '맑은 고딕'
        h.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        h.font.size = Pt(size)
        h.font.bold = True
        r, g, b = int(color[:2], 16), int(color[2:4], 16), int(color[4:], 16)
        h.font.color.rgb = RGBColor(r, g, b)
        h.paragraph_format.space_before = Pt(18 if level == 1 else 12)
        h.paragraph_format.space_after = Pt(6)

    # 여백 (HD현대 기준)
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    return doc
```

---

## 레시피 2. 프로페셔널 표지

```python
def add_cover(doc, title, subtitle, date, author, department):
    """깔끔한 표지 페이지"""
    for _ in range(5):
        doc.add_paragraph('')

    # 상단 장식선
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('━' * 40)
    run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    run.font.size = Pt(10)

    doc.add_paragraph('')

    # 메인 제목
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    run.font.name = '맑은 고딕'

    doc.add_paragraph('')

    # 부제
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(subtitle)
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # 하단 장식선
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('━' * 40)
    run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    run.font.size = Pt(10)

    # 발표 정보
    for _ in range(5):
        doc.add_paragraph('')
    for text in [date, f'{department}  |  {author}']:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_page_break()
```

---

## 레시피 3. 스타일링된 표 (헤더 + 교대행 + Bold 강조)

```python
# ──────── 컬러 상수 (HD현대 CI 기반 — 초록 메인) ────────
NAVY    = '008233'   # HD Prosperity Green (메인 — 제목, 표 헤더)
BLUE    = '00AD1D'   # HD Heritage Green (보조 강조)
LIGHT   = 'E8F5E9'   # Light Green (박스/하이라이트 배경)
RED     = 'E74C3C'   # 위험/경고
GREEN   = '008233'   # 양호/정상 (메인과 동일)
YELLOW  = 'F39C12'   # 주의
WHITE   = 'FFFFFF'
DARK_GREEN = '004D1A' # 진한 초록 (Phase 1 등)
HD_BLUE = '003087'    # HD Discovery Blue (보조 대비용)

def set_cell_shading(cell, color_hex):
    """셀 배경색 설정"""
    tcPr = cell._element.get_or_add_tcPr()
    shading = tcPr.makeelement(qn('w:shd'), {
        qn('w:fill'): color_hex,
        qn('w:val'): 'clear',
    })
    tcPr.append(shading)

def set_cell_text(cell, text, bold=False, color_rgb=None, size=Pt(10), align=WD_ALIGN_PARAGRAPH.LEFT):
    """셀 텍스트 스타일 설정"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.font.size = size
    run.font.bold = bold
    run.font.name = '맑은 고딕'
    if color_rgb:
        run.font.color.rgb = color_rgb

def add_styled_table(doc, headers, rows, col_widths=None):
    """
    프로페셔널 스타일 표 생성

    Parameters:
        headers: ['열1', '열2', '열3']
        rows: [['값1', '값2', '값3'], ...]
        col_widths: [Inches(2), Inches(3), ...] (선택)
    """
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # ── 헤더 행 ──
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_text(cell, header, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF), size=Pt(10))

    # ── 데이터 행 (교대 배경색) ──
    for row_idx, row_data in enumerate(rows):
        for col_idx, value in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            # 교대 행 배경
            if row_idx % 2 == 1:
                set_cell_shading(cell, LIGHT)
            # Bold 처리: **텍스트** 패턴
            if isinstance(value, str) and value.startswith('**') and value.endswith('**'):
                set_cell_text(cell, value[2:-2], bold=True, size=Pt(10))
            else:
                set_cell_text(cell, value, size=Pt(10))

    # ── 열 너비 ──
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width

    # 표 아래 여백
    doc.add_paragraph('')
    return table
```

---

## 레시피 4. Gap 분석 표 (🔴🟡🟢 컬러 코딩)

```python
def add_gap_table(doc, headers, rows):
    """
    Gap 분석 전용 표 — 마지막 열의 🔴🟡🟢에 따라 배경색 적용

    rows 예시:
    [['고객용 웹 포털', 'Wärtsilä Online, K-IMS', '없음', '🔴 큼'], ...]
    """
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_text(cell, header, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF))

    # 데이터
    for row_idx, row_data in enumerate(rows):
        for col_idx, value in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]

            # 교대 행
            if row_idx % 2 == 1:
                set_cell_shading(cell, LIGHT)

            # 마지막 열 (Gap 컬러)
            if col_idx == len(headers) - 1:
                if '🔴' in str(value):
                    set_cell_shading(cell, 'FDEDEC')  # 연한 빨강
                    set_cell_text(cell, value, bold=True,
                                  color_rgb=RGBColor(0xE7, 0x4C, 0x3C))
                elif '🟡' in str(value):
                    set_cell_shading(cell, 'FEF9E7')  # 연한 노랑
                    set_cell_text(cell, value, bold=True,
                                  color_rgb=RGBColor(0xF3, 0x9C, 0x12))
                elif '🟢' in str(value):
                    set_cell_shading(cell, 'EAFAF1')  # 연한 초록
                    set_cell_text(cell, value, bold=True,
                                  color_rgb=RGBColor(0x27, 0xAE, 0x60))
                else:
                    set_cell_text(cell, value)
            else:
                set_cell_text(cell, value)

    doc.add_paragraph('')
    return table
```

---

## 레시피 5. 인사이트 박스 (Callout)

```python
def add_insight_box(doc, text, bg_color='E8F0FE', border_color='1B3A5C'):
    """
    왼쪽 파란 보더 + 연한 배경의 핵심 인사이트 박스

    사용 예:
        add_insight_box(doc, '기술력과 네트워크는 선도기업 이상이다.\n'
                             '부족한 건 고객이 우리 서비스에 접근하는 "디지털 창구"다.')
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)

    # 배경색
    set_cell_shading(cell, bg_color)

    # 왼쪽 보더
    tcPr = cell._element.get_or_add_tcPr()
    borders = tcPr.makeelement(qn('w:tcBorders'), {})
    for side, width in [('left', '24'), ('top', '4'), ('bottom', '4'), ('right', '4')]:
        border = borders.makeelement(qn(f'w:{side}'), {
            qn('w:val'): 'single',
            qn('w:sz'): width,
            qn('w:color'): border_color if side == 'left' else 'D5D5D5',
            qn('w:space'): '0',
        })
        borders.append(border)
    tcPr.append(borders)

    # 텍스트
    cell.text = ''
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(
        int(border_color[:2], 16),
        int(border_color[2:4], 16),
        int(border_color[4:], 16)
    )
    run.font.name = '맑은 고딕'

    doc.add_paragraph('')
    return table
```

---

## 레시피 6. Before/After 비교표

```python
def add_before_after_table(doc, title, rows):
    """
    Before(회색)/After(블루) 비교표

    rows 예시:
    [('서비스 요청', '전화/이메일 (업무시간)', '웹 포털 24/7 (즉시)'), ...]
    """
    doc.add_heading(title, level=3)

    table = doc.add_table(rows=1 + len(rows), cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더
    headers = ['항목', 'Before (현재)', 'After (웹 포털)']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_text(cell, header, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF))

    # 데이터
    for row_idx, (item, before, after) in enumerate(rows):
        # 항목명
        set_cell_text(table.rows[row_idx + 1].cells[0], item, bold=True)

        # Before — 연한 회색 배경
        cell_before = table.rows[row_idx + 1].cells[1]
        set_cell_shading(cell_before, 'F5F5F5')
        set_cell_text(cell_before, before, color_rgb=RGBColor(0x99, 0x99, 0x99))

        # After — 연한 블루 배경 + Bold
        cell_after = table.rows[row_idx + 1].cells[2]
        set_cell_shading(cell_after, 'E8F0FE')
        set_cell_text(cell_after, after, bold=True,
                      color_rgb=RGBColor(0x1B, 0x3A, 0x5C))

    doc.add_paragraph('')
    return table
```

---

## 레시피 7. 그래프 생성 + 워드 삽입

```python
def create_and_insert_line_chart(doc, title, x_data, y_data, xlabel, ylabel, save_path):
    """꺾은선 그래프 생성 후 워드에 삽입"""
    fig, ax = plt.subplots(figsize=(7, 3.5))

    ax.plot(x_data, y_data, color='#1B3A5C', linewidth=2.5,
            marker='o', markerfacecolor='#2B579A', markersize=7)
    ax.fill_between(x_data, y_data, alpha=0.08, color='#2B579A')

    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A5C', pad=15)
    ax.set_xlabel(xlabel, fontsize=9, color='#666666')
    ax.set_ylabel(ylabel, fontsize=9, color='#666666')
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(5.8))
    # 캡션
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph('')


def create_and_insert_bar_chart(doc, title, labels, values, save_path,
                                 colors=None, horizontal=True):
    """막대 차트 생성 후 워드에 삽입"""
    fig, ax = plt.subplots(figsize=(7, max(3, len(labels) * 0.5)))

    if colors is None:
        colors = ['#1B3A5C'] * len(labels)

    if horizontal:
        bars = ax.barh(labels, values, color=colors, height=0.6)
        ax.set_xlabel(title.split('(')[-1].rstrip(')') if '(' in title else '')
        # 값 표시
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max(values) * 0.02, bar.get_y() + bar.get_height() / 2,
                    f'{val}', va='center', fontsize=9, color='#333333')
    else:
        bars = ax.bar(labels, values, color=colors, width=0.6)
        ax.set_ylabel(title.split('(')[-1].rstrip(')') if '(' in title else '')
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.02,
                    f'{val}', ha='center', fontsize=9, color='#333333')

    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A5C', pad=15)
    ax.grid(axis='x' if horizontal else 'y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(5.8))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph('')


def create_and_insert_donut_chart(doc, title, labels, sizes, save_path, colors=None):
    """도넛 차트 생성 후 워드에 삽입"""
    fig, ax = plt.subplots(figsize=(5, 5))

    if colors is None:
        colors = ['#1B3A5C', '#2B579A', '#5B9BD5', '#A5C8E1', '#D6E8F5']

    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors[:len(labels)],
        autopct='%1.0f%%', startangle=90, pctdistance=0.75,
        wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2)
    )

    for text in texts:
        text.set_fontsize(9)
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')

    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A5C', pad=20)
    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(4))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph('')
```

---

## 레시피 7.5. 고급 차트 기법 — 주석(Annotation) + 스타일 시트

```python
def create_annotated_line_chart(doc, title, x_data, y_data, annotations, save_path):
    """
    핵심 포인트에 주석 화살표가 달린 꺾은선 그래프

    annotations 예시:
    [
        (2024, 371, '현재\n371억$'),
        (2032, 532, '목표\n532억$'),
    ]
    """
    fig, ax = plt.subplots(figsize=(7, 3.5))

    # 메인 라인
    ax.plot(x_data, y_data, color='#1B3A5C', linewidth=2.5,
            marker='o', markerfacecolor='#2B579A', markersize=7, zorder=3)
    ax.fill_between(x_data, y_data, alpha=0.06, color='#2B579A')

    # 주석 화살표
    for x, y, label in annotations:
        ax.annotate(label,
            xy=(x, y), xytext=(0, 30),
            textcoords='offset points',
            fontsize=9, fontweight='bold', color='#1B3A5C',
            ha='center', va='bottom',
            bbox=dict(boxstyle='round,pad=0.4', facecolor='#E8F0FE',
                      edgecolor='#2B579A', alpha=0.9),
            arrowprops=dict(arrowstyle='->', color='#2B579A', lw=1.5))

    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A5C', pad=15)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(5.8))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


def create_comparison_bar_chart(doc, title, categories, group1_vals, group2_vals,
                                 group1_label, group2_label, save_path):
    """경쟁사 비교용 나란히 막대 차트"""
    import numpy as np
    fig, ax = plt.subplots(figsize=(7, max(3.5, len(categories) * 0.6)))

    y = np.arange(len(categories))
    height = 0.35

    bars1 = ax.barh(y - height/2, group1_vals, height, color='#AECDE0',
                     label=group1_label, edgecolor='white')
    bars2 = ax.barh(y + height/2, group2_vals, height, color='#1B3A5C',
                     label=group2_label, edgecolor='white')

    # 값 레이블
    for bar, val in zip(bars1, group1_vals):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2,
                f'{val}', va='center', fontsize=8, color='#666666')
    for bar, val in zip(bars2, group2_vals):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2,
                f'{val}', va='center', fontsize=8, color='#1B3A5C', fontweight='bold')

    ax.set_yticks(y)
    ax.set_yticklabels(categories, fontsize=9)
    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A5C', pad=15)
    ax.legend(loc='lower right', fontsize=8, framealpha=0.9)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='x', alpha=0.3, linestyle='--')

    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(5.8))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


def create_gauge_chart(doc, title, score, save_path, max_score=100):
    """건강도 점수 게이지 차트 (반원형)"""
    import numpy as np

    fig, ax = plt.subplots(figsize=(4, 2.5), subplot_kw={'projection': 'polar'})

    # 게이지 범위 (반원)
    theta = np.linspace(np.pi, 0, 100)
    colors_map = ['#E74C3C'] * 30 + ['#F39C12'] * 20 + ['#27AE60'] * 50

    # 배경 (회색 반원)
    for i in range(99):
        ax.bar(theta[i], 1, width=np.pi/100, bottom=0.6,
               color='#E8E8E8', edgecolor='none')

    # 점수 채우기
    for i in range(min(score, 99)):
        ax.bar(theta[i], 1, width=np.pi/100, bottom=0.6,
               color=colors_map[i], edgecolor='none')

    # 중앙 점수 텍스트
    score_color = '#27AE60' if score >= 70 else '#F39C12' if score >= 40 else '#E74C3C'
    ax.text(np.pi/2, 0.1, f'{score}', ha='center', va='center',
            fontsize=28, fontweight='bold', color=score_color)
    ax.text(np.pi/2, -0.25, f'/ {max_score}', ha='center', va='center',
            fontsize=10, color='#999999')

    ax.set_ylim(0, 1.6)
    ax.set_thetamin(0)
    ax.set_thetamax(180)
    ax.axis('off')

    plt.title(title, fontsize=12, fontweight='bold', color='#1B3A5C', pad=10, y=1.05)
    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(3.5))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

---

## 레시피 8. 로드맵 타임라인 표

```python
def add_roadmap_table(doc, phases):
    """
    Phase별 색상 그라데이션 로드맵 표

    phases 예시:
    [
        ('Phase 1: 즉시', '1~3개월', '메일 AI + 리포트 시각화', '낮음'),
        ('Phase 2: 단기', '3~6개월', '웹 포털 MVP', '중간'),
        ...
    ]
    """
    phase_colors = ['27AE60', '2B579A', '8E44AD', '1B3A5C']
    headers = ['Phase', '기간', '내용', '투자 규모']

    table = doc.add_table(rows=1 + len(phases), cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_text(cell, header, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF))

    # Phase 행
    for idx, (phase, period, content, invest) in enumerate(phases):
        color = phase_colors[idx % len(phase_colors)]
        row = table.rows[idx + 1]

        # Phase 이름 셀 — 컬러 배경
        set_cell_shading(row.cells[0], color)
        set_cell_text(row.cells[0], phase, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF), size=Pt(9))

        # 나머지 셀
        set_cell_text(row.cells[1], period)
        set_cell_text(row.cells[2], content)
        set_cell_text(row.cells[3], invest,
                      align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph('')
    return table
```

---

## 레시피 9. 페이지 헤더/푸터 + 페이지 번호

```python
def add_header_footer(doc, header_text, confidential=False):
    """헤더에 로고/제목, 푸터에 페이지 번호"""
    for section in doc.sections:
        # 헤더
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(header_text)
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

        if confidential:
            run = p.add_run('  |  CONFIDENTIAL')
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0xE7, 0x4C, 0x3C)
            run.font.bold = True

        # 푸터 — 페이지 번호
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 페이지 번호 필드
        run = p.add_run()
        fldChar1 = run._element.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'begin'})
        run._element.append(fldChar1)

        run2 = p.add_run()
        instrText = run2._element.makeelement(qn('w:instrText'), {})
        instrText.text = ' PAGE '
        run2._element.append(instrText)

        run3 = p.add_run()
        fldChar2 = run3._element.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'end'})
        run3._element.append(fldChar2)

        for r in [run, run2, run3]:
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
```

---

## 레시피 10. 자동 목차 (TOC) 생성

```python
from docx.oxml import OxmlElement

def add_table_of_contents(doc, title='목 차'):
    """
    자동 목차 삽입 — Word에서 열면 "필드 업데이트"로 자동 갱신됨
    (Heading 1~3 스타일 기반)
    """
    doc.add_heading(title, level=1)

    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' TOC \\o "1-3" \\h \\z \\u '  # Heading 1~3, 하이퍼링크, 페이지번호

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    run._element.append(fldChar1)
    run._element.append(instrText)
    run._element.append(fldChar2)

    # 플레이스홀더 텍스트
    placeholder = paragraph.add_run('[Word에서 이 문서를 열고 Ctrl+A → F9 를 누르면 목차가 자동 생성됩니다]')
    placeholder.font.size = Pt(9)
    placeholder.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    placeholder.font.italic = True

    run2 = paragraph.add_run()
    run2._element.append(fldChar3)

    doc.add_page_break()
```

---

## 레시피 11. 셀 패딩/여백 + 테두리 제어

```python
from docx.shared import Twips

def set_cell_margins(cell, top=0, bottom=0, start=0, end=0):
    """
    셀 내부 여백 설정 (단위: Twips, 1인치=1440Twips)

    권장값:
        - 일반 표: top=60, bottom=60, start=80, end=80
        - 넉넉한 표: top=80, bottom=80, start=120, end=120
    """
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for side, val in [('top', top), ('bottom', bottom), ('start', start), ('end', end)]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)

    tcPr.append(tcMar)


def set_table_borders(table, color='D5D5D5', size='4', style='single'):
    """표 전체 테두리 스타일 설정 (연한 회색 기본)"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')

    borders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), style)
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        border.set(qn('w:space'), '0')
        borders.append(border)

    tblPr.append(borders)


def remove_table_borders(table):
    """표 테두리 완전 제거 (인사이트 박스 등에 사용)"""
    set_table_borders(table, color='FFFFFF', size='0', style='none')


def set_cell_vertical_alignment(cell, align='center'):
    """셀 세로 정렬 (top/center/bottom)"""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align)
    tcPr.append(vAlign)


def add_padded_styled_table(doc, headers, rows, col_widths=None):
    """패딩이 넉넉한 프로페셔널 표 (기존 add_styled_table의 업그레이드 버전)"""
    table = add_styled_table(doc, headers, rows, col_widths)

    # 모든 셀에 여유로운 패딩 적용
    for row in table.rows:
        for cell in row.cells:
            set_cell_margins(cell, top=60, bottom=60, start=100, end=100)
            set_cell_vertical_alignment(cell, 'center')

    # 연한 회색 테두리
    set_table_borders(table, color='D5D5D5', size='4')

    return table
```

---

## 레시피 12. 프로세스 플로우 다이어그램

```python
def create_process_flow(doc, title, steps, save_path):
    """
    프로세스 흐름도 이미지 생성 후 워드 삽입

    steps 예시:
    [
        ('접수', '#27AE60'),
        ('PM 배정', '#2B579A'),
        ('견적 발송', '#2B579A'),
        ('고객 승인', '#F39C12'),
        ('공사 진행', '#1B3A5C'),
        ('완료', '#27AE60'),
    ]
    """
    fig, ax = plt.subplots(figsize=(8, 1.8))

    n = len(steps)
    box_w = 0.12
    gap = (1 - box_w * n) / (n + 1)

    for i, (label, color) in enumerate(steps):
        x = gap + i * (box_w + gap)

        # 둥근 박스
        bbox = dict(boxstyle='round,pad=0.4', facecolor=color, edgecolor='none', alpha=0.9)
        ax.text(x + box_w / 2, 0.5, label, ha='center', va='center',
                fontsize=9, fontweight='bold', color='white',
                bbox=bbox, transform=ax.transAxes)

        # 화살표 (마지막 제외)
        if i < n - 1:
            arrow_x = x + box_w + gap * 0.15
            ax.annotate('', xy=(arrow_x + gap * 0.7, 0.5), xytext=(arrow_x, 0.5),
                        xycoords='axes fraction', textcoords='axes fraction',
                        arrowprops=dict(arrowstyle='->', color='#AAAAAA', lw=2))

    ax.set_title(title, fontsize=12, fontweight='bold', color='#1B3A5C', pad=12)
    ax.axis('off')
    plt.tight_layout()
    plt.savefig(save_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()

    doc.add_picture(save_path, width=Inches(6))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[그림] {title}')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
```

---

## 레시피 13. 핵심 수치 강조 카드 (KPI 박스)

```python
def add_kpi_cards(doc, kpis):
    """
    핵심 수치를 시각적 카드로 나열

    kpis 예시:
    [
        ('9,890척', 'AS 관리 선박', '#1B3A5C'),
        ('371억$', '글로벌 AM 시장', '#2B579A'),
        ('4.6%', '연평균 성장률', '#27AE60'),
    ]
    """
    table = doc.add_table(rows=2, cols=len(kpis))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, (number, label, color) in enumerate(kpis):
        # 숫자 행 — 큰 Bold
        cell_num = table.rows[0].cells[i]
        r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:], 16)
        set_cell_shading(cell_num, color[1:])  # 배경색
        set_cell_text(cell_num, number, bold=True,
                      color_rgb=RGBColor(0xFF, 0xFF, 0xFF),
                      size=Pt(20), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_margins(cell_num, top=80, bottom=20, start=40, end=40)

        # 라벨 행 — 작은 텍스트
        cell_label = table.rows[1].cells[i]
        set_cell_shading(cell_label, color[1:])
        set_cell_text(cell_label, label, bold=False,
                      color_rgb=RGBColor(0xDD, 0xDD, 0xDD),
                      size=Pt(9), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_margins(cell_label, top=0, bottom=80, start=40, end=40)

    # 테두리 제거 (카드 느낌)
    remove_table_borders(table)
    doc.add_paragraph('')
    return table
```

---

## 레시피 14. 섹션 구분선 + 빈 줄 유틸

```python
def add_divider(doc, style='thin'):
    """섹션 구분선 (thin/thick/double)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if style == 'thin':
        run = p.add_run('─' * 60)
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    elif style == 'thick':
        run = p.add_run('━' * 40)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    elif style == 'double':
        run = p.add_run('═' * 40)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x2B, 0x57, 0x9A)

    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)


def add_spacer(doc, height_pt=12):
    """지정된 높이의 빈 줄 삽입"""
    p = doc.add_paragraph('')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run('')
    run.font.size = Pt(height_pt)


def add_caption(doc, text, number=None, prefix='표'):
    """표/그림 캡션 — '[표 1] 제목' 형태"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    label = f'[{prefix} {number}] {text}' if number else f'[{prefix}] {text}'
    run = p.add_run(label)
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
```

---

## 레시피 15. 전체 조합 예시 — AM 전략 보고서 (확장판)

```python
def generate_report():
    doc = Document()
    init_document(doc)

    # ── 표지 ──
    add_cover(doc,
              title='선도기업 AM 전략 벤치마킹을 통한\nHD현대마린솔루션테크의 미래 방향 제안',
              subtitle='바르질라·콩스버그 사례 분석 및 웹 포털 구축 제안',
              date='2026년 4월',
              author='홍길동',
              department='기술서비스팀')

    # ── 헤더/푸터 ──
    add_header_footer(doc, 'HD현대마린솔루션테크 | AM 전략 제안', confidential=True)

    # ── 목차 ──
    add_table_of_contents(doc)

    # ── Executive Summary ──
    doc.add_heading('Executive Summary', level=1)

    # KPI 카드 — 핵심 수치를 시각적으로 먼저 보여줌
    add_kpi_cards(doc, [
        ('9,890척', 'AS 관리 선박', '#1B3A5C'),
        ('371억$', '글로벌 AM 시장 (2024)', '#2B579A'),
        ('532억$', '글로벌 AM 시장 (2032)', '#27AE60'),
    ])

    add_insight_box(doc,
        '선도기업과의 차이는 기술력이 아닙니다.\n'
        '우리는 9,890척의 AS 네트워크와 최고의 기술력을 이미 갖고 있습니다.\n'
        '부족한 건 고객이 그 서비스에 접근하는 "디지털 창구"입니다.')

    # ── 시장 현황 ──
    doc.add_heading('1. AM 시장 현황', level=1)

    # 주석이 달린 그래프 — 현재/목표가 한눈에 보임
    create_annotated_line_chart(doc,
        title='글로벌 선박 수리·보수 시장 규모 (억 달러)',
        x_data=[2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032],
        y_data=[371, 389, 408, 428, 449, 470, 493, 517, 532],
        annotations=[
            (2024, 371, '현재\n371억$'),
            (2028, 449, '연평균\n4.6% 성장'),
            (2032, 532, '2032년\n532억$'),
        ],
        save_path='chart_market_annotated.png')
    add_caption(doc, '글로벌 선박 수리·보수 시장 성장 전망 (2024~2032)', number=1, prefix='그림')

    add_spacer(doc, 6)

    add_padded_styled_table(doc,
        headers=['시장', '2024년', '2032년', '연평균 성장률'],
        rows=[
            ['선박 수리·보수', '371억$ (55조원)', '532억$ (78조원)', '4.6%'],
            ['해양 소프트웨어', '29.9억$', '78.4억$', '**11.3%**'],
        ])
    add_caption(doc, 'AM 관련 시장 규모 비교', number=1, prefix='표')

    # ── 벤치마킹 요약 ──
    doc.add_heading('2. 선도기업 벤치마킹 요약', level=1)

    # 경쟁사 비교 막대 차트
    create_comparison_bar_chart(doc,
        title='선도기업 vs 우리 — 디지털 서비스 역량 비교 (10점 만점)',
        categories=['웹 포털', '서비스 온라인화', '리포트 시각화', '통합 플랫폼', 'AI/데이터', 'AS 네트워크'],
        group1_vals=[9, 9, 8, 8, 7, 7],
        group2_vals=[1, 2, 2, 5, 7, 9],
        group1_label='바르질라/콩스버그',
        group2_label='HD현대마린솔루션(테크)',
        save_path='chart_comparison.png')
    add_caption(doc, '선도기업 대비 디지털 서비스 역량 비교', number=2, prefix='그림')

    # ── Gap 분석 ──
    doc.add_heading('3. Gap 분석', level=1)
    add_gap_table(doc,
        headers=['영역', '바르질라/콩스버그', 'HD현대마린솔루션(테크)', 'Gap'],
        rows=[
            ['고객용 웹 포털', 'Wärtsilä Online, K-IMS', '없음 (전화/이메일)', '🔴 큼'],
            ['서비스 요청 온라인화', '웹에서 접수→추적→이력', '전화/메일로 접수', '🔴 큼'],
            ['리포트 시각화', '데이터 대시보드+그래프', '텍스트/숫자 나열', '🔴 큼'],
            ['통합 플랫폼', '레거시 통합 완료/진행 중', 'Hi-4S + ISS 2.0 (분리)', '🟡 중간'],
            ['AI/데이터 분석', '예지정비, 성능 최적화', 'OceanWise, ISS 2.0', '🟢 양호'],
            ['AS 네트워크/기술력', '글로벌 네트워크 보유', '9,890척 독점 AS', '🟢 우위'],
        ])
    add_caption(doc, '선도기업 대비 Gap 분석', number=2, prefix='표')

    add_insight_box(doc,
        '기술력과 네트워크는 선도기업 이상\n'
        '→ 부족한 건 고객이 우리 서비스에 접근하는 "디지털 창구"')

    # ── 제안: 서비스 요청 프로세스 ──
    doc.add_heading('4. 웹 포털 핵심 기능', level=1)
    doc.add_heading('서비스 요청 프로세스', level=2)

    create_process_flow(doc, '서비스 요청 → 완료 흐름', [
        ('웹 접수', '#27AE60'),
        ('PM 배정', '#2B579A'),
        ('견적 발송', '#2B579A'),
        ('고객 승인', '#F39C12'),
        ('공사 진행', '#1B3A5C'),
        ('완료', '#27AE60'),
    ], save_path='flow_service.png')
    add_caption(doc, '웹 포털 기반 서비스 요청 프로세스', number=3, prefix='그림')

    # ── 리포트 시각화 예시 ──
    doc.add_heading('5. 리포트 시각화 대시보드', level=1)

    # 게이지 차트 예시
    create_gauge_chart(doc, '배전반 종합 건강도', 78, save_path='gauge_health.png')
    add_caption(doc, '배전반 건강도 게이지 차트 예시 (78/100점)', number=4, prefix='그림')

    # ── 기대효과 Before/After ──
    add_divider(doc, 'thick')
    doc.add_heading('6. 기대효과', level=1)
    add_before_after_table(doc, '고객 관점 변화', [
        ('서비스 요청', '전화/이메일 (업무시간)', '웹 포털 24/7 (즉시)'),
        ('호선정보 확인', '전화 → 담당자 확인 → 회신', '웹에서 즉시 검색'),
        ('서비스 이력', '요청 → 내부 조회 → 회신', '웹에서 직접 조회'),
        ('리포트 품질', '텍스트/숫자 나열', '그래프/대시보드 시각화'),
        ('견적 회신', '2~3일', '당일 (AI 자동 산출)'),
    ])

    # ── 로드맵 ──
    doc.add_heading('7. 구현 로드맵', level=1)
    add_roadmap_table(doc, [
        ('Phase 1: 즉시', '1~3개월', 'PM 업무 AI 자동화 + 리포트 시각화', '낮음'),
        ('Phase 2: 단기', '3~6개월', '웹 포털 MVP (호선정보, 서비스 요청)', '중간'),
        ('Phase 3: 중기', '6~12개월', '리포트 대시보드 + 문서 라이브러리', '중간'),
        ('Phase 4: 확장', '12개월~', '모회사 Hi-4S/ISS 2.0 연동', '그룹 협의'),
    ])

    # ── 마무리 ──
    add_divider(doc, 'thick')
    doc.add_heading('마무리', level=1)
    add_insight_box(doc,
        '웹 포털 하나가 그 창구가 됩니다.\n'
        '고객은 24/7 서비스에 접근하고, 우리는 데이터를 축적하고,\n'
        '리포트는 시각화되어 수주를 앞당깁니다.\n\n'
        '그 창구를 만드는 것이 이 제안의 핵심입니다.',
        bg_color='F0F8E8', border_color='27AE60')

    doc.save('AM전략_벤치마킹_보고서.docx')
    print('보고서 생성 완료: AM전략_벤치마킹_보고서.docx')

# 실행
generate_report()
```

---

## 요약 — 레시피 맵

| # | 레시피 | 함수명 | 용도 |
|---|--------|--------|------|
| 1 | 문서 초기 설정 | `init_document()` | 폰트, 여백, 줄간격 |
| 2 | 표지 | `add_cover()` | 프로페셔널 표지 |
| 3 | 스타일 표 | `add_styled_table()` | 일반 데이터 표 |
| 4 | Gap 분석 표 | `add_gap_table()` | 🔴🟡🟢 컬러 코딩 |
| 5 | 인사이트 박스 | `add_insight_box()` | 핵심 메시지 강조 |
| 6 | Before/After 표 | `add_before_after_table()` | 변화 비교 |
| 7 | 그래프 (꺾은선/막대/도넛) | `create_and_insert_*_chart()` | 기본 데이터 시각화 |
| 7.5 | **고급 차트** | `create_annotated_line_chart()` | **주석 달린 꺾은선** |
| 7.5 | **비교 막대 차트** | `create_comparison_bar_chart()` | **경쟁사 나란히 비교** |
| 7.5 | **게이지 차트** | `create_gauge_chart()` | **건강도 점수 반원형** |
| 8 | 로드맵 표 | `add_roadmap_table()` | Phase 타임라인 |
| 9 | 헤더/푸터 | `add_header_footer()` | 페이지 번호, 제목 |
| 10 | **자동 목차 (TOC)** | `add_table_of_contents()` | **Heading 기반 자동 목차** |
| 11 | **셀 패딩/테두리** | `set_cell_margins()`, `set_table_borders()` | **표 여백·테두리 제어** |
| 11 | **패딩 표** | `add_padded_styled_table()` | **넉넉한 여백의 프로 표** |
| 12 | **프로세스 플로우** | `create_process_flow()` | **단계별 흐름도** |
| 13 | **KPI 카드** | `add_kpi_cards()` | **핵심 수치 시각 카드** |
| 14 | **구분선/여백/캡션** | `add_divider()`, `add_spacer()`, `add_caption()` | **레이아웃 유틸** |
| 15 | 전체 조합 (확장판) | `generate_report()` | 보고서 완성본 |

---

## 참고 자료 (추가)

- [python-docx TOC 생성 가이드 (2026)](https://copyprogramming.com/howto/update-the-toc-table-of-content-of-ms-word-docx-documents-with-python)
- [python-docx Section Breaks 가이드](https://pytutorial.com/python-docx-section-breaks-advanced-layout-control/)
- [python-docx Cell Margins Issue #641](https://github.com/python-openxml/python-docx/issues/641)
- [python-docx Table Cell Borders Issue #433](https://github.com/python-openxml/python-docx/issues/433)
- [From Default Chart to Journal-Quality Infographics](https://towardsdatascience.com/from-default-python-line-chart-to-journal-quality-infographics-80e3949eacc3/)
- [How to Make Extremely Beautiful Charts with Python](https://wangari.substack.com/p/how-to-make-extremely-beautiful-charts)
- [Matplotlib Style Sheets Reference](https://matplotlib.org/stable/gallery/style_sheets/style_sheets_reference.html)
- [Effective Use of White Space in Reports](https://www.moreport.co/insights/design/effective-use-of-white-space-in-reports)
- [Visual Hierarchy in Word Documents (LinkedIn)](https://www.linkedin.com/advice/0/how-do-you-use-white-space-hierarchy-enhance)
- [10 DOs and DON'Ts of Report Design (Datylon)](https://www.datylon.com/blog/10-report-design-tips)
- [12 Visual Hierarchy Principles (Visme)](https://visme.co/blog/visual-hierarchy/)
- [개조식 보고서 작성법](https://0409-care.senior-health-life.kr/entry/%EA%B0%9C%EC%A1%B0%EC%8B%9D%EC%9D%B4%EB%9E%80-%EC%9E%91%EC%84%B1-%EB%B3%B4%EA%B3%A0%EC%84%9C-%EB%9C%BB-%ED%91%9C%EA%B8%B0-%EB%AC%B8%EC%9E%A5-%EC%98%88%EC%8B%9C-%EC%95%8C%EC%95%84%EB%B3%B4%EA%B8%B0)
- [한 번에 통과되는 보고서 작성법 (PDF)](https://www.classpod.co.kr/prohub_file/uploadfiles/tutor/1672148327pro00089315.pdf)
