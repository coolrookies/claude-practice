from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.dml import MSO_THEME_COLOR
import copy
from lxml import etree

# ── 색상 상수 ──────────────────────────────────
C_WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
C_BG          = RGBColor(0xFF, 0xFF, 0xFF)       # 슬라이드 배경
C_TOP_BAR_L   = RGBColor(0xC9, 0xB8, 0xE0)      # 상단 바 왼쪽 (연보라)
C_TOP_BAR_R   = RGBColor(0xE8, 0xDD, 0xD0)      # 상단 바 오른쪽 (베이지)
C_PURPLE_DK   = RGBColor(0x6D, 0x4F, 0x96)      # 소제목 보라
C_PURPLE_MD   = RGBColor(0xA9, 0x8E, 0xC5)      # 섹션 라벨
C_PURPLE_LT   = RGBColor(0xC9, 0xB8, 0xE0)      # 언더라인·구분선
C_PURPLE_BG   = RGBColor(0xF3, 0xEE, 0xF9)      # 키워드 배경
C_TITLE_TEXT  = RGBColor(0x2A, 0x22, 0x30)      # 제목 텍스트
C_BODY_TEXT   = RGBColor(0x4A, 0x43, 0x50)      # 본문 텍스트
C_MUTED       = RGBColor(0x6A, 0x60, 0x70)      # 보조 텍스트
C_CARD_BG     = RGBColor(0xFA, 0xF8, 0xF5)      # 카드 배경
C_CARD_BDR    = RGBColor(0xE8, 0xDD, 0xD0)      # 카드 테두리
C_RED_DARK    = RGBColor(0x7A, 0x20, 0x30)      # 건성피부 카드 헤더
C_BAR_LIGHT   = RGBColor(0xD4, 0xC8, 0xE0)      # 차트 바 (이전)
C_BAR_DARK    = RGBColor(0x7A, 0x5F, 0xA0)      # 차트 바 (이후)
C_STEP_BDR    = RGBColor(0xC9, 0xB8, 0xE0)      # 루틴 스텝 테두리
C_STEP_BG     = RGBColor(0xF3, 0xEE, 0xF9)      # 루틴 스텝 배경
C_DIVIDER     = RGBColor(0xED, 0xE6, 0xDB)      # 구분선
C_REF_TEXT    = RGBColor(0x6A, 0x5F, 0x78)      # 참고문헌 텍스트
C_COVER_BG1   = RGBColor(0xFD, 0xFC, 0xFB)
C_COVER_BG2   = RGBColor(0xED, 0xE6, 0xF5)

W = Inches(13.33)   # 16:9 width
H = Inches(7.5)     # 16:9 height

# ── 헬퍼 ──────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, l, t, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.width = line_width
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width == Pt(0):
            shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, text, l, t, w, h,
                 font_name="맑은 고딕", font_size=Pt(11), bold=False,
                 color=C_BODY_TEXT, align=PP_ALIGN.LEFT,
                 wrap=True, v_anchor=None):
    txBox = slide.shapes.add_textbox(l, t, w, h)
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    if v_anchor:
        tf.vertical_anchor = v_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_multiline_textbox(slide, lines, l, t, w, h,
                          font_name="맑은 고딕", font_size=Pt(10.5),
                          bold=False, color=C_BODY_TEXT,
                          align=PP_ALIGN.LEFT, line_spacing=Pt(4)):
    """lines = list of (text, bold, color, size) or just strings"""
    txBox = slide.shapes.add_textbox(l, t, w, h)
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            txt, b, c, sz = item, bold, color, font_size
        else:
            txt = item.get('text', '')
            b   = item.get('bold', bold)
            c   = item.get('color', color)
            sz  = item.get('size', font_size)
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = line_spacing
        run = p.add_run()
        run.text = txt
        run.font.name = font_name
        run.font.size = sz
        run.font.bold = b
        run.font.color.rgb = c
    return txBox


def top_bar(slide):
    """상단 5px 그라데이션 바 (단색 2분할로 근사)"""
    bar_h = Inches(0.06)
    add_rect(slide, 0, 0, W * 0.55, bar_h, fill=C_TOP_BAR_L)
    add_rect(slide, W * 0.55, 0, W * 0.45, bar_h, fill=C_TOP_BAR_R)


def section_label(slide, text, l=Inches(0.6), t=Inches(0.22)):
    add_text_box(slide, text, l, t, Inches(5), Inches(0.25),
                 font_size=Pt(9), bold=True, color=C_PURPLE_MD,
                 align=PP_ALIGN.LEFT)


def slide_title(slide, text, l=Inches(0.6), t=Inches(0.5)):
    add_text_box(slide, text, l, t, Inches(8), Inches(0.5),
                 font_size=Pt(22), bold=True, color=C_TITLE_TEXT)
    # 언더라인
    add_rect(slide, l, t + Inches(0.52), Inches(0.45), Inches(0.035),
             fill=C_PURPLE_LT)


def sub_heading(slide, text, l, t, w=Inches(5.8)):
    line = add_rect(slide, l, t + Inches(0.02), Inches(0.04), Inches(0.22),
                    fill=C_PURPLE_LT)
    add_text_box(slide, text, l + Inches(0.1), t, w, Inches(0.28),
                 font_size=Pt(11.5), bold=True, color=C_PURPLE_DK)


def keyword_bar(slide, keywords):
    """하단 키워드 바"""
    bar_t = H - Inches(0.45)
    add_rect(slide, 0, bar_t, W, Inches(0.45), fill=C_CARD_BG)
    add_rect(slide, 0, bar_t, W, Inches(0.01), fill=C_DIVIDER)

    x = Inches(0.6)
    add_text_box(slide, "KEY", x, bar_t + Inches(0.12), Inches(0.5), Inches(0.25),
                 font_size=Pt(8), bold=True, color=C_PURPLE_MD)
    x += Inches(0.65)

    for kw in keywords:
        w_est = Inches(len(kw) * 0.14 + 0.3)
        box = add_rect(slide, x, bar_t + Inches(0.1), w_est, Inches(0.26),
                       fill=C_PURPLE_BG)
        add_text_box(slide, kw, x + Inches(0.1), bar_t + Inches(0.12),
                     w_est - Inches(0.1), Inches(0.22),
                     font_size=Pt(9), bold=False, color=C_PURPLE_DK)
        x += w_est + Inches(0.1)


def slide_number(slide, num, total=6):
    add_text_box(slide, f"{num:02d} / {total:02d}",
                 W - Inches(1.0), H - Inches(0.28), Inches(0.9), Inches(0.22),
                 font_size=Pt(9), color=C_PURPLE_LT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank_layout = prs.slide_layouts[6]  # 완전 빈 레이아웃


# ══ SLIDE 1 : 표지 ════════════════════════════
s1 = prs.slides.add_slide(blank_layout)
# 배경 그라데이션 근사 (두 직사각형)
add_rect(s1, 0, 0, W, H, fill=C_COVER_BG1)
add_rect(s1, W * 0.55, 0, W * 0.45, H, fill=C_COVER_BG2)

top_bar(s1)

# 장식 원 (연보라)
circ = s1.shapes.add_shape(9, W - Inches(3.8), -Inches(0.8), Inches(4.2), Inches(4.2))
circ.fill.solid(); circ.fill.fore_color.rgb = RGBColor(0xED, 0xE6, 0xF5)
circ.line.fill.background()

# 학과 라벨
add_text_box(s1, "뷰티아트산업학과 · 피부미용",
             Inches(2.5), Inches(1.9), Inches(8.3), Inches(0.35),
             font_size=Pt(11), bold=True, color=C_PURPLE_MD, align=PP_ALIGN.CENTER)

# 메인 제목
add_text_box(s1, "본인이 사용하는 화장품 및 관리 방법",
             Inches(1.0), Inches(2.4), Inches(11.33), Inches(0.9),
             font_size=Pt(32), bold=True, color=C_TITLE_TEXT, align=PP_ALIGN.CENTER)

# 부제목
add_text_box(s1, "나의 피부 유형에 맞는 보습 루틴과 제품 사용 경험",
             Inches(1.0), Inches(3.4), Inches(11.33), Inches(0.4),
             font_size=Pt(13), color=C_MUTED, align=PP_ALIGN.CENTER)

# 구분선
add_rect(s1, Inches(5.92), Inches(3.9), Inches(1.5), Inches(0.04), fill=C_PURPLE_LT)

# 메타 정보
for i, (label, val) in enumerate([
    ("학과", "뷰티아트산업학과"),
    ("이름", "이현주"),
    ("학번", "2024215607"),
]):
    y = Inches(4.1) + i * Inches(0.34)
    add_text_box(s1, label, Inches(4.5), y, Inches(0.8), Inches(0.3),
                 font_size=Pt(10), color=C_MUTED, align=PP_ALIGN.RIGHT)
    add_text_box(s1, val, Inches(5.5), y, Inches(3.0), Inches(0.3),
                 font_size=Pt(10.5), bold=True, color=C_TITLE_TEXT)

slide_number(s1, 1)


# ══ SLIDE 2 : 나의 피부 상태 ══════════════════
s2 = prs.slides.add_slide(blank_layout)
add_rect(s2, 0, 0, W, H, fill=C_BG)
top_bar(s2)
section_label(s2, "My Skin Condition")
slide_title(s2, "나의 피부 상태")

# 본문 (왼쪽 영역)
LX = Inches(0.6)
body_top = Inches(1.25)
sec_gap  = Inches(0.08)

# ① 피부 유형
sub_heading(s2, "피부 유형 : 건성 피부", LX, body_top)
add_multiline_textbox(s2, [
    "나의 피부는 건성 피부로, 피지 분비가 적고 피부 장벽이 약한 편입니다.",
    "계절이나 생활환경에 따라 피부 상태가 민감하게 반응하며,",
    "특히 세안 후 보습을 적절히 해주지 않으면 피부가 빠르게 당기는 느낌이 납니다.",
], LX, body_top + Inches(0.32), Inches(7.3), Inches(0.65),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# ② 주요 피부 고민
sh2_t = body_top + Inches(1.05)
sub_heading(s2, "주요 피부 고민", LX, sh2_t)
add_multiline_textbox(s2, [
    "세안 후 피부 당김이 자주 느껴지며, 겉으로는 유분이 없어 보이지만",
    "피부 속은 건조한 속건조 상태가 지속됩니다.",
    "특히 실내 냉난방 환경에서 더 심해지는 경향이 있어",
    "하루 중 수분감이 떨어지는 시간이 빠른 편입니다.",
], LX, sh2_t + Inches(0.32), Inches(7.3), Inches(0.8),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# ③ 관리 방향
sh3_t = sh2_t + Inches(1.22)
sub_heading(s2, "관리 방향", LX, sh3_t)
add_multiline_textbox(s2, [
    "자극이 적고 보습력이 좋은 제품을 중심으로 스킨케어 루틴을 구성하고 있으며,",
    "피부 장벽을 강화하면서 하루 종일 수분감을 유지하는 것을 목표로 관리하고 있습니다.",
], LX, sh3_t + Inches(0.32), Inches(7.3), Inches(0.55),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# 오른쪽 건성피부 특징 카드
cx = Inches(8.3)
cy = Inches(1.1)
cw = Inches(4.5)
ch = Inches(5.6)
add_rect(s2, cx, cy, cw, ch, fill=C_CARD_BG, line_color=C_CARD_BDR, line_width=Pt(0.75))
# 카드 헤더
add_rect(s2, cx + Inches(0.2), cy + Inches(0.18), cw - Inches(0.4), Inches(0.38), fill=C_RED_DARK)
add_text_box(s2, "건성 피부의 특징",
             cx + Inches(0.2), cy + Inches(0.18), cw - Inches(0.4), Inches(0.38),
             font_size=Pt(11.5), bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

chars = [
    "피부가 자주 당기는 느낌이 난다.",
    "가려워서 긁고 싶은 느낌이 빈번하게 나타난다.",
    "손으로 살짝 만졌을 때 매끄럽지 않고\n거친 느낌이 난다.",
    "각질이 심하게 발생한다.",
    "주변 환경에 따라서 피부가\n예민하게 반응한다.",
]
for i, txt in enumerate(chars):
    iy = cy + Inches(0.72) + i * Inches(0.9)
    add_text_box(s2, f"{i+1}.", cx + Inches(0.22), iy, Inches(0.28), Inches(0.5),
                 font_size=Pt(11), bold=True, color=C_RED_DARK)
    add_text_box(s2, txt, cx + Inches(0.5), iy, cw - Inches(0.65), Inches(0.55),
                 font_size=Pt(10.5), color=C_BODY_TEXT)

keyword_bar(s2, ["건성 피부", "속건조", "피부 당김", "수분 공급", "보습 관리", "피부 장벽"])
slide_number(s2, 2)


# ══ SLIDE 3 : 현재 사용 중인 화장품 ════════════
s3 = prs.slides.add_slide(blank_layout)
add_rect(s3, 0, 0, W, H, fill=C_BG)
top_bar(s3)
section_label(s3, "Products I Use")
slide_title(s3, "현재 사용 중인 화장품")

LX = Inches(0.6)

# ① 디마프 세럼
sub_heading(s3, "① 디마프 히어로 마이 퍼스트 세럼", LX, Inches(1.25))
body1 = (
    "가장 먼저 사용하게 된 제품으로, 건성 피부의 수분 보충을 목적으로 도입하였습니다.\n"
    "세럼 타입으로 피부에 가볍게 흡수되는 것이 특징입니다.\n"
    "주요 성분으로는 수분을 끌어당겨 피부 표면에 수분막을 형성하는 히알루론산,\n"
    "피부 보습과 장벽 기능을 돕는 판테놀, 수분 손실을 억제하는 세라마이드,\n"
    "피부를 부드럽게 가꾸는 호호바씨오일과 마카다미아씨오일이 포함되어 있습니다."
)
add_multiline_textbox(s3, body1.split("\n"), LX, Inches(1.57), Inches(7.3), Inches(1.35),
                      font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# ② 보습 크림
sub_heading(s3, "② 보습 크림", LX, Inches(3.0))
add_multiline_textbox(s3, [
    "세럼 사용 후 마지막 단계에서 수분이 날아가지 않도록 밀봉하는 역할로 사용하고 있습니다.",
    "피부에 수분을 잡아두는 효과를 위해 꾸준히 활용 중입니다.",
], LX, Inches(3.32), Inches(7.3), Inches(0.55),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# 성분 태그
comp_y = Inches(4.0)
add_text_box(s3, "주요 성분", LX, comp_y, Inches(1.0), Inches(0.28),
             font_size=Pt(9.5), bold=True, color=C_PURPLE_MD)
comps = ["히알루론산", "판테놀", "세라마이드", "호호바씨오일", "마카다미아씨오일"]
cx2 = LX + Inches(1.15)
for c in comps:
    w_c = Inches(len(c) * 0.155 + 0.35)
    add_rect(s3, cx2, comp_y - Inches(0.01), w_c, Inches(0.3), fill=C_PURPLE_BG, line_color=C_PURPLE_LT, line_width=Pt(0.5))
    add_text_box(s3, c, cx2 + Inches(0.1), comp_y + Inches(0.02), w_c, Inches(0.25),
                 font_size=Pt(9.5), color=C_PURPLE_DK)
    cx2 += w_c + Inches(0.1)

# 오른쪽 제품 이미지 영역
px = Inches(8.3)
py = Inches(1.1)
pw = Inches(4.5)
ph = Inches(5.6)
add_rect(s3, px, py, pw, ph, fill=RGBColor(0xF5, 0xF2, 0xEE), line_color=C_CARD_BDR, line_width=Pt(0.75))
# 이미지 안내 문구
add_text_box(s3,
    "De:maf\nHere-Oh My First Serum",
    px, py + Inches(1.8), pw, Inches(0.9),
    font_size=Pt(13), bold=True, color=C_PURPLE_DK, align=PP_ALIGN.CENTER)
add_text_box(s3,
    "※ 이 자리에 제품 사진(demaf_product.jpg)을\n삽입하세요",
    px, py + Inches(2.8), pw, Inches(0.6),
    font_size=Pt(9), color=C_MUTED, align=PP_ALIGN.CENTER)

keyword_bar(s3, ["히알루론산", "판테놀", "세라마이드", "호호바씨오일", "마카다미아씨오일", "보습 크림"])
slide_number(s3, 3)


# ══ SLIDE 4 : 화장품 사용 방법 ════════════════
s4 = prs.slides.add_slide(blank_layout)
add_rect(s4, 0, 0, W, H, fill=C_BG)
top_bar(s4)
section_label(s4, "How I Use It")
slide_title(s4, "화장품 사용 방법")

LX = Inches(0.6)

# Step 1
sub_heading(s4, "Step 1 · 세럼 기본 사용", LX, Inches(1.25))
add_multiline_textbox(s4, [
    "세안 후 피부를 가볍게 정리한 다음, 디마프 히어로 마이 퍼스트 세럼을 손에 덜어",
    "얼굴 전체에 고루 펴 발라줍니다. 적당량을 사용하여 피부에 스며들도록",
    "가볍게 눌러 흡수시키는 방식으로 적용합니다.",
], LX, Inches(1.57), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# Step 2
sub_heading(s4, "Step 2 · 레이어링", LX, Inches(2.42))
add_multiline_textbox(s4, [
    "피부 상태에 따라 같은 세럼을 2~3회 반복해서 겹쳐 바르는 레이어링 방법을 활용합니다.",
    "피부가 특히 건조한 날에는 레이어링 횟수를 늘려 수분감을 보충하며,",
    "각 레이어가 충분히 흡수된 후 다음 단계로 넘어갑니다.",
], LX, Inches(2.74), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# Step 3
sub_heading(s4, "Step 3 · 보습 크림 혼합 사용", LX, Inches(3.6))
add_multiline_textbox(s4, [
    "현재 사용 중인 보습 크림은 제형이 다소 뻑뻑하여 단독 사용 시 발림성이 떨어졌습니다.",
    "이를 보완하기 위해 보습 크림에 디마프 세럼을 소량 혼합하여 사용합니다.",
    "두 제품을 함께 사용하면 크림의 제형이 부드러워지고 흡수가 자연스럽게 이루어집니다.",
], LX, Inches(3.92), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# 오른쪽 루틴 다이어그램
dx = Inches(8.3)
dy = Inches(1.1)
dw = Inches(4.5)

add_rect(s4, dx, dy, dw, Inches(5.6), fill=C_CARD_BG, line_color=C_CARD_BDR, line_width=Pt(0.75))

# 다이어그램 라벨
add_text_box(s4, "Skincare Routine", dx, dy + Inches(0.15), dw, Inches(0.3),
             font_size=Pt(9), bold=True, color=C_PURPLE_MD, align=PP_ALIGN.CENTER)

steps = [
    ("STEP 1", "세안", "피부 정돈", False),
    ("STEP 2", "세럼 도포", "디마프 세럼 1회", True),
    ("STEP 3", "레이어링", "건조 시 2~3회 중첩", True),
    ("STEP 4", "보습 크림", "세럼 소량 혼합 후 도포", False),
]
step_h = Inches(0.9)
for i, (label, main, sub, highlight) in enumerate(steps):
    sy = dy + Inches(0.55) + i * Inches(1.15)
    bg  = C_STEP_BG if highlight else C_WHITE
    bdr = C_PURPLE_DK if highlight else C_STEP_BDR
    bw  = Pt(1.5) if highlight else Pt(0.75)
    add_rect(s4, dx + Inches(0.3), sy, dw - Inches(0.6), step_h, fill=bg, line_color=bdr, line_width=bw)
    add_text_box(s4, label, dx + Inches(0.3), sy + Inches(0.06), dw - Inches(0.6), Inches(0.22),
                 font_size=Pt(8.5), bold=True,
                 color=C_PURPLE_DK if highlight else C_PURPLE_MD, align=PP_ALIGN.CENTER)
    add_text_box(s4, main, dx + Inches(0.3), sy + Inches(0.28), dw - Inches(0.6), Inches(0.3),
                 font_size=Pt(12), bold=True, color=C_TITLE_TEXT, align=PP_ALIGN.CENTER)
    add_text_box(s4, sub, dx + Inches(0.3), sy + Inches(0.58), dw - Inches(0.6), Inches(0.25),
                 font_size=Pt(9.5), color=C_MUTED, align=PP_ALIGN.CENTER)
    # 화살표
    if i < 3:
        add_text_box(s4, "▼", dx + Inches(0.3), sy + step_h + Inches(0.05),
                     dw - Inches(0.6), Inches(0.22),
                     font_size=Pt(10), color=C_PURPLE_LT, align=PP_ALIGN.CENTER)

keyword_bar(s4, ["세안 후 사용", "레이어링", "2~3회 중첩", "크림 혼합", "발림성 개선", "흡수력"])
slide_number(s4, 4)


# ══ SLIDE 5 : 사용 후 느낀 변화 ══════════════
s5 = prs.slides.add_slide(blank_layout)
add_rect(s5, 0, 0, W, H, fill=C_BG)
top_bar(s5)
section_label(s5, "Changes I Noticed")
slide_title(s5, "사용 후 느낀 변화")

LX = Inches(0.6)

# 본문
sub_heading(s5, "피부 당김 감소", LX, Inches(1.25))
add_multiline_textbox(s5, [
    "세안 후 느껴지던 피부 당김이 줄어들었습니다. 이전에는 세안 직후부터",
    "피부가 팽팽하게 당기는 느낌이 강했는데, 현재 루틴을 꾸준히 실천하면서",
    "그 불편함이 많이 완화되었습니다.",
], LX, Inches(1.57), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

sub_heading(s5, "피부 촉촉함 증가 및 결 개선", LX, Inches(2.42))
add_multiline_textbox(s5, [
    "하루 중 피부가 건조해지는 시간이 눈에 띄게 늦어졌습니다. 수분감이 지속되는",
    "시간이 더 길어진 것을 느꼈으며, 피부 결이 전보다 고르게 정돈된 느낌이 듭니다.",
    "화장이 얼굴에 자연스럽게 밀착되는 것도 개선된 부분 중 하나입니다.",
], LX, Inches(2.74), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

sub_heading(s5, "피부의 전반적인 편안함", LX, Inches(3.6))
add_multiline_textbox(s5, [
    "자극이 적고 가벼운 세럼을 중심으로 루틴을 구성한 덕분에, 예민해지거나",
    "트러블이 생기는 일 없이 피부가 전반적으로 이전보다 편안하게 느껴집니다.",
    "세럼과 크림을 혼합하여 사용하는 방식이 보습 효과를 높이는 데 도움이 됩니다.",
], LX, Inches(3.92), Inches(7.3), Inches(0.72),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(2))

# 오른쪽 전후 비교 차트
bx = Inches(8.3)
by = Inches(1.1)
bw = Inches(4.5)
bh = Inches(5.6)
add_rect(s5, bx, by, bw, bh, fill=C_CARD_BG, line_color=C_CARD_BDR, line_width=Pt(0.75))

add_text_box(s5, "사용 전후 비교 (10점 만점)", bx, by + Inches(0.15), bw, Inches(0.3),
             font_size=Pt(9.5), bold=True, color=C_PURPLE_MD, align=PP_ALIGN.CENTER)

# 범례
add_rect(s5, bx + Inches(0.4), by + Inches(0.6), Inches(0.18), Inches(0.18), fill=C_BAR_LIGHT)
add_text_box(s5, "기타 제품", bx + Inches(0.62), by + Inches(0.59), Inches(0.9), Inches(0.22),
             font_size=Pt(8.5), color=C_MUTED)
add_rect(s5, bx + Inches(1.7), by + Inches(0.6), Inches(0.18), Inches(0.18), fill=C_BAR_DARK)
add_text_box(s5, "디마프 사용", bx + Inches(1.92), by + Inches(0.59), Inches(0.9), Inches(0.22),
             font_size=Pt(8.5), color=C_MUTED)

# 차트: 피부 당김
charts = [
    ("피부 당김 감소",    5, 8,  Inches(1.0)),
    ("피부 촉촉함 증가",  5, 9,  Inches(2.9)),
]
max_bar_w = bw - Inches(1.8)

for (cat, before, after, cy_offset) in charts:
    cy_ = by + cy_offset
    add_text_box(s5, cat, bx + Inches(0.25), cy_, bw - Inches(0.4), Inches(0.28),
                 font_size=Pt(10), bold=True, color=C_TITLE_TEXT)
    # 이전
    add_text_box(s5, "기타", bx + Inches(0.25), cy_ + Inches(0.32), Inches(0.5), Inches(0.22),
                 font_size=Pt(8.5), color=C_MUTED, align=PP_ALIGN.RIGHT)
    bar_bg_w = max_bar_w
    add_rect(s5, bx + Inches(0.82), cy_ + Inches(0.32), bar_bg_w, Inches(0.22),
             fill=RGBColor(0xED, 0xE8, 0xF0))
    add_rect(s5, bx + Inches(0.82), cy_ + Inches(0.32), bar_bg_w * before / 10, Inches(0.22),
             fill=C_BAR_LIGHT)
    add_text_box(s5, str(before), bx + Inches(0.82) + bar_bg_w * before / 10 - Inches(0.25),
                 cy_ + Inches(0.32), Inches(0.22), Inches(0.22),
                 font_size=Pt(8.5), bold=True, color=C_PURPLE_DK, align=PP_ALIGN.CENTER)
    # 이후
    add_text_box(s5, "디마프", bx + Inches(0.25), cy_ + Inches(0.6), Inches(0.5), Inches(0.22),
                 font_size=Pt(8.5), bold=True, color=C_PURPLE_DK, align=PP_ALIGN.RIGHT)
    add_rect(s5, bx + Inches(0.82), cy_ + Inches(0.6), bar_bg_w, Inches(0.22),
             fill=RGBColor(0xED, 0xE8, 0xF0))
    add_rect(s5, bx + Inches(0.82), cy_ + Inches(0.6), bar_bg_w * after / 10, Inches(0.22),
             fill=C_BAR_DARK)
    add_text_box(s5, str(after), bx + Inches(0.82) + bar_bg_w * after / 10 - Inches(0.25),
                 cy_ + Inches(0.6), Inches(0.22), Inches(0.22),
                 font_size=Pt(8.5), bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# 구분선
add_rect(s5, bx + Inches(0.25), by + Inches(2.82), bw - Inches(0.5), Inches(0.01), fill=C_DIVIDER)

add_text_box(s5, "* 개인 사용 경험 기반 주관적 평가",
             bx + Inches(0.25), by + Inches(5.2), bw - Inches(0.4), Inches(0.25),
             font_size=Pt(8), color=C_MUTED, align=PP_ALIGN.CENTER)

keyword_bar(s5, ["당김 완화", "수분 지속", "피부결 개선", "편안한 피부", "보습 효과"])
slide_number(s5, 5)


# ══ SLIDE 6 : 결론 및 참고문헌 ═══════════════
s6 = prs.slides.add_slide(blank_layout)
add_rect(s6, 0, 0, W, H, fill=C_BG)
top_bar(s6)
section_label(s6, "Conclusion & References")
slide_title(s6, "결론 및 참고문헌")

# 세로 구분선
add_rect(s6, W * 0.565, Inches(1.0), Inches(0.01), Inches(5.9), fill=C_DIVIDER)

# ── 왼쪽: 결론 ──
LX = Inches(0.6)
sub_heading(s6, "결론", LX, Inches(1.25))
add_multiline_textbox(s6, [
    "건성 피부의 가장 핵심적인 관리 과제는 충분한 수분 공급과 피부 장벽 유지입니다.",
    "이번 루틴을 실천하면서 히알루론산, 판테놀, 세라마이드와 같이",
    "피부 보습에 직접적으로 관여하는 성분을 중심으로 제품을 선택하는 것이",
    "건성 피부 관리에 실질적으로 도움이 된다는 것을 경험할 수 있었습니다.",
    "",
    "또한 제형이 뻑뻑한 크림을 세럼과 혼합하여 사용하는 방식은",
    "제품의 발림성을 높이고 흡수를 도울 수 있었습니다.",
    "앞으로도 피부 상태에 맞는 성분과 사용 방법을 꾸준히 연구하고",
    "실천함으로써 피부 건강을 유지해 나가고자 합니다.",
], LX, Inches(1.57), Inches(6.7), Inches(2.5),
   font_size=Pt(10.5), color=C_BODY_TEXT, line_spacing=Pt(3))

# 키워드 태그
kw_y = Inches(4.3)
kws6 = ["보습 루틴", "피부 장벽", "성분 중심 스킨케어", "건성 피부 관리"]
kx6 = LX
for kw in kws6:
    w_k = Inches(len(kw) * 0.155 + 0.35)
    add_rect(s6, kx6, kw_y, w_k, Inches(0.3), fill=C_PURPLE_BG, line_color=C_PURPLE_LT, line_width=Pt(0.5))
    add_text_box(s6, kw, kx6 + Inches(0.1), kw_y + Inches(0.03), w_k, Inches(0.25),
                 font_size=Pt(9.5), color=C_PURPLE_DK)
    kx6 += w_k + Inches(0.12)

# ── 오른쪽: 참고문헌 ──
RX = W * 0.585
add_text_box(s6, "REFERENCES", RX, Inches(1.25), Inches(5.5), Inches(0.28),
             font_size=Pt(9), bold=True, color=C_PURPLE_MD)

refs = [
    "디마프(De:maf) 공식 홈페이지.\n히어로 마이 퍼스트 세럼 제품 정보.\n(demaf.global)",
    "네이버 지식백과. 「건성 피부」.\n피부 유형 및 특성 관련 항목.\n서울대학교병원 의학정보.",
    "대한피부과학회. 피부 건강 정보 —\n피부 건조증과 보습 관리.\n(www.derma.or.kr)",
    "피부미용학 수업 교재 참고\n(수업 중 배부된 강의자료 기반).",
]
for i, ref in enumerate(refs):
    ry = Inches(1.7) + i * Inches(1.1)
    add_rect(s6, RX, ry, Inches(0.04), Inches(0.26), fill=C_PURPLE_LT)
    add_multiline_textbox(s6, ref.split("\n"), RX + Inches(0.12), ry,
                          Inches(5.2), Inches(0.75),
                          font_size=Pt(10), color=C_REF_TEXT, line_spacing=Pt(2))

slide_number(s6, 6)


# ── 저장 ────────────────────────────────────
out = "/home/user/claude-practice/hyunjoo_skincare.pptx"
prs.save(out)
print(f"저장 완료: {out}")
