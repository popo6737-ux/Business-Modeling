from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn

prs = Presentation('/home/user/Business-Modeling/Tesla_Filled.pptx')

FONT  = '맑은 고딕'
RED   = RGBColor(0xE3, 0x19, 0x37)
DARK  = RGBColor(0x1A, 0x1A, 0x1A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY = RGBColor(0xF0, 0xF0, 0xF0)
DGRAY = RGBColor(0x55, 0x55, 0x55)
GREEN = RGBColor(0x1E, 0x8B, 0x4C)
BLUE  = RGBColor(0x16, 0x4F, 0x96)
SLATE = RGBColor(0x2C, 0x3E, 0x50)
GOLD  = RGBColor(0xFF, 0xCC, 0x00)
ORANGE= RGBColor(0xD3, 0x54, 0x00)

def add_rect(slide, l, t, w, h, bg):
    sh = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = bg
    sh.line.fill.background()
    return sh

def set_para(tf, items):
    for p in tf._txBody.findall(qn('a:p')):
        tf._txBody.remove(p)
    tf.word_wrap = True
    first = True
    for text, size, bold, color, align, sp_b in items:
        if first:
            p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
            first = False
        else:
            p = tf.add_paragraph()
        p.text = text
        p.alignment = align
        p.space_before = Pt(sp_b)
        p.space_after  = Pt(1)
        r = p.runs[0] if p.runs else p.add_run()
        r.font.name  = FONT
        r.font.size  = Pt(size)
        r.font.bold  = bold
        r.font.color.rgb = color

def remove_added(slide, keep_names):
    to_rm = [sh for sh in slide.shapes if sh.name not in keep_names]
    for sh in reversed(to_rm):
        sh._element.getparent().remove(sh._element)

# ══════════════════════════════════════════════════════
# SLIDE 4 — Q1: 제품 전략 + 마케팅 전략
# ══════════════════════════════════════════════════════
s4 = prs.slides[3]
remove_added(s4, {'직사각형 1', '직선 연결선 3', 'Title 1'})

LX = 450000
TY = 2090000
W  = 8200000
LW = 3850000   # 왼쪽 컬럼
RW = 4100000   # 오른쪽 컬럼
RX = LX + LW + 100000
GAP = 60000

# ── 왼쪽: 제품 전략 ────────────────────────────────────
h1 = add_rect(s4, LX, TY, LW, 340000, RED)
set_para(h1.text_frame, [
    ("▸  제품 전략 : 계단식 확장", 13, True, WHITE, PP_ALIGN.LEFT, 7),
])

# Tier 1
t1 = add_rect(s4, LX, TY+360000, LW, 870000, SLATE)
set_para(t1.text_frame, [
    ("①  Premium",                                     13, True,  GOLD,                   PP_ALIGN.LEFT, 10),
    ("Model S / X  ·  Cybertruck",                     11, False, RGBColor(0xDD,0xDD,0xFF), PP_ALIGN.LEFT,  4),
    ("대상  :  고소득 테크 얼리어답터",                  11, False, RGBColor(0xDD,0xDD,0xFF), PP_ALIGN.LEFT,  3),
    ("전략  :  브랜드 핵심 수호 — 마진·이미지 최우선",   11, False, RGBColor(0xDD,0xDD,0xFF), PP_ALIGN.LEFT,  3),
])

# Tier 2
t2 = add_rect(s4, LX, TY+1270000, LW, 870000, DARK)
set_para(t2.text_frame, [
    ("②  Mid-Market",                                  13, True,  WHITE,                  PP_ALIGN.LEFT, 10),
    ("Model 3  /  Y",                                  11, False, RGBColor(0xBB,0xCC,0xFF), PP_ALIGN.LEFT,  4),
    ("대상  :  중산층 ($50K~$100K)",                    11, False, RGBColor(0xBB,0xCC,0xFF), PP_ALIGN.LEFT,  3),
    ("전략  :  실용 프리미엄 — TCO 절감 메시지",         11, False, RGBColor(0xBB,0xCC,0xFF), PP_ALIGN.LEFT,  3),
])

# Tier 3
t3 = add_rect(s4, LX, TY+2180000, LW, 870000, BLUE)
set_para(t3.text_frame, [
    ("③  Mass Market  (신규)",                         13, True,  WHITE,                  PP_ALIGN.LEFT, 10),
    ("보급형  $25,000↓",                               11, False, RGBColor(0xCC,0xEE,0xFF), PP_ALIGN.LEFT,  4),
    ("대상  :  저·중소득층 / 다양한 고객군",             11, False, RGBColor(0xCC,0xEE,0xFF), PP_ALIGN.LEFT,  3),
    ("전략  :  Apple SE 방식 — 작지만 완벽한 Tesla",    11, False, RGBColor(0xCC,0xEE,0xFF), PP_ALIGN.LEFT,  3),
])

# ── 오른쪽: 마케팅 전략 ────────────────────────────────
h2 = add_rect(s4, RX, TY, RW, 340000, DARK)
set_para(h2.text_frame, [
    ("▸  마케팅 전략 : 4가지 실행 방향", 13, True, WHITE, PP_ALIGN.LEFT, 7),
])

cards = [
    (RED,    "①  세그먼트별 메시지 분리",
             ["고소득층  →  기술 · 혁신 · 지위재 강조",
              "중산층    →  TCO 절감 (5년 후 휘발유차보다 저렴)",
              "저소득층  →  IRA 보조금 $7,500 + 연료비 절약 전면 강조"]),
    (SLATE,  "②  유통 채널 다각화",
             ["직영 온라인 + 직영 쇼룸 유지",
              "농촌 · 교외 서비스센터 단계적 확대",
              "충전 인프라 선(先) 투자 후 차량 판매 전략"]),
    (DGRAY,  "③  금융 접근성 강화",
             ["리스(Lease) 프로그램 확대 — 초기 비용 부담 완화",
              "인증 중고(CPO) 시장 활성화",
              "IRA 세액공제 $7,500 구매 시점 즉시 적용"]),
    (GREEN,  "④  커뮤니티 파트너십",
             ["비영리단체 · CDFI와 저소득 커뮤니티 접근",
              "지역 체험 · 시승 프로그램으로 신뢰 구축",
              "다양성 반영 마케팅 — 현재 타겟과 다른 고객 공감"]),
]

card_h = 730000
for i, (bg, title, lines) in enumerate(cards):
    cy = TY + 360000 + i * (card_h + GAP)
    c = add_rect(s4, RX, cy, RW, card_h, bg)
    items = [(title, 12, True, WHITE, PP_ALIGN.LEFT, 8)]
    for line in lines:
        items.append(("• " + line, 10, False, RGBColor(0xEE,0xEE,0xEE), PP_ALIGN.LEFT, 3))
    set_para(c.text_frame, items)

# ── 핵심 원칙 바 ────────────────────────────────────────
by = TY + 3090000 + GAP
b = add_rect(s4, LX, by, W, 320000, RGBColor(0xF5,0xF0,0xE8))
b.line.color.rgb = RED
b.line.width = Pt(1)
set_para(b.text_frame, [(
    '핵심 원칙  ·  "누구나 갖고 싶은 차 중에 살 수 있는 것"  —  브랜드 희석 없는 계단식 시장 확장',
    11, True, RED, PP_ALIGN.CENTER, 8
)])


# ══════════════════════════════════════════════════════
# SLIDE 5 — Q2: 5가지 추가 역량
# ══════════════════════════════════════════════════════
s5 = prs.slides[4]
remove_added(s5, {'직사각형 1', '직선 연결선 3', 'Title 1'})

LX5 = 450000
TY5 = 1870000
W5  = 8200000
GAP5= 70000

caps = [
    (RED,    "①  저비용 제조 역량",
             "목표: $25K EV 수익성 확보",
             ["4680 배터리 셀 양산 원가 절감 가속화",
              "차량 플랫폼 단순화 — 부품 수 최소화",
              "기가팩토리 생산 수율(Yield) 극대화",
              "BYD처럼 배터리 수직계열화 심화"]),
    (BLUE,   "②  충전 인프라 확장",
             "Charging Desert 해소가 선결 과제",
             ["농촌 · 교외 Supercharger 네트워크 확장",
              "NACS 개방 → 타사 충전 수익화",
              "V2G(차량→전력망) 기술 상용화",
              "무선충전 R&D (WiTricity 등 파트너십)"]),
    (SLATE,  "③  금융 · 접근성 서비스",
             "살 수 있는 조건을 만드는 역량",
             ["리스(Lease) 프로그램 대폭 확대",
              "인증 중고(CPO) — 저가 진입 경로 제공",
              "저신용층 파이낸싱 (CDFI 협력)",
              "IRA $7,500 세액공제 즉시 현장 적용"]),
    (DGRAY,  "④  세입자 충전 솔루션",
             "미국 인구 36% — 미개척 최대 시장",
             ["공동주택 · 아파트 충전기 설치 지원 프로그램",
              "부동산 개발사 · 아파트 운영사 파트너십",
              "신축 건물 EV 충전 기본 설치 의무화 로비",
              "흑인 55% · 라틴계 52% 세입자 → 형평성 직결"]),
    (GREEN,  "⑤  커뮤니티 신뢰 구축",
             "기술 이전에 '신뢰'가 먼저",
             ["저소득 커뮤니티 EV 교육 · 시승 프로그램",
              "BlueHub Capital 등 비영리단체 협력",
              "다양성 반영 마케팅 콘텐츠 제작",
              "도심 저소득 지역 서비스센터 확대"]),
]

# 2열 레이아웃: 상단 3개 / 하단 2개
CW_3 = (W5 - 2*GAP5) // 3
CW_2 = (W5 - GAP5) // 2
CH_TOP = 2100000
CH_BOT = 2100000

for i, (bg, title, subtitle, lines) in enumerate(caps[:3]):
    cx = LX5 + i * (CW_3 + GAP5)
    c = add_rect(s5, cx, TY5, CW_3, CH_TOP, bg)
    items = [
        (title,    12, True,  WHITE,                   PP_ALIGN.LEFT, 10),
        (subtitle, 10, False, RGBColor(0xFF,0xEE,0xAA), PP_ALIGN.LEFT,  3),
    ]
    for line in lines:
        items.append(("• " + line, 10, False, RGBColor(0xEE,0xEE,0xEE), PP_ALIGN.LEFT, 4))
    set_para(c.text_frame, items)

for i, (bg, title, subtitle, lines) in enumerate(caps[3:]):
    cx = LX5 + i * (CW_2 + GAP5)
    c = add_rect(s5, cx, TY5 + CH_TOP + GAP5, CW_2, CH_BOT, bg)
    items = [
        (title,    12, True,  WHITE,                   PP_ALIGN.LEFT, 10),
        (subtitle, 10, False, RGBColor(0xFF,0xEE,0xAA), PP_ALIGN.LEFT,  3),
    ]
    for line in lines:
        items.append(("• " + line, 10, False, RGBColor(0xEE,0xEE,0xEE), PP_ALIGN.LEFT, 4))
    set_para(c.text_frame, items)

# ── 우선순위 바 ─────────────────────────────────────────
bot_y = TY5 + CH_TOP + GAP5 + CH_BOT + GAP5
sub = add_rect(s5, LX5, bot_y, W5, 300000, DARK)
set_para(sub.text_frame, [(
    "우선순위  ①저비용제조  ≥  ②충전인프라  >  ③금융접근성  ≥  ④세입자솔루션  >  ⑤커뮤니티신뢰  —  생태계 전체를 함께 구축",
    10, True, GOLD, PP_ALIGN.CENTER, 8
)])

prs.save('/home/user/Business-Modeling/Tesla_Filled.pptx')
print("Done.")
