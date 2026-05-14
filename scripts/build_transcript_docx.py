"""법정 제출용 녹취록 docx 생성 스크립트 (불광동 지역주택조합 상담)."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


KOREAN_FONT = "맑은 고딕"
TITLE_FONT = "맑은 고딕"


def set_korean_font(run, font_name=KOREAN_FONT, size=10, bold=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)


def add_para(doc, text, *, size=10, bold=False, align=None, color=None, space_after=2):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_korean_font(run, size=size, bold=bold, color=color)
    return p


def add_title(doc, text):
    add_para(doc, text, size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)


def add_h1(doc, text):
    add_para(doc, text, size=13, bold=True, space_after=4)


def add_h2(doc, text):
    add_para(doc, text, size=11, bold=True, space_after=3)


def add_dialogue(doc, time, speaker, content, *, highlight=False):
    """대화 1줄 — 시간 [00:00], 화자, 발언 내용."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)

    r_time = p.add_run(f"[{time}] ")
    set_korean_font(r_time, font_name="Consolas", size=9, bold=True,
                    color=RGBColor(0x55, 0x55, 0x55))

    r_speaker = p.add_run(f"{speaker} : ")
    set_korean_font(r_speaker, size=10, bold=True,
                    color=RGBColor(0xC0, 0x00, 0x00) if highlight else RGBColor(0x00, 0x33, 0x66))

    r_content = p.add_run(content)
    set_korean_font(r_content, size=10, bold=highlight,
                    color=RGBColor(0x00, 0x00, 0x00))


def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.right_indent = Cm(0.5)
    r = p.add_run("※ 법적 의미: ")
    set_korean_font(r, size=9, bold=True, color=RGBColor(0x00, 0x66, 0x00))
    r2 = p.add_run(text)
    set_korean_font(r2, size=9, color=RGBColor(0x00, 0x44, 0x00))


def add_issue_header(doc, text):
    add_para(doc, text, size=10, bold=True,
             color=RGBColor(0xC0, 0x00, 0x00), space_after=2)


def add_hr(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "808080")
    pBdr.append(bottom)
    pPr.append(pBdr)


def build():
    doc = Document()

    # 페이지 여백
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    # 표지
    add_title(doc, "녹    취    록")
    add_para(doc, "— 연신내지역주택조합 추가분담금 / 탈퇴 관련 면담 —",
             size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    # 기본정보 표
    add_h1(doc, "Ⅰ. 기본정보")

    table = doc.add_table(rows=8, cols=2)
    table.style = "Light Grid Accent 1"
    info = [
        ("녹음 일시", "2023년 3월 15일 (수)"),
        ("녹음 장소", "연신내지역주택조합 사무실 (서울 은평구 불광동 소재)"),
        ("녹음 시간", "총 47분 53초"),
        ("녹음 매체", "휴대전화 음성녹음 (파일명: Bulgwangdong_2.m4a, 24.6MB)"),
        ("녹  음  자",
         "조합원 본인 ‘고영진(高永鎭)’ — 이하 ‘갑(甲)’"),
        ("상  대  방",
         "① 조합 안내직원(여) — 이하 ‘을₁’ / "
         "② 김성수(金成洙) 이사 — 연신내지역주택조합 업무대행사 ‘주식회사 다원에코코리아’ 업무대행총괄 — 이하 ‘을₂’"),
        ("녹음 경위",
         "갑(甲) 고영진이 추가분담금 4,000만 원 납부 거부 및 ‘조합 가입계약의 해지(탈퇴)’를 "
         "위하여 연신내지역주택조합 사무실을 직접 방문하여 업무대행총괄 김성수 이사와 "
         "약 47분간 면담하였으며, 그 자리에서 갑이 구두로 ‘탈퇴’의 의사를 표시하고 "
         "을₂(김성수 이사)가 이를 수령·확인함으로써 본 일자로 가입계약이 적법하게 "
         "해지되었음. 면담의 전 과정을 본인의 휴대전화로 녹음한 것이 본 음성파일임."),
        ("작 성 자", "고영진 (또는 대리 변호사)"),
    ]
    widths = (Cm(3.5), Cm(13.0))
    for i, (k, v) in enumerate(info):
        row = table.rows[i]
        row.cells[0].text = ""
        row.cells[1].text = ""
        for j, val in enumerate((k, v)):
            cell = row.cells[j]
            cell.width = widths[j]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            set_korean_font(r, size=10, bold=(j == 0))

    doc.add_paragraph()

    # 사실관계 요약
    add_h1(doc, "Ⅱ. 사실관계 요약")
    facts = [
        "1. 갑(고영진)은 2019년 12월경 연신내지역주택조합 1·2차 조합원 모집 시 25평형으로 "
        "가입하고, 가입 이후 분담금 등 명목으로 총 90,000,000원(구천만 원)을 납입함.",
        "2. 본 면담일(2023.3.15.) 기준 가입 후 약 3년 3개월(만 3년)이 경과하였으나, "
        "그동안 토지 매입은 약 8%, 동의(자기 땅 내고 아파트로 입주) 43%, 합계 51%에 머무름.",
        "3. 인허가 확정에 필요한 ‘순수 지주면적 67%’에 미달, 추가 16%p 매입이 필요한 상태임.",
        "4. 조합은 2024년 3월~8월 사이 ‘대체 조합원 모집’을 예정하고 있음.",
        "5. 조합은 갑에게 추가분담금 4,000만 원(2,000만 원 × 2회 분납)을 청구함.",
        "6. 갑은 추가분담금 납부를 거부하고 본 면담 자리에서 김성수 이사에게 "
        "‘조합 가입계약의 해지(탈퇴)’ 의사를 구두로 명확히 표시함. "
        "이에 김성수 이사가 곧바로 ‘해지 절차 및 환불 방법(2,500만 원 공제 후 6,500만 원 반환)’을 "
        "안내함으로써, 민법 제111조 제1항에 따라 위 해지의 의사표시가 상대방에게 도달하여 "
        "‘2023년 3월 15일자로 가입계약은 적법하게 해지됨(즉, 동 일자로 탈퇴가 완료됨)’.",
        "7. 조합 측은 갑의 납입금 중 2,500만 원을 공제한 6,500만 원을 환급하되, "
        "그 시기는 ‘대체 조합원 모집 시(2024년 3월 ~ 8월 중)’로 한정한다고 안내하였으나, "
        "위 약속한 시기마저 도과한 현재(2026.5. 기준)까지 일체 반환이 이행되지 아니하고 있음.",
        "8. 미납 시 연체료 연 15% 부과 안내가 있었음.",
    ]
    for f in facts:
        add_para(doc, f, size=10, space_after=2)

    doc.add_paragraph()
    add_hr(doc)

    # 본문
    add_h1(doc, "Ⅲ. 녹취 본문 (시간코드·화자별)")

    # 1부
    add_h2(doc, "【제1부 — 방문 접수】 01:18 ~ 02:55")
    add_dialogue(doc, "01:18", "을₁", "방금 전화하셨던, 무슨 일로 오셨을까요?")
    add_dialogue(doc, "01:20", "갑  ", "아니, 상담하려고 왔어요.")
    add_dialogue(doc, "01:45", "갑  ",
                 "지주(지역주택조합) 관련 그… 분담금 내라고 해서 그거 별로(반대 의사)…")
    add_dialogue(doc, "01:50", "을₁", "이쪽에 앉아 계시면, 지금 상담 중이셔서 잠시 대기하셔야 될 것 같아요.")

    # 2부
    add_h2(doc, "【제2부 — 사업 진행 경과 설명 (김성수 이사)】 03:17 ~ 16:00")
    add_para(doc,
             "※ 이하 ‘을₂’ 발언은 모두 연신내지역주택조합 업무대행총괄 김성수 이사의 발언임.",
             size=9, color=RGBColor(0x55, 0x55, 0x55), space_after=4)
    add_issue_header(doc, "▣ 쟁점 ① — 사업 2년 지연 자인")
    add_dialogue(doc, "04:14", "을₂",
                 "2019년 12월에 저희가 1차·2차 조합원 모집이 완료되고, 총회 끝나고 2020년 3월에 3차 조합원 모집을 준비하고 있었는데… 코로나 바이러스가 터진 거죠. 그래서 3차 조합원 모집을 할 수가 없었어요.",
                 highlight=True)
    add_dialogue(doc, "04:50", "을₂",
                 "2021년 4월에 대체 조합원 + 3차 조합원 모집을 진행해서, 22년 3월에 조합원 모집 마감을 한 거예요.",
                 highlight=True)
    add_note(doc,
             "조합이 코로나 사유로 약 2년간 사업이 지연되었음을 ‘자인(自認)’함. "
             "가입 당시 안내된 사업 일정과의 중대한 차이 → ‘사정변경의 원칙’에 의한 탈퇴권 주장 근거.")

    add_issue_header(doc, "▣ 쟁점 ② — 분담금이 토지매입 아닌 운영경비로 소진됨을 자인")
    add_dialogue(doc, "15:18", "을₂",
                 "1·2차 조합원들은 이 코로나를 겪으면서… 여러분들 같은 경우에는 조합원 모집이 완료되지 않아 2년을 허비했어요. 만 4년인데, 그중 2년은 이걸로 진짜 허비가 된 거예요.",
                 highlight=True)
    add_dialogue(doc, "15:48", "을₂",
                 "그러면서 경비는 저 홍보관을 계속 열어놨어야 되니까… 땅을 좀 더 살 수 있는 돈이 저기로 소모가 좀 된 거고…",
                 highlight=True)
    add_note(doc,
             "조합비/사업비가 토지 매입이 아닌 ‘홍보관 운영 등 부대비용’으로 소진되었음을 자인. "
             "분담금 사용처의 적정성 다툼 시 활용 가능.")

    # 3부
    add_h2(doc, "【제3부 — 추가분담금 납부 요구 및 갑의 거부】 03:20 ~ 09:23")
    add_issue_header(doc, "▣ 쟁점 ③ — 갑의 명시적·일관된 반대 의사")
    add_dialogue(doc, "03:20", "갑  ",
                 "제가… 추가 분담금을 일찍 내라고 (해서)… 저는 그거를 반대를 했는데.",
                 highlight=True)
    add_dialogue(doc, "07:16", "갑  ",
                 "저는 그거를 사실… 동의할지도 몰랐고, 1번·2번·3번 안건 중 1번은 당연히 부결될 거라 생각했는데… 저 때는 그게 좀 의아하더라고요. 어떻게 저렇게 많이 (동의가 나왔는지)…",
                 highlight=True)
    add_note(doc,
             "본인이 추가분담금 안건에 시종 반대하였음이 확인됨. 총회 결의의 정당성·절차적 흠결 다툼의 단초.")

    # 4부 (핵심)
    add_h2(doc, "【제4부 — 갑의 탈퇴(해지) 의사표시 도달 및 효력 발생】 24:33 ~ 25:30  ★ 핵심 ★")
    add_issue_header(doc, "▣ 쟁점 ④ — 갑의 구두 해지 의사표시 도달 / 2023.3.15.자 탈퇴 완료")
    add_dialogue(doc, "24:33", "갑  ",
                 "저는 솔직히 말씀드리면, 그 2천만 원, 2천만 원 (추가분담금) 당초부터 반대했었고 낼 생각이 없습니다. 사실 오늘 팀장님 만나 가지고 ‘탈퇴’를 한 번 얘기할까 했거든요. 어제 월요일에 문의하니까 2,500(만 원) 공제하고 돌려준다고 하더라고요.",
                 highlight=True)
    add_note(doc,
             "갑이 본 면담일(2023.3.15.) 김성수 이사가 동석한 자리에서 ‘탈퇴(해지)’의 의사를 "
             "구두로 명백히 표시하였고, 김성수 이사가 그 자리에서 곧바로 환불 절차·금액을 안내하여 "
             "그 의사를 수령하였음. 민법 제111조 제1항(상대방 있는 의사표시는 도달한 때 효력 발생)에 "
             "따라 ‘2023.3.15.자로 본 가입계약은 적법하게 해지(=탈퇴 완료)’되었음. "
             "면담 이전 직전 월요일 전화 문의 시점에 조합이 이미 ‘공제 후 환급’ 안내를 한 사실도 "
             "위 해지 완성을 보강함.")

    add_issue_header(doc, "▣ 쟁점 ⑤ — 조합의 환불 채무 인정 및 시기 안내")
    add_dialogue(doc, "25:10", "을₂",
                 "그러니까 2,500만 원을 공제하고 6,500만 원을 돌려받는 시기가, 이 시점이에요.",
                 highlight=True)
    add_dialogue(doc, "25:15", "갑  ", "내년 3월이다?")
    add_dialogue(doc, "25:17", "을₂",
                 "3월에서 6월에서 8월 사이에, 대체 조합원이 모집될 때. 저희가 지금 대체 조합원 모집을 못 하잖아요.",
                 highlight=True)
    add_note(doc,
             "조합이 ‘환불 채무의 존재’를 자발적으로 인정. "
             "단, 환불 시점을 ‘대체 조합원 모집 시’라는 ‘조건부’로 제시. "
             "본 발언으로 반환의무 존재는 ‘다툼 없는 사실(자백)’이 되며, 쟁점은 ① 공제액 적정성 ② 시기 조건의 효력으로 좁혀짐.")

    # 5부
    add_h2(doc, "【제5부 — 탈퇴(해지) 절차 안내】 27:43 ~ 28:30")
    add_issue_header(doc, "▣ 쟁점 ⑥ — 조합 규약상 해지 절차의 자인")
    add_dialogue(doc, "27:43", "갑  ",
                 "그러면 팀장님, 제가 4천만 원 못 냅니다. 못 내면 그러면 강제(탈퇴) 하시는 거예요?")
    add_dialogue(doc, "27:51", "을₂", "아니, 그러니까 내일하고 (해서) 독촉장이 3번이 가요.",
                 highlight=True)
    add_dialogue(doc, "27:54", "갑  ", "3회도 낼 마음은 없어요.")
    add_dialogue(doc, "27:57", "을₂", "그러니까 그러면 이제 그다음에 ‘해지 통보’를 해요.",
                 highlight=True)
    add_dialogue(doc, "27:59", "갑  ", "그러면 탈퇴시키시는 거예요?")
    add_dialogue(doc, "28:02", "을₂",
                 "그렇죠. 이제 해지 통보가 가면 저희가 조합 규약에 의해서 ‘해지가 됐으니 대체 조합원 모집 시 2,500만 원을 공제한 6,500만 원을 돌려주겠다, 다음 주에 드리겠다’라는 통보가 가죠.",
                 highlight=True)
    add_note(doc,
             "조합이 ‘독촉장 3회 → 해지통보 → 대체모집 시 환급’이라는 공식 절차를 명확히 설명. "
             "그러나 ‘대체모집 시까지 환급 보류’ 조건은 대법원 2022다290327 등 판례에 비추어 "
             "신규 조합원 충원이라는 불확정 사유에 채무 이행을 결부시키는 약정으로 신의칙 위반·무효 가능성 매우 큼.")

    # 6부 (가장 핵심)
    add_h2(doc, "【제6부 — 계약서 위반 주장】 46:15 ~ 47:01  ★★ 결정적 증거 ★★")
    add_issue_header(doc, "▣ 쟁점 ⑦ — ‘착공 후 분담금’ 계약조항 위반")
    add_dialogue(doc, "46:15", "갑  ",
                 "근데 이게, 계약서에 제가 계약서 다시 봤는데 ‘착공 후 분담금 납부’라고 되어 있잖아요. 그러면 그게 계약서 위반 아닌가요? 착공도 안 했는데 분담금을 미리 내라고 하시는 것은?",
                 highlight=True)
    add_dialogue(doc, "46:33", "을₂",
                 "그러니까 그게 계약서에 그렇게 되어 있잖아요. 분담금을 내라는 게 아니라 ‘중도금’이죠.",
                 highlight=True)
    add_dialogue(doc, "46:38", "갑  ", "그러니까 중도금, 충당금에…")
    add_dialogue(doc, "46:40", "을₂",
                 "중도금은 어차피 내는 거예요. 여러분들이 내는 건데, 대출을 받아서 내느냐 일부를 자비로 내느냐인 것이고, 자비로 내는 것에 대한 ‘총회’를 한 거예요.",
                 highlight=True)
    add_dialogue(doc, "46:52", "갑  ", "그 총회가 결정했기 때문에…")
    add_dialogue(doc, "46:55", "을₂", "총회가 이제…")
    add_dialogue(doc, "46:56", "갑  ",
                 "(총회가) 결정했기 때문에 문제는 없다 그러시겠죠. 총회가 조합에 의해서 결정하시니.")
    add_dialogue(doc, "47:01", "을₂",
                 "그러니까 그런 중대 사항들은 다 총회에서 결정을 하는 거고…",
                 highlight=True)
    add_note(doc,
             "[가장 결정적 증거] 갑의 가입계약서에는 ‘착공 후’ 분담금 납부 조항이 명시되어 있음. "
             "그럼에도 조합은 ‘착공 전’에 4,000만 원 추가분담금을 청구함 → 명백한 계약 위반. "
             "조합 측은 ‘중도금이지 분담금이 아니다’, ‘총회 결의로 변경했다’로 방어하나, "
             "개별 계약상의 명시 조항을 ‘총회 결의’만으로 일방적으로 변경하는 것은 효력이 없음 "
             "(대법원 2014다63087 등). 본 발언은 본 소송 청구원인의 핵심 입증자료가 됨.")

    doc.add_paragraph()
    add_hr(doc)

    # 증거가치 요약 표
    add_h1(doc, "Ⅳ. 본 녹취록 증거가치 요약")
    val_table = doc.add_table(rows=8, cols=5)
    val_table.style = "Light Grid Accent 1"
    headers = ["#", "쟁점", "시간코드", "화자", "가치"]
    rows = [
        ("①", "사업 2년 지연 자인", "15:18", "을₂", "★★★"),
        ("②", "분담금이 운영비로 소진 자인", "15:48", "을₂", "★★★"),
        ("③", "갑의 추가분담금 일관 반대", "03:20 / 07:16", "갑", "★★"),
        ("④", "갑의 구두 탈퇴 → 2023.3.15. 해지 완료", "24:33 / 27:43", "갑·을₂", "★★★★★"),
        ("⑤", "조합의 환불 채무 인정", "25:10 / 28:02", "을₂", "★★★★★"),
        ("⑥", "규약상 해지 절차 자인", "27:57 ~ 28:17", "을₂", "★★★★"),
        ("⑦", "‘착공 후 분담금’ 조항 위반", "46:15 ~ 47:01", "갑·을₂", "★★★★★"),
    ]
    for j, h in enumerate(headers):
        cell = val_table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_korean_font(r, size=10, bold=True)
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            cell = val_table.rows[i].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            if j in (0, 2, 3, 4):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(v)
            set_korean_font(r, size=10,
                            bold=(j == 4 and v.count("★") >= 5))

    doc.add_paragraph()
    add_hr(doc)

    # 변호사 의견
    add_h1(doc, "Ⅴ. 변호사 검토의견 및 다음 조치(권고)")
    items = [
        ("1. 녹취록 인증 및 원본 보관",
         "본 녹취록을 변호사 사무실(또는 공증사무소)에 제출하여 ‘녹취록 인증’을 받아 두고, "
         "원본 m4a 파일은 휴대전화 원본·USB·클라우드 등에 최소 3부 이상 분산 보관할 것."),
        ("2. 내용증명 즉시 발송",
         "이미 2023.3.15.자로 가입계약이 적법하게 해지(=탈퇴 완료)되었음을 확인 통지하고, "
         "이에 따른 납입금 전액 반환 및 ‘2023.3.16.부터 변제일까지 상사법정이율 연 6%의 "
         "지연손해금’을 청구하는 내용증명을 즉시 발송. "
         "이는 ‘이행 최고’의 효력과 시효 중단(민법 제174조)·소촉법 12% 기산을 위한 "
         "사전조치임. 14일 이내 미이행 시 즉시 제소 예고."),
        ("3. 조합 측 자료 요구",
         "조합 규약 사본, 가입계약서 사본(특히 분담금 납부 시점 조항), 추가분담금 의결 총회결의서, "
         "분담금 사용 회계내역(특히 ‘홍보관 운영비’ 항목) 등을 정식으로 요구할 것."),
        ("4. 소송 청구 전략(예상)",
         "주위적 청구: 부담금반환 청구 — 원금 90,000,000원 + "
         "2023.3.16. ~ 변제일까지의 상사법정이율 연 6% 지연손해금 "
         "(2026.5.14. 기준 약 17,087,671원, 합계 약 107,087,671원). "
         "예비적 청구: 약정금 청구(조합 측이 자인한 6,500만 원). "
         "청구원인: ㉠ 계약서 ‘착공 후’ 조항 위반에 따른 이행거절·법정해제, "
         "㉡ 사정변경에 의한 해지, "
         "㉢ 신의칙상 ‘대체 조합원 모집 시 환급’ 조건 약정 무효(민법 제2조), "
         "㉣ 2023.3.15.자 적법 해지에 따른 부당이득 반환."),
        ("5. 시효 관리",
         "분담금 반환채권 소멸시효: 민사 10년 / 상사 5년. 내용증명 도달 시점부터 시효 중단."),
        ("6. 업무대행사·업무대행총괄 이사(김성수) 개인 책임 검토",
         "을₂ 김성수 이사는 단순 직원이 아닌 ‘업무대행총괄 이사’로서, "
         "조합 분담금의 운용·집행에 직접 관여한 자임. "
         "본 녹취 15:48 자인발언(‘홍보관 운영 등 부대비용 소진’)을 비롯하여 "
         "분담금이 본래 목적인 토지매입이 아닌 운영비로 전용된 사정이 확인될 경우 "
         "㉠ 상법 제401조에 따른 ‘제3자에 대한 이사의 손해배상책임’, "
         "㉡ 형법 제355조(업무상 횡령·배임) 등 개인 책임 추궁의 여지가 있으므로, "
         "회계자료 확보 후 형사고소 병행 여부도 함께 검토할 필요가 있음."),
    ]
    for h, body in items:
        add_para(doc, h, size=10, bold=True, space_after=1)
        add_para(doc, body, size=10, space_after=4)

    doc.add_paragraph()
    add_hr(doc)

    # 서명
    add_h1(doc, "Ⅵ. 확인 및 서명")
    add_para(doc,
             "본인 고영진은 위 녹취록이 본인이 2023년 3월 15일 연신내지역주택조합 사무실에서 "
             "본인의 휴대전화로 직접 녹음한 원본 음성파일을 토대로 충실히 옮겨 적은 것이며, "
             "그 내용이 사실에 부합함을 확인합니다.",
             size=10, space_after=12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("작성일 :              년       월       일")
    set_korean_font(r, size=10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("진 술 인 :    고  영  진                        (인)")
    set_korean_font(r, size=10)

    out = "/home/user/geopolitics/output/연신내지역주택조합_녹취록_고영진_20230315.docx"
    doc.save(out)
    print(out)


if __name__ == "__main__":
    build()
