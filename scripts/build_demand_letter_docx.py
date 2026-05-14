"""연신내지역주택조합 분담금 반환 청구 내용증명 docx 생성."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


FONT = "맑은 고딕"


def kfont(run, *, size=11, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), FONT)
    rFonts.set(qn("w:ascii"), FONT)
    rFonts.set(qn("w:hAnsi"), FONT)


def para(doc, text, *, size=11, bold=False, align=None,
         color=None, space_after=4, left_indent=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if left_indent is not None:
        p.paragraph_format.left_indent = Cm(left_indent)
    p.paragraph_format.line_spacing = 1.4
    r = p.add_run(text)
    kfont(r, size=size, bold=bold, color=color)
    return p


def hr(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def build():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(2.5)
        s.right_margin = Cm(2.5)

    # 표제
    para(doc, "통    고    서", size=22, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    para(doc, "(분담금 반환 청구 및 조합 가입계약 해지 통지)",
         size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         space_after=14)

    hr(doc)

    # 수신인
    para(doc, "[수    신    인]", size=12, bold=True, space_after=4)
    para(doc,
         "1. 연신내지역주택조합  조합장 ○○○ 귀하\n"
         "    (주소 :                                                              )",
         size=11, left_indent=0.5, space_after=4)
    para(doc,
         "2. 위 조합의 업무대행사   ○○○○ 주식회사  대표이사 ○○○ 귀하\n"
         "    (주소 :                                                              )",
         size=11, left_indent=0.5, space_after=4)
    para(doc,
         "3. 업무대행총괄  김성수(金成洙) 이사 귀하\n"
         "    (주소 :                                                              )",
         size=11, left_indent=0.5, space_after=10)

    # 발신인
    para(doc, "[발    신    인]", size=12, bold=True, space_after=4)
    para(doc,
         "성      명 : 고  영  진 (高永鎭)\n"
         "주      소 :                                                                  \n"
         "연  락  처 :                                                                  \n"
         "조합원번호 :                            ( 25평형 / 2019년 12월 가입 )",
         size=11, left_indent=0.5, space_after=14)

    hr(doc)
    doc.add_paragraph()

    # 본문 제목
    para(doc, "제    목 :  이미 해지된 조합 가입계약에 따른 분담금 전액 반환 및 지연손해금 청구의 건",
         size=12, bold=True, space_after=14)

    # 인사말
    para(doc,
         "1. 귀 조합 및 업무대행사의 무궁한 발전을 기원합니다.",
         size=11, space_after=10)

    # 항목 1. 사실관계
    para(doc, "2. 사실관계", size=12, bold=True, space_after=6)
    para(doc,
         "가. 본인 고영진(이하 ‘통고인’)은 2019년 12월경 귀 조합의 1·2차 조합원 모집 당시 "
         "25평형 1세대로 정식 가입하고, 그 무렵부터 가입계약서에 따라 분담금(이른바 ‘계약금·"
         "업무추진비’)을 적법하게 납부하여 왔습니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "나. 그러나 가입 후 만 3년이 경과한 2023년 3월 15일 현재까지도 본 사업은 "
         "토지 매입 약 8%, 동의 43%(합계 51%)에 머물러 인허가에 필요한 ‘순수 지주면적 67%’에 "
         "현저히 미달하는 상태이며, 당초 안내된 사업 일정 대비 약 2년 이상의 지연이 발생하였습니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "다. 그럼에도 귀 조합은 통고인이 명시적으로 반대 의사를 표명한 추가분담금 "
         "4,000만 원(2,000만 원 × 2회 분납)을 본인에게 청구하고 있으며, "
         "이는 통고인이 체결한 가입계약서 분담금 납부 시기 조항(‘착공 후 분담금 납부’)에 "
         "정면으로 반하는 것입니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "라. 이에 통고인은 2023년 3월 15일 귀 조합 사무실을 직접 방문하여 "
         "업무대행총괄 김성수 이사와 약 47분간 면담하였고, 그 자리에서 "
         "① 통고인이 본인의 가입계약을 ‘해지(탈퇴)’한다는 의사를 명백하고 종국적으로 "
         "구두 표시함과 동시에, ② 이를 수령한 김성수 이사가 그 자리에서 "
         "‘조합 규약에 따라 해지가 됐으니, 납입금 중 2,500만 원을 공제한 6,500만 원을 "
         "대체 조합원 모집 시(2024년 3월~8월 중) 반환한다’고 답변함으로써, "
         "민법 제111조 제1항에 따라 통고인의 해지 의사표시가 상대방(귀 조합 및 그 "
         "업무대행총괄 이사)에게 도달하여 ‘2023년 3월 15일자로 본 가입계약은 적법하게 해지’"
         "되었으며, 같은 날자로 귀 조합의 통고인에 대한 분담금 반환채무가 발생하였습니다. "
         "(위 면담 전 과정은 통고인이 휴대전화로 녹음하여 원본을 보관 중이며, "
         "이를 정리한 녹취록 일체를 본 통고서에 첨부합니다.)",
         size=11, left_indent=0.7, space_after=6)

    para(doc,
         "마. 그러나 귀 조합은 위 해지일(2023.3.15.)로부터 본 통고서 발송일 현재까지 "
         "약 3년 2개월이 경과하였음에도, 위 6,500만 원조차 일체 반환하지 아니하고 있으며, "
         "조합이 임의로 약속하였던 ‘2024년 3월~8월 대체 조합원 모집 시’라는 시기도 "
         "이미 도과한 상태로 ‘이행지체’가 명백합니다.",
         size=11, left_indent=0.7, space_after=12)

    # 항목 2. 법적 주장
    para(doc, "3. 통고인의 법적 주장", size=12, bold=True, space_after=6)

    para(doc,
         "가. (가입계약 위반) 통고인의 가입계약서에는 분담금을 ‘착공 후’ 납부하도록 "
         "명시되어 있음에도, 귀 조합은 착공조차 이루어지지 아니한 상태에서 ‘중도금’이라는 "
         "명목을 빌어 사실상의 분담금을 강제 청구하고 있습니다. 이는 개별 조합원과의 "
         "가입계약상 명시된 ‘납부 시점에 관한 본질적 조항’을 총회 결의만으로 일방 변경한 것으로, "
         "대법원 2014다63087 판결 등에 비추어 통고인에 대하여는 그 효력을 미칠 수 없습니다.",
         size=11, left_indent=0.7, space_after=6)

    para(doc,
         "나. (사정변경에 의한 해지) 가입 후 만 3년이 경과하였음에도 토지 매입률이 8%에 "
         "불과하고, 인허가 단계에 이르지 못한 점, 그동안 분담금이 본래 목적인 토지 매입이 "
         "아닌 홍보관 운영 등 부대비용으로 상당 부분 소진된 점은 당초 가입계약 체결 당시 "
         "예상할 수 없었던 ‘중대한 사정변경’에 해당합니다(대법원 2007다1326 등).",
         size=11, left_indent=0.7, space_after=6)

    para(doc,
         "다. (탈퇴의 자유 — 임의탈퇴 금지 약정의 무효) 주택법상 지역주택조합 조합원의 "
         "임의 탈퇴를 일률적으로 제한하거나, 신규 조합원 충원이라는 우연한 사정에 환불 시기를 "
         "결부시키는 약정은 신의성실의 원칙(민법 제2조)에 반하여 무효입니다 "
         "(대법원 2022다290327, 서울고등법원 2021나2017439 등).",
         size=11, left_indent=0.7, space_after=6)

    para(doc,
         "라. (해지의 완성) 위 가입계약은 통고인이 2023년 3월 15일 김성수 이사 면담 자리에서 "
         "구두로 해지의 의사표시를 하고 김성수 이사가 이를 수령하여 곧바로 환불 절차를 안내한 "
         "시점에 ‘이미 적법하게 해지’되었습니다(민법 제111조 제1항). 본 통고서는 위 사실의 "
         "확인 및 그에 따른 환불채무의 이행을 ‘다시 한번 최고(催告)’하는 의미를 가집니다 "
         "(민법 제387조 제2항, 제174조).",
         size=11, left_indent=0.7, space_after=12)

    # 항목 3. 청구
    para(doc, "4. 구체적 청구", size=12, bold=True, space_after=6)

    para(doc,
         "가. 통고인이 그동안 귀 조합에 납입한 분담금 등 일체의 금원 "
         "(원금 ○○,○○○,○○○원 — ※ 영수증·이체내역 별첨)을 "
         "본 통고서 도달일로부터 14일 이내에 아래 계좌로 전액 반환하실 것을 청구합니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "         ▷ 반 환 계 좌 :                          은행      "
         "예금주 : 고영진      계좌번호 :                                  ",
         size=11, left_indent=1.0, space_after=10)

    para(doc,
         "나. 위 기한 내에 전액 반환이 어려운 경우, 최소한 귀 조합이 2023년 3월 15일 "
         "면담 시 김성수 이사를 통하여 자인(自認)한 금 6,500만 원 "
         "(공제액 2,500만 원에 대하여는 그 산정근거의 정당성·약관 효력을 별도로 다투며 "
         "추가 청구를 유보함)을 동 기한 내에 우선 반환할 것을 청구합니다.",
         size=11, left_indent=0.7, space_after=10)

    para(doc,
         "다. 위 분담금 반환채무에 대하여, 해지일인 2023년 3월 15일 다음 날부터 "
         "변제일까지 상사법정이율 연 6%의 비율에 의한 ‘지연손해금’을 가산하여 청구하며, "
         "소제기 후 ‘소장 부본 송달일 다음 날’부터는 ‘소송촉진 등에 관한 특례법’ 제3조에 "
         "따라 연 12%의 비율에 의한 지연손해금을 청구하겠습니다. "
         "(보조 청구: 조합이 임의로 약속한 환불 시기 ‘2024년 3월 ~ 8월’ 도과를 기산점으로 "
         "삼는 것도 가능함을 부기합니다.)",
         size=11, left_indent=0.7, space_after=10)

    para(doc,
         "라. 아울러 본 통고서 도달일로부터 7일 이내에 ① 통고인의 분담금 납부내역 일체, "
         "② 추가분담금 4,000만 원 의결 총회의 회의록·결의서 사본, "
         "③ 2019년 가입 이후 현재까지 조합 분담금의 사용내역(특히 ‘홍보관 운영비’ 항목)에 "
         "관한 회계자료 일체를 통고인에게 송부하여 주실 것을 요청합니다.",
         size=11, left_indent=0.7, space_after=12)

    # 항목 4. 경고
    para(doc, "5. 미이행 시 조치 예고", size=12, bold=True, space_after=6)
    para(doc,
         "가. 본 통고서 도달일로부터 14일이 경과하도록 위 청구가 이행되지 아니할 경우, "
         "통고인은 즉시 관할 법원에 ‘분담금반환 청구의 소’를 제기할 것임을 알려드립니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "나. 위 민사 소송과는 별도로, 조합 분담금의 본래 용도 외 사용 정황 및 "
         "허위·과장 광고 정황에 관하여는 업무대행사 및 업무대행총괄 김성수 이사에 대하여 "
         "상법 제401조에 의한 손해배상 청구 및 형법 제355조(업무상 횡령·배임)에 의한 "
         "형사고소를 함께 검토할 예정임을 분명히 통지합니다.",
         size=11, left_indent=0.7, space_after=6)
    para(doc,
         "다. 또한 본 통고서의 도달 사실 및 그 내용은 향후 소송에서 ‘이행 최고’의 효력과 "
         "‘소멸시효 중단’의 효력이 있음을 부기합니다(민법 제174조).",
         size=11, left_indent=0.7, space_after=14)

    # 결어
    para(doc,
         "이상의 점을 깊이 헤아리시어, 본 통고서에서 정한 기한 내에 통고인의 청구를 "
         "원만히 이행하여 주실 것을 정중히 요청드립니다.",
         size=11, space_after=20)

    # 작성일
    para(doc, "                                                          2026년     월     일",
         size=11, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=10)

    para(doc, "통  고  인 :  고  영  진             (인)",
         size=12, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    para(doc, "                  연락처 :                              ",
         size=11, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=20)

    hr(doc)

    # 별첨
    para(doc, "[ 별   첨 ]", size=11, bold=True, space_after=4)
    para(doc,
         "1. 조합 가입계약서 사본 1부.\n"
         "2. 분담금 입금내역 일체(통장사본·계좌이체 내역) 1부.\n"
         "3. 2023년 3월 15일 김성수 이사 면담 녹취록 1부.\n"
         "4. 녹음 원본파일(USB) 1매.   끝.",
         size=11, left_indent=0.5, space_after=10)

    out = "/home/user/geopolitics/output/내용증명_연신내지역주택조합_분담금반환청구_고영진.docx"
    doc.save(out)
    print(out)


if __name__ == "__main__":
    build()
