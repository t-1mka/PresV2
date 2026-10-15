"""Build the 14-slide academic presentation "Будущее труда в автоматизированной экономике".

Design rules from the brief:
- 16:9, light neutral background, dark text, blue/orange accents.
- Arial (Liberation Sans is metric-compatible, so the PPTX/PDF match).
- Body text >= 16 pt, headings >= 20 pt.
- No animation. Speaker notes carry the spoken text, not the slides.
- Slide number on every slide.
"""

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

ROOT = os.path.dirname(os.path.abspath(__file__))
CH = os.path.join(ROOT, "assets", "charts")

BLUE = RGBColor(0x1F, 0x4E, 0x79)
BLUE_L = RGBColor(0xDE, 0xEB, 0xF7)
ORANGE = RGBColor(0xC5, 0x5A, 0x11)
ORANGE_L = RGBColor(0xFB, 0xE5, 0xD6)
INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x59, 0x59, 0x59)
LINE = RGBColor(0xD0, 0xD0, 0xD0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)

FONT = "Arial"

SW = Inches(13.333)
SH = Inches(7.5)
MARGIN = Inches(0.62)
CONTENT_W = SW - 2 * MARGIN

# Author / event metadata for the title slide.
AUTHOR = "Иванов Иван Иванович"
GROUP = "Группа ЭК-21, экономический факультет"
ORG = "Учебная организация"
EMAIL = "student@example.edu"
EVENT = "Практическое занятие №10"
DATE = "15 октября 2026 г."


def set_font(run, size, bold=False, color=INK, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_textbox(slide, x, y, w, h, text, size=16, bold=False, color=INK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.12,
                italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = line_spacing
        set_font(para.add_run(), size, bold=bold, color=color, italic=italic)
        para.runs[0].text = line
    return box


def add_rect(slide, x, y, w, h, fill, line=None, line_w=0.75):
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)
    shape.shadow.inherit = False
    return shape


def blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SW, SH, BG)
    return slide


def add_slide_number(slide, n):
    add_textbox(slide, SW - Inches(1.15), SH - Inches(0.52), Inches(0.55), Inches(0.3),
                str(n), size=16, color=GREY, align=PP_ALIGN.RIGHT)


def add_title(slide, text, sub=None):
    """Heading band at the top of a content slide."""
    add_rect(slide, MARGIN, Inches(0.42), Inches(0.09), Inches(0.62), BLUE)
    add_textbox(slide, MARGIN + Inches(0.24), Inches(0.40), CONTENT_W - Inches(0.3), Inches(0.7),
                text, size=28, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
    y = Inches(1.22)
    if sub:
        add_textbox(slide, MARGIN + Inches(0.24), Inches(1.10), CONTENT_W - Inches(0.3), Inches(0.4),
                    sub, size=17, color=GREY)
        y = Inches(1.62)
    return y


def add_bullets(slide, x, y, w, h, items, size=17, gap=10):
    """Bulleted takeaways. items = list of (bold_lead, rest) or plain strings."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.line_spacing = 1.18
        para.space_after = Pt(gap)
        marker = para.add_run()
        set_font(marker, size, bold=True, color=ORANGE)
        marker.text = "▪  "
        if isinstance(item, tuple):
            lead, rest = item
            r1 = para.add_run()
            set_font(r1, size, bold=True, color=INK)
            r1.text = lead
            if rest:
                r2 = para.add_run()
                set_font(r2, size, color=INK)
                r2.text = rest
        else:
            r = para.add_run()
            set_font(r, size, color=INK)
            r.text = item
    return box


def add_picture(slide, path, x, y, w, h):
    """Place an image at 1:1.

    Charts are rendered at exactly the size they occupy here, so scaling them
    would also shrink their labels below the 16 pt minimum.
    """
    from PIL import Image

    with Image.open(path) as im:
        iw, ih = im.size
        assert abs(iw / ih - w / h) < 0.02, (
            f"aspect mismatch for {os.path.basename(path)}: "
            f"image {iw}x{ih} vs box {w}x{h}")
    return slide.shapes.add_picture(path, x, y, w, h)


def style_table(table, header_size=17, body_size=16, header_fill=BLUE, first_col_bold=False):
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.fore_color.rgb = WHITE if r % 2 else RGBColor(0xF2, 0xF2, 0xF2)
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.LEFT
                for run in para.runs:
                    if r == 0:
                        set_font(run, header_size, bold=True, color=WHITE)
                    else:
                        set_font(run, body_size, bold=(first_col_bold and c == 0), color=INK)


def add_table(slide, x, y, w, h, data, col_widths=None, header_size=17, body_size=16,
              first_col_bold=False, header_fill=BLUE):
    rows, cols = len(data), len(data[0])
    shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = shape.table
    table.first_row = True
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            table.columns[i].width = Emu(int(w * cw / total))
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            table.cell(r, c).text = str(val)
    style_table(table, header_size, body_size, header_fill, first_col_bold)
    return table


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ====================================================================== slides
def slide_01(prs):
    slide = blank(prs)
    add_rect(slide, 0, 0, SW, Inches(0.22), BLUE)
    add_rect(slide, 0, Inches(0.22), SW, Inches(0.06), ORANGE)

    add_picture(slide, os.path.join(CH, "01_logo.png"),
                    SW - Inches(2.0), Inches(0.62), Inches(1.35), Inches(1.35))

    add_textbox(slide, MARGIN, Inches(1.42), Inches(9.9), Inches(0.4),
                "УЧЕБНАЯ ПРЕЗЕНТАЦИЯ · " + EVENT.upper(), size=16, bold=True, color=ORANGE)
    add_textbox(slide, MARGIN, Inches(1.88), Inches(10.2), Inches(2.20),
                "Будущее труда\nв автоматизированной\nэкономике", size=40, bold=True,
                color=BLUE, line_spacing=1.0)
    add_rect(slide, MARGIN, Inches(4.20), Inches(2.2), Inches(0.05), ORANGE)
    add_textbox(slide, MARGIN, Inches(4.42), Inches(10.4), Inches(0.9),
                "Не только ИИ: как роботы и алгоритмы вместе меняют рынок труда",
                size=20, color=INK, line_spacing=1.15)

    meta = [
        ("Автор", AUTHOR),
        ("Организация", GROUP + " · " + ORG),
        ("Email", EMAIL),
        ("Мероприятие", EVENT),
        ("Дата", DATE),
    ]
    y = Inches(5.40)
    for i, (k, v) in enumerate(meta):
        row_y = y + Inches(i * 0.34)
        add_textbox(slide, MARGIN, row_y, Inches(1.7), Inches(0.3), k, size=16, color=GREY)
        add_textbox(slide, MARGIN + Inches(1.8), row_y, Inches(9.6), Inches(0.3), v,
                    size=16, bold=True, color=INK)
    add_slide_number(slide, 1)
    set_notes(slide,
              "Тема касается каждого. Автоматизация — это не одна технология, а два потока: "
              "цифровой и физический. И они по-разному влияют на разные профессии.\n\n"
              "Зацепка для зала (первые 20 секунд): «Кто из вас считает, что его профессия может "
              "исчезнуть из-за ИИ?» — пауза 2–3 секунды.")
    return slide


def slide_02(prs):
    slide = blank(prs)
    add_title(slide, "План доклада")
    items = [
        ("01", "Два вектора автоматизации", "ИИ и роботы: в чём разница"),
        ("02", "Профессии под ударом", "Кто уязвим и почему"),
        ("03", "Масштаб изменений", "Статистика и география"),
        ("04", "Кейсы: транспорт и логистика", "Что уже происходит"),
        ("05", "Экономика и навыки", "Когда робот дешевле человека"),
        ("06", "Выводы", "Что делать и что учить"),
    ]
    col_w = Inches(5.86)
    for idx, (num, title, sub) in enumerate(items):
        col = idx % 2
        row = idx // 2
        x = MARGIN + col * (col_w + Inches(0.33))
        y = Inches(1.50) + row * Inches(1.68)
        add_rect(slide, x, y, col_w, Inches(1.42), WHITE, line=LINE)
        add_rect(slide, x, y, Inches(0.08), Inches(1.42), BLUE if col == 0 else ORANGE)
        add_textbox(slide, x + Inches(0.26), y + Inches(0.22), Inches(0.9), Inches(0.75),
                    num, size=30, bold=True, color=BLUE_L)
        add_textbox(slide, x + Inches(1.15), y + Inches(0.20), col_w - Inches(1.4), Inches(0.5),
                    title, size=19, bold=True, color=BLUE)
        add_textbox(slide, x + Inches(1.15), y + Inches(0.78), col_w - Inches(1.4), Inches(0.45),
                    sub, size=16, color=GREY)
    add_slide_number(slide, 2)
    set_notes(slide,
              "Мы последовательно разберём, что такое автоматизированная экономика, "
              "кого она затрагивает, и что делать, чтобы остаться востребованным.\n\n"
              "Регламент — 10–12 минут, шесть блоков.")
    return slide


def slide_03(prs):
    slide = blank(prs)
    add_title(slide, "ИИ и роботы: два вектора автоматизации")
    add_bullets(slide, MARGIN, Inches(1.35), Inches(5.85), Inches(3.3), [
        ("ИИ: ", "заменяет когнитивные задачи — анализ, текст, код"),
        ("Роботы: ", "заменяют физический труд — сборку, перемещение"),
        ("Скорость: ", "ИИ внедряется быстро, роботы — медленнее"),
        ("Стоимость: ", "ИИ — подписка, роботы — покупка и сервис"),
    ], size=17, gap=14)
    add_table(slide, Inches(6.55), Inches(1.35), Inches(6.16), Inches(2.5), [
        ["Параметр", "ИИ (цифровая)", "Роботы (физическая)"],
        ["Что заменяет", "Когнитивные задачи", "Физический труд"],
        ["Кого затрагивает", "«Белые воротнички»", "«Синие воротнички»"],
        ["Скорость внедрения", "Быстро", "Медленнее"],
        ["Стоимость внедрения", "Низкая", "Высокая"],
    ], col_widths=[1.5, 1.6, 1.7], first_col_bold=True)
    add_picture(slide, os.path.join(CH, "03_vectors_bar.png"),
                    Inches(6.45), Inches(4.45), Inches(6.23), Inches(2.35))
    add_slide_number(slide, 3)
    set_notes(slide,
              "Долгое время говорили только об ИИ и офисных профессиях. Но роботы — это отдельная волна.\n\n"
              "McKinsey: ИИ может автоматизировать 44% рабочих часов, роботы — ещё 13%. "
              "Вместе — 57%.")
    return slide


def slide_04(prs):
    slide = blank(prs)
    add_title(slide, "Профессии с высоким риском автоматизации ИИ")
    add_bullets(slide, MARGIN, Inches(1.45), Inches(4.6), Inches(4.6), [
        ("Программисты — ", "до 94% задач покрываются LLM"),
        ("Офисный персонал — ", "около 90%"),
        ("Юристы, бизнес, финансы — ", "высокий риск"),
        ("Парадокс: ", "под удар попадают высокооплачиваемые и образованные"),
    ], size=17, gap=16)
    add_picture(slide, os.path.join(CH, "04_ai_exposure.png"),
                    Inches(5.75), Inches(1.45), Inches(6.93), Inches(5.25))
    add_slide_number(slide, 4)
    set_notes(slide,
              "Это переворачивает традиционное представление: автоматизация угрожает не только "
              "низкоквалифицированным, но и самым образованным.\n\n"
              "Источник: ILO — GenAI and Jobs (2025), экспозиция профессий.")
    return slide


def slide_05(prs):
    slide = blank(prs)
    add_title(slide, "Профессии с высоким риском автоматизации роботами")
    add_bullets(slide, MARGIN, Inches(1.45), Inches(4.85), Inches(4.6), [
        ("9 из 10 ", "наиболее уязвимых — операторы транспортных средств"),
        ("Водители такси — ", "97% работы может выполнить робот"),
        ("Сельхозтехника, погрузчики, шаттлы — ", "высокий риск"),
        ("Барьер — ", "экономический, а не технологический"),
    ], size=17, gap=16)
    add_picture(slide, os.path.join(CH, "05_robot_exposure.png"),
                    Inches(5.75), Inches(1.45), Inches(6.93), Inches(5.25))
    add_slide_number(slide, 5)
    set_notes(slide,
              "Вождение — первая массовая профессия, где роботы уже конкурентоспособны. "
              "Это не будущее — это происходит сейчас.\n\n"
              "Индекс экспозиции: 0 — робот не справляется, 3 — справляется полностью.")
    return slide


def slide_06(prs):
    slide = blank(prs)
    add_title(slide, "Кто именно теряет работу: два профиля")
    add_picture(slide, os.path.join(CH, "06_profiles.png"),
                    MARGIN, Inches(1.42), CONTENT_W, Inches(4.15))
    add_table(slide, MARGIN, Inches(5.62), CONTENT_W, Inches(1.4), [
        ["Параметр", "Уязвимые для ИИ", "Уязвимые для роботов"],
        ["Образование", "Высокое", "8% с высшим"],
        ["Доход", "Высокий", "≈ $23/час против $53 у защищённых"],
    ], col_widths=[1.5, 1.6, 2.6], first_col_bold=True, body_size=16)
    add_slide_number(slide, 6)
    set_notes(slide,
              "Автоматизация не нейтральна. Последствия для «синих воротничков» тяжелее: "
              "у них меньше ресурсов для переквалификации.\n\n"
              "Безработица: 5,2% у уязвимых для роботов против 2,2% у защищённых. "
              "69% уязвимых для роботов — мужчины.")
    return slide


def slide_07(prs):
    slide = blank(prs)
    add_title(slide, "Сколько рабочих мест исчезнет и появится")
    add_bullets(slide, MARGIN, Inches(1.45), Inches(4.7), Inches(4.6), [
        ("Исчезнет к 2030: ", "92 млн рабочих мест"),
        ("Появится новых: ", "170 млн"),
        ("Чистый прирост: ", "+78 млн"),
        ("Робототехника — ", "крупнейший нетто-вытеснитель: −5 млн"),
    ], size=17, gap=16)
    add_picture(slide, os.path.join(CH, "07_jobs.png"),
                    Inches(5.90), Inches(1.70), Inches(6.78), Inches(4.80))
    add_slide_number(slide, 7)
    set_notes(slide,
              "Баланс положительный, но проблема не в количестве, а в несовпадении навыков. "
              "Роботы чаще заменяют, ИИ чаще дополняет.\n\n"
              "Источник: WEF Future of Jobs Report 2025.")
    return slide


def slide_08(prs):
    slide = blank(prs)
    add_title(slide, "Где роботов больше всего")
    add_bullets(slide, MARGIN, Inches(1.45), Inches(4.2), Inches(4.6), [
        ("Южная Корея — ", "1220 роботов на 10 000 работников"),
        ("Сингапур — ", "818, Германия — 449"),
        ("США — ", "307, Китай — 166"),
        ("Мировой средний — ", "132"),
    ], size=17, gap=16)
    add_picture(slide, os.path.join(CH, "08_robot_density.png"),
                    Inches(5.40), Inches(1.50), Inches(7.28), Inches(5.20))
    add_slide_number(slide, 8)
    set_notes(slide,
              "Разрыв в 9 раз между лидером и средним уровнем. Страны с высокой автоматизацией "
              "первыми почувствуют её последствия.\n\n"
              "Источник: IFR World Robotics 2025.")
    return slide


def slide_09(prs):
    slide = blank(prs)
    add_title(slide, "Автономные грузовики: самая уязвимая профессия")
    add_bullets(slide, MARGIN, Inches(1.45), Inches(5.0), Inches(4.6), [
        ("В США — ", "3,5 млн водителей грузовиков"),
        ("Роботы выполняют ", "97% работы таксиста"),
        ("В местах работы автономных такси ", "зарплаты упали на 6,9%"),
        ("24 000 ", "водительских позиций ликвидированы (2023)"),
    ], size=17, gap=16)
    add_picture(slide, os.path.join(CH, "09_trucking.png"),
                    Inches(6.00), Inches(1.75), Inches(6.68), Inches(4.80))
    add_slide_number(slide, 9)
    set_notes(slide,
              "Это самый уязвимый сектор. Массовая потеря рабочих мест возможна к 2026–2030.\n\n"
              "Источник: UC ITS — Autonomous Trucking (2026).")
    return slide


def slide_10(prs):
    slide = blank(prs)
    add_title(slide, "Склады и логистика: роботы уже рядом")
    add_bullets(slide, MARGIN, Inches(1.35), CONTENT_W, Inches(1.5), [
        ("Amazon: ", "более 1 млн роботов на складах; за полгода 2026 заказано 18 000 роботов на $1,2 млрд"),
        ("55% руководителей ", "ожидают сокращения найма; при этом роботы создают новые роли — техники, интеграторы"),
    ], size=17, gap=10)
    add_picture(slide, os.path.join(CH, "10_logistics.png"),
                    Inches(1.55), Inches(3.20), Inches(10.23), Inches(3.45))
    add_slide_number(slide, 10)
    set_notes(slide,
              "Логистика — пример трансформации, а не просто замены. "
              "Вопрос в том, успеют ли работники переквалифицироваться.")
    return slide


def slide_11(prs):
    slide = blank(prs)
    add_title(slide, "Когда робот дешевле человека")
    add_bullets(slide, MARGIN, Inches(1.35), Inches(4.6), Inches(4.6), [
        ("74% ", "физических задач роботы способны выполнять технически"),
        ("0,3% ", "— экономически конкурентоспособны сегодня"),
        ("При снижении цены на 20% ", "число замен вырастет с 300 тыс. до 2,8 млн"),
        ("Барьер ", "не технологический, а экономический"),
    ], size=17, gap=14)
    add_picture(slide, os.path.join(CH, "11_economics.png"),
                    Inches(5.45), Inches(1.90), Inches(7.23), Inches(4.55))
    add_slide_number(slide, 11)
    set_notes(slide,
              "Как только роботы подешевеют, замена ускорится лавинообразно. "
              "Это ключевой экономический триггер.\n\n"
              "Источник: Anthropic — Labor Market Impacts (2026); McKinsey (2025).")
    return slide


def slide_12(prs):
    slide = blank(prs)
    add_title(slide, "Кто выиграет от автоматизации")
    add_bullets(slide, MARGIN, Inches(1.35), Inches(5.1), Inches(4.6), [
        ("Инженеры-робототехники ", "— в топе WEF по росту спроса"),
        ("Техники по обслуживанию роботов ", "— $70–80 тыс./год без 4-летнего диплома"),
        ("Эмоциональный интеллект ", "— спрос вырос на 173%"),
        ("Креативность ", "— спрос вырос на 168%"),
        ("43% рабочих часов ", "остаются уникально человеческими"),
    ], size=17, gap=11)
    add_picture(slide, os.path.join(CH, "12_skills.png"),
                    Inches(6.00), Inches(1.95), Inches(6.68), Inches(4.55))
    add_slide_number(slide, 12)
    set_notes(slide,
              "Технические навыки устаревают. «Power skills» — эмпатия, креативность, суждение — "
              "становятся главным конкурентным преимуществом человека.")
    return slide


def slide_13(prs):
    slide = blank(prs)
    add_title(slide, "Основные результаты")
    add_bullets(slide, MARGIN, Inches(1.35), Inches(7.1), Inches(4.0), [
        ("Два вектора — разные последствия. ", "ИИ бьёт по офисным профессиям быстро и дёшево; роботы — по физическим, медленнее, но глубже."),
        ("Риск не в безработице, а в неравенстве. ", "Уязвимые для роботов — менее образованные, ниже оплачиваемые, с меньшими ресурсами для адаптации."),
        ("Будущее — партнёрство. ", "Выиграют те, кто умеет направлять, а не конкурировать с машинами."),
        ("Ключевой навык — адаптивность. ", "Готовность к смене 3–5 профессий за карьеру."),
    ], size=17, gap=12)
    add_picture(slide, os.path.join(CH, "13_conclusion.png"),
                    Inches(7.95), Inches(2.50), Inches(4.75), Inches(2.30))
    add_slide_number(slide, 13)
    set_notes(slide,
              "Автоматизация — не конец работы. Это переопределение того, что значит работать. "
              "И те, кто понимает оба вектора, будут лучше подготовлены к будущему.")
    return slide


def slide_14(prs):
    slide = blank(prs)
    add_rect(slide, 0, 0, SW, Inches(0.22), BLUE)
    add_rect(slide, 0, Inches(0.22), SW, Inches(0.06), ORANGE)
    add_textbox(slide, MARGIN, Inches(1.15), Inches(8.4), Inches(1.0),
                "Спасибо за внимание", size=40, bold=True, color=BLUE)
    add_textbox(slide, MARGIN, Inches(2.15), Inches(8.4), Inches(0.5),
                "Вопросы и обсуждение", size=20, color=ORANGE)

    add_textbox(slide, MARGIN, Inches(3.05), Inches(4.3), Inches(0.4),
                "КОНТАКТЫ", size=16, bold=True, color=GREY)
    add_textbox(slide, MARGIN, Inches(3.45), Inches(5.2), Inches(2.4),
                AUTHOR + "\n" + GROUP + "\n" + ORG + "\n" + EMAIL + "\n" + DATE,
                size=17, color=INK, line_spacing=1.3)

    add_textbox(slide, Inches(6.35), Inches(3.05), Inches(6.4), Inches(0.4),
                "ИСТОЧНИКИ — НАВЕДИТЕ КАМЕРУ", size=16, bold=True, color=GREY)
    add_picture(slide, os.path.join(CH, "14_qr_sources.png"),
                    Inches(6.35), Inches(3.45), Inches(1.9), Inches(1.9))
    add_textbox(slide, Inches(8.45), Inches(3.42), Inches(4.4), Inches(2.6),
                "WEF Future of Jobs Report 2025\nMcKinsey — Agents, Robots, and Us\n"
                "IFR World Robotics 2025\nAnthropic — Labor Market Impacts (2026)\n"
                "ILO — GenAI and Jobs (2025)\nUC ITS — Autonomous Trucking (2026)\n\n"
                "github.com/t-1mka/PresV2",
                size=16, color=INK, line_spacing=1.22)
    add_slide_number(slide, 14)
    set_notes(slide,
              "Готов ответить на ваши вопросы. Особенно интересно обсудить, "
              "какой из двух векторов коснётся лично вас.\n\n"
              "QR-код ведёт на список источников в репозитории.")
    return slide


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    for fn in (slide_01, slide_02, slide_03, slide_04, slide_05, slide_06, slide_07,
               slide_08, slide_09, slide_10, slide_11, slide_12, slide_13, slide_14):
        fn(prs)
    out = os.path.join(ROOT, "Будущее_труда_в_автоматизированной_экономике.pptx")
    prs.save(out)
    print("saved", out, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))


if __name__ == "__main__":
    main()
