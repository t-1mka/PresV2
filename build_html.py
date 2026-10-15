"""Render an online (HTML) version of the presentation.

The HTML mirrors the PPTX slide by slide: same headings, same data, same
charts (embedded as base64 PNG), plus the speaker notes behind a toggle. It is
a static page with no dependencies, so it can be served straight from GitHub
Pages or opened locally.
"""

import base64
import html
import os
import re

from pptx import Presentation

ROOT = os.path.dirname(os.path.abspath(__file__))
CH = os.path.join(ROOT, "assets", "charts")
PPTX = os.path.join(ROOT, "Будущее_труда_в_автоматизированной_экономике.pptx")
OUT = os.path.join(ROOT, "index.html")

# Structured content for the web version. Keeping it explicit (rather than
# scraping the PPTX) keeps the markup semantic and the notes readable.
SLIDES = [
    dict(
        n=1, kind="title",
        title="Будущее труда в автоматизированной экономике",
        sub="Не только ИИ: как роботы и алгоритмы вместе меняют рынок труда",
        meta=[("Автор", "Иванов Иван Иванович"),
              ("Организация", "Группа ЭК-21, экономический факультет · Учебная организация"),
              ("Email", "student@example.edu"),
              ("Мероприятие", "Практическое занятие №10"),
              ("Дата", "15 октября 2026 г.")],
        notes="Тема касается каждого. Автоматизация — это не одна технология, а два потока: "
              "цифровой и физический. И они по-разному влияют на разные профессии.",
    ),
    dict(n=2, title="План доклада", notes=(
        "Мы последовательно разберём, что такое автоматизированная экономика, кого она "
        "затрагивает, и что делать, чтобы остаться востребованным."),
        cards=[("01", "Два вектора автоматизации", "ИИ и роботы: в чём разница"),
               ("02", "Профессии под ударом", "Кто уязвим и почему"),
               ("03", "Масштаб изменений", "Статистика и география"),
               ("04", "Кейсы: транспорт и логистика", "Что уже происходит"),
               ("05", "Экономика и навыки", "Когда робот дешевле человека"),
               ("06", "Выводы", "Что делать и что учить")]),
    dict(n=3, title="ИИ и роботы: два вектора автоматизации",
         bullets=[("ИИ: ", "заменяет когнитивные задачи — анализ, текст, код"),
                  ("Роботы: ", "заменяют физический труд — сборку, перемещение"),
                  ("Скорость: ", "ИИ внедряется быстро, роботы — медленнее"),
                  ("Стоимость: ", "ИИ — подписка, роботы — покупка и сервис")],
         table=[["Параметр", "ИИ (цифровая)", "Роботы (физическая)"],
                ["Что заменяет", "Когнитивные задачи", "Физический труд"],
                ["Кого затрагивает", "«Белые воротнички»", "«Синие воротнички»"],
                ["Скорость внедрения", "Быстро", "Медленнее"],
                ["Стоимость внедрения", "Низкая", "Высокая"]],
         chart="03_vectors_bar.png",
         notes="Долгое время говорили только об ИИ и офисных профессиях. Но роботы — это "
               "отдельная волна. McKinsey: ИИ может автоматизировать 44% рабочих часов, "
               "роботы — ещё 13%. Вместе — 57%."),
    dict(n=4, title="Профессии с высоким риском автоматизации ИИ",
         bullets=[("Программисты — ", "до 94% задач покрываются LLM"),
                  ("Офисный персонал — ", "около 90%"),
                  ("Юристы, бизнес, финансы — ", "высокий риск"),
                  ("Парадокс: ", "под удар попадают высокооплачиваемые и образованные")],
         chart="04_ai_exposure.png",
         notes="Это переворачивает традиционное представление: автоматизация угрожает не "
               "только низкоквалифицированным, но и самым образованным. "
               "Источник: ILO — GenAI and Jobs (2025)."),
    dict(n=5, title="Профессии с высоким риском автоматизации роботами",
         bullets=[("9 из 10 ", "наиболее уязвимых — операторы транспортных средств"),
                  ("Водители такси — ", "97% работы может выполнить робот"),
                  ("Сельхозтехника, погрузчики, шаттлы — ", "высокий риск"),
                  ("Барьер — ", "экономический, а не технологический")],
         chart="05_robot_exposure.png",
         notes="Вождение — первая массовая профессия, где роботы уже конкурентоспособны. "
               "Это не будущее — это происходит сейчас. Индекс экспозиции: 0 — робот не "
               "справляется, 3 — справляется полностью."),
    dict(n=6, title="Кто именно теряет работу: два профиля",
         chart="06_profiles.png",
         table=[["Параметр", "Уязвимые для ИИ", "Уязвимые для роботов"],
                ["Образование", "Высокое", "8% с высшим"],
                ["Доход", "Высокий", "≈ $23/час против $53 у защищённых"]],
         notes="Автоматизация не нейтральна. Последствия для «синих воротничков» тяжелее: "
               "у них меньше ресурсов для переквалификации. Безработица: 5,2% у уязвимых "
               "для роботов против 2,2% у защищённых."),
    dict(n=7, title="Сколько рабочих мест исчезнет и появится",
         bullets=[("Исчезнет к 2030: ", "92 млн рабочих мест"),
                  ("Появится новых: ", "170 млн"),
                  ("Чистый прирост: ", "+78 млн"),
                  ("Робототехника — ", "крупнейший нетто-вытеснитель: −5 млн")],
         chart="07_jobs.png",
         notes="Баланс положительный, но проблема не в количестве, а в несовпадении навыков. "
               "Роботы чаще заменяют, ИИ чаще дополняет. Источник: WEF Future of Jobs 2025."),
    dict(n=8, title="Где роботов больше всего",
         bullets=[("Южная Корея — ", "1220 роботов на 10 000 работников"),
                  ("Сингапур — ", "818, Германия — 449"),
                  ("США — ", "307, Китай — 166"),
                  ("Мировой средний — ", "132")],
         chart="08_robot_density.png",
         notes="Разрыв в 9 раз между лидером и средним уровнем. Страны с высокой "
               "автоматизацией первыми почувствуют её последствия. "
               "Источник: IFR World Robotics 2025."),
    dict(n=9, title="Автономные грузовики: самая уязвимая профессия",
         bullets=[("В США — ", "3,5 млн водителей грузовиков"),
                  ("Роботы выполняют ", "97% работы таксиста"),
                  ("В местах работы автономных такси ", "зарплаты упали на 6,9%"),
                  ("24 000 ", "водительских позиций ликвидированы (2023)")],
         chart="09_trucking.png",
         notes="Это самый уязвимый сектор. Массовая потеря рабочих мест возможна к "
               "2026–2030. Источник: UC ITS — Autonomous Trucking (2026)."),
    dict(n=10, title="Склады и логистика: роботы уже рядом",
         bullets=[("Amazon: ", "более 1 млн роботов на складах; за полгода 2026 заказано "
                               "18 000 роботов на $1,2 млрд"),
                  ("55% руководителей ", "ожидают сокращения найма; при этом роботы создают "
                                         "новые роли — техники, интеграторы")],
         chart="10_logistics.png",
         notes="Логистика — пример трансформации, а не просто замены. Вопрос в том, успеют "
               "ли работники переквалифицироваться."),
    dict(n=11, title="Когда робот дешевле человека",
         bullets=[("74% ", "физических задач роботы способны выполнять технически"),
                  ("0,3% ", "— экономически конкурентоспособны сегодня"),
                  ("При снижении цены на 20% ", "число замен вырастет с 300 тыс. до 2,8 млн"),
                  ("Барьер ", "не технологический, а экономический")],
         chart="11_economics.png",
         notes="Как только роботы подешевеют, замена ускорится лавинообразно. Это ключевой "
               "экономический триггер. Источник: Anthropic — Labor Market Impacts (2026); "
               "McKinsey (2025)."),
    dict(n=12, title="Кто выиграет от автоматизации",
         bullets=[("Инженеры-робототехники ", "— в топе WEF по росту спроса"),
                  ("Техники по обслуживанию роботов ", "— $70–80 тыс./год без 4-летнего диплома"),
                  ("Эмоциональный интеллект ", "— спрос вырос на 173%"),
                  ("Креативность ", "— спрос вырос на 168%"),
                  ("43% рабочих часов ", "остаются уникально человеческими")],
         chart="12_skills.png",
         notes="Технические навыки устаревают. «Power skills» — эмпатия, креативность, "
               "суждение — становятся главным конкурентным преимуществом человека."),
    dict(n=13, title="Основные результаты",
         bullets=[("Два вектора — разные последствия. ",
                   "ИИ бьёт по офисным профессиям быстро и дёшево; роботы — по физическим, "
                   "медленнее, но глубже."),
                  ("Риск не в безработице, а в неравенстве. ",
                   "Уязвимые для роботов — менее образованные, ниже оплачиваемые, с меньшими "
                   "ресурсами для адаптации."),
                  ("Будущее — партнёрство. ",
                   "Выиграют те, кто умеет направлять, а не конкурировать с машинами."),
                  ("Ключевой навык — адаптивность. ",
                   "Готовность к смене 3–5 профессий за карьеру.")],
         chart="13_conclusion.png",
         notes="Автоматизация — не конец работы. Это переопределение того, что значит "
               "работать. И те, кто понимает оба вектора, будут лучше подготовлены к будущему."),
    dict(n=14, kind="final", title="Спасибо за внимание", sub="Вопросы и обсуждение",
         contacts=["Иванов Иван Иванович", "Группа ЭК-21, экономический факультет",
                   "Учебная организация", "student@example.edu", "15 октября 2026 г."],
         sources=["WEF Future of Jobs Report 2025",
                  "McKinsey — Agents, Robots, and Us (2025)",
                  "IFR World Robotics 2025",
                  "Anthropic — Labor Market Impacts (2026)",
                  "ILO — GenAI and Jobs (2025)",
                  "UC ITS — Autonomous Trucking (2026)"],
         link="https://github.com/t-1mka/PresV2/blob/main/SOURCES.md",
         notes="Готов ответить на ваши вопросы. Особенно интересно обсудить, какой из двух "
               "векторов коснётся лично вас."),
]


def data_uri(name):
    with open(os.path.join(CH, name), "rb") as fh:
        return "data:image/png;base64," + base64.b64encode(fh.read()).decode()


def esc(s):
    return html.escape(s)


def render_table(rows):
    out = ["<table>"]
    out.append("<thead><tr>" + "".join(f"<th>{esc(c)}</th>" for c in rows[0]) + "</tr></thead>")
    out.append("<tbody>")
    for row in rows[1:]:
        out.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
    out.append("</tbody></table>")
    return "\n".join(out)


def render_bullets(bullets):
    items = []
    for lead, rest in bullets:
        items.append(f"<li><strong>{esc(lead)}</strong>{esc(rest)}</li>")
    return "<ul class='bullets'>" + "".join(items) + "</ul>"


def render_slide(s):
    n = s["n"]
    kind = s.get("kind", "")
    cls = "slide" + (f" slide--{kind}" if kind else "")
    parts = [f"<section class='{cls}' id='slide-{n}' aria-label='Слайд {n}'>"]
    parts.append(f"<div class='num'>{n}</div>")

    if kind == "title":
        parts.append("<p class='kicker'>УЧЕБНАЯ ПРЕЗЕНТАЦИЯ · ПРАКТИЧЕСКОЕ ЗАНЯТИЕ №10</p>")
        parts.append(f"<h1>{esc(s['title'])}</h1>")
        parts.append(f"<p class='sub'>{esc(s['sub'])}</p>")
        rows = "".join(f"<div class='meta-row'><span>{esc(k)}</span><b>{esc(v)}</b></div>"
                       for k, v in s["meta"])
        parts.append(f"<div class='meta'>{rows}</div>")
    elif kind == "final":
        parts.append(f"<h2 class='big'>{esc(s['title'])}</h2>")
        parts.append(f"<p class='sub accent'>{esc(s['sub'])}</p>")
        contacts = "".join(f"<div>{esc(c)}</div>" for c in s["contacts"])
        srcs = "".join(f"<li>{esc(x)}</li>" for x in s["sources"])
        parts.append(
            "<div class='final-grid'>"
            f"<div class='col'><h3>Контакты</h3>{contacts}</div>"
            "<div class='col'><h3>Источники — наведите камеру</h3>"
            "<div class='qr-row'>"
            f"<img class='qr' src='{data_uri('14_qr_sources.png')}' alt='QR-код на список источников'>"
            f"<ul class='srclist'>{srcs}</ul></div>"
            f"<p class='link'><a href='{esc(s['link'])}'>{esc(s['link'])}</a></p>"
            "</div></div>")
    else:
        parts.append(f"<h2>{esc(s['title'])}</h2>")
        body = ["<div class='cols'>"]
        if s.get("cards"):
            body.append("<div class='cards'>")
            for num, t, sub in s["cards"]:
                body.append(f"<div class='card'><span class='cnum'>{esc(num)}</span>"
                            f"<h3>{esc(t)}</h3><p>{esc(sub)}</p></div>")
            body.append("</div>")
        else:
            left = []
            if s.get("bullets"):
                left.append(render_bullets(s["bullets"]))
            if s.get("table"):
                left.append(render_table(s["table"]))
            body.append("<div class='text-col'>" + "".join(left) + "</div>")
            if s.get("chart"):
                body.append(f"<div class='chart-col'><img src='{data_uri(s['chart'])}' "
                            f"alt='Диаграмма к слайду {n}'></div>")
        body.append("</div>")
        parts.append("".join(body))

    parts.append("<details class='notes'><summary>Речь докладчика</summary>"
                 f"<p>{esc(s['notes'])}</p></details>")
    parts.append("</section>")
    return "\n".join(parts)


CSS = """
:root{--blue:#1F4E79;--blue-l:#DEBEF7;--blue-l2:#8FB8DC;--orange:#C55A11;
--ink:#1A1A1A;--grey:#595959;--line:#D0D0D0;--bg:#FAFAFA;--card:#F2F2F2}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:#EDEDED;color:var(--ink);
font-family:Arial,"Liberation Sans",Helvetica,sans-serif;font-size:17px;line-height:1.35}
header.page{position:sticky;top:0;z-index:10;background:#fff;border-bottom:3px solid var(--blue);
padding:10px 20px;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
header.page h1{font-size:18px;margin:0;color:var(--blue)}
header.page nav{display:flex;gap:6px;flex-wrap:wrap}
header.page nav a{font-size:13px;text-decoration:none;color:var(--grey);
border:1px solid var(--line);border-radius:4px;padding:3px 7px}
header.page nav a:hover{background:var(--blue);color:#fff;border-color:var(--blue)}
main{max-width:1240px;margin:22px auto 60px;padding:0 18px}
.slide{position:relative;background:var(--bg);border:1px solid var(--line);border-radius:6px;
margin:0 0 26px;padding:34px 44px 30px;aspect-ratio:16/9;overflow:hidden;
display:flex;flex-direction:column}
.slide h2{color:var(--blue);font-size:29px;margin:0 0 4px;padding-left:16px;
border-left:5px solid var(--blue)}
.slide .num{position:absolute;right:16px;bottom:10px;font-size:16px;color:var(--grey)}
.kicker{color:var(--orange);font-weight:bold;font-size:16px;margin:6px 0 4px}
.slide h1{color:var(--blue);font-size:44px;line-height:1.02;margin:6px 0 10px;max-width:74%}
.sub{font-size:21px;margin:0 0 14px}
.sub.accent{color:var(--orange)}
.slide--title{justify-content:flex-start}
.slide--title:after,.slide--final:before{content:"";position:absolute;left:0;right:0;top:0;
height:16px;background:linear-gradient(var(--blue) 70%,var(--orange) 70%)}
.meta{margin-top:auto;display:grid;grid-template-columns:170px 1fr;gap:2px 14px;font-size:16px}
.meta-row{display:contents}
.meta-row span{color:var(--grey)}
.meta-row b{font-weight:bold}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:26px;align-items:start;flex:1;margin-top:12px}
.cols .text-col,.cols .chart-col{min-width:0}
.chart-col img{width:100%;height:auto;display:block}
ul.bullets{list-style:none;padding:0;margin:0}
ul.bullets li{position:relative;padding-left:22px;margin:0 0 13px}
ul.bullets li:before{content:"▪";position:absolute;left:0;color:var(--orange);font-weight:bold}
table{border-collapse:collapse;width:100%;margin:14px 0 0;font-size:16px}
th{background:var(--blue);color:#fff;text-align:left;padding:8px 10px}
td{padding:7px 10px;border-bottom:1px solid #E6E6E6}
tbody tr:nth-child(odd){background:#F2F2F2}
.cards{display:grid;grid-template-columns:1fr 1fr;gap:14px;width:100%;align-content:start}
.card{background:#fff;border:1px solid var(--line);border-left:5px solid var(--blue);
padding:12px 16px;display:grid;grid-template-columns:auto 1fr;gap:2px 14px;align-items:center}
.card:nth-child(even){border-left-color:var(--orange)}
.cnum{font-size:30px;font-weight:bold;color:var(--blue-l2);grid-row:span 2}
.card h3{margin:0;color:var(--blue);font-size:19px}
.card p{margin:0;color:var(--grey);font-size:16px}
.slide h2.big{font-size:40px;border:0;padding:0;margin-top:18px}
.final-grid{display:grid;grid-template-columns:1fr 1.25fr;gap:34px;margin-top:10px}
.final-grid h3{font-size:16px;color:var(--grey);text-transform:uppercase;margin:0 0 10px}
.final-grid .col>div{font-size:17px;line-height:1.5}
.qr-row{display:grid;grid-template-columns:150px 1fr;gap:18px;align-items:start}
.qr{width:150px;height:auto;border:1px solid var(--line);background:#fff;padding:4px}
ul.srclist{margin:0;padding-left:18px;font-size:16px}
ul.srclist li{margin-bottom:5px}
.link{margin-top:12px;font-size:16px}
.link a{color:var(--blue);font-weight:bold}
.notes{margin-top:auto;padding-top:10px;font-size:16px;color:var(--grey)}
.notes summary{cursor:pointer;color:var(--blue);font-weight:bold}
.notes p{margin:8px 0 0;color:var(--ink)}
@media(max-width:900px){
 .slide{aspect-ratio:auto;padding:22px 20px}
 .cols,.cards,.final-grid,.qr-row{grid-template-columns:1fr}
 .slide h1{font-size:30px;max-width:100%}
 .slide h2,.slide h2.big{font-size:23px}
 .meta{grid-template-columns:130px 1fr}
}
@media print{
 header.page{display:none}
 body{background:#fff}
 .slide{break-after:page;border:0;margin:0;border-radius:0}
 .notes{display:none}
}
"""


def main():
    slides = "\n".join(render_slide(s) for s in SLIDES)
    nav = "".join(f"<a href='#slide-{s['n']}'>{s['n']}</a>" for s in SLIDES)
    doc = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Будущее труда в автоматизированной экономике — презентация</title>
<meta name="description" content="Учебная презентация: как ИИ и роботы меняют рынок труда.
14 слайдов, практическое занятие №10.">
<style>{CSS}</style>
</head>
<body>
<header class="page">
<h1>Будущее труда в автоматизированной экономике</h1>
<nav>{nav}</nav>
</header>
<main>
{slides}
</main>
</body>
</html>
"""
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(doc)
    size = os.path.getsize(OUT) / 1024
    print(f"wrote {OUT} ({size:.0f} KB, {len(SLIDES)} slides)")


if __name__ == "__main__":
    main()
