"""Generate all chart / infographic assets for the presentation.

Two rules drive the sizing:

1. Charts are generated at the exact size they occupy on the slide, so the
   PPTX never has to scale them. Scaling a bitmap down also scales its text,
   which would break the "body text >= 16 pt" requirement.
2. Every label is placed inside the axes limits, so nothing is clipped and the
   full canvas is used (the brief asks for 60-70% slide coverage).

Style: light background, dark text, neutral palette plus two accents.
Blue = AI / cognitive automation. Orange = robots / physical automation.
"""

import os

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import qrcode
from matplotlib.patches import FancyBboxPatch

ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "charts")
os.makedirs(ASSETS, exist_ok=True)

BLUE = "#1F4E79"
BLUE_L = "#8FB8DC"
ORANGE = "#C55A11"
ORANGE_L = "#E8B18C"
GREEN = "#548235"
GREY = "#595959"
INK = "#1A1A1A"
BG = "#FFFFFF"
CARD = "#F2F2F2"

DPI = 200
plt.rcParams.update(
    {
        "font.family": "Liberation Sans",
        "font.size": 16,
        "text.color": INK,
        "axes.labelcolor": INK,
        "xtick.color": GREY,
        "ytick.color": GREY,
        "axes.edgecolor": "#BFBFBF",
        "axes.linewidth": 0.8,
        "figure.facecolor": BG,
        "axes.facecolor": BG,
        "savefig.facecolor": BG,
    }
)


def fig_inches(w_in, h_in):
    """A figure whose saved canvas is exactly w_in x h_in at 100% zoom."""
    return plt.figure(figsize=(w_in, h_in), dpi=DPI)


def finish(fig, name, rect=None):
    """Save the figure. `rect` is a [l, b, w, h] axes rectangle in figure
    fractions, letting a figure use its whole canvas without tight bbox."""
    path = os.path.join(ASSETS, name)
    if rect is None:
        fig.tight_layout(pad=0.4)
        fig.savefig(path, dpi=DPI, bbox_inches="tight", pad_inches=0.06)
    else:
        fig.savefig(path, dpi=DPI)
    plt.close(fig)
    print("wrote", os.path.basename(path))


def clean_axes(ax, grid_axis="x"):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.grid(axis=grid_axis, color="#E6E6E6", linewidth=0.9, zorder=0)
    ax.set_axisbelow(True)


# ---------------------------------------------------------------- slide 3
def chart_vectors():
    fig = fig_inches(6.23, 2.35)
    ax = fig.add_axes([0.05, 0.42, 0.90, 0.42])
    ai, rob, human = 44, 13, 43
    ax.barh([""], [ai], color=BLUE, height=0.55)
    ax.barh([""], [rob], left=[ai], color=ORANGE, height=0.55)
    ax.barh([""], [human], left=[ai + rob], color="#D9D9D9", height=0.55)
    for x, val, col in ((ai / 2, ai, "white"),
                        (ai + rob / 2, rob, "white"),
                        (ai + rob + human / 2, human, INK)):
        ax.text(x, 0, f"{val}%", ha="center", va="center", fontsize=19,
                fontweight="bold", color=col)
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, 0.5)
    ax.set_xticks([])
    ax.set_yticks([])
    for side in ("top", "right", "left", "bottom"):
        ax.spines[side].set_visible(False)
    handles = [
        plt.Rectangle((0, 0), 1, 1, color=BLUE),
        plt.Rectangle((0, 0), 1, 1, color=ORANGE),
        plt.Rectangle((0, 0), 1, 1, color="#D9D9D9"),
    ]
    ax.legend(handles, ["ИИ (цифровая)", "Роботы (физическая)", "Остаётся человеку"],
              loc="upper center", bbox_to_anchor=(0.5, 1.42), ncol=3,
              frameon=False, fontsize=16)
    finish(fig, "03_vectors_bar.png", rect=True)


# ---------------------------------------------------------------- slide 4
def chart_ai_exposure():
    roles = ["Программисты", "Офисный персонал", "Юристы", "Бизнес и финансы", "Дизайн и медиа"]
    vals = [94, 90, 85, 80, 75]
    fig = fig_inches(6.93, 5.25)
    ax = fig.add_axes([0.30, 0.07, 0.68, 0.90])
    bars = ax.barh(roles[::-1], vals[::-1], color=BLUE, height=0.62)
    for bar, val in zip(bars, vals[::-1]):
        ax.text(bar.get_width() + 2, bar.get_y() + bar.get_height() / 2, f"{val}%",
                va="center", fontsize=17, fontweight="bold", color=BLUE)
    ax.set_xlim(0, 112)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"])
    ax.tick_params(axis="y", labelsize=17, length=0)
    clean_axes(ax)
    finish(fig, "04_ai_exposure.png", rect=True)


# ---------------------------------------------------------------- slide 5
def chart_robot_exposure():
    roles = ["Водители такси", "Операторы сельхозтехники", "Водители лёгких грузовиков",
             "Водители шаттлов", "Операторы погрузчиков"]
    vals = [2.2, 2.1, 2.1, 2.0, 2.0]
    shades = ["#E8A06E", "#E08E55", "#D77C3C", "#CE6A22", ORANGE]
    fig = fig_inches(6.93, 5.25)
    ax = fig.add_axes([0.30, 0.13, 0.68, 0.84])
    bars = ax.barh(roles[::-1], vals[::-1], color=shades, height=0.62)
    for bar, val in zip(bars, vals[::-1]):
        ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2, f"{val:.1f}",
                va="center", fontsize=17, fontweight="bold", color=ORANGE)
    ax.set_xlim(0, 2.7)
    ax.set_xticks([0, 0.5, 1.0, 1.5, 2.0, 2.5])
    ax.set_xlabel("Индекс экспозиции (0–3)", fontsize=16)
    ax.tick_params(axis="y", labelsize=17, length=0)
    clean_axes(ax)
    finish(fig, "05_robot_exposure.png", rect=True)


# ---------------------------------------------------------------- slide 6
def chart_profiles():
    """Symmetric paired portrait cards."""
    fig = fig_inches(12.09, 4.15)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 3.43)
    ax.axis("off")

    cards = [
        (0.10, BLUE, "Уязвимые для ИИ",
         [("Пол", "сбалансировано"), ("Образование", "высокое"),
          ("Доход", "высокий"), ("Безработица", "ниже среднего")]),
        (5.10, ORANGE, "Уязвимые для роботов",
         [("Пол", "69% мужчины"), ("Образование", "8% с высшим"),
          ("Доход", "≈ $23/час"), ("Безработица", "5,2%")]),
    ]
    for x, color, title, rows in cards:
        ax.add_patch(FancyBboxPatch((x, 0.06), 4.80, 3.31,
                                    boxstyle="round,pad=0.03,rounding_size=0.12",
                                    linewidth=0, facecolor=CARD))
        ax.add_patch(FancyBboxPatch((x, 2.80), 4.80, 0.57,
                                    boxstyle="round,pad=0.03,rounding_size=0.12",
                                    linewidth=0, facecolor=color))
        ax.text(x + 2.40, 3.08, title, ha="center", va="center", fontsize=19,
                fontweight="bold", color="white")
        for i, (k, v) in enumerate(rows):
            y = 2.44 - i * 0.63
            ax.text(x + 0.28, y, k, ha="left", va="center", fontsize=17, color=GREY)
            ax.text(x + 4.52, y, v, ha="right", va="center", fontsize=17,
                    fontweight="bold", color=INK)
            if i < len(rows) - 1:
                ax.plot([x + 0.28, x + 4.52], [y - 0.315, y - 0.315],
                        color="#D0D0D0", linewidth=0.9)
    finish(fig, "06_profiles.png", rect=True)


# ---------------------------------------------------------------- slide 7
def chart_jobs():
    labels = ["Исчезнет\nк 2030", "Появится\nновых", "Чистый\nприрост", "Вклад\nроботов"]
    vals = [92, 170, 78, -5]
    colors = ["#9DC3E6", BLUE, GREEN, ORANGE]
    fig = fig_inches(6.78, 4.80)
    ax = fig.add_axes([0.10, 0.12, 0.88, 0.84])
    bars = ax.bar(labels, vals, color=colors, width=0.6)
    texts = ["92 млн", "170 млн", "+78 млн", "−5 млн"]
    for bar, val, txt in zip(bars, vals, texts):
        if val >= 0:
            ax.text(bar.get_x() + bar.get_width() / 2, val + 6, txt,
                    ha="center", va="bottom", fontsize=17, fontweight="bold", color=INK)
        else:
            ax.text(bar.get_x() + bar.get_width() / 2, val - 6, txt,
                    ha="center", va="top", fontsize=17, fontweight="bold", color=ORANGE)
    ax.axhline(0, color="#BFBFBF", linewidth=0.9)
    ax.set_ylim(-34, 200)
    ax.set_yticks([0, 50, 100, 150])
    ax.tick_params(axis="x", labelsize=16, length=0)
    clean_axes(ax, grid_axis="y")
    finish(fig, "07_jobs.png", rect=True)


# ---------------------------------------------------------------- slide 8
def chart_robot_density():
    rows = [("Южная Корея", 1220), ("Сингапур", 818), ("Германия", 449), ("Япония", 446),
            ("США", 307), ("Китай", 166), ("Мировой средний", 132)][::-1]
    names = [r[0] for r in rows]
    vals = [r[1] for r in rows]
    colors = [GREY if n == "Мировой средний" else BLUE for n in names]
    fig = fig_inches(7.28, 5.20)
    ax = fig.add_axes([0.27, 0.13, 0.71, 0.84])
    bars = ax.barh(names, vals, color=colors, height=0.62)
    for bar, val in zip(bars, vals):
        ax.text(bar.get_width() + 22, bar.get_y() + bar.get_height() / 2, f"{val}",
                va="center", fontsize=17, fontweight="bold", color=INK)
    ax.set_xlim(0, 1480)
    ax.set_xticks([0, 300, 600, 900, 1200])
    ax.set_xlabel("Роботов на 10 000 работников", fontsize=16)
    ax.tick_params(axis="y", labelsize=17, length=0)
    clean_axes(ax)
    finish(fig, "08_robot_density.png", rect=True)


# ---------------------------------------------------------------- slide 9
def chart_trucking():
    years = ["2020", "2023", "2026\n(прогноз)", "2030\n(прогноз)"]
    vals = [100, 94, 85, 70]
    fig = fig_inches(6.68, 4.80)
    ax = fig.add_axes([0.11, 0.13, 0.87, 0.83])
    ax.axvspan(1.5, 3.45, color="#FBE5D6", zorder=0)
    ax.plot(years, vals, marker="o", markersize=11, linewidth=3.2, color=BLUE, zorder=3)
    ax.plot(years[1:], vals[1:], marker="o", markersize=11, linewidth=3.2,
            linestyle="--", color=ORANGE_L, zorder=4)
    for i, val in enumerate(vals):
        ax.annotate(str(val), (i, val), textcoords="offset points", xytext=(0, 15),
                    ha="center", fontsize=18, fontweight="bold", color=INK)
    ax.text(2.5, 66, "зона прогноза", ha="center", fontsize=16, color=ORANGE)
    ax.set_ylim(60, 112)
    ax.set_yticks([60, 70, 80, 90, 100, 110])
    ax.tick_params(axis="x", labelsize=16, length=0)
    clean_axes(ax, grid_axis="y")
    finish(fig, "09_trucking.png", rect=True)


# ---------------------------------------------------------------- slide 10
def chart_logistics():
    fig = fig_inches(10.23, 3.45)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 3.37)
    ax.axis("off")
    pairs = [
        ("Сборщик заказов", "Техник по робототехнике"),
        ("Упаковщик", "Оператор систем управления"),
        ("Кладовщик", "Аналитик цепочек поставок"),
    ]
    ax.text(1.55, 3.14, "БЫЛО", ha="center", fontsize=17, fontweight="bold", color=GREY)
    ax.text(8.45, 3.14, "СТАЛО", ha="center", fontsize=17, fontweight="bold", color=BLUE)
    for i, (old, new) in enumerate(pairs):
        y = 2.48 - i * 0.90
        ax.add_patch(FancyBboxPatch((0.05, y - 0.30), 2.95, 0.60,
                                    boxstyle="round,pad=0.03,rounding_size=0.08",
                                    linewidth=0, facecolor=CARD))
        ax.text(1.525, y, old, ha="center", va="center", fontsize=16, color=GREY)
        ax.annotate("", xy=(6.93, y), xytext=(3.12, y),
                    arrowprops=dict(arrowstyle="-|>", color=ORANGE, linewidth=2.4))
        ax.add_patch(FancyBboxPatch((7.00, y - 0.30), 2.95, 0.60,
                                    boxstyle="round,pad=0.03,rounding_size=0.08",
                                    linewidth=0, facecolor=BLUE))
        ax.text(8.475, y, new, ha="center", va="center", fontsize=16, color="white")
    finish(fig, "10_logistics.png", rect=True)


# ---------------------------------------------------------------- slide 11
def chart_economics():
    fig = fig_inches(7.23, 4.55)
    ax_l = fig.add_axes([0.13, 0.14, 0.33, 0.76])
    ax_r = fig.add_axes([0.52, 0.04, 0.47, 0.92])

    cats = ["Технически\nспособны", "Экономически\nзаменены"]
    bars = ax_l.bar(cats, [74, 0.3], color=[BLUE_L, ORANGE], width=0.55)
    ax_l.text(0, 76, "74%", ha="center", fontsize=19, fontweight="bold", color=BLUE)
    ax_l.text(1, 4, "0,3%", ha="center", fontsize=19, fontweight="bold", color=ORANGE)
    ax_l.set_ylim(0, 96)
    ax_l.set_yticks([0, 25, 50, 75])
    ax_l.set_yticklabels(["0%", "25%", "50%", "75%"])
    ax_l.set_title("Физические задачи", fontsize=16, color=GREY, pad=8)
    ax_l.tick_params(axis="x", labelsize=16, length=0)
    clean_axes(ax_l, grid_axis="y")

    ax_r.set_title("Цена и доступность", fontsize=16, color=GREY, pad=8)
    ax_r.set_xlim(0, 10)
    ax_r.set_ylim(0, 4.3)
    ax_r.axis("off")
    rows = [
        ("Человек", "$30–45/час", "8 ч/день", GREY),
        ("Робот (пром.)", "€10–14/час", "16+ ч/день", ORANGE),
        ("Гуманоид (прогноз)", "< $20 тыс.", "24/7", BLUE),
    ]
    for i, (name, rate, hours, color) in enumerate(rows):
        y = 3.25 - i * 1.18
        ax_r.add_patch(FancyBboxPatch((0.05, y - 0.42), 9.85, 0.90,
                                      boxstyle="round,pad=0.03,rounding_size=0.10",
                                      linewidth=0, facecolor="#F7F7F7"))
        ax_r.add_patch(FancyBboxPatch((0.05, y - 0.42), 0.16, 0.90,
                                      boxstyle="round,pad=0,rounding_size=0.04",
                                      linewidth=0, facecolor=color))
        ax_r.text(0.48, y, name, ha="left", va="center", fontsize=16,
                  fontweight="bold", color=INK)
        ax_r.text(6.35, y, rate, ha="center", va="center", fontsize=16, color=INK)
        ax_r.text(9.72, y, hours, ha="right", va="center", fontsize=16, color=GREY)
    finish(fig, "11_economics.png", rect=True)


# ---------------------------------------------------------------- slide 12
def chart_skills():
    skills = ["Аналитическое\nмышление", "Эмоциональный\nинтеллект", "Креативность", "AI-грамотность"]
    series = {"2020": [70, 40, 45, 10], "2025": [85, 60, 65, 50], "2030": [95, 80, 85, 90]}
    fig = fig_inches(6.68, 4.55)
    ax = fig.add_axes([0.12, 0.14, 0.86, 0.70])
    n = len(skills)
    width = 0.26
    colors = ["#BDD7EE", BLUE, "#12395C"]
    for j, (year, vals) in enumerate(series.items()):
        xs = [i + (j - 1) * width for i in range(n)]
        bars = ax.bar(xs, vals, width=width, color=colors[j], label=year)
        for bar, val in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width() / 2, val + 2.0, str(val),
                    ha="center", fontsize=16, fontweight="bold", color=INK)
    ax.set_xticks(range(n))
    ax.set_xticklabels(skills, fontsize=16)
    ax.set_ylim(0, 124)
    ax.set_yticks([0, 25, 50, 75, 100])
    ax.set_yticklabels(["0%", "25%", "50%", "75%", "100%"])
    ax.legend(frameon=False, fontsize=16, ncol=3, loc="upper center",
              bbox_to_anchor=(0.5, 1.30))
    clean_axes(ax, grid_axis="y")
    finish(fig, "12_skills.png", rect=True)


# ---------------------------------------------------------------- slide 13
def chart_conclusion():
    fig = fig_inches(4.75, 2.30)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.8)
    ax.axis("off")
    items = [("ИИ", BLUE), ("Робот", ORANGE), ("Человек", GREEN)]
    for i, (label, color) in enumerate(items):
        x = 0.05 + i * 2.45
        ax.add_patch(FancyBboxPatch((x, 1.15), 1.95, 2.5,
                                    boxstyle="round,pad=0.04,rounding_size=0.14",
                                    linewidth=0, facecolor=color))
        ax.text(x + 0.975, 2.4, label, ha="center", va="center", fontsize=18,
                fontweight="bold", color="white")
        ax.text(x + 2.18, 2.4, "+", ha="center", va="center", fontsize=24, color=GREY)
    ax.text(7.60, 2.4, "=", ha="center", va="center", fontsize=24, color=GREY)
    ax.add_patch(FancyBboxPatch((7.72, 1.02), 2.23, 2.76,
                                boxstyle="round,pad=0.04,rounding_size=0.14",
                                linewidth=0, facecolor=CARD))
    ax.text(8.835, 2.4, "Новая\nэкономика", ha="center", va="center", fontsize=17,
            fontweight="bold", color=INK, linespacing=1.25)
    finish(fig, "13_conclusion.png", rect=True)


# ---------------------------------------------------------------- slide 14
def qr_code():
    url = "https://github.com/t-1mka/PresV2/blob/main/SOURCES.md"
    qrcode.make(url, box_size=12, border=2).save(
        os.path.join(ASSETS, "14_qr_sources.png"))
    print("wrote 14_qr_sources.png ->", url)


def logo_placeholder():
    """Monogram used on the title slide until a real logo is supplied."""
    fig = fig_inches(1.35, 1.35)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 2)
    ax.set_ylim(0, 2)
    ax.axis("off")
    ax.add_patch(FancyBboxPatch((0.06, 0.06), 1.88, 1.88,
                                boxstyle="round,pad=0.02,rounding_size=0.20",
                                linewidth=0, facecolor=BLUE))
    ax.text(1.0, 1.0, "АТ", ha="center", va="center", fontsize=42,
            fontweight="bold", color="white")
    finish(fig, "01_logo.png", rect=True)


if __name__ == "__main__":
    chart_vectors()
    chart_ai_exposure()
    chart_robot_exposure()
    chart_profiles()
    chart_jobs()
    chart_robot_density()
    chart_trucking()
    chart_logistics()
    chart_economics()
    chart_skills()
    chart_conclusion()
    qr_code()
    logo_placeholder()
    print("\nAll assets generated in", ASSETS)
