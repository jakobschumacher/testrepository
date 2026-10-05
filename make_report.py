"""Erzeugt die Folie "Wo stehen wir heute?" (DART-Ziele AMR) aus einer Excel-Datei.

Aufruf: python make_report.py [input.xlsx] [output.pdf] [Datum]
"""
import os
import sys
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib.ticker import MultipleLocator, FuncFormatter
import openpyxl

INPUT = sys.argv[1] if len(sys.argv) > 1 else "data/input_updated.xlsx"
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/bericht_neu.pdf"
LOGO = "data/logo.png"
DATUM = sys.argv[3] if len(sys.argv) > 3 else date.today().strftime("%d.%m.%Y")

BLUE = "#4472C4"
ORANGE = "#ED7D31"
RKI_BLUE = "#0B5CA8"
GRID = "#D9D9D9"
TEXT = "#595959"
plt.rcParams["font.family"] = "Liberation Sans"

# Reihenfolge und Titel wie in der Vorlage; (kursiver Teil, normaler Teil)
CHARTS = {
    "3GCECOLI": ("E. coli", " Cephalosporine 3. Gen."),
    "MRSA": ("", "MRSA"),
    "CRKP": ("K. pneumoniae", " Carbapeneme"),
    "VRE": ("", "VRE"),
}
TARGET_YEAR = 2030


def de(x, nd):
    return f"{x:.{nd}f}".replace(".", ",")


def read_data(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    years = [h for h in header if isinstance(h, int)]
    target_col = len(header) - 1  # letzte Spalte: Zielwert 2030 (absolut)
    data = {}
    for r in rows[1:]:
        if r[0] is None:
            continue
        values = [r[header.index(y)] for y in years]
        data[r[0]] = (years, values, r[target_col])
    return data


def ytick_fmt(step):
    nd = 0 if step >= 1 else len(f"{step:.10f}".rstrip("0").split(".")[1])
    return FuncFormatter(lambda v, _: de(v, nd).rstrip("0").rstrip(",") if nd else f"{v:.0f}")


def axis_scale(vmax):
    # Achsenmaximum und Schrittweite ähnlich der Excel-Automatik
    for step in (0.05, 0.1, 0.2, 0.5, 1, 2, 5):
        n = int(round(vmax / step, 6)) + 1
        if n <= 8:
            return n * step, step
    return vmax * 1.1, None


def label_fmt(code, v):
    # Beschriftung wie in der Vorlage: 2 Nachkommastellen, bei CRKP gekürzt (0,2 / 0,196)
    if code == "CRKP":
        return de(v, 3).rstrip("0").rstrip(",")
    return de(v, 2)


def draw(ax, code, years, values, target):
    first_year, last_year = years[0], years[-1]
    ax.plot(years, values, color=BLUE, lw=2.2, marker="o", ms=5, zorder=3,
            label=f"{first_year} - {last_year}")
    ax.plot([first_year, TARGET_YEAR], [values[0], target], color=ORANGE, lw=0,
            marker="o", ms=5, zorder=4)
    # Legendeneintrag wie in der Vorlage (Linie mit Marker)
    ax.plot([], [], color=ORANGE, lw=2.2, marker="o", ms=5, label="start - ziel")
    ax.plot([first_year, TARGET_YEAR], [values[0], target], color=ORANGE, lw=1,
            ls=(0, (1, 1.5)), zorder=2, label="Linear (start - ziel)")

    ymax, step = axis_scale(max(values + [target]))
    ax.set_ylim(0, ymax)
    ax.yaxis.set_major_locator(MultipleLocator(step))
    ax.yaxis.set_major_formatter(ytick_fmt(step))
    ax.set_xlim(first_year - 0.5, TARGET_YEAR + 0.5)
    ax.set_xticks(range(first_year, TARGET_YEAR + 1))
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.tick_params(colors=TEXT, labelsize=7, length=0)

    off = ymax * 0.04
    ax.text(first_year, values[0] + off, label_fmt(code, values[0]), fontsize=7,
            color=TEXT, ha="center", va="bottom")
    ax.text(TARGET_YEAR, target + off, label_fmt(code, target), fontsize=7,
            color=TEXT, ha="center", va="bottom")

    ax.text(0.0, 1.06, "Inzidenz", transform=ax.transAxes, fontsize=7.5,
            color=TEXT, ha="right", va="bottom")
    ital, normal = CHARTS[code]
    kw = dict(transform=ax.transAxes, fontsize=9, fontweight="bold", color="#404040", va="bottom")
    if ital:
        # kursiver und normaler Teil nebeneinander, gemeinsam zentriert
        fig = ax.figure
        r = fig.canvas.get_renderer()
        t1 = ax.text(0, 1.1, ital, style="italic", **kw)
        t2 = ax.text(0, 1.1, normal, **kw)
        w1 = t1.get_window_extent(r).width
        w2 = t2.get_window_extent(r).width
        ax_w = ax.get_window_extent(r).width
        x0 = 0.5 - (w1 + w2) / 2 / ax_w
        t1.set_x(x0)
        t2.set_x(x0 + w1 / ax_w)
    else:
        ax.text(0.5, 1.1, normal, ha="center", **kw)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=3,
              frameon=False, fontsize=7, labelcolor=TEXT, handlelength=3)


def main():
    data = read_data(INPUT)
    fig = plt.figure(figsize=(11.69, 8.27))  # A4 quer

    fig.text(0.05, 0.795, "Wo stehen wir heute?", fontsize=24, fontweight="bold",
             color=RKI_BLUE, va="bottom")
    # Logo nur, wenn lokal vorhanden (nicht im Repository enthalten)
    if os.path.exists(LOGO):
        lax = fig.add_axes([0.77, 0.79, 0.19, 0.07])
        lax.imshow(mpimg.imread(LOGO))
        lax.axis("off")

    boxes = [(0.052, 0.48), (0.43, 0.48), (0.052, 0.165), (0.43, 0.165)]
    for (x, y), code in zip(boxes, CHARTS):
        frame = fig.add_axes([x, y, 0.36, 0.30])
        frame.set_xticks([]); frame.set_yticks([])
        for s in frame.spines.values():
            s.set_color(GRID)
        frame.set_facecolor("none")
        ax = fig.add_axes([x + 0.035, y + 0.075, 0.31, 0.18])
        years, values, target = data[code]
        draw(ax, code, years, values, target)

    # Fußzeile
    for xl in (0.052, 0.287, 0.877, 0.94):
        fig.add_artist(plt.Line2D([xl, xl], [0.105, 0.155], color=RKI_BLUE, lw=0.8))
    fig.text(0.062, 0.14, DATUM, fontsize=8, color=RKI_BLUE)
    fig.text(0.297, 0.14, "DART Ziele Antibiotikar-Resistenz", fontsize=8, color=RKI_BLUE)
    fig.text(0.905, 0.14, "5", fontsize=8, color=RKI_BLUE)

    fig.savefig(OUTPUT)
    print("geschrieben:", OUTPUT)


if __name__ == "__main__":
    main()
