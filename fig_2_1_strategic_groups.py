"""
Рис. 2.1. Карта стратегічних груп ринку роздрібної торгівлі пальним.

Вісь X — ціна бензину А-95 (грн/л) станом на 22.05.2026 (Gazeta.ua).
Вісь Y — широта сервісу, бали 1–10 (оцінка автора: харчування, програма
лояльності, преміальні марки пального, електрозарядки).
Площа кола пропорційна кількості АЗС на початок 2026 р. (Факти ICTV).

Запуск:  python fig_2_1_strategic_groups.py   ->  fig_2_1_strategic_groups.png
Потрібно: matplotlib (pip install matplotlib)
"""
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse

# мережа: (ціна А-95, бал сервісу, кількість АЗС, група)
NETWORKS = {
    "ОККО":       (79.90, 9.5, 416, "I"),
    "WOG":        (79.90, 8.5, 359, "I"),
    "SOCAR":      (79.90, 7.0,  55, "I"),
    "UPG":        (77.40, 6.0, 201, "II"),
    "KLO":        (78.40, 5.0,  71, "II"),
    "Укрнафта":   (73.90, 5.5, 662, "III"),
    "БРСМ-Нафта": (71.99, 3.0, 227, "III"),
}
GROUPS = {
    # назва, центр еліпса, ширина, висота, позиція підпису, вирівнювання
    "I":   ("Преміум-сервіс",   (80.3, 8.2), 3.2, 4.1, (80.0, 10.6), "center"),
    "II":  ("Середній сегмент", (77.90, 5.50), 2.8, 2.4, (77.9, 3.95),  "center"),
    "III": ("Ціновий сегмент",  (72.90, 4.25), 4.4, 4.4, (72.9, 6.75),  "center"),
}
FOCUS = "ОККО"
SIZE_PER_STATION = 3.6          # площа маркера (pt²) на одну АЗС
FONT = "DejaVu Sans"            # кирилиця є в DejaVu Sans


def main(out="fig_2_1_strategic_groups.png"):
    plt.rcParams["font.family"] = FONT
    total = sum(v[2] for v in NETWORKS.values())

    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)

    # еліпси стратегічних груп і їх підписи
    for g, (name, centre, w, h, label_xy, ha) in GROUPS.items():
        ax.add_patch(Ellipse(centre, w, h, facecolor="#F3F3F3",
                             edgecolor="#AAAAAA", linestyle="--", zorder=1))
        n = sum(v[2] for v in NETWORKS.values() if v[3] == g)
        ax.text(*label_xy, f"{g}. {name} · {n} АЗС ({n / total:.0%})",
                ha=ha, va="center", fontsize=13, fontweight="bold")

    # кола мереж
    for name, (price, score, stations, _) in NETWORKS.items():
        focus = name == FOCUS
        ax.scatter(price, score, s=stations * SIZE_PER_STATION,
                   color="#3B82D6" if focus else "#DADAD2",
                   edgecolor="#1F5FAD" if focus else "#B5B5AA",
                   linewidth=1, zorder=3)
        dx = 0.22 + (stations ** 0.5) * 0.0105
        left = name == "Укрнафта"
        ax.text(price - dx if left else price + dx, score, name,
                ha="right" if left else "left", va="center", zorder=4,
                fontsize=14, fontweight="bold" if focus else "normal")

    ax.set_xlim(70.4, 82.4)
    ax.set_ylim(1.5, 11)
    ax.set_xticks([72, 74, 76, 78, 80])
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_xlabel("Ціна бензину А-95, грн/л (22.05.2026)", fontsize=14)
    ax.set_ylabel("Широта сервісу, бали", fontsize=14)
    ax.grid(axis="y", color="#E5E5E5")
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.tick_params(length=0, labelsize=13)

    fig.tight_layout()
    fig.savefig(out, dpi=200)
    print("saved", out)


if __name__ == "__main__":
    main()
