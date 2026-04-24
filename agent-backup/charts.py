"""Clustered horizontal bar charts (Budget vs Actual per region). Returns PNG bytes.

Used for the 3 exec-summary charts: MTD, QTD, YTD.
"""
from io import BytesIO

import matplotlib
matplotlib.use("Agg")  # no display
import matplotlib.pyplot as plt
import numpy as np

from agent import config


def _format_millions(x: float) -> str:
    """Format as $X.XM (e.g., $12.3M, $-1.2M, $0.4M)."""
    m = x / 1_000_000
    if abs(m) >= 10:
        return f"${m:,.1f}M"
    return f"${m:,.1f}M"


def clustered_bar_png(
    title: str,
    regions: list[str],
    budget_values: list[float],
    actual_values: list[float],
) -> bytes:
    """Render a clustered horizontal bar chart. Returns PNG bytes."""
    if len(regions) != len(budget_values) or len(regions) != len(actual_values):
        raise ValueError("regions, budget_values, actual_values must be same length")

    fig, ax = plt.subplots(figsize=config.CHART_FIGSIZE, dpi=config.CHART_DPI)

    # Order top-to-bottom: reverse for matplotlib (which stacks bottom-up)
    y = np.arange(len(regions))
    bar_height = 0.38
    budget_bars = ax.barh(
        y + bar_height / 2, budget_values, height=bar_height,
        color=config.CHART_BUDGET_COLOR, edgecolor="#999999",
        linewidth=0.5, label="Budget",
    )
    actual_bars = ax.barh(
        y - bar_height / 2, actual_values, height=bar_height,
        color=config.CHART_ACTUAL_COLOR, label="Actual",
    )

    # Data labels at end of bars
    for bars, vals in [(budget_bars, budget_values), (actual_bars, actual_values)]:
        labels = [_format_millions(v) for v in vals]
        ax.bar_label(bars, labels=labels, padding=4, fontsize=9)

    ax.set_yticks(y)
    ax.set_yticklabels(regions)
    ax.invert_yaxis()  # NA at top
    ax.set_xlabel("Total Bookings (CAD)")
    ax.set_title(title, fontsize=11, fontweight="bold")
    ax.legend(loc="lower right", frameon=False)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.xaxis.set_major_formatter(
        plt.FuncFormatter(lambda x, _pos: f"${x/1_000_000:,.0f}M")
    )

    max_val = max(max(budget_values, default=0), max(actual_values, default=0), 1)
    ax.set_xlim(right=max_val * 1.15)  # headroom for labels

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=config.CHART_DPI, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()
