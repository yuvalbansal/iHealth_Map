# utils/ppt_builder.py

import io
import numpy as np
import pandas as pd

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches

import plotly.express as px

# -------------------------------------------------
# Small helpers copied exactly from legacy app
# -------------------------------------------------

def as_num(s):
    return pd.to_numeric(s, errors="coerce")


def safe_pct(num, den):
    return float(num) * 100.0 / float(den) if den not in (0, None) else 0.0


def set_autofit(tf, maxsize=28, minsize=14):
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass
    try:
        tf.fit_text(max_size=maxsize, min_size=minsize)
    except Exception:
        pass


def build_population_ppt(
    view_df: pd.DataFrame,
    df_full: pd.DataFrame,
    cols: dict,
    location_title: str,
    year_title: str,
) -> io.BytesIO:
    """
    IDENTICAL behavior to legacy single-file Downloads page.
    """

    prs = Presentation()

    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]
    try:
        title_only_layout = prs.slide_layouts[5]
    except IndexError:
        title_only_layout = bullet_layout

    # -------------------------------------------------
    # Slide helpers
    # -------------------------------------------------

    def add_bullet_slide(title_text, lines):
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = title_text
        set_autofit(slide.shapes.title.text_frame, 32, 18)

        tf = slide.placeholders[1].text_frame
        tf.clear()

        if not isinstance(lines, (list, tuple)):
            lines = [lines]

        for i, txt in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = str(txt)
            p.level = 0

        set_autofit(tf, 26, 12)

    def add_figure_slide(title_text, fig):
        slide = prs.slides.add_slide(title_only_layout)
        slide.shapes.title.text = title_text
        set_autofit(slide.shapes.title.text_frame, 28, 14)

        left, top, width = Inches(0.5), Inches(1.5), Inches(9)

        try:
            img = fig.to_image(format="png", width=1200, height=700)
            slide.shapes.add_picture(io.BytesIO(img), left, top, width=width)
        except Exception:
            box = slide.shapes.add_textbox(left, top, width, Inches(2))
            tf = box.text_frame
            tf.text = "Plot export failed. Install 'kaleido'."
            set_autofit(tf, 20, 12)

    # -------------------------------------------------
    # Common metrics
    # -------------------------------------------------

    total_screened = len(view_df)
    total_dataset = len(df_full)

    if "Health Status" in view_df.columns:
        normal = (view_df["Health Status"] == "Healthy").sum()
    else:
        normal = 0

    abnormal = total_screened - normal

    # -------------------------------------------------
    # Slide 1: Title
    # -------------------------------------------------

    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = f"Health of {location_title} – {year_title}".strip(" –")
    slide.placeholders[1].text = "Platform: Bharat_iHealthMap"

    set_autofit(slide.shapes.title.text_frame, 40, 24)
    set_autofit(slide.placeholders[1].text_frame, 24, 14)

    # -------------------------------------------------
    # Slide 2: Overall summary
    # -------------------------------------------------

    add_bullet_slide(
        "Overall Screening Summary",
        [
            f"Total persons in dataset: {total_dataset:,}",
            f"Screened persons (after filters): {total_screened:,}",
            f"Total with abnormal health status: {abnormal:,} ({safe_pct(abnormal, total_screened):.1f}%)",
            f"Total with normal health status: {normal:,} ({safe_pct(normal, total_screened):.1f}%)",
        ],
    )

    # -------------------------------------------------
    # Gender participation
    # -------------------------------------------------

    if cols.get("gender") and cols["gender"] in view_df.columns:
        g = view_df[cols["gender"]].astype(str).value_counts()
        add_bullet_slide(
            "Gender-wise Participation",
            [f"{k}: {v:,} ({safe_pct(v, g.sum()):.1f}%)" for k, v in g.items()],
        )
    else:
        add_bullet_slide(
            "Gender-wise Participation",
            ["Gender column not available."],
        )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
