import io
import numpy as np
import pandas as pd

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches

import plotly.express as px


# ---------------------------------------------------------------------
# Helper utilities (same logic as Initial_Code.py)
# ---------------------------------------------------------------------

def as_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def safe_pct(num: float, den: float) -> float:
    return float(num) * 100.0 / float(den) if den not in (0, None) else 0.0


def set_autofit(tf, maxsize=28, minsize=12):
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass
    try:
        tf.fit_text(max_size=maxsize, min_size=minsize)
    except Exception:
        pass


def donut_normal_abnormal(series, label, thresholds):
    s = pd.to_numeric(series, errors="coerce")
    s = s[s.notna()]

    if len(s) == 0:
        df_plot = pd.DataFrame({"Category": ["No data"], "Count": [1]})
    else:
        low_ok, high_ok = thresholds
        if low_ok is not None and high_ok is not None:
            normal = s.between(low_ok, high_ok)
        elif low_ok is None and high_ok is not None:
            normal = s <= high_ok
        elif low_ok is not None and high_ok is None:
            normal = s >= low_ok
        else:
            normal = pd.Series(False, index=s.index)

        abnormal = ~normal
        df_plot = pd.DataFrame(
            {
                "Category": ["Normal", "Abnormal"],
                "Count": [int(normal.sum()), int(abnormal.sum())],
            }
        )

    fig = px.pie(
        df_plot,
        values="Count",
        names="Category",
        hole=0.55,
        title=label,
    )
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(showlegend=False, margin=dict(l=0, r=0, t=40, b=0))
    return fig


# ---------------------------------------------------------------------
# Main public API
# ---------------------------------------------------------------------

def build_population_ppt(
    view_df: pd.DataFrame,
    df_full: pd.DataFrame,
    cols_map: dict,
    location_title: str,
    year_title: str,
) -> io.BytesIO:
    """
    Combined A + B style PPT:
    - Keeps all text / bullet analytics
    - Adds image slides for all key Plotly charts across:
        Overview, Clinical, Lifestyle, Community, Socioeconomic & Profession
    """
    prs = Presentation()

    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]
    # title-only or content layout fallback for figures
    try:
        title_only_layout = prs.slide_layouts[5]
    except IndexError:
        title_only_layout = bullet_layout

    # ---------------- small helpers -----------------

    def set_autofit(tf, maxsize=28, minsize=12):
        tf.word_wrap = True
        try:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
        try:
            tf.fit_text(max_size=maxsize, min_size=minsize)
        except Exception:
            # Some pptx versions don’t have fit_text
            pass

    def add_bullet_slide(title_text: str, lines):
        slide = prs.slides.add_slide(bullet_layout)
        title_shape = slide.shapes.title
        title_shape.text = title_text
        set_autofit(title_shape.text_frame, maxsize=32, minsize=18)

        body = slide.placeholders[1].text_frame
        body.clear()
        if not isinstance(lines, (list, tuple)):
            lines = [lines]

        for i, txt in enumerate(lines):
            p = body.paragraphs[0] if i == 0 else body.add_paragraph()
            p.text = str(txt)
            p.level = 0
        set_autofit(body, maxsize=26, minsize=12)

    def add_figure_slide(title_text: str, fig):
        """
        Insert a Plotly fig as a PNG image slide.
        Requires `pip install kaleido` in the environment.
        """
        slide = prs.slides.add_slide(title_only_layout)
        title_shape = slide.shapes.title
        title_shape.text = title_text
        set_autofit(title_shape.text_frame, maxsize=28, minsize=14)

        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9.0)

        try:
            # This calls Plotly's built-in static export (needs kaleido)
            img_bytes = fig.to_image(format="png", width=1200, height=700)
            stream = io.BytesIO(img_bytes)
            slide.shapes.add_picture(stream, left, top, width=width)
        except Exception:
            # Graceful fallback if kaleido not installed
            box = slide.shapes.add_textbox(left, top, width, Inches(2.0))
            tf = box.text_frame
            tf.text = (
                "Chart image export failed. "
                "Install 'kaleido' (pip install kaleido) to enable Plotly → PNG export."
            )
            set_autofit(tf, maxsize=20, minsize=12)

    # ---------------- Common metrics -----------------

    total_screened = len(view_df)
    total_dataset = len(df_full)

    if "Health Status" in view_df.columns:
        normal_mask = view_df["Health Status"] == "Healthy"
        total_normal = int(normal_mask.sum())
        total_abnormal = int(total_screened - total_normal)
    else:
        total_normal = 0
        total_abnormal = 0

    pct_normal = safe_pct(total_normal, total_screened)
    pct_abnormal = safe_pct(total_abnormal, total_screened)

    # Convenience aliases
    cols = cols_map

    # ---------------- Slide 1: Title -----------------
    slide1 = prs.slides.add_slide(title_layout)
    t = slide1.shapes.title
    st_sub = slide1.placeholders[1]

    t.text = f"Health of {location_title} – {year_title}".strip(" –")
    st_sub.text = "Platform: Bharat_iHealthMap"

    set_autofit(t.text_frame, maxsize=40, minsize=24)
    set_autofit(st_sub.text_frame, maxsize=24, minsize=14)

    # ---------------- Slide 2: Overall Screening Summary --------
    add_bullet_slide(
        "Overall Screening Summary",
        [
            f"Total persons in dataset: {total_dataset:,}",
            f"Screened persons (after filters): {total_screened:,}",
            f"Total with abnormal health status: {total_abnormal:,} ({pct_abnormal:.1f}%)",
            f"Total with normal health status: {total_normal:,} ({pct_normal:.1f}%)",
        ],
    )

    # ---------------- Slide 3: Gender-wise Participation --------
    if cols["gender"]:
        g_series = view_df[cols["gender"]].astype(str)
        g_counts = g_series.value_counts()
        total_g = int(g_counts.sum())
        lines = [
            f"{g}: {cnt:,} ({safe_pct(cnt, total_g):.1f}%)"
            for g, cnt in g_counts.items()
        ]
    else:
        lines = ["Gender column not available in this dataset."]

    add_bullet_slide("Gender-wise Participation", lines)

    # -----------------------------------------------------------------
    # Parameter abnormality helpers (same as your earlier version)
    # -----------------------------------------------------------------
    glu_col = cols["glucose"]
    chol_col = cols["chol"]
    bp_sys = cols["bp_sys"]
    bp_dia = cols["bp_dia"]
    creat_col = cols["creatinine"]
    alt_col = cols["alt"]
    prot_col = cols["protein_total"]
    bmi_col = cols["bmi"]
    locality_col = cols["locality"]

    def abnormal_bp_mask(df_in):
        if not (bp_sys and bp_dia):
            return pd.Series([False] * len(df_in))
        s = as_num(df_in[bp_sys])
        d = as_num(df_in[bp_dia])
        return (s >= 130) | (d >= 80)

    def abnormal_glu_mask(df_in):
        if not glu_col:
            return pd.Series([False] * len(df_in))
        g = as_num(df_in[glu_col])
        return g >= 126

    def abnormal_chol_mask(df_in):
        if not chol_col:
            return pd.Series([False] * len(df_in))
        ch = as_num(df_in[chol_col])
        return ch >= 240

    def abnormal_creat_mask(df_in):
        if not creat_col:
            return pd.Series([False] * len(df_in))
        c = as_num(df_in[creat_col])
        return c > 1.2

    def abnormal_protein_mask(df_in):
        if not prot_col:
            return pd.Series([False] * len(df_in))
        p = as_num(df_in[prot_col])
        return (p < 6.0) | (p > 8.0)

    def abnormal_alt_mask(df_in):
        if not alt_col:
            return pd.Series([False] * len(df_in))
        a = as_num(df_in[alt_col])
        return a > 40

    # ---------------- Slides: Abnormality summaries ---------------
    def abnormal_summary(title_text, series, predicate):
        valid = series.dropna()
        if len(valid) == 0:
            add_bullet_slide(title_text, ["No data available."])
            return
        ab = predicate(valid)
        cnt_ab = int(ab.sum())
        cnt_norm = len(valid) - cnt_ab
        add_bullet_slide(
            title_text,
            [
                f"Abnormal: {cnt_ab:,} ({safe_pct(cnt_ab, len(valid)):.1f}%)",
                f"Normal: {cnt_norm:,} ({safe_pct(cnt_norm, len(valid)):.1f}%)",
            ],
        )

    if glu_col:
        abnormal_summary(
            "Diabetes (fasting glucose ≥126 mg/dL)",
            as_num(view_df[glu_col]),
            lambda x: x >= 126,
        )
    if bp_sys and bp_dia:
        valid_bp = view_df[[bp_sys, bp_dia]].dropna()
        if len(valid_bp) == 0:
            add_bullet_slide("Blood Pressure Status", ["No BP data available."])
        else:
            ab = abnormal_bp_mask(valid_bp)
            add_bullet_slide(
                "Blood Pressure Status",
                [
                    f"Hypertensive (>130/80): {ab.sum():,} ({safe_pct(ab.sum(), len(valid_bp)):.1f}%)",
                    f"Normotensive: {len(valid_bp) - ab.sum():,} ({safe_pct(len(valid_bp) - ab.sum(), len(valid_bp)):.1f}%)",
                ],
            )
    if chol_col:
        abnormal_summary(
            "Cholesterol (≥240 mg/dL)",
            as_num(view_df[chol_col]),
            lambda x: x >= 240,
        )
    if creat_col:
        abnormal_summary(
            "Renal (Creatinine >1.2 mg/dL)",
            as_num(view_df[creat_col]),
            lambda x: x > 1.2,
        )
    if prot_col:
        abnormal_summary(
            "Total Protein (outside 6.0–8.0 g/dL)",
            as_num(view_df[prot_col]),
            lambda x: (x < 6.0) | (x > 8.0),
        )
    if alt_col:
        abnormal_summary(
            "Liver Enzyme (ALT/SGPT >40 U/L)",
            as_num(view_df[alt_col]),
            lambda x: x > 40,
        )

    # ---------------- Age 18–20 yrs abnormality -------------------
    if "__AGE__" in view_df.columns:
        grp_1820 = view_df[view_df["__AGE__"].between(18, 20)]
        if len(grp_1820) == 0:
            add_bullet_slide(
                "Abnormal Parameters – Age 18–20 years",
                ["No records in the 18–20 years age group."],
            )
        else:
            bullets = []
            for label, fn in [
                ("Hypertension", abnormal_bp_mask),
                ("Diabetes (≥126 mg/dL)", abnormal_glu_mask),
                ("High cholesterol (≥240 mg/dL)", abnormal_chol_mask),
                ("Abnormal creatinine", abnormal_creat_mask),
                ("Abnormal total protein", abnormal_protein_mask),
                ("Abnormal SGPT", abnormal_alt_mask),
            ]:
                ab = fn(grp_1820)
                bullets.append(
                    f"{label}: {safe_pct(ab.sum(), len(grp_1820)):.1f}%"
                )
            add_bullet_slide(
                "Abnormal Parameters – 18–20 years", bullets
            )
    else:
        add_bullet_slide(
            "Abnormal Parameters – 18–20 years",
            ["Age column not available."],
        )

    # ---------------- Community-wise abnormal parameters ----------
    if locality_col:
        sub_loc = view_df.copy()
        any_ab = (
            abnormal_bp_mask(sub_loc)
            | abnormal_glu_mask(sub_loc)
            | abnormal_chol_mask(sub_loc)
            | abnormal_creat_mask(sub_loc)
            | abnormal_protein_mask(sub_loc)
            | abnormal_alt_mask(sub_loc)
        )
        sub_loc["__ANY_ABNORMAL__"] = any_ab.astype(int)
        grp = (
            sub_loc.groupby(locality_col)["__ANY_ABNORMAL__"]
            .agg(["count", "sum"])
            .reset_index()
        )
        grp["abnormal_pct"] = grp["sum"] * 100.0 / grp["count"]
        grp = grp.sort_values("count", ascending=False).head(10)

        bullets = []
        for _, row in grp.iterrows():
            name = str(row[locality_col])
            n_pop = int(row["count"])
            p_ab = float(row["abnormal_pct"])
            bullets.append(
                f"{name}: {p_ab:.1f}% with ≥1 abnormal parameter (n={n_pop})"
            )
        add_bullet_slide(
            "Abnormal Parameters – Community-wise", bullets
        )
    else:
        add_bullet_slide(
            "Abnormal Parameters – Community-wise",
            ["Community/locality column not available."],
        )

    # ---------------- BMI distribution (text) --------------------
    if bmi_col:
        b_ = as_num(view_df[bmi_col]).dropna()
        n_b = len(b_)
        if n_b == 0:
            add_bullet_slide("BMI Distribution", ["No BMI data available."])
        else:
            under = (b_ < 18.5).sum()
            normal_b = ((b_ >= 18.5) & (b_ < 25)).sum()
            over = ((b_ >= 25) & (b_ < 30)).sum()
            obese = (b_ >= 30).sum()

            add_bullet_slide(
                "BMI Distribution",
                [
                    f"Underweight (<18.5): {under:,} ({safe_pct(under, n_b):.1f}%)",
                    f"Normal (18.5–24.9): {normal_b:,} ({safe_pct(normal_b, n_b):.1f}%)",
                    f"Overweight (25–29.9): {over:,} ({safe_pct(over, n_b):.1f}%)",
                    f"Obese (≥30): {obese:,} ({safe_pct(obese, n_b):.1f}%)",
                    "Classification: <18.5 underweight; 18.5–24.9 healthy; "
                    "25–29.9 overweight; ≥30 obese.",
                ],
            )
    else:
        add_bullet_slide(
            "BMI Distribution", ["BMI column not available."]
        )

    # ---------------- BMI vs age bands (text) --------------------
    if bmi_col and "__AGE__" in view_df.columns:
        sub = view_df[[ "__AGE__", bmi_col]].dropna()
        if len(sub) > 0:
            bins = [0, 10, 20, 30, 40, 50, 60, 70, 200]
            labels = [
                "01–10",
                "11–20",
                "21–30",
                "31–40",
                "41–50",
                "51–60",
                "61–70",
                "71+",
            ]
            sub["age_band"] = pd.cut(sub["__AGE__"], bins=bins, labels=labels, right=True)
            b_ = as_num(sub[bmi_col])
            sub["over_obese"] = (b_ >= 25).astype(int)

            grp = (
                sub.groupby("age_band")["over_obese"]
                .agg(["count", "sum"])
                .reset_index()
            )
            grp["pct_over_obese"] = grp["sum"] * 100.0 / grp["count"]
            grp = grp.dropna(subset=["age_band"])

            lines = ["Overweight & obesity (BMI ≥25 kg/m²) by age band:"]
            for _, row in grp.iterrows():
                band = str(row["age_band"])
                pct_ = float(row["pct_over_obese"])
                lines.append(f"{band} years: {pct_:.1f}% overweight/obese")

            peak_row = grp.loc[grp["pct_over_obese"].idxmax()]
            peak_band = str(peak_row["age_band"])
            lines.append(
                f"Peak overweight/obesity seen in {peak_band} years group."
            )
            add_bullet_slide("BMI in Different Age Groups", lines)
        else:
            add_bullet_slide(
                "BMI in Different Age Groups",
                ["No BMI+age records available."],
            )
    else:
        add_bullet_slide(
            "BMI in Different Age Groups",
            ["BMI or age column not available."],
        )

    # ---------------- Rural vs Urban (Indore vs Mhow) ------------
    if locality_col:
        loc_series = view_df[locality_col].astype(str).str.lower()

        def rural_urban_tag(val: str) -> str:
            if "indore" in val:
                return "Urban (Indore)"
            if "mhow" in val:
                return "Rural (Mhow)"
            return "Other"

        df_ru = view_df.copy()
        df_ru["__RU_TAG__"] = loc_series.map(rural_urban_tag)

        metrics = []
        for tag_name in ["Rural (Mhow)", "Urban (Indore)"]:
            sub_tag = df_ru[df_ru["__RU_TAG__"] == tag_name]
            if len(sub_tag) == 0:
                continue
            glu_mean = as_num(sub_tag[glu_col]).mean() if glu_col else np.nan
            chol_mean = as_num(sub_tag[chol_col]).mean() if chol_col else np.nan
            alt_mean = as_num(sub_tag[alt_col]).mean() if alt_col else np.nan
            metrics.append((tag_name, len(sub_tag), glu_mean, chol_mean, alt_mean))

        if metrics:
            lines = ["Comparison between Rural (Mhow) and Urban (Indore):"]
            for tag_name, n_rec, g_m, c_m, a_m in metrics:
                text_line = f"{tag_name}: n={n_rec}"
                if not np.isnan(g_m):
                    text_line += f"; mean fasting glucose ≈ {g_m:.1f} mg/dL"
                if not np.isnan(c_m):
                    text_line += f"; mean cholesterol ≈ {c_m:.1f} mg/dL"
                if not np.isnan(a_m):
                    text_line += f"; mean SGPT ≈ {a_m:.1f} U/L"
                lines.append(text_line)
            add_bullet_slide("Rural vs Urban Comparison", lines)
        else:
            add_bullet_slide(
                "Rural vs Urban Comparison",
                [
                    "Rural (Mhow) vs Urban (Indore) tags could not be derived.",
                    "Ensure locality text contains 'Mhow' and 'Indore'.",
                ],
            )
    else:
        add_bullet_slide(
            "Rural vs Urban Comparison",
            ["Locality column not available; comparison not generated."],
        )

    # =================================================================
    #  NEW SECTION: ADD ONE SLIDE PER PLOTLY FIGURE (A + B COMBINED)
    # =================================================================

    # ---------------- OVERVIEW CHARTS ----------------
    # Status pie
    if "Health Status" in view_df.columns:
        # Safe extraction of status column
        status_col = None
        for cand in ["Health Status", "Health_Status", "health_status", "health status"]:
            if cand in view_df.columns:
                status_col = cand
                break

        if status_col:
            temp = (
                view_df[status_col]
                .astype(str)
                .value_counts()
                .reset_index()
            )
            temp.columns = ["Status", "Count"]   # force the correct names

            fig_status = px.pie(
                temp,
                names="Status",      # guaranteed to exist
                values="Count",      # guaranteed to exist
                title="Overall health status distribution",
            )
            add_figure_slide("Overview – Health Status (Pie)", fig_status)
        else:
            # Fallback slide
            add_bullet_slide(
                "Overview – Health Status (Pie)",
                ["No usable 'Health Status' column found in dataset."]
            )


    # Gender bar
    # --------- SAFE GENDER COLUMN DETECTION ----------
    gender_col = None

    # Possible names your dataset may use
    gender_candidates = [
        cols.get("gender"),         # primary from your mapping
        "Gender",
        "gender",
        "SEX",
        "Sex",
        "sex",
        "Male_Female",
        "M_F",
        "Gender ",
        "Gender (M/F)"
    ]

    for cand in gender_candidates:
        if cand and cand in view_df.columns:
            gender_col = cand
            break

    # --------- IF FOUND, PLOT GENDER BAR ----------
    if gender_col:
        temp = (
            view_df[gender_col]
            .astype(str)
            .value_counts()
            .reset_index()
        )
        temp.columns = ["Gender", "Count"]   # normalize names

        fig_gender = px.bar(
            temp,
            x="Gender",        # guaranteed to exist
            y="Count",
            title="Gender-wise participation",
            text="Count",
        )
        fig_gender.update_traces(textposition="outside")
        add_figure_slide("Overview – Gender-wise Participation", fig_gender)

    else:
        # Fallback safe slide
        add_bullet_slide(
            "Overview – Gender-wise Participation",
            ["No usable Gender column found in dataset."]
        )


    # Donut charts: glucose, cholesterol, BMI
    if glu_col:
        fig_g = donut_normal_abnormal(
            view_df[glu_col], "Fasting glucose", (70, 99)
        )
        add_figure_slide("Overview – Fasting Glucose (Normal vs Abnormal)", fig_g)
    if chol_col:
        fig_c = donut_normal_abnormal(
            view_df[chol_col], "Total cholesterol", (0, 199)
        )
        add_figure_slide("Overview – Cholesterol (Normal vs Abnormal)", fig_c)
    if bmi_col:
        b = as_num(view_df[bmi_col])
        series = pd.Series(
            np.where((b >= 18.5) & (b < 25), b, np.nan)
        )
        fig_b = donut_normal_abnormal(
            series, "BMI (18.5–24.9 normal)", (0, np.inf)
        )
        add_figure_slide("Overview – BMI (Normal vs Others)", fig_b)

    # ---------------- CLINICAL CHARTS ----------------
    # Glucose distribution + categories
    if glu_col:
        fig_glu_hist = px.histogram(
            view_df,
            x=glu_col,
            nbins=40,
            title="Fasting glucose distribution",
        )
        fig_glu_hist.update_layout(xaxis_title="Glucose (mg/dL)")
        add_figure_slide("Clinical – Fasting Glucose Distribution", fig_glu_hist)

        g = as_num(view_df[glu_col])
        cat = {
            "Normal (<100)": (g < 100).mean(),
            "Prediabetes (100–125)": ((g >= 100) & (g < 126)).mean(),
            "Diabetes (≥126)": (g >= 126).mean(),
        }
        df_cat = pd.DataFrame(
            {
                "Category": list(cat.keys()),
                "Share (%)": [v * 100 for v in cat.values()],
            }
        )
        fig_glu_cat = px.bar(
            df_cat,
            x="Category",
            y="Share (%)",
            title="Glycemic categories (approx.)",
            text="Share (%)",
        )
        fig_glu_cat.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig_glu_cat.update_layout(yaxis_title="Share (%)")
        add_figure_slide("Clinical – Glycemic Categories", fig_glu_cat)

    # Cholesterol distribution + categories
    if chol_col:
        fig_chol_hist = px.histogram(
            view_df,
            x=chol_col,
            nbins=40,
            title="Cholesterol distribution",
        )
        fig_chol_hist.update_layout(xaxis_title="Cholesterol (mg/dL)")
        add_figure_slide("Clinical – Cholesterol Distribution", fig_chol_hist)

        cchol = as_num(view_df[chol_col])
        catc = {
            "Desirable (<200)": (cchol < 200).mean(),
            "Borderline (200–239)": ((cchol >= 200) & (cchol < 240)).mean(),
            "High (≥240)": (cchol >= 240).mean(),
        }
        dfc = pd.DataFrame(
            {
                "Category": list(catc.keys()),
                "Share (%)": [v * 100 for v in catc.values()],
            }
        )
        fig_chol_cat = px.bar(
            dfc,
            x="Category",
            y="Share (%)",
            title="Cholesterol risk categories",
            text="Share (%)",
        )
        fig_chol_cat.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig_chol_cat.update_layout(yaxis_title="Share (%)")
        add_figure_slide("Clinical – Cholesterol Risk Categories", fig_chol_cat)

    # BP scatter + 3D density
    if bp_sys and bp_dia:
        bp_view = view_df[[bp_sys, bp_dia]].dropna()
        if len(bp_view) > 60000:
            bp_view = bp_view.sample(60000, random_state=42)

        fig_bp_scatter = px.scatter(
            bp_view,
            x=bp_sys,
            y=bp_dia,
            opacity=0.4,
            labels={bp_sys: "Systolic (mmHg)", bp_dia: "Diastolic (mmHg)"},
            title="BP scatter (sampled if very large)",
        )
        add_figure_slide("Clinical – BP Scatter (2D)", fig_bp_scatter)

        # 3D density (as in corrected plot)
        bp_df = view_df[[bp_sys, bp_dia]].dropna().astype(float)
        if len(bp_df) > 0:
            bp_counts = (
                bp_df.groupby([bp_sys, bp_dia])
                .size()
                .reset_index(name="Count")
            )
            fig_bp3d = px.scatter_3d(
                bp_counts,
                x=bp_sys,
                y=bp_dia,
                z="Count",
                color="Count",
                size="Count",
                opacity=0.9,
                title="3D BP Density (Systolic × Diastolic)",
            )
            fig_bp3d.update_layout(
                width=1000,
                height=700,
                margin=dict(l=0, r=0, t=80, b=20),
                scene=dict(
                    xaxis_title="Systolic (mmHg)",
                    yaxis_title="Diastolic (mmHg)",
                    zaxis_title="Number of persons",
                    aspectmode="cube",
                ),
            )
            add_figure_slide("Clinical – BP Density (3D)", fig_bp3d)

    # BMI histogram
    if bmi_col:
        fig_bmi_hist = px.histogram(
            view_df,
            x=bmi_col,
            nbins=40,
            title="BMI distribution",
        )
        fig_bmi_hist.update_layout(xaxis_title="BMI (kg/m²)")
        add_figure_slide("Clinical – BMI Distribution", fig_bmi_hist)

    # ---------------- LIFESTYLE CHARTS ----------------
    if cols["diet"]:
        d_series = view_df[cols["diet"]].map(parse_diet)
        d_counts = d_series.value_counts()
        fig_diet = px.pie(
            values=d_counts.values,
            names=d_counts.index,
            hole=0.55,
            title="Diet pattern",
        )
        add_figure_slide("Lifestyle – Diet Pattern", fig_diet)

    if cols["sleep"]:
        s_counts = view_df[cols["sleep"]].astype(str).value_counts().head(8)
        fig_sleep = px.bar(
            x=s_counts.index,
            y=s_counts.values,
            labels={"x": "Sleep pattern", "y": "Count"},
            title="Sleep pattern (top responses)",
            text=s_counts.values,
        )
        fig_sleep.update_traces(textposition="outside")
        add_figure_slide("Lifestyle – Sleep Pattern", fig_sleep)

    if cols["tobacco"]:
        tob = view_df[cols["tobacco"]].map(flag_yes).value_counts()
        fig_tob = px.bar(
            x=tob.index,
            y=tob.values,
            labels={"x": "Tobacco history", "y": "Count"},
            title="Tobacco / smoking history",
            text=tob.values,
        )
        fig_tob.update_traces(textposition="outside")
        add_figure_slide("Lifestyle – Tobacco / Smoking", fig_tob)

    if cols["alcohol"]:
        alc = view_df[cols["alcohol"]].map(flag_yes).value_counts()
        fig_alc = px.bar(
            x=alc.index,
            y=alc.values,
            labels={"x": "Alcohol/drugs", "y": "Count"},
            title="Alcohol / substance use",
            text=alc.values,
        )
        fig_alc.update_traces(textposition="outside")
        add_figure_slide("Lifestyle – Alcohol / Substance Use", fig_alc)

    if cols["tobacco"] and cols["alcohol"]:
        combo = (
            view_df.assign(
                Tobacco=view_df[cols["tobacco"]].map(flag_yes),
                Alcohol=view_df[cols["alcohol"]].map(flag_yes),
            )
            .groupby(["Tobacco", "Alcohol"])
            .size()
            .reset_index(name="Count")
        )
        fig_combo = px.treemap(
            combo,
            path=["Tobacco", "Alcohol"],
            values="Count",
            title="Joint distribution of tobacco & alcohol exposure",
        )
        add_figure_slide(
            "Lifestyle – Combined Tobacco × Alcohol Risk Map", fig_combo
        )

    # ---------------- COMMUNITY CHARTS ----------------
    if locality_col:
        agg = view_df.copy()
        agg["is_unhealthy"] = (
            agg["Health Status"] != "Healthy"
            if "Health Status" in agg.columns
            else 0
        ).astype(int)
        if glu_col:
            agg["high_glu"] = (agg[glu_col] >= 126).astype(int)
        else:
            agg["high_glu"] = 0
        if bmi_col:
            agg["obese"] = (agg[bmi_col] >= 30).astype(int)
        else:
            agg["obese"] = 0

        grp_c = (
            agg.groupby(locality_col)
            .agg(
                Population=(ID_COL, "count"),
                Unhealthy_rate=("is_unhealthy", "mean"),
                Diabetes_rate=("high_glu", "mean"),
                Obesity_rate=("obese", "mean"),
            )
            .reset_index()
        )
        grp_c["Unhealthy_rate"] *= 100
        grp_c["Diabetes_rate"] *= 100
        grp_c["Obesity_rate"] *= 100

        grp_top = grp_c.sort_values("Population", ascending=False).head(25)

        fig_pop = px.bar(
            grp_top,
            x=locality_col,
            y="Population",
            title="Top localities by number of records",
            text="Population",
        )
        fig_pop.update_traces(textposition="outside")
        fig_pop.update_layout(xaxis_tickangle=-45)
        add_figure_slide("Community – Population by Locality", fig_pop)

        heat_df = grp_top.set_index(locality_col)[
            ["Unhealthy_rate", "Diabetes_rate", "Obesity_rate"]
        ]
        fig_heat = px.imshow(
            heat_df,
            labels=dict(
                x="Metric",
                y="Locality",
                color="Percentage (%)",
            ),
            title="Locality-wise risk indicators",
            aspect="auto",
        )
        add_figure_slide("Community – Locality Risk Heatmap", fig_heat)

    # ---------------- SOCIOECONOMIC / PROFESSION CHARTS -----------

    if cols["income"]:
        inc_counts = (
            view_df[cols["income"]].astype(str).value_counts().head(10)
        )
        fig_inc = px.bar(
            x=inc_counts.index,
            y=inc_counts.values,
            title="Top income categories (filtered)",
            labels={"x": "Income group", "y": "Count"},
            text=inc_counts.values,
        )
        fig_inc.update_traces(textposition="outside")
        fig_inc.update_layout(xaxis_tickangle=-30)
        add_figure_slide("Socioeconomic – Income Group Distribution", fig_inc)

        if "Health Status" in view_df.columns:
            inc_status = (
                view_df[[cols["income"], "Health Status"]]
                .dropna()
                .groupby([cols["income"], "Health Status"])
                .size()
                .reset_index(name="Count")
            )
            fig_inc_h = px.bar(
                inc_status,
                x=cols["income"],
                y="Count",
                color="Health Status",
                barmode="group",
                title="Health status by income group",
                text="Count",
            )
            fig_inc_h.update_layout(xaxis_tickangle=-30)
            fig_inc_h.update_traces(textposition="outside")
            add_figure_slide(
                "Socioeconomic – Health by Income Group", fig_inc_h
            )

    if cols["occupation"]:
        occ_counts = (
            view_df[cols["occupation"]]
            .astype(str)
            .value_counts()
            .head(15)
        )
        fig_occ = px.bar(
            x=occ_counts.index,
            y=occ_counts.values,
            title="Top professions (filtered)",
            labels={"x": "Profession", "y": "Count"},
            text=occ_counts.values,
        )
        fig_occ.update_traces(textposition="outside")
        fig_occ.update_layout(xaxis_tickangle=-30)
        add_figure_slide(
            "Socioeconomic – Profession Distribution", fig_occ
        )

        if "Health Status" in view_df.columns:
            occ_status = (
                view_df[[cols["occupation"], "Health Status"]]
                .dropna()
                .groupby([cols["occupation"], "Health Status"])
                .size()
                .reset_index(name="Count")
            )
            # Keep only top 15 professions by total count
            rank = (
                occ_status.groupby(cols["occupation"])["Count"]
                .sum()
                .sort_values(ascending=False)
                .head(15)
                .index
            )
            occ_status_top = occ_status[
                occ_status[cols["occupation"]].isin(rank)
            ]
            fig_occ2 = px.bar(
                occ_status_top,
                x=cols["occupation"],
                y="Count",
                color="Health Status",
                barmode="group",
                title="Health status distribution by profession (Top 15)",
                text="Count",
            )
            fig_occ2.update_layout(
                xaxis_title="Profession",
                yaxis_title="Number of persons",
                xaxis_tickangle=-65,
            )
            fig_occ2.update_traces(textposition="outside")
            add_figure_slide(
                "Socioeconomic – Health by Profession", fig_occ2
            )

    # ---------------- Save PPT -------------------------
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf