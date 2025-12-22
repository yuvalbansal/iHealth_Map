import pandas as pd
import numpy as np
from utils.formatting import parse_diet, flag_yes


def assess_health(row: pd.Series, cols: dict) -> pd.Series:
    diagnosis, prognosis, rx = [], [], []

    def val(key):
        return row.get(cols[key], np.nan) if cols.get(key) else np.nan

    bmi = val("bmi")
    sys = val("bp_sys")
    dia = val("bp_dia")
    glu = val("glucose")
    chol = val("chol")
    creat = val("creatinine")
    alt = val("alt")

    # Blood pressure
    if pd.notna(sys) and pd.notna(dia):
        if sys >= 140 or dia >= 90:
            diagnosis.append("Hypertension")
            prognosis.append("Elevated cardiovascular risk")
            rx.append("Lifestyle modification and BP monitoring")

    # Glucose
    if pd.notna(glu):
        if glu >= 126:
            diagnosis.append("Diabetes")
            prognosis.append("High metabolic risk")
            rx.append("Glycemic control and medical review")
        elif 100 <= glu < 126:
            diagnosis.append("Pre-diabetes")
            prognosis.append("Risk of progression to diabetes")
            rx.append("Diet and physical activity")

    # Cholesterol
    if pd.notna(chol) and chol >= 240:
        diagnosis.append("High cholesterol")
        prognosis.append("Atherosclerotic risk")
        rx.append("Dietary fat reduction")

    # BMI
    if pd.notna(bmi):
        if bmi >= 30:
            diagnosis.append("Obesity")
            prognosis.append("Metabolic syndrome risk")
            rx.append("Structured weight management")
        elif 25 <= bmi < 30:
            diagnosis.append("Overweight")
            prognosis.append("Increased cardiometabolic risk")
            rx.append("Lifestyle optimisation")

    # Kidney / liver
    if pd.notna(creat) and creat > 1.2:
        diagnosis.append("Possible renal impairment")
    if pd.notna(alt) and alt > 40:
        diagnosis.append("Elevated liver enzymes")

    # Lifestyle
    if cols.get("tobacco") and flag_yes(row.get(cols["tobacco"])) == "Yes":
        diagnosis.append("Tobacco exposure")
    if cols.get("alcohol") and flag_yes(row.get(cols["alcohol"])) == "Yes":
        diagnosis.append("Alcohol use")

    # Diet note
    if cols.get("diet"):
        diet = parse_diet(row.get(cols["diet"]))
        if diet == "Vegetarian":
            rx.append("Ensure adequate protein intake")

    # Final status
    if not diagnosis:
        status = "Healthy"
        diagnosis = ["No major risk detected"]
        prognosis = ["Maintain healthy lifestyle"]
        rx = ["Regular screening"]
    elif len(diagnosis) == 1:
        status = "At Risk"
    else:
        status = "Needs Attention"

    return pd.Series(
        {
            "Health Status": status,
            "Diagnosis": "; ".join(dict.fromkeys(diagnosis)),
            "Prognosis": "; ".join(dict.fromkeys(prognosis)),
            "Prescription": "; ".join(dict.fromkeys(rx)),
        }
    )


def apply_health_assessment(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    assessed = df.apply(lambda r: assess_health(r, cols), axis=1)
    return pd.concat([df.reset_index(drop=True), assessed.reset_index(drop=True)], axis=1)
