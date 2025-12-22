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
    
    diet_raw = row.get(cols["diet"], None) if cols.get("diet") else None
    diet = parse_diet(diet_raw)

    # Blood pressure
    if pd.notna(sys) and pd.notna(dia):
        if sys >= 180 or dia >= 120:
            diagnosis.append("Hypertensive crisis")
            prognosis.append("Immediate stroke / organ damage risk.")
            rx.append("Urgent medical evaluation.")
        elif sys >= 140 or dia >= 90:
            diagnosis.append("Hypertension")
            prognosis.append("High risk of CVD and stroke.")
            rx.append("Low-salt diet, weight control, activity; consult physician.")
        elif sys >= 120 or dia >= 80:
            diagnosis.append("Pre-hypertension")
            prognosis.append("Likely to progress without lifestyle change.")
            rx.append("Lifestyle optimisation, regular BP monitoring.")

    # Glucose
    if pd.notna(glu):
        if glu >= 126:
            diagnosis.append("Diabetes (fasting)")
            prognosis.append("Micro/macro-vascular complication risk.")
            rx.append("HbA1c, medical review, diet + activity plan.")
        elif 100 <= glu < 126:
            diagnosis.append("Pre-diabetes (fasting)")
            prognosis.append("High risk of progression to diabetes.")
            rx.append("Weight reduction, 150+ min/wk activity, low refined carbs.")

    # Cholesterol
    if pd.notna(chol):
        if chol >= 240:
            diagnosis.append("High cholesterol")
            prognosis.append("Atherosclerotic CVD risk.")
            rx.append("Reduce saturated/trans fats, full lipid panel.")
        elif 200 <= chol < 240:
            diagnosis.append("Borderline high cholesterol")
            prognosis.append("Moderate CVD risk.")
            rx.append("Dietary modification, recheck in 3–6 months.")

    # BMI
    if pd.notna(bmi):
        if bmi >= 30:
            diagnosis.append("Obesity")
            prognosis.append("Metabolic syndrome & CVD risk.")
            rx.append("Structured weight management, caloric deficit, exercise.")
        elif 25 <= bmi < 30:
            diagnosis.append("Overweight")
            prognosis.append("Increased cardiometabolic risk.")
            rx.append("Increase activity, portion control, diet optimisation.")

    # Kidney / liver
    if pd.notna(creat) and creat > 1.2:
        diagnosis.append("Possible renal impairment")
        prognosis.append("Needs eGFR & urine protein assessment.")
        rx.append("Avoid nephrotoxins, check renal profile.")
    
    if pd.notna(alt) and alt > 40:
        diagnosis.append("Elevated ALT (transaminitis)")
        prognosis.append("Fatty liver / hepatitis / drug effect possible.")
        rx.append("Review alcohol, obesity, hepatotoxic drugs; liver evaluation.")

    # Lifestyle
    tob_flag = flag_yes(row.get(cols["tobacco"])) if cols.get("tobacco") else "Unknown"
    alc_flag = flag_yes(row.get(cols["alcohol"])) if cols.get("alcohol") else "Unknown"
    
    if tob_flag == "Yes":
        diagnosis.append("Tobacco exposure")
        prognosis.append("↑ risk of CVD, COPD, cancer.")
        rx.append("Cessation counselling, nicotine replacement as needed.")
    if alc_flag == "Yes":
        diagnosis.append("Alcohol / substance use")
        prognosis.append("Liver, mental health, injury risk.")
        rx.append("Limit intake; consider de-addiction help for heavy use.")

    # Diet note
    if diet == "Vegetarian":
        rx.append("Ensure adequate protein: dals, soy, paneer, nuts & seeds.")
    elif diet == "Non-Vegetarian":
        rx.append("Prefer fish & skinless poultry; limit fried & red meat.")

    # Final status
    if not diagnosis:
        status = "Healthy"
        diagnosis.append("No major risk flags detected.")
        prognosis.append("Maintain healthy lifestyle & periodic screening.")
        rx.append("Balanced diet, physical activity, regular check-ups.")
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
