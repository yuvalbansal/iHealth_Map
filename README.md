# 🩺 iHealth Map — Community Health Analytics Dashboard

**iHealth Map** is a Streamlit-based interactive dashboard for analyzing population health screening data.  
It converts Excel-based lab and lifestyle datasets into actionable clinical, lifestyle, community, and socioeconomic insights with downloadable reports.

---

## 🚀 Getting Started

Follow the steps below to clone the repository, set up the environment, and run the application locally.

---

## 📦 1. Clone the Repository

```bash
git clone https://github.com/yuvalbansal/iHealth_Map.git
cd iHealth_Map
```

---

## 🐍 2. Create & Activate a Python Virtual Environment

- **Linux / macOS**

```bash
python3 -m venv myenv
source myenv/bin/activate
```

- **Windows (PowerShell)**

```powershell
python -m venv myenv
myenv\Scripts\activate
```

> 💡 Make sure you are using **Python 3.9 or newer**.

---

## 📥 3. Install Dependencies

All required packages are listed in `requirements.txt`.

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

---

## ▶️ 4. Run the Streamlit App

```bash
streamlit run Big_B081225.py
```

Once started, Streamlit will automatically open the app in your browser.

If it does not, open the URL shown in the terminal (usually `http://localhost:8501`).

---

## 📊 5. Using the App

1. Upload an **Excel (.xlsx)** file containing health screening data

1. Apply filters (age, gender, diet, lifestyle)

1. Navigate between pages:

   - Overview
   - Clinical
   - Lifestyle
   - Community
   - Socioeconomic
   - Downloads

1. Download:

   - Filtered Excel data
   - Individual PDF health reports
   - Population-level PPT summaries

---

## 🛠️ Troubleshooting

- **Arrow / DataFrame warnings**

  Ensure all numeric columns are properly formatted in the Excel file.

- **PPT export disabled**

  Install `python-pptx` and restart the app:

  ```bash
  pip install python-pptx
  ```

- **Port already in use**

  ```bash
  streamlit run Big_B081225.py --server.port 8502
  ```
