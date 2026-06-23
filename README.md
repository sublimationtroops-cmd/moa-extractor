# MOA Count Sheet Extractor — Streamlit App

No Google account. No Colab. Runs locally in your browser.

---

## Setup (one time only)

**Step 1 — Install Python** (if not already installed)
Download from https://python.org — version 3.9 or higher.

**Step 2 — Open a terminal / command prompt**
- Windows: press `Win + R`, type `cmd`, press Enter
- Mac: open Terminal from Applications

**Step 3 — Install dependencies**
Navigate to the folder where you saved these files, then run:

```
pip install -r requirements.txt
```

---

## Run the app

```
streamlit run app.py
```

Your browser will open automatically at http://localhost:8501

---

## How to use

1. Click **Browse files** and select one or more PO PDF files
2. (Optional) Enter SUBF number mapping if needed
3. Click **Extract to Excel**
4. Click **Download Excel** when it appears

---

## Files in this folder

| File | Purpose |
|------|---------|
| `app.py` | The main app — run this |
| `requirements.txt` | Python packages to install |
| `README.md` | This guide |
