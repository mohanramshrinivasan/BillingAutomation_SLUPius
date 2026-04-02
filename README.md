# 📚 Billing Automation - SLU Pius Library

## 🚀 Overview

This project is a **data-driven billing automation tool** designed for preparing library billing notice emails efficiently.

It:
- 📥 Reads billing data from Excel
- 👥 Groups multiple items for the same patron into one email
- 📝 Generates email drafts in **Microsoft Outlook**
- ⚠️ Marks emails as **High Importance**
- 🔄 Updates invoice dates in Excel
- ✅ Changes `Run` from `Yes` → `Done`
- 📄 Logs all activity

👉 No coding required — just update Excel and double-click the batch file.

---

## 📂 Folder Structure

```
BillingAutomation_SLUPius/
├── billing_prepare.py
├── run_billing.bat
├── requirements.txt
├── README.md
├── .gitignore
├── data/
│   ├── templates/
│   │   ├── template_first_notice_1.txt
│   │   ├── template_second_notice_2.txt
│   │   ├── template_third_notice_3.txt
│   │   └── template_final_notice_4.txt
│   ├── input/
│   │   └── input_data.json
│   └── excel/
│       └── billing_data.xlsx
├── logs/
└── output/
```

---

## 🧰 Software Required

Before running, install:

### 1️⃣ Python
Download from: https://www.python.org/downloads/windows/

✔️ During installation:
```
☑ Add Python to PATH
```

---

### 2️⃣ Microsoft Outlook (Classic)

⚠️ MUST use **Classic Outlook Desktop App**

❌ Do NOT use:
- New Outlook (does not support automation)

---

### 3️⃣ Install Required Python Packages

Open Command Prompt in the project folder:

```
python -m pip install -r requirements.txt
```

Or manually:

```
python -m pip install pandas openpyxl pywin32
```

---

## 📥 Step 1 — Clone the Repository

```
git clone https://github.com/mohanramshrinivasan/BillingAutomation_SLUPius.git
cd BillingAutomation_SLUPius
```

📍 Recommended location:
```
C:\BillingAutomation_SLUPius
```

---

## ⚙️ Step 2 — Configure Input JSON

Open:
```
data/input/input_data.json
```

Update:
- 📧 sender email
- 📬 reply-to email
- 📊 column names
- 📄 template file names
- 📑 subject lines

---

## 📊 Step 3 — Prepare Excel File

Open:
```
data/excel/billing_data.xlsx
```

### 🔑 Important Rules

- Column 2 must be `Run`
- Only rows with `Run = Yes` will be processed
- Each item must be on its own row
- Same person must have the same email
- Grouped rows must be in the same invoice stage

---

### 📌 Example

| Run | Title | Email | Fines |
|-----|------|------|------|
| Yes | Rich Dad Poor Dad | test@slu.edu | 120 |
| Yes | Education | test@slu.edu | 120 |

➡️ Output:

```
Rich Dad Poor Dad - $120.00
Education - $120.00

TOTAL: $240.00
```

---

## 📝 Step 4 — Update Templates

Location:
```
data/templates/
```

---

### 🔁 Replace OLD format:

```
{Title} - ${Fines}

TOTAL: ${Fines}
```

---

### ✅ Use NEW format:

```
{Item_Lines}

TOTAL: ${Total_Fines}
```

---

### 📄 Detailed section:

```
Item Details:
{Item_Details}
```

---

### 🧩 Available Placeholders

- `{Preferred_Name}`
- `{Item_Lines}`
- `{Item_Details}`
- `{Total_Fines}`
- `{Today}`
- `{From_Email}`

---

## ▶️ Step 5 — Run the Script

### Option 1 (Recommended)

Double-click:
```
run_billing.bat
```

---

### Option 2 (Manual)

```
python billing_prepare.py
```

---

## ⚙️ What Happens When You Run

1. Loads configuration from JSON  
2. Reads Excel data  
3. Filters rows where `Run = Yes`  
4. Groups rows by email  
5. Determines invoice stage  
6. Builds grouped email content  
7. Opens Outlook draft  
8. Sets **High Importance**  
9. Updates Excel invoice date  
10. Changes `Run` to `Done`  
11. Writes logs  

---

## 📤 Output

### 📧 Outlook
- Draft email created
- Opens automatically
- You review and click **Send**

---

### 📊 Excel Updates
- Invoice date updated
- `Run` becomes `Done`

---

### 📄 Logs

Stored in:
```
logs/
```

Example:
```
logs-040126-203746.txt
```

---

## ⚠️ Important Notes

- ✔️ Use Classic Outlook only  
- ✔️ Keep Excel clean and consistent  
- ✔️ Do not upload real patron data to GitHub  
- ✔️ Keep repository private for production  

---

## 🛠️ Troubleshooting

### ❌ Python not found
Install Python and enable PATH

---

### ❌ pandas not found
```
python -m pip install -r requirements.txt
```

---

### ❌ Outlook error
- Open Outlook manually first  
- Ensure account is signed in  
- Use Classic Outlook  

---

### ❌ Mixed invoice stages
Ensure grouped rows are in the same stage

---

## 🔄 Recommended Workflow

1. Pull latest code  
2. Update Excel  
3. Set `Run = Yes`  
4. Double-click batch file  
5. Review drafts  
6. Click Send  
7. Check logs  

---

## 👨‍💻 Author

**Mohanram Shrinivasan**  
Saint Louis University  
Pius XII Memorial Library  

---

## ⭐ Summary

✔ Fully automated  
✔ Cross-machine compatible  
✔ Outlook integrated  
✔ Easy for non-technical users  

---