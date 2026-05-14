# 📋 Google Form Automation — Excel to Google Form Generator

> **Upload an Excel file with your questions → Get a fully built Google Form in seconds.**

[![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D14-brightgreen)](https://nodejs.org/)

---

## 🎯 What This Does

| You provide | You get |
|---|---|
| An Excel file with question definitions | A live Google Form with all your questions |
| Column: Question, Type, Options, Required | Questions appear exactly as defined |
| — | A Google Sheet auto-linked for collecting responses |
| — | An email with your form link when it's ready |

**No manual form building. No copy-pasting. Just fill in a spreadsheet and run.**

---

## 📁 Repository Structure

```
google-form-automation/
│
├── scripts/
│   └── Code.gs                  # Main script — reads Excel, builds form
│
├── data/
│   └── sample_questions.csv     # Template showing the required Excel format
│
├── README.md                    # This file
├── .gitignore                   # Keeps credentials & node_modules out of Git
├── package.json                 # npm scripts for clasp commands
└── appsscript.json              # Apps Script permissions manifest
```

---

## 📊 Excel File Format

Your Excel (or Google Sheet) must have **5 columns** with this exact header row:

| Column | Header | Required? | Description |
|---|---|---|---|
| A | `Question` | ✅ Yes | The question text shown to the user |
| B | `Type` | ✅ Yes | The type of question (see table below) |
| C | `Options` | ⚠️ Some types | Comma-separated choices |
| D | `Required` | ✅ Yes | `TRUE` or `FALSE` |
| E | `Help Text` | ❌ Optional | Hint shown below the question |

### Supported Question Types

| Type value | What it creates | Options needed? |
|---|---|---|
| `short_answer` | Single-line text box | No |
| `paragraph` | Multi-line text box | No |
| `multiple_choice` | Radio buttons — pick one | ✅ Yes |
| `checkbox` | Checkboxes — pick many | ✅ Yes |
| `dropdown` | Dropdown menu — pick one | ✅ Yes |
| `linear_scale` | Rating scale (e.g. 1–5) | ✅ Yes — format: `min, max, lowLabel, highLabel` |
| `date` | Date picker | No |
| `time` | Time picker | No |

### Example rows:

```
Question                          | Type            | Options                        | Required | Help Text
----------------------------------|-----------------|--------------------------------|----------|--------------------
What is your name?                | short_answer    |                                | TRUE     |
Which department?                 | dropdown        | HR, Engineering, Sales         | TRUE     |
Select interests                  | checkbox        | Tech, Art, Sports              | FALSE    | Pick all that apply
Rate your experience              | linear_scale    | 1, 5, Terrible, Amazing        | TRUE     | 1=worst, 5=best
Tell us more                      | paragraph       |                                | FALSE    | Optional
```

> 📎 See `data/sample_questions.csv` for a ready-to-use template.

---

## ✅ Prerequisites

| Tool | Purpose | Download |
|---|---|---|
| **Node.js** (v14+) | Runs clasp | [nodejs.org](https://nodejs.org) |
| **Git** | Version control | [git-scm.com](https://git-scm.com) |
| **A Google Account** | Access Apps Script, Drive, Forms | [accounts.google.com](https://accounts.google.com) |
| **A GitHub Account** | Host your repository | [github.com](https://github.com) |

---

## 🚀 Setup Guide — Step by Step

### Step 1 — Prepare your Excel file

1. Open `data/sample_questions.csv` — this is your template
2. Fill in your questions following the format above
3. Save it as `.xlsx` (Excel) or keep it as `.csv`

---

### Step 2 — Upload Excel to Google Drive

1. Go to [drive.google.com](https://drive.google.com)
2. Click **"+ New" → "File upload"** and upload your Excel file
3. Once uploaded, right-click the file → **"Open with Google Sheets"**
   - Google automatically converts Excel → Google Sheet
4. Copy the **Spreadsheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_ID/edit
                                          ↑ copy this part
   ```
5. Make sure your questions are on a tab named **`Questions`**
   (or change `QUESTIONS_TAB_NAME` in the script)

---

### Step 3 — Install clasp and connect the script

```bash
# Install dependencies (installs clasp)
npm install

# Log in to Google with clasp
npm run login
```

> A browser window opens — sign in and grant access.

```bash
# Create a new Apps Script project
npx clasp create --title "Google Form Automation" --type standalone
```

This creates a `.clasp.json` file with your Script ID.

---

### Step 4 — Configure the script

Open `scripts/Code.gs` and update the `CONFIG` block at the top:

```javascript
var CONFIG = {
  // Your Google Sheet ID (from Step 2)
  QUESTIONS_SHEET_ID: "paste-your-spreadsheet-id-here",

  // Tab name inside your sheet (must match exactly)
  QUESTIONS_TAB_NAME: "Questions",

  // What the generated Google Form will be called
  FORM_TITLE: "My Survey Form",

  // Description shown at the top of the form
  FORM_DESCRIPTION: "Please fill in this form.",

  // Send an email when the form is created?
  SEND_EMAIL_ON_COMPLETION: true,
  NOTIFICATION_EMAIL: "your@email.com",
};
```

---

### Step 5 — Push script to Google Apps Script

```bash
# Push your local code to Google's servers
npm run push
```

---

### Step 6 — Run the form generator

1. Open your Apps Script project in the browser:
   ```bash
   npx clasp open
   ```
2. In the editor, select **`generateForm`** from the function dropdown
3. Click **▶ Run**
4. On first run, click **"Review Permissions"** and grant access
5. Check the **Execution log** — it will print your form link!

### ✅ Test first (no Excel needed)

Before using your real data, test the script works:
1. In the Apps Script editor, select **`testWithSampleData`**
2. Click **▶ Run**
3. It creates a form with 8 built-in sample questions
4. Check the log for the form link

---

## 🔄 Day-to-Day Workflow

```
Edit your Excel / Sheet questions
          ↓
npm run push       ← send code changes to Google
          ↓
Run generateForm   ← in Apps Script editor
          ↓
Get form link from email or logs
          ↓
Share form with respondents
          ↓
git add + git commit + git push   ← save your code to GitHub
```

### Useful commands

```bash
npm run push      # Push local code → Apps Script
npm run pull      # Pull any changes made in the browser editor → local
npm run open      # Open your Apps Script project in the browser
npm run logs      # View recent execution logs
npm run deploy    # Create a versioned deployment
```

---

## 🐙 GitHub Setup

### Create a GitHub repository:

1. Go to [github.com/new](https://github.com/new)
2. Name it: `google-form-automation`
3. Do NOT initialize with a README (we already have one)
4. Click **"Create repository"**

### Push to GitHub:

```bash
git init
git add .
git commit -m "feat: excel to google form generator"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/google-form-automation.git
git push -u origin main
```

> **Authentication**: Use a Personal Access Token as your password.  
> Get one at: **GitHub → Settings → Developer settings → Personal access tokens**

---

## 🔑 Creating a GitHub Personal Access Token (PAT)

1. GitHub → click your avatar → **Settings**
2. Scroll to **"Developer settings"** → **"Personal access tokens"** → **"Tokens (classic)"**
3. Click **"Generate new token (classic)"**
4. Set:
   - Note: `google-form-automation`
   - Expiration: 90 days
   - Scope: ✅ `repo`
5. Click **"Generate token"** — **copy it immediately!**
6. When Git asks for password, paste your token

---

## ✍️ Example Commit Messages

```bash
git commit -m "feat: add linear scale question type support"
git commit -m "fix: handle empty rows in Excel gracefully"
git commit -m "chore: update SPREADSHEET_ID in CONFIG"
git commit -m "docs: add dropdown type to README examples"
git commit -m "feat: send completion email with form links"
```

---

## 🔧 Troubleshooting

### ❌ "Sheet tab not found"
- Make sure your Google Sheet has a tab named exactly `Questions` (capital Q)
- Or update `QUESTIONS_TAB_NAME` in the `CONFIG` block

### ❌ "clasp: command not found"
```bash
npx clasp login    # Use npx prefix always
```

### ❌ Script permission error on first run
1. Click **"Review Permissions"** in the consent screen
2. Click **"Advanced"** → **"Go to [project name] (unsafe)"**
3. Click **"Allow"**

### ❌ Apps Script API not enabled
1. Go to [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
2. Toggle **"Google Apps Script API"** → ON

### ❌ Form created but no questions
- Check that Column B (Type) uses exactly the supported type names (all lowercase, underscores)
- Check that Column A is not empty for any row you expect to be added

---

## 📄 License

MIT — free to use, modify, and distribute.

---

*Built with ❤️ using [Google Apps Script](https://developers.google.com/apps-script) + [clasp](https://github.com/google/clasp)*
