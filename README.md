# 🏥 Department Wise Test Price Updater

A clean, responsive web app for updating lab test prices in Excel files based on department-wise percentage rules — no backend required.

![HTML](https://img.shields.io/badge/HTML5-E34F26?style=flat&logo=html5&logoColor=white)
![CSS](https://img.shields.io/badge/CSS3-1572B6?style=flat&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat&logo=javascript&logoColor=black)
![SheetJS](https://img.shields.io/badge/SheetJS-217346?style=flat&logo=microsoft-excel&logoColor=white)

---

## ✨ Features

- 📂 **Drag & Drop** Excel upload (`.xlsx` / `.xls`)
- ✏️ **Department rules** via simple textarea input
- ⚡ **Instant processing** — runs entirely in the browser
- 📊 **Result summary** — total rows, updated, skipped, departments matched
- ⬇️ **One-click download** of the updated Excel file
- 🛡️ **Full error handling** — missing columns, bad format, empty rules
- 📱 **Fully responsive** — works on mobile and desktop

---

## 📁 Project Structure

```
dept-price-updater/
├── index.html          ← Main HTML page
├── assets/
│   ├── style.css       ← All styles & theme variables
│   └── app.js          ← All JavaScript logic
└── README.md           ← This file
```

---

## 🚀 How to Use

### Option 1 — Open directly
Just open `index.html` in any modern browser. No server needed.

### Option 2 — GitHub Pages
1. Fork or upload this repo to GitHub
2. Go to **Settings → Pages**
3. Set source to `main` branch, `/ (root)`
4. Visit `https://your-username.github.io/repo-name/`

---

## 📋 Excel File Format

Your uploaded Excel file must contain these **exact column headers** (case-insensitive):

| Column      | Description                        |
|-------------|-------------------------------------|
| Test Id     | Unique ID for the test              |
| Test Name   | Name of the test                    |
| Test Code   | Short code for the test             |
| Department  | Department the test belongs to      |
| Test Price  | ✎ **This column will be updated**   |

---

## 📝 Rules Format

Enter one rule per line in the textarea:

```
CLINICAL PATHOLOGY - 30%
BIOCHEMISTRY - 50%
MICROBIOLOGY - 20%
HEMATOLOGY - 40%
```

- **Case-insensitive** — `biochemistry` and `BIOCHEMISTRY` both work
- Departments **not mentioned** in rules are left unchanged
- Rows with non-numeric prices are safely **skipped**

---

## 🧮 Formula

```
New Price = (Old Price × Percentage) / 100
```

**Example:**
- Old Price = `170`, Rule = `CLINICAL PATHOLOGY - 30%`
- New Price = `(170 × 30) / 100` = **`51`**

---

## ⚠️ Error Handling

| Situation                  | Behavior                            |
|----------------------------|--------------------------------------|
| No file uploaded           | Shows error toast                    |
| Wrong file type            | Rejected with error message          |
| Empty rules textarea       | Shows error toast                    |
| Invalid rule format        | Line is skipped, others still work   |
| Missing required column    | Shows which columns are missing      |
| Non-numeric Test Price     | Row is skipped safely                |

---

## 🛠️ Tech Stack

| Technology | Purpose                          |
|------------|----------------------------------|
| HTML5      | Page structure & markup          |
| CSS3       | Styling, animations, responsive  |
| JavaScript | Logic, DOM manipulation          |
| [SheetJS (xlsx)](https://sheetjs.com/) | Read & write Excel files |
| Google Fonts | DM Mono + Syne typefaces       |

---

## 📦 Dependencies

All dependencies are loaded via CDN — no `npm install` required.

```html
<!-- SheetJS -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<!-- Google Fonts -->
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&display=swap" rel="stylesheet"/>
```

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

> Built with ❤️ for lab management workflows.
