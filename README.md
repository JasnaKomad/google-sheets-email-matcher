# google-sheets-email-matcher
Automatically populate emails in Google Sheets by matching records from a database tab. Supports 150k+ rows, real-time triggers, and conditional formatting.
# Google Sheets Email Matcher

Automatically populate email addresses in Google Sheets by matching records from a database tab.

## 🚀 Features

- **Real-time auto-fill** - Emails appear automatically when you type name/company
- **3 execution modes** - Fill all, fill only empty, or recheck all
- **Handles 150k+ records** - Optimized with in-memory mapping
- **Conditional formatting ready** - Green for found, red for not found
- **Clickable hyperlinks** - One click to send email
- **Audit trail** - Timestamp notes on every change

## 📋 How It Works
Database Tab (150k rows) Teams Tab (50k rows)

┌─────┬─────┬──────────┬─────┐ ┌─────┬─────┬──────────┬─────┐
│First│Last │Company │Email│ 			│First│Last │Company │Email│
├─────┼─────┼──────────┼─────┤ ├─────┼─────┼──────────┼─────┤
│John │Brown│Archetype │j...│ 		──► 	│John │Brown│Archetype │j...│
│Jane │Doe │Google │j...│ │Aad..│P... │Archetype │?...│
└─────┴─────┴──────────┴─────┘ └─────┴─────┴──────────┴─────┘

## 🔧 Installation

### Step 1: Open Apps Script
1. Open your Google Sheet
2. Click **Extensions** → **Apps Script**

### Step 2: Add the Code
1. Delete all existing code
2. Copy everything from `Code.gs`
3. Paste into the editor

### Step 3: Save & Run
1. Click **Save** (💾 icon)
2. Click **Run** → **Review permissions** → **Allow**

### Step 4: Use the Menu
After installation, a new menu appears: **📧 Email Generator**

## 📖 Usage Guide

| Menu Option | When to Use |
|-------------|-------------|
| **Fill All Emails** | First time running the script |
| **Fill Only Empty Emails** | After adding new rows |
| **Recheck All Emails** | After fixing typos in names/companies |
| **Setup Trigger** | Enable auto-fill on edit (runs automatically) |
| **Delete Trigger** | Disable auto-fill |

## 🎨 Conditional Formatting (Optional)

To add color coding:

1. Select column D
2. **Format** → **Conditional formatting**
3. Add rule: `Text contains` → `NOT_FOUND_IN_DATABASE` → Red background
4. Add rule: `Custom formula is` → `=ISTEXT(D2)` → Green background

## 📊 Performance

| Metric | Value |
|--------|-------|
| Max Database rows | 150,000+ |
| Max Teams rows | 50,000+ |
| Processing time | 2-5 seconds |
| Auto-fill delay | 200ms |

## 🛠️ Requirements

- Google account (free)
- Two sheets in the same spreadsheet:
  - `Teams` (columns: First Name, Last Name, Company, Email)
  - `Database` (columns: First Name, Last Name, Company, Email)

## 📝 File Structure
google-sheets-email-matcher/
├── Code.gs # Main Apps Script code
├── README.md # This file
└── LICENSE # MIT license (optional)


## 📧 Support

For questions or customizations, open an issue on GitHub.

## 📄 License

MIT License - Free for personal and commercial use.

