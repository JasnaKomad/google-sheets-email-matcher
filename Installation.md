# Detailed Installation Guide

## Prerequisites
- Google account (Gmail or Workspace)
- Google Sheet with two tabs named exactly:
  - `Teams` (columns: First Name, Last Name, Company, Email)
  - `Database` (columns: First Name, Last Name, Company, Email)

## Installation Steps

### 1. Open Apps Script Editor
- Open your Google Sheet
- Click `Extensions` → `Apps Script`
- A new tab will open

### 2. Add the Code
- Delete any existing code in the editor (Ctrl+A, Delete)
- Copy the entire code from `Code.gs`
- Paste into the editor (Ctrl+V)

### 3. Save the Project
- Click the 💾 **Save** icon
- Name your project: `Email Matcher`

### 4. Authorize the Script
- Click ▶️ **Run** button
- A popup appears: "Authorization Required"
- Click **Review Permissions**
- Select your Google account
- Click **Advanced** (bottom left)
- Click **Go to Email Matcher (unsafe)**
- Click **Allow**

### 5. Refresh Your Sheet
- Go back to your Google Sheet
- Refresh the page (F5)

### 6. Verify Installation
- You should see a new menu: **📧 Email Generator**

## First Time Use

1. Click **📧 Email Generator** → **Fill All Emails**
2. Wait for the completion message
3. Column D in Teams tab is now populated

## How to Test

Add a test row in Teams tab:
First Name: John
Last Name: Brown
Company: Archetype Ventures


The email should appear automatically within 2 seconds.

## Troubleshooting

### "Authorization Required" loop
- Make sure you added `/** @OnlyCurrentDoc */` at the top of the code
- Try incognito/private browsing mode

### Script timeout error
- Run **Fill Only Empty Emails** instead of Fill All
- Or run in smaller batches

### "Cannot read property 'getRange' of null"
- Verify sheet names are exactly `Teams` and `Database` (case-sensitive)

### Nothing happens when I edit
- Run **Setup Trigger** from the menu
- Or run **Delete Trigger** then **Setup Trigger** again

## Need Help?
Contact your developer for support.
