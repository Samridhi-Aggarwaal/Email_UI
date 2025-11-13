# 📨 Outlook Email Automation Tool (PyQt5 + Outlook Integration)

A desktop application built using Python (PyQt5) that automates the process of composing and sending standardized Outlook emails through an interactive GUI.

This tool is ideal for teams handling employee block requests or onboarding/offboarding communications, allowing users to quickly fill in employee details, select recipient groups, and send formatted emails directly via Microsoft Outlook.

---

##  🚀 Features

- 🖥️ **Intuitive PyQt5 GUI** for seamless user interaction  
- 📬 **Dynamic recipient selection** from predefined groups or custom entries  
- 👥 **Multi-employee entry support** with add/remove options  
- 🏢 **Customizable roles and company dropdowns** with “Custom” entry fields  
- ✅ **Built-in validation** for missing or incorrect fields  
- 👀 **Preview before sending** — opens draft email in Outlook  
- ✉️ **Direct send option** via Outlook automation using `win32com`  
- 🎨 **Styled, responsive layout** with visual error highlighting

---

## ⚙️ Tech Stack

- **Language:** Python  
- **Libraries:** PyQt5, win32com.client  
- **Platform:** Windows (requires Microsoft Outlook installed)

---

## 🧠 How It Works

1. Select a recipient group (e.g., *CDT Block* or *DS Block*) or enter custom emails.  
2. Add employee details — *name, code, role, and company.*  
3. Review the auto-generated subject and email body.  
4. Choose **Preview and Send** (to open draft) or **Send** (to send directly).

---

## 🪄 Use Cases

- Automating repetitive email workflows for **HR, PM, or IT operations** teams  
- Sending **employee block or access requests**  
- Generating **formatted Outlook emails** without manual typing

---

## 💻 Run Locally

```bash
# Install dependencies
pip install pyqt5 pywin32

# Run the application
python outlook_email_app.py
```

---

## 🏷️ Repository Tags
`python` `pyqt5` `email-automation` `outlook` `win32com` `gui` `automation` `desktop-app`
