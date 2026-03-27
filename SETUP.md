# Developer Setup Guide

This guide walks you through setting up a full development environment for **Superbill Processor** on a Windows workstation using VS Code and GitHub Copilot.

---

## 1. Install Python

1. Open the **Microsoft Store**, search for **Python 3.13**, and install it.  
   *(Alternatively, download the installer from [python.org](https://www.python.org/downloads/) and make sure to tick **"Add Python to PATH"** during setup.)*
2. Open **PowerShell** and verify:
   ```
   python --version
   ```
   You should see `Python 3.13.x`.

---

## 2. Install Git

1. Download Git from [git-scm.com](https://git-scm.com/download/win) and run the installer (all defaults are fine).
2. Verify in PowerShell:
   ```
   git --version
   ```

---

## 3. Install Visual Studio Code

1. Download VS Code from [code.visualstudio.com](https://code.visualstudio.com/) and run the installer.
2. During setup, tick **"Add to PATH"** and **"Add 'Open with Code' action"** for convenience.

---

## 4. Install the Python and GitHub Copilot extensions in VS Code

1. Open VS Code.
2. Click the **Extensions** icon in the left sidebar (or press `Ctrl+Shift+X`).
3. Search for and install:
   - **Python** (by Microsoft)
   - **GitHub Copilot** (by GitHub)
   - **GitHub Copilot Chat** (by GitHub)
4. After installing GitHub Copilot, click **Sign in to GitHub** when prompted and follow the browser flow to authorise your GitHub account.  
   *(GitHub Copilot requires a Copilot subscription or a free trial on your GitHub account.)*

---

## 5. Clone the repository

1. In PowerShell, navigate to where you want the project folder, for example:
   ```
   cd "$HOME\projects"
   ```
2. Clone the repo:
   ```
   git clone <repository-url> superbill
   cd superbill
   ```
   *(Replace `<repository-url>` with the actual GitHub URL provided to you.)*

---

## 6. Open the project in VS Code

```
code .
```

VS Code will open the `superbill` folder. You should see `superbill_processor.py` in the Explorer panel on the left.

---

## 7. Select the Python interpreter

1. Press `Ctrl+Shift+P` and type **Python: Select Interpreter**.
2. Choose the Python 3.13 entry from the list.

---

## 8. Install project dependencies

Open the **integrated terminal** in VS Code (`Ctrl+`` ` ``) and run:

```
pip install pandas openpyxl tkinterdnd2
```

To also install **Playwright** and **python-dotenv** for `fetch_superbill_and_merge.py` (or install everything from the project file):

```
pip install -r requirements.txt
python -m playwright install chromium
```

To also be able to build the `.exe` file, install PyInstaller:

```
pip install pyinstaller
```

---

## 9. Run the application

In the integrated terminal:

```
python superbill_processor.py
```

The GUI window should appear. You can also press **F5** in VS Code to launch it with the debugger attached (useful for setting breakpoints).

---

## 10. Build the standalone .exe (optional)

When you want to produce a distributable single-file executable:

```
python -m PyInstaller SuperbillProcessor.spec
```

The finished `.exe` will be placed in the `dist\` folder.

---

## 11. Using GitHub Copilot

With the Copilot extension installed and signed in, you have two main ways to get AI assistance:

- **Inline suggestions** — as you type in any file, Copilot will suggest completions. Press `Tab` to accept.
- **Copilot Chat** — click the chat icon in the left sidebar (or press `Ctrl+Alt+I`). You can ask questions like:
  - *"Explain what this function does"*
  - *"Add error handling to this block"*
  - *"Write a function that does X"*

You can also highlight any block of code, right-click, and choose **Copilot → Explain This** or **Copilot → Fix This**.

---

## Lynx fetch: download confirmation popup

The script looks for the in-page prompt **“Access other apps and services on this device”** and clicks **Allow** automatically.

If that does not match your UI, set one of these in `.env` after inspecting the control in DevTools:

- **`LYNX_DOWNLOAD_CONFIRM_SELECTOR`** — CSS selector (highest precedence).
- **`LYNX_DOWNLOAD_CONFIRM_BUTTON`** — accessible name if it is not **Allow**.

Native `window.confirm` / `alert` is handled automatically by the fetch script.

---

## Quick reference

| Task | Command |
|------|---------|
| Run the app | `python superbill_processor.py` |
| Run top-level fetch + merge launcher UI | `python superbill_workflow_ui.py` |
| Fetch Lynx + merge | `python fetch_superbill_and_merge.py --month 03/2026 --master "path\to\master.xlsx" --yes` |
| Save download to custom folder | add `--download-dir "path\to\downloads"` |
| Interactive mode | default is interactive; add `--no-interactive` / `--batch` / `-b` for unattended runs |
| Build the EXE | `python -m PyInstaller SuperbillProcessor.spec` |
| Install / update dependencies | `pip install pandas openpyxl tkinterdnd2 pyinstaller` |
| Install including Playwright | `pip install -r requirements.txt` then `python -m playwright install chromium` |
| Open Copilot Chat | `Ctrl+Alt+I` |
| Run with debugger | `F5` in VS Code |
