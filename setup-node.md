# Developer Setup Guide — Node.js

This guide walks you through setting up a full development environment for the **Superbill Processor Node.js CLI** on a Windows workstation using VS Code and GitHub Copilot.

---

## 1. Install Node.js

1. Download the **LTS** installer from [nodejs.org](https://nodejs.org/) and run it (all defaults are fine).  
   The installer includes both `node` and `npm`.
2. Open **PowerShell** and verify:
   ```
   node --version
   npm --version
   ```
   You should see `v20.x.x` (or later) and `10.x.x` (or later).

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

## 4. Install the JavaScript and GitHub Copilot extensions in VS Code

1. Open VS Code.
2. Click the **Extensions** icon in the left sidebar (or press `Ctrl+Shift+X`).
3. Search for and install:
   - **ESLint** (by Microsoft) — optional but recommended for linting
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

VS Code will open the `superbill` folder. You should see `superbill_processor.js` in the Explorer panel on the left.

---

## 7. Install project dependencies

Open the **integrated terminal** in VS Code (`` Ctrl+` ``) and run:

```
npm install
```

This reads `package.json` and installs `exceljs` and `xlsx` into the `node_modules` folder.

---

## 8. Run the application

```
node superbill_processor.js <input.xlsx> <output.xlsx>
```

See the [CLI usage section in the README](README.md#nodejs-cli) for all available switches.

---

## 9. Build the standalone .exe (optional)

When you want to produce a distributable single-file executable (requires the `pkg` dev dependency, which is already listed in `package.json`):

```
npm run build
```

The finished `.exe` will be placed in the `dist\` folder as `SuperbillProcessor.exe`.

---

## 10. Using GitHub Copilot

With the Copilot extension installed and signed in, you have two main ways to get AI assistance:

- **Inline suggestions** — as you type in any file, Copilot will suggest completions. Press `Tab` to accept.
- **Copilot Chat** — click the chat icon in the left sidebar (or press `Ctrl+Alt+I`). You can ask questions like:
  - *"Explain what this function does"*
  - *"Add error handling to this block"*
  - *"Write a function that does X"*

You can also highlight any block of code, right-click, and choose **Copilot → Explain This** or **Copilot → Fix This**.

---

## Quick reference

| Task | Command |
|------|---------|
| Install dependencies | `npm install` |
| Run the CLI (interactive) | `node superbill_processor.js <input.xlsx> <output.xlsx>` |
| Run the CLI (batch, no prompts) | `node superbill_processor.js <input.xlsx> <output.xlsx> --yes` |
| Show usage | `node superbill_processor.js --help` |
| Build the EXE | `npm run build` |
| Open Copilot Chat | `Ctrl+Alt+I` |
