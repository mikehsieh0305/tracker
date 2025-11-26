# 📊 Green Rock PV Tracker (Electron App)

A desktop application built with **Electron** for visualizing and tracking construction or logistics progress using Gantt charts, Excel parsing, and data visualization.  
Supports **Windows & macOS** packaging using `electron-builder`.

> 🚧 *This project uses Chart.js, Frappe-Gantt, and XLSX to provide real-time progress visualization.*

---

## 🚀 Features

- 📂 **Import Excel Files** (xlsx) for project or material tracking
- 📅 **Gantt Chart Visualization** using `frappe-gantt`
- 📈 **Progress & Statistics Charting** using `chart.js`
- 💻 **Cross-Platform** (Windows `.exe` / macOS `.dmg`)
- 🔧 Easy packaging via `electron-builder`

---

## 📦 Tech Stack

| Tool | Usage |
|------|-------|
| **Electron** | Desktop application |
| **Chart.js** | Visualization & statistics |
| **Frappe-Gantt** | Gantt chart timeline |
| **XLSX (SheetJS)** | Excel import / data parsing |
| **Electron-Builder** | Packaging for macOS / Windows |

---

## 📁 Project Structure

tracker/
├─ build/
│ └─ icons/
│ ├─ mac/icon.icns
│ └─ win/icon.ico
├─ main.js # Main Electron process
├─ renderer/ # UI frontend pages (optional)
├─ package.json
└─ README.md


> ⚠️ Make sure your icons are placed correctly inside `build/icons/mac/` & `build/icons/win/`

---

## 🔧 Installation & Setup

### 1️⃣ Clone & Install Dependencies

```bash
git clone <your-repo-url>
cd tracker
npm install


2️⃣ Start Development Mode
npm start

📦 Build & Packaging

🛑 ❗ Before building, ensure you already supplied correct icons.

🐧 macOS Build
npm run dist:mac


Output: .dmg and .zip

Requires macOS system to build mac binaries

🪟 Windows Build
npm run dist:win


Output: .exe installer and .zip

Cross-Platform Build
npm run dist

⚙️ Build Configuration (package.json)
"build": {
  "appId": "com.example.myapp",
  "mac": {
    "category": "public.app-category.utilities",
    "target": ["dmg", "zip"],
    "icon": "build/icons/mac/icon.icns"
  },
  "win": {
    "target": ["nsis", "zip"],
    "icon": "build/icons/win/icon.ico"
  }
}

❓ Troubleshooting
🧱 electron-builder fails due to missing dependencies

🔧 Install build tools:

macOS:
xcode-select --install
brew install node

Windows:

Install Visual Studio Build Tools

Enable Desktop development with C++

📎 App icon not loading?

Check icon paths:

build/icons/mac/icon.icns
build/icons/win/icon.ico