# 🟨 Excel Crosshair Highlighter

**Author:** Akshay Solanki
**License:** Unlicensed
**Version:** 1.0 — Final (Viewport-Aware + Cached Scope + No StatusBar)

---

## 📘 Overview

The **Excel Crosshair Highlighter** is a lightweight, high-performance VBA add-in that highlights the **active row and column** as you move around your worksheet. Once installed, it runs automatically every time Excel opens and works across all workbooks — just like a built-in feature.

---

## 🧩 Installation Guide (One-Time Setup)

Follow these steps carefully. You only need to do this once.

### 1️⃣ Copy the VBA Code

1.  Download or open the file **`CrosshairHighlighter.txt`** from this repository.
2.  Copy **all** of the VBA code from the file.

### 2️⃣ Add the Code in Excel

1.  Open Excel.
2.  Press **Alt + F11** to open the **Visual Basic Editor (VBE)**.
3.  From the top menu, select **Insert → Module**.
4.  Paste the copied VBA code into the new module.

### 3️⃣ Save as an Excel Add-in (.xlam)

This is crucial for permanent, automatic loading.

1.  In the VBE, press **Ctrl + S** or go to **File → Save As**.
2.  In the “Save as type” dropdown, select **Excel Add-In (\*.xlam)**.
3.  Navigate to Excel’s default Add-ins location. You can paste this path directly into the file dialog's address bar:
    
    `%appdata%\Microsoft\AddIns`
    
4.  Save the file as:
    
    `CrosshairHighlighter.xlam`

### 4️⃣ Enable the Add-in

1.  Go back to Excel.
2.  Click **File → Options → Add-ins**.
3.  At the bottom, in **Manage: Excel Add-ins**, click **Go…**
4.  Check **CrosshairHighlighter** in the list and click **OK**.
    ✅ The add-in will now load automatically every time Excel starts.

### 5️⃣ Add a Custom Ribbon Button

To easily toggle the Crosshair Highlighter ON/OFF:

1.  Go to **File → Options → Customize Ribbon**.
2.  On the right side, click **New Tab** (you can rename it, e.g., “Tools”).
3.  Inside your new tab, click **New Group** (rename it “Crosshair”).
4.  On the left side (“Choose commands from”), select **Macros**.
5.  Find the macro **Crosshair\_Run** in the list.
6.  Click **Add >>** to move it into your new group.
7.  (Optional but Recommended) Click **Modify…** to rename it to something like **"Toggle Crosshair"** and pick a suitable icon.
    You’ll now have a custom **Ribbon button** to turn the feature ON or OFF instantly.

---

## 🪄 Usage

* Click your **Toggle Crosshair** button to enable or disable the feature.
* When enabled, the active **row and column** will be softly highlighted.
* Use arrow keys, Tab, or Enter — the highlight follows your selection smoothly.
* The add-in remains active across all new Excel files automatically.

---

## ⚠️ Known Limitation

When you first enable the Crosshair Highlighter via the ribbon button, the highlight **won’t appear immediately**.

👉 Simply press **any arrow key** (↑, ↓, ←, or →) once — the highlight will then follow and function normally after that.

---

## 🧑‍💻 Author

**Akshay Solanki**
Final optimized version — November 2025.
*Viewport-aware • Cached scope • Flicker-free • StatusBar disabled.*

---

## 📄 License

Unlicensed — free for personal and educational use.

---

## 🌟 Summary

| Feature | Status |
| :--- | :--- |
| **Viewport-aware** highlighting | ✅ |
| **Persistent** across sessions | ✅ |
| Works in **all** Excel workbooks | ✅ |
| Smooth, **instant** response | ✅ |
| One-click **toggle** in Ribbon | ✅ |

Enjoy a faster, cleaner, and smarter Excel experience with the **Crosshair Highlighter**.
