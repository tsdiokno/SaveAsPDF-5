# SaveAsPDF-5 PowerPoint Add-in

One-click PDF export for PowerPoint — no menus, no hassle.

Are you tired of the endless clicks it takes to save your PowerPoint as a PDF in the right folder? You’re not alone. Many PowerPoint users struggle with this exact hassle every single day.

That’s why I built the **SaveAsPDF-5 Add-in** — with some help from ChatGPT, vibecoded into reality — the one-click solution that puts an end to the frustration.

![Demo GIF](https://raw.githubusercontent.com/tsdiokno/SaveAsPDF-5/refs/heads/main/SaveAsPDF-5_Demo.gif)


---

## 😫 The Problem

PowerPoint already offers multiple official ways to export:

* **File > Save As > Save as Adobe PDF**
* **File > Save As Adobe PDF** (choose folder each time)
* **File > Export > Create Adobe PDF**

With the modern file interface, each of these paths involves extra dialogs, folder navigation, and repeated clicks. If you only export occasionally, it’s fine — but if you do it often, the friction quickly adds up.

---

## 😍 The Agitation

Think about it: You just wrapped up a presentation. You’re ready to share it. But instead of sending it right away, you’re stuck digging through menus and dialogs — over and over again.

It feels like a small thing… but repeated a hundred times, it adds up to hours wasted every month. Hours you could spend on actual work.

---

## 🚀 The Solution

The **SaveAsPDF-5 Add-in** adds a simple macro to PowerPoint that does exactly what you wish Microsoft had built in:

👉 **Save your presentation as a PDF directly into the same folder as your PPTX — instantly.**

No browsing. No menus. Just one click.

---

## ✨ Key Features

* One-click **Save As PDF** in the same folder as your `.pptx`
* Simple installer with cleanup/uninstall support
* Works with PowerPoint 2010, 2013, 2016, 2019, 2021, and Office 365
* Fully trusted add-in — no scary security prompts

---

## 📦 How to Install

1. Download this repo (or clone it):

   ```sh
   git clone https://github.com/tsdiokno/SaveAsPDF-5.git
   ```

2. Run the installer:

   * Open a Command Prompt or PowerShell window in the folder
   * Run:

     ```sh
     .\install.bat
     ```
   * The script will:

     * Copy `SaveAsPDF-5.ppam` into your AddIns folder
     * Register it in the registry so it auto-loads

3. Restart PowerPoint. Done.

---

## 🗑 How to Uninstall

If you ever want to remove the add-in:

```sh
.\install.bat /uninstall
```

This will:

* Delete the `.ppam` file from your AddIns folder
* Remove all registry entries
* Restore PowerPoint to its original state

---

## 🔧 How to Use

* Add the **SaveAsPDF** macro to your Ribbon or Quick Access Toolbar (QAT)
* Next time you finish a deck, just click the button → PDF saved in the same folder

No menus. No hassle.

---

## 🖥 Compatibility

* Office 2010 (`14.0`)
* Office 2013 (`15.0`)
* Office 2016+ (`16.0`)
* Office 365

---

## 📜 License

MIT License – free to use, modify, and share.

---

## ⚡ Why You’ll Love It

If you’re constantly making decks and sending PDFs, this tool will save you **hours every month**. It’s fast, lightweight, and built for real users who just want less friction.

👉 Install it once. Enjoy it forever.
