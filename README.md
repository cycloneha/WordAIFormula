# Word AI Formula Assistant 🤖📐

A VSTO-based Microsoft Word Add-in designed to streamline the workflow of converting AI-generated LaTeX formulas (from ChatGPT, Claude, DeepSeek, etc.) into editable **MathType** equations in Word.

The add-in supports one-click conversion of LaTeX code ($$...$$, \[...\]) and intelligently handles mixed text–formula layouts, enabling seamless pasting of entire AI-generated paragraphs into Word.

## ✨ Features

* **Smart Cleaning**: Automatically removes AI-generated Markdown markers and non-standard LaTeX delimiters to ensure compatibility with MathType.
* **Mixed Layout Support**: Uses a Regex-based lexer to accurately separate plain text from mathematical formulas, allowing entire paragraphs containing interleaved text and equations to be pasted and converted correctly.
* **Deep MathType Integration**: Leverages MathType’s exposed COM automation interfaces to render LaTeX formulas as fully editable MathType objects in Microsoft Word.
* **Robust Design**:
    * **Strategy Pattern**: Automatically detects and adapts to different MathType versions (6.9 / 7.0 / Wiris editions).
    * **Performance Optimization**: Uses Word `ScreenUpdating` locks to prevent UI flickering during bulk conversions.

## 🛠️ Tech Stack

* **Framework**: .NET Framework 4.8
* **Platform**: VSTO (Visual Studio Tools for Office)
* **Language**: C#
* **Design Patterns**: Strategy Pattern, Factory Pattern
* **IDE**: Visual Studio 2022

## 🚀 Getting Started

### Prerequisites
* Windows 10 / 11
* Microsoft Word 2016 or later
* **MathType 6.9 or later (Required Dependency)**

> [!IMPORTANT]
> MathType must be installed locally and properly licensed. This add-in **does not** function without MathType.

### Installation
1.  Go to the [Releases](../../releases) page.
2.  Download the latest `Installer.zip`.
3.  Unzip the package and run `setup.exe`.
4.  Click **Install**.
5.  The plugin button will appear in the Word ribbon under the **Add-ins** tab.

### Usage
1.  Copy any text containing LaTeX formulas from an AI chatbot (e.g., `According to $$E=mc^2$$, energy and mass...`).
2.  Open Microsoft Word.
3.  Click the **“aiLatex”** button in the Word **Add-ins** tab.
4.  The text will be pasted, and all formulas will be automatically converted into editable MathType equations.

---

## 🔍 Scope of Functionality
* This project **does not** implement a mathematical typesetting or rendering engine.
* All formula rendering, layout, and editing capabilities are provided **exclusively by MathType**.
* The purpose of this add-in is to improve user workflow efficiency by automating repetitive copy-and-paste and conversion operations.
* This project is **not** a replacement for MathType.

## ⚠️ Disclaimer
* **Open Source**: This project is an open-source utility tool released under the MIT License.
* **Dependency Notice**: This plugin relies on the user's local installation of MathType. It does not include, distribute, embed, or crack any MathType binary files. Users must obtain MathType in accordance with its official license terms.
* **Liability**: To the maximum extent permitted by law, the author shall not be liable for any data loss, document corruption, or other damages. **Users are strongly advised to back up important documents before use.**
* **Trademark Notice**: MathType is a registered trademark of Wiris. This project is not affiliated with, endorsed by, or sponsored by Wiris.

## ⚖️ Legal & Compliance Notice
* This project **does not** bypass, disable, modify, or interfere with any MathType licensing, activation, or copy-protection mechanisms.
* All MathType-related functionality is executed through user-level automation of COM interfaces exposed by MathType itself.
* If MathType is not properly licensed or its trial period has expired, related features will fail as enforced by MathType.
* This project **does not** include, redistribute, reverse-engineer, or modify any MathType binaries, components, or proprietary assets.

## 📄 License
This project is released under the [MIT License](LICENSE).
