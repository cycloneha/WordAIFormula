\# Word AI Formula Assistant 🤖📐



A VSTO-based Word Add-in designed to streamline the workflow of converting AI-generated LaTeX formulas (from ChatGPT, Claude, DeepSeek, etc.) into editable \*\*MathType\*\* equations in Microsoft Word.



It supports one-click conversion of LaTeX code (`$$...$$`, `\\\[...\\]`) and intelligently handles mixed text-formula layouts.



\## ✨ Features



\* \*\*Smart Cleaning\*\*: Automatically removes AI-generated Markdown markers and non-standard LaTeX delimiters.

\* \*\*Mixed Layout Support\*\*: Uses a Regex-based lexer to accurately separate plain text from formulas, allowing for perfect pasting of entire paragraphs.

\* \*\*Deep MathType Integration\*\*: Leverages underlying COM interfaces to render LaTeX as editable MathType objects.

\* \*\*Robust Design\*\*:

&nbsp;   \* \*\*Strategy Pattern\*\*: Automatically detects and adapts to different MathType versions (6.9/7.0/Wiris).

&nbsp;   \* \*\*Performance Optimization\*\*: utilizing `ScreenUpdating` locks to prevent screen flickering during bulk processing.



\## 🛠️ Tech Stack



\* \*\*Framework\*\*: .NET Framework 4.8 / VSTO (C#)

\* \*\*Patterns\*\*: Strategy Pattern, Factory Pattern

\* \*\*IDE\*\*: Visual Studio 2022



\## 🚀 Getting Started



\### Prerequisites

1\.  Windows 10/11

2\.  Microsoft Word (2016 or later)

3\.  \*\*MathType 6.9 or later is required\*\* (Core dependency)



\### Installation

1\.  Go to the \[Releases](../../releases) page on the right side.

2\.  Download the latest `Installer.zip`.

3\.  Unzip the file and run `setup.exe`.

4\.  Click "Install". The plugin button will appear in the Word ribbon.



\### Usage

1\.  Copy any text containing formulas from an AI chatbot (e.g., `According to $$E=mc^2$$, energy and mass...`).

2\.  Click the \*\*"AI Formula"\*\* button in the Word Add-ins tab.

3\.  The text will be pasted, and formulas will be automatically converted.



\## ⚠️ Disclaimer



This project is an open-source utility tool released under the MIT License.

1\.  \*\*Dependency\*\*: This plugin relies on the user's local installation of MathType for rendering. This project \*\*does not contain, distribute, or crack\*\* any MathType binary files. Please support genuine software.

2\.  \*\*Liability\*\*: The author is not responsible for any document damage or data loss caused by the use of this tool. Please backup important documents before use.

3\.  \*\*Trademark\*\*: MathType is a trademark of Wiris. This project is not affiliated with Wiris.



\## 📄 License



MIT License

