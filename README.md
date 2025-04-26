# Script Collector → Word  
Generate a single **.docx** that contains every C# script in your Unity project – ready to share with teammates, teachers, reviewers or clients.

![Unity](https://img.shields.io/badge/Unity-2020.3%2B-black?logo=unity)
![License](https://camo.githubusercontent.com/3318cb25b502f7a6e23c3574c614fe969dc4d8316a6e71ab55bf62f0a65c675e/68747470733a2f2f696d672e736869656c64732e696f2f6769746875622f6c6963656e73652f5365616565657353616e2f53696d706c65466f6c64657249636f6e)

---

## Features
| ✔ | What it does |
|---|--------------|
| **One-click export** | Collects all `*.cs` files (optionally only selected sub-folders) and saves them into a single Word document. |
| **Keeps file names** | Each script is preceded by its file name for quick navigation. |
| **Editor-only** | Lives entirely inside the Unity Editor – no code touches your build. |
| **Folder filters** | Include / exclude `Editor` folders and choose sub-folders manually. |

---

## Quick Start

1. **Install NuGetForUnity**
   *Window → Package Manager → Add package from Git URL…*
```jsonc
  "https://github.com/GlitchEnzo/NuGetForUnity.git?path=/src/NuGetForUnity"
  ```
2. **Install Open XML SDK**  
NuGet → Manage NuGet Packages → search **DocumentFormat.OpenXml** → **Install**.
  ```jsonc
  "DocumentFormat.OpenXml"
  ```
3. **Install ScriptCollectorToWord**  
Window → Package Manager → **Add package from Git URL…**  
  ```jsonc
  "https://github.com/Je1rei/ScriptCollector.git"
  ```

4. **Generate the document**  
Tools → Generate Word from Scripts → choose folders → **Generate Word file**.

---

## ⚙️ Options
- **Advanced options** – tick first-level sub-folders manually
- **Include all sub-folders** – scan entire tree  
- **Exclude “Editor” folders** – ignore editor-only code  

---

## 📄 Output
- Word 2007+ compatible `.docx`  
- UTF-8 preserved, ready for diff/grep  
- File names use **Heading 1** style → _Insert → Table of Contents_ works instantly

---

## 🗺️ Roadmap
- Syntax-highlighted export (Word styles)  
- Extra file types (`.shader`, `.uss`, `.asmdef` …)  

---

## 🤝 Contributing
Pull requests and issues are welcome – please open an issue first for large changes.

---

## 📜 License

This project is licensed under the **MIT License**. See the full text below.

---

## MIT License
MIT License

Copyright (c) 2025 Je1rei

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

