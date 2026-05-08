# PaperReady | Local AI Architect 

**PaperReady** is a professional-grade, hybrid AI application designed to transform local LLM inference into high-quality document automation. It bridges the gap between CLI-based AI interaction and real-time visual output in Microsoft Office and a web-based Canvas.

![PaperReady Banner](https://img.shields.io/badge/Version-3.0-4285F4?style=for-the-badge)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=for-the-badge&logo=windows)
![Privacy](https://img.shields.io/badge/Privacy-100%25_Offline-34A853?style=for-the-badge)

---

##  Key Capabilities

*   **Live COM Automation**: Watch the AI build PowerPoint slides and Word documents directly inside the actual Microsoft Office applications in real-time.
*   **Interactive Canvas**: A persistent web dashboard that renders live Markdown summaries and professional **Graphviz (DOT)** architectural diagrams.
*   **Workspace RAG**: Retrieval-Augmented Generation that allows the AI to "read" and summarize your local folders, including `.pptx`, `.docx`, and `.txt` files.
*   **Data Sovereignty**: 100% offline. All inference and file generation happen on your local hardware. Zero data leakage.
*   **Smart Web Search**: Automated Selenium-based scraping to fetch real-world data and images for your documents when needed.

---

## Technology Stack

- **Model**: Microsoft Phi-3 Mini 4K Instruct (GGUF)
- **Inference**: `llama-cpp-python` (CPU-optimized)
- **Automation**: `pywin32` / Win32 COM
- **Visuals**: `Viz.js` / Graphviz
- **Dashboard**: Python `http.server` (Asynchronous state sync)

---

##  Installation & Setup

PaperReady features a **Smart Launcher** designed to handle all dependencies automatically, even on "resource-less" machines without C++ compilers or CMake.

### 1. Requirements
- **Windows 10/11**
- **Python 3.10 or 3.11** (Recommended)
- **Microsoft Office** (PowerPoint & Word)

### 2. Deployment
Open your terminal and run:

```powershell
# Clone the repository
git clone https://github.com/Bhargavmupparisetty/paperReady.git
cd paperReady

# Launch the Smart Architect
.\Run_PaperReady.bat
