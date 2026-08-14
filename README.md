
# Devgent 🤖

> **An interactive desktop AI developer agent that connects LLMs directly to your local workspace.** Automate document generation, run terminal commands, host local web servers, and maintain project context safely.

![Devgent Preview](demo.mp4)

## ⚡ Quick Reference (TL;DR)

### 3-Minute Setup
```bash
git clone https://github.com/your-username/devgent.git
cd devgent
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
python devgent.py
```

### Key Highlights
* **Inline Confirmation Cards**: Requires explicit `Yes/No` approval before touching files or running commands.
* **Full Document Suite**: Native generation of `.pptx` (with charts), `.pdf`, `.xlsx`, and `.docx`.
* **Workspace Server**: Built-in HTTP server to host and preview web builds immediately.
* **Dual LLM Engines**: Supports both Google Gemini API (BYOK) and local Ollama instances.

---

## 📖 Complete Documentation

### 🛠️ Detailed Installation & Environment Setup

#### Prerequisites
* **Python 3.9+** installed on your system.
* Operating System: Windows 11 / 10, macOS, or Linux.

#### Step-by-Step Instructions
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/your-username/devgent.git
   cd devgent
   ```
2. **Configure Virtual Environment**:
   * **Windows**:
     ```bash
     python -m venv venv
     venv\Scripts\activate
     ```
   * **macOS / Linux**:
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```
3. **Install Dependencies**:
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

---

## 💡 How It Works & Architecture

Devgent connects LLMs to your local environment using an asynchronous background execution pipeline.

* **Lazy-Loading Core**: Heavy document libraries and SDKs are loaded in background threads upon startup. This ensures the GUI launches instantly without lag.
* **Persistent Workspace Memory**: Automatically maintains a `.devgent_memory.json` file inside your active project folder. It remembers user preferences, active tech stacks, and architecture decisions across separate chat sessions.
* **File Content Parser**: Extracts text and structural data from PDF, Word, Excel, and code source files to supply context directly into prompt windows.

---

## ⚙️ Configuration & Engine Options

When launching Devgent, configure your preferred backend via the sidebar:

1. **Google Gemini Engine**:
   * Select **Google Gemini** in the engine menu.
   * Enter your Gemini API key. 
   * *Note: API keys are stored locally at `~/.devgent/config.json` and are never committed or exposed in builds.*

2. **Ollama (Local LLM)**:
   * Select **Ollama** in the engine menu.
   * Ensure your local Ollama server is running (default endpoint: `http://localhost:11434`).

3. **Workspace Path**:
   * Click **Set Directory** to designate your working root directory. All server hosting, file reading, and script execution are strictly locked to this root folder.

---

## 🛡️ Security & Approvals

Safety is built directly into the execution flow. Devgent will **never** perform file modifications or execute shell scripts silently. 

Whenever the agent attempts to:
* Create or modify a document (`.pdf`, `.pptx`, `.docx`, `.xlsx`, `.txt`)
* Perform batch file creation
* Execute terminal commands (`run_command`)
* Spin up local web servers (`start_dev_server`)

An **Interactive Permission Card** renders directly in the chat stream requiring you to click **Approve** or **Deny** before the task executes.

---

## 📜 License

Distributed under the MIT License. See `LICENSE` for more information.
