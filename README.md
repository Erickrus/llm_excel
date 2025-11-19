# Excel-Python LLM Bridge

This project enables you to run LLMs, Web Crawlers, and File Readers directly inside Microsoft Excel. It uses a robust file-based communication system with Base64 encoding to ensure full support for **UTF-8 text**.

## Prerequisites

1.  **Microsoft Excel** (Windows recommended).
2.  **Python 3.8+** installed.
3.  **Ollama** installed and running (for the LLM feature).

---

## Part 1: Python Setup

1.  **Install Dependencies**
    Open your terminal or command prompt and install the required libraries:
    ```bash
    pip install -r requirements.txt
    ```

2.  **Prepare the Model**
    Make sure you have the specific model pulled in Ollama (matches the model defined in `functions.py`):
    ```bash
    ollama pull gemma3n:latest
    ```

3.  **Save the Python Scripts**
    Save the following two files in the same folder on your computer (e.g., `C:\Projects\ExcelLLM\`):
    *   `llm_main.py` (The main listener script)
    *   `functions.py` (The logic definitions)

4.  **Run the Listener**
    Start the Python agent. It must be running in the background to process Excel requests.
    ```bash
    python llm_main.py
    ```
    *You should see: "Agent started. Watching folder: .../Temp/llm_temp"*

---

## Part 2: Excel Setup (VBA)

1.  **Open the Visual Basic Editor**
    *   Open Excel.
    *   Press **`Alt + F11`** to open the VBA Editor.
    *   If nothing appears, go to **Insert > Module**.

2.  **Import the Code**
    *   Copy the entire VBA code.
    *   Paste it into the Module window.
    *   Save the file as an **Excel Macro-Enabled Workbook (.xlsm)**.

3.  **Assign a Shortcut (Recommended)**
    *   Close the VBA window and go back to Excel.
    *   Press **`Alt + F8`**.
    *   Select `TriggerProcessing`.
    *   Click **Options...**
    *   Assign a shortcut key (e.g., `Ctrl + q`).

---

## Part 3: How to Use

1.  **Write a Formula**
    In any cell, write one of the supported functions. You can use direct text or cell references.

    *   **LLM Chat:**
        ```excel
        =LLM("Explain quantum physics briefly")
        ```
        *Or refer to a cell:* `=LLM(A1)`

    *   **Web Crawl:**
        ```excel
        =CRAWL("https://en.wikipedia.org/wiki/Python_(programming_language)")
        ```

    *   **Read File:**
        ```excel
        =READ("C:\Users\Name\Documents\notes.txt")
        ```

2.  **Trigger the Process**
    *   **Select the cell(s)** containing the formula.
    *   Run the macro (press the Shortcut you assigned, e.g., `Ctrl + q`, or use `Alt + F8` -> Run).

3.  **View Results**
    *   Excel will create a placeholder text (e.g., `[LLM trigger for...]`).
    *   Wait a few seconds.
    *   The Python script processes the request.
    *   The result will appear automatically in the **cell to the right** of your formula.


