# 📝 Math PDF Maker (AlvinRocks Edition)

**"The Bridge Between AI Math and Professional Documents."**

Most math tools make you choose between the speed of **Markdown/LaTeX** and the formatting of **Microsoft Word**. **Math PDF Maker** lets you have both. It’s a dedicated workspace for competitive math enthusiasts, students, and educators to create high-quality PDFs from scratch or from screenshots.

---

## **🚀 Key Features**

### **1. AI-Powered OCR & Vision**
Stop re-typing equations. 
* **Local OCR**: Uses `rapidocr_onnxruntime` for instant, offline text extraction from your clipboard.
* **Gemini Vision**: For complex formulas or handwritten notes, the integrated Gemini 1.5 Flash model translates images into structured LaTeX instantly.

### **2. Professional Export Engine**
We don't just "paste" math as images. 
* **Native Equations**: Generates true Microsoft Word equation objects using `math2docx`.
* **PDF Mastery**: Automatically triggers a background Word instance to export your document as a high-fidelity PDF.

### **3. Workflow Optimization**
* **Auto-Install**: The script checks your Python environment and installs all necessary libraries (Pillow, PyMuPDF, `math2docx`, etc.) automatically.
* **Presets**: Switch between **"IOQM Worksheet,"** **"Simple Note,"** or **"Raw Text"** templates to handle different document styles instantly.
* **Draft System**: Built-in autosave and draft management so you never lose your progress.

---

## **⌨️ Keyboard Shortcuts**

| Shortcut | Action |
| :--- | :--- |
| **Ctrl + V** | **Smart Paste**: Triggers Local OCR on clipboard images. |
| **Ctrl + Shift + V** | **AI Paste**: Sends clipboard image to Gemini for LaTeX extraction. |
| **Ctrl + S** | **Export**: Instantly generates the Word/PDF document. |
| **Ctrl + N** | **New Draft**: Clears the editor and starts a fresh project. |
| **Ctrl + Q** | **Quick Save**: Saves current progress to the Drafts tab. |

---

## **🛠️ Installation & Setup**

1.  **Clone the Repository**:
    ```bash
    git clone https://github.com/AlvinKirath/Math-PDF-Maker.git
    ```
2.  **Run the Script**:
    ```bash
    python math_pdf_maker.py
    ```
    *Note: The script will automatically prompt you to install dependencies if they are missing.*

3.  **Gemini API (Optional)**:
    To use the AI-Vision feature, obtain a free API key from [Google AI Studio](https://aistudio.google.com/) and enter it into the UI when prompted.

---

## **⚠️ Requirements**
* **Windows OS**: Required for the PDF export feature (uses Word COM API).
* **Microsoft Word**: Must be installed for `.docx` to `.pdf` conversion.
* **Python 3.10+**

---

### **Final Note**
This application was built to solve the friction of documenting complex math. It uses a combination of local OCR for speed and Cloud AI for accuracy, ensuring that no matter how messy your source material is, your final PDF looks professional.
