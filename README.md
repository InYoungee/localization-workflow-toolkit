# Localization Workflow Toolkit
A comprehensive Python toolkit designed to automate common localization tasks for language professionals. This collection of scripts streamlines word counting, cost estimation, and quality assurance across multiple file formats.

## 🎯 Overview
This toolkit was created by a localization project manager to address real-world workflow challenges in the gaming and software localization industry. Whether you're managing translation projects, performing QA, or analyzing content volume, these tools help automate repetitive tasks and improve accuracy.

## ✨ Features
### Word Counter Tools

<ul>
<li><b>Multi-Format Counter</b> (multi_format_counter.py) - Count words across various file formats (json, xml, pdf, docx) in a single operation</li>
<li><b>GUI Counter</b> (gui_counter.py) - Desktop application with file uploader, visual word count analysis, and Excel report export</li>
<li><b>Streamlit Counter</b> (streamlit_counter.py) - Web-based interface for word count analysis with CSV export</li>
<li><b>Advanced Streamlit Counter</b> (streamlit_counter_advanced_counter.py) - Enhanced web interface with target language-specific cost calculations and Excel export</li>
</ul>

### QA Tools
<ul>
<li>QA Auditor (qa_tools/qa_auditor.py) - Automated quality checks comparing source and target files</li>
<ul>
<li>Missing placeholder detection</li>
<li>String length validation for UI constraints</li>
<li>HTML tag corruption checks</li>
<li>Excel report generation with flagged issues</li>
</ul>
</ul>

### Excel Processing Tools
<ul>
<li>Sample File Generator (excel_counter/create_sample_excel_files.py) - Creates sample Excel files for testing (game strings: skill descriptions, dialogues, UI strings)</li>
<li>Excel Column Counter (excel_counter/excel_column_counter_with_tag_stripping.py) - Counts words from source columns while ignoring markup tags, with cost calculation and Excel export</li>
</ul>

## 🚀 Getting Started
### Prerequisites
```bash
Python 3.7+
```
### Installation
```bash
# Clone this repository:

git clone https://github.com/InYoungee/localization-workflow-toolkit.git
cd localization-workflow-toolkit

# Install required dependencies:

pip install -r requirements.txt
```

## 📖 Usage
### Word Counting
<b>Multi-Format Counter</b>
```bash
python multi_format_counter.py
```
Batch process and count words across multiple file formats (e.g., .txt, .json, .xml/xlf, .docx, .pdf) in a single operation. Ideal for quickly analyzing diverse localization file types without format-specific tools.

<b>GUI Counter (Desktop Application)</b>
```bash
python gui_counter.py
````
Upload files through the interface, view analysis, and export Excel reports.

<b>Streamlit Web Counter</b>
```bash
streamlit run streamlit_counter.py
````
Access the web interface at `http://localhost:8501` to upload files and export CSV reports.

<b>Advanced Counter with Cost Calculation</b>
```bash
streamlit run streamlit_counter_advanced_counter.py
````
Includes target language-specific pricing and Excel export functionality.

### Quality Assurance
<b>Run QA Auditor</b>
```bash
python qa_tools/qa_auditor.py
````
Checks target files against source files and generates detailed Excel reports with flagged issues.

### Excel Processing
<b>Generate Sample Files</b>
```bash
python excel_counter/create_sample_excel_files.py
```
Creates sample game localization files (KO→EN & JP) with string IDs and info comments.

<b>Count Excel Column Words</b>
```bash
python excel_counter/excel_column_counter_with_tag_stripping.py
```
Analyzes source column word counts while stripping markup tags, includes cost estimation.

## 🗂️ Project Structure
```
localization-workflow-toolkit/
├── word_counter/
│   ├── test_files                  # Sameple files
│   ├── gui_counter.py
│   ├── multi_format_counter.py
│   ├── streamlit_advanced_counter.py
│   └── streamlit_counter.py
├── qa_tools/
│   ├── qa_auditor.py
│   ├── qa_en-US.json               # Sameple file
│   └── qa_ko-KR.json               # Sameple file
├── excel_counter/
│   ├── sample_excel_files          # Sameple files
│   ├── create_sample_excel_files.py
│   └── excel_column_counter_with_tag_stripping.py
├── README.md
├── .gitignore
└── requirements.txt
```
## 🎮 Use Cases
<ul>
<li><b>Project Managers</b>: Quickly estimate translation costs and volume across multiple formats</li>
<li><b>QA Engineers</b>: Automate quality checks for common localization issues</li>
<li><b>Freelance Translators</b>: Calculate word counts and generate client reports</li>
<li><b>Game Localization Teams</b>: Process multi-column Excel files with tagged content</li>
</ul>

## 🛠️ Technologies
<ul>
<li>Python 3.x</li>
<li>Streamlit (Web interfaces)</li>
<li>Tkinter (GUI applications)</li>
<li>pandas (Data processing)</li>
<li>openpyxl (Excel operations)</li>
</ul>

## 🔧 Technical Highlights

### Pattern Matching with Regular Expressions
This toolkit leverages regex extensively for:
- **Tag Stripping**: Removes HTML/XML tags while preserving text content for accurate word counts
- **Placeholder Detection**: Identifies patterns like `{0}`, `%s`, `${variable}` in QA checks
- **Format Recognition**: Automatically detects file formats and content structures
- **String Validation**: Checks for malformed tags and syntax errors

Example regex patterns used:
- HTML tag removal: `<[^>]+>`

## 📝 License
This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing
Contributions, issues, and feature requests are welcome! Feel free to check the issues page.

## 👤 Author
Inyoung Kim
<ul>
<li>LinkedIn: https://www.linkedin.com/in/inyoungee/</li>
<li>GitHub: @InYoungee</li>
</ul>

## 🙏 Acknowledgments
Built from real-world experience in gaming localization to help the broader localization community work more efficiently.


*If you find this toolkit helpful, please consider giving it a ⭐ on GitHub!*

## Demo
![streamlit WC counter](https://github.com/InYoungee/portfolio/blob/main/assets/img/portfolio/streamlit_wordcounter.gif)
