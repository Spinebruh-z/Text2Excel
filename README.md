# Text to Excel Converter

A powerful Streamlit-based application that converts unstructured text data from various file formats into structured Excel spreadsheets using Google Gemini LLM for intelligent extraction.

## 📋 Features

- **AI-Powered Extraction**: Uses Google Gemini (gemini-2.0-flash-lite) to intelligently extract key-value pairs from unstructured text
- **Multiple File Format Support**: Process `.txt`, `.docx`, and `.pdf` files
- **Smart Text Processing**: 
  - UTF-8 encoding with NFKC Unicode normalization
  - Automatic text chunking with overlap for context preservation
  - Paragraph-based segmentation
- **Rich Data Extraction**: Extracts with detailed context:
  - Key-value pairs with confidence scores
  - Contextual comments explaining temporal context, format details, and analytical significance
  - Provenance tracking for data lineage
- **Batch Processing**: Process multiple files simultaneously with automatic document ID tracking
- **Interactive Data Editor**: Review and edit extracted data in real-time before exporting
- **Professional Excel Export**: Generate formatted Excel files with:
  - 3-column structure (key, value, comments)
  - Professional styling
  - Auto-adjusted column widths
  - Frozen header rows
- **CSV Export**: Alternative lightweight export format

## 🚀 Installation

### Prerequisites

- Python 3.11 or higher
- pip package manager
- **Google Gemini API Key** (Get from: https://aistudio.google.com/app/apikey)

### Setup

1. **Clone or download the project**

```bash
cd Text_2_Excel
```

2. **Create a virtual environment (recommended)**

```bash
# Windows
python -m venv t2e
t2e\Scripts\activate

# Linux/Mac
python -m venv t2e
source t2e/bin/activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

Required packages:
- `streamlit>=1.28.0` - Web UI framework
- `pandas>=2.0.0` - Data manipulation
- `openpyxl>=3.1.0` - Excel file handling
- `python-docx>=1.0.0` - Word document processing
- `PyPDF2>=3.0.0` - PDF text extraction
- `xlsxwriter>=3.0.0` - Advanced Excel formatting
- `google-genai>=0.3.0` - Google Gemini LLM integration

## 💻 Usage

### Running the Application

```bash
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`

### Step-by-Step Guide

#### 1. **Enter Your Gemini API Key** (Sidebar - Required)
   - In the left sidebar, locate the "🔑 Google Gemini API Key" section
   - Enter your API key in the password field
   - Don't have an API key? Get one free at: https://aistudio.google.com/app/apikey
   - The key is used securely and never stored

#### 2. **Configure Optional Settings** (Sidebar)
   - **Key Patterns (Optional)**: Enter custom keys to prioritize extraction
     - One key per line
     - Example: `Name`, `Email`, `Phone`, `Address`
     - Helpful when you know specific fields to extract
   
   - **Excel Options**:
     - ☑️ **Include Headers**: Add column headers (key, value, comments) to Excel
     - ☑️ **Auto-adjust Column Width**: Automatically size columns to fit content

#### 3. **Upload Your Files**
   - Click "Choose text files" button or drag-and-drop files
   - **Supported formats**: `.txt`, `.docx`, `.pdf`
   - **Batch processing**: Upload multiple files at once
   - Each file gets a unique document ID (D1, D2, D3, etc.)
   - View uploaded files and their sizes in the expandable section

#### 4. **Extract Data with AI**
   - Click the **"🚀 Extract Data"** button
   - The app processes your files using Gemini LLM:
     - Text is normalized and chunked for optimal processing
     - Each chunk is analyzed by Gemini to extract facts
     - Results include key-value pairs with rich contextual comments
   - Processing time depends on file size and complexity
   - Progress is shown with a spinner animation

#### 5. **Review and Edit Extracted Data**
   - After extraction, the data appears in an **interactive table**
   - **Edit any cell**: Click on a cell to modify its content
   - **Add rows**: Click "+ Add row" at the bottom of the table
   - **Delete rows**: Select rows and press Delete key
   - **Reorder columns**: Drag column headers to rearrange
   - The table shows:
     - **key**: The data field name (e.g., "person_name", "birth_date")
     - **value**: The extracted value
     - **comments**: Detailed context including temporal info, format notes, and analytical significance
     - **source_file**: Which file the data came from

#### 6. **Export Your Results**
   - Two export options available:
   
   **📥 Download Excel** (Recommended)
   - Professional formatted `.xlsx` file
   - Includes only 3 columns: key, value, comments
   - Auto-adjusted column widths
   - Frozen header row for easy scrolling
   - Professional styling with borders and colors
   - Timestamp in filename: `Output_YYYYMMDD_HHMMSS.xlsx`
   
   **📥 Download CSV**
   - Plain text `.csv` format
   - Compatible with any spreadsheet software
   - Lightweight and easy to process programmatically
   - Same 3-column structure
   - Timestamp in filename: `Output_YYYYMMDD_HHMMSS.csv`

#### 7. **View Statistics** (Optional)
   - Expand the "📈 Data Statistics" section to see:
     - **Total Records**: Number of key-value pairs extracted
     - **Total Columns**: Number of data fields
     - **Files Processed**: Count of source files analyzed

## 📁 Project Structure

```
Text_2_Excel/
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
├── README.md                   # Documentation
├── utils/                      # Utility modules
│   ├── __init__.py
│   ├── text_extractor.py      # Text extraction from files
│   ├── parser.py              # Data parsing and extraction logic
│   ├── excel_generator.py     # Excel file generation
│   └── logger.py              # Logging configuration
├── logs/                       # Application logs (auto-created)
└── ref/                        # Reference documents
```

## 🔧 Configuration

### How AI Extraction Works

**Intelligent Processing Pipeline:**

1. **Text Normalization**
   - UTF-8 encoding with NFKC Unicode normalization
   - Removal of non-printable characters
   - Whitespace normalization while preserving paragraph structure

2. **Smart Chunking**
   - Text split into ~2000 character chunks with 200 character overlap
   - Ensures context is preserved across chunk boundaries
   - Each chunk tracked with document ID and paragraph index

3. **AI Analysis (Gemini LLM)**
   - Each chunk analyzed by Google Gemini (gemini-2.0-flash-lite)
   - Low temperature (0.1) for deterministic, accurate extraction
   - Extracts ALL factual information as key-value pairs
   - Generates detailed contextual comments for each extraction

4. **Data Consolidation**
   - Results from all chunks merged
   - 6-field LLM output (key, value, raw_value, comments, provenance, confidence) consolidated to 3 columns
   - Document provenance tracked throughout

### What Gets Extracted

The AI automatically identifies and extracts:

- **Personal Information**: Names, birth dates/places, age, nationality, languages
- **Education**: Degrees, institutions, graduation years, grades, majors
- **Professional**: Job titles, companies, employment periods, salaries, responsibilities
- **Skills**: Certifications, technical skills, soft skills, competencies
- **Contact**: Email addresses, phone numbers, physical addresses, LinkedIn profiles
- **Achievements**: Awards, publications, projects, recognitions
- **Financial**: Salaries, amounts, currency values
- **Temporal**: Dates, periods, durations
- **Locations**: Cities, states, countries, addresses

### Example Output Format

For unstructured text like: *"John graduated from MIT in 2010 with a Computer Science degree and now earns $85,000 annually."*

**Extracted Data:**

| key | value | comments |
|-----|-------|----------|
| person_name | John | First name mentioned in context. Likely refers to primary subject of the document. |
| education_institution | MIT | Massachusetts Institute of Technology. Prestigious technical university indicating strong academic background. |
| education_year | 2010 | Graduation year. Indicates approximately 14 years of potential career experience as of 2024. |
| education_degree_field | Computer Science | Technical degree in computing. Relevant for technology sector positions and indicates analytical skillset. |
| salary | $85,000 | Annual compensation in USD. Formatted with currency symbol. Represents competitive mid-level compensation for CS graduates with ~14 years experience. |

## 🛠️ Technical Details

### Dependencies

- **streamlit**: Web interface framework
- **pandas**: Data manipulation and DataFrame operations
- **openpyxl**: Excel file creation and formatting
- **xlsxwriter**: Advanced Excel styling and export
- **python-docx**: Word document (.docx) text extraction
- **PyPDF2**: PDF text extraction
- **google-genai**: Google Gemini LLM integration for AI-powered extraction

### Architecture

1. **Text Extraction Layer** (`text_extractor.py`)
   - Handles `.txt`, `.docx`, `.pdf` formats
   - UTF-8 encoding and NFKC normalization
   - Non-printable character removal

2. **Preprocessing & Chunking** (`parser.py`)
   - Paragraph-based segmentation
   - Overlapping chunks for context preservation
   - Document ID and offset tracking

3. **AI Extraction Engine** (`parser.py`)
   - Google Gemini LLM integration
   - Structured prompt engineering
   - JSON schema validation
   - Confidence scoring

4. **Data Consolidation** (`parser.py`)
   - 6-field to 3-column transformation
   - Metadata consolidation into comments
   - Provenance tracking

5. **Export Layer** (`excel_generator.py`)
   - Professional Excel formatting
   - 3-column structure enforcement
   - Auto-width adjustment and styling

## 📝 Use Cases

### Use Case 1: Resume/CV Processing

**Input**: Upload multiple resume files (PDF, DOCX)

**Output**: Structured Excel with:
- Names, contact information
- Education history with institutions and degrees
- Work experience with companies and dates
- Skills and certifications
- Rich comments explaining context and relevance

### Use Case 2: Biographical Data Extraction

**Input**: Unstructured biographical narratives

**Output**: Organized data points:
- Birth dates and places
- Education milestones
- Career progression
- Achievements and awards
- Detailed temporal and contextual notes

### Use Case 3: Document Batch Analysis

**Input**: Multiple business documents (invoices, reports, emails)

**Output**: Consolidated spreadsheet with:
- All extracted facts across documents
- Source file tracking (source_file column)
- Context-rich comments for each data point
- Ready for further analysis or database import

## 🐛 Troubleshooting

### Common Issues

**Issue**: "Invalid Gemini API key" error
- **Solution**: 
  - Get a valid API key from https://aistudio.google.com/app/apikey
  - Ensure you're copying the entire key without extra spaces
  - Check if your API key has the necessary permissions
  - Try generating a new API key if the current one doesn't work

**Issue**: Empty or minimal comments column
- **Solution**: 
  - The AI generates comments based on available context
  - More detailed source text produces richer comments
  - Try uploading documents with complete sentences and paragraphs
  - Fragments or lists may produce shorter comments

**Issue**: Cannot extract text from PDF
- **Solution**: 
  - Ensure PDF contains actual text (not scanned images)
  - Image-based PDFs require OCR preprocessing (not currently supported)
  - Try converting to .txt or .docx format first

**Issue**: Missing expected data
- **Solution**: 
  - Add custom keys in the sidebar to prioritize specific fields
  - Ensure source text contains the expected information
  - Check if text formatting is preventing extraction
  - Review the extracted data - it may be under a different key name

**Issue**: Application is slow or times out
- **Solution**: 
  - Large files take longer to process
  - Each chunk requires an API call to Gemini
  - Try splitting very large files into smaller documents
  - Check your internet connection (API calls require connectivity)

**Issue**: Excel file formatting looks wrong
- **Solution**: 
  - Ensure you're using Excel 2007+ or compatible software
  - Try the CSV export as an alternative
  - Check "Include Headers" option in sidebar
  - Verify "Auto-adjust Column Width" is enabled

**Issue**: Application won't start
- **Solution**: 
  - Verify Python version: `python --version` (requires 3.11+)
  - Reinstall dependencies: `pip install -r requirements.txt`
  - Ensure virtual environment is activated
  - Check for port conflicts (default: 8501)

### Logs

Check the `logs/` directory for detailed error messages and debugging information. Each run creates timestamped log files with full extraction details.

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📄 License

This project is provided as-is for educational and commercial use.

## 📧 Support

For issues, questions, or suggestions:
- Check the logs directory for error details
- Review the troubleshooting section
- Refer to the assignment documentation in the `ref/` folder

## 🔄 Version History

**v2.0.0** (Current - December 2025)
- 🤖 **AI-Powered Extraction**: Integrated Google Gemini LLM (gemini-2.0-flash-lite)
- 📝 **Rich Contextual Comments**: Detailed analytical comments for each extracted field
- 🔍 **Smart Text Processing**: UTF-8/NFKC normalization, intelligent chunking
- 📊 **3-Column Output**: Streamlined key-value-comments structure per industry standards
- 🎯 **Provenance Tracking**: Document ID and paragraph indexing
- 🔐 **Secure API Integration**: Session-based API key handling
- ✨ **Enhanced UI**: Interactive data editor with real-time preview

**v1.0.0** (Legacy)
- Basic regex-based extraction
- Multi-format file support
- Simple key-value extraction

## 🎯 Future Enhancements

- [ ] **OCR Integration**: Support for image-based PDFs using OCR
- [ ] **Advanced Deduplication**: Smart merging of duplicate extractions across chunks
- [ ] **Key Canonicalization**: Automatic synonym mapping (e.g., "DOB" → "birth_date")
- [ ] **Confidence Filtering**: Auto-accept high confidence (≥0.9), flag low confidence (<0.8)
- [ ] **Provenance Highlighting**: Click a row to highlight source text in original document
- [ ] **Vector Store Integration**: FAISS-based storage for full metadata and embeddings
- [ ] **LangChain Orchestration**: Pipeline visualization and parallel processing
- [ ] **Type Normalization**: Automatic date, currency, and phone number standardization
- [ ] **Template Library**: Pre-built extraction templates for common document types
- [ ] **API Endpoints**: RESTful API for programmatic access
- [ ] **Database Export**: Direct export to PostgreSQL, MySQL, MongoDB
- [ ] **Multi-language Support**: Extraction from non-English documents
- [ ] **Cloud Integration**: Azure Blob, AWS S3, Google Drive connectors

---

**Made with ❤️ using Streamlit**
