 🏭 Cement Quality Analytics Dashboard

> **The most advanced cement QC analytics app with OCR, camera capture, and beautiful modern UI**

![Version](https://img.shields.io/badge/version-4.0-blue)
![Python](https://img.shields.io/badge/python-3.9+-blue)
![Streamlit](https://img.shields.io/badge/streamlit-1.28+-red)

---

## ✨ What's New in v4.0

### 🎨 Beautiful Modern UI
- **Dark glassmorphism theme** with gradient backgrounds
- **Animated KPI cards** with glow effects
- **Responsive layout** optimized for desktop and tablet
- **Color-coded status badges** (Excellent/Good/Warning/Danger)

### 📸 Image & Camera Support (NEW!)
| Feature | Description |
|---------|-------------|
| **📸 Image Upload** | Upload screenshots/photos of control registers — OCR extracts data automatically |
| **📷 Camera Capture** | Take live photos from plant floor, QC lab, or control room screens |
| **🔍 OCR Processing** | Uses Tesseract OCR with OpenCV preprocessing for accurate text extraction |
| **✅ Auto-Validation** | Validates extracted data structure before analysis |

### 📋 All Input Methods
| Method | Supported Formats |
|--------|-------------------|
| 📂 **File Upload** | Excel (.xlsx/.xls) with multi-sheet, CSV (.csv) |
| 📋 **Copy & Paste** | Tab, comma, semicolon, space-separated |
| 📄 **Word Document** | .docx files with data tables |
| 📸 **Image / Screenshot** | PNG, JPG, JPEG — OCR extraction |
| 📷 **Camera Capture** | Live photo with automatic OCR |

---

## 🚀 Installation

### Prerequisites
- Python 3.9 or higher
- Tesseract OCR (for image/camera support)

### Step 1: Install Tesseract OCR

**Windows:**
```bash
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
# Install and add to PATH
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
sudo apt-get install libtesseract-dev
```

**macOS:**
```bash
brew install tesseract
```

### Step 2: Install Python Dependencies

```bash
pip install -r requirements_v4.txt
```

Or manually:
```bash
pip install streamlit pandas numpy plotly openpyxl xlrd python-docx opencv-python pillow pytesseract
```

### Step 3: Run the App

```bash
streamlit run cement_quality_app_v4.py
```

Open browser at: `http://localhost:8501`

---

## 📖 How to Use

### 📸 Using Image Upload (Screenshot of Control Register)

1. Select **"📸 Image / Screenshot"** in the sidebar
2. Upload a clear photo of your control register screen or printed sheet
3. The app shows a preview of your image
4. Click **"🔍 Extract Data with OCR"**
5. The app processes the image and extracts tabular data
6. Review the extracted preview and proceed to analysis

**Tips for best OCR results:**
- Ensure good lighting and no glare
- Text should be horizontal and clearly visible
- Crop to include only the data table
- Use higher resolution images (min 800px width)

### 📷 Using Camera Capture (Live Photo)

1. Select **"📷 Camera Capture"** in the sidebar
2. Allow camera access when prompted
3. Point camera at the control register screen/sheet
4. Click the capture button
5. Click **"🔍 Extract Data from Photo"**
6. Data auto-extracts and analysis begins

**Best for:** Plant floor operators, QC lab technicians, control room operators

### 📂 Using Excel Multi-Sheet

1. Select **"📂 Upload File"**
2. Upload multi-sheet Excel workbook
3. App detects all sheets automatically
4. Choose sheet from dropdown
5. Other sheets listed in expandable section

---

## 🔬 Analysis Modules

| Tab | Features |
|-----|----------|
| **📊 Overview** | 6 KPI cards, data preview, statistics, animated quality grade |
| **📈 Trends** | Interactive trends with USL/LSL control limits, mineral phase stacked area |
| **🔬 Capability** | Cp/Cpk table with color-coded status, CV% variability chart |
| **🎯 Coating & Strength** | Coating index risk assessment, burnability, 28-day strength prediction |
| **💰 Operations** | Fuel efficiency, raw material utilization, cost impact calculator |
| **📄 Export** | Downloadable CSV/text reports, actionable recommendations |

---

## 🧪 Key Formulas

### Coating Index
```
Coating Index = (Liquid Phase % × C3A / C4AF) / 10
```
| Value | Risk | Action |
|-------|------|--------|
| < 1.55 | Low | Continue normal monitoring |
| 1.55 – 1.65 | Moderate | Inspect kiln regularly |
| > 1.65 | High | Immediate kiln inspection |

### 28-Day Strength Prediction
```
Strength = 35 + (C3S-45)×0.85 + (C2S-25)×0.25 + (C3A-7)×0.5
           - max(0, F-CaO-1)×3 + (LSF-90)×0.3
```

---

## 🌐 Deploy on Streamlit Cloud

1. Push to GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with GitHub
4. Select your repository
5. Click **Deploy**

**Note:** For OCR/camera features on Streamlit Cloud, you may need to:
- Use `packages.txt` to install system dependencies (tesseract-ocr)
- Or deploy on a server with Tesseract pre-installed

---

## 🛠️ Project Structure

```
cement-quality-analytics/
├── cement_quality_app_v4.py    # Main application (v4.0)
├── requirements_v4.txt        # Python dependencies
├── sample_clinker_data.csv    # Test data
├── test_multi_sheet.xlsx      # Multi-sheet test file
├── test_word_doc.docx         # Word test file
├── README.md                  # Documentation
└── .gitignore                 # Git ignore rules
```

---

## 🐛 Troubleshooting

| Problem | Solution |
|---------|----------|
| Tesseract not found | Install Tesseract OCR and ensure it's in system PATH |
| OCR poor accuracy | Use clearer images, better lighting, crop to table only |
| Camera not working | Ensure browser has camera permissions, use HTTPS |
| Multi-sheet not detected | Use `.xlsx` format (not `.xls`) |
| Material not detected | Check column headers match expected format |
| App crashes | Check traceback in error message for details |

---

## 📄 License

MIT License — Copyright (c) 2026 FAUJI CEMENT COMPANY LIMITED (NIZAMPUR)

---

<div align="center">

### 🏭 Built for Cement Quality Excellence

*Upload anything — Excel, Word, screenshots, or live photos. Get instant analytics.*

</div>
