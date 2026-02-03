# ✈️ AirSial SMS Form Scanner & Dashboard Generator

## 🚀 Live Demo
[![Hugging Face Spaces](https://img.shields.io/badge/🤗-Hugging%20Face%20Spaces-blue)](https://huggingface.co/spaces)
[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://airsial-sms-scanner.streamlit.app)

## 📋 Features

### 1. **SMS Form Scanner**
- Upload multiple images/PDFs (50+ at once)
- Automatic OCR with PaddleOCR (free, unlimited)
- Handles printed & handwritten text
- Extracts all form fields automatically
- Generates audit-ready Excel with 5 sheets

### 2. **Monthly Dashboard Generator**
- Automatic KPI calculation
- Interactive charts & visualizations
- Risk heat maps
- CAP tracking with traffic lights
- AI-powered safety insights

### 3. **Audit-Ready Excel Template**
- **Sheet 1**: Raw SMS Data (locked for integrity)
- **Sheet 2**: Standardized Lists (dropdown controls)
- **Sheet 3**: CAP Tracker (auto-calculated overdue)
- **Sheet 4**: Monthly Dashboard (auto-linked charts)
- **Sheet 5**: Audit Evidence Log

## 🛠️ Quick Start

### Local Installation
```bash
# Clone repository
git clone https://github.com/yourusername/AirSial-SMS-Scanner.git
cd AirSial-SMS-Scanner

# Install dependencies
pip install -r requirements.txt

# Run application
streamlit run sms_scanner_app.py
