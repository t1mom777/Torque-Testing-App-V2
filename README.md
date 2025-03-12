# 🚀 Modern Torque Testing Application

## 📖 Overview
Welcome to the **Modern Torque Testing Application**—your ultimate tool for precise torque testing, intuitive data management, and smart reporting powered by cutting-edge AI technology!

## 🌟 Key Features

### 🔧 Torque Testing
- ⚡ **Real-time Torque Measurement** via serial connections
- 🟢 Visual feedback for **instant pass/fail results** (green for pass, red for fail)

### 📊 Data Management
- ✏️ Easily add, edit, or delete torque specifications
- 📐 Customizable allowance ranges for precise control
- 🖼️ **AI-powered Data Extraction** from images—effortlessly upload customer and equipment info through:
  - 📸 Images from your device
  - 📋 Clipboard screenshots
  - 📹 Webcam captures

_The AI automatically recognizes and organizes customer data into the correct fields, saving you valuable time!_

### 💾 Data Storage
- 🗃️ Robust integration with DuckDB for efficient local data handling
- 🔐 Secure storage of torque tables, raw test data, and app settings

### 📋 Report Generation
- 📈 Generate sleek and professional reports in **Excel and PDF formats**
- ✨ **Highly Customizable Excel Templates** to match your brand and reporting needs

## 🤖 AI Integration
- 🌐 **Advanced OpenAI Integration** to automatically extract accurate data from uploaded images, significantly reducing manual entry!

## 📥 Installation

### 📌 Requirements
- Python 3.x
- PyQt6
- DuckDB
- PySerial
- openpyxl
- pywin32 (required for Excel to PDF conversion on Windows)
- OpenAI API

### 🛠️ Setup
Install all dependencies quickly using:
```bash
pip install PyQt6 duckdb pandas openpyxl pyserial pywin32 openai
```

## 🚀 Quick Start
Launch your torque testing journey by running:
```bash
python main.py
```

### 🧪 Conducting Tests
1. 🔽 Select the torque entry from the dropdown.
2. 🔌 Choose your serial port connection.
3. ▶️ Click `Begin Test` to start.

### ⚙️ Customization
- 🖌️ Fully customize your torque data entries.
- 🔑 Configure OpenAI API settings.
- 📁 Tailor your Excel and PDF report export preferences.

## 📂 Project Structure
```
.
├── db_handler_local.py          # Database handler
├── main.py                      # App entry point
├── modern_torque_app.py         # Core application
├── serial_reader.py             # Serial communication
├── openai_handler.py            # OpenAI image extraction
├── editor.html                  # Report template editor
├── data.duckdb                  # Local database file
└── templates/
    └── summary_template.xlsx    # Customizable Excel template
```

## 🔗 Dependencies
- Python 3.x
- PyQt6
- DuckDB
- pandas
- OpenAI API
- pyserial
- openpyxl
- pywin32 (Windows only)

## 🤝 Contributing
We welcome your contributions! Submit your pull requests or report any issues you encounter.

## 📜 License
This project is proudly open-source under the **MIT License**.
