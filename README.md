# 🛠️ AI-SolidWorks


<div align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-blue.svg" alt="Python 3.8+"/>
  <img src="https://img.shields.io/badge/SolidWorks-Compatible-red.svg" alt="SolidWorks Compatible"/>
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License MIT"/>
  <img src="https://img.shields.io/badge/AI-Powered-purple.svg" alt="AI Powered"/>
</div>

<div dir="rtl">

## 🇮🇷 نسخه فارسی

**SoliPy** یک رابط هوشمند بین کاربر و نرم‌افزار SolidWorks است که امکان کنترل و ایجاد اشکال با استفاده از **دستورات زبان طبیعی فارسی** را فراهم می‌کند.

### ✨ ویژگی‌های اصلی

- **🗣️ ارتباط با زبان طبیعی**: ارسال دستورات طراحی با زبان فارسی ساده
- **🤖 پشتیبانی از هوش مصنوعی**: استفاده از مدل‌های پیشرفته هوش مصنوعی برای درک دستورات
- **📝 تبدیل خودکار به اسکریپت**: تولید اسکریپت‌های VBS برای SolidWorks
- **🎨 رابط کاربری آسان**: محیط کاربری ساده و کاربرپسند با تم تاریک

### 🚀 نحوه استفاده

1. **نصب و راه‌اندازی**:
   ```bash
   git clone https://github.com/your-username/solipy.git
   cd solipy
   pip install -r requirements.txt
   python sw_api_panel.py
   ```

2. **تنظیم API**: 
   - کلید API خود را در فایل `.env` یا از طریق دکمه "تنظیمات API" وارد کنید
   - امکان استفاده از سرویس‌های OpenAI یا OpenRouter

3. **ارسال دستورات**: 
   - مثال: "یک دایره بکش"
   - مثال: "یک مستطیل به ضلع ۴۰ متر بکش"
   - مثال: "یک خط بکش"

### 📂 ساختار فایل‌ها

- `sw_api_panel.py`: برنامه اصلی با رابط کاربری گرافیکی
- `scripts/`: پوشه حاوی اسکریپت‌های VBS برای SolidWorks
  - `connect_to_sw.vbs`: اسکریپت اتصال به SolidWorks
  - `draw_circle.vbs`: اسکریپت نمونه برای رسم دایره
  - `create_sketch.vbs`: اسکریپت ایجاد اسکچ اصلی
  - `create_sketch_from_input.vbs`: اسکریپت پارامتریک برای ایجاد اشکال
  - `create_extrude.vbs`: اسکریپت برای اکسترود کردن اشکال

### 📋 نیازمندی‌ها

- **Python 3.8+**
- **SolidWorks** (نصب شده روی سیستم)
- **کلید API** از OpenAI یا OpenRouter
- **سیستم عامل**: ویندوز

</div>

---

## 🇬🇧 English Version

**SoliPy** is an intelligent interface between users and SolidWorks software that enables control and shape creation using **natural language commands** in both Persian and English.

### ✨ Key Features

- **🗣️ Natural Language Communication**: Send design commands using simple language
- **🤖 AI-Powered**: Utilizes advanced AI models to understand commands
- **📝 Automatic Script Generation**: Generates VBS scripts for SolidWorks
- **🎨 User-Friendly Interface**: Simple and intuitive UI with dark theme

### 🚀 How to Use

1. **Installation**:
   ```bash
   git clone https://github.com/your-username/solipy.git
   cd solipy
   pip install -r requirements.txt
   python sw_api_panel.py
   ```

2. **API Setup**:
   - Enter your API key in the `.env` file or through the "API Settings" button
   - Compatible with OpenAI or OpenRouter services

3. **Sending Commands**:
   - Example: "Draw a circle"
   - Example: "Create a rectangle with 40 meters sides"
   - Example: "Draw a line"

### 📂 File Structure

- `sw_api_panel.py`: Main program with graphical user interface
- `scripts/`: Folder containing VBS scripts for SolidWorks
  - `connect_to_sw.vbs`: Script for connecting to SolidWorks
  - `draw_circle.vbs`: Sample script for drawing a circle
  - `create_sketch.vbs`: Script for creating main sketch
  - `create_sketch_from_input.vbs`: Parametric script for creating shapes
  - `create_extrude.vbs`: Script for extruding shapes

### 📋 Requirements

- **Python 3.8+**
- **SolidWorks** (installed on the system)
- **API key** from OpenAI or OpenRouter
- **Operating System**: Windows

## 📸 Screenshots

<div align="center">
  <p><em>Main application interface with SolidWorks integration</em></p>
  <p><em>Command input and response flow</em></p>
</div>

## 🔄 Workflow

1. **User Input**: Enter a command in natural language
2. **AI Processing**: The command is processed by AI to understand intent
3. **Script Generation**: A VBS script is automatically generated
4. **Execution**: The script is executed in SolidWorks
5. **Result**: The requested operation is performed in SolidWorks

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details. "# SolidWorks-MCP-Server" 
"# SolidWorks-MCP-Server" 
