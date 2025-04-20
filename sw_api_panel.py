#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SoliPy - SolidWorks API Panel
-----------------------------
An AI-powered interface for SolidWorks using VBScript

Version: 1.0.0
Author: SoliPy Team
License: MIT
Repository: https://github.com/sina-salim/solipy
"""

import os
import sys
import time
import json
import glob
import shutil
import logging
import uuid
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import threading
import queue
import datetime
import requests
from typing import Dict, List, Any, Optional, Tuple

# تنظیم کدگذاری برای خروجی
if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# تنظیم لاگینگ
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("sw_api_panel.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("SolidWorksPanel")

# تنظیمات مسیرها
SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
HISTORY_DIR = os.path.join(SCRIPTS_DIR, "history")
MAX_HISTORY = 20  # حداکثر تعداد اسکریپت‌های ذخیره شده در تاریخچه

# کلیدهای API و تنظیمات
API_KEY = ""
BASE_URL = "https://api.openai.com/v1/chat/completions"
API_MODEL = "gpt-4o-mini"

# خواندن کلید API از فایل .env یا doc.txt
try:
    if os.path.exists(".env"):
        with open(".env", "r") as f:
            for line in f:
                if line.startswith("OPENAI_API_KEY="):
                    API_KEY = line.strip().split("=", 1)[1].strip('"\'')
                if line.startswith("OPENAI_BASE_URL="):
                    BASE_URL = line.strip().split("=", 1)[1].strip('"\'')
                if line.startswith("OPENAI_MODEL="):
                    API_MODEL = line.strip().split("=", 1)[1].strip('"\'')
    if not API_KEY and os.path.exists("doc.txt"):
        with open("doc.txt", "r") as f:
            for line in f:
                if "api key" in line.lower():
                    parts = line.split("=", 1)
                    if len(parts) == 2:
                        API_KEY = parts[1].strip().strip('"\'')
                if "base_url" in line.lower():
                    parts = line.split("=", 1)
                    if len(parts) == 2:
                        BASE_URL = parts[1].strip().strip('"\'')
                if "model" in line.lower():
                    parts = line.split("=", 1)
                    if len(parts) == 2:
                        API_MODEL = parts[1].strip().strip('"\'')
except Exception as e:
    logger.error(f"خطا در خواندن کلید API: {e}")

# اطمینان از وجود پوشه‌های مورد نیاز
os.makedirs(SCRIPTS_DIR, exist_ok=True)
os.makedirs(HISTORY_DIR, exist_ok=True)

# کپی اسکریپت نمونه به عنوان اسکریپت فعلی اگر وجود نداشته باشد
current_script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
sample_script_path = os.path.join(SCRIPTS_DIR, "create_simple_part.vbs")
if not os.path.exists(current_script_path) and os.path.exists(sample_script_path):
    shutil.copy2(sample_script_path, current_script_path)
    logger.info(f"اسکریپت نمونه به عنوان اسکریپت فعلی کپی شد.")

class APITester:
    """کلاس تست کننده API"""
    
    @staticmethod
    def test_api_connection(api_key: str, base_url: str, api_model: str) -> Tuple[bool, str]:
        """تست اتصال به API

        Args:
            api_key: کلید API
            base_url: آدرس API
            api_model: مدل هوش مصنوعی

        Returns:
            (موفقیت, پیام): وضعیت تست و پیام مربوطه
        """
        if not api_key or not base_url or not api_model:
            return False, "لطفاً تمام فیلدها را پر کنید"
        
        try:
            # ایجاد هدرهای درخواست
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
                "HTTP-Referer": "https://solipy.app"
            }
            
            # ایجاد پیام بسیار کوتاه و ساده برای تست
            payload = {
                "model": api_model,
                "messages": [
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": "Say 'Test connection successful' if you can read this."}
                ],
                "temperature": 0.2,
                "max_tokens": 20
            }
            
            # ارسال درخواست با timeout کوتاه
            logger.info(f"در حال تست اتصال به API: {base_url}")
            response = requests.post(base_url, headers=headers, json=payload, timeout=10)
            
            if response.status_code == 200:
                try:
                    response_data = response.json()
                    content = response_data.get("choices", [{}])[0].get("message", {}).get("content", "")
                    logger.info(f"تست اتصال به API موفقیت‌آمیز بود. پاسخ: {content[:50]}...")
                    return True, "اتصال موفقیت‌آمیز"
                except Exception as e:
                    logger.warning(f"تست اتصال موفق بود اما خطا در پردازش پاسخ: {e}")
                    return True, "اتصال موفق، خطا در پردازش پاسخ"
            else:
                logger.error(f"خطا در تست اتصال به API: {response.status_code} - {response.text}")
                return False, f"خطا: {response.status_code}"
                
        except requests.exceptions.Timeout:
            logger.error("زمان پاسخگویی API به پایان رسید")
            return False, "خطا: زمان پاسخ به پایان رسید"
        except requests.exceptions.ConnectionError:
            logger.error("خطا در اتصال به سرور API")
            return False, "خطا: مشکل در اتصال به سرور"
        except Exception as e:
            logger.error(f"خطای کلی در تست API: {e}")
            return False, f"خطا: {str(e)[:40]}"

class APISettingsDialog(tk.Toplevel):
    """دیالوگ تنظیمات API برای وارد کردن API key، base URL و model"""
    
    def __init__(self, parent, api_key="", base_url="", api_model=""):
        """راه‌اندازی دیالوگ تنظیمات API

        Args:
            parent: پنجره والد
            api_key: کلید API فعلی
            base_url: آدرس API فعلی
            api_model: مدل هوش مصنوعی فعلی
        """
        super().__init__(parent)
        self.title("تنظیمات API")
        self.geometry("500x350")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # رنگ‌های تم تیره
        self.bg_color = "#1A1A2E"  # پس زمینه تیره
        self.card_color = "#0F3460"  # کارت‌های تیره با کمی رنگ
        self.text_color = "#FFFFFF"  # متن سفید
        self.input_bg = "#1E2A4A"  # پس زمینه فیلد ورودی
        self.button_bg = "#6A2CB3"  # بنفش تیره
        self.button_active_bg = "#8A42D8"  # بنفش روشن‌تر
        
        # تنظیم رنگ پس‌زمینه
        self.config(bg=self.bg_color)
        
        self.api_key = api_key
        self.base_url = base_url
        self.api_model = api_model
        
        self.result = {"api_key": self.api_key, "base_url": self.base_url, "api_model": self.api_model}
        
        # ایجاد فریم اصلی
        main_frame = tk.Frame(self, bg=self.bg_color, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # فیلدهای ورودی
        tk.Label(main_frame, text="API Key:", anchor=tk.W, bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 10)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.api_key_entry = tk.Entry(main_frame, width=50, bg=self.input_bg, fg=self.text_color, insertbackground="white", relief=tk.FLAT)
        self.api_key_entry.grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        self.api_key_entry.insert(0, self.api_key)
        
        tk.Label(main_frame, text="Base URL:", anchor=tk.W, bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 10)).grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.base_url_entry = tk.Entry(main_frame, width=50, bg=self.input_bg, fg=self.text_color, insertbackground="white", relief=tk.FLAT)
        self.base_url_entry.grid(row=1, column=1, sticky=tk.W, pady=(0, 5))
        self.base_url_entry.insert(0, self.base_url)
        
        tk.Label(main_frame, text="مدل:", anchor=tk.W, bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 10)).grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.model_entry = tk.Entry(main_frame, width=50, bg=self.input_bg, fg=self.text_color, insertbackground="white", relief=tk.FLAT)
        self.model_entry.grid(row=2, column=1, sticky=tk.W, pady=(0, 5))
        self.model_entry.insert(0, self.api_model)
        
        # دکمه تست API
        test_frame = tk.Frame(main_frame, bg=self.bg_color)
        test_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 5))
        
        self.test_btn = tk.Button(test_frame, text="تست اتصال به API", 
                               command=self._test_api,
                               bg=self.button_bg,
                               fg="white",
                               activebackground=self.button_active_bg,
                               activeforeground="white",
                               font=("Segoe UI", 10),
                               bd=0,
                               relief=tk.FLAT,
                               padx=10,
                               pady=5)
        self.test_btn.pack(side=tk.LEFT, padx=2)
        
        self.test_result_label = tk.Label(test_frame, text="", anchor=tk.W, bg=self.bg_color, fg=self.text_color)
        self.test_result_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # بخش توضیحات
        tk.Label(main_frame, text="راهنما:", anchor=tk.W, bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 10)).grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        help_text = (
            "برای استفاده از OpenAI:\n"
            "- Base URL: https://api.openai.com/v1/chat/completions\n"
            "- مدل: gpt-4o, gpt-3.5-turbo\n\n"
            "برای استفاده از OpenRouter:\n"
            "- Base URL: https://openrouter.ai/api/v1/chat/completions\n"
            "- مدل: google/gemini-1.5-pro"
        )
        help_label = tk.Label(main_frame, text=help_text, anchor=tk.W, justify=tk.LEFT, bg=self.bg_color, fg=self.text_color)
        help_label.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # دکمه‌های پایین
        buttons_frame = tk.Frame(main_frame, bg=self.bg_color)
        buttons_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        save_btn = tk.Button(buttons_frame, text="ذخیره", 
                           command=self._on_save,
                           bg=self.button_bg,
                           fg="white",
                           activebackground=self.button_active_bg,
                           activeforeground="white",
                           font=("Segoe UI", 10),
                           bd=0,
                           relief=tk.FLAT,
                           padx=15,
                           pady=5)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(buttons_frame, text="انصراف", 
                             command=self._on_cancel,
                             bg=self.button_bg,
                             fg="white",
                             activebackground=self.button_active_bg,
                             activeforeground="white",
                             font=("Segoe UI", 10),
                             bd=0,
                             relief=tk.FLAT,
                             padx=15,
                             pady=5)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # حذف تغییرات در صورت بستن پنجره
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        # تمرکز بر فیلد اول
        self.api_key_entry.focus_set()
        
    def _test_api(self):
        """تست اتصال به API"""
        self.test_result_label.config(text="در حال تست...", foreground="white", background=self.bg_color)
        self.update_idletasks()
        
        api_key = self.api_key_entry.get().strip()
        base_url = self.base_url_entry.get().strip()
        api_model = self.model_entry.get().strip()
        
        success, message = APITester.test_api_connection(api_key, base_url, api_model)
        
        if success:
            self.test_result_label.config(text=f"✓ {message}", foreground="#4CAF50", background=self.bg_color)
        else:
            self.test_result_label.config(text=f"✗ {message}", foreground="#F44336", background=self.bg_color)
    
    def _on_save(self):
        """ذخیره تنظیمات و بستن دیالوگ"""
        self.result = {
            "api_key": self.api_key_entry.get().strip(),
            "base_url": self.base_url_entry.get().strip(),
            "api_model": self.model_entry.get().strip()
        }
        self.destroy()
    
    def _on_cancel(self):
        """انصراف از تغییرات و بستن دیالوگ"""
        self.result = None
        self.destroy()
    
    @staticmethod
    def show_dialog(parent, api_key="", base_url="", api_model=""):
        """نمایش دیالوگ و دریافت نتیجه

        Args:
            parent: پنجره والد
            api_key: کلید API فعلی
            base_url: آدرس API فعلی
            api_model: مدل هوش مصنوعی فعلی

        Returns:
            dict: دیکشنری حاوی تنظیمات API یا None در صورت انصراف
        """
        dialog = APISettingsDialog(parent, api_key, base_url, api_model)
        parent.wait_window(dialog)
        return dialog.result

class SolidWorksScriptGenerator:
    """کلاس تولید کننده اسکریپت‌های VBS برای SolidWorks"""
    
    def __init__(self, api_key: str = "", base_url: str = "", api_model: str = ""):
        """راه اندازی تولید کننده اسکریپت

        Args:
            api_key: کلید API برای استفاده از سرویس هوش مصنوعی
            base_url: آدرس API
            api_model: مدل هوش مصنوعی
        """
        self.api_key = api_key if api_key else API_KEY
        self.api_url = base_url if base_url else BASE_URL
        self.api_model = api_model if api_model else API_MODEL
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
            "HTTP-Referer": "https://solipy.app"
        }
    
    def generate_script(self, query: str) -> Tuple[bool, str, Optional[str]]:
        """تولید اسکریپت VBS بر اساس درخواست کاربر

        Args:
            query: متن درخواست کاربر

        Returns:
            (موفقیت, پیام, مسیر_اسکریپت): وضعیت تولید اسکریپت، پیام و مسیر فایل اسکریپت تولید شده
        """
        try:
            # لاگ کردن درخواست
            logger.info(f"درخواست جدید: {query}")
            
            if not self.api_key:
                return False, "کلید API تنظیم نشده است. لطفاً کلید API را در فایل .env یا doc.txt تنظیم کنید.", None
            
            # ایجاد درخواست API
            system_prompt = """You are an expert in SolidWorks automation with VBScript. 
Your task is to generate VBScript code that can automate SolidWorks operations.
You MUST understand user instructions in both English and Persian (Farsi) language.
When user instructions are in Persian, you should understand words like "بکش", "دایره", "خط", "مستطیل", etc.
Keep your responses focused only on the VBScript code without any explanations.

IMPORTANT: Always structure your script in this sequence:
1. Start with 'Option Explicit'
2. Include connection code that connects to SolidWorks first:

Dim swApp
On Error Resume Next
WScript.Echo "در حال اتصال به SolidWorks..."
Set swApp = GetObject(, "SldWorks.Application")
If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo "SolidWorks در حال اجرا نیست. تلاش برای اجرای SolidWorks..."
    Set swApp = CreateObject("SldWorks.Application")
    If Err.Number <> 0 Then
        WScript.Echo "خطا در اتصال به SolidWorks: " & Err.Description
        WScript.Quit(1)
    End If
End If
swApp.Visible = True
WScript.Echo "اتصال به SolidWorks با موفقیت انجام شد."
On Error Goto 0

3. Create or open a document
4. Implement the requested feature or operation
5. Include proper error handling

The script must run independently and include all necessary code to connect to SolidWorks, 
not relying on any external functions or files.
Always end with a success message and return 0 exit code on success.

PERSIAN COMMANDS GLOSSARY:
- "دایره بکش" or "یک دایره بکش" or "رسم دایره" = Draw a circle
- "مستطیل بکش" or "یک مستطیل بکش" = Draw a rectangle
- "خط بکش" = Draw a line
- "اکسترود کن" = Extrude
- "برش بزن" = Cut
- "ذخیره کن" = Save
"""
            payload = {
                "model": self.api_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Create a VBScript to automate the following SolidWorks task: {query}. "
                                              f"Only respond with the complete VBScript code without any explanations."}
                ],
                "temperature": 0.2,
                "max_tokens": 2000
            }
            
            # ارسال درخواست به API
            logger.info(f"ارسال درخواست به API... ({self.api_url}, {self.api_model})")
            response = requests.post(self.api_url, headers=self.headers, json=payload)
            
            if response.status_code != 200:
                logger.error(f"خطا در پاسخ API: {response.status_code} - {response.text}")
                return False, f"خطا در درخواست API: {response.status_code}", None
            
            # استخراج کد اسکریپت از پاسخ
            response_data = response.json()
            script_content = response_data['choices'][0]['message']['content'].strip()
            
            # حذف بخش‌های توضیحی احتمالی و نگه داشتن فقط کد
            if "```vb" in script_content or "```vbs" in script_content:
                script_parts = script_content.split("```")
                for part in script_parts:
                    if part.startswith("vb") or part.startswith("vbs"):
                        script_content = part[part.index("\n")+1:]
                    elif not part.startswith("`") and len(part.strip()) > 10:
                        script_content = part
            
            # ایجاد نام فایل با تاریخ و زمان
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            script_name = f"sw_script_{timestamp}.vbs"
            script_path = os.path.join(HISTORY_DIR, script_name)
            
            # ذخیره اسکریپت در فایل
            with open(script_path, "w", encoding='utf-8') as f:
                f.write(script_content)
            
            logger.info(f"اسکریپت ایجاد شد: {script_path}")
            
            # حذف اسکریپت‌های قدیمی اگر تعداد آنها از حد مجاز بیشتر شد
            self._cleanup_history()
            
            # کپی اسکریپت به پوشه اصلی اسکریپت‌ها
            current_script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
            shutil.copy2(script_path, current_script_path)
            
            return True, "اسکریپت با موفقیت ایجاد شد.", script_path
            
        except Exception as e:
            logger.error(f"خطا در تولید اسکریپت: {e}")
            return False, f"خطا در تولید اسکریپت: {str(e)}", None
    
    def _cleanup_history(self):
        """حذف اسکریپت‌های قدیمی اگر تعداد آنها از حد مجاز بیشتر شد"""
        try:
            # دریافت لیست همه اسکریپت‌ها
            script_files = glob.glob(os.path.join(HISTORY_DIR, "sw_script_*.vbs"))
            script_files.sort()  # مرتب‌سازی بر اساس نام (تاریخ و زمان)
            
            # حذف اسکریپت‌های قدیمی
            if len(script_files) > MAX_HISTORY:
                files_to_remove = script_files[0:len(script_files) - MAX_HISTORY]
                for file_path in files_to_remove:
                    try:
                        os.remove(file_path)
                        logger.info(f"اسکریپت قدیمی حذف شد: {file_path}")
                    except Exception as e:
                        logger.error(f"خطا در حذف فایل {file_path}: {e}")
        
        except Exception as e:
            logger.error(f"خطا در پاکسازی تاریخچه: {e}")
    
    def execute_script(self, script_path: str) -> Tuple[bool, str, str]:
        """اجرای اسکریپت VBS

        Args:
            script_path: مسیر فایل اسکریپت

        Returns:
            (موفقیت, پیام, خروجی): وضعیت اجرا، پیام و خروجی اسکریپت
        """
        try:
            if not os.path.exists(script_path):
                return False, f"فایل اسکریپت وجود ندارد: {script_path}", ""
            
            # اجرای اسکریپت با تنظیم encoding=None برای دریافت خروجی به صورت bytes
            result = subprocess.run(["cscript", "//NoLogo", script_path], 
                                   capture_output=True, text=False, check=False)
            
            exit_code = result.returncode
            
            # تبدیل خروجی با مدیریت خطای کدگذاری
            try:
                output = result.stdout.decode('utf-8', errors='replace')
            except Exception as e:
                logger.warning(f"خطا در تبدیل خروجی: {e}")
                output = str(result.stdout)
                
            try:
                error = result.stderr.decode('utf-8', errors='replace')
            except Exception as e:
                logger.warning(f"خطا در تبدیل خطا: {e}")
                error = str(result.stderr)
            
            if exit_code == 0:
                logger.info(f"اسکریپت با موفقیت اجرا شد: {script_path}")
                return True, "اسکریپت با موفقیت اجرا شد.", output
            else:
                logger.error(f"خطا در اجرای اسکریپت {script_path}: {error or output}")
                return False, f"خطا در اجرای اسکریپت (کد خروج: {exit_code})", error or output
            
        except Exception as e:
            logger.error(f"خطا در اجرای اسکریپت: {e}")
            return False, f"خطا در اجرای اسکریپت: {str(e)}", ""

# کلاس دیباگر اسکریپت 
class ScriptDebugger:
    """کلاس دیباگر اسکریپت برای شناسایی و رفع باگ‌های VBScript با کمک LLM"""
    
    def __init__(self, api_key: str, base_url: str, api_model: str):
        """راه‌اندازی دیباگر اسکریپت

        Args:
            api_key: کلید API برای استفاده از LLM
            base_url: آدرس API
            api_model: مدل هوش مصنوعی
        """
        self.api_key = api_key
        self.api_url = base_url
        self.api_model = api_model
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
            "HTTP-Referer": "https://solipy.app"
        }
    
    def provide_user_guidance(self, user_query: str) -> Tuple[bool, str]:
        """ارائه راهنمایی به کاربر با استفاده از LLM برای سوالات مرتبط با دیباگ یا طراحی اسکریپت

        Args:
            user_query: سوال یا درخواست کاربر

        Returns:
            (موفقیت, پاسخ): وضعیت درخواست و پاسخ دریافتی
        """
        try:
            # ایجاد پرامپت برای LLM
            system_prompt = """You are an expert in VBScript programming for SolidWorks automation. 
You help users debug their SolidWorks scripts and provide guidance on script development.
Focus on providing practical, specific advice that users can immediately apply to fix their code or improve their scripts.
Your expertise includes:
1. SolidWorks API methods and best practices
2. VBScript syntax and common errors
3. Debugging techniques for automation scripts
4. Best practices for SolidWorks automation

Respond in Persian (Farsi) language with:
1. Clear, step-by-step advice for the user's specific question
2. Code examples when relevant
3. Explanations of common mistakes or misconceptions
"""
            
            payload = {
                "model": self.api_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_query}
                ],
                "temperature": 0.3,
                "max_tokens": 2000
            }
            
            # ارسال درخواست به API
            logger.info(f"ارسال درخواست راهنمایی به API... ({self.api_url}, {self.api_model})")
            response = requests.post(self.api_url, headers=self.headers, json=payload, timeout=30)
            
            if response.status_code != 200:
                logger.error(f"خطا در پاسخ API راهنمایی: {response.status_code} - {response.text}")
                return False, f"خطا در درخواست API راهنمایی: {response.status_code}"
            
            # استخراج پاسخ از LLM
            response_data = response.json()
            llm_response = response_data['choices'][0]['message']['content'].strip()
            
            return True, llm_response
        
        except Exception as e:
            logger.error(f"خطا در دریافت راهنمایی: {e}")
            return False, f"خطا در دریافت راهنمایی: {str(e)}"
    
    def debug_script(self, script_content: str, error_message: str) -> Tuple[bool, str, str]:
        """دیباگ اسکریپت VBS با استفاده از LLM

        Args:
            script_content: محتوای اسکریپت دارای خطا
            error_message: پیام خطای دریافت شده هنگام اجرا

        Returns:
            (موفقیت, اسکریپت_اصلاح_شده, توضیحات): وضعیت دیباگ، اسکریپت اصلاح شده و توضیحات
        """
        try:
            # ایجاد پرامپت برای LLM
            system_prompt = """You are an expert VBScript debugger for SolidWorks automation. 
Your task is to analyze the provided VBScript code and error message, then fix the issue.
Focus on common VBScript errors such as:
1. Syntax errors (missing parentheses, wrong variable names)
2. "Cannot use parentheses when calling a Sub" error - in VBScript function calls that return values use parentheses, but Sub calls do not use parentheses
3. Invalid characters or encoding issues
4. Incorrect method calls or parameters for SolidWorks API
5. Issues with object references or method parameters

Respond with:
1. The fixed script - provide the complete corrected script
2. A brief explanation of what you fixed and why
"""
            
            payload = {
                "model": self.api_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Debug this VBScript code for SolidWorks automation. Here's the script:\n\n```vbs\n{script_content}\n```\n\nHere's the error message:\n{error_message}\n\nPlease fix the code and explain what was wrong."}
                ],
                "temperature": 0.3,
                "max_tokens": 2500
            }
            
            # ارسال درخواست به API
            logger.info(f"ارسال درخواست دیباگ به API... ({self.api_url}, {self.api_model})")
            response = requests.post(self.api_url, headers=self.headers, json=payload, timeout=30)
            
            if response.status_code != 200:
                logger.error(f"خطا در پاسخ API دیباگ: {response.status_code} - {response.text}")
                return False, "", f"خطا در درخواست API دیباگ: {response.status_code}"
            
            # استخراج پاسخ از LLM
            response_data = response.json()
            llm_response = response_data['choices'][0]['message']['content'].strip()
            
            # جداسازی کد اصلاح شده و توضیحات
            fixed_script = ""
            explanation = ""
            
            if "```vbs" in llm_response or "```vbscript" in llm_response:
                # استخراج کد از بخش‌های کد در پاسخ
                code_parts = []
                for part in llm_response.split("```"):
                    if part.startswith("vbs") or part.startswith("vbscript"):
                        code_parts.append(part[part.index("\n")+1:])
                    elif not part.startswith("```") and not part.strip().startswith("vbs") and not part.strip().startswith("vbscript") and len(part.strip()) > 10:
                        explanation += part.strip() + "\n"
                
                if code_parts:
                    fixed_script = code_parts[0].strip()
            else:
                # اگر ساختار کد مشخص نشده، تلاش برای تشخیص بخش کد و توضیحات
                lines = llm_response.split("\n")
                is_code = False
                code_lines = []
                explanation_lines = []
                
                for line in lines:
                    if line.strip().startswith("Option Explicit") or line.strip().startswith("Dim ") or line.strip().startswith("Set "):
                        is_code = True
                        code_lines.append(line)
                    elif is_code and (line.strip() == "" or line.strip().startswith("'")):
                        code_lines.append(line)
                    elif is_code and ("=" in line or line.strip().startswith("If ") or line.strip().startswith("End ") or line.strip().startswith("WScript.")):
                        code_lines.append(line)
                    else:
                        is_code = False
                        explanation_lines.append(line)
                
                fixed_script = "\n".join(code_lines).strip()
                explanation = "\n".join(explanation_lines).strip()
            
            # اگر نتوانستیم کد را استخراج کنیم، از کل پاسخ استفاده کنیم
            if not fixed_script:
                explanation = llm_response
                return False, "", f"نتوانستم کد اصلاح شده را استخراج کنم. لطفاً پاسخ زیر را بررسی کنید:\n\n{explanation}"
            
            return True, fixed_script, explanation
        
        except Exception as e:
            logger.error(f"خطا در دیباگ اسکریپت: {e}")
            return False, "", f"خطا در دیباگ اسکریپت: {str(e)}"

class SolidWorksPanel:
    """پنل گرافیکی برای تعامل با SolidWorks از طریق اسکریپت‌های VBS"""
    
    def __init__(self, root):
        """راه‌اندازی پنل

        Args:
            root: ریشه برنامه Tkinter
        """
        self.root = root
        self.root.title("SolidWorks API Panel")
        self.root.geometry("1100x700")
        self.root.minsize(900, 650)
        
        # تنظیم رنگ‌ها و استایل کلی (Dark Wallet Theme)
        self.bg_color = "#1A1A2E"  # پس زمینه تیره
        self.sidebar_color = "#16213E"  # سایدبار تیره‌تر
        self.card_color = "#0F3460"  # کارت‌های تیره با کمی رنگ
        self.accent_color = "#E94560"  # رنگ تاکیدی قرمز-صورتی
        self.text_color = "#FFFFFF"  # متن سفید
        self.secondary_text = "#B2B1B9"  # متن ثانویه خاکستری روشن
        self.highlight_color = "#E94560"  # هایلایت همان رنگ تاکیدی
        
        # تعریف رنگ بنفش برای دکمه‌ها
        self.button_bg = "#6A2CB3"  # بنفش تیره
        self.button_active_bg = "#8A42D8"  # بنفش روشن‌تر
        self.button_pressed_bg = "#551D99"  # بنفش تیره‌تر
        
        # تنظیم رنگ پس‌زمینه اصلی
        self.root.config(bg=self.bg_color)
        
        # تنظیم استایل‌ها
        self._configure_styles()
        
        # ایجاد تولید کننده اسکریپت
        self.script_generator = SolidWorksScriptGenerator(API_KEY, BASE_URL, API_MODEL)
        
        # ایجاد دیباگر اسکریپت
        self.script_debugger = ScriptDebugger(API_KEY, BASE_URL, API_MODEL)
        
        # ایجاد صف برای ارتباط با ترد
        self.queue = queue.Queue()
        
        # لیست مسیرهای فایل‌های تاریخچه
        self.history_paths = []
        
        # ایجاد اجزای رابط کاربری
        self._create_ui()
        
        # بررسی تاریخچه
        self._update_history_list()
        
        # بررسی دوره‌ای صف
        self.root.after(100, self._process_queue)
    
    def _configure_styles(self):
        """تنظیم استایل‌های مختلف برای ویجت‌ها"""
        style = ttk.Style()
        
        # استایل کلی
        style.configure(".", 
                       font=("Segoe UI", 10),
                       background=self.bg_color,
                       foreground=self.text_color)
                       
        # روش اول: تلاش برای استفاده از استایل clam که بیشتر قابل سفارشی‌سازی است
        style.theme_use('clam')
        
        # دکمه‌های معمولی
        style.configure("TButton", 
                       padding=10,
                       relief="flat",
                       background=self.button_bg,
                       foreground="white",
                       borderwidth=0,
                       font=("Segoe UI", 10))
        
        style.map("TButton",
                 background=[('active', self.button_active_bg), ('pressed', self.button_pressed_bg)],
                 foreground=[('active', 'white'), ('pressed', 'white')])
        
        # دکمه اصلی
        style.configure("Primary.TButton", 
                       padding=10,
                       background=self.button_bg,
                       foreground="white",
                       borderwidth=0,
                       font=("Segoe UI", 10, "bold"))
        
        style.map("Primary.TButton",
                 background=[('active', self.button_active_bg), ('pressed', self.button_pressed_bg)],
                 foreground=[('active', 'white'), ('pressed', 'white')])
        
        # دکمه‌های سایدبار
        style.configure("Sidebar.TButton", 
                       padding=15,
                       background=self.button_bg,
                       foreground="white",
                       borderwidth=0,
                       font=("Segoe UI", 10))
        
        style.map("Sidebar.TButton",
                 background=[('active', self.button_active_bg), ('pressed', self.button_pressed_bg)],
                 foreground=[('active', 'white'), ('pressed', 'white')])
        
        # فریم‌ها
        style.configure("TFrame", background=self.bg_color)
        
        # فریم سایدبار
        style.configure("Sidebar.TFrame", background=self.sidebar_color)
        
        # فریم کارت
        style.configure("Card.TFrame", 
                       background=self.card_color)
        
        # لیبل‌ها
        style.configure("TLabel", 
                       background=self.bg_color, 
                       foreground=self.text_color,
                       font=("Segoe UI", 10))
        
        # لیبل سایدبار
        style.configure("Sidebar.TLabel", 
                       background=self.sidebar_color, 
                       foreground=self.text_color,
                       font=("Segoe UI", 11, "bold"),
                       padding=5)
        
        # لیبل کارت
        style.configure("Card.TLabel", 
                       background=self.card_color, 
                       foreground=self.text_color,
                       font=("Segoe UI", 10))
        
        # هدر لیبل
        style.configure("Header.TLabel", 
                       foreground=self.text_color,
                       background=self.sidebar_color,
                       font=("Segoe UI", 14, "bold"),
                       padding=15)
        
        # هدر کارت
        style.configure("CardHeader.TLabel", 
                       foreground=self.text_color,
                       background=self.card_color,
                       font=("Segoe UI", 12, "bold"),
                       padding=10)
        
        # لیبل فریم‌ها
        style.configure("TLabelframe", 
                       background=self.card_color,
                       foreground=self.text_color,
                       font=("Segoe UI", 10, "bold"))
        
        style.configure("TLabelframe.Label", 
                       background=self.bg_color,
                       foreground=self.text_color,
                       font=("Segoe UI", 10, "bold"))
        
        # سپراتور
        style.configure("TSeparator", 
                      background=self.secondary_text)
        
        # نوار وضعیت
        style.configure("Status.TLabel", 
                       background=self.sidebar_color,
                       foreground=self.text_color,
                       font=("Segoe UI", 9),
                       relief=tk.FLAT,
                       padding=8)
    
    def _create_custom_button(self, parent, text, command, style="custom", width=None, height=None):
        """ایجاد دکمه سفارشی با پس‌زمینه بنفش و حاشیه سبز"""
        bg_color = self.button_bg
        if style == "primary":
            bg_color = self.button_bg
            font = ("Segoe UI", 10, "bold")
        elif style == "sidebar":
            bg_color = self.button_bg
            font = ("Segoe UI", 10)
        else:
            bg_color = self.button_bg
            font = ("Segoe UI", 10)
            
        # ایجاد دکمه با حاشیه سبز
        btn = tk.Button(parent, 
            text=text,
            command=self._button_click_effect(command),  # اضافه کردن افکت حرکتی
            bg=bg_color,
            fg="white",
            activebackground=self.button_active_bg,
            activeforeground="white",
            font=font,
            bd=2,  # عرض حاشیه
            relief=tk.FLAT,
            padx=10,
            pady=5,
            width=width,
            height=height,
            cursor="hand2", # تغییر شکل موس وقتی روی دکمه میره
            highlightbackground="#4CAF50",  # رنگ حاشیه سبز
            highlightcolor="#4CAF50",       # رنگ حاشیه در حالت فوکوس
            highlightthickness=2)           # ضخامت حاشیه
        return btn
    
    def _button_click_effect(self, callback):
        """افزودن افکت حرکتی به دکمه هنگام کلیک"""
        def wrapper(*args, **kwargs):
            # گرفتن دکمه از رویداد
            event = args[0] if args and hasattr(args[0], 'widget') else None
            
            if event and hasattr(event, 'widget'):
                button = event.widget
                # ذخیره موقعیت اصلی
                original_x = button.winfo_x()
                original_y = button.winfo_y()
                
                # حرکت به سمت پایین و راست (افکت فشرده شدن)
                button.place(x=original_x+2, y=original_y+2)
                button.update_idletasks()
                
                # بازگشت به موقعیت اصلی پس از 100 میلی‌ثانیه
                button.after(100, lambda: button.place(x=original_x, y=original_y))
            
            # فراخوانی عملکرد اصلی دکمه با تمام پارامترها
            return callback(*args, **kwargs)
        return wrapper
    
    def _create_ui(self):
        """ایجاد اجزای رابط کاربری"""
        # فریم اصلی برای چیدمان
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # بخش بالایی
        top_frame = tk.Frame(main_frame, bg=self.bg_color)
        top_frame.pack(fill=tk.X, pady=(0, 15))
        
        # دکمه راهنمایی دیباگ
        debug_help_btn = self._create_custom_button(top_frame, "راهنمای دیباگ", self._show_debug_guidance)
        debug_help_btn.pack(side=tk.LEFT, padx=5)
        
        # متن کپی‌رایت در سمت راست نوار بالایی
        copyright_label = tk.Label(
            top_frame,
            text="Designed by Sina-Salim",
            bg=self.bg_color,
            fg=self.accent_color,
            font=("Segoe UI", 12, "bold")
        )
        copyright_label.pack(side=tk.RIGHT, padx=10)
        
        # === فریم سایدبار (سمت چپ) ===
        sidebar_frame = ttk.Frame(main_frame, style="Sidebar.TFrame", width=250)
        sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        sidebar_frame.pack_propagate(False)  # ثابت نگه‌داشتن عرض
        
        # لوگو و عنوان
        logo_frame = ttk.Frame(sidebar_frame, style="Sidebar.TFrame")
        logo_frame.pack(fill=tk.X, padx=10, pady=(20, 30))
        
        title_label = ttk.Label(logo_frame, text="SolidWorks API Panel", 
                               style="Sidebar.TLabel",
                               font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=5)
        
        # خط جداکننده
        separator = ttk.Separator(sidebar_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=15, pady=5)
        
        # بخش منوی سایدبار
        menu_frame = ttk.Frame(sidebar_frame, style="Sidebar.TFrame")
        menu_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # دکمه‌های منو (با استفاده از دکمه سفارشی)
        self.api_settings_btn = self._create_custom_button(menu_frame, "تنظیمات API", 
                                                          self._on_api_settings, 
                                                          style="sidebar", 
                                                          width=25)
        self.api_settings_btn.pack(fill=tk.X, pady=5)
        
        self.test_api_btn = self._create_custom_button(menu_frame, "تست API", 
                                                      self._on_test_api, 
                                                      style="sidebar", 
                                                      width=25)
        self.test_api_btn.pack(fill=tk.X, pady=5)
        
        self.samples_btn = self._create_custom_button(menu_frame, "استفاده از نمونه آماده", 
                                                     self._on_use_sample, 
                                                     style="sidebar", 
                                                     width=25)
        self.samples_btn.pack(fill=tk.X, pady=5)
        
        # خط جداکننده
        separator2 = ttk.Separator(sidebar_frame, orient='horizontal')
        separator2.pack(fill=tk.X, padx=15, pady=15)
        
        # بخش تاریخچه
        history_label = ttk.Label(sidebar_frame, text="تاریخچه اسکریپت‌ها", style="Sidebar.TLabel")
        history_label.pack(fill=tk.X, padx=15, pady=5, anchor=tk.W)
        
        history_container = ttk.Frame(sidebar_frame, style="Sidebar.TFrame")
        history_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # اسکرول‌بار برای لیست
        history_scrollbar = ttk.Scrollbar(history_container)
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.history_list = tk.Listbox(history_container, 
                                      selectmode=tk.SINGLE, 
                                      bg=self.sidebar_color,
                                      fg=self.text_color,
                                      font=("Segoe UI", 9),
                                      borderwidth=0,
                                      highlightthickness=0,
                                      activestyle="none",
                                      selectbackground=self.accent_color,
                                      selectforeground="white")
        self.history_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # اتصال اسکرول‌بار
        self.history_list.config(yscrollcommand=history_scrollbar.set)
        history_scrollbar.config(command=self.history_list.yview)
        
        self.history_list.bind('<<ListboxSelect>>', self._on_history_select)
        
        # دکمه‌های تاریخچه
        history_btn_frame = ttk.Frame(sidebar_frame, style="Sidebar.TFrame")
        history_btn_frame.pack(fill=tk.X, padx=10, pady=(5, 15))
        
        btn_container = ttk.Frame(history_btn_frame, style="Sidebar.TFrame")
        btn_container.pack(fill=tk.X, expand=True)
        
        self.load_btn = self._create_custom_button(btn_container, "بارگذاری", 
                                                  self._on_load_history, 
                                                  style="custom", 
                                                  width=10)
        self.load_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        self.run_history_btn = self._create_custom_button(btn_container, "اجرا", 
                                                         self._on_run_history, 
                                                         style="custom", 
                                                         width=10)
        self.run_history_btn.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        # === فریم محتوا (سمت راست) ===
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # فریم بالایی برای ورود درخواست
        request_card = ttk.Frame(content_frame, style="Card.TFrame")
        request_card.pack(fill=tk.X, expand=False, padx=20, pady=20)
        
        request_header = ttk.Label(request_card, text="درخواست جدید", style="CardHeader.TLabel")
        request_header.pack(fill=tk.X, anchor=tk.W)
        
        request_content = ttk.Frame(request_card, style="Card.TFrame")
        request_content.pack(fill=tk.X, expand=True, padx=15, pady=(0, 15))
        
        ttk.Label(request_content, text="عملیات مورد نظر را توضیح دهید:", 
                 style="Card.TLabel").pack(anchor=tk.W, pady=(10, 5))
        
        self.query_entry = scrolledtext.ScrolledText(request_content, height=4, width=40, wrap=tk.WORD,
                                                    font=("Segoe UI", 10),
                                                    background="#1E2A4A",
                                                    foreground="white",
                                                    insertbackground="white")
        self.query_entry.pack(fill=tk.X, expand=True, pady=5)
        
        # تغییر رنگ و ظاهر اسکرول‌تکست
        self.query_entry.config(borderwidth=0, relief=tk.FLAT)
        
        buttons_frame = ttk.Frame(request_content, style="Card.TFrame")
        buttons_frame.pack(fill=tk.X, expand=False, pady=(10, 5))
        
        self.submit_btn = self._create_custom_button(buttons_frame, "ارسال درخواست و تولید اسکریپت", 
                                                    self._on_submit, 
                                                    style="primary")
        self.submit_btn.pack(side=tk.LEFT, padx=2)
        
        self.run_btn = self._create_custom_button(buttons_frame, "اجرای اسکریپت فعلی", 
                                                 self._on_run_current)
        self.run_btn.pack(side=tk.LEFT, padx=2)
        
        self.debug_btn = self._create_custom_button(buttons_frame, "دیباگ اسکریپت فعلی", 
                                                  self._on_debug_current)
        self.debug_btn.pack(side=tk.LEFT, padx=2)
        
        # فریم میانی برای نمایش اسکریپت
        script_card = ttk.Frame(content_frame, style="Card.TFrame")
        script_card.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        script_header = ttk.Label(script_card, text="کد اسکریپت", style="CardHeader.TLabel")
        script_header.pack(fill=tk.X, anchor=tk.W)
        
        script_content = ttk.Frame(script_card, style="Card.TFrame")
        script_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        self.script_text = scrolledtext.ScrolledText(script_content, wrap=tk.NONE,
                                                    font=("Consolas", 10),
                                                    background="#1E2A4A",
                                                    foreground="white",
                                                    insertbackground="white")
        self.script_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # رنگ بندی کد
        self._setup_code_highlighting()
        
        # فریم پایینی برای نمایش خروجی
        output_card = ttk.Frame(content_frame, style="Card.TFrame")
        output_card.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        output_header = ttk.Label(output_card, text="خروجی اجرا", style="CardHeader.TLabel")
        output_header.pack(fill=tk.X, anchor=tk.W)
        
        output_content = ttk.Frame(output_card, style="Card.TFrame")
        output_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        self.output_text = scrolledtext.ScrolledText(output_content, height=8, wrap=tk.WORD,
                                                   font=("Consolas", 10),
                                                   background="#1E2A4A",
                                                   foreground="white")
        self.output_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.output_text.config(state=tk.DISABLED)
        
        # === بخش وضعیت ===
        self.status_bar = ttk.Label(self.root, text="آماده", style="Status.TLabel", anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _setup_code_highlighting(self):
        """تنظیم هایلایت ساده برای کد VBScript"""
        # ایجاد تگ‌های رنگ‌بندی
        self.script_text.tag_configure("keyword", foreground="#569CD6")  # آبی روشن
        self.script_text.tag_configure("comment", foreground="#57A64A")  # سبز
        self.script_text.tag_configure("string", foreground="#D69D85")   # نارنجی-قهوه‌ای
        self.script_text.tag_configure("function", foreground="#4EC9B0") # فیروزه‌ای
        
        # ایجاد بایند برای تشخیص کلیدواژه‌ها
        self.script_text.bind("<KeyRelease>", self._highlight_code)
    
    def _highlight_code(self, event=None):
        """هایلایت کردن کد VBScript به صورت ساده"""
        # حذف تمام تگ‌های موجود
        for tag in ["keyword", "comment", "string", "function"]:
            self.script_text.tag_remove(tag, "1.0", "end")
        
        # لیست کلمات کلیدی VBScript
        keywords = ["Option", "Explicit", "Dim", "Set", "Sub", "Function", "End", "If", "Then", 
                   "Else", "ElseIf", "While", "Wend", "For", "To", "Next", "On", "Error", 
                   "Resume", "Call", "Exit", "Loop", "Do", "Until", "Select", "Case"]
        
        # بررسی متن برای یافتن کلمات کلیدی
        for keyword in keywords:
            start_index = "1.0"
            while True:
                start_index = self.script_text.search(r'\y' + keyword + r'\y', start_index, "end", regexp=True)
                if not start_index:
                    break
                end_index = f"{start_index}+{len(keyword)}c"
                self.script_text.tag_add("keyword", start_index, end_index)
                start_index = end_index
        
        # هایلایت کامنت‌ها
        start_index = "1.0"
        while True:
            start_index = self.script_text.search("'", start_index, "end")
            if not start_index:
                break
            line = self.script_text.get(start_index.split('.')[0] + ".0", start_index.split('.')[0] + ".end")
            comment_part = line[int(start_index.split('.')[1]):]
            end_index = f"{start_index.split('.')[0]}.end"
            self.script_text.tag_add("comment", start_index, end_index)
            start_index = f"{int(start_index.split('.')[0]) + 1}.0"
        
        # هایلایت رشته‌ها
        start_index = "1.0"
        while True:
            start_index = self.script_text.search('"', start_index, "end")
            if not start_index:
                break
            end_index = self.script_text.search('"', f"{start_index}+1c", "end")
            if not end_index:
                break
            end_index = f"{end_index}+1c"
            self.script_text.tag_add("string", start_index, end_index)
            start_index = end_index
    
    def _on_submit(self):
        """پردازش درخواست کاربر برای تولید اسکریپت"""
        query = self.query_entry.get("1.0", tk.END).strip()
        
        if not query:
            messagebox.showwarning("خطا", "لطفاً درخواست خود را وارد کنید.")
            return
        
        # غیرفعال کردن دکمه ارسال
        self.submit_btn.config(state=tk.DISABLED)
        self.status_bar.config(text="در حال تولید اسکریپت...")
        
        # اجرای پردازش در ترد جداگانه
        threading.Thread(target=self._generate_script_thread, args=(query,), daemon=True).start()
    
    def _generate_script_thread(self, query):
        """پردازش درخواست در ترد جداگانه

        Args:
            query: متن درخواست کاربر
        """
        try:
            # تولید اسکریپت
            success, message, script_path = self.script_generator.generate_script(query)
            
            # قرار دادن نتیجه در صف برای پردازش در ترد اصلی
            self.queue.put(("generate_result", success, message, script_path))
            
        except Exception as e:
            logger.error(f"خطا در تولید اسکریپت: {e}")
            self.queue.put(("generate_result", False, f"خطا: {str(e)}", None))
    
    def _process_queue(self):
        """پردازش صف پیام‌ها از تردهای دیگر"""
        try:
            while True:
                message = self.queue.get_nowait()
                message_type = message[0]
                
                if message_type == "generate_result":
                    success, script, script_path = message[1], message[2], message[3]
                    self._handle_generate_result(success, script, script_path)
                
                elif message_type == "execute_result":
                    success, result_message, output, script_path = message[1], message[2], message[3], message[4]
                    self._handle_execute_result(success, result_message, output, script_path)
                
                elif message_type == "debug_result":
                    success, fixed_script, explanation, script_path = message[1], message[2], message[3], message[4]
                    self._handle_debug_result(success, fixed_script, explanation, script_path)
                
                elif message_type == "api_test_result":
                    success, message_text, result_label = message[1], message[2], message[3]
                    self._handle_api_test_result(success, message_text, result_label)
                
                elif message_type == "guidance_result":
                    success, answer, answer_widget, status_label = message[1], message[2], message[3], message[4]
                    self._handle_guidance_result(success, answer, answer_widget, status_label)
                
                self.queue.task_done()
                
        except queue.Empty:
            pass
        
        # دوباره پردازش را برنامه‌ریزی کن
        self.root.after(100, self._process_queue)
    
    def _on_run_current(self):
        """اجرای اسکریپت فعلی"""
        script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
        
        if not os.path.exists(script_path):
            messagebox.showwarning("خطا", "هیچ اسکریپتی برای اجرا وجود ندارد.")
            return
        
        # اجرا در ترد جداگانه
        threading.Thread(target=self._execute_script_thread, args=(script_path,), daemon=True).start()
        
        self.status_bar.config(text="در حال اجرای اسکریپت...")
    
    def _execute_script_thread(self, script_path):
        """اجرای اسکریپت در ترد جداگانه

        Args:
            script_path: مسیر فایل اسکریپت
        """
        try:
            # اجرای اسکریپت
            success, message, output = self.script_generator.execute_script(script_path)
            
            # قرار دادن نتیجه در صف
            self.queue.put(("execute_result", success, message, output, script_path))
            
        except Exception as e:
            logger.error(f"خطا در اجرای اسکریپت: {e}")
            self.queue.put(("execute_result", False, f"خطا: {str(e)}", "", script_path))
    
    def _on_debug_current(self):
        """دیباگ کردن اسکریپت فعلی با استفاده از LLM"""
        script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
        
        if not os.path.exists(script_path):
            messagebox.showwarning("خطا", "هیچ اسکریپتی برای دیباگ وجود ندارد.")
            return
        
        # تنظیم پرچم برای نشان دادن درخواست دیباگ
        self._debug_requested = True
        
        # اجرای اسکریپت برای دریافت خطا
        self.status_bar.config(text="در حال اجرای اسکریپت برای شناسایی خطا...")
        threading.Thread(target=self._execute_script_thread, args=(script_path,), daemon=True).start()
    
    def _debug_script_thread(self, script_path, error_message):
        """دیباگ اسکریپت در ترد جداگانه

        Args:
            script_path: مسیر فایل اسکریپت
            error_message: پیام خطای دریافت شده
        """
        try:
            # خواندن محتوای اسکریپت
            with open(script_path, "r", encoding='utf-8') as f:
                script_content = f.read()
            
            # دیباگ اسکریپت
            success, fixed_script, explanation = self.script_debugger.debug_script(script_content, error_message)
            
            # قرار دادن نتیجه در صف
            self.queue.put(("debug_result", success, fixed_script, explanation, script_path))
            
        except Exception as e:
            logger.error(f"خطا در دیباگ اسکریپت: {e}")
            self.queue.put(("debug_result", False, "", f"خطا در دیباگ اسکریپت: {str(e)}", script_path))
    
    def _show_debug_dialog(self, script_path, fixed_script, explanation):
        """نمایش دیالوگ نتیجه دیباگ

        Args:
            script_path: مسیر فایل اسکریپت
            fixed_script: اسکریپت اصلاح شده
            explanation: توضیحات خطا و اصلاحات
        """
        debug_dialog = tk.Toplevel(self.root)
        debug_dialog.title("نتیجه دیباگ اسکریپت")
        debug_dialog.geometry("900x700")
        debug_dialog.transient(self.root)
        debug_dialog.grab_set()
        
        # تنظیم رنگ پس‌زمینه
        debug_dialog.config(bg=self.bg_color)
        
        # فریم اصلی
        main_frame = tk.Frame(debug_dialog, bg=self.bg_color, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # بخش توضیحات
        explanation_frame = tk.LabelFrame(main_frame, text="توضیحات خطا و اصلاحات", 
                                        bg=self.card_color, fg=self.text_color, 
                                        font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        explanation_frame.pack(fill=tk.X, expand=False, pady=(0, 15))
        
        explanation_text = scrolledtext.ScrolledText(explanation_frame, height=7, wrap=tk.WORD,
                                                   font=("Segoe UI", 10),
                                                   background="#1E2A4A",
                                                   foreground="white")
        explanation_text.pack(fill=tk.BOTH, expand=True, pady=5)
        explanation_text.insert("1.0", explanation)
        
        # بخش کد اصلاح شده
        code_frame = tk.LabelFrame(main_frame, text="کد اصلاح شده", 
                                 bg=self.card_color, fg=self.text_color, 
                                 font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        code_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        code_text = scrolledtext.ScrolledText(code_frame, wrap=tk.NONE,
                                            font=("Consolas", 10),
                                            background="#1E2A4A",
                                            foreground="white")
        code_text.pack(fill=tk.BOTH, expand=True, pady=5)
        code_text.insert("1.0", fixed_script)
        
        # فریم دکمه‌ها
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        apply_btn = self._create_custom_button(button_frame, "اعمال تغییرات", 
                                             lambda: self._apply_debug_changes(debug_dialog, script_path, code_text.get("1.0", tk.END)),
                                             style="primary")
        apply_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = self._create_custom_button(button_frame, "انصراف", 
                                              lambda: debug_dialog.destroy())
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def _apply_debug_changes(self, dialog, script_path, fixed_script):
        """اعمال تغییرات دیباگ شده به اسکریپت

        Args:
            dialog: دیالوگ دیباگ
            script_path: مسیر فایل اسکریپت
            fixed_script: محتوای اصلاح شده اسکریپت
        """
        try:
            # ذخیره نسخه اصلاح شده
            with open(script_path, "w", encoding='utf-8') as f:
                f.write(fixed_script)
            
            # بروزرسانی محتوای اسکریپت در ویرایشگر اصلی
            self.script_text.delete("1.0", tk.END)
            self.script_text.insert("1.0", fixed_script)
            self._highlight_code()
            
            # بستن دیالوگ
            dialog.destroy()
            
            # نمایش پیام موفقیت
            self.status_bar.config(text="اسکریپت با موفقیت اصلاح شد.")
            messagebox.showinfo("اصلاح موفق", "اسکریپت با موفقیت اصلاح شد. حالا می‌توانید آن را اجرا کنید.")
            
            logger.info(f"اسکریپت با موفقیت اصلاح شد: {script_path}")
            
        except Exception as e:
            logger.error(f"خطا در اعمال تغییرات دیباگ: {e}")
            messagebox.showerror("خطا", f"خطا در اعمال تغییرات: {str(e)}")
    
    def _update_history_list(self):
        """بروزرسانی لیست تاریخچه اسکریپت‌ها"""
        try:
            # دریافت لیست اسکریپت‌ها
            script_files = glob.glob(os.path.join(HISTORY_DIR, "sw_script_*.vbs"))
            script_files.sort(reverse=True)  # جدیدترین در بالا
            
            # پاک کردن لیست قبلی
            self.history_list.delete(0, tk.END)
            
            # افزودن اسکریپت‌ها به لیست
            self.history_paths = []  # ذخیره مسیرها در یک لیست جداگانه
            
            for script_path in script_files:
                filename = os.path.basename(script_path)
                # استخراج تاریخ و زمان از نام فایل
                date_part = filename.replace("sw_script_", "").replace(".vbs", "")
                try:
                    date_obj = datetime.datetime.strptime(date_part, "%Y%m%d_%H%M%S")
                    display_name = date_obj.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    display_name = filename
                
                self.history_list.insert(tk.END, display_name)
                # ذخیره مسیر کامل در لیست جداگانه
                self.history_paths.append(script_path)
            
        except Exception as e:
            logger.error(f"خطا در بروزرسانی لیست تاریخچه: {e}")
    
    def _on_history_select(self, event):
        """انتخاب یک اسکریپت از لیست تاریخچه"""
        # هیچ عملیاتی انجام نمی‌شود، فقط انتخاب
        pass
    
    def _on_load_history(self):
        """بارگذاری اسکریپت انتخابی از تاریخچه"""
        selected_indices = self.history_list.curselection()
        
        if not selected_indices:
            messagebox.showwarning("خطا", "لطفاً یک اسکریپت را از لیست انتخاب کنید.")
            return
        
        # دریافت مسیر اسکریپت انتخابی
        idx = selected_indices[0]
        script_path = self.history_paths[idx]
        
        try:
            # نمایش محتوای اسکریپت
            with open(script_path, "r", encoding='utf-8') as f:
                script_content = f.read()
            
            self.script_text.delete("1.0", tk.END)
            self.script_text.insert("1.0", script_content)
            
            # کپی به اسکریپت فعلی
            current_script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
            shutil.copy2(script_path, current_script_path)
            
            self.status_bar.config(text=f"اسکریپت بارگذاری شد: {os.path.basename(script_path)}")
            
        except Exception as e:
            logger.error(f"خطا در بارگذاری اسکریپت: {e}")
            messagebox.showerror("خطا", f"خطا در بارگذاری اسکریپت: {str(e)}")
    
    def _on_run_history(self):
        """اجرای اسکریپت انتخابی از تاریخچه"""
        selected_indices = self.history_list.curselection()
        
        if not selected_indices:
            messagebox.showwarning("خطا", "لطفاً یک اسکریپت را از لیست انتخاب کنید.")
            return
        
        # دریافت مسیر اسکریپت انتخابی
        idx = selected_indices[0]
        script_path = self.history_paths[idx]
        
        # اجرای اسکریپت
        threading.Thread(target=self._execute_script_thread, args=(script_path,), daemon=True).start()
        
        self.status_bar.config(text="در حال اجرای اسکریپت...")

    def _on_use_sample(self):
        """استفاده از اسکریپت نمونه"""
        sample_script_path = os.path.join(SCRIPTS_DIR, "create_simple_part.vbs")
        
        if not os.path.exists(sample_script_path):
            messagebox.showwarning("خطا", "فایل اسکریپت نمونه یافت نشد.")
            return
        
        try:
            # نمایش محتوای اسکریپت
            with open(sample_script_path, "r", encoding='utf-8') as f:
                script_content = f.read()
            
            self.script_text.delete("1.0", tk.END)
            self.script_text.insert("1.0", script_content)
            
            # کپی به اسکریپت فعلی
            current_script_path = os.path.join(SCRIPTS_DIR, "current_script.vbs")
            shutil.copy2(sample_script_path, current_script_path)
            
            self.status_bar.config(text=f"اسکریپت نمونه بارگذاری شد")
            
        except Exception as e:
            logger.error(f"خطا در بارگذاری اسکریپت نمونه: {e}")
            messagebox.showerror("خطا", f"خطا در بارگذاری اسکریپت نمونه: {str(e)}")

    def _on_api_settings(self):
        """نمایش دیالوگ تنظیمات API"""
        # دریافت تنظیمات فعلی
        api_key = self.script_generator.api_key
        base_url = self.script_generator.api_url
        api_model = self.script_generator.api_model
        
        # نمایش دیالوگ
        result = APISettingsDialog.show_dialog(self.root, api_key, base_url, api_model)
        
        if result:
            # اعمال تنظیمات جدید
            self.script_generator.api_key = result["api_key"]
            self.script_generator.api_url = result["base_url"]
            self.script_generator.api_model = result["api_model"]
            
            # بروزرسانی هدرها
            self.script_generator.headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.script_generator.api_key}",
                "HTTP-Referer": "https://solipy.app"
            }
            
            # ذخیره تنظیمات در فایل .env
            try:
                self._save_api_settings(result["api_key"], result["base_url"], result["api_model"])
                self.status_bar.config(text="تنظیمات API با موفقیت ذخیره شد.")
            except Exception as e:
                logger.error(f"خطا در ذخیره تنظیمات API: {e}")
                messagebox.showerror("خطا", f"خطا در ذخیره تنظیمات API: {str(e)}")
    
    def _save_api_settings(self, api_key, base_url, api_model):
        """ذخیره تنظیمات API در فایل .env
        
        Args:
            api_key: کلید API جدید
            base_url: آدرس API جدید
            api_model: مدل هوش مصنوعی جدید
        """
        env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
        
        try:
            # خواندن محتوای فایل .env موجود (اگر وجود دارد)
            env_lines = []
            if os.path.exists(env_path):
                with open(env_path, 'r', encoding='utf-8') as f:
                    env_lines = f.readlines()
            
            # جایگزینی یا افزودن تنظیمات API
            api_key_found = False
            base_url_found = False
            api_model_found = False
            
            for i, line in enumerate(env_lines):
                if line.startswith('OPENAI_API_KEY='):
                    env_lines[i] = f'OPENAI_API_KEY="{api_key}"\n'
                    api_key_found = True
                elif line.startswith('OPENAI_BASE_URL='):
                    env_lines[i] = f'OPENAI_BASE_URL="{base_url}"\n'
                    base_url_found = True
                elif line.startswith('OPENAI_MODEL='):
                    env_lines[i] = f'OPENAI_MODEL="{api_model}"\n'
                    api_model_found = True
            
            # افزودن تنظیمات جدید اگر قبلاً وجود نداشته باشند
            if not api_key_found:
                env_lines.append(f'OPENAI_API_KEY="{api_key}"\n')
            if not base_url_found:
                env_lines.append(f'OPENAI_BASE_URL="{base_url}"\n')
            if not api_model_found:
                env_lines.append(f'OPENAI_MODEL="{api_model}"\n')
            
            # نوشتن در فایل
            with open(env_path, 'w', encoding='utf-8') as f:
                f.writelines(env_lines)
            
            logger.info("تنظیمات API در فایل .env ذخیره شد.")
            
        except Exception as e:
            logger.error(f"خطا در ذخیره تنظیمات API: {e}")
            raise

    def _on_test_api(self):
        """تست اتصال به API فعلی"""
        self.status_bar.config(text="در حال تست اتصال به API...")
        
        # اجرای تست در یک ترد جداگانه
        threading.Thread(target=self._test_api_thread, daemon=True).start()
    
    def _test_api_thread(self):
        """اجرای تست API در ترد جداگانه"""
        try:
            api_key = self.script_generator.api_key
            base_url = self.script_generator.api_url
            api_model = self.script_generator.api_model
            
            if not api_key or not base_url or not api_model:
                self.queue.put(("api_test_result", False, "لطفاً ابتدا تنظیمات API را کامل کنید"))
                return
            
            success, message = APITester.test_api_connection(api_key, base_url, api_model)
            
            # ارسال نتیجه به صف برای پردازش در ترد اصلی
            self.queue.put(("api_test_result", success, message))
            
        except Exception as e:
            logger.error(f"خطا در تست API: {e}")
            self.queue.put(("api_test_result", False, f"خطا: {str(e)}"))

    def _handle_generate_result(self, success, script, script_path):
        """پردازش نتیجه درخواست تولید اسکریپت

        Args:
            success: وضعیت موفقیت درخواست
            script: محتوای اسکریپت تولید شده
            script_path: مسیر فایل اسکریپت تولید شده
        """
        if success and script_path:
            # نمایش اسکریپت در بخش متن
            with open(script_path, "r", encoding='utf-8') as f:
                script_content = f.read()
            
            self.script_text.delete("1.0", tk.END)
            self.script_text.insert("1.0", script_content)
            
            # بروزرسانی لیست تاریخچه
            self._update_history_list()
            
            self.status_bar.config(text=f"اسکریپت با موفقیت ایجاد شد: {os.path.basename(script_path)}")
        else:
            self.status_bar.config(text=f"خطا در تولید اسکریپت: {script or 'خطا در درخواست API'}")
            messagebox.showerror("خطا در تولید اسکریپت", script or "خطا در درخواست API")
    
    def _handle_execute_result(self, success, result_message, output, script_path):
        """پردازش نتیجه درخواست اجرای اسکریپت

        Args:
            success: وضعیت موفقیت درخواست
            result_message: پیام نتیجه اجرای اسکریپت
            output: خروجی اسکریپت
            script_path: مسیر فایل اسکریپت
        """
        if success:
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", output or result_message)
            self.output_text.config(state=tk.DISABLED)
            self.status_bar.config(text="اسکریپت با موفقیت اجرا شد.")
        else:
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", output or result_message)
            self.output_text.config(state=tk.DISABLED)
            self.status_bar.config(text=f"خطا در اجرای اسکریپت: {result_message}")
            messagebox.showerror("خطا در اجرای اسکریپت", result_message)
    
    def _handle_debug_result(self, success, fixed_script, explanation, script_path):
        """پردازش نتیجه درخواست دیباگ اسکریپت

        Args:
            success: وضعیت موفقیت درخواست
            fixed_script: اسکریپت اصلاح شده
            explanation: توضیحات خطا و اصلاحات
            script_path: مسیر فایل اسکریپت
        """
        if success and fixed_script:
            self.status_bar.config(text="اسکریپت با موفقیت دیباگ شد.")
            self._show_debug_dialog(script_path, fixed_script, explanation)
        else:
            self.status_bar.config(text="خطا در دیباگ اسکریپت.")
            messagebox.showerror("خطا در دیباگ", explanation)
    
    def _handle_api_test_result(self, success, message_text, result_label):
        """پردازش نتیجه تست API

        Args:
            success: وضعیت موفقیت درخواست
            message_text: متن پیام تست API
            result_label: لیبل نتیجه تست API
        """
        if success:
            result_label.config(text=f"تست API: ✓ {message_text}", foreground="#4CAF50", background=self.bg_color)
            messagebox.showinfo("تست API", "اتصال به API با موفقیت برقرار شد.")
        else:
            result_label.config(text=f"تست API: ✗ {message_text}", foreground="#F44336", background=self.bg_color)
            messagebox.showerror("تست API", f"خطا در اتصال به API: {message_text}")
    
    def _handle_guidance_result(self, success, answer, answer_widget, status_label):
        """پردازش نتیجه درخواست راهنمایی

        Args:
            success: وضعیت موفقیت درخواست
            answer: پاسخ دریافتی
            answer_widget: ویجت متن برای نمایش پاسخ
            status_label: لیبل وضعیت
        """
        answer_widget.delete("1.0", tk.END)
        
        if success:
            answer_widget.insert("1.0", answer)
            status_label.config(text="پاسخ دریافت شد", fg="lightgreen")
        else:
            answer_widget.insert("1.0", f"خطا در دریافت پاسخ: {answer}")
            status_label.config(text="خطا در دریافت پاسخ", fg="red")
    
    def _show_debug_guidance(self):
        """نمایش دیالوگ راهنمایی دیباگ و طراحی اسکریپت"""
        guidance_dialog = tk.Toplevel(self.root)
        guidance_dialog.title("راهنمای دیباگ و طراحی اسکریپت")
        guidance_dialog.geometry("900x700")
        guidance_dialog.transient(self.root)
        guidance_dialog.grab_set()
        
        # تنظیم رنگ پس‌زمینه
        guidance_dialog.config(bg=self.bg_color)
        
        # فریم اصلی
        main_frame = tk.Frame(guidance_dialog, bg=self.bg_color, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # بخش ورود سوال
        question_frame = tk.LabelFrame(main_frame, text="سوال یا مشکل خود را وارد کنید", 
                                     bg=self.card_color, fg=self.text_color, 
                                     font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        question_frame.pack(fill=tk.X, expand=False, pady=(0, 15))
        
        question_text = scrolledtext.ScrolledText(question_frame, height=5, wrap=tk.WORD,
                                                font=("Segoe UI", 10),
                                                background="#1E2A4A",
                                                foreground="white")
        question_text.pack(fill=tk.BOTH, expand=True, pady=5)
        question_text.insert("1.0", "سوال یا مشکل خود را درباره دیباگ یا طراحی اسکریپت‌های SolidWorks اینجا بنویسید...")
        
        # پاک کردن متن نمونه هنگام کلیک
        def _on_question_click(event):
            if question_text.get("1.0", "end-1c") == "سوال یا مشکل خود را درباره دیباگ یا طراحی اسکریپت‌های SolidWorks اینجا بنویسید...":
                question_text.delete("1.0", tk.END)
        question_text.bind("<Button-1>", _on_question_click)
        
        # دکمه ارسال سوال
        send_frame = tk.Frame(main_frame, bg=self.bg_color)
        send_frame.pack(fill=tk.X, pady=(0, 15))
        
        status_label = tk.Label(send_frame, text="", bg=self.bg_color, fg=self.text_color, font=("Segoe UI", 10))
        status_label.pack(side=tk.RIGHT, padx=5)
        
        send_btn = self._create_custom_button(send_frame, "ارسال سوال", 
                                            lambda: self._send_guidance_request(question_text.get("1.0", tk.END), answer_text, status_label),
                                            style="primary")
        send_btn.pack(side=tk.LEFT, padx=5)
        
        # بخش پاسخ
        answer_frame = tk.LabelFrame(main_frame, text="پاسخ راهنما", 
                                    bg=self.card_color, fg=self.text_color, 
                                    font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        answer_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        answer_text = scrolledtext.ScrolledText(answer_frame, wrap=tk.WORD,
                                              font=("Segoe UI", 10),
                                              background="#1E2A4A",
                                              foreground="white")
        answer_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # فریم دکمه‌های پایین
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        close_btn = self._create_custom_button(button_frame, "بستن", 
                                             lambda: guidance_dialog.destroy())
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # پیشنهادات رایج
        suggestions_frame = tk.LabelFrame(main_frame, text="سوالات پیشنهادی", 
                                        bg=self.card_color, fg=self.text_color, 
                                        font=("Segoe UI", 11, "bold"), padx=10, pady=10)
        suggestions_frame.pack(fill=tk.X, expand=False, pady=(0, 15))
        
        suggestions = [
            "چگونه می‌توانم یک خطای Object referenced has been disconnected را رفع کنم؟",
            "چطور باید یک مدل سه‌بعدی در SolidWorks با VBScript ایجاد کنم؟",
            "روش اضافه کردن قید بین قطعات در SolidWorks چیست؟",
            "چگونه می‌توانم از خواص سفارشی در مدل‌های SolidWorks استفاده کنم؟",
            "چرا اسکریپت من با خطای Type mismatch مواجه می‌شود؟"
        ]
        
        for i, suggestion in enumerate(suggestions):
            def make_click_handler(sugg):
                return lambda: question_text.delete("1.0", tk.END) or question_text.insert("1.0", sugg)
            
            suggestion_btn = tk.Button(suggestions_frame, text=suggestion, 
                                      font=("Segoe UI", 9),
                                      bg="#283A5B", fg=self.text_color,
                                      relief=tk.FLAT, padx=5, pady=3,
                                      command=make_click_handler(suggestion))
            suggestion_btn.pack(fill=tk.X, pady=2, padx=3, anchor=tk.W)
    
    def _send_guidance_request(self, question, answer_text_widget, status_label):
        """ارسال درخواست راهنمایی به LLM

        Args:
            question: سوال کاربر
            answer_text_widget: ویجت متن برای نمایش پاسخ
            status_label: لیبل وضعیت
        """
        if not question or question.strip() == "" or question.strip() == "سوال یا مشکل خود را درباره دیباگ یا طراحی اسکریپت‌های SolidWorks اینجا بنویسید...":
            status_label.config(text="لطفاً سوال خود را وارد کنید", fg="red")
            return
        
        # نمایش وضعیت
        status_label.config(text="در حال دریافت پاسخ...", fg="white")
        answer_text_widget.delete("1.0", tk.END)
        answer_text_widget.insert("1.0", "لطفاً صبر کنید...")
        
        # ارسال درخواست در یک ترد جداگانه
        threading.Thread(
            target=self._guidance_request_thread,
            args=(question, answer_text_widget, status_label),
            daemon=True
        ).start()
    
    def _guidance_request_thread(self, question, answer_text_widget, status_label):
        """ارسال درخواست راهنمایی به LLM در یک ترد جداگانه

        Args:
            question: سوال کاربر
            answer_text_widget: ویجت متن برای نمایش پاسخ
            status_label: لیبل وضعیت
        """
        try:
            # ارسال درخواست
            success, answer = self.script_debugger.provide_user_guidance(question)
            
            # قرار دادن نتیجه در صف برای پردازش در ترد اصلی
            self.queue.put(("guidance_result", success, answer, answer_text_widget, status_label))
            
        except Exception as e:
            logger.error(f"خطا در ارسال درخواست راهنمایی: {e}")
            self.queue.put((
                "guidance_result", 
                False, 
                f"خطا در ارسال درخواست: {str(e)}", 
                answer_text_widget, 
                status_label
            ))

def main():
    """تابع اصلی برنامه"""
    try:
        # بررسی سیستم عامل
        if not sys.platform.startswith('win'):
            logger.error("این برنامه فقط در سیستم عامل ویندوز قابل اجراست.")
            sys.exit(1)
        
        # راه‌اندازی رابط کاربری
        root = tk.Tk()
        app = SolidWorksPanel(root)
        root.mainloop()
        
    except Exception as e:
        logger.error(f"خطا در اجرای برنامه: {e}")
        print(f"خطا در اجرای برنامه: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 