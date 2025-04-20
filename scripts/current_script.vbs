Option Explicit

' اسکریپت ساده ایجاد فایل جدید در SolidWorks
' این اسکریپت یک فایل قطعه جدید ایجاد می‌کند
' نویسنده: تیم SoliPy

' متغیرهای اصلی
Dim swApp, swModel

' تلاش برای اتصال به SolidWorks
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

' نمایش نرم‌افزار
swApp.Visible = True
WScript.Echo "اتصال به SolidWorks با موفقیت انجام شد."
On Error Goto 0

' ایجاد سند جدید
On Error Resume Next
WScript.Echo "در حال ایجاد سند جدید..."

' مسیر قالب
Dim templatePath
templatePath = ""

' بررسی مسیرهای معمول برای قالب پارت
If templatePath = "" Then
    ' تلاش برای استفاده از تنظیمات کاربر
    templatePath = swApp.GetUserPreferenceStringValue(9)
End If

' اگر هنوز پیدا نشد، استفاده از مسیر پیش‌فرض
If templatePath = "" Then
    templatePath = swApp.GetExecutablePath
    templatePath = Left(templatePath, InStrRev(templatePath, "\\"))
    templatePath = templatePath & "data\\templates\\Part.prtdot"
    
    If Err.Number <> 0 Then
        WScript.Echo "خطا در دریافت مسیر فایل‌های قالب: " & Err.Description
        ' استفاده از مسیر مطلق به عنوان آخرین راه حل
        templatePath = "C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2023\\templates\\Part.prtdot"
        Err.Clear
    End If
End If

WScript.Echo "استفاده از قالب: " & templatePath

' ایجاد سند جدید
Set swModel = swApp.NewDocument(templatePath, 0, 0, 0)

If Err.Number <> 0 Then
    WScript.Echo "خطا در ایجاد سند جدید: " & Err.Description
    WScript.Quit(1)
End If

If swModel Is Nothing Then
    WScript.Echo "خطا: مدل ایجاد نشد."
    WScript.Quit(1)
End If

WScript.Echo "سند جدید با موفقیت ایجاد شد."

' نمایش در نمای ایزومتریک
swModel.ShowNamedView2 "*Isometric", 7
swModel.ViewZoomtofit2

WScript.Echo "SUCCESS: فایل جدید با موفقیت ایجاد شد."
WScript.Quit(0) 