Option Explicit

' اسکریپت رسم دایره در SolidWorks
' این اسکریپت به SolidWorks متصل شده و یک دایره می‌کشد

' متغیرهای اصلی
Dim swApp, swModel, swModelDoc, swSketchManager

' اتصال به SolidWorks
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

' ایجاد سند جدید یا استفاده از سند فعلی
On Error Resume Next
WScript.Echo "در حال آماده‌سازی سند..."

' بررسی سند فعال
Set swModel = swApp.ActiveDoc

If swModel Is Nothing Then
    ' اگر سندی باز نیست، ایجاد سند جدید
    Dim templatePath
    templatePath = swApp.GetUserPreferenceStringValue(0)  ' مسیر قالب پارت
    
    If templatePath = "" Then
        ' استفاده از مسیر پیش‌فرض
        templatePath = "C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Part.prtdot"
    End If
    
    WScript.Echo "در حال ایجاد سند جدید با قالب: " & templatePath
    Set swModel = swApp.NewDocument(templatePath, 0, 0, 0)
    
    If Err.Number <> 0 Then
        WScript.Echo "خطا در ایجاد سند جدید: " & Err.Description
        WScript.Quit(1)
    End If
Else
    WScript.Echo "استفاده از سند فعال موجود..."
End If

' تبدیل به modelDoc2 برای دسترسی به توابع بیشتر
Set swModelDoc = swModel

On Error GoTo 0

' انتخاب صفحه یا سطح برای رسم
On Error Resume Next
swModelDoc.ClearSelection2 True
Dim boolStatus
boolStatus = swModelDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

If boolStatus = False Then
    ' تلاش برای انتخاب با نام فارسی
    boolStatus = swModelDoc.Extension.SelectByID2("صفحه جلو", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    
    If boolStatus = False Then
        ' استفاده از صفحه بالا به عنوان گزینه دوم
        boolStatus = swModelDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        
        If boolStatus = False Then
            ' تلاش برای انتخاب با نام فارسی
            boolStatus = swModelDoc.Extension.SelectByID2("صفحه بالا", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            
            If boolStatus = False Then
                WScript.Echo "خطا: نمی‌توان صفحه مناسب برای رسم پیدا کرد."
                WScript.Quit(1)
            End If
        End If
    End If
End If

On Error GoTo 0

' ایجاد اسکچ جدید
Set swSketchManager = swModelDoc.SketchManager
swSketchManager.InsertSketch True
WScript.Echo "اسکچ جدید ایجاد شد."

' رسم دایره
Dim centerX, centerY, radius
centerX = 0        ' مرکز دایره، مختصات X
centerY = 0        ' مرکز دایره، مختصات Y
radius = 0.05      ' شعاع دایره (متر)

WScript.Echo "در حال رسم دایره با شعاع " & (radius * 1000) & " میلی‌متر..."
swSketchManager.CreateCircleByRadius centerX, centerY, 0, radius

' پایان اسکچ
swSketchManager.InsertSketch True

' نمایش در نمای ایزومتریک و مقیاس مناسب
swModelDoc.ShowNamedView2 "*Isometric", 7
swModelDoc.ViewZoomtofit2

WScript.Echo "SUCCESS: دایره با موفقیت رسم شد."
WScript.Quit(0) 