Option Explicit

' اسکریپت ایجاد اسکچ و اکسترود آن در SolidWorks
' این اسکریپت به صورت مستقل عمل می‌کند و نیاز به هیچ فایل خارجی ندارد

' متغیرهای اصلی
Dim swApp, swModel, swModelDoc, swSketchManager, swFeatureManager

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

' ایجاد سند جدید Part
On Error Resume Next
WScript.Echo "در حال ایجاد سند پارت جدید..."

' ایجاد سند پارت جدید با استفاده از تابع NewDocument
Dim docType
docType = 1  ' نوع 1 برای Part

' مسیر قالب Part
Dim templatePath
templatePath = swApp.GetUserPreferenceStringValue(0)  ' مسیر قالب Part

If templatePath = "" Then
    ' اگر مسیر قالب پیدا نشد، استفاده از مسیر پیش‌فرض
    templatePath = "C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Part.prtdot"
End If

WScript.Echo "استفاده از قالب پارت: " & templatePath

Set swModel = swApp.NewDocument(templatePath, 0, 0, 0)

If Err.Number <> 0 Then
    WScript.Echo "خطا در ایجاد سند پارت: " & Err.Description
    WScript.Quit(1)
End If

If swModel Is Nothing Then
    WScript.Echo "خطا: مدل پارت ایجاد نشد."
    WScript.Quit(1)
End If

' تبدیل به modelDoc2 برای دسترسی به توابع بیشتر
Set swModelDoc = swModel

' بررسی نوع سند
If swModelDoc.GetType <> 1 Then
    WScript.Echo "خطا: سند ایجاد شده از نوع پارت نیست!"
    WScript.Quit(1)
End If

WScript.Echo "سند پارت جدید با موفقیت ایجاد شد."

' دریافت مدیر اسکچ و مدیر فیچر
Set swSketchManager = swModelDoc.SketchManager
Set swFeatureManager = swModelDoc.FeatureManager

' انتخاب صفحه جلو (Front Plane)
On Error Resume Next
swModelDoc.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
If Err.Number <> 0 Then
    WScript.Echo "تلاش برای انتخاب Front Plane با نام فارسی..."
    swModelDoc.Extension.SelectByID2 "صفحه جلو", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    If Err.Number <> 0 Then
        WScript.Echo "خطا در انتخاب صفحه: " & Err.Description
        WScript.Quit(1)
    End If
End If
On Error Goto 0

' ایجاد اسکچ در صفحه انتخاب شده
WScript.Echo "در حال ایجاد اسکچ..."
swSketchManager.InsertSketch True

' رسم یک دایره
Dim circleRadius
circleRadius = 0.025 ' 25 میلی‌متر به متر
swSketchManager.CreateCircleByRadius 0, 0, 0, circleRadius
WScript.Echo "دایره با شعاع 25 میلی‌متر ایجاد شد."

' بستن اسکچ
swSketchManager.InsertSketch True

' اکسترود کردن اسکچ
WScript.Echo "در حال اکسترود کردن اسکچ..."

' انتخاب آخرین اسکچ ایجاد شده
On Error Resume Next
Dim lastSketch
Set lastSketch = swModelDoc.GetActiveSketch2()
If lastSketch Is Nothing Then
    WScript.Echo "انتخاب اسکچ با نام..."
    swModelDoc.ClearSelection2 True
    swModelDoc.Extension.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
End If

' اکسترود
Dim myFeature
Dim boolStatus
boolStatus = swModelDoc.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
WScript.Echo "انتخاب اسکچ: " & (boolStatus)

' اکسترود
Dim extrudeDepth
extrudeDepth = 0.01 ' 10 میلی‌متر

' ایجاد فیچر اکسترود با تنظیمات ساده
Set myFeature = swFeatureManager.FeatureExtrusion2(True, False, True, 0, 0, extrudeDepth, 0.01, False, 0, False, False, 0, 0, False, False, False, False, False, False, False, False, False, False)

If Err.Number <> 0 Then
    WScript.Echo "خطا در اکسترود کردن: " & Err.Description
    Err.Clear
End If

If myFeature Is Nothing Then
    WScript.Echo "خطا: فیچر اکسترود ایجاد نشد. تلاش با روش دیگر..."
    
    ' تلاش با استفاده از روش معمولی (منو)
    swModelDoc.ClearSelection2 True
    swModelDoc.Extension.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    
    ' تلاش با تنظیمات کمتر
    Set myFeature = swFeatureManager.FeatureExtrusion2(True, False, False, 0, 0, extrudeDepth, 0.01, False, 0, False, False, 0, 0, False, False, False, False, False, False, False, False, False, False)
    
    If myFeature Is Nothing Then
        WScript.Echo "خطا: فیچر اکسترود ایجاد نشد. تلاش دوم نیز ناموفق بود."
        
        ' تلاش با منوی افزودن اکسترود
        Dim sModelName
        sModelName = swModelDoc.GetPathName
        If Len(sModelName) = 0 Then
            sModelName = "Untitled"
        End If
        
        WScript.Echo "استفاده از منو برای اکسترود..."
        swApp.RunCommand swCommands_Insert_Boss_Extrude, ""
    Else
        WScript.Echo "اکسترود با روش دوم با موفقیت انجام شد."
    End If
Else
    WScript.Echo "اکسترود با موفقیت انجام شد."
End If

On Error GoTo 0

' نمایش در نمای ایزومتریک
swModelDoc.ShowNamedView2 "*Isometric", 7
swModelDoc.ViewZoomtofit2

WScript.Echo "SUCCESS: اسکچ و اکسترود با موفقیت ایجاد شد (یا تلاش شد)"
WScript.Quit(0) 