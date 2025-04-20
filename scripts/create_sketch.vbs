Option Explicit

' اسکریپت ایجاد اسکچ در SolidWorks
' این اسکریپت به SolidWorks متصل شده و یک اسکچ با شکل‌های درخواستی کاربر ایجاد می‌کند

' اینکلود کردن اسکریپت اتصال
' توجه: در مسیر اسکریپت قرار دهید
Dim fso, connectScriptPath, fileContents
Set fso = CreateObject("Scripting.FileSystemObject")
connectScriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\connect_to_sw.vbs"

If fso.FileExists(connectScriptPath) Then
    ' فایل اتصال را بخوان و شامل کن
    Dim file
    Set file = fso.OpenTextFile(connectScriptPath, 1)
    fileContents = file.ReadAll()
    file.Close
    
    ' اجرای محتوای اسکریپت کانکشن
    ExecuteGlobal fileContents
Else
    WScript.Echo "خطا: فایل connect_to_sw.vbs یافت نشد"
    WScript.Quit(1)
End If

' متغیرهای اصلی
Dim swApp, swModel, swSketch, swSketchManager

' اتصال به SolidWorks
Set swApp = ConnectToSolidWorks()

' ایجاد یا باز کردن سند
On Error Resume Next
WScript.Echo "در حال باز کردن یا ایجاد سند جدید..."
Dim documentType
documentType = "Part"  ' نوع سند: Part, Assembly, Drawing

' مسیر قالب
Dim templatePath
templatePath = ""

' اگر سندی باز است، از آن استفاده کن
Set swModel = swApp.ActiveDoc

If swModel Is Nothing Then
    ' بررسی مسیرهای معمول برای قالب
    If templatePath = "" Then
        templatePath = swApp.GetUserPreferenceStringValue(9)
    End If

    ' اگر قالب پیدا نشد، از مسیر پیش‌فرض استفاده کن
    If templatePath = "" Then
        If documentType = "Part" Then
            templatePath = swApp.GetUserPreferenceStringValue(0)
        ElseIf documentType = "Assembly" Then
            templatePath = swApp.GetUserPreferenceStringValue(1)
        ElseIf documentType = "Drawing" Then
            templatePath = swApp.GetUserPreferenceStringValue(2)
        End If
    End If

    ' اگر هنوز قالب پیدا نشد، از مسیر سیستمی استفاده کن
    If templatePath = "" Then
        Dim templateFolder
        templateFolder = swApp.GetExecutablePath
        templateFolder = Left(templateFolder, InStrRev(templateFolder, "\"))
        templateFolder = templateFolder & "data\templates\"
        
        If documentType = "Part" Then
            templatePath = templateFolder & "Part.prtdot"
        ElseIf documentType = "Assembly" Then
            templatePath = templateFolder & "Assembly.asmdot"
        ElseIf documentType = "Drawing" Then
            templatePath = templateFolder & "Drawing.drwdot"
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
Else
    WScript.Echo "استفاده از سند فعال موجود..."
End If

On Error GoTo 0

' نمایش در نمای از بالا (Top)
swModel.ShowNamedView2 "Top", 1

' ایجاد اسکچ در صفحه بالا
Set swSketchManager = swModel.SketchManager
swSketchManager.InsertSketch True

' ایجاد اشکال در اسکچ
' مثال: رسم یک مربع
WScript.Echo "در حال ایجاد اسکچ..."
swSketchManager.CreateCenterRectangle 0, 0, 0, 0.05, 0.05, 0

' مثال: رسم یک دایره
swSketchManager.CreateCircleByRadius 0, 0, 0, 0.025

' مثال: رسم یک خط
swSketchManager.CreateLine -0.05, -0.05, 0, 0.05, -0.05, 0

' بستن اسکچ
swSketchManager.InsertSketch True

' نمایش در نمای ایزومتریک
swModel.ShowNamedView2 "*Isometric", 7
swModel.ViewZoomtofit2

WScript.Echo "SUCCESS: اسکچ با موفقیت ایجاد شد"
WScript.Quit(0) 