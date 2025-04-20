Option Explicit

' اسکریپت ایجاد اسکچ بر اساس ورودی در SolidWorks
' این اسکریپت به SolidWorks متصل شده و اسکچ‌هایی را بر اساس پارامترهای ورودی ایجاد می‌کند
' نحوه استفاده:
' cscript create_sketch_from_input.vbs shape=circle x=0 y=0 radius=10
' cscript create_sketch_from_input.vbs shape=rectangle x=0 y=0 width=20 height=10
' cscript create_sketch_from_input.vbs shape=line x1=0 y1=0 x2=30 y2=20

' اینکلود کردن اسکریپت اتصال
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

' تابع پارس کردن آرگومان‌ها
Function ParseArguments()
    Dim params
    Set params = CreateObject("Scripting.Dictionary")
    Dim i, arg, argParts, key, value
    
    ' بررسی تعداد آرگومان‌ها
    If WScript.Arguments.Count < 1 Then
        WScript.Echo "خطا: پارامترهای کافی ارائه نشده است"
        WScript.Echo "نحوه استفاده:"
        WScript.Echo "cscript create_sketch_from_input.vbs shape=circle x=0 y=0 radius=10"
        WScript.Echo "cscript create_sketch_from_input.vbs shape=rectangle x=0 y=0 width=20 height=10"
        WScript.Echo "cscript create_sketch_from_input.vbs shape=line x1=0 y1=0 x2=30 y2=20"
        WScript.Quit(1)
    End If
    
    ' پردازش همه آرگومان‌ها
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        argParts = Split(arg, "=")
        
        If UBound(argParts) = 1 Then
            key = LCase(Trim(argParts(0)))
            value = Trim(argParts(1))
            
            ' تبدیل مقادیر عددی به عدد
            If IsNumeric(value) Then
                value = CDbl(value)
            End If
            
            params.Add key, value
        End If
    Next
    
    Set ParseArguments = params
End Function

' متغیرهای اصلی
Dim swApp, swModel, swSketchManager

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

' دریافت و پردازش پارامترها
Dim params, shape, x, y, radius, width, height, x1, y1, x2, y2
Set params = ParseArguments()

' استخراج نوع شکل
If params.Exists("shape") Then
    shape = LCase(params("shape"))
Else
    WScript.Echo "خطا: نوع شکل (shape) مشخص نشده است"
    WScript.Quit(1)
End If

' پردازش بر اساس نوع شکل
WScript.Echo "در حال ایجاد اسکچ با شکل: " & shape
Select Case shape
    Case "circle"
        ' پارامترهای دایره
        If params.Exists("x") Then x = params("x") Else x = 0
        If params.Exists("y") Then y = params("y") Else y = 0
        If params.Exists("radius") Then 
            radius = params("radius") / 1000  ' تبدیل به متر
        Else
            radius = 0.01  ' مقدار پیش‌فرض 10mm
        End If
        
        ' رسم دایره
        WScript.Echo "ایجاد دایره با مرکز (" & x & ", " & y & ") و شعاع " & radius * 1000 & "mm"
        swSketchManager.CreateCircleByRadius x/1000, y/1000, 0, radius
        
    Case "rectangle"
        ' پارامترهای مستطیل
        If params.Exists("x") Then x = params("x") Else x = 0
        If params.Exists("y") Then y = params("y") Else y = 0
        If params.Exists("width") Then 
            width = params("width") / 1000  ' تبدیل به متر
        Else
            width = 0.02  ' مقدار پیش‌فرض 20mm
        End If
        If params.Exists("height") Then 
            height = params("height") / 1000  ' تبدیل به متر
        Else
            height = 0.01  ' مقدار پیش‌فرض 10mm
        End If
        
        ' رسم مستطیل
        WScript.Echo "ایجاد مستطیل با مرکز (" & x & ", " & y & "), عرض " & width * 1000 & "mm و ارتفاع " & height * 1000 & "mm"
        swSketchManager.CreateCenterRectangle x/1000, y/1000, 0, (x/1000) + (width/2), (y/1000) + (height/2), 0
        
    Case "line"
        ' پارامترهای خط
        If params.Exists("x1") Then x1 = params("x1") Else x1 = 0
        If params.Exists("y1") Then y1 = params("y1") Else y1 = 0
        If params.Exists("x2") Then x2 = params("x2") Else x2 = 0.03  ' 30mm
        If params.Exists("y2") Then y2 = params("y2") Else y2 = 0.02  ' 20mm
        
        ' رسم خط
        WScript.Echo "ایجاد خط از نقطه (" & x1 & ", " & y1 & ") به نقطه (" & x2 & ", " & y2 & ")"
        swSketchManager.CreateLine x1/1000, y1/1000, 0, x2/1000, y2/1000, 0
        
    Case Else
        WScript.Echo "خطا: نوع شکل '" & shape & "' پشتیبانی نمی‌شود"
        WScript.Echo "شکل‌های پشتیبانی شده: circle, rectangle, line"
        WScript.Quit(1)
End Select

' بستن اسکچ
swSketchManager.InsertSketch True

' نمایش در نمای ایزومتریک
swModel.ShowNamedView2 "*Isometric", 7
swModel.ViewZoomtofit2

WScript.Echo "SUCCESS: اسکچ با موفقیت ایجاد شد"
WScript.Quit(0) 