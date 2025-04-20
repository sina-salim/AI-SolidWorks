Option Explicit

' اسکریپت اتصال به SolidWorks
' این اسکریپت فقط به SolidWorks متصل می‌شود و شیء اپلیکیشن را برمی‌گرداند

Function ConnectToSolidWorks()
    ' متغیرهای اصلی
    Dim swApp
    
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
    
    ' برگرداندن شیء اپلیکیشن
    Set ConnectToSolidWorks = swApp
End Function

' اجرای تابع اتصال اگر این اسکریپت به طور مستقیم اجرا شود
If WScript.ScriptName = "connect_to_sw.vbs" Then
    Dim swApp
    Set swApp = ConnectToSolidWorks()
    WScript.Echo "SUCCESS: اتصال موفقیت‌آمیز"
    WScript.Quit(0)
End If 