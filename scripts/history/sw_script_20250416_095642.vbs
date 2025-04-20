Option Explicit

' Main function to create a circle in SolidWorks
Sub Main()
    On Error Resume Next
    
    Dim swApp, swModel, swSketchMgr, boolstatus
    Dim center(2) As Double
    Dim radius As Double
    
    ' Initialize SolidWorks application
    Set swApp = CreateObject("SldWorks.Application")
    If swApp Is Nothing Then
        MsgBox "Failed to connect to SolidWorks", vbCritical
        Exit Sub
    End If
    
    ' Create a new part document
    Set swModel = swApp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\Part.prtdot", 0, 0, 0)
    If swModel Is Nothing Then
        MsgBox "Failed to create new part document", vbCritical
        Exit Sub
    End If
    
    ' Set circle parameters
    center(0) = 0#  ' X coordinate
    center(1) = 0#  ' Y coordinate
    center(2) = 0#  ' Z coordinate
    radius = 2      ' Radius in meters
    
    ' Start sketch on front plane
    Set swSketchMgr = swModel.SketchManager
    boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    swSketchMgr.InsertSketch True
    
    ' Create circle
    boolstatus = swSketchMgr.CreateCircle(center(0), center(1), center(2), center(0) + radius, center(1), center(2))
    
    ' Exit sketch
    swModel.ClearSelection2 True
    swSketchMgr.InsertSketch True
    
    ' Save the document (optional)
    ' swModel.SaveAs "C:\CirclePart.SLDPRT"
    
    ' Clean up
    Set swSketchMgr = Nothing
    Set swModel = Nothing
    Set swApp = Nothing
    
    MsgBox "Circle created successfully", vbInformation
End Sub

' Execute main function
Call Main
