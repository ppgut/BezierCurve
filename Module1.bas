Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
#End If

Public TimerID As LongLong

Private i As Long
Private iMax As Long

Function BeziereCurve(DataRng As Range, Optional t As Integer = 10) As Variant()
    'Returns array with 't+1' number of Beziere Curve points (either X values or Y values - depends on provided data range)
    'Should be run twice (once to get X values of Bezier Curve points and second time to get Y values)
    
    Dim u As Single
    Dim i As Integer
    Dim j As Integer
    Dim rng As Range
    Dim DataIn() As Single
    Dim DataOut() As Single
    Dim DataTemp() As Variant
    
    ReDim DataIn(DataRng.Count - 1)
    ReDim DataOut(DataRng.Count - 1)
    ReDim DataTemp(t)
    ReDim BeziereCurve(t)
    
    'Read input data to DataIn array
    i = 0
    For Each rng In DataRng
        DataIn(i) = rng.Value
        i = i + 1
    Next rng
    
    For i = 0 To t
        'each section of guiding curve is divided into 't' equal pieces
        'in the end Bezier Curve will consist of 't+1' points so 't' parameter represents the smoothness of Beziere Curve
        'u - fraction of guiding curve section
        u = i / t

        DataOut = DataIn
        
        'guiding curve has n points which divide it into n-1 sections
        'by taking one new point from each section (from 'u' position of this section), new curve is received
        'new curve has n-1 points and n-2 sections
        'these steps are repeated till we get only one point (zero sections)
        
        Do While UBound(DataOut) > 0
            For j = 0 To UBound(DataOut) - 1
                DataOut(j) = (1 - u) * DataOut(j) + u * DataOut(j + 1)
            Next j
            ReDim Preserve DataOut(0 To UBound(DataOut) - 1)
        Loop
        
        'then the loop is repeated for further 'u' positions
        'points for each 'u' position are stored in temp array
        DataTemp(i) = DataOut(0)
        
    Next i
    
    BeziereCurve = Application.WorksheetFunction.Transpose(DataTemp())

End Function

Sub UpdateBezierCurve()
    'checks the data provided in the worksheet and updates the chart if everything is ok
    
    Dim XRng As Range
    Dim YRng As Range
    
    Set XRng = Range(Range("A4"), Range("A3").End(xlDown))
    Set YRng = Range(Range("B4"), Range("B4").End(xlDown))
        
    If XRng(1) = "" Or YRng(1) = "" Then Exit Sub
    If XRng.Count < 2 Or YRng.Count < 2 Then Exit Sub
    If XRng.Count <> YRng.Count Then Exit Sub
    
    With Sheet1.ChartObjects(1).Chart
    
        With .SeriesCollection("GuidingPoints")
            .XValues = XRng
            .Values = YRng
        End With
        
        With .SeriesCollection("BezierCurve")
            .XValues = BeziereCurve(XRng, Range("B1"))
            .Values = BeziereCurve(YRng, Range("B1"))
        End With
            
    End With
    
End Sub

Sub Animate()
    i = 1
    iMax = [b1].Value
    If i <= iMax Then StartTimer
End Sub

Private Sub StartTimer()
    'Run TimerEvent every 100/1000s of a second
    Dim AnimationTime As Single
    Dim StepTime As Single
    
    AnimationTime = 3000 '[miliseconds]
    StepTime = AnimationTime / iMax
    
    TimerID = SetTimer(0, 0, StepTime, AddressOf TimerEvent)
End Sub

Private Sub StopTimer()
    On Error Resume Next
    KillTimer 0, TimerID
    On Error GoTo 0
End Sub

Private Sub TimerEvent()

    On Error Resume Next
    Application.EnableEvents = False
    
    If i = 1 Then
        [b1].Value = 1
    Else
        [b1].Value = [b1].Value + 1
    End If
    UpdateBezierCurve
    
    Application.EnableEvents = True
    
    If i >= iMax Then StopTimer
    i = i + 1
    On Error GoTo 0
    
End Sub


