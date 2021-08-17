VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

Public WithEvents ExcelChartEvent As Excel.Chart
Attribute ExcelChartEvent.VB_VarHelpID = -1

Private bActivated As Boolean
Private PointClicked As Long
Private XVal As Single
Private YVal As Single

Private i As Long
Private iMax As Long

'=== Get Chart XY values based on mouse pointer position ==========================================

Private Function PixelsToPointsX(dPixelsHorizontally As Long) As Double
    Dim x#
    x = 72 / GetDeviceCaps(GetDC(0), 88) / (ActiveWindow.Zoom / 100) 'points per pixel
    PixelsToPointsX = dPixelsHorizontally * x
    Call ReleaseDC(0, GetDC(0))
End Function
Private Function PixelsToPointsY(dPixelsVertically As Long) As Double
    Dim y#
    y = 72 / GetDeviceCaps(GetDC(0), 90) / (ActiveWindow.Zoom / 100) 'points per pixel
    PixelsToPointsY = dPixelsVertically * y
    Call ReleaseDC(0, GetDC(0))
End Function

Private Sub FindXval(MouseXCoord As Long)
    With ExcelChartEvent
        XVal = .Axes(xlCategory).MinimumScale + _
                (PixelsToPointsX(MouseXCoord) - .PlotArea.InsideLeft - 4) / .PlotArea.InsideWidth _
                * (.Axes(xlCategory).MaximumScale - .Axes(xlCategory).MinimumScale)
    End With
End Sub
Private Sub FindYval(MouseYCoord As Long)
    With ExcelChartEvent
        YVal = .Axes(xlValue).MaximumScale - _
                (PixelsToPointsY(MouseYCoord) - .PlotArea.InsideTop - 4) / .PlotArea.InsideHeight _
                * (.Axes(xlValue).MaximumScale - .Axes(xlValue).MinimumScale)
    End With
End Sub

'=== Chart Events ======================================

Private Sub ExcelChartEvent_Activate()
    bActivated = True
    'switch to exclude activate action from mouse down event
End Sub

Private Sub ExcelChartEvent_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    FindXval x
    FindYval y
    Application.StatusBar = "(" & Application.Round(XVal, 2) & ", " & Application.Round(YVal, 2) & ")"
End Sub

Private Sub ExcelChartEvent_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    If bActivated Then
        'if this click has activated the chart, do nothing
        bActivated = False
    ElseIf Button = xlPrimaryButton Then
        'in case of lmb click
        'check if any point has been clicked
        CheckIfPointClicked
        
        If PointClicked <> 0 Then
            'in case point is clicked, do nothing - mouse up event will change clicked point XYvalues in the data table to new one (in case it has been dragged)
        Else
            'add new point
            'XY values determined in MouseMove event
            Cells(100, 1).End(xlUp).Offset(1, 0).Value = XVal
            Cells(100, 2).End(xlUp).Offset(1, 0).Value = YVal
        End If
    End If
End Sub

Private Sub ExcelChartEvent_MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    If Button = xlPrimaryButton And PointClicked <> 0 Then
        'in case any point on the chart has been clicked, change its XY values in data table to new one
        'thanks to that point can be dragged on the chart
        FindXval x
        FindYval y
        Range("a3").Offset(PointClicked, 0).Value = XVal
        Range("b3").Offset(PointClicked, 0).Value = YVal
        PointClicked = 0
    End If
End Sub

Private Sub CheckIfPointClicked()
    PointClicked = 0
    Dim s As Series
    Dim xv As Variant
    Dim yv As Variant
    Dim p As Point
    Dim eta As Single
    Dim i As Integer
    
    eta = 0.1
    
    Set s = ExcelChartEvent.SeriesCollection("GuidingPoints")
    xv = s.XValues
    yv = s.Values
    
    For i = 1 To s.Points.Count
        If ((XVal - xv(i)) ^ 2 + (YVal - yv(i)) ^ 2) ^ 0.5 < eta Then
            PointClicked = i
        End If
    Next i
End Sub

Private Sub Worksheet_Activate()
    frmBezier.Show 0
End Sub

'=== Worksheet Events ==================================

Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Intersect(Target, Range("A:B")) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    If Not Intersect(Target, Range("B1")) Is Nothing Then
        Module1.Animate
    Else
        Module1.UpdateBezierCurve
    End If

    Application.EnableEvents = True
End Sub



