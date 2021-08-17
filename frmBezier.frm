VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBezier 
   Caption         =   "Bezier"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "frmBezier.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbActivate_Click()
    Set Sheet1.ExcelChartEvent = Sheet1.ChartObjects("Chart 1").Chart
End Sub

Private Sub cbAnimate_Click()
    Sheet1.Range("B1").Value = Sheet1.Range("B1").Value 'to trigger change event
End Sub

Private Sub cbDeleteLastPoint_Click()
    With Sheet1
        .Cells(1000, 1).End(xlUp).ClearContents
        .Cells(1000, 2).End(xlUp).ClearContents
    End With
End Sub

Private Sub UserForm_Activate()
    Me.Left = 800
    Me.Top = 200
End Sub

Private Sub UserForm_Initialize()
    
End Sub



