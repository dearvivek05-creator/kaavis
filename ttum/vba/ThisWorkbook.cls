VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Re-creates the Dashboard buttons every time the workbook is opened, so the
' utility still works if a button is accidentally deleted or the file is copied
' between machines.
Private Sub Workbook_Open()
    On Error Resume Next
    modTTUM.TTUM_Setup
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub
