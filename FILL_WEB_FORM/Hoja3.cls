VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim row_sel As Long
    Dim col_sel As Long
    
    row_sel = Target.EntireRow.Row
    col_sel = Target.EntireColumn.Column
        
    ' Detect a change in this sheet "Hide" and continue with the macro
    If Hoja3.Cells(row_sel, col_sel) = "Finished" Then
        Application.EnableEvents = False
        Hoja1.Cells(row_sel, col_sel) = ""
        Application.EnableEvents = True
        MsgBox ("Work done")
    End If
     
End Sub
