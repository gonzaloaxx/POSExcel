
Option Explicit

Private Sub Workbook_Open()
    Call ResetPointers
    Call FormCredits.Show
End Sub


Private Sub ResetPointers()
    ' Los apuntadores de Excel vuelven a situarse en el rango A1
    ' La vista de hoja al abrir el documento es la hoja 1
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        ws.Range("A1").Select
    Next ws
    
    ThisWorkbook.Sheets("Principal").Activate
End Sub
