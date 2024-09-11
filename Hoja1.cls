
Option Explicit

Sub TextBoxTax_GetFocus()
    ThisWorkbook.Sheets("Principal").TextBoxTax.Activate
    
    TextBoxTax.SelStart = 0
    TextBoxTax.SelLength = Len(TextBoxTax)
End Sub



Private Sub TextBoxTax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    handleKeypress KeyAscii
End Sub



Private Sub handleKeypress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim keyChar As String
    keyChar = Chr(KeyAscii)
    
    ' Permitir solo numeros (0-9) y backspace (8)
    If Not (IsNumeric(keyChar) Or KeyAscii = 8) Then
        KeyAscii = 0 ' Cancelar la entrada de la tecla
    End If
End Sub
