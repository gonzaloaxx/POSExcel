Attribute VB_Name = "Modulo1"

Option Explicit

Private TextFrameNumSalesStatus As String
Private TextFrameSummaryStatus As String


Sub FormSales_OpenForm()
    Call FormSales.Show
    Call refreshWorkbook
End Sub


Sub FormCredits_OpenForm()
    Call FormCredits.Show
End Sub


Sub FormProducts_OpenForm()
    Call FormProducts.Show
    Call refreshWorkbook
    Call refreshWorkbook
End Sub


Sub TextFrameNumSales_Switch()
    ' Alternar el cuadro de texto (forma) entre diario y mensual

    If TextFrameSummaryStatus = "" Then
        TextFrameSummaryStatus = "Day"
    End If
    
    If TextFrameSummaryStatus = "Day" Then
        ThisWorkbook.Sheets("Principal").Shapes.Range(Array("TextFrameNumSales")).Select
        Selection.Formula = "=Analiticas!B8"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Name = "+mn-lt"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 20
        Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
        
        ActiveSheet.Shapes.Range(Array("RoudedRectangeNumSales")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "                Numero de Ventas" & Chr(13) & _
        "                          (Diario)"
        
        TextFrameSummaryStatus = "Month"
        Range("M7").Select
        Exit Sub
    End If
    
    If TextFrameSummaryStatus = "Month" Then
        ThisWorkbook.Sheets("Principal").Shapes.Range(Array("TextFrameNumSales")).Select
        Selection.Formula = "=Analiticas!B11"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Name = "+mn-lt"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 20
        Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
        
        ActiveSheet.Shapes.Range(Array("RoudedRectangeNumSales")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "                Numero de Ventas" & Chr(13) & _
        "                       (Mensual)"
        
        TextFrameSummaryStatus = "Day"
        Range("M7").Select
        Exit Sub
    End If
End Sub



Sub TextFrameSummary_Switch()
    ' Alternar el cuadro de texto (forma) entre diario y mensual

    If TextFrameSummaryStatus = "" Then
        TextFrameSummaryStatus = "Day"
    End If
    
    If TextFrameSummaryStatus = "Day" Then
        ThisWorkbook.Sheets("Principal").Shapes.Range(Array("TextFrameSummary")).Select
        Selection.Formula = "=Analiticas!B2"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Name = "+mn-lt"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 20
        Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
        
        ActiveSheet.Shapes.Range(Array("RoundedRectangleSummary")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "                Recaudacion" & Chr(13) & _
        "                     (Diario)"
        
        TextFrameSummaryStatus = "Month"
        Range("M3").Select
        Exit Sub
    End If
    
    If TextFrameSummaryStatus = "Month" Then
        ThisWorkbook.Sheets("Principal").Shapes.Range(Array("TextFrameSummary")).Select
        Selection.Formula = "=Analiticas!B5"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Name = "+mn-lt"
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 20
        Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
        
        ActiveSheet.Shapes.Range(Array("RoundedRectangleSummary")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "                Recaudacion" & Chr(13) & _
        "                   (Mensual)"
        
        TextFrameSummaryStatus = "Day"
        Range("M3").Select
        Exit Sub
    End If
End Sub


Private Sub refreshWorkbook()
    ThisWorkbook.RefreshAll
End Sub
