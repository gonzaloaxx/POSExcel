VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProducts 
   Caption         =   "Administrador de productos"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9030.001
   OleObjectBlob   =   "FormProducts.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    ' Valores por defecto de los campos del formulario
    
    Call ComboBoxProductCode_SetCodes
    Me.ComboBoxProductCode.Value = ""
    
    TextBoxDate.Value = Format(Date, "dd/mm/yyyy")
    TextBoxDescription.Value = ""
    
    Call ComboBoxCategory_SetCodes
    Me.ComboBoxCategory.Value = ""
    
    TextBoxQuantity.Value = 1
    SpinButtonQuantity.Value = 1
    TextBoxUnitCost.Value = Format(0, "Currency")
    
    Call ListBoxProducts_UpdateList
    ComboBoxProductCode.SetFocus
End Sub



Private Sub ComboBoxProductCode_SetCodes()
    ' Establece en el ComboBox los codigos de productos ya registrados
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    ' Verificar si la tabla tiene filas
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = tbl.ListColumns(1).DataBodyRange
    
    ComboBoxProductCode.Clear
    Dim cell As Range
    
    For Each cell In rng
        ComboBoxProductCode.AddItem (cell.Value)
    Next cell
End Sub



Private Sub ComboBoxCategory_SetCodes()
    ' Establece en el ComboBox las categorias previamente registradas
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    ' Verificar si la tabla tiene filas
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = tbl.ListColumns(4).DataBodyRange
    
    ComboBoxCategory.Clear
    Dim cell As Range
    
    For Each cell In rng
        ComboBoxCategory.AddItem (cell.Value)
    Next cell
End Sub



Private Sub ListBoxProducts_UpdateList()
    ' Actualizar lista de productos en el ListBox
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer

    Set ws = ThisWorkbook.Sheets("Productos")
    Set tbl = ws.ListObjects("TblProductos")
    
    With ListBoxProducts
        .Clear
        .RowSource = ""
        
        ' Configurar encabezados
        .ColumnCount = 6
        .ColumnWidths = "45;65;120;90;40;80"
        
        ' Agregar encabezados de columna
        .AddItem ""
        .List(0, 0) = "Codigo"
        .List(0, 1) = "Fecha"
        .List(0, 2) = "Descripcion"
        .List(0, 3) = "Categoria"
        .List(0, 4) = "Unid"
        .List(0, 5) = "PCosto"
        
        ' Agregar registros de la tabla
        For i = 1 To tbl.ListRows.Count
            .AddItem
            .List(i, 0) = tbl.ListColumns(1).DataBodyRange.Cells(i, 1).Value
            .List(i, 1) = Format(tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value, "dd/mm/yyyy")
            .List(i, 2) = tbl.ListColumns(3).DataBodyRange.Cells(i, 1).Value
            .List(i, 3) = tbl.ListColumns(4).DataBodyRange.Cells(i, 1).Value
            .List(i, 4) = tbl.ListColumns(5).DataBodyRange.Cells(i, 1).Value
            .List(i, 5) = Format(tbl.ListColumns(6).DataBodyRange.Cells(i, 1).Value, "Currency")
        Next i
    End With
End Sub



Private Sub SpinButtonQuantity_Change()
    ' Enlazar valor del SpinButton con el valor que se muestra en el
    ' TextBox TextBoxQuantity
    
    TextBoxQuantity.Value = SpinButtonQuantity.Value
    TextBoxQuantity.SetFocus
End Sub



Private Sub TextBoxQuantity_Change()
    ' Enlazar valor del TextBox con el valor que se muestra en el
    ' TextBox SpinButtonQuantity
    
    If IsNumeric(TextBoxQuantity.Value) Then
        SpinButtonQuantity.Value = TextBoxQuantity.Value
    End If
End Sub



Private Sub TextBoxQuantity_AfterUpdate()
    ' Evaluar si se introdujo un valor incorrecto
    
    If Not IsNumeric(TextBoxQuantity.Value) Then
        TextBoxQuantity.Value = 1
    End If
End Sub



Private Sub ComboBoxProductCode_Change()
    ' Usar el valor del ComboBox para completar campos del formulario
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    ' Verificar si la tabla esta vacia
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = tbl.ListColumns(1).DataBodyRange
    
    Dim cell As Range
    Dim found As Boolean
    found = False
    
    For Each cell In rng
        If cell.Value = ComboBoxProductCode.Value Then
            ' Descripcion del producto
            TextBoxDescription.Value = cell.Offset(0, 2).Value
            ' Categoria de producto
            ComboBoxCategory.Value = cell.Offset(0, 3).Value
            ' Ultimo valor de costo unitario registrado en la tabla
            TextBoxUnitCost.Value = Format(cell.Offset(0, 5).Value, "Currency")
            found = True
            Exit For ' Salir del bucle una vez encontrado
        End If
    Next cell
    
    ' Si no se encontro, limpiar textbox(es)
    If Not found Then
        TextBoxDescription.Value = ""
        ComboBoxCategory.Value = ""
        TextBoxUnitCost.Value = Format(0, "Currency")
    End If
End Sub


Private Sub ComboBoxProductCode_AfterUpdate()
    ' Formateo del texto introducido a mayusculas
    Me.ComboBoxProductCode.Value = UCase(Me.ComboBoxProductCode.Value)
End Sub


Private Sub TextBoxDescription_AfterUpdate()
    ' Formateo del texto introducido a mayusculas
    TextBoxDescription.Value = UCase(TextBoxDescription.Value)
End Sub


Private Sub ComboBoxCategory_AfterUpdate()
    ' Formateo del texto introducido a mayusculas
    ComboBoxCategory.Value = UCase(ComboBoxCategory.Value)
End Sub


Private Sub TextBoxUnitCost_AfterUpdate()
    ' El campo de costo unitario admite solamente valores numericos
    
    If Not IsNumeric(TextBoxUnitCost.Value) Then
        TextBoxUnitCost.Value = Format(0, "Currency")
        TextBoxUnitCost.SetFocus
    Else
        TextBoxUnitCost.Value = Format(TextBoxUnitCost.Value, "Currency")
    End If
End Sub


Private Sub BtExit_Click()
    Unload Me
End Sub


Private Sub BtAdd_Click()
    ' Se registra una actualizacion del stock de productos
    ' Si el producto ya existe. Se actualizan los datos de los campos
    ' excepto el codigo de producto
    ' Si el producto no existe. Se agrega uno nuevo

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")

    Dim rng As Range
    Dim productCode As String
    productCode = ComboBoxProductCode.Value

    ' Verificar si la tabla tiene filas antes de definir el rango
    If tbl.ListRows.Count > 0 Then
        Set rng = tbl.ListColumns(1).DataBodyRange
    End If

    Dim foundRow As Variant
    If Not rng Is Nothing Then
        ' Buscar el código de producto en la columna 1
        foundRow = Application.Match(productCode, rng, 0)
    Else
        ' Si no hay filas, simular que no se encuentra el producto
        foundRow = CVErr(xlErrNA)
    End If

    If Not IsError(foundRow) Then
        ' Si se encuentra el producto. Actualizar datos de la tabla
        tbl.DataBodyRange(foundRow, 2).Value = Format(TextBoxDate.Value, "mm/dd/yyyy")
        tbl.DataBodyRange(foundRow, 3).Value = UCase(TextBoxDescription.Value)
        tbl.DataBodyRange(foundRow, 4).Value = UCase(ComboBoxCategory.Value)
        tbl.DataBodyRange(foundRow, 5).Value = tbl.DataBodyRange(foundRow, 4).Value + TextBoxQuantity.Value
        tbl.DataBodyRange(foundRow, 6).Value = CDbl(TextBoxUnitCost.Value)
    Else
        ' Si no se encuentra el producto. Agregar una nueva fila al final de la tabla
        With tbl.ListRows.Add(1)
            .Range(1, 1).Value = ComboBoxProductCode.Value
            .Range(1, 2).Value = Format(TextBoxDate.Value, "mm/dd/yyyy")
            .Range(1, 3).Value = UCase(TextBoxDescription.Value)
            .Range(1, 4).Value = UCase(ComboBoxCategory.Value)
            .Range(1, 5).Value = TextBoxQuantity.Value
            .Range(1, 6).Value = CDbl(TextBoxUnitCost.Value)
        End With
    End If
    
    Call UserForm_Initialize
End Sub



Private Sub ListBoxProducts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Doble click en un registro del listbox
    ' usar los datos de la fila seleccionada para completar los valores
    ' de los elementos del formulario
    
    Dim selectedIndex As Integer
    selectedIndex = Me.ListBoxProducts.listIndex
    
    ' Verifica si se ha seleccionado una fila
    If selectedIndex <> -1 Then
        ' Asigna los valores de la fila seleccionada a los widgets correspondientes
        Me.ComboBoxProductCode.Value = Me.ListBoxProducts.List(selectedIndex, 0)
        Me.TextBoxDate.Value = Me.ListBoxProducts.List(selectedIndex, 1)
        Me.TextBoxDescription.Value = Me.ListBoxProducts.List(selectedIndex, 2)
        Me.ComboBoxCategory.Value = Me.ListBoxProducts.List(selectedIndex, 3)
        Me.TextBoxQuantity.Value = Me.ListBoxProducts.List(selectedIndex, 4)
        Me.TextBoxUnitCost.Value = Me.ListBoxProducts.List(selectedIndex, 5)
        
        ComboBoxProductCode.SetFocus
    End If
End Sub


Private Sub BtDelete_Click()
    ' Eliminar el registro de un producto en la tabla "TblProductos"
    ' que coincida con el valor del ComboBoxProductCode
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    Dim rng As Range
    Set rng = tbl.ListColumns(1).DataBodyRange
    
    Dim productCode As String
    productCode = Me.ComboBoxProductCode.Value
    
    ' Buscar el codigo de producto en la columna 1
    Dim foundRow As Variant
    foundRow = Application.Match(productCode, rng, 0)
    
    If Not IsError(foundRow) Then
        ' Si se encuentra el producto, eliminar la fila correspondiente
        tbl.ListRows(foundRow).Delete
        MsgBox "Producto eliminado correctamente.", vbInformation
    Else
        MsgBox "Código de producto no encontrado.", vbExclamation
    End If
    
    Call UserForm_Initialize
End Sub


