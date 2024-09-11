VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSales 
   Caption         =   "Panel de ventas"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12180
   OleObjectBlob   =   "FormSales.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub ComboBoxPayment_Change()

End Sub

Private Sub UserForm_Initialize()
    ' Valores por defecto de los objetos del formulario
    
    LabelTicketID.Caption = Format(GetNextTicketID, "00000")
    Call ComboBoxProductCode_SetCodes
    ComboBoxProductCode.Value = ""
    ComboBoxProductCode.SetFocus
    TextBoxQuantity.Value = 1
    SpinButtonQuantity.Value = 1
    TextBoxDate.Value = Format(Date, "dd/mm/yyyy")
    TextBoxTime.Value = Format(Time, "hh:mm")
    Call ComboBoxPayment_SetMethods
    ComboBoxPayment.Value = "EFECTIVO"
    TextBoxTax.Value = GetTaxFromSheet
    Call ListBoxSales_UpdateList
End Sub



Private Function GetNextID() As Long
    ' Devuelve el max ID de la tabla incrementado en 1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblVentas")
    
    Dim rng As Range
    Set rng = tbl.ListColumns(1).DataBodyRange
    Dim maxID As Long
    
    If rng Is Nothing Then
        maxID = 0
    Else
        maxID = Application.WorksheetFunction.Max(rng)
    End If
    
    GetNextID = maxID + 1
End Function



Private Function GetNextTicketID() As Long
    ' Devuelve el max TicketID de la tabla incrementado en 1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblVentas")
    
    Dim rng As Range
    Set rng = tbl.ListColumns(2).DataBodyRange
    Dim maxID As Long
    
    If rng Is Nothing Then
        maxID = 0
    Else
        maxID = Application.WorksheetFunction.Max(rng)
    End If
    
    GetNextTicketID = maxID + 1
End Function



Private Sub ComboBoxProductCode_SetCodes()
    ' Establece en el ComboBox los codigos de productos cuya cantidad es mayor a cero
    
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
        ' Verificar la cantidad es mayor a cero
        If cell.Offset(0, 4).Value > 0 Then
            ComboBoxProductCode.AddItem (cell.Value)
        End If
    Next cell
End Sub



Private Sub ComboBoxPayment_SetMethods()
    ' Establece en el ComboBox los metodos de pago unicos registrados en ventas anteriores
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblVentas")
    
    ' Verificar si la tabla tiene filas
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = tbl.ListColumns(1).DataBodyRange
    
    ComboBoxPayment.Clear
    Dim cell As Range
    Dim uniquePayments As Collection
    Set uniquePayments = New Collection
    
    ' Agregar solo valores unicos
    ' Omitir error al agregar duplicados
    On Error Resume Next
    For Each cell In rng
        uniquePayments.Add cell.Offset(0, 5).Value, CStr(cell.Offset(0, 5).Value)
    Next cell
    On Error GoTo 0
    
    ' Poblar el ComboBox con los valores unicos
    Dim payment As Variant
    For Each payment In uniquePayments
        ComboBoxPayment.AddItem payment
    Next payment
End Sub



Private Function GetTaxFromSheet() As Double
    ' Obtener tasa de interes desde la hoja principal
    
    Dim tax As Double
    tax = CInt(Split(ThisWorkbook.Sheets("Principal").TextBoxTax.Value, "%")(0))
    GetTaxFromSheet = tax
End Function



Private Sub ListBoxSales_UpdateList()
    ' Se muestran los registros de ventas en el listbox
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer

    Set ws = ThisWorkbook.Sheets("Ventas")
    Set tbl = ws.ListObjects("TblVentas")
    
    With ListBoxSales
        .Clear
        .RowSource = ""
        
        ' Configurar encabezados
        .ColumnCount = 8
        .ColumnWidths = "45;45;70;60;60;100;50;110"
        
        ' Agregar encabezados de columna
        .AddItem ""
        .List(0, 0) = "ID"
        .List(0, 1) = "IDTicket"
        .List(0, 2) = "Fecha"
        .List(0, 3) = "Hora"
        .List(0, 4) = "Codigo"
        .List(0, 5) = "Pago"
        .List(0, 6) = "Interes"
        .List(0, 7) = "Total"
        
        ' Agregar registros de la tabla
        For i = 1 To tbl.ListRows.Count
            .AddItem
            .List(i, 0) = Format(tbl.ListColumns(1).DataBodyRange.Cells(i, 1).Value, "00000")
            .List(i, 1) = Format(tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value, "00000")
            .List(i, 2) = Format(tbl.ListColumns(3).DataBodyRange.Cells(i, 1).Value, "dd/mm/yyyy")
            .List(i, 3) = Format(tbl.ListColumns(4).DataBodyRange.Cells(i, 1).Value, "hh:mm")
            .List(i, 4) = tbl.ListColumns(5).DataBodyRange.Cells(i, 1).Value
            .List(i, 5) = tbl.ListColumns(6).DataBodyRange.Cells(i, 1).Value
            .List(i, 6) = Format(tbl.ListColumns(7).DataBodyRange.Cells(i, 1).Value * 100) & "%"
            .List(i, 7) = Format(tbl.ListColumns(8).DataBodyRange.Cells(i, 1).Value, "Currency")
        Next i
    End With
End Sub



Private Sub ComboBoxProductCode_Change()
    ' Se establece la cantidad maxima del producto seleccionado
    ' Se muestra la descripcion del producto
    
    If ComboBoxProductCode.Value <> "" Then
        Call SpinButtonQuantity_SetMax
        Call LabelDescription_Set
    End If
End Sub



Private Sub SpinButtonQuantity_SetMax()
    ' Buscar el producto seleccionado en el ComboBox y establecer el valor máximo
    ' del SpinButton de cantidad SpinButtonQuantity
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    ' Verificar si la tabla tiene filas
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim productCode As String
    productCode = Me.ComboBoxProductCode.Value

    ' Buscar el producto
    Dim productRow As Range
    Set productRow = tbl.ListColumns(1).DataBodyRange.Find(What:=productCode, LookIn:=xlValues, LookAt:=xlWhole)

    ' Si se encuentra el producto, establecer el valor maximo del SpinButton
    Dim maxQuantity As Long
    
    If Not productRow Is Nothing Then
        ' Obtener valor de cantidad en la columna 5 de la tabla de productos
        maxQuantity = productRow.Offset(0, 4).Value
        Me.SpinButtonQuantity.Max = maxQuantity
    Else
        ' Si no se encuentra el producto, establecer el valor maximo en 1
        Me.SpinButtonQuantity.Max = 1
        Me.SpinButtonQuantity.Value = 1
    End If
End Sub



Private Sub LabelDescription_Set()
    ' Usar el valor del ComboBox para mostrar en el Label
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
    
    Dim cell As Range
    Dim found As Boolean
    found = False
    
    For Each cell In rng
        If cell.Value = ComboBoxProductCode.Value Then
            ' Mostrar la descripcion del producto (3ra columna de la tabla)
            LabelDescription.Caption = cell.Offset(0, 2).Value
            found = True
            Exit For ' Salir del bucle una vez encontrado
        End If
    Next cell
    
    ' Si no se encontro, limpiar el Label
    If Not found Then
        LabelDescription.Caption = ""
    End If
End Sub



Private Sub TextBoxQuantity_AfterUpdate()
    ' Verificar que no se haya ingresado un valor mayor al
    ' numero de productos disponibles
    
    If TextBoxQuantity.Value <= SpinButtonQuantity.Max Then
        SpinButtonQuantity.Value = TextBoxQuantity.Value
    Else
        MsgBox "No existen productos suficientes para la cantidad ingresada", vbExclamation
    End If
End Sub



Private Sub SpinButtonQuantity_Change()
    ' Vincular valor del SpinButton con el TextBox
    TextBoxQuantity.Value = SpinButtonQuantity.Value
End Sub



Private Sub BtNext_Click()
    ' Enviar los datos del formulario de venta al ticket
    
    If LabelDescription.Caption <> "" Then
        ' Evaluar que se haya seleccionado un producto valido
        ' si se ha encontrado descripcion
        Call ListBoxSummary_AddSale
        Call LabelTotalSale_Update
        Call UserForm_Initialize
    End If
End Sub



Private Sub ListBoxSummary_AddSale()
    ' Agregar ventas al resumen
    
    ' Buscar el precio de costo del producto y
    ' calcular el precio de venta unitario
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    Dim productCode As String
    productCode = ComboBoxProductCode.Value
    Dim taxRate As Double
    taxRate = GetTaxFromSheet / 100
    Dim salePrice As Double
    salePrice = 0
    
    Dim row As ListRow
    
    For Each row In tbl.ListRows
        If row.Range(1).Value = productCode Then
            Dim costPrice As Double
            costPrice = row.Range(1, 6).Value
            salePrice = costPrice * (1 + taxRate)
            Exit For
        End If
    Next row
    
    ' Obtener la cantidad de ventas a agregar
    Dim quantity As Integer
    quantity = val(TextBoxQuantity.Value)
    
    Dim listIndex As Integer
    Dim i As Integer
    
    With ListBoxSummary
        ' Obtener el indice de la última fila existente en el ListBox
        listIndex = .ListCount
        
        ' Agregar registros al ListBox
        For i = 1 To quantity
            .AddItem
            .List(listIndex, 0) = ComboBoxProductCode.Value
            .List(listIndex, 1) = LabelDescription.Caption
            .List(listIndex, 2) = Format(salePrice, "Currency")
            listIndex = listIndex + 1
        Next i
    End With
End Sub



Private Sub LabelTotalSale_Update()
    ' Sumar todos los valores de la tercera columna del listbox ListBoxSummary
    
    Dim totalSale As Double
    totalSale = 0
    
    Dim i As Long
    For i = 0 To ListBoxSummary.ListCount - 1
        ' Sumar el valor de la tercera columna (precio de venta) de cada fila
        totalSale = totalSale + CDbl(ListBoxSummary.List(i, 2))
    Next i
    
    ' Mostrar el total en el LabelTotalSale
    LabelTotalSale.Caption = Format(totalSale, "Currency")
End Sub



Private Sub TextBoxTax_AfterUpdate()
    ' Al cambiar la tasa de interes se debe actualizar el valor de venta
    ' de cada producto agregado al ListBoxSummary
    
    ' Verificar que sea un valor valido
    If Not IsNumeric(TextBoxTax.Value) Then
        TextBoxTax.Value = GetTaxFromSheet
        TextBoxTax.SetFocus
        Exit Sub
    End If
    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    ' Verificar si la tabla tiene filas
    If tbl.ListRows.Count = 0 Then
        Exit Sub
    End If
    
    Dim productCode As String
    Dim row As ListRow
    Dim found As Boolean
    Dim costPrice As Double
    Dim salePrice As Double
    
    Dim taxRate As Double
    taxRate = TextBoxTax.Value / 100
    
    Dim i As Long
    For i = 0 To ListBoxSummary.ListCount - 1
        ' Obtener el codigo de producto de la primera columna del listbox
        productCode = ListBoxSummary.List(i, 0)
        found = False
        
        ' Buscar el producto en la tabla TblProductos
        For Each row In tbl.ListRows
            If row.Range(1).Value = productCode Then
                ' Obtener el precio de costo (columna 6)
                costPrice = row.Range(1, 6).Value
                found = True
                Exit For
            End If
        Next row
        
        ' Actualizar el valor de la tercera columna del ListBoxSummary
        If found = True Then
            salePrice = costPrice * (1 + taxRate)
            ListBoxSummary.List(i, 2) = salePrice
        End If
    Next i
    
    ' Actualizar el precio total de venta
    Call LabelTotalSale_Update
End Sub


Private Sub BtConfirm_Click()
    If LabelTotalSale > 0 Then
        Call TblVentas_WriteSale
        Call TblProductos_UpdateQuantity
        ListBoxSummary.Clear
        Call UserForm_Initialize
    End If
End Sub

Private Sub TblVentas_WriteSale()
    ' Recorrer el ListBoxSummary y agregar una nueva fila al principio de la tabla TblVentas
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblVentas")
    
    ' Valores constantes de una venta
    Dim ticketID As String
    ticketID = LabelTicketID.Caption
    Dim saleDate As String
    saleDate = Format(TextBoxDate.Value, "mm/dd/yyyy")
    Dim saleTime As String
    saleTime = TextBoxTime.Value
    Dim paymentMethod As String
    paymentMethod = ComboBoxPayment.Value
    Dim taxRate As Double
    taxRate = TextBoxTax.Value / 100
    
    Dim i As Long
    Dim newRow As ListRow
    For i = 0 To ListBoxSummary.ListCount - 1
        ' Insertar una nueva fila al principio de la tabla
        Set newRow = tbl.ListRows.Add(1)
        
        ' Llenar la fila con los datos correspondientes
        newRow.Range(1, 1).Value = GetNextID
        newRow.Range(1, 2).Value = ticketID
        newRow.Range(1, 3).Value = saleDate
        newRow.Range(1, 4).Value = saleTime
        newRow.Range(1, 5).Value = ListBoxSummary.List(i, 0) ' Codigo de producto
        newRow.Range(1, 6).Value = UCase(paymentMethod)
        newRow.Range(1, 7).Value = taxRate
        newRow.Range(1, 8).Value = CLng(ListBoxSummary.List(i, 2)) ' Precio de venta
    Next i
End Sub



Private Sub TblProductos_UpdateQuantity()
    ' Se registraron ventas del listbox ListBoxSummary.
    ' Restar los productos vendidos de la tabla TblProductos
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Productos")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TblProductos")
    
    Dim i As Long
    Dim productCode As String
    Dim row As ListRow
    Dim currentQty As Long
    Dim found As Boolean
    
    For i = 0 To ListBoxSummary.ListCount - 1
        ' Obtener el codigo de producto en la columna 0
        productCode = ListBoxSummary.List(i, 0)
        found = False
        
        ' Buscar el codigo de producto en la tabla TblProductos
        For Each row In tbl.ListRows
            If row.Range(1).Value = productCode Then
                ' Obtener la cantidad actual en la columna 5
                currentQty = row.Range(1, 5).Value
                ' Restar uno a la cantidad actual
                row.Range(1, 5).Value = currentQty - 1
                found = True
                Exit For
            End If
        Next row
    Next i
End Sub


Private Sub BtExit_Click()
    Unload Me
End Sub

Private Sub BtAddProduct_Click()
    Call FormProducts.Show
    Call ComboBoxProductCode_SetCodes
End Sub
