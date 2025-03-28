VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConsultaBancoDeDados 
   Caption         =   "Ferramentas de Custos"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "DataBaseQuery.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConsultaBancoDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub AlteraItemNoBD_Click()
    
    countExecution "AlteraItemNoBD_Click", True, "Sub", "ConsultaBancoDeDados"

    SetListItemAsDisplay
        
    UpdateRecordConsulta listBoxSelectedItem, True
    
    countExecution "AlteraItemNoBD_Click", False, "Sub", "ConsultaBancoDeDados"
    
End Sub

Private Sub SalvarItemNoBD_Click()

    'countExecution "SalvarItemNoBD_Click", True, "Sub", "ConsultaBancoDeDados"
    
    OptimizedMode True
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.activeSheet
    
    'Verify if the current sheet is a product sheet
    If ws.Range("A1").Value <> "NOME DO PAINEL>>>" Then
        MsgBox "A célula selecionada não se encontra na planilha certa. A macro será encerrada."
        Exit Sub
    End If
    
    Dim selectedRows As Range
    Dim Row As Range
    
    Set selectedRows = Selection.Rows.EntireRow
    
    For Each Row In selectedRows
        If Not IsEmpty(Row.Cells(1, 2)) Or Not IsEmpty(Row.Cells(1, 6)) Then
            
            ConvertRowToListItem Row
            
            UpdateRecordConsulta listBoxSelectedItem, True
            
        End If
    Next
    
    TextBox1.Value = "TODOS OS ITENS FORAM SALVOS NO BANCO DE DADOS COM SUCESSO"
    'TextBox2.value = listBoxSelectedItem(1, 4)
    
    'RunSearch
    
    ' Set the ListBox column widths
    'ConsultaBancoDeDados.ListBox1.ColumnWidths = "0 pt;300 pt;0 pt;50 pt;75 pt;0 pt;50 pt;0 pt;0 pt;0 pt;50 pt;0 pt;0 pt;0 pt"
    
    OptimizedMode False
    
    'countExecution "SalvarItemNoBD_Click", False, "Sub", "ConsultaBancoDeDados"
    
End Sub

Private Sub BuscarSelecionado_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.activeSheet
    
    'Verify if the current sheet is a product sheet
    If ws.Range("A1").Value <> "NOME DO PAINEL>>>" Then
        Exit Sub
    End If
    
    Dim selectedRows As Range
    
    Set selectedRows = Selection.Rows.EntireRow
    
    SearchSelectedItem selectedRows
    
    ConvertRowToListItem selectedRows
    
    DisplaySelectedItem
    
End Sub

Private Sub ListBox1_Click()
    
    SetListBoxSelectedItem ConsultaBancoDeDados.ListBox1.ListIndex
    
    DisplaySelectedItem
        
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim SelectedItem As Long
    Dim selectedRows As Range
    Dim quantity As String
    
    OptimizedMode True
    
    ' Get the selected row in the ListBox
    SelectedItem = ConsultaBancoDeDados.ListBox1.ListIndex

    ' Check if an item is selected
    If SelectedItem >= 0 Then
        ' Set the selected cell
        Set selectedRows = Selection.Rows(1).EntireRow
        
        If selectedRows.Row < 4 Then
            Exit Sub
        End If
         
        quantity = selectedRows.Cells(1, qtdeColumn).Value
        
        If quantity = "" Then
        
        ' Insert a popup asking the user for the quantity of the item
        quantity = InputBox("Inserir quantidade:", "Quantidade", 0)
        
        ' Check if the user provided a quantity
        If quantity = vbNullString Then
            Exit Sub
        ElseIf Not IsNumeric(quantity) Then
            MsgBox "Quantiadde inválida. Insira um valor numérico.", vbExclamation
            Exit Sub
        End If
        End If
        
        Application.EnableEvents = False

        ConvertListItemToRow selectedRows
        
        selectedRows.Cells(1, qtdeColumn).Value = quantity
        
        selectedRows.Cells(1, 1).Offset(1, 1).Select
        
        TextBox1.Value = ""
        TextBox2.Value = ""
        
        Application.EnableEvents = True
        
        SearchSelectedItem (Selection.Rows(1).EntireRow)
    End If
    
    OptimizedMode False
End Sub

Private Sub DeletaItemNoBD_Click()

    countExecution "DeletaItemNoBD_Click", True, "Sub", "ConsultaBancoDeDados"
    
    Dim SelectedItem As Variant
    Dim SelectedItemIndex As Long
    Dim i As Integer
    
    ReDim SelectedItem(1 To 1, 1 To 13)

    ' Get the selected index in the ListBox
    SelectedItemIndex = ConsultaBancoDeDados.ListBox1.ListIndex
    
    ' Check if an item is selected
    If SelectedItemIndex >= 0 Then
        For i = 1 To 13
            ' Get the selected item from the ListBox
            SelectedItem(1, i) = ConsultaBancoDeDados.ListBox1.List(SelectedItemIndex, i - 1)
        Next i
        ' Delete selected item from the DB
        If Left(SelectedItem(1, 2), 2) = "##" Then
            SelectedItem(1, 2) = Replace(SelectedItem(1, 2), "##", "")
            DeleteRowFromTodosTable SelectedItem
        Else
            DeleteRowFromConsultaTable SelectedItem
        End If
    End If
    
    RunSearch
    
    countExecution "DeletaItemNoBD_Click", False, "Sub", "ConsultaBancoDeDados"
    
End Sub

Private Sub SalvarTudoBD_Click()
    
    CheckForChanges
    
End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Initialize()
    'RunSearch
    GetColumnsNumbers
    SetUpAccess
    TextBox8.MultiLine = True
End Sub

Private Sub UserForm_Terminate()
    CheckForChanges
    
    SetDownAccess
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        RunSearch
    End If
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        RunSearch
    End If
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> 13 Then
        Exit Sub
    End If
    
    Dim SelectedItem As Long
    Dim selectedRows As Range
    Dim quantity As String
    
    OptimizedMode True
    
    ' Get the selected row in the ListBox
    SelectedItem = ConsultaBancoDeDados.ListBox1.ListIndex

    ' Check if an item is selected
    If SelectedItem >= 0 Then
        ' Set the selected cell
        Set selectedRows = Selection.Rows(1).EntireRow
        
        If selectedRows.Row < 4 Then
            Exit Sub
        End If
         
        quantity = selectedRows.Cells(1, qtdeColumn).Value
        
        If quantity = "" Then
        
        ' Insert a popup asking the user for the quantity of the item
        quantity = InputBox("Inserir quantidade:", "Quantidade", 0)
        
        ' Check if the user provided a quantity
        If quantity = vbNullString Then
            Exit Sub
        ElseIf Not IsNumeric(quantity) Then
            MsgBox "Quantiadde inválida. Insira um valor numérico.", vbExclamation
            Exit Sub
        End If
        End If
        
        Application.EnableEvents = False

        ConvertListItemToRow selectedRows
        
        selectedRows.Cells(1, qtdeColumn).Value = quantity
        
        selectedRows.Cells(1, 1).Offset(1, 1).Select
        
        TextBox1.Value = ""
        TextBox2.Value = ""
        
        Application.EnableEvents = True
        
        SearchSelectedItem (Selection.Rows(1).EntireRow)
    End If
    
    OptimizedMode False
End Sub

Function SetListBoxSelectedItem(ListIndex As Long)
Dim i As Integer
    
    listBoxSelected = ListIndex
    
    ReDim listBoxSelectedItem(1 To 1, 1 To 13)
    
    If ConsultaBancoDeDados.ListBox1.List(0, 0) <> "Nenhum item encontrado no banco de dados Oficial." Or IsNull(ConsultaBancoDeDados.ListBox1.List(0, 0)) Then
        For i = 0 To 12
            listBoxSelectedItem(1, i + 1) = ConsultaBancoDeDados.ListBox1.List(listBoxSelected, i)
        Next i
    End If
End Function

Function RunSearch()
    
    'countExecution "RunSearch", True, "Function", "ConsultaBancoDeDados"

' Call the search function and get the results
    Dim searchResults As Variant
    searchResults = SearchByDescritivo(TextBox1.Value, TextBox2.Value)

    ' Populate the ListBox with the search results
    ' Clear existing items in the ListBox
    ConsultaBancoDeDados.ListBox1.Clear
    
    ' Check if there are any search results
    If Not IsEmpty(searchResults) Then
        Dim i As Long
       
        ConsultaBancoDeDados.ListBox1.ColumnCount = 18
       
        ' Sort the data using QuickSort based on the second column (index 1)
        QuickSort searchResults, LBound(searchResults, 1), UBound(searchResults, 1), 2  ' 1 is the column index to be sorted

        ' Modify the seventh element (index 6) to add "R$ "
        For i = LBound(searchResults, 1) To UBound(searchResults, 1)
            searchResults(i, 7) = "R$ " & searchResults(i, 7)
        Next i
       
        ' Add data to the ListBox
        ConsultaBancoDeDados.ListBox1.List = searchResults
        
        ' Set the ListBox column widths
        ConsultaBancoDeDados.ListBox1.ColumnWidths = "0 pt;300 pt;0 pt;50 pt;75 pt;0 pt;50 pt;0 pt;0 pt;0 pt;50 pt;0 pt;0 pt;0 pt"
        
    Else
        ' If there are no search results, display a message
        ConsultaBancoDeDados.ListBox1.AddItem "Nenhum item encontrado no banco de dados Oficial."
        ConsultaBancoDeDados.ListBox1.AddItem "Se deseja salvar esse item no banco de dados oficial utilize os campos abaixo."
        ConsultaBancoDeDados.ListBox1.ColumnWidths = "500 pt"
    End If
    
     ' Set focus to ListBox1
    ListBox1.SetFocus
    
    ' Highlight the first item
    ListBox1.ListIndex = 0
    
    countExecution "RunSearch", True, "Function", "ConsultaBancoDeDados"
    
End Function

Function ConvertRowToListItem(Row As Range)
    
    ReDim listBoxSelectedItem(1 To 1, 1 To 13)
    
    listBoxSelectedItem(1, 1) = Replace(Trim(Row.Cells(1, codigoColumn)), vbLf, " ") 'Código
    listBoxSelectedItem(1, 2) = Replace(Trim(Row.Cells(1, componenteColumn)), vbLf, " ") 'Descritivo técnico
    listBoxSelectedItem(1, 3) = Replace(Trim(Row.Cells(1, descritivoColumn)), vbLf, " ") 'Descritivo comercial
    listBoxSelectedItem(1, 4) = Replace(Trim(Row.Cells(1, fabricanteColumn)), vbLf, " ") 'Fabricante
    listBoxSelectedItem(1, 5) = Replace(Trim(Row.Cells(1, modeloColumn)), vbLf, " ") 'Modelo
    listBoxSelectedItem(1, 6) = Replace(Trim(Row.Cells(1, unColumn)), vbLf, " ") 'Unidade
    If Row.Cells(1, preçoColumn) <> "" Or InStr(1, Row.Cells(1, preçoColumn), "R$") > 0 Then
        listBoxSelectedItem(1, 7) = Replace(Trim(Row.Cells(1, preçoColumn)), vbLf, " ") 'Preço
    Else
        listBoxSelectedItem(1, 7) = 0
    End If
    listBoxSelectedItem(1, 9) = Replace(Trim(Row.Cells(1, icmsColumn)), vbLf, " ") 'ICMS
    listBoxSelectedItem(1, 10) = Replace(Trim(Row.Cells(1, ipiColumn)), vbLf, " ") 'IPI
    listBoxSelectedItem(1, 8) = Replace(Trim(Row.Cells(1, pisConfinsColumn)), vbLf, " ") 'PIS/COFINS
    If Row.Cells(1, dataColumn) <> "" Then
        listBoxSelectedItem(1, 11) = Replace(Trim(Row.Cells(1, dataColumn)), vbLf, " ") 'Data da cotação
    End If

End Function

Function ConvertListItemToRow(Row As Range)
    
    Row.Cells(1, codigoColumn) = listBoxSelectedItem(1, 1) 'Código
    Row.Cells(1, componenteColumn) = listBoxSelectedItem(1, 2) 'Descritivo técnico
    Row.Cells(1, descritivoColumn) = listBoxSelectedItem(1, 3) 'Descritivo comercial
    Row.Cells(1, fabricanteColumn) = listBoxSelectedItem(1, 4) 'Fabricante
    Row.Cells(1, modeloColumn) = listBoxSelectedItem(1, 5) 'Modelo
    Row.Cells(1, unColumn) = listBoxSelectedItem(1, 6) 'Unidade
    Row.Cells(1, preçoColumn) = Replace(listBoxSelectedItem(1, 7), "R$", "") 'Preço
    Row.Cells(1, icmsColumn) = listBoxSelectedItem(1, 9) 'ICMS
    Row.Cells(1, ipiColumn) = listBoxSelectedItem(1, 10) 'IPI
    Row.Cells(1, pisConfinsColumn) = listBoxSelectedItem(1, 8) 'PIS/COFINS
    Row.Cells(1, dataColumn) = listBoxSelectedItem(1, 11) 'Data da cotação

End Function

Function SetListItemAsDisplay()
    
    ReDim listBoxSelectedItem(1 To 1, 1 To 13)
    
    listBoxSelectedItem(1, 1) = TextBox6.Value 'Código
    listBoxSelectedItem(1, 2) = TextBox7.Value 'Descritivo técnico
    listBoxSelectedItem(1, 3) = TextBox8.Value 'Descritivo comercial
    listBoxSelectedItem(1, 4) = TextBox9.Value 'Fabricante
    listBoxSelectedItem(1, 5) = TextBox10.Value 'Modelo
    listBoxSelectedItem(1, 6) = TextBox11.Value 'Unidade
    If IsNumeric(listBoxSelectedItem(1, 7)) And IsNumeric(TextBox13.Value) Then 'Preço
        listBoxSelectedItem(1, 7) = TextBox12.Value
    Else
        listBoxSelectedItem(1, 7) = 0
    End If
    If IsNumeric(listBoxSelectedItem(1, 8)) And IsNumeric(TextBox13.Value) Then 'ICMS
        listBoxSelectedItem(1, 8) = TextBox13.Value / 100
    Else
        listBoxSelectedItem(1, 8) = 0
    End If
    If IsNumeric(listBoxSelectedItem(1, 9)) And IsNumeric(TextBox14.Value) Then 'IPI
        listBoxSelectedItem(1, 9) = TextBox14.Value / 100
    Else
        listBoxSelectedItem(1, 9) = 0
    End If
    If IsNumeric(listBoxSelectedItem(1, 10)) And IsNumeric(TextBox15.Value) Then 'PIS/COFINS
        listBoxSelectedItem(1, 10) = TextBox15.Value / 100
    Else
        listBoxSelectedItem(1, 10) = 0
    End If
    If TextBox16.Value <> 0 And TextBox16.Value <> "" Then   'Data da cotação Then 'Data da cotação
        listBoxSelectedItem(1, 11) = TextBox16.Value
    Else
        listBoxSelectedItem(1, 11) = 0
    End If
    listBoxSelectedItem(1, 12) = Now 'Última atualização
End Function

Function DisplaySelectedItem()

'If Not IsEmpty(listBoxSelected) Then
        TextBox6.Value = listBoxSelectedItem(1, 1)  'Código
        TextBox7.Value = listBoxSelectedItem(1, 2) 'Descritivo técnico
        TextBox8.Value = listBoxSelectedItem(1, 3) 'Descritivo comercial
        TextBox9.Value = listBoxSelectedItem(1, 4) 'Fabricante
        TextBox10.Value = listBoxSelectedItem(1, 5)  'Modelo
        TextBox11.Value = listBoxSelectedItem(1, 6)  'Unidade
        TextBox12.Value = listBoxSelectedItem(1, 7)  'Preço
        If IsNumeric(listBoxSelectedItem(1, 8)) Then
        TextBox13.Value = listBoxSelectedItem(1, 8) * 100 'ICMS
        Else
        TextBox13.Value = ""
        End If
        If IsNumeric(listBoxSelectedItem(1, 8)) Then
        TextBox14.Value = listBoxSelectedItem(1, 9) * 100 'IPI
        Else
        TextBox14.Value = ""
        End If
        If IsNumeric(listBoxSelectedItem(1, 10)) Then
        TextBox15.Value = listBoxSelectedItem(1, 10) * 100 'PIS/COFINS
        Else
        TextBox15.Value = ""
        End If
        TextBox16.Value = listBoxSelectedItem(1, 11) 'Data da cotação
        'TextBox17.value = listBoxSelectedItem(1, 12) 'Última atualzição
        'TextBox18.value = listBoxSelectedItem(1, 13) 'Data de criação
    'End If
    
    'ReDim listBoxSelectedItem(1 To 1, 1 To 13)
End Function

Function SearchSelectedItem(selectedRows As Range)

    countExecution "SearchSelectedItem", True, "Function", "ConsultaBancoDeDados"
    
    If selectedRows.Cells(1, componenteColumn) = "" Then
        Exit Function
    End If
    
    TextBox1.Value = selectedRows.Cells(1, componenteColumn)
    
    TextBox2.Value = selectedRows.Cells(1, fabricanteColumn)
    
    RunSearch
    
End Function

Function CheckForChanges()
    Dim lastRow As Long

    Dim ws As Worksheet
    Set ws = Application.activeSheet

' Lock the autocheck
    If ThisWorkbook.Sheets("S.PROP").Range("A1").Value <> "" Or ws.Range("C1").Value = "" Then
        Exit Function
    End If

Dim temp As Double
temp = Timer

    'Verify if the current sheet is a product sheet
    If ws.Range("A1").Value <> "NOME DO PAINEL>>>" Then
        'MsgBox "A célula selecionada não se encontra na planilha certa. A macro será encerrada."
        GoTo ExitSubrutine
    End If

    OptimizedMode True

    ' Find the last row with a number in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim selectedRows As Range
    Dim Row As Range

    Set selectedRows = ws.Range("B4:B" & lastRow).EntireRow

    For Each Row In selectedRows
        If Not IsEmpty(Row.Cells(1, 2)) Then

            ConsultaBancoDeDados.ConvertRowToListItem Row

            'Add selected item to the DB
            AddItemToTodosTable listBoxSelectedItem, ws, Row

        End If
        
        If Not IsEmpty(Row.Cells(1, 2)) Or Not IsEmpty(Row.Cells(1, 6)) Then

            ConsultaBancoDeDados.ConvertRowToListItem Row

            'Add selected item to the DB
            AddItemToConsultaTable listBoxSelectedItem, ws, Row, True

        End If
    Next

ExitSubrutine:
    OptimizedMode False

Debug.Print "The time it takes to verify the sheet is: " & Timer - temp

End Function

Function QuickSort(arr As Variant, low As Long, high As Long, columnIndex As Long)
    Dim pivot As Variant
    Dim temp As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long

    i = low
    j = high
    pivot = arr((low + high) \ 2, columnIndex)

    Do While i <= j
        Do While arr(i, columnIndex) < pivot
            i = i + 1
        Loop

        Do While arr(j, columnIndex) > pivot
            j = j - 1
        Loop

        If i <= j Then
            ' Swap the values in the specified column
            For k = LBound(arr, 2) To UBound(arr, 2)
                temp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = temp
            Next k

            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSort arr, low, j, columnIndex
    If i < high Then QuickSort arr, i, high, columnIndex
End Function

Private Sub ExportarResultados_Click()

    countExecution "ExportarResultados_Click", True, "Sub", "ConsultaBancoDeDados"

    OptimizedMode True
    
    Dim i, j As Long
    
    ' Create a new Excel workbook and set a reference to it
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    ' Create a new workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = xlApp.Workbooks.Add
    
    ' Create a new worksheet in the workbook
    Dim xlWorksheet As Object
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    ' Headers in order
    Dim headers As Variant
    headers = Array("Código", "Descritivo técnico", "Descritivo comercial", _
                    "Fabricante", "Modelo", "Unidade", "Preço", "ICMS", _
                    "IPI", "PIS/COFINS", "Data da cotação", "Última atualização", _
                    "Data de criação", "Custo Unitário")
    
    ' Write headers to the first row of the worksheet
    For i = LBound(headers) To UBound(headers)
        xlWorksheet.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' Get the data from ConsultaBancoDeDados.ListBox1.List
    Dim data As Variant
    data = ConsultaBancoDeDados.ListBox1.List
    
    ' Write data to the worksheet
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            xlWorksheet.Cells(i + 2, j + 1).Value = data(i, j)
        Next j
    Next i
    
    ' Make Excel visible and activate the new workbook
    xlApp.Visible = True
    xlWorkbook.Activate
    
    ' Clean up
    Set xlApp = Nothing
    Set xlWorkbook = Nothing
    Set xlWorksheet = Nothing
    
    OptimizedMode False
    
    countExecution "ExportarResultados_Click", False, "Sub", "ConsultaBancoDeDados"
    
End Sub
