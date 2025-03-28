Attribute VB_Name = "Tools"
Sub CorrigirPlanilha()
    OptimizedMode False
    
    Freeze
    
    GetColumnsNumbers
    
    countExecution "CorrigirPlanilha", False, "Sub", "Tools"
End Sub

Private Sub GerarPedido()
    Dim ws As Worksheet
    Dim listaPedido As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim index As Integer
    
    OptimizedMode True
    
    ' Set the first row for items list sheets
    firstRowWs = 4
    
    On Error Resume Next
    Set listaPedido = ThisWorkbook.Sheets("Pedido")
    On Error GoTo 0
    
    If listaPedido Is Nothing Then
        Set listaPedido = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("S.PROP"))
        listaPedido.Name = "Pedido"
    Else
        listaPedido.Cells.Clear ' Clear the sheet
    End If
    
    ' Loop through each worksheet
    For index = 1 To 30
        Set ws = ThisWorkbook.Sheets("" & index)
        ' Check if the first cell in A1 is "NOME DO PAINEL>>>"
        If ws.Range("A1").Value = "NOME DO PAINEL>>>" And Not IsEmpty(ws.Range("C1")) And ws.Range("Q1") > 0 Then
            lastRowWs = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Check if the last row is greater than 4 (to avoid copying headers)
            If lastRowWs > firstRowWs Then
                ' Put value from worksheet cell C1 in column A of "Lista de Materiais" and merge columns A to E
                With listaDeMateriais
                    .Cells(lastRow, 1).Value = ws.Range("C1").Value
                    .Range(.Cells(lastRow, 1), .Cells(lastRow, columnsQuant)).Merge
                    With .Cells(lastRow, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 9
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThick
                    End With
                    lastRow = lastRow + 1
                End With
            End If
        End If
    Next index
    
End Sub

Sub GerarListaDeMateriais()
    Dim ws As Worksheet
    Dim listaDeMateriais As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim fabricanteFilter As String
    Dim userInput As String
    Dim newRow As Range
    Dim index As Integer
    
    OptimizedMode True
    
    ' Set the first row for items list sheets
    firstRowWs = 4
    
    ' Locate or create "Lista de Materiais" sheet
    On Error Resume Next
    Set listaDeMateriais = ThisWorkbook.Sheets("Lista de Materiais")
    On Error GoTo 0
    
    If listaDeMateriais Is Nothing Then
        Set listaDeMateriais = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Materiais"))
        listaDeMateriais.Name = "Lista de Materiais"
    Else
        listaDeMateriais.Cells.Clear ' Clear the sheet
    End If
    
    ' Find the last used row in the sheet
    lastRow = listaDeMateriais.Cells(listaDeMateriais.Rows.Count, "A").End(xlUp).Row
    
    ' Ask user if they want to filter by manufacturer
    comercialOnly = MsgBox("Você quer listar todos os componentes?" & vbCrLf & "Itens com quantidades zeradas não são listados.", vbYesNo + vbQuestion, "Filtrar Componentes?")
    
    If comercialOnly = vbNo Then
        ' Ask user if they want to filter by manufacturer
        comercialOnly = True
        columnsQuant = 6
    Else
        comercialOnly = False
        columnsQuant = 7
            
        userInput = MsgBox("Você quer filtrar algum fabricante?" & vbCrLf & "Itens com quantidades zeradas não são listados.", vbYesNo + vbQuestion, "Filtrar Fabricante?")
        If userInput = vbYes Then
            fabricanteFilter = InputBox("Informe o fabricante para filtrar:", "Filtrar Fabricante")
        End If
    End If
    
    ' Loop through each worksheet
    For index = 1 To 30
        Set ws = ThisWorkbook.Sheets("" & index)
        ' Check if the first cell in A1 is "NOME DO PAINEL>>>"
        If ws.Range("A1").Value = "NOME DO PAINEL>>>" And Not IsEmpty(ws.Range("C1")) And ws.Range("Q1") > 0 Then
            lastRowWs = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Check if the last row is greater than 4 (to avoid copying headers)
            If lastRowWs > firstRowWs Then
                ' Put value from worksheet cell C1 in column A of "Lista de Materiais" and merge columns A to E
                With listaDeMateriais
                    .Cells(lastRow, 1).Value = ws.Range("C1").Value
                    .Range(.Cells(lastRow, 1), .Cells(lastRow, columnsQuant)).Merge
                    With .Cells(lastRow, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThick
                    End With
                    lastRow = lastRow + 1
                End With
            
            
            If comercialOnly Then
                For i = firstRowWs To lastRowWs
                If ws.Range("C" & i).Value <> "" And ws.Range("H" & i).Value > 0 And ws.Range("R" & i).Value > 0 And (fabricanteFilter = "" Or InStr(1, ws.Range("D" & i).Value, fabricanteFilter, vbTextCompare) > 0) Then
                    Set newRow = listaDeMateriais.Cells(lastRow, 1).EntireRow
                    
                    ' Copy values from the found rows to the new rows in "Materiais" table
                    newRow.Cells(1, 1).Resize(1, columnsQuant).Value = ws.Range("C" & i & ":H" & i).Value
                    If newRow.Cells(1, 3).Value = "" Then
                        newRow.Cells(1, 3).Value = "-"
                    End If
                    With newRow.Cells(1, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    lastRow = lastRow + 1
                End If
                Next i
            Else
                For i = firstRowWs To lastRowWs
                If ws.Range("B" & i).Value <> "" And ws.Range("H" & i).Value > 0 And ws.Range("R" & i).Value > 0 And (fabricanteFilter = "" Or InStr(1, ws.Range("D" & i).Value, fabricanteFilter, vbTextCompare) > 0) Then
                    Set newRow = listaDeMateriais.Cells(lastRow, 1).EntireRow
                    
                    ' Copy values from the found rows to the new rows in "Materiais" table
                    newRow.Cells(1, 1).Resize(1, columnsQuant).Value = ws.Range("B" & i & ":H" & i).Value
                    
                    If newRow.Cells(1, 5).Value <> "" And newRow.Cells(1, 4).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 4).Value & " - " & newRow.Cells(1, 5).Value
                    ElseIf newRow.Cells(1, 5).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 5).Value
                    End If
                    
                    With newRow.Cells(1, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    lastRow = lastRow + 1
                End If
                Next i
            End If
            
                ' Put value "MATERIAIS DE FORNECIMENTO CLIENTE" in column A of "Lista de Materiais" and merge columns A to E
                With listaDeMateriais
                    .Cells(lastRow, 1).Value = "MATERIAIS DE FORNECIMENTO CLIENTE"
                    .Range(.Cells(lastRow, 1), .Cells(lastRow, columnsQuant)).Merge
                    With .Cells(lastRow, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Font.Color = RGB(255, 0, 0)
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThick
                    End With
                    lastRow = lastRow + 1
                End With
                
            If comercialOnly Then
                ' List items with the value 0 that the client must provide
                For i = firstRowWs To lastRowWs
                If ws.Range("B" & i).Value <> "" And ws.Range("H" & i).Value > 0 And ws.Range("R" & i).Value = 0 And ws.Range("G" & i).Value <> "CJ" And (fabricanteFilter = "" Or InStr(1, ws.Range("D" & i).Value, fabricanteFilter, vbTextCompare) > 0) Then
                    Set newRow = listaDeMateriais.Cells(lastRow, 1).EntireRow
                    
                    ' Copy values from the found rows to the new rows in "Materiais" table
                    newRow.Cells(1, 1).Resize(1, columnsQuant).Value = ws.Range("C" & i & ":H" & i).Value
                    
                    If newRow.Cells(1, 3).Value = "" Then
                        newRow.Cells(1, 3).Value = "-"
                    End If
                    
                    If newRow.Cells(1, 5).Value <> "" And newRow.Cells(1, 4).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 4).Value & " - " & newRow.Cells(1, 5).Value
                    ElseIf newRow.Cells(1, 5).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 5).Value
                    End If
                    
                    With newRow.Cells(1, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Font.Color = RGB(255, 0, 0)
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    lastRow = lastRow + 1
                End If
                Next i
            Else
                ' List items with the value 0 that the client must provide
                For i = firstRowWs To lastRowWs
                If ws.Range("B" & i).Value <> "" And ws.Range("H" & i).Value > 0 And ws.Range("R" & i).Value = 0 And ws.Range("G" & i).Value <> "CJ" And (fabricanteFilter = "" Or InStr(1, ws.Range("D" & i).Value, fabricanteFilter, vbTextCompare) > 0) Then
                    Set newRow = listaDeMateriais.Cells(lastRow, 1).EntireRow
                    
                    ' Copy values from the found rows to the new rows in "Materiais" table
                    newRow.Cells(1, 1).Resize(1, columnsQuant).Value = ws.Range("B" & i & ":H" & i).Value
                    
                    If newRow.Cells(1, 5).Value <> "" And newRow.Cells(1, 4).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 4).Value & " - " & newRow.Cells(1, 5).Value
                    ElseIf newRow.Cells(1, 5).Value <> "" Then
                        newRow.Cells(1, 4).Value = newRow.Cells(1, 5).Value
                    End If
                    
                    With newRow.Cells(1, 1).Resize(1, columnsQuant)
                        .Font.Name = "Calibri"
                        .Font.Size = 10
                        .Font.Color = RGB(255, 0, 0)
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    lastRow = lastRow + 1
                End If
                Next i
            End If
            
            End If
        End If
    Next index
    
    If comercialOnly And fabricanteFilter = "" Then
        ' Delete column D
        listaDeMateriais.Columns("D").Delete
    ElseIf fabricanteFilter <> "" Then
        ' Delete column B and E
        listaDeMateriais.Columns("D").Delete
        listaDeMateriais.Columns("B").Delete
    Else
        ' Delete column B and E
        listaDeMateriais.Columns("E").Delete
        listaDeMateriais.Columns("B").Delete
    End If
    
    ' Autosize columns A to E
    listaDeMateriais.Columns("A:E").AutoFit
    listaDeMateriais.Range("B:E").HorizontalAlignment = xlCenter
    listaDeMateriais.Range("B:E").VerticalAlignment = xlVAlignCenter
    ' Check if the current width is greater than 50 points
    If listaDeMateriais.Columns("A").ColumnWidth > 50 Then
        ' If it is, set the width to 50 points
        listaDeMateriais.Columns("A").ColumnWidth = 50
    End If
    
    OptimizedMode False

End Sub

Private Sub Freeze()

countExecution "Freeze", True, "Sub", "Tools"

OptimizedMode True

For i = 1 To 30
    Worksheets("" & i).Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 2
        .SplitRow = 3
        .FreezePanes = True
    End With
Next

OptimizedMode False

countExecution "Freeze", False, "Sub", "Tools"

End Sub

Function CheckDate()
    Dim fs As Object
    Dim f As Object
    Dim creationDate As Date
        
    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    creationDate = f.DateCreated
    
    If creationDate <= Date - 7 Then
        ThisWorkbook.Sheets("S.PROP").Range("A1").Value = "Atualização automática desativada"
    End If
End Function

Sub CopiarPlanilha()

    countExecution "CopiarPlanilha", True, "Sub", "Tools"
    
    OptimizedMode True
    
    ' Check worksheet version
    Dim activeVersion As String
    Dim inactiveVersion As String
    Dim activeWorkbook As Workbook
    Dim inactiveWorkbook As Workbook
    Dim inactiveSheet As Worksheet
    Dim activeSheet As Worksheet
    Dim columnHeader As Range
    Dim result As Integer
    Dim Workbook As Workbook
    Dim Sheet As Worksheet
    Dim found As Boolean
    Dim versionModule As CodeModule
    Dim versionFunction As String
    Dim sheetIndex As Integer

    Set activeWorkbook = ThisWorkbook
    
    For Each Workbook In Workbooks
        found = False
            
        Dim sheet1Exists As Boolean
        sheet1Exists = False
        
        On Error Resume Next
        sheet1Exists = Workbook.Sheets("1").Range("A1").Value = "NOME DO PAINEL>>>"
        On Error GoTo 0
                        
        If Workbook.Name <> activeWorkbook.Name And Workbook.Name <> "PERSONAL.XLSB" And sheet1Exists Then
            ' Check if the sheet exists in the inactive workbook
            On Error Resume Next
            Set inactiveWorkbook = Workbook
            On Error GoTo 0
                                
            If Not inactiveWorkbook Is Nothing Then
                ' Sheet found in the inactive workbook
                found = True
                
                result = MsgBox("A pasta de trabalho encontrada foi:" & vbCrLf & inactiveWorkbook.Name & _
                        vbCrLf & "Deseja copiar todas as planilhas para esta pasta de trabalho?" & vbCrLf & vbCrLf & _
                        "Clicar em 'Não' vai copiar apenas a planilha ativa.", vbYesNoCancel + vbQuestion, "Confirmação")
                
                If result = vbCancel Then
                    found = False
                Else
                    Exit For
                End If
            End If
        End If
    Next Workbook
    
    If inactiveWorkbook Is Nothing Or Not found Then
        MsgBox "Não foi encontrada outra planilha da qual copiar"
        GoTo ExitSub
    End If
    
    ' Get version of workbooks
    activeVersion = VersionAndUpdate.CheckVersion
    
    On Error Resume Next
    Set versionModule = inactiveWorkbook.VBProject.VBComponents("VersionAndUpdate").CodeModule
    If Not versionModule Is Nothing Then
        versionFunction = "VersionAndUpdate.CheckVersion"
        inactiveVersion = Application.Run("'" & inactiveWorkbook.Name & "'!" & versionFunction)
    Else
        inactiveVersion = "0.0.0" ' Function not found, return default version
    End If

' Check the user's response
If result = vbYes Then
     ' Loop through each sheet in the active workbook
    For sheetIndex = 1 To 30
        ' Set the active sheet
        Set activeSheet = activeWorkbook.Worksheets("" & sheetIndex)
        
        ' Try to find the corresponding sheet in the inactive workbook
            found = False
            Set inactiveSheet = Nothing
            If inactiveWorkbook.Name <> activeWorkbook.Name Then
                ' Check if the sheet exists in the inactive workbook
                On Error Resume Next
                Set inactiveSheet = inactiveWorkbook.Worksheets("" & sheetIndex)
                On Error GoTo 0
                
                If Not inactiveSheet Is Nothing Then
                    ' Sheet found in the inactive workbook
                    found = True
                    Debug.Print "The inactiveSheet name found was: " & inactiveSheet.Name
                End If
            End If
        
        If Not inactiveSheet Is Nothing Then
        If found And inactiveSheet.Range("C1").Value <> "" And inactiveSheet.Name = activeSheet.Name And inactiveSheet.Range("A1").Value = "NOME DO PAINEL>>>" Then
        ' Call the copying rutine based on the sheets versions
        If InStr(1, inactiveVersion, "0.0.0", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            
            ' Check if the column header "DESCRITIVO" exists in row 3
            Set columnHeader = inactiveSheet.Rows(3).Find(What:="DESCRITIVO", LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not columnHeader Is Nothing Then
                Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
                CopyBetaToC0p1 activeSheet, inactiveSheet
            Else
                Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
                CopyV0p0ToC0p1 activeSheet, inactiveSheet
            End If
        ElseIf InStr(1, inactiveVersion, "Beta", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
            CopyBetaToC0p1 activeSheet, inactiveSheet
        ElseIf InStr(1, inactiveVersion, "0.1.", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
            CopyV0p1ToC0p1 activeSheet, inactiveSheet
        Else
            MsgBox "Não é possível copiar a versão V" & inactiveVersion & " para a versão V" & activeVersion
        End If
        End If
        End If
    Next sheetIndex

ElseIf result = vbNo Then
    Set activeSheet = activeWorkbook.activeSheet

    On Error Resume Next
    Set inactiveSheet = inactiveWorkbook.activeSheet
    On Error GoTo 0
        
    If Not inactiveSheet Is Nothing Then
        If inactiveSheet.Range("C1").Value <> "" Then
        ' Call the copying rutine based on the sheets versions
        If InStr(1, inactiveVersion, "0.0.0", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            ' Check if the column header "DESCRITIVO" exists in row 3
            Set columnHeader = inactiveSheet.Rows(3).Find(What:="DESCRITIVO", LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not columnHeader Is Nothing Then
                'Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
                CopyBetaToC0p1 activeSheet, inactiveSheet
            Else
                'Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
                CopyV0p0ToC0p1 activeSheet, inactiveSheet
            End If
        ElseIf InStr(1, inactiveVersion, "Beta", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            'Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
            CopyBetaToC0p1 activeSheet, inactiveSheet
        ElseIf InStr(1, inactiveVersion, "0.1.", vbTextCompare) > 0 And InStr(1, activeVersion, "0.1.", vbTextCompare) > 0 Then
            'Debug.Print "Copying sheet " & inactiveSheet.Name & " FROM " & inactiveWorkbook.Name & " TO " & activeSheet.Name & " FROM " & activeWorkbook.Name
            CopyV0p1ToC0p1 activeSheet, inactiveSheet
        Else
            MsgBox "Não é possível copiar a versão V" & inactiveVersion & " para a versão V" & activeVersion
        End If
        End If
    End If
End If

ExitSub:
    
    OptimizedMode False
    
    countExecution "CopyOldSheet", False, "Sub", "Tools"
    
End Sub

Function CopyV0p1ToC0p1(activeSheet As Worksheet, inactiveSheet As Worksheet)

    Dim lastRowInactiveSheet As Long
    Dim lastRowActiveSheet As Long
            
    If activeSheet.Range("C1") = "" Then
        activeSheet.Range("C1") = inactiveSheet.Range("C1")
    End If
    
    lastRowActiveSheet = activeSheet.Cells(activeSheet.Rows.Count, "A").End(xlUp).Row
    lastRowInactiveSheet = inactiveSheet.Cells(inactiveSheet.Rows.Count, "A").End(xlUp).Row
            
    'Copy assembly hours
    ' Copy rows 105 to 109 from column F of InactiveSheet
    inactiveSheet.Range("H" & lastRowInactiveSheet + 2 & ":H" & lastRowInactiveSheet + 6).Copy

    ' Paste into rows 105 to 109 of column G in ActiveSheet
    activeSheet.Range("H" & lastRowActiveSheet + 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Clear the clipboard
    Application.CutCopyMode = False
    
    ' Copy lines 4 to last row from the inactive sheet
    inactiveSheet.Rows("4:" & lastRowInactiveSheet).Copy
            
    ' Insert the copied rows below row 4 in the active sheet
    activeSheet.Rows(5).Insert Shift:=xlDown
        
    ' Clear the clipboard
    Application.CutCopyMode = False
            
    activeSheet.Range("B4").Resize(1, 13).Copy
    activeSheet.Range("B5").Resize(lastRowInactiveSheet, 13).PasteSpecial Paste:=xlPasteFormats
            
    activeSheet.Range("O4").Resize(1, 20).Copy
    activeSheet.Range("O5").Resize(lastRowInactiveSheet, 20).PasteSpecial
    
    activeSheet.Range("A4").EntireRow.Delete
    activeSheet.Range("A" & lastRowInactiveSheet + 1).Resize(lastRowActiveSheet - 4, 1).EntireRow.Delete

    ' Autosize columns A to AH
    activeSheet.Columns("A:AH").AutoFit
    
    ' Check if the current width is greater than 50 points
    For Each col In activeSheet.Columns
        ' Check if column width is greater than 50
        If col.ColumnWidth > 50 Then
            ' If so, set the column width to 50
            col.ColumnWidth = 50
        End If
    Next col

End Function

Function CopyBetaToC0p1(activeSheet As Worksheet, inactiveSheet As Worksheet)
    
    Dim lastRowInactiveSheet As Long
    Dim lastRowActiveSheet As Long
    
    If activeSheet.Range("C1") = "" Then
        activeSheet.Range("C1") = inactiveSheet.Range("C1")
    End If
    
    lastRowActiveSheet = activeSheet.Cells(activeSheet.Rows.Count, "A").End(xlUp).Row
    lastRowInactiveSheet = inactiveSheet.Cells(inactiveSheet.Rows.Count, "A").End(xlUp).Row
            
    'Copy assembly hours (only works if both sheets have 100 items lines)
    ' Copy rows 105 to 109 from column F of InactiveSheet
    inactiveSheet.Range("G" & lastRowInactiveSheet + 2 & ":G" & lastRowInactiveSheet + 6).Copy

    ' Paste into rows 105 to 109 of column G in ActiveSheet
    activeSheet.Range("H" & lastRowActiveSheet + 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Clear the clipboard
    Application.CutCopyMode = False
            
    ' Copy lines 4 to 103 from the inactive sheet
    inactiveSheet.Rows("4:" & lastRowInactiveSheet).Copy
            
    ' Insert the copied rows above row 5 in the active sheet
    Application.DisplayAlerts = False
    activeSheet.Rows(5).Insert Shift:=xlDown
    Application.DisplayAlerts = True
            
    ' Clear the clipboard
    Application.CutCopyMode = False
            
    ' Insert empty cells on the left of column C in the pasted lines
    activeSheet.Range("F5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
    activeSheet.Range("M5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
    activeSheet.Range("N5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
            
    activeSheet.Range("A" & lastRowActiveSheet + lastRowInactiveSheet - 3).EntireRow.Copy
    activeSheet.Range("A5:N" & lastRowInactiveSheet + 1).PasteSpecial xlPasteFormats
            
    ' Clear the clipboard
    Application.CutCopyMode = False
            
    activeSheet.Range("O4").Resize(1, 20).Copy
    activeSheet.Range("O5").Resize(lastRowInactiveSheet - 3, 20).PasteSpecial
        
    ' Clear the clipboard
    Application.CutCopyMode = False
        
    activeSheet.Range("A4").EntireRow.Delete
    activeSheet.Range("A" & lastRowInactiveSheet + 1).Resize(lastRowActiveSheet - 4, 1).EntireRow.Delete
           
    ' Autosize columns A to AH
    activeSheet.Columns("A:AH").AutoFit
    
    ' Check if the current width is greater than 50 points
    For Each col In activeSheet.Columns
        ' Check if column width is greater than 50
        If col.ColumnWidth > 50 Then
            ' If so, set the column width to 50
            col.ColumnWidth = 50
        End If
    Next col
    
End Function

Function CopyV0p0ToC0p1(activeSheet As Worksheet, inactiveSheet As Worksheet)
    
    If activeSheet.Range("C1") = "" Then
        activeSheet.Range("C1") = inactiveSheet.Range("C1")
    End If
            
    lastRowActiveSheet = activeSheet.Cells(activeSheet.Rows.Count, "A").End(xlUp).Row
    lastRowInactiveSheet = inactiveSheet.Cells(inactiveSheet.Rows.Count, "A").End(xlUp).Row
            
    'Copy assembly hours (only works if both sheets have 100 items lines)
    ' Copy rows 105 to 109 from column F of InactiveSheet
    inactiveSheet.Range("F" & lastRowInactiveSheet + 2 & ":F" & lastRowInactiveSheet + 6).Copy

    ' Paste into rows 105 to 109 of column G in ActiveSheet
    activeSheet.Range("H" & lastRowActiveSheet + 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Clear the clipboard
    Application.CutCopyMode = False
            
    ' Copy lines 4 to 103 from the inactive sheet
    inactiveSheet.Rows("4:" & lastRowInactiveSheet).Copy
            
    ' Insert the copied rows above row 5 in the active sheet
    Application.DisplayAlerts = False
    activeSheet.Rows(5).Insert Shift:=x1Down
    Application.DisplayAlerts = True
            
    ' Clear the clipboard
    Application.CutCopyMode = False
            
    ' Insert empty cells on the left of column C in the pasted lines
    activeSheet.Range("C5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
    activeSheet.Range("F5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
    activeSheet.Range("M5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
    activeSheet.Range("N5").Resize(lastRowInactiveSheet - 3, 1).Insert Shift:=xlToRight
            
    activeSheet.Range("A" & lastRowActiveSheet + lastRowInactiveSheet - 3).EntireRow.Copy
    activeSheet.Range("A5:N" & lastRowInactiveSheet + 1).PasteSpecial xlPasteFormats
            
    ' Clear the clipboard
    Application.CutCopyMode = False
            
    activeSheet.Range("O4").Resize(1, 20).Copy
    activeSheet.Range("O5").Resize(lastRowInactiveSheet - 3, 20).PasteSpecial
        
    ' Clear the clipboard
    Application.CutCopyMode = False
        
    activeSheet.Range("A4").EntireRow.Delete
    activeSheet.Range("A" & lastRowInactiveSheet + 1).Resize(lastRowActiveSheet - 4, 1).EntireRow.Delete
    
    ' Autosize columns A to AH
    activeSheet.Columns("A:AH").AutoFit
    
    ' Check if the current width is greater than 50 points
    For Each col In activeSheet.Columns
        ' Check if column width is greater than 50
        If col.ColumnWidth > 50 Then
            ' If so, set the column width to 50
            col.ColumnWidth = 50
        End If
    Next col
    
End Function

