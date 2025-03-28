Attribute VB_Name = "ADODBControl"
    Option Explicit
    
    Public conn As Object
    Public rsConsulta As Object
    Public rsTodos As Object
    
Function SetUpAccess()
    
    On Error GoTo ErrorHandler
    
    ' Set up the connection to the Access database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\Database.accdb;"
    On Error GoTo 0
    
    ' Set up a recordset
    Set rsTodos = CreateObject("ADODB.Recordset")
    rsTodos.Open "Todos_Itens", conn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    ' Set up a recordset
    Set rsConsulta = CreateObject("ADODB.Recordset")
    rsConsulta.Open "Consulta_Itens", conn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    ' Exit function if successful
    Exit Function
    
ErrorHandler:
    ' Handle the error
    MsgBox "O banco de dados não pode ser acessado." & Err.Description, vbExclamation, "Error"
    
    ' Clean up objects
    If Not rsTodos Is Nothing Then
        If rsTodos.State = 1 Then rsTodos.Close ' adStateOpen = 1
        Set rsTodos = Nothing
    End If
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Function

Function SetDownAccess()
    ' Close connections
    On Error Resume Next
    rsTodos.Close
    conn.Close
    On Error GoTo 0
End Function

Sub AddItemToTodosTable(listItem As Variant, Optional ws As Worksheet = Nothing, Optional Row As Range = Nothing)
    
    'countExecution "AddItemToTodosTable", True, "Sub", "ADODBControl"
    
    Dim descritivo As String
    descritivo = listItem(1, 2)
    
    Dim codigo As String
    codigo = listItem(1, 1)
    
    Dim CheckReturn As String
    
    'SetUpAccess
    Dim rsCheck As Object
    Set rsCheck = CreateObject("ADODB.Recordset")
    
    'Look for repeated items
    Dim sql As String
    sql = "SELECT * FROM Todos_Itens WHERE " & _
      "[DESCRITIVO TÉCNICO] = '" & listItem(1, 2) & "' AND " & _
      "[CÓDIGO] = '" & listItem(1, 1) & "' AND " & _
      "[DESCRITIVO COMERCIAL] = '" & listItem(1, 3) & "' AND " & _
      "[MODELO] = '" & listItem(1, 5) & "' AND " & _
      "[FABRICANTE] = '" & listItem(1, 4) & "';"
    
    On Error GoTo ErrorHandler
    rsCheck.Open sql, conn, 1, 3
    On Error GoTo 0
    
    Dim foundCell As Range
    Dim foundCellPreviousColor As Long

    ' Check if the "DESCRITIVO TÉCNICO" already exists in the database
    If rsCheck.RecordCount > 0 Then
        
        ' Check if all fields have the same values
        CheckReturn = "Record"
        
        If CheckReturn = "Record" Or listItem(1, 7) = "" Then
            
            UpdateRecordTodos listItem
            
        End If
        
    ElseIf descritivo <> "" Then
        ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
        rsTodos.AddNew
        ' Add a new record
        'Código
        rsTodos.Fields("CÓDIGO").Value = listItem(1, 1)
        'Descritivo Técnico
        rsTodos.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
        'Descritivo Comercial
        rsTodos.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
        'Faricante
        rsTodos.Fields("FABRICANTE").Value = listItem(1, 4)
        'Modelo
        rsTodos.Fields("MODELO").Value = listItem(1, 5)
        'Un
        rsTodos.Fields("UN").Value = listItem(1, 6)
        'Preço Unitário
        If listItem(1, 7) = "" Then
        rsTodos.Fields("PREÇO UNITÁRIO").Value = 0
        Else
        rsTodos.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
        End If
        'PIS/COFINS
        If listItem(1, 8) = "" Then
        rsTodos.Fields("PIS/COFINS").Value = 0
        Else
        rsTodos.Fields("PIS/COFINS").Value = listItem(1, 8)
        End If
        'ICMS
        If listItem(1, 9) = "" Then
        rsTodos.Fields("ICMS").Value = 0
        Else
        rsTodos.Fields("ICMS").Value = listItem(1, 9)
        End If
        'IPI
        If listItem(1, 10) = "" Then
        rsTodos.Fields("IPI").Value = 0
        Else
        rsTodos.Fields("IPI").Value = listItem(1, 10)
        End If
        'Data da Cotação
        If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
        rsTodos.Fields("DATA DA COTAÇÃO").Value = Date
        Else
        rsTodos.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
        End If
        'Última Atualização
        If listItem(1, 12) = "" Then
        rsTodos.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
        Else
        rsTodos.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
        End If
        'Data de Criação
        rsTodos.Fields("DATA DE CRIAÇÃO").Value = Now
        
        rsTodos.Update
    End If

    rsCheck.Close
    
    'SetDownAccess
    
    'countExecution "AddItemToTodosTable", False, "Sub", "ADODBControl"
    
    ' Exit function if successful
    Exit Sub
    
ErrorHandler:
    
    ' Clean up objects
    If Not rsTodos Is Nothing Then
        If rsTodos.State = 1 Then rsTodos.Close ' adStateOpen = 1
        Set rsTodos = Nothing
    End If
    
    If Not rsCheck Is Nothing Then
        If rsCheck.State = 1 Then rsCheck.Close ' adStateOpen = 1
        Set rsCheck = Nothing
    End If
    
    On Error GoTo 0

End Sub

Sub AddItemToConsultaTable(listItem As Variant, Optional ws As Worksheet = Nothing, Optional Row As Range = Nothing, Optional readOnly As Boolean = False, Optional forceUpdate As Boolean = False)
    
    'countExecution "AddItemToTodosTable", True, "Sub", "ADODBControl"
    
    Dim descritivo As String
    descritivo = listItem(1, 2)
    
    Dim codigo As String
    codigo = listItem(1, 1)
    
    Dim CheckReturn As String
    
    'SetUpAccess
    Dim rsCheck As Object
    Set rsCheck = CreateObject("ADODB.Recordset")
    
    Dim sql As String
    sql = "SELECT * FROM Consulta_Itens WHERE [DESCRITIVO TÉCNICO] = '" & descritivo & "';"
    
    On Error GoTo ErrorHandler
    rsCheck.Open sql, conn, 1, 3
    On Error GoTo 0
    
    Dim foundCell As Range
    Dim foundCellPreviousColor As Long
    
    ' Check if the "DESCRITIVO TÉCNICO" already exists in the database
    If rsCheck.RecordCount > 0 Then
        
        ' Check if all fields have the same values
        CheckReturn = CheckTheNeedToUpdate(listItem, rsCheck)
            
        ' Find the cell containing "descritivo" in column B
        'Set foundCell = ws.Columns("B").Find(What:=descritivo, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
        If Not Row Is Nothing Then
            Set foundCell = Row.Cells(1, 2)
        End If
        ' If "descritivo" is found, highlight the cell in yellow
        If Not foundCell Is Nothing Then
            If foundCell.Interior.Color = RGB(255, 0, 0) Then
                foundCellPreviousColor = foundCell.Cells(1, 2).Interior.Color
            Else
                foundCellPreviousColor = foundCell.Interior.Color
            End If
            foundCell.Interior.Color = RGB(255, 0, 0) ' Red
        End If
        
        If Not readOnly And (CheckReturn = "Record" Or listItem(1, 7) = "") Then
            
            UpdateRecordConsulta listItem
            
            ' If "descritivo" is found, highlight the cell in light green
            If Not foundCell Is Nothing Then
                foundCell.Interior.Color = foundCellPreviousColor
                'foundCell.Interior.Color = RGB(0, 0, 255)
            End If
        ElseIf CheckReturn = "Sheet" Or readOnly Then
            If UpdateSheet(rsCheck, ws, foundCell) Then
                ' If "descritivo" is found, highlight the cell in light green
                If Not foundCell Is Nothing Then
                    foundCell.Interior.Color = foundCellPreviousColor
                    'foundCell.Interior.Color = RGB(0, 255, 0)
                End If
            End If
        ElseIf CheckReturn = "Ignored" Then
            foundCell.Interior.Color = foundCellPreviousColor
            Exit Sub
        Else
            If Not foundCell Is Nothing Then
                foundCell.Interior.Color = foundCellPreviousColor
            End If
            Exit Sub
        End If
    ElseIf codigo <> "" Then
        sql = "SELECT * FROM Consulta_Itens WHERE [CÓDIGO] = '" & codigo & "';"
    
        On Error GoTo ErrorHandler
        rsCheck.Close
        rsCheck.Open sql, conn, 1, 3
        On Error GoTo 0
        
        If rsCheck.RecordCount > 0 Then
            ' Check if all fields have the same values
            CheckReturn = CheckTheNeedToUpdate(listItem, rsCheck, readOnly)
                
            ' Find the cell containing "descritivo" in column B
            'Set foundCell = ws.Columns("B").Find(What:=descritivo, LookIn:=xlValues, _
                        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False)
            If Not Row Is Nothing Then
                Set foundCell = Row.Cells(1, 2)
            End If
            ' If "descritivo" is found, highlight the cell in yellow
            If Not foundCell Is Nothing Then
                If foundCell.Interior.Color = RGB(255, 0, 0) Then
                    foundCellPreviousColor = foundCell.Cells(1, 2).Interior.Color
                Else
                    foundCellPreviousColor = foundCell.Interior.Color
                End If
                foundCell.Interior.Color = RGB(255, 0, 0) ' Red
            End If
            
            If Not readOnly And CheckReturn = "Record" And descritivo <> "" And forceUpdate Then
                ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
                rsConsulta.AddNew
                ' Add a new record
                'Código
                rsConsulta.Fields("CÓDIGO").Value = listItem(1, 1)
                'Descritivo Técnico
                rsConsulta.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
                'Descritivo Comercial
                rsConsulta.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
                'Faricante
                rsConsulta.Fields("FABRICANTE").Value = listItem(1, 4)
                'Modelo
                rsConsulta.Fields("MODELO").Value = listItem(1, 5)
                'Un
                rsConsulta.Fields("UN").Value = listItem(1, 6)
                'Preço Unitário
                If listItem(1, 7) = "" Then
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = 0
                Else
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
                End If
                'PIS/COFINS
                If listItem(1, 8) = "" Then
                    rsConsulta.Fields("PIS/COFINS").Value = 0
                Else
                    rsConsulta.Fields("PIS/COFINS").Value = listItem(1, 8)
                End If
                'ICMS
                If listItem(1, 9) = "" Then
                    rsConsulta.Fields("ICMS").Value = 0
                Else
                    rsConsulta.Fields("ICMS").Value = listItem(1, 9)
                End If
                'IPI
                If listItem(1, 10) = "" Then
                    rsConsulta.Fields("IPI").Value = 0
                Else
                    rsConsulta.Fields("IPI").Value = listItem(1, 10)
                End If
                'Data da Cotação
                If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = Date
                Else
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
                End If
                'Última Atualização
                If listItem(1, 12) = "" Then
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
                Else
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
                End If
                'Data de Criação
                rsConsulta.Fields("DATA DE CRIAÇÃO").Value = Now
                
                rsConsulta.Update
            ElseIf Not readOnly And (CheckReturn = "Record" Or listItem(1, 7) = "") Then
            
                UpdateRecordConsulta listItem
            
                ' If "descritivo" is found, highlight the cell in light green
                If Not foundCell Is Nothing Then
                    foundCell.Interior.Color = foundCellPreviousColor
                    'foundCell.Interior.Color = RGB(0, 0, 255)
                End If
            ElseIf readOnly And CheckReturn = "False" Then
                ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
                rsConsulta.AddNew
                ' Add a new record
                'Código
                rsConsulta.Fields("CÓDIGO").Value = listItem(1, 1)
                'Descritivo Técnico
                rsConsulta.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
                'Descritivo Comercial
                rsConsulta.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
                'Faricante
                rsConsulta.Fields("FABRICANTE").Value = listItem(1, 4)
                'Modelo
                rsConsulta.Fields("MODELO").Value = listItem(1, 5)
                'Un
                rsConsulta.Fields("UN").Value = listItem(1, 6)
                'Preço Unitário
                If listItem(1, 7) = "" Then
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = 0
                Else
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
                End If
                'PIS/COFINS
                If listItem(1, 8) = "" Then
                    rsConsulta.Fields("PIS/COFINS").Value = 0
                Else
                    rsConsulta.Fields("PIS/COFINS").Value = listItem(1, 8)
                End If
                'ICMS
                If listItem(1, 9) = "" Then
                    rsConsulta.Fields("ICMS").Value = 0
                Else
                    rsConsulta.Fields("ICMS").Value = listItem(1, 9)
                End If
                'IPI
                If listItem(1, 10) = "" Then
                    rsConsulta.Fields("IPI").Value = 0
                Else
                    rsConsulta.Fields("IPI").Value = listItem(1, 10)
                End If
                'Data da Cotação
                If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = Date
                Else
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
                End If
                'Última Atualização
                If listItem(1, 12) = "" Then
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
                Else
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
                End If
                'Data de Criação
                rsConsulta.Fields("DATA DE CRIAÇÃO").Value = Now
                
                rsConsulta.Update
            ElseIf CheckReturn = "Sheet" Or readOnly Then
                If UpdateSheet(rsCheck, ws, foundCell) Then
                    ' If "descritivo" is found, highlight the cell in light green
                    If Not foundCell Is Nothing Then
                        foundCell.Interior.Color = foundCellPreviousColor
                        'foundCell.Interior.Color = RGB(0, 255, 0)
                    End If
                End If
            ElseIf CheckReturn = "Ignored" Then
                foundCell.Interior.Color = foundCellPreviousColor
                Exit Sub
            Else
                If Not foundCell Is Nothing Then
                    foundCell.Interior.Color = foundCellPreviousColor
                End If
                Exit Sub
            End If
        ElseIf descritivo <> "" And Not readOnly Then
            ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
            rsConsulta.AddNew
            ' Add a new record
            'Código
            rsConsulta.Fields("CÓDIGO").Value = listItem(1, 1)
            'Descritivo Técnico
            rsConsulta.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
            'Descritivo Comercial
            rsConsulta.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
            'Faricante
            rsConsulta.Fields("FABRICANTE").Value = listItem(1, 4)
            'Modelo
            rsConsulta.Fields("MODELO").Value = listItem(1, 5)
            'Un
            rsConsulta.Fields("UN").Value = listItem(1, 6)
            'Preço Unitário
            If listItem(1, 7) = "" Then
                rsConsulta.Fields("PREÇO UNITÁRIO").Value = 0
            Else
                rsConsulta.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
            End If
            'PIS/COFINS
            If listItem(1, 8) = "" Then
                rsConsulta.Fields("PIS/COFINS").Value = 0
            Else
                rsConsulta.Fields("PIS/COFINS").Value = listItem(1, 8)
            End If
            'ICMS
            If listItem(1, 9) = "" Then
                rsConsulta.Fields("ICMS").Value = 0
            Else
                rsConsulta.Fields("ICMS").Value = listItem(1, 9)
            End If
            'IPI
            If listItem(1, 10) = "" Then
                rsConsulta.Fields("IPI").Value = 0
            Else
                rsConsulta.Fields("IPI").Value = listItem(1, 10)
            End If
            'Data da Cotação
            If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
                rsConsulta.Fields("DATA DA COTAÇÃO").Value = Date
            Else
                rsConsulta.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
            End If
            'Última Atualização
            If listItem(1, 12) = "" Then
                rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
            Else
                rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
            End If
            'Data de Criação
            rsConsulta.Fields("DATA DE CRIAÇÃO").Value = Now
            
            rsConsulta.Update
        End If
    ElseIf descritivo <> "" And Not readOnly Then
        
        sql = "SELECT * FROM Consulta_Itens WHERE [DESCRITIVO TÉCNICO] LIKE '%" & descritivo & "%';"
        rsCheck.Close
        rsCheck.Open sql, conn, 1, 3
        
        ' Check if the "descritivo" is too vague and avoid saving it to the database, but highlight it
        If rsCheck.RecordCount >= 1 Then
            
            If Not Row Is Nothing Then
                Set foundCell = Row.Cells(1, 2)
            End If
            
            If Not foundCell Is Nothing Then
                If foundCell.Interior.Color = RGB(255, 255, 0) Then
                    foundCellPreviousColor = foundCell.Cells(1, 2).Interior.Color
                Else
                    foundCellPreviousColor = foundCell.Interior.Color
                End If
            
                foundCell.Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
            
            ConsultaBancoDeDados.ConvertRowToListItem Row
            ConsultaBancoDeDados.DisplaySelectedItem
            
            Dim response As VbMsgBoxResult
            response = MsgBox("Foram encontrados outros itens parecidos com esse descritivo:" & _
                        vbCrLf & vbCrLf & descritivo & vbCrLf & vbCrLf & _
                        "Você realmente deseja salvar esse item no banco de dados?", _
                        vbYesNo + vbExclamation + vbDefaultButton1, "Confirmar Registro")
        
            ' If the user clicks Yes, save the workbook
            If response = vbYes And Not readOnly Then
                ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
                rsConsulta.AddNew
                ' Add a new record
                'Código
                rsConsulta.Fields("CÓDIGO").Value = listItem(1, 1)
                'Descritivo Técnico
                rsConsulta.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
                'Descritivo Comercial
                rsConsulta.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
                'Faricante
                rsConsulta.Fields("FABRICANTE").Value = listItem(1, 4)
                'Modelo
                rsConsulta.Fields("MODELO").Value = listItem(1, 5)
                'Un
                rsConsulta.Fields("UN").Value = listItem(1, 6)
                'Preço Unitário
                If listItem(1, 7) = "" Then
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = 0
                Else
                    rsConsulta.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
                End If
                'PIS/COFINS
                If listItem(1, 8) = "" Then
                    rsConsulta.Fields("PIS/COFINS").Value = 0
                Else
                    rsConsulta.Fields("PIS/COFINS").Value = listItem(1, 8)
                End If
                'ICMS
                If listItem(1, 9) = "" Then
                    rsConsulta.Fields("ICMS").Value = 0
                Else
                    rsConsulta.Fields("ICMS").Value = listItem(1, 9)
                End If
                'IPI
                If listItem(1, 10) = "" Then
                    rsConsulta.Fields("IPI").Value = 0
                Else
                    rsConsulta.Fields("IPI").Value = listItem(1, 10)
                End If
                'Data da Cotação
                If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = Date
                Else
                    rsConsulta.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
                End If
                'Última Atualização
                If listItem(1, 12) = "" Then
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
                Else
                    rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
                End If
                'Data de Criação
                rsConsulta.Fields("DATA DE CRIAÇÃO").Value = Now
            
                rsConsulta.Update
    
                foundCell.Interior.Color = foundCellPreviousColor
            End If
        ElseIf Not readOnly Then
            ' Add a new record if DESCRITIVO TÉCNICO doesn't exist
            rsConsulta.AddNew
            ' Add a new record
            'Código
            rsConsulta.Fields("CÓDIGO").Value = listItem(1, 1)
            'Descritivo Técnico
            rsConsulta.Fields("DESCRITIVO TÉCNICO").Value = listItem(1, 2)
            'Descritivo Comercial
            rsConsulta.Fields("DESCRITIVO COMERCIAL").Value = listItem(1, 3)
            'Faricante
            rsConsulta.Fields("FABRICANTE").Value = listItem(1, 4)
            'Modelo
            rsConsulta.Fields("MODELO").Value = listItem(1, 5)
            'Un
            rsConsulta.Fields("UN").Value = listItem(1, 6)
            'Preço Unitário
            If listItem(1, 7) = "" Then
                rsConsulta.Fields("PREÇO UNITÁRIO").Value = 0
            Else
                rsConsulta.Fields("PREÇO UNITÁRIO").Value = listItem(1, 7)
            End If
            'PIS/COFINS
            If listItem(1, 8) = "" Then
                rsConsulta.Fields("PIS/COFINS").Value = 0
            Else
                rsConsulta.Fields("PIS/COFINS").Value = listItem(1, 8)
            End If
            'ICMS
            If listItem(1, 9) = "" Then
                rsConsulta.Fields("ICMS").Value = 0
            Else
                rsConsulta.Fields("ICMS").Value = listItem(1, 9)
            End If
            'IPI
            If listItem(1, 10) = "" Then
                rsConsulta.Fields("IPI").Value = 0
            Else
                rsConsulta.Fields("IPI").Value = listItem(1, 10)
            End If
            'Data da Cotação
            If listItem(1, 11) = "" Or listItem(1, 11) = 0 Then
                rsConsulta.Fields("DATA DA COTAÇÃO").Value = Date
            Else
                rsConsulta.Fields("DATA DA COTAÇÃO").Value = listItem(1, 11)
            End If
            'Última Atualização
            If listItem(1, 12) = "" Then
                rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = Now
            Else
                rsConsulta.Fields("ÚLTIMA ATUALIZAÇÃO").Value = listItem(1, 12)
            End If
            'Data de Criação
            rsConsulta.Fields("DATA DE CRIAÇÃO").Value = Now
            
            rsConsulta.Update
        End If
    Else
        MsgBox descritivo & vbCrLf & "Esse descritivo não pode ser adicionado."
    End If

    rsCheck.Close
    
    'SetDownAccess
    
    'countExecution "AddItemToTodosTable", False, "Sub", "ADODBControl"
    
    ' Exit function if successful
    Exit Sub
    
ErrorHandler:
    
    ' Clean up objects
    If Not rsConsulta Is Nothing Then
        If rsConsulta.State = 1 Then rsConsulta.Close ' adStateOpen = 1
        Set rsConsulta = Nothing
    End If
    
    If Not rsCheck Is Nothing Then
        If rsCheck.State = 1 Then rsCheck.Close ' adStateOpen = 1
        Set rsCheck = Nothing
    End If
    
    On Error GoTo 0

End Sub

Sub UpdateRecordTodos(listItem As Variant)
    
    'countExecution "UpdateRecordConsulta", True, "Sub", "ADODBControl"
    
    ' Delete existing item
    DeleteRowFromTodosTable listItem
    
    ' Add the modified item
    AddItemToTodosTable listItem
    
    'countExecution "UpdateRecordConsulta", False, "Sub", "ADODBControl"
    
End Sub

Sub UpdateRecordConsulta(listItem As Variant, Optional forceUpdate As Boolean = False)
    
    'countExecution "UpdateRecordConsulta", True, "Sub", "ADODBControl"
    
    ' Delete existing item
    DeleteRowFromConsultaTable listItem
    
    ' Add the modified item
    AddItemToConsultaTable listItem, , , , forceUpdate
    
    'countExecution "UpdateRecordConsulta", False, "Sub", "ADODBControl"
    
End Sub

Sub DeleteRowFromConsultaTable(valuesArray As Variant)
    
    'countExecution "DeleteRowFromConsultaTable", True, "Sub", "ADODBControl"
    
    Dim i As Long
    
    ' Build the WHERE clause for the DELETE query
    Dim whereClause As String
    whereClause = ""
    If valuesArray(1, 2) <> "" Then
        whereClause = "TRIM([DESCRITIVO TÉCNICO]) ALIKE '" & Replace(Replace(Replace(valuesArray(1, 2), vbCrLf, "%"), vbLf, "%"), vbCr, "%") & "'"
    End If
        
    whereClause = whereClause & ");"

    ' Execute the DELETE query
    Dim sql As String
    sql = "DELETE FROM (SELECT TOP 1 * FROM Consulta_Itens WHERE " & whereClause
    On Error GoTo ErrorHandler
    conn.Execute sql
    
    If False = True Then
ErrorHandler:
        Debug.Print "Nenhum item foi encontrado para ser deletado."
    End If
    
    On Error GoTo 0
    
    'countExecution "DeleteRowFromConsultaTable", False, "Sub", "ADODBControl"
    
End Sub

Sub DeleteRowFromTodosTable(valuesArray As Variant)
    
    'countExecution "DeleteRowFromTodosTable", True, "Sub", "ADODBControl"
    
    Dim i As Long
    
    ' Build the WHERE clause for the DELETE query
    Dim whereClause As String
    whereClause = ""
    If valuesArray(1, 2) <> "" Then
        whereClause = "TRIM([DESCRITIVO TÉCNICO]) ALIKE '" & Replace(Replace(Replace(valuesArray(1, 2), vbCrLf, "%"), vbLf, "%"), vbCr, "%") & "'"
    End If
        
    whereClause = whereClause & " AND " & _
      "[CÓDIGO] = '" & valuesArray(1, 1) & "' AND " & _
      "[DESCRITIVO COMERCIAL] = '" & valuesArray(1, 3) & "' AND " & _
      "[FABRICANTE] = '" & valuesArray(1, 4) & "' AND " & _
      "[MODELO] = '" & valuesArray(1, 5) & "'"
        
    whereClause = whereClause & ");"

    ' Execute the DELETE query
    Dim sql As String
    sql = "DELETE FROM (SELECT TOP 1 * FROM Todos_Itens WHERE " & whereClause
    On Error GoTo ErrorHandler
    conn.Execute sql
    
    If False = True Then
ErrorHandler:
        'MsgBox "Nenhum item foi encontrado para ser deletado."
    End If
    
    On Error GoTo 0
    
    'countExecution "DeleteRowFromTodosTable", False, "Sub", "ADODBControl"
    
End Sub

Function AreValuesEqual(value1 As Variant, value2 As Variant) As Boolean
    
    countExecution "AreValuesEqual", True, "Function", "ADODBControl"
    
    ' Check if two values are equal, handling Null and Empty
    If IsNull(value1) And IsEmpty(value2) Then
        AreValuesEqual = True
    ElseIf IsNull(value2) And IsEmpty(value1) Then
        AreValuesEqual = True
    ElseIf Trim(value1) = Trim(value2) Then
        AreValuesEqual = True
    Else
        AreValuesEqual = False
    End If
End Function

Function SearchByDescritivo(searchText As String, Optional searchBrand As String = "", Optional searchAllDB As Boolean = False) As Variant
    
    'countExecution "SearchByDescritivo", True, "Function", "ADODBControl"
    
    Dim resultArray() As Variant
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long
    Dim rsCheck As Object
    Dim word As Variant
    Dim words As Variant

    'Set error handler for database connection
    On Error GoTo ErrorHandler
    If False = True Then
ErrorHandler:
        On Error GoTo 0
        SetUpAccess
    End If
    
    ' Check if the recordset is not empty
    If Not rsTodos.EOF Then
        'Reset the error handler after confirming database connection
        On Error GoTo 0
        
        ' Initialize the result array as an empty array
        ' Initialize the result array with an initial size (adjust as needed)
        ReDim resultArray(1 To rsTodos.Fields.Count, 1 To 1)
            
        ' Split the search text into words
        words = Split(searchText)
            
        ' Check if a record with the specified value in the specified field exists
        Set rsCheck = CreateObject("ADODB.Recordset")
    
        Dim sql As String
        sql = "SELECT * FROM Consulta_Itens WHERE "
    
        For Each word In words
            sql = sql & "[DESCRITIVO TÉCNICO] LIKE '%" & word & "%' AND "
        Next word
    
        sql = sql & "[FABRICANTE] LIKE '%" & searchBrand & "%';"

        'SELECT * FROM Todos_Itens WHERE [DESCRITIVO TÉCNICO] LIKE '*DISJ*' AND [DESCRITIVO TÉCNICO] LIKE '*TRIPOLAR*' AND [FABRICANTE] LIKE '*ABB*';
        rsCheck.Open sql, conn, 1, 1

        If Not rsCheck.EOF Then
            ' Get number of rows and columns in the recordset
            rowCount = rsCheck.RecordCount
            colCount = rsCheck.Fields.Count
            
            ' Resize array
            ReDim resultArray(1 To rowCount, 1 To colCount)
            
            ' Load the data into an array by columns
            dataArray = rsCheck.GetRows(rowCount)
        
            ' Resize the resultArray to fit the data
            ReDim resultArray(1 To rowCount, 1 To colCount)
        
            ' Transfer data from the array to the resultArray (transposed)
            For i = 1 To rowCount
                For j = 1 To colCount
                    resultArray(i, j) = dataArray(j - 1, i - 1)
                Next j
            Next i

        SearchByDescritivo = Not Empty
        Else
            SearchByDescritivo = Empty
        End If
        
        ' Close recordset
        rsCheck.Close
        
        searchAllDB = ConsultaBancoDeDados.AmbosDB.Value
        
        If searchAllDB Then
            
            If IsEmpty(SearchByDescritivo) Then
                ' Get number of rows and columns in the recordset
                rowCount = 1
                colCount = 1
                
                ' Resize array
                ReDim resultArray(1 To rowCount, 1 To colCount)
            
                ' Resize the resultArray to fit the data
                ReDim resultArray(1 To rowCount, 1 To colCount)
            End If
        
            sql = "SELECT * FROM Todos_Itens WHERE "
        
            For Each word In words
                sql = sql & "[DESCRITIVO TÉCNICO] LIKE '%" & word & "%' AND "
            Next word
        
            sql = sql & "[FABRICANTE] LIKE '%" & searchBrand & "%';"
    
            'SELECT * FROM Todos_Itens WHERE [DESCRITIVO TÉCNICO] LIKE '*DISJ*' AND [DESCRITIVO TÉCNICO] LIKE '*TRIPOLAR*' AND [FABRICANTE] LIKE '*ABB*';
            rsCheck.Open sql, conn, 1, 1
            
            If Not rsCheck.EOF Then
                ' Get number of rows and columns in the recordset
                rowCount = rsCheck.RecordCount
                colCount = rsCheck.Fields.Count

                ' Load the data into an array by columns
                dataArray = rsCheck.GetRows(rowCount)
                
                ' Modify the first row of dataArray to add a "#" before each value
                For i = 0 To rowCount - 1
                    dataArray(1, i) = "##" & dataArray(1, i)
                Next i
                
                ReDim Preserve dataArray(0 To colCount - 1, 0 To rowCount + UBound(resultArray, 1) - 1)
                
                 ' Transfer data from the array to the resultArray (transposed)
                For i = rowCount To UBound(dataArray, 2) - 1
                    For j = 1 To colCount
                        dataArray(j - 1, i) = resultArray(i - rowCount + 1, j)
                    Next j
                Next i
            
                ' Resize the resultArray to fit the data
                ReDim resultArray(1 To UBound(dataArray, 2) + 1, 1 To UBound(dataArray, 1) + 1)
            
               ' Transfer data from the array to the resultArray (transposed)
                For i = 1 To UBound(resultArray, 1)
                    For j = 1 To UBound(resultArray, 2)
                        resultArray(i, j) = dataArray(j - 1, i - 1)
                    Next j
                Next i
                
                ' Close recordset
                rsCheck.Close
                
                SearchByDescritivo = Not Empty
            End If
        End If
        
        If Not IsEmpty(SearchByDescritivo) Then
            SearchByDescritivo = resultArray
        End If
    End If
    
    'countExecution "SearchByDescritivo", True, "Function", "ADODBControl"
    
End Function
Function CheckTheNeedToUpdate(listItem As Variant, rsCheck As Object, Optional readOnly As Boolean = False) As String
    
    'countExecution "CheckTheNeedToUpdate", True, "Function", "ADODBControl"
    
    Dim msg As String
    Dim optA As String
    Dim optB As String
    
    optA = ""
    optB = ""
    msg = listItem(1, 2) & vbCrLf & "O item não está conforme o banco de dados" & vbCrLf
    
    Dim AutoUpdateRecord As Boolean
    Dim AutoUpdateSheet As Boolean
    
    CheckTheNeedToUpdate = False
    AutoUpdateRecord = True
    AutoUpdateSheet = True
    
    Dim dataArray() As Variant
    
    dataArray = rsCheck.GetRows(rsCheck.RecordCount)
    
    If UCase(Left(listItem(1, 2), 6)) = "MISCEL" Then
        CheckTheNeedToUpdate = "Ignored"
        AutoUpdateSheet = False
        AutoUpdateRecord = False
        Exit Function
    End If
    
    If dataArray(0, 0) = listItem(1, 1) And Not IsNull(dataArray(0, 0)) And readOnly Then
        ' Avoid an iten been changed if the CÓDIGO already exists
        AutoUpdateSheet = True
        AutoUpdateRecord = False
        CheckTheNeedToUpdate = True
    ElseIf dataArray(1, 0) = listItem(1, 2) Then
        ' Check if all fields have the same values as the existing record
        
        If dataArray(0, 0) <> listItem(1, 1) Then
            optA = optA & "CÓDIGO: " & dataArray(0, 0) & vbCrLf
            optB = optB & "CÓDIGO: " & listItem(1, 1) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(0, 0) = "" And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 1) = "" And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If dataArray(2, 0) <> listItem(1, 3) Or IsNull(dataArray(2, 0)) Then
            optA = optA & "DESCRITIVO COMERCIAL: " & dataArray(2, 0) & vbCrLf
            optB = optB & "DESCRITIVO COMERCIAL: " & listItem(1, 3) & vbCrLf
            CheckTheNeedToUpdate = True
            If (dataArray(2, 0) = "" Or IsNull(dataArray(2, 0))) And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 3) = "" And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If dataArray(3, 0) <> listItem(1, 4) Or IsNull(dataArray(3, 0)) Then
            optA = optA & "FABRICANTE: " & dataArray(3, 0) & vbCrLf
            optB = optB & "FABRICANTE: " & listItem(1, 4) & vbCrLf
            CheckTheNeedToUpdate = True
            If (dataArray(3, 0) = "" Or IsNull(dataArray(3, 0))) And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 4) = "" And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If dataArray(4, 0) <> listItem(1, 5) Then
            optA = optA & "MODELO: " & dataArray(4, 0) & vbCrLf
            optB = optB & "MODELO: " & listItem(1, 5) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(4, 0) = "" And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 5) = "" And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If dataArray(5, 0) <> listItem(1, 6) Then
            optA = optA & "UN: " & dataArray(5, 0) & vbCrLf
            optB = optB & "UN: " & listItem(1, 6) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(5, 0) = "" And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 6) = "" And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If Trim(dataArray(7, 0)) <> Trim(listItem(1, 8)) And Trim(listItem(1, 8)) <> "" Then
            optA = optA & "PIS/COFINS: " & dataArray(7, 0) & vbCrLf
            optB = optB & "PIS/COFINS: " & listItem(1, 8) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(7, 0) = 0 And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 8) = 0 And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If Trim(dataArray(8, 0)) <> Trim(listItem(1, 9)) And Trim(listItem(1, 9)) <> "" Then
            optA = optA & "ICMS: " & Trim(dataArray(8, 0)) & vbCrLf
            optB = optB & "ICMS: " & listItem(1, 9) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(8, 0) = 0 And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 9) = 0 And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        If Trim(dataArray(9, 0)) <> Trim(listItem(1, 10)) And Trim(listItem(1, 10)) <> "" Then
            optA = optA & "IPI: " & Trim(dataArray(9, 0)) & vbCrLf
            optB = optB & "IPI: " & listItem(1, 10) & vbCrLf
            CheckTheNeedToUpdate = True
            If dataArray(9, 0) = 0 And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf listItem(1, 10) = 0 And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
                
        If Trim("R$ " & dataArray(6, 0)) <> Trim(listItem(1, 7)) And Trim(dataArray(6, 0)) <> Trim(listItem(1, 7)) Then
            optA = optA & "PREÇO UNITÁRIO: " & Trim("R$ " & dataArray(6, 0)) & vbCrLf
            optB = optB & "PREÇO UNITÁRIO: " & "R$ " & listItem(1, 7) & vbCrLf
            CheckTheNeedToUpdate = True
            If listItem(1, 7) <> 0 And Trim(dataArray(10, 0)) < Trim(listItem(1, 11)) And AutoUpdateRecord Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf Trim(dataArray(10, 0)) <> Trim(listItem(1, 11)) And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
        
        If Trim(dataArray(10, 0)) <> Trim(listItem(1, 11)) Or IsNull(Trim(dataArray(10, 0))) Then
            If Trim(dataArray(10, 0)) <> Trim(listItem(1, 11)) Then
                optA = optA & "DATA DA COTAÇÃO: " & Trim(dataArray(10, 0)) & vbCrLf
                optB = optB & "DATA DA COTAÇÃO: " & listItem(1, 11) & vbCrLf
            End If
            CheckTheNeedToUpdate = True
            If (dataArray(10, 0) = "" Or Trim(dataArray(10, 0)) <= Trim(listItem(1, 11)) Or IsNull(Trim(dataArray(10, 0)))) And Not AutoUpdateSheet Then
                AutoUpdateRecord = True
                AutoUpdateSheet = False
            ElseIf (IsNull(Trim(dataArray(10, 0))) And Trim(listItem(1, 11)) = "") And AutoUpdateSheet Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            ElseIf Not Trim(dataArray(10, 0)) <= Trim(listItem(1, 11)) And Not AutoUpdateRecord Then
                AutoUpdateSheet = True
                AutoUpdateRecord = False
            Else
                AutoUpdateSheet = False
                AutoUpdateRecord = False
            End If
        End If
    Else
        AutoUpdateSheet = False
        AutoUpdateRecord = True
        CheckTheNeedToUpdate = True
    End If
    
    msg = msg & vbCrLf & "Banco de dados: " & vbCrLf & optA & vbCrLf & "Planilha: " & vbCrLf & optB & vbCrLf & _
            "Deseja atualizar a planilha conforme o banco de dados?" & vbCrLf & vbCrLf & "**Clicar em Não atualizará o banco de dados.'"
            
    Dim response As VbMsgBoxResult
            
    If CheckTheNeedToUpdate = True And AutoUpdateRecord And Not readOnly Then
        CheckTheNeedToUpdate = "Record"
        Exit Function
    ElseIf CheckTheNeedToUpdate = True And AutoUpdateSheet Then
        CheckTheNeedToUpdate = "Sheet"
        Exit Function
    ElseIf CheckTheNeedToUpdate = True Then
        response = MsgBox(msg, vbYesNoCancel + vbQuestion + vbMsgBoxSetForeground, "Item desatualizado")
        If response = vbYes Then
            CheckTheNeedToUpdate = "Sheet"
        ElseIf response = vbNo Then
            CheckTheNeedToUpdate = "Record"
        ElseIf response = vbCancel Then
            CheckTheNeedToUpdate = "Ignored"
        Else
            CheckTheNeedToUpdate = "False"
        End If
    Else
        CheckTheNeedToUpdate = "False"
    End If
End Function

Function UpdateSheet(rsCheck As Object, ws As Worksheet, foundCell As Range) As Boolean
    
    rsCheck.MoveFirst

    Dim dataArray() As Variant
    dataArray = rsCheck.GetRows(rsCheck.RecordCount)
    
    Dim Row As Range
    Set Row = foundCell.EntireRow
    
    If Not Row Is Nothing Then
        Row.EntireRow.Cells(1, codigoColumn) = dataArray(0, 0) 'Código
        Row.EntireRow.Cells(1, componenteColumn) = dataArray(1, 0) 'Descritivo técnico
        Row.EntireRow.Cells(1, descritivoColumn) = dataArray(2, 0) 'Descritivo comercial
        Row.EntireRow.Cells(1, fabricanteColumn) = dataArray(3, 0) 'Fabricante
        Row.EntireRow.Cells(1, modeloColumn) = dataArray(4, 0) 'Modelo
        Row.EntireRow.Cells(1, unColumn) = dataArray(5, 0) 'Unidade
        Row.EntireRow.Cells(1, preçoColumn) = dataArray(6, 0) 'Preço
        Row.EntireRow.Cells(1, pisConfinsColumn) = dataArray(7, 0) 'PIS/COFINS
        Row.EntireRow.Cells(1, icmsColumn) = dataArray(8, 0) 'ICMS
        Row.EntireRow.Cells(1, ipiColumn) = dataArray(9, 0) 'IPI
        Row.EntireRow.Cells(1, dataColumn) = dataArray(10, 0) 'Data da cotação
        
        UpdateSheet = True
    Else
        MsgBox "Não foi possível atualizar" & vbCrLf & dataArray(1, 0), vbExclamation
        UpdateSheet = False
    End If
    
End Function

