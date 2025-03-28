Attribute VB_Name = "GlobalAndPublic"
Option Explicit

'Manipulated item
Public listBoxSelected As Variant
Public listBoxSelectedItem As Variant

'Columns
Public itemColumn As Integer
Public componenteColumn As Integer
Public descritivoColumn As Integer
Public fabricanteColumn As Integer
Public modeloColumn As Integer
Public codigoColumn As Integer
Public unColumn As Integer
Public qtdeColumn As Integer
Public preçoColumn As Integer
Public ipiColumn As Integer
Public pisConfinsColumn As Integer
Public icmsColumn As Integer
Public dataColumn As Integer

Public Sub OptimizedMode(ByVal enable As Boolean)
     Application.EnableEvents = Not enable
     Application.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
     Application.ScreenUpdating = Not enable
     Application.EnableAnimations = Not enable
     'Application.DisplayStatusBar = Not enable
     Application.PrintCommunication = Not enable
End Sub

Public Function GetColumnsNumbers()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("1")
    
    ' Search for each word and get the corresponding column number
    itemColumn = FindColumnByHeader(ws, "ITEM")
    componenteColumn = FindColumnByHeader(ws, "COMPONENTE")
    descritivoColumn = FindColumnByHeader(ws, "DESCRITIVO")
    fabricanteColumn = FindColumnByHeader(ws, "FABRICANTE")
    modeloColumn = FindColumnByHeader(ws, "MODELO")
    codigoColumn = FindColumnByHeader(ws, "CÓDIGO")
    unColumn = FindColumnByHeader(ws, "UN")
    qtdeColumn = FindColumnByHeader(ws, "QTDE")
    preçoColumn = FindColumnByHeader(ws, "SEM IPI")
    pisConfinsColumn = FindColumnByHeader(ws, "CONFINS")
    icmsColumn = FindColumnByHeader(ws, "COMPRA")
    ipiColumn = FindColumnByHeader(ws, "IPI")
    dataColumn = FindColumnByHeader(ws, "COTAÇÃO")

    CompareSheets

End Function

Function FindColumnByHeader(ws As Worksheet, header As String) As Integer
    Dim foundRange As Range
    Dim headersRow As Integer
    headersRow = 3
    
    On Error Resume Next
    Set foundRange = ws.Rows(headersRow).Find(header, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not foundRange Is Nothing Then
        FindColumnByHeader = foundRange.Column
    Else
        ' Return -1 or handle the case when the header is not found
        FindColumnByHeader = -1
    End If
End Function

Function CompareSheets()
    Dim wsReference As Worksheet
    Dim wsCompare As Worksheet
    Dim sheetNumber As Integer
    
    ' Set the reference sheet
    Set wsReference = ThisWorkbook.Sheets("1")
    
    ' Loop through sheets with names "2" to "30"
    For sheetNumber = 2 To 30
        On Error Resume Next
        ' Attempt to set the compare sheet
        Set wsCompare = ThisWorkbook.Sheets(CStr(sheetNumber))
        On Error GoTo 0
        
        ' Check if the compare sheet exists
        If Not wsCompare Is Nothing Then
            ' Compare the contents of the sheets
            If Not CompareRanges(wsReference, wsCompare) Then
                'MsgBox "All sheets are equal to Sheet 1", vbInformation
                Exit Function
            End If
        Else
            MsgBox "Sheet " & sheetNumber & " does not exist", vbExclamation
        End If
    Next sheetNumber
    
End Function

Function CompareRanges(ws1 As Worksheet, ws2 As Worksheet) As Boolean
    Dim rng1 As Range
    Dim rng2 As Range
    
    ' Assuming data starts from cell A1 and ends at the last used cell in both sheets
    Set rng1 = ws1.Range("A1").CurrentRegion
    Set rng2 = ws2.Range("A1").CurrentRegion
    
    If rng1.Columns.Count = rng2.Columns.Count Then
        Dim cell1 As Range
        Dim cell2 As Range

        ' Loop through each cell in the ranges and compare values
        For Each cell1 In rng1
            For Each cell2 In rng2
                ' If any pair of cells do not match, return False
                If (cell1.Row = 2 Or cell1.Row = 3) And (cell2.Row = 2 Or cell2.Row = 3) And cell1.Value <> cell2.Value And cell1.Address = cell2.Address Then
                    CompareRanges = False
                    MsgBox "Os nomes das colunas foram definidos. Existem diferenças na planilha " & ws2.Name & " célula " & cell2.Address & vbCrLf & cell2.Value, vbExclamation
                    Exit Function
                End If
                
                If cell1.Row = 1 Or cell2.Row > 3 Then
                    Exit For
                End If
            Next cell2
            
            If cell1.Row > 3 Then
                Exit For
            End If
        Next cell1

        ' If all cells match, return True
        CompareRanges = True
    Else
        ' If ranges are not of the same size or have different addresses, return False
        CompareRanges = False
    End If
End Function

Public Function countExecution(Name As String, executionType As Boolean, Optional FType As String, Optional Location As String)
    Dim connFE As Object
    Dim rsFE As Object
    
    ' Set up the connection to the Access database
    Set connFE = CreateObject("ADODB.Connection")
    connFE.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\Database.accdb;"

    ' Set up a recordset
    Set rsFE = CreateObject("ADODB.Recordset")
    rsFE.Open "Function_Execution", connFE, adOpenKeyset, adLockOptimistic, adCmdTable
    
    rsFE.AddNew
    rsFE.Fields("Name").Value = Name
    rsFE.Fields("Type").Value = FType
    rsFE.Fields("Location").Value = Location
    
    If executionType Then
        rsFE.Fields("ExecutionStart").Value = Now
    Else
        rsFE.Fields("ExecutionEnd").Value = Now
    End If
    
    rsFE.Fields("Version").Value = CheckVersion
    rsFE.Fields("User").Value = Application.UserName
    
    rsFE.Update
    
    ' Close connections
    On Error Resume Next
    rsFE.Close
    connFE.Close
    On Error GoTo 0
End Function
