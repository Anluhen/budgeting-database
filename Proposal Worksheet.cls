VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    ' Tentativa de forçar a verificação do banco de dados
    
    OptimizedMode True

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim selectedRows As Range
    Dim Row As Range
    
    If ThisWorkbook.Sheets("S.PROP").Range("A1").Value <> "" Then
        GoTo ExitSub
    End If
    
    ConsultaBancoDeDados.Show
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Range("A1").Value = "NOME DO PAINEL>>>" And ws.Range("C1").Value <> "" Then
            ws.Activate
        
            ' Find the last row with a number in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
            Set selectedRows = ws.Range("B4:B" & lastRow).EntireRow
        
            For Each Row In selectedRows
                If Not IsEmpty(Row.Cells(1, 2)) Then
                
                    ConsultaBancoDeDados.ConvertRowToListItem Row
            
                    'Add selected item to the DB
                    AddItemToTodosTable listBoxSelectedItem, ws, Row
            
                End If
            Next
        End If
    Next ws
    
    ConsultaBancoDeDados.Hide

ExitSub:
    ThisWorkbook.Sheets("S.PROP").Activate
    
    OptimizedMode False
        
End Sub
