VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    Option Explicit

Private Sub Worksheet_Activate()
       
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim Item As Long
    Dim newRow As Range
    Dim i As Long, j As Long, lastRow As Long
    Dim sheetTime As Double
    Dim materiaisSheet As Worksheet
    Dim matchCount As Integer
    Dim materiaisSheetLastRow As Long
    Dim col As Range
    Dim index As Integer
    
countExecution "Worksheet_Activate", True, "Event", "Planilha4(Materiais)"
    
    OptimizedMode True
    
' This bit is here for correction porpouses of old sheets
Application.Run ("VersionAndUpdate.CorrectErrors")

    OptimizedMode True

Dim tempIni As Long
tempIni = Timer
    
    Item = 1
    
    ' Set the first row for items list sheets
    firstRow = 4
    
    ' Set reference to the "Materiais" sheet

    Set materiaisSheet = ThisWorkbook.Sheets("Materiais")
    
    materiaisSheet.Range("A2:Z" & materiaisSheet.Rows.Count).ClearContents
        
    materiaisSheetLastRow = materiaisSheet.Cells(materiaisSheet.Rows.Count, "A").End(xlUp).Count + 1
    
sheetTime = Timer

        ' Loop through each worksheet
    For index = 1 To 30
        Set ws = ThisWorkbook.Sheets("" & index)
        
        ' Check if cell A1 is equal to "NOME DO PAINEL>>>"
        If ws.Range("A1").Value = "NOME DO PAINEL>>>" And Not IsEmpty(ws.Range("C1")) And ws.Range("Q1") > 0 Then
            Debug.Print "Sheet " & ws.Name & " done in " & Timer - sheetTime
        
            sheetTime = Timer
            
            ' Find the last used row in the sheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Check if the last row is greater than 4 (to avoid copying headers)
            If lastRow < 4 Then
                GoTo ContinueToNextSheet
            End If
            For i = firstRow To lastRow
                If (ws.Range("B" & i).Value <> "" Or ws.Range("C" & i).Value <> "") And ws.Range("H" & i).Value > 0 Then
                    Set newRow = materiaisSheet.Cells(materiaisSheetLastRow, 1).EntireRow
                    
                     newRow.Cells(1, 2).Value = ws.Name
                     
                    ' Copy values from the found rows to the new rows in "Materiais" table
                    newRow.Cells(1, 3).Resize(1, 17).Value = ws.Range("B" & i & ":R" & i).Value
                    
                    newRow.Cells(1, 20).Value = ws.Range("Q1").Value
                     
                    newRow.Cells(1, 21).Formula = "=T" & newRow.Row & "*S" & newRow.Row
                    
                    With newRow.Cells(1, 1).Resize(1, 21)
                        .Font.Name = "Calibri"
                        .Font.Size = 9
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                    End With
                    materiaisSheetLastRow = materiaisSheetLastRow + 1
                End If
            Next i
        End If
ContinueToNextSheet:
    Next index
    
    Range("A2").Value = 1
    Range("A3").Value = 2

    ' Autofill numbers from A2 to the last row in column A
    On Error Resume Next
    Range("A2:A3").AutoFill Destination:=Range("A2:A" & materiaisSheetLastRow - 1), Type:=xlFillDefault
    On Error GoTo 0
    
    materiaisSheet.Range("A" & materiaisSheetLastRow & ":Z" & materiaisSheet.Rows.Count).Clear
    
    ' Autosize columns A to U
    materiaisSheet.Columns("A:U").AutoFit
    
    ' Check if the current width is greater than 50 points
    For Each col In materiaisSheet.Columns
        ' Check if column width is greater than 50
        If col.ColumnWidth > 50 Then
            ' If so, set the column width to 50
            col.ColumnWidth = 50
        End If
    Next col
    
    OptimizedMode False
    
    Debug.Print "Total Time " & Timer - tempIni
    
    countExecution "Worksheet_Activate", False, "Event", "Planilha4(Materiais)"
    
End Sub

