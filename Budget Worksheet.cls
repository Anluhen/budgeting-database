VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    
    Tools.CheckDate

    'GetColumnsNumbers
    
    'CompareSheets
    
    'Application.Run ("VersionAndUpdate.CorrectErrors")
    
    VersionAndUpdate.CorrectErrors
    
    UpdateUpdater ThisWorkbook
    
    UpdateWorkbookVBA True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim BackupPath As String
    Dim FileName As String
    Dim NewFileName As String
    
    ' Get the path of the current workbook
    BackupPath = ThisWorkbook.Path
    ' Get the file name without extension
    FileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ' Generate new file name for backup
    NewFileName = "BACKUP - " & FileName & ".xlsm"
    
    ' Save a copy of the file in the same folder with "BACKUP - " prefix
    On Error Resume Next
    ThisWorkbook.SaveCopyAs BackupPath & "\" & NewFileName
    On Error GoTo 0
    
' This part creates a full backup file withh all sheets for error handling when this version is released.
' After enough tests this maybe can be deleted.
    
    BackupPath = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\Backups"
    ' Get the file name without extension
    FileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ' Generate new file name for backup
    NewFileName = FileName & " - " & Format(Now, "dd-mm-yy hh\hmm") & ".xlsm"
    
    ' Save a copy of the file in the folder \\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\Backupk
    On Error Resume Next
    ThisWorkbook.SaveCopyAs BackupPath & "\" & NewFileName
    On Error GoTo 0
    
    OptimizedMode False
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    ' Variables for date comparison
    Dim fs As Object
    Dim f As Object
    Dim thisModifiedDate As Date
    Dim currentTime As Date
    Dim minutesSinceLastSave As Double
    
    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    thisModifiedDate = f.DateLastModified
    
    ' Get the current time
    currentTime = Now
    
    ' Calculate the difference in minutes
    minutesSinceLastSave = DateDiff("n", thisModifiedDate, currentTime)
    
    ' Check if the workbook has been saved in the last 10 minutes
    If minutesSinceLastSave >= 10 And minutesSinceLastSave <= 480 Then
        ThisWorkbook.Save
    End If

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Target.Address = "$A$1:$B$1" Then
        With ConsultaBancoDeDados
        '.StartUpPosition = 1 ' CenterOwner
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless
        End With
        
        If componenteColumn <> 2 Then
            GetColumnsNumbers
        End If
    End If
    
End Sub
    
Sub UpdateUpdater(targetWorkbook As Workbook)

'countExecution "UpdateUpdater", True, "Sub", "VersionAndUpdate"
    
    ' Copy all code from every sheet, module, and form from the source workbook
    Dim sourceVbComp As VBComponent
    Dim targetVbComp As VBComponent
    Dim lineNum As Long
    Dim Prompt As Boolean
    
    'countExecution "UpdateWorkbookVBA", True, "Sub", "VersionAndUpdate"
    
    OptimizedMode True
    
    Dim sourceWorkbookPath As String
    Dim sourceWorkbook As Workbook
    Dim result As VbMsgBoxResult
    Dim firstUpdateConfirmed As Boolean
    
    If targetWorkbook.Sheets("S.PROP").Range("A1").Value <> "" Then
        GoTo ExitSub
    End If
    
    ' Check if the opened workbook path is the source path
    If targetWorkbook.Path & "\" & targetWorkbook.Name = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" Then
        GoTo ExitSub
    End If
    
    ' Check if current workbook name is the same as the origin file
    If targetWorkbook.Name = "FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" Then
        MsgBox "Não foi possível verificar atualizações da planilha. Renomeie a planilha primeiro."
        GoTo ExitSub
    End If
    
    ' Verificar a data da planilha e forçar a atualização automática
    Dim fs As Object
    Dim f As Object
    Dim thisModfiedDate As Date
    Dim baseModfiedDate As Date
    
    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(targetWorkbook.Path & "\" & targetWorkbook.Name)
    thisModfiedDate = f.DateLastModified
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile("\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS" & "\" & "FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm")
    baseModfiedDate = f.DateLastModified
    
    If baseModfiedDate <= thisModfiedDate And thisModfiedDate <= Date - 7 Then
        GoTo ExitSub
    ElseIf thisModfiedDate >= Date - 7 Then
        Prompt = False
    End If
    
    ' Set the path to the source workbook
    sourceWorkbookPath = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" ' Change this path
    
    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourceWorkbookPath)

    ' Check if the source workbook has macros
    On Error GoTo ErrorHandler
    If Not sourceWorkbook.VBProject Is Nothing Then
        On Error GoTo 0
    End If
    
    If True = False Then
ErrorHandler:
        On Error GoTo 0
        MsgBox "A planilha não pode ser atualizada. Verifique a Central de Confiabilidade para permissão de acesso."
        sourceWorkbook.Close savechanges:=False
        GoTo ExitSub
    End If

    ' Copy all code from every sheet, module, and form from the source workbook
    Dim vbCompNotFound As Boolean
    
    If Prompt Then
        firstUpdateConfirmed = False
    Else
        firstUpdateConfirmed = True
    End If
    
    ' Loop through each VBComponent in the source workbook
    For Each sourceVbComp In sourceWorkbook.VBProject.VBComponents
         For Each targetVbComp In targetWorkbook.VBProject.VBComponents
            ' Check if the targetVbComp is the same as sourceVbComp
            If targetVbComp.Name = sourceVbComp.Name And targetVbComp.Name = "VersionAndUpdate" Then
                If sourceVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines) <> targetVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines) Then
                    ' Replace the existing code in the target VBComponent with the code from the source VBComponent
                    targetVbComp.CodeModule.DeleteLines 1, targetVbComp.CodeModule.CountOfLines
                    targetVbComp.CodeModule.AddFromString sourceVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines)
                End If
            End If
        Next targetVbComp
    Next sourceVbComp
    
    targetWorkbook.Activate
    
    targetWorkbook.Save
    
ExitSub:

    OptimizedMode False
    
'countExecution "UpdateUpdater", False, "Sub", "VersionAndUpdate"

    'sourceWorkbook.Close savechanges:=False
    
End Sub
