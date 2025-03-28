Attribute VB_Name = "VersionAndUpdate"
Function CheckVersion() As String
    CheckVersion = "0.1.1"
End Function

Sub UpdateWorkbookVBA(Optional Prompt As Boolean = True)
    
    'countExecution "UpdateWorkbookVBA", True, "Sub", "VersionAndUpdate"
    
    OptimizedMode True
    
    Dim sourceWorkbookPath As String
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim result As VbMsgBoxResult
    Dim firstUpdateConfirmed As Boolean
    
    If ThisWorkbook.Sheets("S.PROP").Range("A1").Value <> "" Then
        GoTo ExitSub
    End If
    
    ' Check if the opened workbook path is the source path
    If ThisWorkbook.Path & "\" & ThisWorkbook.Name = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" Then
        GoTo ExitSub
    End If

    ' Set the path to the source workbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourceWorkbookPath = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" ' Change this path
    
    ' Check if the source path exists
    If Not fso.FolderExists(sourceWorkbookPath) Then
        Debug.Print "The source directory was not found"
        'GoTo ExitSub
    End If
    
    ' Check if current workbook name is the same as the origin file
    If ThisWorkbook.Name = "FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm" Then
        MsgBox "Não foi possível verificar atualizações da planilha. Renomeie a planilha primeiro."
        GoTo ExitSub
    End If
    
    If ThisWorkbook.Sheets("S.PROP").Range("A1").Value <> "" Then
        GoTo ExitSub
    End If
    
    ' Verificar a data da planilha e forçar a atualização automática
    Dim fs As Object
    Dim f As Object
    Dim thisModfiedDate As Date
    Dim baseModfiedDate As Date
    
    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    thisModfiedDate = f.DateLastModified
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile("\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS" & "\" & "FOR-COM-01 PLANILHA DE ORÇAMENTO_V0.1.1.xlsm")
    baseModfiedDate = f.DateLastModified
    
    If baseModfiedDate <= thisModfiedDate And thisModfiedDate <= Date - 7 Then
        GoTo ExitSub
    ElseIf thisModfiedDate >= Date - 7 Then
        Prompt = False
    End If
    
    If Prompt Then
        Update = MsgBox("Versão: " & CheckVersion & vbCrLf & vbCrLf & "Deseja verificar atualizações?" & vbCrLf & vbCrLf & "(Esta mensagem não significa que existem atualizações a serem feitas)", vbYesNo + vbQuestion, "Confirmation")
    End If
    
    If Update = vbYes Then
        Prompt = False
        firstUpdateConfirmed = True
    ElseIf Update = vbNo Then
        GoTo ExitSub
    End If
    
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

    ' Set the reference to the opened workbook
    Set targetWorkbook = ThisWorkbook

    ' Copy all code from every sheet, module, and form from the source workbook
    Dim sourceVbComp As VBComponent
    Dim targetVbComp As VBComponent
    Dim lineNum As Long
    Dim vbCompNotFound As Boolean
    
    If Prompt Then
        firstUpdateConfirmed = False
    Else
        firstUpdateConfirmed = True
    End If
    
    'Correct errors before moving on. This makes sure that you can change the VBComponents before checking each one of them.
    Application.Run "'" & targetWorkbook.Name & "'!" & "VersionAndUpdate.CorrectErrors"
    
    ' Loop through each VBComponent in the source workbook
    For Each sourceVbComp In sourceWorkbook.VBProject.VBComponents
         vbCompNotFound = True
         
         For Each targetVbComp In targetWorkbook.VBProject.VBComponents
            ' Check if the targetVbComp is the same as sourceVbComp
            If targetVbComp.Name = sourceVbComp.Name And targetVbComp.Name <> "VersionAndUpdate" And sourceVbComp.CodeModule.CountOfLines > 0 Then
                vbCompNotFound = False
                
                'Check if the code is diferent
                If sourceVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines) <> targetVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines) Then
                    ' Ask for update confirmation only for the first code module
                    If Not firstUpdateConfirmed Then
                        If Prompt Then
                            result = MsgBox("A pasta de trabalho está desatualizada. Deseja atualizar?", vbYesNo + vbQuestion, "Confirmation")
                        Else
                            result = vbYes
                        End If
                        
                        ' Check the user's response
                        If result = vbYes Then
                            ' User clicked Yes
                            firstUpdateConfirmed = True ' Set the flag to true after the first update confirmation
                        Else
                            sourceWorkbook.Close False
                            GoTo ExitSub
                        End If
                    End If
                    
                    If firstUpdateConfirmed Then
                        ' Replace the existing code in the target VBComponent with the code from the source VBComponent
                        targetVbComp.CodeModule.DeleteLines 1, targetVbComp.CodeModule.CountOfLines
                        targetVbComp.CodeModule.AddFromString sourceVbComp.CodeModule.Lines(1, sourceVbComp.CodeModule.CountOfLines)
                    End If
                End If
            End If
        Next targetVbComp
        
        If vbCompNotFound And sourceVbComp.Name <> "VersionAndUpdate" And sourceVbComp.CodeModule.CountOfLines > 0 Then
            MsgBox ("O módulo VBA da planilha " & sourceVbComp.Name & " não foi encontrado. Impossível atualizar.")
        End If
    Next sourceVbComp
    
    targetWorkbook.Save
    
    sourceWorkbook.Close savechanges:=False
    
ExitSub:

    OptimizedMode False

    'countExecution "UpdateWorkbookVBA", False, "Sub", "VersionAndUpdate"
    
End Sub

Sub CorrectErrors(Optional lockFromMacroList As Boolean = True)
' This function server the purpose to correct formula or formating errors on the copies of the sheet.
' It executes before checking for updates.
' This can be errased when the version of the workbook changes.

'countExecution "UpdateUpdater", True, "Function", "CorrectErrors"

'OptimizedMode True

    Dim vbComp As VBComponent
    Dim FilePath As String
    FilePath = "\\dados\comercial_vendas\300 - PLANILHA DE CUSTOS E BANCO DE DADOS\ConsultaBancoDeDados.frm" ' Update this with the path to your UserForm file

    ' Delete existing UserForm if it exists
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents("ConsultaBancoDeDados")
    If vbComp.Designer.Controls("AmbosDB") Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
        
        ' Import new UserForm
        ThisWorkbook.VBProject.VBComponents.Import FilePath

    End If
    On Error GoTo 0

    
'OptimizedMode False

'countExecution "UpdateUpdater", False, "Function", "CorrectErrors"

End Sub

Sub UpdateUpdater(sourceWorkbook As Workbook, targetWorkbook As Workbook)
    Application.Run "'" & sourceWorkbook.Name & "'!" & "EstaPastaDeTrabalho.UpdateUpdater", targetWorkbook
End Sub

