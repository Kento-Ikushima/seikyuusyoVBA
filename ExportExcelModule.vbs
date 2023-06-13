Dim objParams, strFullPath, strFileName, objExcel, objWorkBook
Dim objTempComponent, strExportPath, strCode 
Dim FSO

strFullPath=""
strExportPath = ""
strFileName = ""
strFilePath = ""

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objParams = WScript.Arguments
If objParams.Count = 2 Then
    strFullPath = objParams.Item(0)
    strExportPath = objParams.Item(1)
    strFileName = FSO.GetFileName(strFullPath)
    strFilePath = FSO.GetParentFolderName(strFullPath)
    'WScript.Echo "strFullPath---->" & strFullPath
    'WScript.Echo "strFileName---->" & strFileName
    'WScript.Echo "strFilePath---->" & strFilePath
    'WScript.Echo "strExportPath---->" & strExportPath

Else
    WScript.Quit 0
End If

'Excel操作準備
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False
'マクロが無効の状態で開く
'※だめだった！！objExcel.AutomationSecurity = msoAutomationSecurityForceDisable
Set objWorkBook = objExcel.Workbooks.Open(strFullPath)

'ソースをエクスポートする
Call ExportSource()

'Excelをクローズ
Set FSO = nothing
Set objParams = nothing
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True
objWorkBook.Close False
objExcel.Quit
Set objWorkBook = nothing
Set objExcel = nothing

'--------------------------------------------------------------------------
'ソースをエクスポートする
'--------------------------------------------------------------------------
Sub ExportSource()
    For Each TempComponent In objWorkBook.VBProject.VBComponents
        If TempComponent.CodeModule.CountOfDeclarationLines <> TempComponent.CodeModule.CountOfLines Then
            Select Case TempComponent.Type
                'STANDARD_MODULE
                Case 1
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".bas"
                'CLASS_MODULE
                Case 2
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".cls"
                'USER_FORM
                Case 3
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".frm"
                'SHEETとThisWorkBook
                Case 100
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".bas"
            End Select
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)                    
            End With
        End If
    Next

End Sub
