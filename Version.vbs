Attribute VB_Name = "Version"
Sub SaveCodeModules()
'This code Exports all VBA modules
Dim i%, sName$
With ThisWorkbook.VBProject
For i% = 1 To .VBComponents.Count
If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
sName$ = .VBComponents(i%).CodeModule.Name
.VBComponents(i%).Export "C:\share\DriveZ\LGE\documents\newrepository\" & sName$ & ".vbs"
End If
Next i
End With
End Sub

'Sub SaveVBACode()
'
''This code Exports all VBA modules
'Dim i As Integer
'Dim mName As String
'Dim Fname As String
'
'
'    With ThisWorkbook
'        For i = 1 To .VBProject.VBComponents.Count
'            If .VBProject.VBComponents(i).CodeModule.CountOfLines > 0 Then
'                mName = .VBProject.VBComponents(i%).CodeModule.Name
'
'                 Fname = .Path & "\" & mName & ".txt"
'                .VBProject.VBComponents(mName).Export Fname
'            End If
'        Next i
'    End With
'
'End Sub
