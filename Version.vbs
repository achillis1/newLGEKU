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
