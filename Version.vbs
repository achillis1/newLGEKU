Attribute VB_Name = "Version"
Sub SaveCodeModules()
' http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
'This code Exports all VBA modules
Dim i%, sName$
With ThisWorkbook.VBProject
For i% = 1 To .VBComponents.Count
If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
sName$ = .VBComponents(i%).CodeModule.Name
.VBComponents(i%).Export sName$ & ".vbs"
End If
Next i
End With
End Sub
