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

ActiveWorkbook.VBProject.VBComponents.Item(1).Export "C:\Users\dli\Documents\Trash\11\a.frm"
ActiveWorkbook.VBProject.VBComponents("frmAdmin").Export "C:\Users\dli\Documents\Trash\11\b.frm"
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

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub

