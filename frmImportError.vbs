VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportError 
   Caption         =   "Import Error"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   OleObjectBlob   =   "frmImportError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Me.Hide
    frmImport.Hide
    frmServiceCenter.Show vbModelless
    
End Sub





Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

