VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProcessing 
   Caption         =   "Processing"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   OleObjectBlob   =   "frmProcessing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

