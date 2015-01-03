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


Private Sub UserForm_Initialize()
    lastenroll = Worksheets("Enrollment").Range("C" & Rows.Count).End(xlUp).Row
    For i = 2 To lastenroll
        frmEnrollment.Controls("lstEnrollments").AddItem (Cells(i, 3).Value)
    Next i
End Sub

Private Sub UserForm_Terminate()
    frmServiceCenter.Show
End Sub


