VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Results_ROSA 
   Caption         =   "Load_Results_ROSA"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   OleObjectBlob   =   "Load_Results_ROSA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Results_ROSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cancel_Load_Results_Click()
Me.Hide
frmProcessing.Show vbModeless
End Sub


Private Sub Scheduled_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.Enrollment_ID_ROSA = Scheduled_Listbox.Value



End Sub

Private Sub UserForm_Initialize()
Set wsdb = Worksheets("Enrollments")

'last row database
wsDblr = wsdb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "Scheduled" Then
        'push data from database to form
        'ROSA Scheduling
        With Scheduled_Listbox
            .AddItem wsdb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA)
        End With
              
    End If
Next x

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
