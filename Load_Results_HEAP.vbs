VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Results_HEAP 
   Caption         =   "Load_Results_HEAP"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   OleObjectBlob   =   "Load_Results_HEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Results_HEAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cancel_Load_Results_Click()
Me.Hide
frmProcessing.Show vbModeless
End Sub


Private Sub Scheduled_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.Enrollment_ID_HEAP = Scheduled_Listbox.Value



End Sub

Private Sub UserForm_Initialize()
Set wsdb = Worksheets("Enrollments")

'last row database
wsDblr = wsdb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Status_HEAP) = "Scheduled" Then
        'push data from database to form
        'HEAP Scheduling
        With Scheduled_Listbox
            .AddItem wsdb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
        End With
              
    End If
Next x

End Sub

