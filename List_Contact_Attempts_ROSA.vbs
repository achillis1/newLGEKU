VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} List_Contact_Attempts_ROSA 
   Caption         =   "List_Contact_Attempts_ROSA"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15600
   OleObjectBlob   =   "List_Contact_Attempts_ROSA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "List_Contact_Attempts_ROSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub Cancel_ROSA_Click()
Unload Me
frmAdmin.Show
End Sub

Private Sub Clear_ROSA_Click()
'Clear Fields
Enrollment_Listbox.Clear
        Me.Enrollment_ID_ROSA = ""
        Me.First_Contact_Attempt_Date_ROSA.Value = ""
        Me.First_Contact_Attempt_Notes_ROSA = ""
        Me.First_Contact_Attempt_Type_ROSA = ""
        Me.Second_Contact_Attempt_Date_ROSA = ""
        Me.Second_Contact_Attempt_Notes_ROSA = ""
        Me.Second_Contact_Attempt_Type_ROSA = ""
        Me.Third_Contact_Attempt_Date_ROSA = ""
        Me.Third_Contact_Attempt_Notes_ROSA = ""
        Me.Third_Contact_Attempt_Type_ROSA = ""
        Me.Fourth_Contact_Attempt_Date_ROSA = ""
        Me.Fourth_Contact_Attempt_Notes_ROSA = ""
        Me.Fourth_Contact_Attempt_Type_ROSA = ""
        Me.Fifth_Contact_Attempt_Date_ROSA = ""
        Me.Fifth_Contact_Attempt_Notes_ROSA = ""
        Me.Fifth_Contact_Attempt_Type_ROSA = ""
        Me.Schedule_Date_ROSA = ""
        Me.Schedule_Time_ROSA = ""

Call UserForm_Initialize
End Sub
Private Sub Enrollment_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Set wsdb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Enrollment_Listbox.Value
'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'Retrive values from Database
Me.Enrollment_ID_ROSA = EID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = EID Then
        'push data from database to form
        'ROSA Scheduling

        Me.First_Contact_Attempt_Date_ROSA = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA)
        Me.First_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA)
        Me.First_Contact_Attempt_Type_ROSA = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA)
        Me.Second_Contact_Attempt_Date_ROSA = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA)
        Me.Second_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_ROSA)
        Me.Second_Contact_Attempt_Type_ROSA = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_ROSA)
        Me.Third_Contact_Attempt_Date_ROSA = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA)
        Me.Third_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_ROSA)
        Me.Third_Contact_Attempt_Type_ROSA = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_ROSA)
        Me.Fourth_Contact_Attempt_Date_ROSA = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA)
        Me.Fourth_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_ROSA)
        Me.Fourth_Contact_Attempt_Type_ROSA = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_ROSA)
        Me.Fifth_Contact_Attempt_Date_ROSA = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA)
        Me.Fifth_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_ROSA)
        Me.Fifth_Contact_Attempt_Type_ROSA = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_ROSA)
        Me.Schedule_Date_ROSA = wsdb.Cells(x, NexantEnrollments.Schedule_Date_ROSA)
        Me.Schedule_Time_ROSA = wsdb.Cells(x, NexantEnrollments.Schedule_Time_ROSA)
        
      

    End If
Next x

'IF DATA CHECK HERE THEN
First_Contact_Attempt_Date_ROSA.Enabled = False
First_Contact_Attempt_Date_ROSA.BackColor = &H80000005


End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Save_ROSA_Click()

Set wsdb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = currentEnrollment
'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = EID Then
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA) = Me.First_Contact_Attempt_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA) = Me.First_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA) = Me.First_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA) = Me.Second_Contact_Attempt_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_ROSA) = Me.Second_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_ROSA) = Me.Second_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA) = Me.Third_Contact_Attempt_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_ROSA) = Me.Third_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_ROSA) = Me.Third_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA) = Me.Fourth_Contact_Attempt_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_ROSA) = Me.Fourth_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_ROSA) = Me.Fourth_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA) = Me.Fifth_Contact_Attempt_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_ROSA) = Me.Fifth_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_ROSA) = Me.Fifth_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Schedule_Date_ROSA) = Me.Schedule_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Schedule_Time_ROSA) = Me.Schedule_Time_ROSA
    End If
Next x

End Sub

Private Sub UserForm_Initialize()

Set wsdb = Worksheets("Enrollments")

'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "RECEIVED AT VENDOR" Or _
    wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "FIRST CONTACT" Or _
    wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING" Then
        'push data from database to form
        'ROSA Scheduling
        With Enrollment_Listbox
            .AddItem wsdb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA)
        End With
              
    End If
Next x




End Sub
