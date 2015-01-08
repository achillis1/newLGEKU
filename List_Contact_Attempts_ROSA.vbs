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
frmProcessing.Show
End Sub

Private Sub Clear_ROSA_Click()
'Clear Fields
Enrollment_Listbox.Clear
        Me.Enrollment_ID_ROSA = ""
        Me.First_Contact_Attempt_Notes_ROSA = ""
        Me.First_Contact_Attempt_Type_ROSA = ""
        Me.Second_Contact_Attempt_Notes_ROSA = ""
        Me.Second_Contact_Attempt_Type_ROSA = ""
        Me.Third_Contact_Attempt_Notes_ROSA = ""
        Me.Third_Contact_Attempt_Type_ROSA = ""
        Me.Fourth_Contact_Attempt_Notes_ROSA = ""
        Me.Fourth_Contact_Attempt_Type_ROSA = ""
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
wsDblr = wsdb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).Row


Me.First_Contact_Attempt_Notes_ROSA.Enabled = True
Me.First_Contact_Attempt_Notes_ROSA.BackColor = rgbWhite
Me.First_Contact_Attempt_Type_ROSA.Enabled = True
Me.First_Contact_Attempt_Type_ROSA.BackColor = rgbWhite
Me.Second_Contact_Attempt_Notes_ROSA.Enabled = True
Me.Second_Contact_Attempt_Notes_ROSA.BackColor = rgbWhite
Me.Second_Contact_Attempt_Type_ROSA.Enabled = True
Me.Second_Contact_Attempt_Type_ROSA.BackColor = rgbWhite
Me.Third_Contact_Attempt_Notes_ROSA.Enabled = True
Me.Third_Contact_Attempt_Notes_ROSA.BackColor = rgbWhite
Me.Third_Contact_Attempt_Type_ROSA.Enabled = True
Me.Third_Contact_Attempt_Type_ROSA.BackColor = rgbWhite
Me.Fourth_Contact_Attempt_Notes_ROSA.Enabled = True
Me.Fourth_Contact_Attempt_Notes_ROSA.BackColor = rgbWhite
Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = True
Me.Fourth_Contact_Attempt_Type_ROSA.BackColor = rgbWhite
Me.Fifth_Contact_Attempt_Notes_ROSA.Enabled = True
Me.Fifth_Contact_Attempt_Notes_ROSA.BackColor = rgbWhite
Me.Fifth_Contact_Attempt_Type_ROSA.Enabled = True
Me.Fifth_Contact_Attempt_Type_ROSA.BackColor = rgbWhite


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

If Me.First_Contact_Attempt_Type_ROSA = "" Then
    Me.Second_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
ElseIf Me.Second_Contact_Attempt_Type_ROSA = "" Then
    Me.First_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.First_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_ROSA.Enabled = False
    Me.First_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
ElseIf Me.Third_Contact_Attempt_Type_ROSA = "" Then
    Me.First_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.First_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_ROSA.Enabled = False
    Me.First_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
ElseIf Me.Fourth_Contact_Attempt_Type_ROSA = "" Then
    Me.First_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.First_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_ROSA.Enabled = False
    Me.First_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fifth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
ElseIf Me.Fifth_Contact_Attempt_Type_ROSA = "" Then
    Me.First_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.First_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_ROSA.Enabled = False
    Me.First_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Second_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Third_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_ROSA.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = False
    Me.Fourth_Contact_Attempt_Type_ROSA.BackColor = rgbLightGrey
End If
'IF DATA CHECK HERE THEN
'First_Contact_Attempt_Date_ROSA.Enabled = False
'First_Contact_Attempt_Date_ROSA.BackColor = &H80000005


End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Save_ROSA_Click()

Set wsdb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Me.Enrollment_ID_ROSA
'last row database
wsDblr = wsdb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).Row

'Verify that the values have been added to the Fields
If Me.First_Contact_Attempt_Type_ROSA.Enabled = True Then
    If Me.First_Contact_Attempt_Type_ROSA = "" Or Me.First_Contact_Attempt_Notes_ROSA = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Second_Contact_Attempt_Type_ROSA.Enabled = True Then
    If Me.Second_Contact_Attempt_Type_ROSA = "" Or Me.Second_Contact_Attempt_Notes_ROSA = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Third_Contact_Attempt_Type_ROSA.Enabled = True Then
    If Me.Third_Contact_Attempt_Type_ROSA = "" Or Me.Third_Contact_Attempt_Notes_ROSA = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Fourth_Contact_Attempt_Type_ROSA.Enabled = True Then
    If Me.Fourth_Contact_Attempt_Type_ROSA = "" Or Me.Fourth_Contact_Attempt_Notes_ROSA = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Fifth_Contact_Attempt_Date_ROSA.Enabled = True Then
    If Me.Fifth_Contact_Attempt_Type_ROSA = "" Or Me.Fifth_Contact_Attempt_Notes_ROSA = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If


For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = EID Then
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA) = Me.First_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA) = Me.First_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_ROSA) = Me.Second_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_ROSA) = Me.Second_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_ROSA) = Me.Third_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_ROSA) = Me.Third_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_ROSA) = Me.Fourth_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_ROSA) = Me.Fourth_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_ROSA) = Me.Fifth_Contact_Attempt_Notes_ROSA
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_ROSA) = Me.Fifth_Contact_Attempt_Type_ROSA
         wsdb.Cells(x, NexantEnrollments.Schedule_Date_ROSA) = Me.Schedule_Date_ROSA
         wsdb.Cells(x, NexantEnrollments.Schedule_Time_ROSA) = Me.Schedule_Time_ROSA
         wsdb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
         wsdb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
         
         
       'Select the appropiate status and update the specific dates/times
        If Me.Fifth_Contact_Attempt_Type_ROSA <> "" Then
            If Me.Schedule_Date_ROSA = "" Then
                wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsdb.Cells(x, NexantEnrollments.PENDING_5_date_set_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.PENDING_5_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Fifth_Contact_Attempt_Notes_ROSA
                wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Fifth_Contact_Attempt_Type_ROSA
                wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            Else
                wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Fifth_Contact_Attempt_Notes_ROSA
                wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Fifth_Contact_Attempt_Type_ROSA
                wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            End If

            ElseIf Me.Fourth_Contact_Attempt_Type_ROSA <> "" Then
                If Me.Schedule_Date_ROSA = "" Then
                    wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
                    wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                    wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsdb.Cells(x, NexantEnrollments.PENDING_4_date_set_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.PENDING_4_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Fourth_Contact_Attempt_Notes_ROSA
                    wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Fourth_Contact_Attempt_Type_ROSA
                    wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                Else
                    wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
                    wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                    wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Fourth_Contact_Attempt_Notes_ROSA
                    wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Fourth_Contact_Attempt_Type_ROSA
                    wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                    wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                End If
                ElseIf Me.Third_Contact_Attempt_Type_ROSA <> "" Then
                    If Me.Schedule_Date_ROSA = "" Then
                        wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
                        wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                        wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsdb.Cells(x, NexantEnrollments.PENDING_3_date_set_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.PENDING_3_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Third_Contact_Attempt_Notes_ROSA
                        wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Third_Contact_Attempt_Type_ROSA
                        wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    Else
                        wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
                        wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                        wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Third_Contact_Attempt_Notes_ROSA
                        wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Third_Contact_Attempt_Type_ROSA
                        wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                        wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    End If
                    ElseIf Me.Second_Contact_Attempt_Type_ROSA <> "" And Me.Schedule_Date_ROSA = "" Then
                        If Me.Schedule_Date_ROSA = "" Then
                            wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
                            wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                            wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsdb.Cells(x, NexantEnrollments.PENDING_2_date_set_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.PENDING_2_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Second_Contact_Attempt_Notes_ROSA
                            wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Second_Contact_Attempt_Type_ROSA
                            wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        Else
                            wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
                            wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                            wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Second_Contact_Attempt_Notes_ROSA
                            wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Second_Contact_Attempt_Type_ROSA
                            wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                            wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        End If
                        ElseIf Me.First_Contact_Attempt_Type_ROSA <> "" Then
                            If Me.Schedule_Date_ROSA = "" Then
                                wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
                                wsdb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                                wsdb.Cells(x, NexantEnrollments.PENDING_1_date_set_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.PENDING_1_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                                wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.First_Contact_Attempt_Notes_ROSA
                                wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.First_Contact_Attempt_Type_ROSA
                                wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            Else
                                wsdb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
                                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsdb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.First_Contact_Attempt_Notes_ROSA
                                wsdb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.First_Contact_Attempt_Type_ROSA
                                wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA).NumberFormat = "@"
                                wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            End If
                        
        End If
           
    End If
Next x
Call Clear_ROSA_Click
End Sub

Private Sub UserForm_Initialize()

Set wsdb = Worksheets("Enrollments")

'last row database
wsDblr = wsdb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).Row

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

With First_Contact_Attempt_Type_ROSA
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Second_Contact_Attempt_Type_ROSA
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Third_Contact_Attempt_Type_ROSA
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Fourth_Contact_Attempt_Type_ROSA
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Fifth_Contact_Attempt_Type_ROSA
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

End Sub
