VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} List_Contact_Attempts_HEAP 
   Caption         =   "List_Contact_Attempts_HEAP"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15180
   OleObjectBlob   =   "List_Contact_Attempts_HEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "List_Contact_Attempts_HEAP"
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

Private Sub Cancel_HEAP_Click()
Unload Me
frmProcessing.Show
End Sub

Private Sub Clear_HEAP_Click()
'Clear Fields
Enrollment_Listbox.Clear
        Me.Enrollment_ID_HEAP = ""
        Me.First_Contact_Attempt_Notes_HEAP = ""
        Me.First_Contact_Attempt_Type_HEAP = ""
        Me.Second_Contact_Attempt_Notes_HEAP = ""
        Me.Second_Contact_Attempt_Type_HEAP = ""
        Me.Third_Contact_Attempt_Notes_HEAP = ""
        Me.Third_Contact_Attempt_Type_HEAP = ""
        Me.Fourth_Contact_Attempt_Notes_HEAP = ""
        Me.Fourth_Contact_Attempt_Type_HEAP = ""
        Me.Fifth_Contact_Attempt_Notes_HEAP = ""
        Me.Fifth_Contact_Attempt_Type_HEAP = ""
        Me.Schedule_Date_HEAP = ""
        Me.Schedule_Time_HEAP = ""

Call UserForm_Initialize
End Sub
Private Sub Enrollment_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Set wsDb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Enrollment_Listbox.Value
'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row


Me.First_Contact_Attempt_Notes_HEAP.Enabled = True
Me.First_Contact_Attempt_Notes_HEAP.BackColor = rgbWhite
Me.First_Contact_Attempt_Type_HEAP.Enabled = True
Me.First_Contact_Attempt_Type_HEAP.BackColor = rgbWhite
Me.Second_Contact_Attempt_Notes_HEAP.Enabled = True
Me.Second_Contact_Attempt_Notes_HEAP.BackColor = rgbWhite
Me.Second_Contact_Attempt_Type_HEAP.Enabled = True
Me.Second_Contact_Attempt_Type_HEAP.BackColor = rgbWhite
Me.Third_Contact_Attempt_Notes_HEAP.Enabled = True
Me.Third_Contact_Attempt_Notes_HEAP.BackColor = rgbWhite
Me.Third_Contact_Attempt_Type_HEAP.Enabled = True
Me.Third_Contact_Attempt_Type_HEAP.BackColor = rgbWhite
Me.Fourth_Contact_Attempt_Notes_HEAP.Enabled = True
Me.Fourth_Contact_Attempt_Notes_HEAP.BackColor = rgbWhite
Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = True
Me.Fourth_Contact_Attempt_Type_HEAP.BackColor = rgbWhite
Me.Fifth_Contact_Attempt_Notes_HEAP.Enabled = True
Me.Fifth_Contact_Attempt_Notes_HEAP.BackColor = rgbWhite
Me.Fifth_Contact_Attempt_Type_HEAP.Enabled = True
Me.Fifth_Contact_Attempt_Type_HEAP.BackColor = rgbWhite


'Retrive values from Database
Me.Enrollment_ID_HEAP = EID
For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
        'push data from database to form
        'HEAP Scheduling

        Me.First_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP)
        Me.First_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP)
        Me.First_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP)
        Me.Second_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP)
        Me.Second_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP)
        Me.Second_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP)
        Me.Third_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP)
        Me.Third_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP)
        Me.Third_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP)
        Me.Fourth_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP)
        Me.Fourth_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP)
        Me.Fourth_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP)
        Me.Fifth_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP)
        Me.Fifth_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP)
        Me.Fifth_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP)
        Me.Schedule_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP)
        Me.Schedule_Time_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP)
        
    End If
Next x

If Me.First_Contact_Attempt_Type_HEAP = "" Then
    Me.Second_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
ElseIf Me.Second_Contact_Attempt_Type_HEAP = "" Then
    Me.First_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.First_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_HEAP.Enabled = False
    Me.First_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
ElseIf Me.Third_Contact_Attempt_Type_HEAP = "" Then
    Me.First_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.First_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_HEAP.Enabled = False
    Me.First_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
ElseIf Me.Fourth_Contact_Attempt_Type_HEAP = "" Then
    Me.First_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.First_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_HEAP.Enabled = False
    Me.First_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fifth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fifth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
ElseIf Me.Fifth_Contact_Attempt_Type_HEAP = "" Then
    Me.First_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.First_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.First_Contact_Attempt_Type_HEAP.Enabled = False
    Me.First_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Second_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Second_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Third_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Third_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Notes_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Notes_HEAP.BackColor = rgbLightGrey
    Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = False
    Me.Fourth_Contact_Attempt_Type_HEAP.BackColor = rgbLightGrey
End If
'IF DATA CHECK HERE THEN
'First_Contact_Attempt_Date_HEAP.Enabled = False
'First_Contact_Attempt_Date_HEAP.BackColor = &H80000005


End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Save_HEAP_Click()

Set wsDb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Me.Enrollment_ID_HEAP
'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row

'Verify that the values have been added to the Fields
If Me.First_Contact_Attempt_Type_HEAP.Enabled = True Then
    If Me.First_Contact_Attempt_Type_HEAP = "" Or Me.First_Contact_Attempt_Notes_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Second_Contact_Attempt_Type_HEAP.Enabled = True Then
    If Me.Second_Contact_Attempt_Type_HEAP = "" Or Me.Second_Contact_Attempt_Notes_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Third_Contact_Attempt_Type_HEAP.Enabled = True Then
    If Me.Third_Contact_Attempt_Type_HEAP = "" Or Me.Third_Contact_Attempt_Notes_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Fourth_Contact_Attempt_Type_HEAP.Enabled = True Then
    If Me.Fourth_Contact_Attempt_Type_HEAP = "" Or Me.Fourth_Contact_Attempt_Notes_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If
If Me.Fifth_Contact_Attempt_Date_HEAP.Enabled = True Then
    If Me.Fifth_Contact_Attempt_Type_HEAP = "" Or Me.Fifth_Contact_Attempt_Notes_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If
End If


For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
         wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP) = Me.First_Contact_Attempt_Notes_HEAP
         wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP) = Me.First_Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP) = Me.Second_Contact_Attempt_Notes_HEAP
         wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP) = Me.Second_Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP) = Me.Third_Contact_Attempt_Notes_HEAP
         wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP) = Me.Third_Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP) = Me.Fourth_Contact_Attempt_Notes_HEAP
         wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP) = Me.Fourth_Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP) = Me.Fifth_Contact_Attempt_Notes_HEAP
         wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP) = Me.Fifth_Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP) = Me.Schedule_Date_HEAP
         wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP) = Me.Schedule_Time_HEAP
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
         
         
       'Select the appropiate status and update the specific dates/times
        If Me.Fifth_Contact_Attempt_Type_HEAP <> "" Then
            If Me.Schedule_Date_HEAP = "" Then
                wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Fifth_Contact_Attempt_Notes_HEAP
                wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Fifth_Contact_Attempt_Type_HEAP
                wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            Else
                wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Fifth_Contact_Attempt_Notes_HEAP
                wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Fifth_Contact_Attempt_Type_HEAP
                wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            End If

            ElseIf Me.Fourth_Contact_Attempt_Type_HEAP <> "" Then
                If Me.Schedule_Date_HEAP = "" Then
                    wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
                    wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                    wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Fourth_Contact_Attempt_Notes_HEAP
                    wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Fourth_Contact_Attempt_Type_HEAP
                    wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                Else
                    wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
                    wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                    wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Fourth_Contact_Attempt_Notes_HEAP
                    wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Fourth_Contact_Attempt_Type_HEAP
                    wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                    wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                End If
                ElseIf Me.Third_Contact_Attempt_Type_HEAP <> "" Then
                    If Me.Schedule_Date_HEAP = "" Then
                        wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
                        wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                        wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Third_Contact_Attempt_Notes_HEAP
                        wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Third_Contact_Attempt_Type_HEAP
                        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    Else
                        wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
                        wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                        wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Third_Contact_Attempt_Notes_HEAP
                        wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Third_Contact_Attempt_Type_HEAP
                        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                    End If
                    ElseIf Me.Second_Contact_Attempt_Type_HEAP <> "" And Me.Schedule_Date_HEAP = "" Then
                        If Me.Schedule_Date_HEAP = "" Then
                            wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
                            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Second_Contact_Attempt_Notes_HEAP
                            wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Second_Contact_Attempt_Type_HEAP
                            wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        Else
                            wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
                            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Second_Contact_Attempt_Notes_HEAP
                            wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Second_Contact_Attempt_Type_HEAP
                            wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                            wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                        End If
                        ElseIf Me.First_Contact_Attempt_Type_HEAP <> "" Then
                            If Me.Schedule_Date_HEAP = "" Then
                                wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
                                wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                                wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                                wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.First_Contact_Attempt_Notes_HEAP
                                wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.First_Contact_Attempt_Type_HEAP
                                wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            Else
                                wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
                                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                                wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.First_Contact_Attempt_Notes_HEAP
                                wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.First_Contact_Attempt_Type_HEAP
                                wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP).NumberFormat = "@"
                                wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                            End If
                        
        End If
           
    End If
Next x
Call Clear_HEAP_Click
End Sub

Private Sub UserForm_Activate()

Set wsDb = Worksheets("Enrollments")

'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "RECEIVED AT VENDOR" Or _
    wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "FIRST CONTACT" Or _
    wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING" Then
        'push data from database to form
        'HEAP Scheduling
        With Enrollment_Listbox
            .AddItem wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
        End With
    
    End If
Next x

With First_Contact_Attempt_Type_HEAP
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Second_Contact_Attempt_Type_HEAP
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Third_Contact_Attempt_Type_HEAP
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Fourth_Contact_Attempt_Type_HEAP
    .AddItem ""
    .AddItem "EMAIL"
    .AddItem "MAIL"
    .AddItem "LEFT MESSAGE"
    .AddItem "NO ANSWER"
    .AddItem "VOICE MAIL"
    .AddItem "TEXT MESSAGE"
'EMAIL; MAIL; LEFT MESSAGE; NO ANSWER; VOICE MAIL; TEXT MESSAGE
End With

With Fifth_Contact_Attempt_Type_HEAP
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


