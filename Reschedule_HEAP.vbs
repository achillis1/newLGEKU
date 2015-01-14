VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Reschedule_HEAP 
   Caption         =   "Reschedule_HEAP"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15240
   OleObjectBlob   =   "Reschedule_HEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Reschedule_HEAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Cancel_Enrollment_HEAP_Click()

Set wsDb = Worksheets("Enrollments")
Set wsContacts = Worksheets("Contacts")

If MsgBox("Cancelation Requires Management Approval, Has Approval Been Granted?", vbYesNo) = vbYes Then
 
    

    EID = Me.Enrollment_ID_HEAP
    'last row database
    wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row


    'Verify that the values have been added to the Fields

    If Me.Contact_Attempt_Notes_HEAP = "" Or Me.Contact_Attempt_Type_HEAP = "" Then
        MsgBox ("Please fill in the Type and Notes of the Attempt")
        Exit Sub
    End If

    For x = 11 To wsDblr
        If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
'Last Modified Date
            wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Contact_Attempt_Notes_HEAP
'CANCELLED Date Set
            wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "CANCELLED"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                 
        End If
    Next x

    'Append the new Contact to the Contact tab
    wsClr = wsContacts.Cells(Rows.Count, NexantContacts.Contact_ID).End(xlUp).row
    wsContacts.Cells(wsClr + 1, NexantContacts.Enrollment_ID_HEAP) = Me.Enrollment_ID_HEAP
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Number) = Me.Contact_Attempt_Number_HEAP
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Type) = Me.Contact_Attempt_Type_HEAP
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Notes) = Me.Contact_Attempt_Notes_HEAP
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_DateTime).NumberFormat = "@"
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_DateTime) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
    wsContacts.Cells(wsClr + 1, NexantContacts.Contact_ID) = wsContacts.Cells(wsClr, NexantContacts.Contact_ID).Value + 1

    'Clear Results
    MsgBox "Form has been saved"
    Call Clear_HEAP_Click
    MsgBox ("Project Has Been Cancelled")
Else
    Exit Sub
End If
End Sub

Private Sub Previous_Contact_Attempt_Number_HEAP_Change()

    If Previous_Contact_Attempt_Number_HEAP.Value = "" Then
        Exit Sub
    End If
    
    If CInt(Previous_Contact_Attempt_Number_HEAP.Value) <= attemptnum Then
        ir = CInt(Previous_Contact_Attempt_Number_HEAP.Value)
        
        Previous_Contact_Attempt_Date_HEAP.Text = adate(ir - 1)
        Previous_Contact_Attempt_Type_HEAP.Text = atype(ir - 1)
        Previous_Contact_Attempt_Notes_HEAP.Text = anote(ir - 1)
        
    End If
End Sub


Private Sub Schedule_Date_HEAP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Schedule_Date_HEAP) = 8 And IsNumeric(Schedule_Date_HEAP) = True Or Schedule_Date_HEAP = "" Then
Schedule_Date_HEAP.BackColor = &H80000005

Else

Schedule_Date_HEAP.BackColor = &HFF&
MsgBox ("Schedule_Date_HEAP is Formatted Incorrectly")
Cancel = True

End If
End Sub

Private Sub Schedule_Time_HEAP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Schedule_Time_HEAP) = 6 And IsNumeric(Schedule_Time_HEAP) = True Or Schedule_Time_HEAP = "" Then
Schedule_Time_HEAP.BackColor = &H80000005

Else

Schedule_Time_HEAP.BackColor = &HFF&
MsgBox ("Schedule_Time_HEAP is Formatted Incorrectly")
Cancel = True

End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub Cancel_HEAP_Click()
Unload Me
frmProcessing.Show vbModeless
End Sub

Private Sub Clear_HEAP_Click()
'Clear Fields
        Enrollment_Listbox.Clear
        Call formreset
        MsgBox "Form Cleared"

Call UserForm_Activate
End Sub

Private Sub formreset()
Me.Enrollment_ID_HEAP = ""
        Me.Contact_Attempt_Number_HEAP = ""
        Me.Contact_Attempt_Notes_HEAP = ""
        'Me.Contact_Attempt_Type_HEAP.Clear
        Me.Previous_Contact_Attempt_Number_HEAP = ""
        Me.Previous_Contact_Attempt_Date_HEAP = ""
        Me.Previous_Contact_Attempt_Type_HEAP = ""
        Me.Previous_Contact_Attempt_Notes_HEAP = ""
        Me.Schedule_Date_HEAP = ""
        Me.Schedule_Time_HEAP = ""
        Me.Primary_contact_name = ""
        Me.Primary_Contact_Address = ""
        Me.Primary_Contact_Address_City = ""
        Me.Primary_Contact_Address_State = ""
        Me.Primary_Contact_Address_Zip = ""
        Me.Primary_Contact_Email = ""
        Me.Primary_Contact_Phone = ""
        Me.Primary_Contact_phone_extension = ""
        Me.Primary_Contact_mobile_phone = ""
        Me.Contact_Name = ""
        Me.Mailing_Street_Address = ""
        Me.Mailing_City = ""
        Me.Mailing_State = ""
        Me.Mailing_Zipcode = ""
        Me.Customer_Email = ""
        Me.Customer_Home_Phone = ""
        Me.Customer_mobile_phone = ""
        Me.Reason_for_audit = ""
End Sub
Private Sub Enrollment_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Set wsDb = Worksheets("Enrollments")
    Set wsContacts = Worksheets("Contacts")
    
    Call formreset
    
    EID = Enrollment_Listbox.Value
    
    'last row database
    wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row
    Me.Enrollment_ID_HEAP = EID
    
    'Find latest contact attempt from Contacts tab using for loop from the bottom up
    wsClr = wsContacts.Cells(Rows.Count, NexantContacts.Enrollment_ID_HEAP).End(xlUp).row
    'Set Contact attempt number in case there haven't been any prior Contacts
    Me.Contact_Attempt_Number_HEAP = 1
    
    For j = 0 To wsClr - 1
        If wsContacts.Cells(wsClr - j, NexantContacts.Enrollment_ID_HEAP) = EID And wsContacts.Cells(wsClr - j, NexantContacts.HEAP_Contact_Attempt_Number) <> "" Then
            Me.Contact_Attempt_Number_HEAP = wsContacts.Cells(wsClr - j, NexantContacts.HEAP_Contact_Attempt_Number).Value + 1
            j = wsClr - 1
        End If
    Next j
    
        'Retrive values from Database

    For x = 11 To wsDblr
        If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
            'pull data from database to form
            'HEAP Scheduling
    
            Me.Schedule_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP)
            Me.Schedule_Time_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP)
            Me.Primary_contact_name = wsDb.Cells(x, NexantEnrollments.Primary_contact_name)
            Me.Primary_Contact_Address = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Address)
            Me.Primary_Contact_Address_City = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Address_City)
            Me.Primary_Contact_Address_State = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Address_State)
            Me.Primary_Contact_Address_Zip = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Address_Zip)
            Me.Primary_Contact_Email = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Email)
            Me.Primary_Contact_Phone = wsDb.Cells(x, NexantEnrollments.Primary_Contact_Phone)
            Me.Primary_Contact_phone_extension = wsDb.Cells(x, NexantEnrollments.Primary_Contact_phone_extension)
            Me.Primary_Contact_mobile_phone = wsDb.Cells(x, NexantEnrollments.Primary_Contact_mobile_phone)
            Me.Contact_Name = wsDb.Cells(x, NexantEnrollments.Contact_Name)
            Me.Mailing_Street_Address = wsDb.Cells(x, NexantEnrollments.Mailing_Street_Address)
            Me.Mailing_City = wsDb.Cells(x, NexantEnrollments.Mailing_City)
            Me.Mailing_State = wsDb.Cells(x, NexantEnrollments.Mailing_State)
            Me.Mailing_Zipcode = wsDb.Cells(x, NexantEnrollments.Mailing_Zipcode)
            Me.Customer_Email = wsDb.Cells(x, NexantEnrollments.Customer_Email)
            Me.Customer_Home_Phone = wsDb.Cells(x, NexantEnrollments.Customer_Home_Phone)
            Me.Customer_mobile_phone = wsDb.Cells(x, NexantEnrollments.Customer_Home_Phone)
            Me.Reason_for_audit = wsDb.Cells(x, NexantEnrollments.Reason_for_audit)
                 
            
        End If
    Next x
    
    'Ding
    Call updatepreviouscontactattemptnumber

End Sub


Private Sub updatepreviouscontactattemptnumber()

Set wsDb = Worksheets("Enrollments")
Set wsContacts = Worksheets("Contacts")
    
    attemptnum = 0
    ReDim anum(0)
    ReDim adate(0)
    ReDim arow(0)
    If Enrollment_Listbox.ListIndex <> -1 Then
        Previous_Contact_Attempt_Number_HEAP.Clear
        EID = Enrollment_Listbox.Value
        lastcontactrow = wsContacts.Cells(Rows.Count, NexantContacts.Enrollment_ID_HEAP).End(xlUp).row
        For i = 2 To lastcontactrow
            cursorEID = wsContacts.Cells(i, NexantContacts.Enrollment_ID_HEAP).Value
            If cursorEID = EID Then
                
                ReDim Preserve anum(attemptnum)
                ReDim Preserve adate(attemptnum)
                ReDim Preserve arow(attemptnum)
                ReDim Preserve atype(attemptnum)
                ReDim Preserve anote(attemptnum)
                attemptnum = attemptnum + 1

                anum(attemptnum - 1) = attemptnum
                adate(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.HEAP_Contact_DateTime).Value
                anote(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.HEAP_Contact_Attempt_Notes).Value
                atype(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.HEAP_Contact_Attempt_Type).Value
                arow(attemptnum - 1) = i
                wsContacts.Cells(i, NexantContacts.HEAP_Contact_Attempt_Number).Value = attemptnum
                Previous_Contact_Attempt_Number_HEAP.AddItem (attemptnum)
            End If
        Next i
    End If
End Sub

''' Add done'''

Private Sub Save_HEAP_Click()

Set wsDb = Worksheets("Enrollments")
Set wsContacts = Worksheets("Contacts")

EID = Me.Enrollment_ID_HEAP
'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row


'Verify that the values have been added to the Fields
If Me.Schedule_Date_HEAP <> "" And Me.Schedule_Time_HEAP = "" Then
    MsgBox ("Please fill in the Schedule Time")
    Exit Sub
ElseIf Me.Schedule_Date_HEAP = "" And Me.Schedule_Time_HEAP <> "" Then
    MsgBox ("Please fill in the Schedule Date")
    Exit Sub
ElseIf Me.Contact_Attempt_Notes_HEAP = "" Or Me.Contact_Attempt_Type_HEAP = "" Then
    MsgBox ("Please fill in the Type and Notes of the Attempt")
    Exit Sub
End If

For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
         'wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP) = Me.Contact_Attempt_Notes_HEAP
         'wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP) = Me.Contact_Attempt_Type_HEAP
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
         wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Contact_Attempt_Notes_HEAP
         
        'Set First Contact Date/Time
        If Me.Contact_Attempt_Number_HEAP = 1 Then
            wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
        End If
        'Set Status to Pending or Scheduled
        ''''''''''Need a Pending Date Set and Pending Date Interfaced HEAP
        If Me.Schedule_Date_HEAP = "" Then
            wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP) = Me.Schedule_Date_HEAP
            wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP) = Me.Schedule_Time_HEAP
            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = ""
            
         Else
            wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP) = Me.Schedule_Date_HEAP
            wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP) = Me.Schedule_Time_HEAP
         End If
         
    End If
Next x

'Append the new Contact to the Contact tab
wsClr = wsContacts.Cells(Rows.Count, NexantContacts.Contact_ID).End(xlUp).row
wsContacts.Cells(wsClr + 1, NexantContacts.Enrollment_ID_HEAP) = Me.Enrollment_ID_HEAP
wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Number) = Me.Contact_Attempt_Number_HEAP
wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Type) = Me.Contact_Attempt_Type_HEAP
wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Attempt_Notes) = Me.Contact_Attempt_Notes_HEAP
wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_DateTime).NumberFormat = "@"
wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_DateTime) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
If wsClr = 1 Then
wsContacts.Cells(wsClr + 1, NexantContacts.Contact_ID) = 10000
Else
wsContacts.Cells(wsClr + 1, NexantContacts.Contact_ID) = wsContacts.Cells(wsClr, NexantContacts.Contact_ID).Value + 1
End If
'"Discussion, No Scheduled Appt" and "Discussion, Scheduled Appt"
If Me.Schedule_Date_HEAP = "" And Me.Contact_Attempt_Type_HEAP = "Left Message" Then
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Response) = "Discussion, No Scheduled Appt"
ElseIf Me.Schedule_Date_HEAP <> "" Then
    wsContacts.Cells(wsClr + 1, NexantContacts.HEAP_Contact_Response).Value = "Discussion, Scheduled Appt"
End If

'Clear Results
MsgBox "Form has been saved"
Call Clear_HEAP_Click

End Sub

Private Sub UserForm_Activate()

    Set wsDb = Worksheets("Enrollments")
    
    'last row database
    wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row
    
    'find row in Database for Enrollment ID
    For x = 11 To wsDblr
        If wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED" Then  'Or _
        wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "FIRST CONTACT" Or _
        wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "PENDING" Then
            'push data from database to form
            'HEAP Scheduling
            With Enrollment_Listbox
                .AddItem wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
            End With
        
        End If
    Next x
    
    Me.Contact_Attempt_Type_HEAP.Clear
    
    
    With Me.Contact_Attempt_Type_HEAP
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



