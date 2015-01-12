VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} List_Contact_Attempts_ROSA 
   Caption         =   "List_Contact_Attempts_ROSA"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15465
   OleObjectBlob   =   "List_Contact_Attempts_ROSA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "List_Contact_Attempts_ROSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Previous_Contact_Attempt_Number_ROSA_Change()

    If Previous_Contact_Attempt_Number_ROSA.Value = "" Then
        Exit Sub
    End If
    
    If CInt(Previous_Contact_Attempt_Number_ROSA.Value) <= attemptnum Then
        ir = CInt(Previous_Contact_Attempt_Number_ROSA.Value)
        
        Previous_Contact_Attempt_Date_ROSA.Text = adate(ir - 1)
        Previous_Contact_Attempt_Type_ROSA.Text = atype(ir - 1)
        Previous_Contact_Attempt_Notes_ROSA.Text = anote(ir - 1)
        
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub Cancel_ROSA_Click()
Unload Me
frmProcessing.Show vbModeless
End Sub

Private Sub Clear_ROSA_Click()
'Clear Fields
        Enrollment_Listbox.Clear
        Call formreset
        MsgBox "Form Cleared"

Call UserForm_Activate
End Sub

Private Sub formreset()
Me.Enrollment_ID_ROSA = ""
        Me.Previous_Contact_Attempt_Number_ROSA = ""
        Me.Previous_Contact_Attempt_Date_ROSA = ""
        Me.Previous_Contact_Attempt_Type_ROSA = ""
        Me.Previous_Contact_Attempt_Notes_ROSA = ""
        Me.Schedule_Date_ROSA = ""
        Me.Schedule_Time_ROSA = ""
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
    wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row
    
    '''''Find latest contact attempt from Contacts tab using for loop from the bottom up
    '''''Reference the  me.contact_attempt_number = contact number +1
    
    
    'Retrive values from Database
    Me.Enrollment_ID_ROSA = EID
    For x = 11 To wsDblr
        If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = EID Then
            'pull data from database to form
            'ROSA Scheduling
    
            Me.Schedule_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Schedule_Date_ROSA)
            Me.Schedule_Time_ROSA = wsDb.Cells(x, NexantEnrollments.Schedule_Time_ROSA)
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
        Previous_Contact_Attempt_Number_ROSA.Clear
        EID = Enrollment_Listbox.Value
        lastcontactrow = wsContacts.Cells(Rows.Count, NexantContacts.Enrollment_ID_ROSA).End(xlUp).row
        For i = 2 To lastcontactrow
            cursorEID = wsContacts.Cells(i, NexantContacts.Enrollment_ID_ROSA).Value
            If cursorEID = EID Then
                
                ReDim Preserve anum(attemptnum)
                ReDim Preserve adate(attemptnum)
                ReDim Preserve arow(attemptnum)
                ReDim Preserve atype(attemptnum)
                ReDim Preserve anote(attemptnum)
                attemptnum = attemptnum + 1

                anum(attemptnum - 1) = attemptnum
                adate(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.ROSA_Contact_DateTime).Value
                anote(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.ROSA_Contact_Attempt_Notes).Value
                atype(attemptnum - 1) = wsContacts.Cells(i, NexantContacts.ROSA_Contact_Attempt_Type).Value
                arow(attemptnum - 1) = i
                wsContacts.Cells(i, NexantContacts.ROSA_Contact_Attempt_Number).Value = attemptnum
                Previous_Contact_Attempt_Number_ROSA.AddItem (attemptnum)
            End If
        Next i
    End If
End Sub

''' Add done'''

Private Sub Save_ROSA_Click()

Set wsDb = Worksheets("Enrollments")
Set wsContacts = Worksheets("Contacts")
'Enrollment_Listbox.Value = EID
EID = Me.Enrollment_ID_ROSA
'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row


'Verify that the values have been added to the Fields

For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = EID Then
         wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA) = Me.Contact_Attempt_Notes_ROSA
         wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA) = Me.Contact_Attempt_Type_ROSA
         wsDb.Cells(x, NexantEnrollments.Schedule_Date_ROSA) = Me.Schedule_Date_ROSA
         wsDb.Cells(x, NexantEnrollments.Schedule_Time_ROSA) = Me.Schedule_Time_ROSA
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
         wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
         wsDb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Contact_Attempt_Notes_ROSA
         
        'Set First Contact Date/Time
        If Me.Contact_Attempt_Number_ROSA = 1 Then
            wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
        End If
        'Set Status to Pending or Scheduled
        If Me.Schedule_Date_ROSA = "" Then
            wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING"
            wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
            wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
         Else
            wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED"
            wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
            wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA).NumberFormat = "@"
            wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
         End If
    '''''Append the new Contact to the Contact tab
    '''''Find the last row +1 and set values from form to tab
              
    End If
Next x
Call Clear_ROSA_Click
End Sub

Private Sub UserForm_Activate()

    Set wsDb = Worksheets("Enrollments")
    
    'last row database
    wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row
    
    'find row in Database for Enrollment ID
    For x = 11 To wsDblr
        If wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "RECEIVED AT VENDOR" Or _
        wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "FIRST CONTACT" Or _
        wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "PENDING" Then
            'push data from database to form
            'ROSA Scheduling
            With Enrollment_Listbox
                .AddItem wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA)
            End With
        
        End If
    Next x
    
    Me.Contact_Attempt_Type_ROSA.Clear
    
    
    With Me.Contact_Attempt_Type_ROSA
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
