VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ROSA_Scheduling_Status_Info 
   Caption         =   "ROSA_Scheduling_Status_Info"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12840
   OleObjectBlob   =   "ROSA_Scheduling_Status_Info.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ROSA_Scheduling_Status_Info"
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
Me.Hide
frmAdmin.Show vbModeless
End Sub

Private Sub Frame8_Click()

End Sub

Private Sub Revert_ROSA_Click()
Call UserForm_Initialize

End Sub

Private Sub Save_ROSA_Click()

Set wsDb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_ROSA = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsDb.Cells(Rows.Count, 2).End(xlUp).row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, 2) = EID Then
        'push data from form to Database
        'ROSA Scheduling
        wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA) = Me.Customer_contact_mode_ROSA
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA) = Me.Fifth_Contact_Attempt_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_ROSA) = Me.Fifth_Contact_Attempt_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_ROSA) = Me.Fifth_Contact_Attempt_Type_ROSA
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA) = Me.First_Contact_Attempt_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA) = Me.First_Contact_Attempt_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA) = Me.First_Contact_Attempt_Type_ROSA
        wsDb.Cells(x, NexantEnrollments.Follow_up_Date_ROSA) = Me.Follow_up_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA) = Me.Fourth_Contact_Attempt_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_ROSA) = Me.Fourth_Contact_Attempt_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_ROSA) = Me.Fourth_Contact_Attempt_Type_ROSA
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA) = Me.Second_Contact_Attempt_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_ROSA) = Me.Second_Contact_Attempt_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_ROSA) = Me.Second_Contact_Attempt_Type_ROSA
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA) = Me.Third_Contact_Attempt_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_ROSA) = Me.Third_Contact_Attempt_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_ROSA) = Me.Third_Contact_Attempt_Type_ROSA
        'ROSA Status
        wsDb.Cells(x, NexantEnrollments.CANCELLED_date_interfaced_ROSA) = Me.CANCELLED_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_ROSA) = Me.CANCELLED_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.COMPLETE_date_interfaced_ROSA) = Me.COMPLETE_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_ROSA) = Me.COMPLETE_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA) = Me.FIRST_CONTACT_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA) = Me.FIRST_CONTACT_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_interfaced_ROSA) = Me.ON_HOLD_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_set_ROSA) = Me.ON_HOLD_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_1_date_interfaced_ROSA) = Me.PENDING_1_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_ROSA) = Me.PENDING_1_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_2_date_interfaced_ROSA) = Me.PENDING_2_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_ROSA) = Me.PENDING_2_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_interfaced_ROSA) = Me.PENDING_3_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_ROSA) = Me.PENDING_3_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_4_date_interfaced_ROSA) = Me.PENDING_4_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_ROSA) = Me.PENDING_4_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_5_date_interfaced_ROSA) = Me.PENDING_5_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_ROSA) = Me.PENDING_5_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA) = Me.RECEIVED_AT_VENDOR_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA) = Me.RECEIVED_AT_VENDOR_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_interfaced_ROSA) = Me.SCHEDULED_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA) = Me.SCHEDULED_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_ROSA) = Me.SITE_WORK_COMPLETE_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA) = Me.SITE_WORK_COMPLETE_date_set_ROSA
        wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Me.Status_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Status_ROSA) = Me.Status_ROSA
        wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Me.Status_Time_ROSA
        wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_interfaced_ROSA) = Me.WITHDRAWN_date_interfaced_ROSA
        wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_set_ROSA) = Me.WITHDRAWN_date_set_ROSA
        'ROSA Info
        wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_ROSA) = Me.Air_Leakage_Rating_ROSA
        wsDb.Cells(x, NexantEnrollments.Auditor_Notes_ROSA) = Me.Auditor_Notes_ROSA
        wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_ROSA) = Me.Blower_door_post_test_ROSA
        wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_ROSA) = Me.Blower_door_pre_test_ROSA
        wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_ROSA) = Me.Building_occupancy_count_ROSA
        wsDb.Cells(x, NexantEnrollments.Business_Partner_Number_ROSA) = Me.Business_Partner_Number_ROSA
        wsDb.Cells(x, NexantEnrollments.Comments_ROSA) = Me.Comments_ROSA
        wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_ROSA) = Me.Dog_or_Cat_Flag_ROSA
        wsDb.Cells(x, NexantEnrollments.FILE_NAME_ROSA) = Me.FILE_NAME_ROSA
        wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA) = Me.First_and_last_name_of_main_Auditor_ROSA
        wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_ROSA) = Me.Number_of_Auditors_ROSA
        wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_ROSA) = Me.Number_of_stories_above_grade_ROSA
        wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_ROSA) = Me.Occupancy_frequency_ROSA
        wsDb.Cells(x, NexantEnrollments.Ownership_Type_ROSA) = Me.Ownership_Type_ROSA
        wsDb.Cells(x, NexantEnrollments.Schedule_Date_ROSA) = Me.Schedule_Date_ROSA
        wsDb.Cells(x, NexantEnrollments.Schedule_Time_ROSA) = Me.Schedule_Time_ROSA
        wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_ROSA) = Me.Total_conditioned_square_footage_ROSA
        wsDb.Cells(x, NexantEnrollments.WO_Number_ROSA) = Me.WO_Number_ROSA
'Time stamp on Last updated
        wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment).NumberFormat = "@"
        wsDb.Cells(x, NexantEnrollments.Last_Modified_Date_Enrollment) = Format(LocalTimeToET(Now()), "YYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
        
        Exit Sub
        
    Else
        If x = wsDblr Then
            MsgBox ("The Enrollment ID is not found in the Database")
            Exit Sub
        End If
    End If
Next x


End Sub

Private Sub UserForm_Initialize()

Set wsDb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_ROSA = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsDb.Cells(Rows.Count, 2).End(xlUp).row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, 2) = EID Then
        'push data from database to form
        Me.Enrollment_ID_ROSA = wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA)
        'ROSA Scheduling
        Me.Customer_contact_mode_ROSA = wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_ROSA)
        Me.Fifth_Contact_Attempt_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_ROSA)
        Me.Fifth_Contact_Attempt_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_ROSA)
        Me.Fifth_Contact_Attempt_Type_ROSA = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_ROSA)
        Me.First_Contact_Attempt_Date_ROSA = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_ROSA)
        Me.First_Contact_Attempt_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_ROSA)
        Me.First_Contact_Attempt_Type_ROSA = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_ROSA)
        Me.Follow_up_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Follow_up_Date_ROSA)
        Me.Fourth_Contact_Attempt_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_ROSA)
        Me.Fourth_Contact_Attempt_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_ROSA)
        Me.Fourth_Contact_Attempt_Type_ROSA = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_ROSA)
        Me.Second_Contact_Attempt_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_ROSA)
        Me.Second_Contact_Attempt_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_ROSA)
        Me.Second_Contact_Attempt_Type_ROSA = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_ROSA)
        Me.Third_Contact_Attempt_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_ROSA)
        Me.Third_Contact_Attempt_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_ROSA)
        Me.Third_Contact_Attempt_Type_ROSA = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_ROSA)
        'ROSA Status
        Me.CANCELLED_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.CANCELLED_date_interfaced_ROSA)
        Me.CANCELLED_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_ROSA)
        Me.COMPLETE_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.COMPLETE_date_interfaced_ROSA)
        Me.COMPLETE_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_ROSA)
        Me.FIRST_CONTACT_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA)
        Me.FIRST_CONTACT_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_ROSA)
        Me.ON_HOLD_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_interfaced_ROSA)
        Me.ON_HOLD_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_set_ROSA)
        Me.PENDING_1_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_1_date_interfaced_ROSA)
        Me.PENDING_1_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_ROSA)
        Me.PENDING_2_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_2_date_interfaced_ROSA)
        Me.PENDING_2_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_ROSA)
        Me.PENDING_3_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_3_date_interfaced_ROSA)
        Me.PENDING_3_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_ROSA)
        Me.PENDING_4_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_4_date_interfaced_ROSA)
        Me.PENDING_4_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_ROSA)
        Me.PENDING_5_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_5_date_interfaced_ROSA)
        Me.PENDING_5_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_ROSA)
        Me.RECEIVED_AT_VENDOR_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA)
        Me.RECEIVED_AT_VENDOR_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA)
        Me.SCHEDULED_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_interfaced_ROSA)
        Me.SCHEDULED_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_ROSA)
        Me.SITE_WORK_COMPLETE_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_ROSA)
        Me.SITE_WORK_COMPLETE_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA)
        Me.Status_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA)
        Me.Status_ROSA = wsDb.Cells(x, NexantEnrollments.Status_ROSA)
        Me.Status_Time_ROSA = wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA)
        Me.WITHDRAWN_date_interfaced_ROSA = wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_interfaced_ROSA)
        Me.WITHDRAWN_date_set_ROSA = wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_set_ROSA)
        'ROSA Info
        Me.Air_Leakage_Rating_ROSA = wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_ROSA)
        Me.Auditor_Notes_ROSA = wsDb.Cells(x, NexantEnrollments.Auditor_Notes_ROSA)
        Me.Blower_door_post_test_ROSA = wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_ROSA)
        Me.Blower_door_pre_test_ROSA = wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_ROSA)
        Me.Building_occupancy_count_ROSA = wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_ROSA)
        Me.Business_Partner_Number_ROSA = wsDb.Cells(x, NexantEnrollments.Business_Partner_Number_ROSA)
        Me.Comments_ROSA = wsDb.Cells(x, NexantEnrollments.Comments_ROSA)
        Me.Dog_or_Cat_Flag_ROSA = wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_ROSA)
        Me.FILE_NAME_ROSA = wsDb.Cells(x, NexantEnrollments.FILE_NAME_ROSA)
        Me.First_and_last_name_of_main_Auditor_ROSA = wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA)
        Me.Number_of_Auditors_ROSA = wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_ROSA)
        Me.Number_of_stories_above_grade_ROSA = wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_ROSA)
        Me.Occupancy_frequency_ROSA = wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_ROSA)
        Me.Ownership_Type_ROSA = wsDb.Cells(x, NexantEnrollments.Ownership_Type_ROSA)
        Me.Schedule_Date_ROSA = wsDb.Cells(x, NexantEnrollments.Schedule_Date_ROSA)
        Me.Schedule_Time_ROSA = wsDb.Cells(x, NexantEnrollments.Schedule_Time_ROSA)
        Me.Total_conditioned_square_footage_ROSA = wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_ROSA)
        Me.WO_Number_ROSA = wsDb.Cells(x, NexantEnrollments.WO_Number_ROSA)
        
        Exit Sub
        
    Else
        If x = wsDblr Then
            MsgBox ("The Enrollment ID is not found in the Database")
            Exit Sub
        End If
    End If
Next x

End Sub


