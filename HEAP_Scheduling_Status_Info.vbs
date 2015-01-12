VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HEAP_Scheduling_Status_Info 
   Caption         =   "HEAP_Scheduling_Status_Info"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12615
   OleObjectBlob   =   "HEAP_Scheduling_Status_Info.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HEAP_Scheduling_Status_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Set wsDb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_HEAP = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsDb.Cells(Rows.Count, 3).End(xlUp).row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, 3) = EID Then
        'push data from database to form
        'HEAP Scheduling

        Me.Enrollment_ID_HEAP = wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
        Me.Customer_contact_mode_HEAP = wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP)
        Me.Fifth_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP)
        Me.Fifth_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP)
        Me.Fifth_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP)
        Me.First_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP)
        Me.First_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP)
        Me.First_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP)
        Me.Follow_up_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Follow_up_Date_HEAP)
        Me.Fourth_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP)
        Me.Fourth_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP)
        Me.Fourth_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP)
        Me.Second_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP)
        Me.Second_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP)
        Me.Second_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP)
        Me.Third_Contact_Attempt_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP)
        Me.Third_Contact_Attempt_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP)
        Me.Third_Contact_Attempt_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP)
        Me.CANCELLED_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.CANCELLED_date_interfaced_HEAP)
        Me.CANCELLED_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_HEAP)
        Me.COMPLETE_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.COMPLETE_date_interfaced_HEAP)
        Me.COMPLETE_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_HEAP)
        Me.FIRST_CONTACT_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP)
        Me.FIRST_CONTACT_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP)
        Me.ON_HOLD_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_interfaced_HEAP)
        Me.ON_HOLD_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_set_HEAP)
        Me.PENDING_1_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_1_date_interfaced_HEAP)
        Me.PENDING_1_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_HEAP)
        Me.PENDING_2_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_2_date_interfaced_HEAP)
        Me.PENDING_2_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_HEAP)
        Me.PENDING_3_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_3_date_interfaced_HEAP)
        Me.PENDING_3_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_HEAP)
        Me.PENDING_4_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_4_date_interfaced_HEAP)
        Me.PENDING_4_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_HEAP)
        Me.PENDING_5_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_5_date_interfaced_HEAP)
        Me.PENDING_5_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_HEAP)
        Me.RECEIVED_AT_VENDOR_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP)
        Me.RECEIVED_AT_VENDOR_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_HEAP)
        Me.SCHEDULED_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_interfaced_HEAP)
        Me.SCHEDULED_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP)
        Me.SITE_WORK_COMPLETE_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_HEAP)
        Me.SITE_WORK_COMPLETE_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP)
        Me.Status_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP)
        Me.Status_HEAP = wsDb.Cells(x, NexantEnrollments.Status_HEAP)
        Me.Status_Time_HEAP = wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP)
        Me.WITHDRAWN_date_interfaced_HEAP = wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_interfaced_HEAP)
        Me.WITHDRAWN_date_set_HEAP = wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_set_HEAP)
        Me.Air_Leakage_Rating_HEAP = wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_HEAP)
        Me.Auditor_Notes_HEAP = wsDb.Cells(x, NexantEnrollments.Auditor_Notes_HEAP)
        Me.Blower_door_post_test_HEAP = wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_HEAP)
        Me.Blower_door_pre_test_HEAP = wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_HEAP)
        Me.Building_occupancy_count_HEAP = wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_HEAP)
        Me.Business_Partner_Number_HEAP = wsDb.Cells(x, NexantEnrollments.Business_Partner_Number_HEAP)
        Me.Comments_HEAP = wsDb.Cells(x, NexantEnrollments.Comments_HEAP)
        Me.Dog_or_Cat_Flag_HEAP = wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_HEAP)
        Me.FILE_NAME_HEAP = wsDb.Cells(x, NexantEnrollments.FILE_NAME_HEAP)
        Me.First_and_last_name_of_main_Auditor_HEAP = wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP)
        Me.Number_of_Auditors_HEAP = wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_HEAP)
        Me.Number_of_stories_above_grade_HEAP = wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_HEAP)
        Me.Occupancy_frequency_HEAP = wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_HEAP)
        Me.Ownership_Type_HEAP = wsDb.Cells(x, NexantEnrollments.Ownership_Type_HEAP)
        Me.Schedule_Date_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP)
        Me.Schedule_Time_HEAP = wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP)
        Me.Total_conditioned_square_footage_HEAP = wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_HEAP)
        Me.WO_Number_HEAP = wsDb.Cells(x, NexantEnrollments.WO_Number_HEAP)

        Exit Sub
        
    Else
        If x = wsDblr Then
            MsgBox ("The Enrollment ID is not found in the Database")
            Exit Sub
        End If
    End If
Next x
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
Private Sub Cancel_HEAP_Click()
Me.Hide
frmAdmin.Show vbModeless
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub Revert_HEAP_Click()
Call UserForm_Initialize

End Sub

Private Sub Save_HEAP_Click()
Set wsDb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_HEAP = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsDb.Cells(Rows.Count, 3).End(xlUp).row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, 3) = EID Then
        'push data from form to database
        'HEAP Scheduling

        wsDb.Cells(x, NexantEnrollments.Customer_contact_mode_HEAP) = Me.Customer_contact_mode_HEAP
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP) = Me.Fifth_Contact_Attempt_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP) = Me.Fifth_Contact_Attempt_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP) = Me.Fifth_Contact_Attempt_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP) = Me.First_Contact_Attempt_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP) = Me.First_Contact_Attempt_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP) = Me.First_Contact_Attempt_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.Follow_up_Date_HEAP) = Me.Follow_up_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP) = Me.Fourth_Contact_Attempt_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP) = Me.Fourth_Contact_Attempt_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP) = Me.Fourth_Contact_Attempt_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP) = Me.Second_Contact_Attempt_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP) = Me.Second_Contact_Attempt_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP) = Me.Second_Contact_Attempt_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP) = Me.Third_Contact_Attempt_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP) = Me.Third_Contact_Attempt_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP) = Me.Third_Contact_Attempt_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.CANCELLED_date_interfaced_HEAP) = Me.CANCELLED_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.CANCELLED_date_set_HEAP) = Me.CANCELLED_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.COMPLETE_date_interfaced_HEAP) = Me.COMPLETE_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_HEAP) = Me.COMPLETE_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP) = Me.FIRST_CONTACT_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_set_HEAP) = Me.FIRST_CONTACT_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_interfaced_HEAP) = Me.ON_HOLD_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.ON_HOLD_date_set_HEAP) = Me.ON_HOLD_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_1_date_interfaced_HEAP) = Me.PENDING_1_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_1_date_set_HEAP) = Me.PENDING_1_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_2_date_interfaced_HEAP) = Me.PENDING_2_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_2_date_set_HEAP) = Me.PENDING_2_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_interfaced_HEAP) = Me.PENDING_3_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_3_date_set_HEAP) = Me.PENDING_3_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_4_date_interfaced_HEAP) = Me.PENDING_4_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_4_date_set_HEAP) = Me.PENDING_4_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_5_date_interfaced_HEAP) = Me.PENDING_5_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.PENDING_5_date_set_HEAP) = Me.PENDING_5_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP) = Me.RECEIVED_AT_VENDOR_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_HEAP) = Me.RECEIVED_AT_VENDOR_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_interfaced_HEAP) = Me.SCHEDULED_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.SCHEDULED_date_set_HEAP) = Me.SCHEDULED_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_HEAP) = Me.SITE_WORK_COMPLETE_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP) = Me.SITE_WORK_COMPLETE_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Me.Status_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Status_HEAP) = Me.Status_HEAP
        wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Me.Status_Time_HEAP
        wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_interfaced_HEAP) = Me.WITHDRAWN_date_interfaced_HEAP
        wsDb.Cells(x, NexantEnrollments.WITHDRAWN_date_set_HEAP) = Me.WITHDRAWN_date_set_HEAP
        wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_HEAP) = Me.Air_Leakage_Rating_HEAP
        wsDb.Cells(x, NexantEnrollments.Auditor_Notes_HEAP) = Me.Auditor_Notes_HEAP
        wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_HEAP) = Me.Blower_door_post_test_HEAP
        wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_HEAP) = Me.Blower_door_pre_test_HEAP
        wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_HEAP) = Me.Building_occupancy_count_HEAP
        wsDb.Cells(x, NexantEnrollments.Business_Partner_Number_HEAP) = Me.Business_Partner_Number_HEAP
        wsDb.Cells(x, NexantEnrollments.Comments_HEAP) = Me.Comments_HEAP
        wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_HEAP) = Me.Dog_or_Cat_Flag_HEAP
        wsDb.Cells(x, NexantEnrollments.FILE_NAME_HEAP) = Me.FILE_NAME_HEAP
        wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP) = Me.First_and_last_name_of_main_Auditor_HEAP
        wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_HEAP) = Me.Number_of_Auditors_HEAP
        wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_HEAP) = Me.Number_of_stories_above_grade_HEAP
        wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_HEAP) = Me.Occupancy_frequency_HEAP
        wsDb.Cells(x, NexantEnrollments.Ownership_Type_HEAP) = Me.Ownership_Type_HEAP
        wsDb.Cells(x, NexantEnrollments.Schedule_Date_HEAP) = Me.Schedule_Date_HEAP
        wsDb.Cells(x, NexantEnrollments.Schedule_Time_HEAP) = Me.Schedule_Time_HEAP
        wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_HEAP) = Me.Total_conditioned_square_footage_HEAP
        wsDb.Cells(x, NexantEnrollments.WO_Number_HEAP) = Me.WO_Number_HEAP
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


End Sub


