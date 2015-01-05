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
Public EID As String
Private Sub Cancel_ROSA_Click()
    Unload Me
    frmAdmin.Show
End Sub



Private Sub Revert_ROSA_Click()
Call UserForm_Initialize

End Sub

Private Sub Save_ROSA_Click()

Set wsdb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_ROSA = EID

EID = currentEnrollment

'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, 2) = EID Then
        'push data from form to Database
        'ROSA Scheduling
        wsdb.Cells(x, 351) = Me.Customer_contact_mode_ROSA
        wsdb.Cells(x, 352) = Me.Fifth_Contact_Attempt_Date_ROSA
        wsdb.Cells(x, 353) = Me.Fifth_Contact_Attempt_Notes_ROSA
        wsdb.Cells(x, 354) = Me.Fifth_Contact_Attempt_Type_ROSA
        wsdb.Cells(x, 355) = Me.First_Contact_Attempt_Date_ROSA
        wsdb.Cells(x, 356) = Me.First_Contact_Attempt_Notes_ROSA
        wsdb.Cells(x, 357) = Me.First_Contact_Attempt_Type_ROSA
        wsdb.Cells(x, 358) = Me.Follow_up_Date_ROSA
        wsdb.Cells(x, 359) = Me.Fourth_Contact_Attempt_Date_ROSA
        wsdb.Cells(x, 360) = Me.Fourth_Contact_Attempt_Notes_ROSA
        wsdb.Cells(x, 361) = Me.Fourth_Contact_Attempt_Type_ROSA
        wsdb.Cells(x, 362) = Me.Schedule_Date_ROSA
        wsdb.Cells(x, 363) = Me.Schedule_Time_ROSA
        wsdb.Cells(x, 364) = Me.Second_Contact_Attempt_Date_ROSA
        wsdb.Cells(x, 365) = Me.Second_Contact_Attempt_Notes_ROSA
        wsdb.Cells(x, 366) = Me.Second_Contact_Attempt_Type_ROSA
        wsdb.Cells(x, 367) = Me.Third_Contact_Attempt_Date_ROSA
        wsdb.Cells(x, 368) = Me.Third_Contact_Attempt_Notes_ROSA
        wsdb.Cells(x, 369) = Me.Third_Contact_Attempt_Type_ROSA

        'ROSA Status
        wsdb.Cells(x, 370) = Me.CANCELLED_date_interfaced_ROSA
        wsdb.Cells(x, 371) = Me.CANCELLED_date_set_ROSA
        wsdb.Cells(x, 372) = Me.COMPLETE_date_interfaced_ROSA
        wsdb.Cells(x, 373) = Me.COMPLETE_date_set_ROSA
        wsdb.Cells(x, 374) = Me.FIRST_CONTACT_date_interfaced_ROSA
        wsdb.Cells(x, 375) = Me.FIRST_CONTACT_date_set_ROSA
        wsdb.Cells(x, 376) = Me.ON_HOLD_date_interfaced_ROSA
        wsdb.Cells(x, 377) = Me.ON_HOLD_date_set_ROSA
        wsdb.Cells(x, 378) = Me.PENDING_1_date_interfaced_ROSA
        wsdb.Cells(x, 379) = Me.PENDING_1_date_set_ROSA
        wsdb.Cells(x, 380) = Me.PENDING_2_date_interfaced_ROSA
        wsdb.Cells(x, 381) = Me.PENDING_2_date_set_ROSA
        wsdb.Cells(x, 382) = Me.PENDING_3_date_interfaced_ROSA
        wsdb.Cells(x, 383) = Me.PENDING_3_date_set_ROSA
        wsdb.Cells(x, 384) = Me.PENDING_4_date_interfaced_ROSA
        wsdb.Cells(x, 385) = Me.PENDING_4_date_set_ROSA
        wsdb.Cells(x, 386) = Me.PENDING_5_date_interfaced_ROSA
        wsdb.Cells(x, 387) = Me.PENDING_5_date_set_ROSA
        wsdb.Cells(x, 388) = Me.RECEIVED_AT_VENDOR_date_interfaced_ROSA
        wsdb.Cells(x, 389) = Me.RECEIVED_AT_VENDOR_date_set_ROSA
        wsdb.Cells(x, 390) = Me.SCHEDULED_date_interfaced_ROSA
        wsdb.Cells(x, 391) = Me.SCHEDULED_date_set_ROSA
        wsdb.Cells(x, 392) = Me.SITE_WORK_COMPLETE_date_interfaced_ROSA
        wsdb.Cells(x, 393) = Me.SITE_WORK_COMPLETE_date_set_ROSA
        wsdb.Cells(x, 394) = Me.Status_Date_ROSA
        wsdb.Cells(x, 395) = Me.Status_ROSA
        wsdb.Cells(x, 396) = Me.Status_Time_ROSA
        wsdb.Cells(x, 397) = Me.WITHDRAWN_date_interfaced_ROSA
        wsdb.Cells(x, 398) = Me.WITHDRAWN_date_set_ROSA

        'ROSA Info
        wsdb.Cells(x, 399) = Me.Air_Leakage_Rating_ROSA
        wsdb.Cells(x, 400) = Me.Auditor_Notes_ROSA
        wsdb.Cells(x, 401) = Me.Blower_door_post_test_ROSA
        wsdb.Cells(x, 402) = Me.Blower_door_pre_test_ROSA
        wsdb.Cells(x, 403) = Me.Building_occupancy_count_ROSA
        wsdb.Cells(x, 404) = Me.Business_Partner_Number_ROSA
        wsdb.Cells(x, 405) = Me.Comments_ROSA
        wsdb.Cells(x, 406) = Me.Dog_or_Cat_Flag_ROSA
        wsdb.Cells(x, 407) = Me.FILE_NAME_ROSA
        wsdb.Cells(x, 408) = Me.First_and_last_name_of_main_Auditor_ROSA
        wsdb.Cells(x, 409) = Me.Number_of_Auditors_ROSA
        wsdb.Cells(x, 410) = Me.Number_of_stories_above_grade_ROSA
        wsdb.Cells(x, 411) = Me.Occupancy_frequency_ROSA
        wsdb.Cells(x, 412) = Me.Ownership_Type_ROSA
        wsdb.Cells(x, 413) = Me.Total_conditioned_square_footage_ROSA
        wsdb.Cells(x, 414) = Me.WO_Number_ROSA

        
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

Set wsdb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_ROSA = EID


EID = currentEnrollment

'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, 2) = EID Then
        'push data from database to form
        'ROSA Scheduling
        Me.Enrollment_ID_ROSA = wsdb.Cells(x, 2)
        Me.Customer_contact_mode_ROSA = wsdb.Cells(x, 351)
        Me.Fifth_Contact_Attempt_Date_ROSA = wsdb.Cells(x, 352)
        Me.Fifth_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, 353)
        Me.Fifth_Contact_Attempt_Type_ROSA = wsdb.Cells(x, 354)
        Me.First_Contact_Attempt_Date_ROSA = wsdb.Cells(x, 355)
        Me.First_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, 356)
        Me.First_Contact_Attempt_Type_ROSA = wsdb.Cells(x, 357)
        Me.Follow_up_Date_ROSA = wsdb.Cells(x, 358)
        Me.Fourth_Contact_Attempt_Date_ROSA = wsdb.Cells(x, 359)
        Me.Fourth_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, 360)
        Me.Fourth_Contact_Attempt_Type_ROSA = wsdb.Cells(x, 361)
        Me.Schedule_Date_ROSA = wsdb.Cells(x, 362)
        Me.Schedule_Time_ROSA = wsdb.Cells(x, 363)
        Me.Second_Contact_Attempt_Date_ROSA = wsdb.Cells(x, 364)
        Me.Second_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, 365)
        Me.Second_Contact_Attempt_Type_ROSA = wsdb.Cells(x, 366)
        Me.Third_Contact_Attempt_Date_ROSA = wsdb.Cells(x, 367)
        Me.Third_Contact_Attempt_Notes_ROSA = wsdb.Cells(x, 368)
        Me.Third_Contact_Attempt_Type_ROSA = wsdb.Cells(x, 369)

        'ROSA Status
        Me.CANCELLED_date_interfaced_ROSA = wsdb.Cells(x, 370)
        Me.CANCELLED_date_set_ROSA = wsdb.Cells(x, 371)
        Me.COMPLETE_date_interfaced_ROSA = wsdb.Cells(x, 372)
        Me.COMPLETE_date_set_ROSA = wsdb.Cells(x, 373)
        Me.FIRST_CONTACT_date_interfaced_ROSA = wsdb.Cells(x, 374)
        Me.FIRST_CONTACT_date_set_ROSA = wsdb.Cells(x, 375)
        Me.ON_HOLD_date_interfaced_ROSA = wsdb.Cells(x, 376)
        Me.ON_HOLD_date_set_ROSA = wsdb.Cells(x, 377)
        Me.PENDING_1_date_interfaced_ROSA = wsdb.Cells(x, 378)
        Me.PENDING_1_date_set_ROSA = wsdb.Cells(x, 379)
        Me.PENDING_2_date_interfaced_ROSA = wsdb.Cells(x, 380)
        Me.PENDING_2_date_set_ROSA = wsdb.Cells(x, 381)
        Me.PENDING_3_date_interfaced_ROSA = wsdb.Cells(x, 382)
        Me.PENDING_3_date_set_ROSA = wsdb.Cells(x, 383)
        Me.PENDING_4_date_interfaced_ROSA = wsdb.Cells(x, 384)
        Me.PENDING_4_date_set_ROSA = wsdb.Cells(x, 385)
        Me.PENDING_5_date_interfaced_ROSA = wsdb.Cells(x, 386)
        Me.PENDING_5_date_set_ROSA = wsdb.Cells(x, 387)
        Me.RECEIVED_AT_VENDOR_date_interfaced_ROSA = wsdb.Cells(x, 388)
        Me.RECEIVED_AT_VENDOR_date_set_ROSA = wsdb.Cells(x, 389)
        Me.SCHEDULED_date_interfaced_ROSA = wsdb.Cells(x, 390)
        Me.SCHEDULED_date_set_ROSA = wsdb.Cells(x, 391)
        Me.SITE_WORK_COMPLETE_date_interfaced_ROSA = wsdb.Cells(x, 392)
        Me.SITE_WORK_COMPLETE_date_set_ROSA = wsdb.Cells(x, 393)
        Me.Status_Date_ROSA = wsdb.Cells(x, 394)
        Me.Status_ROSA = wsdb.Cells(x, 395)
        Me.Status_Time_ROSA = wsdb.Cells(x, 396)
        Me.WITHDRAWN_date_interfaced_ROSA = wsdb.Cells(x, 397)
        Me.WITHDRAWN_date_set_ROSA = wsdb.Cells(x, 398)
        
        'ROSA Info
        Me.Air_Leakage_Rating_ROSA = wsdb.Cells(x, 399)
        Me.Auditor_Notes_ROSA = wsdb.Cells(x, 400)
        Me.Blower_door_post_test_ROSA = wsdb.Cells(x, 401)
        Me.Blower_door_pre_test_ROSA = wsdb.Cells(x, 402)
        Me.Building_occupancy_count_ROSA = wsdb.Cells(x, 403)
        Me.Business_Partner_Number_ROSA = wsdb.Cells(x, 404)
        Me.Comments_ROSA = wsdb.Cells(x, 405)
        Me.Dog_or_Cat_Flag_ROSA = wsdb.Cells(x, 406)
        Me.FILE_NAME_ROSA = wsdb.Cells(x, 407)
        Me.First_and_last_name_of_main_Auditor_ROSA = wsdb.Cells(x, 408)
        Me.Number_of_Auditors_ROSA = wsdb.Cells(x, 409)
        Me.Number_of_stories_above_grade_ROSA = wsdb.Cells(x, 410)
        Me.Occupancy_frequency_ROSA = wsdb.Cells(x, 411)
        Me.Ownership_Type_ROSA = wsdb.Cells(x, 412)
        Me.Total_conditioned_square_footage_ROSA = wsdb.Cells(x, 413)
        Me.WO_Number_ROSA = wsdb.Cells(x, 414)
        
        Exit Sub
        
    Else
        If x = wsDblr Then
            MsgBox ("The Enrollment ID is not found in the Database")
            Exit Sub
        End If
    End If
Next x

End Sub
