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

Private Sub Cancel_HEAP_Click()
    Unload Me
    frmAdmin.Show
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub Revert_HEAP_Click()
Call UserForm_Initialize

End Sub

Private Sub Save_HEAP_Click()
Set wsdb = Worksheets("Enrollments")
'Dim dbRow As Long

'Enrollment_ID_HEAP = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsdb.Cells(Rows.Count, 3).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, 3) = EID Then
        'push data from form to database
        'HEAP Scheduling

        wsdb.Cells(x, 874) = Me.Customer_contact_mode_HEAP
        wsdb.Cells(x, 875) = Me.Fifth_Contact_Attempt_Date_HEAP
        wsdb.Cells(x, 876) = Me.Fifth_Contact_Attempt_Notes_HEAP
        wsdb.Cells(x, 877) = Me.Fifth_Contact_Attempt_Type_HEAP
        wsdb.Cells(x, 878) = Me.First_Contact_Attempt_Date_HEAP
        wsdb.Cells(x, 879) = Me.First_Contact_Attempt_Notes_HEAP
        wsdb.Cells(x, 880) = Me.First_Contact_Attempt_Type_HEAP
        wsdb.Cells(x, 881) = Me.Follow_up_Date_HEAP
        wsdb.Cells(x, 882) = Me.Fourth_Contact_Attempt_Date_HEAP
        wsdb.Cells(x, 883) = Me.Fourth_Contact_Attempt_Notes_HEAP
        wsdb.Cells(x, 884) = Me.Fourth_Contact_Attempt_Type_HEAP
        wsdb.Cells(x, 885) = Me.Second_Contact_Attempt_Date_HEAP
        wsdb.Cells(x, 886) = Me.Second_Contact_Attempt_Notes_HEAP
        wsdb.Cells(x, 887) = Me.Second_Contact_Attempt_Type_HEAP
        wsdb.Cells(x, 888) = Me.Third_Contact_Attempt_Date_HEAP
        wsdb.Cells(x, 889) = Me.Third_Contact_Attempt_Notes_HEAP
        wsdb.Cells(x, 890) = Me.Third_Contact_Attempt_Type_HEAP
        wsdb.Cells(x, 891) = Me.CANCELLED_date_interfaced_HEAP
        wsdb.Cells(x, 892) = Me.CANCELLED_date_set_HEAP
        wsdb.Cells(x, 893) = Me.COMPLETE_date_interfaced_HEAP
        wsdb.Cells(x, 894) = Me.COMPLETE_date_set_HEAP
        wsdb.Cells(x, 895) = Me.FIRST_CONTACT_date_interfaced_HEAP
        wsdb.Cells(x, 896) = Me.FIRST_CONTACT_date_set_HEAP
        wsdb.Cells(x, 897) = Me.ON_HOLD_date_interfaced_HEAP
        wsdb.Cells(x, 898) = Me.ON_HOLD_date_set_HEAP
        wsdb.Cells(x, 899) = Me.PENDING_1_date_interfaced_HEAP
        wsdb.Cells(x, 900) = Me.PENDING_1_date_set_HEAP
        wsdb.Cells(x, 901) = Me.PENDING_2_date_interfaced_HEAP
        wsdb.Cells(x, 902) = Me.PENDING_2_date_set_HEAP
        wsdb.Cells(x, 903) = Me.PENDING_3_date_interfaced_HEAP
        wsdb.Cells(x, 904) = Me.PENDING_3_date_set_HEAP
        wsdb.Cells(x, 905) = Me.PENDING_4_date_interfaced_HEAP
        wsdb.Cells(x, 906) = Me.PENDING_4_date_set_HEAP
        wsdb.Cells(x, 907) = Me.PENDING_5_date_interfaced_HEAP
        wsdb.Cells(x, 908) = Me.PENDING_5_date_set_HEAP
        wsdb.Cells(x, 909) = Me.RECEIVED_AT_VENDOR_date_interfaced_HEAP
        wsdb.Cells(x, 910) = Me.RECEIVED_AT_VENDOR_date_set_HEAP
        wsdb.Cells(x, 911) = Me.SCHEDULED_date_interfaced_HEAP
        wsdb.Cells(x, 912) = Me.SCHEDULED_date_set_HEAP
        wsdb.Cells(x, 913) = Me.SITE_WORK_COMPLETE_date_interfaced_HEAP
        wsdb.Cells(x, 914) = Me.SITE_WORK_COMPLETE_date_set_HEAP
        wsdb.Cells(x, 915) = Me.Status_Date_HEAP
        wsdb.Cells(x, 916) = Me.Status_HEAP
        wsdb.Cells(x, 917) = Me.Status_Time_HEAP
        wsdb.Cells(x, 918) = Me.WITHDRAWN_date_interfaced_HEAP
        wsdb.Cells(x, 919) = Me.WITHDRAWN_date_set_HEAP
        wsdb.Cells(x, 920) = Me.Air_Leakage_Rating_HEAP
        wsdb.Cells(x, 921) = Me.Auditor_Notes_HEAP
        wsdb.Cells(x, 922) = Me.Blower_door_post_test_HEAP
        wsdb.Cells(x, 923) = Me.Blower_door_pre_test_HEAP
        wsdb.Cells(x, 924) = Me.Building_occupancy_count_HEAP
        wsdb.Cells(x, 925) = Me.Business_Partner_Number_HEAP
        wsdb.Cells(x, 926) = Me.Comments_HEAP
        wsdb.Cells(x, 927) = Me.Dog_or_Cat_Flag_HEAP
        wsdb.Cells(x, 928) = Me.FILE_NAME_HEAP
        wsdb.Cells(x, 929) = Me.First_and_last_name_of_main_Auditor_HEAP
        wsdb.Cells(x, 930) = Me.Number_of_Auditors_HEAP
        wsdb.Cells(x, 931) = Me.Number_of_stories_above_grade_HEAP
        wsdb.Cells(x, 932) = Me.Occupancy_frequency_HEAP
        wsdb.Cells(x, 933) = Me.Ownership_Type_HEAP
        wsdb.Cells(x, 934) = Me.Schedule_Date_HEAP
        wsdb.Cells(x, 935) = Me.Schedule_Time_HEAP
        wsdb.Cells(x, 936) = Me.Total_conditioned_square_footage_HEAP
        wsdb.Cells(x, 937) = Me.WO_Number_HEAP


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

'Enrollment_ID_HEAP = EID
Dim EID As String

EID = currentEnrollment

'last row database
wsDblr = wsdb.Cells(Rows.Count, 3).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, 3) = EID Then
        'push data from database to form
        'HEAP Scheduling

        Me.Enrollment_ID_HEAP = wsdb.Cells(x, 3)
        Me.Customer_contact_mode_HEAP = wsdb.Cells(x, 874)
        Me.Fifth_Contact_Attempt_Date_HEAP = wsdb.Cells(x, 875)
        Me.Fifth_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, 876)
        Me.Fifth_Contact_Attempt_Type_HEAP = wsdb.Cells(x, 877)
        Me.First_Contact_Attempt_Date_HEAP = wsdb.Cells(x, 878)
        Me.First_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, 879)
        Me.First_Contact_Attempt_Type_HEAP = wsdb.Cells(x, 880)
        Me.Follow_up_Date_HEAP = wsdb.Cells(x, 881)
        Me.Fourth_Contact_Attempt_Date_HEAP = wsdb.Cells(x, 882)
        Me.Fourth_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, 883)
        Me.Fourth_Contact_Attempt_Type_HEAP = wsdb.Cells(x, 884)
        Me.Second_Contact_Attempt_Date_HEAP = wsdb.Cells(x, 885)
        Me.Second_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, 886)
        Me.Second_Contact_Attempt_Type_HEAP = wsdb.Cells(x, 887)
        Me.Third_Contact_Attempt_Date_HEAP = wsdb.Cells(x, 888)
        Me.Third_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, 889)
        Me.Third_Contact_Attempt_Type_HEAP = wsdb.Cells(x, 890)
        Me.CANCELLED_date_interfaced_HEAP = wsdb.Cells(x, 891)
        Me.CANCELLED_date_set_HEAP = wsdb.Cells(x, 892)
        Me.COMPLETE_date_interfaced_HEAP = wsdb.Cells(x, 893)
        Me.COMPLETE_date_set_HEAP = wsdb.Cells(x, 894)
        Me.FIRST_CONTACT_date_interfaced_HEAP = wsdb.Cells(x, 895)
        Me.FIRST_CONTACT_date_set_HEAP = wsdb.Cells(x, 896)
        Me.ON_HOLD_date_interfaced_HEAP = wsdb.Cells(x, 897)
        Me.ON_HOLD_date_set_HEAP = wsdb.Cells(x, 898)
        Me.PENDING_1_date_interfaced_HEAP = wsdb.Cells(x, 899)
        Me.PENDING_1_date_set_HEAP = wsdb.Cells(x, 900)
        Me.PENDING_2_date_interfaced_HEAP = wsdb.Cells(x, 901)
        Me.PENDING_2_date_set_HEAP = wsdb.Cells(x, 902)
        Me.PENDING_3_date_interfaced_HEAP = wsdb.Cells(x, 903)
        Me.PENDING_3_date_set_HEAP = wsdb.Cells(x, 904)
        Me.PENDING_4_date_interfaced_HEAP = wsdb.Cells(x, 905)
        Me.PENDING_4_date_set_HEAP = wsdb.Cells(x, 906)
        Me.PENDING_5_date_interfaced_HEAP = wsdb.Cells(x, 907)
        Me.PENDING_5_date_set_HEAP = wsdb.Cells(x, 908)
        Me.RECEIVED_AT_VENDOR_date_interfaced_HEAP = wsdb.Cells(x, 909)
        Me.RECEIVED_AT_VENDOR_date_set_HEAP = wsdb.Cells(x, 910)
        Me.SCHEDULED_date_interfaced_HEAP = wsdb.Cells(x, 911)
        Me.SCHEDULED_date_set_HEAP = wsdb.Cells(x, 912)
        Me.SITE_WORK_COMPLETE_date_interfaced_HEAP = wsdb.Cells(x, 913)
        Me.SITE_WORK_COMPLETE_date_set_HEAP = wsdb.Cells(x, 914)
        Me.Status_Date_HEAP = wsdb.Cells(x, 915)
        Me.Status_HEAP = wsdb.Cells(x, 916)
        Me.Status_Time_HEAP = wsdb.Cells(x, 917)
        Me.WITHDRAWN_date_interfaced_HEAP = wsdb.Cells(x, 918)
        Me.WITHDRAWN_date_set_HEAP = wsdb.Cells(x, 919)
        Me.Air_Leakage_Rating_HEAP = wsdb.Cells(x, 920)
        Me.Auditor_Notes_HEAP = wsdb.Cells(x, 921)
        Me.Blower_door_post_test_HEAP = wsdb.Cells(x, 922)
        Me.Blower_door_pre_test_HEAP = wsdb.Cells(x, 923)
        Me.Building_occupancy_count_HEAP = wsdb.Cells(x, 924)
        Me.Business_Partner_Number_HEAP = wsdb.Cells(x, 925)
        Me.Comments_HEAP = wsdb.Cells(x, 926)
        Me.Dog_or_Cat_Flag_HEAP = wsdb.Cells(x, 927)
        Me.FILE_NAME_HEAP = wsdb.Cells(x, 928)
        Me.First_and_last_name_of_main_Auditor_HEAP = wsdb.Cells(x, 929)
        Me.Number_of_Auditors_HEAP = wsdb.Cells(x, 930)
        Me.Number_of_stories_above_grade_HEAP = wsdb.Cells(x, 931)
        Me.Occupancy_frequency_HEAP = wsdb.Cells(x, 932)
        Me.Ownership_Type_HEAP = wsdb.Cells(x, 933)
        Me.Schedule_Date_HEAP = wsdb.Cells(x, 934)
        Me.Schedule_Time_HEAP = wsdb.Cells(x, 935)
        Me.Total_conditioned_square_footage_HEAP = wsdb.Cells(x, 936)
        Me.WO_Number_HEAP = wsdb.Cells(x, 937)

        Exit Sub
        
    Else
        If x = wsDblr Then
            MsgBox ("The Enrollment ID is not found in the Database")
            Exit Sub
        End If
    End If
Next x

End Sub
