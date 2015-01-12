VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Results_ROSA 
   Caption         =   "Load_Results_ROSA"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   OleObjectBlob   =   "Load_Results_ROSA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Results_ROSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Load_Results_Click()
Me.Hide
frmProcessing.Show
End Sub

Private Sub File_Load_ROSA_Click()
Dim wbLoad As Workbook
Dim wsLoad As Worksheet
Dim ws2Load As Worksheet
Dim fName As String
Dim f2Name As String
Set wsDb = Worksheets("Enrollments")
Set ws2Db = Worksheets("Measures")
Set wbDb = ActiveWorkbook
Dim fs As Object
Dim w As String
Set fs = CreateObject("Scripting.FileSystemObject")

wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row

w = Application.GetOpenFilename( _
Title:="Select the corresponding _assessment.xlsm file to be uploaded", MultiSelect:=False)
If w = "False" Then Exit Sub

fName = Dir(w)


'Check File name to match assessment
If fName = Me.Enrollment_ID_ROSA + "_assessments.xlsm" Then
    'Upload PDF Document
    y = Application.GetOpenFilename( _
    Title:="Select the corresponding _.PDF file to be uploaded", MultiSelect:=False)
    If y = "False" Then Exit Sub
    f2Name = Dir(y)
    
    If Left(f2Name, 12) = Me.Enrollment_ID_ROSA Then
        

        Set wbLoad = Workbooks.Open(Filename:=w)
        Set wsLoad = wbLoad.Worksheets("Enrollments")
        Set ws2Load = wbLoad.Worksheets("Measures")
        'find row in Database for Enrollment ID
        For x = 11 To wsDblr
            If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA) = Me.Enrollment_ID_ROSA Then
                'push data from Selected Audit File to Database
                'ROSA Scheduling
                'Context Data
                wsDb.Range(wsDb.Cells(x, NexantEnrollments.APPLIANCE_AQUARIUM_quantity), wsDb.Cells(x, NexantEnrollments.WINDOW_5_uv_coating)).Value = wsLoad.Range(wsLoad.Cells(11, NexantEnrollments.APPLIANCE_AQUARIUM_quantity), wsLoad.Cells(11, NexantEnrollments.WINDOW_5_uv_coating)).Value
                
                'ROSA Info
                '   Auditors
                wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Number_of_Auditors_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Auditor_Notes_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Auditor_Notes_ROSA).Value
                'Update Comments_ROSA
                wsDb.Cells(x, NexantEnrollments.Comments_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Auditor_Notes_ROSA).Value
                '   Building
                wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Number_of_stories_above_grade_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Occupancy_frequency_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Occupancy_frequency_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Building_occupancy_count_ROSA).Value
                '   Testing
                wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Blower_door_pre_test_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Blower_door_post_test_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Air_Leakage_Rating_ROSA).Value
                wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = wsLoad.Cells(11, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value
                
                wsDb.Cells(x, NexantEnrollments.FILE_NAME_ROSA) = f2Name
                
                'Measures This assumes that the Cell reference for enrollment ID in both Enrollment and Measure tab are identical
                'wsDb.Range(ws2Db.Cells(x, NexantMeasures.Annual_CCF_Savings), ws2Db.Cells(x, NexantMeasures.VRM_Quantity)).Value = ws2Load.Range(wsLoad.Cells(11, NexantEnrollments.Annual_CCF_Savings), ws2Load.Cells(11, NexantMeasures.VRM_Quantity)).Value
                'Need to Add a Column for SITE_VISIT_FILE_ROSA
                'wsDb.Cells(x, NexantEnrollments.SITE_VISIT_FILE_ROSA) = fName
                Me.Site_Visit_File_ROSA = fName
                Me.FILE_NAME_ROSA = f2Name
                
                wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_ROSA).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "COMPLETED"
            End If
        Next x
    
    Else
        MsgBox ("The _.PDF file you selected is incorrect")
        Exit Sub
    End If
Else
    MsgBox ("The _.xlsm file you selected is incorrect")
    Exit Sub
End If

'Copy files to Directory
fs.copyfile w, "C:\Users\bmcgary\Desktop\VBA Coding\Assessment Files" & "\" & fName
fs.copyfile y, "C:\Users\bmcgary\Desktop\VBA Coding\Assessment Files" & "\" & f2Name

Set fs = Nothing

'Closes Audit File
wbLoad.Close SaveChanges:=False
Call UserForm_Activate

End Sub


Private Sub Scheduled_Listbox_Click()
Me.Enrollment_ID_ROSA = Scheduled_Listbox.Value
End Sub

Private Sub UserForm_Activate()
Set wsDb = Worksheets("Enrollments")

'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row

Scheduled_Listbox.Clear
Me.Enrollment_ID_ROSA = ""
Me.FILE_NAME_ROSA = ""
Me.Site_Visit_File_ROSA = ""

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Status_ROSA) = "SCHEDULED" Then
        'push data from database to form
        'ROSA Scheduling
        With Scheduled_Listbox
            .AddItem wsDb.Cells(x, NexantEnrollments.Enrollment_ID_ROSA)
        End With
    
    End If
Next x

End Sub


