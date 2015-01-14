VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Results_HEAP 
   Caption         =   "Load_Results_HEAP"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   OleObjectBlob   =   "Load_Results_HEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Results_HEAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cancel_Load_Results_Click()
Me.Hide
frmProcessing.Show
End Sub

Private Sub File_Load_HEAP_Click()
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

wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row

w = Application.GetOpenFilename( _
Title:="Select the corresponding _assessment.xlsm file to be uploaded", MultiSelect:=False)
If w = "False" Then Exit Sub

fName = Dir(w)


'Check File name to match assessment
If fName = Me.Enrollment_ID_HEAP + "_assessments.xlsm" Then
    'Upload PDF Document
    y = Application.GetOpenFilename( _
    Title:="Select the corresponding _.PDF file to be uploaded", MultiSelect:=False)
    If y = "False" Then Exit Sub
    f2Name = Dir(y)
    
    If Left(f2Name, 12) = Me.Enrollment_ID_HEAP Then
        

        Set wbLoad = Workbooks.Open(filename:=w)
        Set wsLoad = wbLoad.Worksheets("Enrollments")
        Set ws2Load = wbLoad.Worksheets("Measures")
        'find row in Database for Enrollment ID
        For x = 11 To wsDblr
            If wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = Me.Enrollment_ID_HEAP Then
                'push data from Selected Audit File to Database
                'HEAP Scheduling
                'Context Data
                wsDb.Range(wsDb.Cells(x, NexantEnrollments.APPLIANCE_AQUARIUM_quantity), wsDb.Cells(x, NexantEnrollments.WINDOW_5_uv_coating)).Value = wsLoad.Range(wsLoad.Cells(11, NexantEnrollments.APPLIANCE_AQUARIUM_quantity), wsLoad.Cells(11, NexantEnrollments.WINDOW_5_uv_coating)).Value
                
                'HEAP Info
                '   Auditors
                wsDb.Cells(x, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Number_of_Auditors_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Number_of_Auditors_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Auditor_Notes_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Auditor_Notes_HEAP).Value
                'Update Comments_HEAP
                wsDb.Cells(x, NexantEnrollments.Comments_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Auditor_Notes_HEAP).Value
                '   Building
                wsDb.Cells(x, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Number_of_stories_above_grade_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Number_of_stories_above_grade_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Occupancy_frequency_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Occupancy_frequency_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Occupancy_frequency_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Building_occupancy_count_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Building_occupancy_count_HEAP).Value
                '   Testing
                wsDb.Cells(x, NexantEnrollments.Blower_door_pre_test_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Blower_door_pre_test_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Blower_door_post_test_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Blower_door_post_test_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Air_Leakage_Rating_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Air_Leakage_Rating_HEAP).Value
                wsDb.Cells(x, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value = wsLoad.Cells(11, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value
                'File Names
                Me.Site_Visit_File_HEAP = fName
                Me.FILE_NAME_HEAP = f2Name
                wsDb.Cells(x, NexantEnrollments.FILE_NAME_HEAP) = f2Name
                wsDb.Cells(x, NexantEnrollments.Site_Visit_File_HEAP) = fName
                
                'Measures This assumes that the Cell reference for enrollment ID in both Enrollment and Measure tab are identical
                ws2Db.Range(ws2Db.Cells(x, NexantMeasures.Annual_CCF_Savings), ws2Db.Cells(x, NexantMeasures.VRM_Quantity)).Value = ws2Load.Range(ws2Load.Cells(11, NexantMeasures.Annual_CCF_Savings), ws2Load.Cells(11, NexantMeasures.VRM_Quantity)).Value
                
                'Set Dates and Status
                wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.COMPLETE_date_set_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now() + TimeValue("00:00:01")), "HHMMSS")
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
                wsDb.Cells(x, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
                wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "COMPLETED"
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
fs.copyfile w, Application.ActiveWorkbook.Path & "\Assessment Files" & "\" & fName
fs.copyfile y, Application.ActiveWorkbook.Path & "\Assessment Files" & "\" & f2Name

Set fs = Nothing
MsgBox ("The data has been uploaded and the documents saved to the drive")
'Closes Audit File
wbLoad.Close SaveChanges:=False
Call UserForm_Activate

End Sub


Private Sub Scheduled_Listbox_Click()
Me.Enrollment_ID_HEAP = Scheduled_Listbox.Value
End Sub

Private Sub UserForm_Activate()
Set wsDb = Worksheets("Enrollments")

'last row database
wsDblr = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row

Scheduled_Listbox.Clear
Me.Enrollment_ID_HEAP = ""
Me.FILE_NAME_HEAP = ""
Me.Site_Visit_File_HEAP = ""

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsDb.Cells(x, NexantEnrollments.Status_HEAP) = "SCHEDULED" Then
        'push data from database to form
        'HEAP Scheduling
        With Scheduled_Listbox
            .AddItem wsDb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
        End With
    
    End If
Next x

End Sub


