VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Export"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub

Private Sub CommandButton1_Click()
' This routine exports records for enrollments:
'   1. whose Status_date/Status_time is more recent than LastStatusRun

' For enrollments who meet the criteria above and whose status is one of the 6 below,
' records are created for each status with a blank STATUSNAME_date_interfaced field from RECEIVED
' AT VENDOR up to and including the current status.  As a record is created for each status, the
' current date/time are written to the STATUSNAME_date_interfaced field.
'   1. RECEIVED AT VENDOR
'   2. FIRST CONTACT
'   3. PENDING
'   4. SCHEDULED
'   5. SITE WORK COMPLETE
'   6. COMPLETE

' For enrollments who meet the criteria above and whose status is one of the following,
' a record is created only for that status and its corresponding STATUSNAME_date_interfaced
' field is set.
'   1. ON-HOLD
'   2. CANCELLED

Set wsDb = Worksheets("Enrollments")
Set wspm = Worksheets("PM")
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False
Set wbROSA = objExcel.Workbooks.Add
Set wsROSA = wbROSA.Worksheets("Sheet1")
Set wbHEAP = objExcel.Workbooks.Add
Set wsHEAP = wbHEAP.Worksheets("Sheet1")

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'last row database
wsDblr_ROSA = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_ROSA).End(xlUp).row
wsDblr_HEAP = wsDb.Cells(Rows.Count, NexantEnrollments.Enrollment_ID_HEAP).End(xlUp).row
LastRow = WorksheetFunction.Max(wsDblr_ROSA, wsDblr_HEAP)

'collect last datetimes Status reports were run
LastStatusRunROSA = ParseDateTime(wspm.Cells(PMINRows.PMROSAStatus, 2).Value)
LastStatusRunHEAP = ParseDateTime(wspm.Cells(PMINRows.PMHEAPStatus, 2).Value)

'counters for active rows to use on export sheets
jr = 2
jh = 2

For i = 11 To LastRow
    'First check if ROSA enrollment has been created
    If wsDb.Cells(i, NexantEnrollments.Status_Date_ROSA).Value <> "" Then
        Status_datetime_ROSA = ParseDateTime(wsDb.Cells(i, NexantEnrollments.Status_Date_ROSA).Value + _
                                       ":" + wsDb.Cells(i, NexantEnrollments.Status_Time_ROSA).Value)
        If Status_datetime_ROSA > LastStatusRunROSA Then
            'Process ROSA enrollment
            vStatus_ROSA = wsDb.Cells(i, NexantEnrollments.Status_ROSA).Value
            Select Case vStatus_ROSA:
                Case "ON-HOLD":
                    'send OH; store OH interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.ON_HOLD_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                Case "CANCELLED":
                    'send CAN; store CAN interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.CANCELLED_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                Case "RECEIVED AT VENDOR":
                    'send RAV; store RAV interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                Case "FIRST CONTACT": 'FC, RAV
                    'send FC; store FC interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = "" Then
                        WriteRAV jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "PENDING": 'P, FC, and RAV
                    'send P; store P interfaced datetime into appropriate PENDING_X version
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    If wsDb.Cells(i, NexantEnrollments.PENDING_1_date_interfaced_ROSA).Value = "" Then
                        wsDb.Cells(i, NexantEnrollments.PENDING_1_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    ElseIf wsDb.Cells(i, NexantEnrollments.PENDING_2_date_interfaced_ROSA).Value = "" Then
                        wsDb.Cells(i, NexantEnrollments.PENDING_2_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    ElseIf wsDb.Cells(i, NexantEnrollments.PENDING_3_date_interfaced_ROSA).Value = "" Then
                        wsDb.Cells(i, NexantEnrollments.PENDING_3_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    ElseIf wsDb.Cells(i, NexantEnrollments.PENDING_4_date_interfaced_ROSA).Value = "" Then
                        wsDb.Cells(i, NexantEnrollments.PENDING_4_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    ElseIf wsDb.Cells(i, NexantEnrollments.PENDING_5_date_interfaced_ROSA).Value = "" Then
                        wsDb.Cells(i, NexantEnrollments.PENDING_5_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = "" Then
                        WriteFC jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = "" Then
                        WriteRAV jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "SCHEDULED": 'set S, FC, and RAV
                    'send S; store S interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = "" Then
                        WriteFC jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = "" Then
                        WriteRAV jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "SITE WORK COMPLETE": 'set SWC, S, FC, and RAV
                    'send SWC; store SWC interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                    'send S if not already sent; store S interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_ROSA).Value = "" Then
                        WriteS jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = "" Then
                        WriteFC jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = "" Then
                        WriteRAV jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "COMPLETE": 'set C, SWC, S, FC, and RAV
                    'send C; store C interfaced datetime
                    WriteCurrent jr, i, wsROSA, wsDb, "ROSA", 2
                    jr = jr + 1
                    wsDb.Cells(i, NexantEnrollments.COMPLETE_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    
                    'send SWC if not already sent; store SWC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_ROSA).Value = "" Then
                        WriteSWC jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send S if not already sent; store S interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_ROSA).Value = "" Then
                        WriteS jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = "" Then
                        WriteFC jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = "" Then
                        WriteRAV jr, i, wsROSA, wsDb, "ROSA", 2
                        jr = jr + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_ROSA).Value = NexantLGEDateTimeNow
                    End If
                    
            End Select
        End If
    End If
    
    'First check if HEAP enrollment has been created
    If wsDb.Cells(i, NexantEnrollments.Status_Date_HEAP).Value <> "" Then
        Status_datetime_HEAP = ParseDateTime(wsDb.Cells(i, NexantEnrollments.Status_Date_HEAP).Value + _
                                       ":" + wsDb.Cells(i, NexantEnrollments.Status_Time_HEAP).Value)
        If Status_datetime_HEAP > LastStatusRunHEAP Then
            'Process HEAP enrollment
            
        End If
    End If
Next i

wsROSA.SaveAs ActiveWorkbook.Path + "\Export Files\DSM_ROSA_status_in.txt", xlCSV
wbROSA.Close
wsHEAP.SaveAs ActiveWorkbook.Path + "\Export Files\DSM_HEAP_status_in.txt", xlCSV
wbHEAP.Close
objExcel.Quit

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Activate()
    Me.Last_Run_Date_Status.BackColor = rgbLightGrey
    Me.Last_Run_Date_Results.BackColor = rgbLightGrey
    Me.Last_Run_Date_Context.BackColor = rgbLightGrey
    Me.Last_Run_Date_PDF_Control.BackColor = rgbLightGrey
    Me.Last_Run_Date_Recon.BackColor = rgbLightGrey
    Me.Last_Run_Date_Invoice.BackColor = rgbLightGrey
    Me.Last_Run_Date_Status.Enabled = False
    Me.Last_Run_Date_Results.Enabled = False
    Me.Last_Run_Date_Context.Enabled = False
    Me.Last_Run_Date_PDF_Control.Enabled = False
    Me.Last_Run_Date_Recon.Enabled = False
    Me.Last_Run_Date_Invoice.Enabled = False
End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Function ParseDateTime(dt As String) As Date
    If dt = "" Then
        ParseDateTime = DateValue("2000-01-01") + TimeValue("00:00:00")
    Else:
        ParseDateTime = DateValue(Left(dt, 4) + "-" + Mid(dt, 5, 2) + "-" + Mid(dt, 7, 2)) + _
                   TimeValue(Mid(dt, 10, 2) + ":" + Mid(dt, 12, 2) + ":" + Mid(dt, 14, 2))
    End If
End Function

Function NexantLGEDateTimeNow()
   NexantLGEDateTimeNow = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
End Function

Function WriteStaticColumns(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 1).Value = RecordType
    wsDest.Cells(jr, 2).NumberFormat = "@"
    wsDest.Cells(jr, 2).Value = wsSour.Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value
    wsDest.Cells(jr, 3).Value = wsSour.Cells(i, NexantEnrollments.Company_Acronym).Value
    wsDest.Cells(jr, 4).NumberFormat = "@"
    wsDest.Cells(jr, 4).Value = wsSour.Cells(i, NexantEnrollments.Account_Number).Value
    wsDest.Cells(jr, 5).NumberFormat = "@"
    wsDest.Cells(jr, 5).Value = wsSour.Cells(i, NexantEnrollments.Premise_ID).Value
    wsDest.Cells(jr, 6).NumberFormat = "@"
    wsDest.Cells(jr, 6).Value = wsSour.Cells(i, NexantEnrollments.WO_Number_ROSA).Value
    wsDest.Cells(jr, 10).NumberFormat = "@"
    wsDest.Cells(jr, 10).Value = wsSour.Cells(i, NexantEnrollments.Follow_up_Date_ROSA).Value
    wsDest.Cells(jr, 13).Value = wsSour.Cells(i, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value
    wsDest.Cells(jr, 14).Value = wsSour.Cells(i, NexantEnrollments.Residence_Building_Type).Value
    wsDest.Cells(jr, 15).Value = wsSour.Cells(i, NexantEnrollments.Residence_Building_Class).Value
    wsDest.Cells(jr, 16).Value = wsSour.Cells(i, NexantEnrollments.Year_building_constructed).Value
    wsDest.Cells(jr, 17).Value = wsSour.Cells(i, NexantEnrollments.Building_occupancy_count_ROSA).Value
    wsDest.Cells(jr, 18).Value = wsSour.Cells(i, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA).Value
    wsDest.Cells(jr, 19).Value = wsSour.Cells(i, NexantEnrollments.Primary_contact_name).Value
    wsDest.Cells(jr, 20).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Address).Value
    wsDest.Cells(jr, 21).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Address_City).Value
    wsDest.Cells(jr, 22).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Address_State).Value
    wsDest.Cells(jr, 23).NumberFormat = "@"
    wsDest.Cells(jr, 23).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Address_Zip).Value
    wsDest.Cells(jr, 24).NumberFormat = "@"
    wsDest.Cells(jr, 24).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Phone).Value
    wsDest.Cells(jr, 25).NumberFormat = "@"
    wsDest.Cells(jr, 25).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_phone_extension).Value
    wsDest.Cells(jr, 26).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_Email).Value
    wsDest.Cells(jr, 27).NumberFormat = "@"
    wsDest.Cells(jr, 27).Value = wsSour.Cells(i, NexantEnrollments.Primary_Contact_mobile_phone).Value
    wsDest.Cells(jr, 28).Value = wsSour.Cells(i, NexantEnrollments.Ownership_Type_ROSA).Value
    wsDest.Cells(jr, 29).Value = wsSour.Cells(i, NexantEnrollments.Service_Class).Value
    wsDest.Cells(jr, 30).Value = wsSour.Cells(i, NexantEnrollments.Schedule_Date_ROSA).Value
    wsDest.Cells(jr, 31).NumberFormat = "@"
    wsDest.Cells(jr, 31).Value = wsSour.Cells(i, NexantEnrollments.Schedule_Time_ROSA).Value
    wsDest.Cells(jr, 32).Value = wsSour.Cells(i, NexantEnrollments.Number_of_Auditors_ROSA).Value
    wsDest.Cells(jr, 33).Value = wsSour.Cells(i, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value
    wsDest.Cells(jr, 34).Value = wsSour.Cells(i, NexantEnrollments.Occupancy_frequency_ROSA).Value
    wsDest.Cells(jr, 35).Value = wsSour.Cells(i, NexantEnrollments.Number_of_stories_above_grade_ROSA).Value
    wsDest.Cells(jr, 36).Value = wsSour.Cells(i, NexantEnrollments.Air_Leakage_Rating_ROSA).Value
    wsDest.Cells(jr, 37).Value = wsSour.Cells(i, NexantEnrollments.Blower_door_pre_test_ROSA).Value
    wsDest.Cells(jr, 38).Value = wsSour.Cells(i, NexantEnrollments.Blower_door_post_test_ROSA).Value
    wsDest.Cells(jr, 39).Value = wsSour.Cells(i, NexantEnrollments.CFM_Reduction).Value
    wsDest.Cells(jr, 40).Value = wsSour.Cells(i, NexantEnrollments.Auditor_Notes_ROSA).Value
    wsDest.Cells(jr, 41).Value = ProgName
    wsDest.Cells(jr, 42).Value = wsSour.Cells(i, NexantEnrollments.Verification_Class).Value
    wsDest.Cells(jr, 43).Value = 99999
    wsDest.Cells(jr, 44).Value = 99999
    wsDest.Cells(jr, 45).Value = 99999
    wsDest.Cells(jr, 46).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Name).Value
    wsDest.Cells(jr, 47).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Address).Value
    wsDest.Cells(jr, 48).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Address_City).Value
    wsDest.Cells(jr, 49).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Address_State).Value
    wsDest.Cells(jr, 50).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Address_Zip).Value
    wsDest.Cells(jr, 51).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Phone).Value
    wsDest.Cells(jr, 52).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_phone_extension).Value
    wsDest.Cells(jr, 53).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_Email).Value
    wsDest.Cells(jr, 54).Value = wsSour.Cells(i, NexantEnrollments.Remit_to_Contact_mobile_phone).Value
End Function

Function WriteCurrent(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 7).Value = wsSour.Cells(i, NexantEnrollments.Status_ROSA).Value
    wsDest.Cells(jr, 8).Value = wsSour.Cells(i, NexantEnrollments.Status_Date_ROSA).Value
    wsDest.Cells(jr, 9).Value = wsSour.Cells(i, NexantEnrollments.Status_Time_ROSA).Value
    wsDest.Cells(jr, 11).Value = wsSour.Cells(i, NexantEnrollments.Customer_contact_mode_ROSA).Value
    wsDest.Cells(jr, 12).Value = wsSour.Cells(i, NexantEnrollments.Comments_ROSA).Value
    WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
End Function

Function WriteRAV(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 7).Value = "RECEIVED AT VENDOR"
    wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA).Value, 8)
    wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA).Value, 6)
    wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
    wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
    WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
End Function

Function WriteFC(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 7).Value = "FIRST CONTACT"
    wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).Value, 8)
    wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).Value, 6)
    wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
    wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
    WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
End Function

Function WriteS(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 7).Value = "SCHEDULED"
    wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_ROSA).Value, 8)
    wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_ROSA).Value, 6)
    wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
    wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
    WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
End Function

Function WriteSWC(jr, i, wsDest, wsSour, ProgName, RecordType)
    wsDest.Cells(jr, 7).Value = "SITE WORK COMPLETE"
    wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA).Value, 8)
    wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA).Value, 6)
    wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
    wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
    WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
End Function

