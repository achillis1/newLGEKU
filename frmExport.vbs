VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Export"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   OleObjectBlob   =   "frmExport.frx":0000
   ShowModal       =   0   'False
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
' CRITERIA: This routine exports records for enrollments who meet the following criteria:
'   1. Status_date/Status_time is more recent than LastStatusRun

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

' PENDINGS are driven by the Contacts tab. For each row and for ROSA/HEAP within each row, the
' following logic is implemented:
'   1. If ID <> "", AND
'   2. If Date_Interface = "", THEN
'   3. write a PENDING record.

' For enrollments who meet the criteria above and whose status is one of the following,
' a record is created only for that status and its corresponding STATUSNAME_date_interfaced
' field is set.
'   1. ON-HOLD
'   2. CANCELLED

Set wsDb = Worksheets("Enrollments")
Set wspm = Worksheets("PM")
Set wsCont = Worksheets("Contacts")
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False
Set wbROSA = objExcel.Workbooks.Add
Set wsROSA = wbROSA.Worksheets("Sheet1")
Set wbHEAP = objExcel.Workbooks.Add
Set wsHEAP = wbHEAP.Worksheets("Sheet1")

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'last row database
nCont = wsCont.Cells(Rows.Count, 1).End(xlUp).row

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

'write all non-PENDING statuses from Enrollments tab
For i = 11 To LastRow
    'First check if ROSA enrollment has been created
    If wsDb.Cells(i, NexantEnrollments.Status_Date_ROSA).Value <> "" Then
        Status_datetime_ROSA = ParseDateTime(wsDb.Cells(i, NexantEnrollments.Status_Date_ROSA).Value + _
                                       ":" + wsDb.Cells(i, NexantEnrollments.Status_Time_ROSA).Value)
        If Status_datetime_ROSA > LastStatusRunROSA Then
            'Process ROSA enrollment
            vStatus_ROSA = wsDb.Cells(i, NexantEnrollments.Status_ROSA).Value
            
            'write all non-PENDING statuses
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
            vStatus_HEAP = wsDb.Cells(i, NexantEnrollments.Status_HEAP).Value
            
            'write all non-PENDING statuses
            Select Case vStatus_HEAP:
                Case "ON-HOLD":
                    'send OH; store OH interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.ON_HOLD_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                Case "CANCELLED":
                    'send CAN; store CAN interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.CANCELLED_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                Case "RECEIVED AT VENDOR":
                    'send RAV; store RAV interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                Case "FIRST CONTACT": 'FC, RAV
                    'send FC; store FC interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = "" Then
                        WriteRAV jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "SCHEDULED": 'set S, FC, and RAV
                    'send S; store S interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = "" Then
                        WriteFC jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = "" Then
                        WriteRAV jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "SITE WORK COMPLETE": 'set SWC, S, FC, and RAV
                    'send SWC; store SWC interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                    'send S if not already sent; store S interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_HEAP).Value = "" Then
                        WriteS jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = "" Then
                        WriteFC jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = "" Then
                        WriteRAV jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                Case "COMPLETE": 'set C, SWC, S, FC, and RAV
                    'send C; store C interfaced datetime
                    WriteCurrent jh, i, wsHEAP, wsDb, "HEAP", 2
                    jh = jh + 1
                    wsDb.Cells(i, NexantEnrollments.COMPLETE_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    
                    'send SWC if not already sent; store SWC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_HEAP).Value = "" Then
                        WriteSWC jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send S if not already sent; store S interfaced date
                    If wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_HEAP).Value = "" Then
                        WriteS jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.SCHEDULED_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send FC if not already sent; store FC interfaced date
                    If wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = "" Then
                        WriteFC jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
                    'send RAV if not already sent; store RAV interfaced date
                    If wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = "" Then
                        WriteRAV jh, i, wsHEAP, wsDb, "HEAP", 2
                        jh = jh + 1
                        wsDb.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP).Value = NexantLGEDateTimeNow
                    End If
                    
            End Select
        End If
    End If

Next i

'write PENDING statuses from Contacts tab
For k = 2 To nCont
    'ROSA
    If wsCont.Cells(k, NexantContacts.Enrollment_ID_ROSA).Value <> "" Then
        If wsCont.Cells(k, NexantContacts.ROSA_Contact_Date_Interface).Value = "" Then
            Set Enroll_ID = wsDb.Range("B:B").Find(What:=wsCont.Cells(k, NexantContacts.Enrollment_ID_ROSA).Value, _
                After:=wsDb.Range("B11"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Enroll_ID Is Nothing Then
                wsCont.Cells(k, NexantContacts.ROSA_Contact_Date_Interface).Value = "ROSA ID not found"
            Else
                r = Enroll_ID.row
                
                wsROSA.Cells(jr, 7).Value = "PENDING"
                wsROSA.Cells(jr, 8).Value = Left(wsCont.Cells(i, NexantContacts.ROSA_Contact_DateTime).Value, 8)
                wsROSA.Cells(jr, 9).Value = Right(wsCont.Cells(i, NexantContacts.ROSA_Contact_DateTime).Value, 6)
                wsROSA.Cells(jr, 11).Value = wsCont.Cells(i, NexantContacts.ROSA_Contact_Attempt_Type).Value
                wsROSA.Cells(jr, 12).Value = wsCont.Cells(i, NexantContacts.ROSA_Contact_Attempt_Notes).Value
                
                WriteStaticColumns jr, r, wsROSA, wsDb, "ROSA", 2
                jr = jr + 1
                wsCont.Cells(k, NexantContacts.ROSA_Contact_Date_Interface).Value = NexantLGEDateTimeNow
            End If
            
        End If
    End If
    
    'HEAP
    If wsCont.Cells(k, NexantContacts.Enrollment_ID_HEAP).Value <> "" Then
        If wsCont.Cells(k, NexantContacts.HEAP_Contact_Date_Interface).Value = "" Then
            Set Enroll_ID = wsDb.Range("C:C").Find(What:=wsCont.Cells(k, NexantContacts.Enrollment_ID_HEAP).Value, _
                After:=wsDb.Range("C11"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Enroll_ID Is Nothing Then
                wsCont.Cells(k, NexantContacts.HEAP_Contact_Date_Interface).Value = "HEAP ID not found"
            Else
                r = Enroll_ID.row
                
                wsHEAP.Cells(jh, 7).Value = "PENDING"
                wsHEAP.Cells(jh, 8).Value = Left(wsCont.Cells(i, NexantContacts.HEAP_Contact_DateTime).Value, 8)
                wsHEAP.Cells(jh, 9).Value = Right(wsCont.Cells(i, NexantContacts.HEAP_Contact_DateTime).Value, 6)
                wsHEAP.Cells(jh, 11).Value = wsCont.Cells(i, NexantContacts.HEAP_Contact_Attempt_Type).Value
                wsHEAP.Cells(jh, 12).Value = wsCont.Cells(i, NexantContacts.HEAP_Contact_Attempt_Notes).Value
                
                WriteStaticColumns jh, r, wsHEAP, wsDb, "HEAP", 2
                jh = jh + 1
                wsCont.Cells(k, NexantContacts.HEAP_Contact_Date_Interface).Value = NexantLGEDateTimeNow
            End If
            
        End If
    End If

Next k

'write header rows
wsROSA.Cells(1, 1).Value = 1
wsROSA.Cells(1, 2).Value = Format(LocalTimeToET(Now()), "YYYYMMDD")
wsROSA.Cells(1, 3).Value = "ROSA STATUS"
wsHEAP.Cells(1, 1).Value = 1
wsHEAP.Cells(1, 2).Value = Format(LocalTimeToET(Now()), "YYYYMMDD")
wsHEAP.Cells(1, 3).Value = "HEAP STATUS"

'write footer rows
wsROSA.Cells(jr, 1).Value = 3
wsROSA.Cells(jr, 2).Value = jr - 2
wsHEAP.Cells(jh, 1).Value = 3
wsHEAP.Cells(jh, 2).Value = jh - 2

tnow = NexantLGEDateTimeNow

Dim m As Integer
Open ActiveWorkbook.Path + "/fromNexant/dsm_rosa_status_in.txt" For Output As #1
'header
Print #1, writeline(wsROSA, 1, 1, 3)
'data
For m = 2 To jr - 1
    Print #1, writeline(wsROSA, m, 1, 54)
Next m
'footer
Print #1, writeline(wsROSA, jr, 1, 2)
Close #1
wbROSA.Close
FileCopy ActiveWorkbook.Path + "/fromNexant/dsm_rosa_status_in.txt", _
         ActiveWorkbook.Path + "/fromNexantArchive/status_in/dsm_rosa_status_in_" + tnow + ".txt"

Open ActiveWorkbook.Path + "/fromNexant/dsm_heap_status_in.txt" For Output As #1
'header
Print #1, writeline(wsHEAP, 1, 1, 3)
'data
For m = 2 To jh - 1
    Print #1, writeline(wsHEAP, m, 1, 54)
Next m
'footer
Print #1, writeline(wsHEAP, jh, 1, 2)
Close #1
wbHEAP.Close
FileCopy ActiveWorkbook.Path + "/fromNexant/dsm_heap_status_in.txt", _
         ActiveWorkbook.Path + "/fromNexantArchive/status_in/dsm_heap_status_in_" + tnow + ".txt"

objExcel.Quit

wspm.Cells(PMINRows.PMROSAStatus, 2).Value = tnow
wspm.Cells(PMINRows.PMHEAPStatus, 2).Value = tnow
Me.Last_Run_Date_Status.Value = tnow

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "The status_in files were successfully created at " & tnow & "," & Chr(10) & "with " & jr - 2 & " ROSAs and " & jh - 2 & " HEAPs."

End Sub

' ir is the row number, ii is the start column number, jj is the end column number
Function writeline(wsSour, ByVal ir As Integer, ByVal ii As Integer, ByVal jj As Integer)
    writeline = ""
    
    For j = ii To jj
        Select Case j
            Case ii
                writeline = CStr(wsSour.Cells(ir, j).Value)
            Case jj
                writeline = writeline + "," + CStr(wsSour.Cells(ir, j).Value)
            Case Else
                writeline = writeline + "," + CStr(wsSour.Cells(ir, j).Value)
        End Select
    Next j
End Function

Private Sub UserForm_Activate()
    Set wspm = Worksheets("PM")
    'setting single box to value for ROSA since both ROSA/HEAP are exported at same time by same macro
    Me.Last_Run_Date_Status.Value = wspm.Cells(PMINRows.PMROSAStatus, 2).Value
    Me.Last_Run_Date_Results.Value = wspm.Cells(PMINRows.PMROSAResults, 2).Value
    Me.Last_Run_Date_Context.Value = wspm.Cells(PMINRows.PMROSAContextdata, 2).Value
    Me.Last_Run_Date_PDF_Control.Value = wspm.Cells(PMINRows.PMROSAPDFControl, 2).Value
    Me.Last_Run_Date_Recon.Value = wspm.Cells(PMINRows.PMROSARecon, 2).Value
    Me.Last_Run_Date_Invoice.Value = wspm.Cells(PMINRows.PMROSAInvoice, 2).Value
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
    If ProgName = "ROSA" Then
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
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 1).Value = RecordType
        wsDest.Cells(jr, 2).NumberFormat = "@"
        wsDest.Cells(jr, 2).Value = wsSour.Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value
        wsDest.Cells(jr, 3).Value = wsSour.Cells(i, NexantEnrollments.Company_Acronym).Value
        wsDest.Cells(jr, 4).NumberFormat = "@"
        wsDest.Cells(jr, 4).Value = wsSour.Cells(i, NexantEnrollments.Account_Number).Value
        wsDest.Cells(jr, 5).NumberFormat = "@"
        wsDest.Cells(jr, 5).Value = wsSour.Cells(i, NexantEnrollments.Premise_ID).Value
        wsDest.Cells(jr, 6).NumberFormat = "@"
        wsDest.Cells(jr, 6).Value = wsSour.Cells(i, NexantEnrollments.WO_Number_HEAP).Value
        wsDest.Cells(jr, 10).NumberFormat = "@"
        wsDest.Cells(jr, 10).Value = wsSour.Cells(i, NexantEnrollments.Follow_up_Date_HEAP).Value
        wsDest.Cells(jr, 13).Value = wsSour.Cells(i, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value
        wsDest.Cells(jr, 14).Value = wsSour.Cells(i, NexantEnrollments.Residence_Building_Type).Value
        wsDest.Cells(jr, 15).Value = wsSour.Cells(i, NexantEnrollments.Residence_Building_Class).Value
        wsDest.Cells(jr, 16).Value = wsSour.Cells(i, NexantEnrollments.Year_building_constructed).Value
        wsDest.Cells(jr, 17).Value = wsSour.Cells(i, NexantEnrollments.Building_occupancy_count_HEAP).Value
        wsDest.Cells(jr, 18).Value = wsSour.Cells(i, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP).Value
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
        wsDest.Cells(jr, 28).Value = wsSour.Cells(i, NexantEnrollments.Ownership_Type_HEAP).Value
        wsDest.Cells(jr, 29).Value = wsSour.Cells(i, NexantEnrollments.Service_Class).Value
        wsDest.Cells(jr, 30).Value = wsSour.Cells(i, NexantEnrollments.Schedule_Date_HEAP).Value
        wsDest.Cells(jr, 31).NumberFormat = "@"
        wsDest.Cells(jr, 31).Value = wsSour.Cells(i, NexantEnrollments.Schedule_Time_HEAP).Value
        wsDest.Cells(jr, 32).Value = wsSour.Cells(i, NexantEnrollments.Number_of_Auditors_HEAP).Value
        wsDest.Cells(jr, 33).Value = wsSour.Cells(i, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value
        wsDest.Cells(jr, 34).Value = wsSour.Cells(i, NexantEnrollments.Occupancy_frequency_HEAP).Value
        wsDest.Cells(jr, 35).Value = wsSour.Cells(i, NexantEnrollments.Number_of_stories_above_grade_HEAP).Value
        wsDest.Cells(jr, 36).Value = wsSour.Cells(i, NexantEnrollments.Air_Leakage_Rating_HEAP).Value
        wsDest.Cells(jr, 37).Value = wsSour.Cells(i, NexantEnrollments.Blower_door_pre_test_HEAP).Value
        wsDest.Cells(jr, 38).Value = wsSour.Cells(i, NexantEnrollments.Blower_door_post_test_HEAP).Value
        wsDest.Cells(jr, 39).Value = wsSour.Cells(i, NexantEnrollments.CFM_Reduction).Value
        wsDest.Cells(jr, 40).Value = wsSour.Cells(i, NexantEnrollments.Auditor_Notes_HEAP).Value
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
    End If
End Function

Function WriteCurrent(jr, i, wsDest, wsSour, ProgName, RecordType)
    If ProgName = "ROSA" Then
        wsDest.Cells(jr, 7).Value = wsSour.Cells(i, NexantEnrollments.Status_ROSA).Value
        wsDest.Cells(jr, 8).Value = wsSour.Cells(i, NexantEnrollments.Status_Date_ROSA).Value
        wsDest.Cells(jr, 9).Value = wsSour.Cells(i, NexantEnrollments.Status_Time_ROSA).Value
        wsDest.Cells(jr, 11).Value = wsSour.Cells(i, NexantEnrollments.Customer_contact_mode_ROSA).Value
        wsDest.Cells(jr, 12).Value = wsSour.Cells(i, NexantEnrollments.Comments_ROSA).Value
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 7).Value = wsSour.Cells(i, NexantEnrollments.Status_HEAP).Value
        wsDest.Cells(jr, 8).Value = wsSour.Cells(i, NexantEnrollments.Status_Date_HEAP).Value
        wsDest.Cells(jr, 9).Value = wsSour.Cells(i, NexantEnrollments.Status_Time_HEAP).Value
        wsDest.Cells(jr, 11).Value = wsSour.Cells(i, NexantEnrollments.Customer_contact_mode_HEAP).Value
        wsDest.Cells(jr, 12).Value = wsSour.Cells(i, NexantEnrollments.Comments_HEAP).Value
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    End If
End Function

Function WriteRAV(jr, i, wsDest, wsSour, ProgName, RecordType)
    If ProgName = "ROSA" Then
        wsDest.Cells(jr, 7).Value = "RECEIVED AT VENDOR"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_ROSA).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 7).Value = "RECEIVED AT VENDOR"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_HEAP).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.RECEIVED_AT_VENDOR_date_set_HEAP).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    End If
End Function

Function WriteFC(jr, i, wsDest, wsSour, ProgName, RecordType)
    If ProgName = "ROSA" Then
        wsDest.Cells(jr, 7).Value = "FIRST CONTACT"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_ROSA).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 7).Value = "FIRST CONTACT"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_HEAP).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.FIRST_CONTACT_date_set_HEAP).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    End If
End Function

Function WriteS(jr, i, wsDest, wsSour, ProgName, RecordType)
    If ProgName = "ROSA" Then
        wsDest.Cells(jr, 7).Value = "SCHEDULED"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_ROSA).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_ROSA).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 7).Value = "SCHEDULED"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_HEAP).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SCHEDULED_date_set_HEAP).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    End If
End Function

Function WriteSWC(jr, i, wsDest, wsSour, ProgName, RecordType)
    If ProgName = "ROSA" Then
        wsDest.Cells(jr, 7).Value = "SITE WORK COMPLETE"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_ROSA).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    ElseIf ProgName = "HEAP" Then
        wsDest.Cells(jr, 7).Value = "SITE WORK COMPLETE"
        wsDest.Cells(jr, 8).Value = Left(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP).Value, 8)
        wsDest.Cells(jr, 9).Value = Right(wsSour.Cells(i, NexantEnrollments.SITE_WORK_COMPLETE_date_set_HEAP).Value, 6)
        wsDest.Cells(jr, 11).Value = "" 'historical only for PENDING status
        wsDest.Cells(jr, 12).Value = "" 'historical only for PENDING; required only for PENDING, ON-HOLD, CANCELLED
        WriteStaticColumns jr, i, wsDest, wsSour, ProgName, RecordType
    End If
End Function

