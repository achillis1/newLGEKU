VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdmin 
   Caption         =   "Admin"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "frmAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LastRow As Integer

Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub

Private Sub cmdHeap_Click()

    If lstEnrollments.Text = "" Then
        MsgBox "Please select an enrollment."
        Exit Sub
    Else
        Me.Hide
        currentEnrollment = Left(lstEnrollments.Text, Len(lstEnrollments.Text) - 5)
    End If
    HEAP_Scheduling_Status_Info.Show
End Sub

Private Sub cmdHeapContact_Click()
Me.Hide
List_Contact_Attempts_HEAP.Show
End Sub

Private Sub cmdInfo_Click()
Me.Hide
Information_Form.Show
End Sub

Private Sub cmdReset_Click()
    txtEnrollment.Text = ""
    cmdROSA.Enabled = False
    cmdHeap.Enabled = False
    cmdInfo.Enabled = False
    cmdMeasure.Enabled = False
    cmdUsage.Enabled = False
    cmdContextual.Enabled = False
    lstEnrollments.Clear
    
    For i = EnrollmentFirstDataLine To LastRow
        ROSAID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value
        HEAPID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value
        If ROSAID = "" And HEAPID <> "" Then existingID = HEAPID + "-HEAP"
        If ROSAID <> "" And HEAPID = "" Then existingID = ROSAID + "-ROSA"
        lstEnrollments.AddItem (existingID)
    Next i
End Sub

Private Sub cmdROSA_Click()

    If lstEnrollments.Text = "" Then
        MsgBox "Please select an enrollment."
        Exit Sub
    Else
        Me.Hide
        currentEnrollment = Left(lstEnrollments.Text, Len(lstEnrollments.Text) - 5)
    End If

    ROSA_Scheduling_Status_Info.Show
End Sub

Private Sub cmdRosaContact_Click()
Me.Hide
List_Contact_Attempts_ROSA.Show
End Sub

Private Sub cmdSearch_Click()
    Dim ROSAID As Long
    Dim HEAPID As Long
    Dim EnrollmentID As Long
    
    If txtEnrollment = "" Then
        MsgBox "Please enter an enrollment ID. Thanks"
    Else
        EnrollmentID = CLng(txtEnrollment.Text)
        Dim flag As Boolean
        flag = False
        For i = EnrollmentFirstDataLine To LastRow
            ROSAID = CLng(Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value)
            HEAPID = CLng(Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value)
'            If ROSAID = "" And HEAPID <> "" Then existingID = CLng(HEAPID)
'            If ROSAID <> "" And HEAPID = "" Then existingID = CLng(ROSAID)
            
            If EnrollmentID = ROSAID Or EnrollmentID = HEAPID Then
                flag = True
                lstEnrollments.Clear
                If EnrollmentID = ROSAID Then
                    lstEnrollments.AddItem (Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value + "-ROSA")
                    cmdROSA.Enabled = True
                    cmdHeap.Enabled = False
                End If
                If EnrollmentID = HEAPID Then
                    lstEnrollments.AddItem (Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value + "-HEAP")
                    cmdHeap.Enabled = True
                    cmdROSA.Enabled = False
                End If
                cmdInfo.Enabled = True
                cmdMeasure.Enabled = True
            End If
        Next i
        If Not flag Then MsgBox "The enrollment ID is not found. Please enter a valid enrollment ID. Thanks"
    End If
End Sub

Private Sub cmdUsage_Click()
Me.Hide
Utility_Data.Show
End Sub

Private Sub lstEnrollments_Click()
    Dim bRH As Boolean
    Dim rh As String
    rh = Right(lstEnrollments.Text, 4)
    
    If rh <> "" Then
        currentEnrollment = Left(lstEnrollments.Text, Len(lstEnrollments.Text) - 5)
        cmdUsage.Enabled = True
        cmdMeasure.Enabled = True
        cmdContextual.Enabled = True
        cmdInfo.Enabled = True
        Select Case rh
            Case "ROSA"
                cmdHeap.Enabled = False
                cmdROSA.Enabled = True

            Case "HEAP"
                cmdROSA.Enabled = False
                cmdHeap.Enabled = True

            Case Else
        End Select
    End If

    
End Sub

Private Sub UserForm_Activate()
    Dim lastROSA As Integer
    Dim lastHEAP As Integer

    lastROSA = Worksheets(ImportSheetName).Range("B" & Rows.Count).End(xlUp).row
    lastHEAP = Worksheets(ImportSheetName).Range("C" & Rows.Count).End(xlUp).row
    LastRow = WorksheetFunction.Max(lastROSA, lastHEAP)
    
    For i = EnrollmentFirstDataLine To LastRow
        ROSAID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value
        HEAPID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value
        If ROSAID = "" And HEAPID <> "" Then existingID = HEAPID + "-HEAP"
        If ROSAID <> "" And HEAPID = "" Then existingID = ROSAID + "-ROSA"
        lstEnrollments.AddItem (existingID)
    Next i
    
    cmdInfo.Enabled = False
    cmdROSA.Enabled = False
    cmdHeap.Enabled = False
    cmdMeasure.Enabled = False
    cmdUsage.Enabled = False
    cmdContextual.Enabled = False
End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

