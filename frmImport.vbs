VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Import"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EnrollmentFirstDataLine As Integer

Private Sub UserForm_Initialize()
    EnrollmentFirstDataLine = 11
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub cmdImport_Click()
    If opbEnrollment Then Call importfile("Enrollment")
    If opbUsage Then Call importfile("Usage")
End Sub

Private Sub UserForm_Terminate()
    frmServiceCenter.Show
End Sub

Sub importfile(ByVal importfile As String)
    Dim FileNum As Integer
    Dim DataLine As String
    Dim str() As String
    Dim LineNum As Integer
    
    FileNum = FreeFile()
    filetoopen = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    If filetoopen = False Then
        Exit Sub
    End If
    
    LineNum = 0
    Open filetoopen For Input As #FileNum
    While Not EOF(FileNum)
        LineNum = LineNum + 1
        Line Input #FileNum, DataLine
        ReDim Preserve str(0 To LineNum)
        str(UBound(str)) = DataLine
    Wend
    Close #FileNum

    'parse records
    Dim flgImport As Boolean
    flgImport = False
    For k = 2 To LineNum - 1
        Select Case importfile
            Case "Enrollment"
                Call parseenrollment(str(k), flgImport)
            Case "Usage"
                Call parseusage(str(k), flgImport)
        End Select
    Next k

    If flgImport Then
        MsgBox "Import found errors! Please check the OUT enrollment file."
    Else
        MsgBox "Import is completed."
    End If
    
End Sub

Sub parseenrollment(ByVal str1 As String, ByRef flg As Boolean)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim lastROSA As Integer
    Dim lastHEAP As Integer
    Dim lastrow As Integer
    lastROSA = Worksheets("Enrollments").Range("B" & Rows.Count).End(xlUp).Row + 1
    lastHEAP = Worksheets("Enrollments").Range("C" & Rows.Count).End(xlUp).Row + 1
    lastrow = WorksheetFunction.Max(lastROSA, lastHEAP)
    
    x = Split(str1, ",")
    Dim enrollmentID As String
    enrollmentID = x(LGEEnrollments.Enrollment_ID)
    
    Select Case x(LGEEnrollments.Transaction_Type) 'Transaction Type
        Case "N"
            If lastrow > EnrollmentFirstDataLine Then
                For i = 11 To lastrow
                    existingID = Worksheets("Enrollments").Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value 'ROSA ID
                    If existingID = enrollmentID Then
                        flg = True
                        MsgBox "The enrollment ID exists. Please check the Enrollment ID: " + CStr(enrollmentID)
                        Exit Sub
                    End If
                Next i
            End If
        Case "U"
            If lastrow > EnrollmentFirstDataLine Then
                For i = EnrollmentFirstDataLine To lastrow
                    existingID = Worksheets("Enrollments").Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value 'ROSA ID
                    If existingID = enrollmentID Then
                        lastrow = i
                        Exit For
                    End If
                Next i
            End If
            flg = True
            MsgBox "The existing enrollment ID " + CStr(enrollmentID) + " is not found. Please check the Enrollment ID and the Transaction Type " + x(LGEEnrollments.Transaction_Type) + "."
            Exit Sub
        Case Else
            flg = True
            MsgBox "Incorrect enrollment OUT file. Please check the Transaction Type " + x(LGEEnrollments.Transaction_Type) + "."
            Exit Sub
    End Select
    
    'Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Record_Type)
    'Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Transaction_Type)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Enrollment_ID_ROSA).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Enrollment_ID_ROSA).Value = x(LGEEnrollments.Enrollment_ID)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Premise_ID).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Premise_ID).Value = x(LGEEnrollments.Premise_ID)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Company_Code).Value = x(LGEEnrollments.Company_Code)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Account_Number).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Account_Number).Value = x(LGEEnrollments.Customer_Account)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Main_Account_Flag).Value = x(LGEEnrollments.Main_Account_Flag)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Business_Partner_Number_ROSA).Value = x(LGEEnrollments.Business_Partner_Number)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Premise_Type).Value = x(LGEEnrollments.Premise_Type)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_customer_name).Value = x(LGEEnrollments.Service_customer_name)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Customer_Home_Phone).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Customer_Home_Phone).Value = x(LGEEnrollments.Customer_Home_Phone)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Customer_mobile_phone).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Customer_mobile_phone).Value = x(LGEEnrollments.Customer_mobile_phone)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_Street_Address).Value = x(LGEEnrollments.Service_Street_Address)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_City).Value = x(LGEEnrollments.Service_City)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_State).Value = x(LGEEnrollments.Service_State)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_Zipcode).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Service_Zipcode).Value = x(LGEEnrollments.Service_Zipcode)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Mailing_Street_Address).Value = x(LGEEnrollments.Mailing_Street_Address)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Mailing_City).Value = x(LGEEnrollments.Mailing_City)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Mailing_State).Value = x(LGEEnrollments.Mailing_State)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Mailing_Zipcode).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Mailing_Zipcode).Value = x(LGEEnrollments.Mailing_Zipcode)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Customer_Email).Value = x(LGEEnrollments.Customer_Email)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Contact_Name).Value = x(LGEEnrollments.Contact_Name)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Address).Value = x(LGEEnrollments.Primary_Contact_Address)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Address_City).Value = x(LGEEnrollments.Primary_Contact_Address_City)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Address_State).Value = x(LGEEnrollments.Primary_Contact_Address_State)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Address_Zip).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Address_Zip).Value = x(LGEEnrollments.Primary_Contact_Address_Zip)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Email).Value = x(LGEEnrollments.Primary_Contact_Email)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Phone).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_Phone).Value = x(LGEEnrollments.Primary_Contact_Phone)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_phone_extension).Value = x(LGEEnrollments.Primary_Contact_phone_extension)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Primary_Contact_mobile_phone).Value = x(LGEEnrollments.Primary_Contact_mobile_phone)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Nbr_Building_Occupants)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value = x(LGEEnrollments.Total_conditioned_square_footage)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Send_Report_to_Primary_Contact)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = x(LGEEnrollments.Dog_or_Cat_Flag)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Ownership_Type_ROSA).Value = x(LGEEnrollments.Ownership_Type)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Reason_for_audit)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Short_Program_Name)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Verification_Class)
'    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.).Value=x(LGEEnrollments.Baseline_Tier1_vendor_work_order_number)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Name).Value = x(LGEEnrollments.Remit_to_Contact_Name)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Address).Value = x(LGEEnrollments.Remit_to_Contact_Address)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Address_City).Value = x(LGEEnrollments.Remit_to_Contact_Address_City)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Address_State).Value = x(LGEEnrollments.Remit_to_Contact_Address_State)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Address_Zip).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Address_Zip).Value = x(LGEEnrollments.Remit_to_Contact_Address_Zip)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Email).Value = x(LGEEnrollments.Remit_to_Contact_Email)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Phone).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_Phone).Value = x(LGEEnrollments.Remit_to_Contact_Phone)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_phone_extension).Value = x(LGEEnrollments.Remit_to_Contact_phone_extension)
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_mobile_phone).NumberFormat = "@"
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Remit_to_Contact_mobile_phone).Value = x(LGEEnrollments.Remit_to_Contact_mobile_phone)

    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Status_Date_ROSA) = Format(Now(), "YYYYMMDD")
    Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Status_Time_ROSA) = Format(Now(), "YYYYMMDD") + ":" + Format(Now(), "HHMMSS")
    Select Case x(LGEEnrollments.Enrollment_ID) 'Transaction Type
        Case "N"
            Worksheets("Enrollments").Cells(lastrow, NexantEnrollments.Status_ROSA) = "RECEIVED AT VENDOR"
        Case "U"
        '  ???
    End Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub parseusage(ByVal str As String, ByRef flg As Boolean)
    'field mapping?
End Sub
