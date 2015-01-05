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
Private ImportSheetName As String
Private PMSheetName As String
Private InboundLastReadCol As Integer
    
Private Sub UserForm_Initialize()
    EnrollmentFirstDataLine = 11
    ImportSheetName = "Enrollments"
    PMSheetName = "PM"
    InboundLastReadCol = 5
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub cmdImport_Click()
    Call importfile
End Sub

Private Sub UserForm_Terminate()
    frmServiceCenter.Show vbModeless
End Sub

Private Sub importfile()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
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
    Dim ShortProgramName As String
    Dim OUTReportType As String
    Dim x() As String
    
    x1 = Split(str(1), ",")
    ShortProgramName = x1(LGEHeader.Short_Program_Name)
    OUTReportType = x1(LGEHeader.Miscellaneous)
    
    If Not (ShortProgramName = "ROSA" Or ShortProgramName = "HEAP") Then
        MsgBox "Incorrect OUT file. Please check the Short Program Name field, " + ShortProgramName + "."
        Exit Sub
    End If
    
    If Not (OUTReportType = "OUTBOUND ENROLLMENT" Or OUTReportType = "OUTBOUND USAGE") Then
        MsgBox "Incorrect OUT file. Please check the Miscellaneous field, " + OUTReportType + "."
        Exit Sub
    End If

    For k = 2 To LineNum - 1
    
        Dim lastROSA As Integer
        Dim lastHEAP As Integer
        Dim lastrow As Integer
        lastROSA = Worksheets(ImportSheetName).Range("B" & Rows.Count).End(xlUp).Row
        lastHEAP = Worksheets(ImportSheetName).Range("C" & Rows.Count).End(xlUp).Row
        lastrow = WorksheetFunction.Max(lastROSA, lastHEAP)
        
        If lastrow < EnrollmentFirstDataLine - 1 Then
            MsgBox "The " + ImportSheetName + " data has errors. Please contact the developer."
            Exit Sub
        End If
    
        x = Split(str(k), ",")
        Dim enrollmentID As String
        enrollmentID = x(LGEEnrollments.Enrollment_ID) 'or LGEUsage.Enrollment_ID
        
        Dim TransactionType As String
        TransactionType = x(LGEEnrollments.Transaction_Type) ' or LGEUsage.Transaction_Type
        
        Dim RosaID As String
        Dim HeapID As String
        Dim ir As Integer
        ir = 0
        If lastrow = EnrollmentFirstDataLine - 1 Then
            ir = EnrollmentFirstDataLine
        Else
            For i = EnrollmentFirstDataLine To lastrow
                RosaID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value
                HeapID = Worksheets(ImportSheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value
                If RosaID = "" And HeapID <> "" Then existingID = HeapID
                If RosaID <> "" And HeapID = "" Then existingID = RosaID
                
                If existingID = enrollmentID Then
                    Select Case TransactionType
                        Case "N"
                            If OUTReportType = "OUTBOUND ENROLLMENT" Then '  Or OUTReportType = "OUTBOUND USAGE"
                                MsgBox "The enrollment ID exists. Please check the Enrollment ID: " + CStr(enrollmentID)
                                Exit Sub
                            Else
                                ir = i
                                Exit For
                            End If
                        Case "U"
                            ir = i
                            Exit For
                    End Select
                End If
            Next i
                                
            Select Case TransactionType 'N or U
                Case "N"
                    If ir = 0 Then
                        If OUTReportType = "OUTBOUND ENROLLMENT" Then
                            ir = lastrow + 1
                        Else
                            MsgBox "The enrollment ID " + CStr(enrollmentID) + " is not found. Please import the enrollment OUT file first."
                            Exit Sub
                        End If
                    End If
                Case "U"
                    If ir = 0 Then
                        MsgBox "The enrollment ID " + CStr(enrollmentID) + " is not found. Please check the Enrollment ID and the Transaction Type " + x(LGEEnrollments.Transaction_Type) + "."
                        Exit Sub
                    End If
            End Select
        End If
    
        If OUTReportType = "OUTBOUND ENROLLMENT" Then Call parseenrollment(x, ir, ShortProgramName)
        If OUTReportType = "OUTBOUND USAGE" Then Call parseusage(x, ir, ShortProgramName)
    Next k

    MsgBox "Import is completed."
    frmImport.Hide
    frmServiceCenter.Show vbModelless
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub parseenrollment(ByRef x() As String, ByVal ir As Integer, ByVal pn As String)

    'Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.).Value=x(LGEEnrollments.Record_Type)
    'Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.).Value=x(LGEEnrollments.Transaction_Type)
    If pn = "ROSA" Then
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Enrollment_ID_ROSA).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Enrollment_ID_ROSA).Value = x(LGEEnrollments.Enrollment_ID)
    Else
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Enrollment_ID_HEAP).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Enrollment_ID_HEAP).Value = x(LGEEnrollments.Enrollment_ID)
    End If
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Premise_ID).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Premise_ID).Value = x(LGEEnrollments.Premise_ID)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Company_Code).Value = x(LGEEnrollments.Company_Code)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Account_Number).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Account_Number).Value = x(LGEEnrollments.Customer_Account)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Main_Account_Flag).Value = x(LGEEnrollments.Main_Account_Flag)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Business_Partner_Number_ROSA).Value = x(LGEEnrollments.Business_Partner_Number)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Premise_Type).Value = x(LGEEnrollments.Premise_Type)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_customer_name).Value = x(LGEEnrollments.Service_customer_name)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Customer_Home_Phone).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Customer_Home_Phone).Value = x(LGEEnrollments.Customer_Home_Phone)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Customer_mobile_phone).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Customer_mobile_phone).Value = x(LGEEnrollments.Customer_mobile_phone)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_Street_Address).Value = x(LGEEnrollments.Service_Street_Address)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_City).Value = x(LGEEnrollments.Service_City)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_State).Value = x(LGEEnrollments.Service_State)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_Zipcode).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Service_Zipcode).Value = x(LGEEnrollments.Service_Zipcode)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Mailing_Street_Address).Value = x(LGEEnrollments.Mailing_Street_Address)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Mailing_City).Value = x(LGEEnrollments.Mailing_City)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Mailing_State).Value = x(LGEEnrollments.Mailing_State)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Mailing_Zipcode).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Mailing_Zipcode).Value = x(LGEEnrollments.Mailing_Zipcode)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Customer_Email).Value = x(LGEEnrollments.Customer_Email)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Contact_Name).Value = x(LGEEnrollments.Contact_Name)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Address).Value = x(LGEEnrollments.Primary_Contact_Address)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Address_City).Value = x(LGEEnrollments.Primary_Contact_Address_City)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Address_State).Value = x(LGEEnrollments.Primary_Contact_Address_State)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Address_Zip).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Address_Zip).Value = x(LGEEnrollments.Primary_Contact_Address_Zip)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Email).Value = x(LGEEnrollments.Primary_Contact_Email)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Phone).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_Phone).Value = x(LGEEnrollments.Primary_Contact_Phone)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_phone_extension).Value = x(LGEEnrollments.Primary_Contact_phone_extension)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Primary_Contact_mobile_phone).Value = x(LGEEnrollments.Primary_Contact_mobile_phone)
    If pn = "ROSA" Then
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Building_occupancy_count_ROSA).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Building_occupancy_count_ROSA).Value = x(LGEEnrollments.Nbr_Building_Occupants)
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value = x(LGEEnrollments.Total_conditioned_square_footage)
    Else
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Building_occupancy_count_HEAP).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Building_occupancy_count_HEAP).Value = x(LGEEnrollments.Nbr_Building_Occupants)
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value = x(LGEEnrollments.Total_conditioned_square_footage)
    End If
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Send_Report_to_Primary_Contact).Value = x(LGEEnrollments.Send_Report_to_Primary_Contact)
    If pn = "ROSA" Then
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = x(LGEEnrollments.Dog_or_Cat_Flag)
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Ownership_Type_ROSA).Value = x(LGEEnrollments.Ownership_Type)
    Else
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value = x(LGEEnrollments.Dog_or_Cat_Flag)
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Ownership_Type_HEAP).Value = x(LGEEnrollments.Ownership_Type)
    End If
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Reason_for_audit).Value = x(LGEEnrollments.Reason_for_audit)
'    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.).Value=programname
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Verification_Class).Value = x(LGEEnrollments.Verification_Class)
    If pn = "ROSA" Then
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.WO_Number_ROSA).Value = x(LGEEnrollments.Baseline_Tier1_vendor_work_order_number)
    Else
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.WO_Number_HEAP).Value = x(LGEEnrollments.Baseline_Tier1_vendor_work_order_number)
    End If
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Name).Value = x(LGEEnrollments.Remit_to_Contact_Name)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Address).Value = x(LGEEnrollments.Remit_to_Contact_Address)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Address_City).Value = x(LGEEnrollments.Remit_to_Contact_Address_City)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Address_State).Value = x(LGEEnrollments.Remit_to_Contact_Address_State)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Address_Zip).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Address_Zip).Value = x(LGEEnrollments.Remit_to_Contact_Address_Zip)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Email).Value = x(LGEEnrollments.Remit_to_Contact_Email)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Phone).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_Phone).Value = x(LGEEnrollments.Remit_to_Contact_Phone)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_phone_extension).Value = x(LGEEnrollments.Remit_to_Contact_phone_extension)
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_mobile_phone).NumberFormat = "@"
    Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Remit_to_Contact_mobile_phone).Value = x(LGEEnrollments.Remit_to_Contact_mobile_phone)

    If pn = "ROSA" Then
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Date_ROSA).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Date_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD")
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Time_ROSA).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Time_ROSA) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    Else
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Date_HEAP).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Date_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD")
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Time_HEAP).NumberFormat = "@"
        Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_Time_HEAP) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    End If
    
    If x(LGEEnrollments.Transaction_Type) = "N" Then
        If pn = "ROSA" Then
            Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_ROSA) = "RECEIVED AT VENDOR"
        Else
            Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Status_HEAP) = "RECEIVED AT VENDOR"
        End If
    End If
    
    If pn = "ROSA" Then
        Worksheets(PMSheetName).Cells(PMROSAEnrollRow, InboundLastReadCol).NumberFormat = "@"
        Worksheets(PMSheetName).Cells(PMROSAEnrollRow, InboundLastReadCol) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    Else
        Worksheets(PMSheetName).Cells(PMHEAPEnrollRow, InboundLastReadCol).NumberFormat = "@"
        Worksheets(PMSheetName).Cells(PMHEAPEnrollRow, InboundLastReadCol) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    End If
    
End Sub

Private Sub parseusage(ByRef x() As String, ByVal ir As Integer, ByVal pn As String)
    
    Dim dt As String
    Dim Year As String
    Dim Month As String
    Dim Day As String
    Dim ratecategory As String
    Dim il As Integer
    
    dt = x(LGEUsage.Billing_Date)
    Year = Mid(dt, 1, 4)
    Month = Mid(dt, 5, 2)
    Day = Mid(dt, 7, 2)
    ratecategory = x(LGEUsage.Rate_Category_Text)
    
    il = InStr(1, ratecategory, "Gas")
    
    If il = 0 Then 'Electric
        Select Case Month
            Case "01"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jan_Service_Division) = x(LGEUsage.Service_Division)
            Case "02"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Feb_Service_Division) = x(LGEUsage.Service_Division)
            Case "03"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Mar_Service_Division) = x(LGEUsage.Service_Division)
            Case "04"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Apr_Service_Division) = x(LGEUsage.Service_Division)
            Case "05"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_May_Service_Division) = x(LGEUsage.Service_Division)
            Case "06"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jun_Service_Division) = x(LGEUsage.Service_Division)
            Case "07"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Jul_Service_Division) = x(LGEUsage.Service_Division)
            Case "08"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Aug_Service_Division) = x(LGEUsage.Service_Division)
            Case "09"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Sep_Service_Division) = x(LGEUsage.Service_Division)
            Case "10"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Oct_Service_Division) = x(LGEUsage.Service_Division)
            Case "11"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Nov_Service_Division) = x(LGEUsage.Service_Division)
            Case "12"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_PF_On_Peak_Electric) = x(LGEUsage.PF_On_Peak_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Power_Factor_on_adjustment_Electric) = x(LGEUsage.Power_Factor_on_adjustment_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_KW_Billed_on_Demand_Electric) = x(LGEUsage.KW_Billed_on_Demand_Electric)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Electricity_Dec_Service_Division) = x(LGEUsage.Service_Division)
        End Select
    Else 'Gas
        Select Case Month
            Case "01"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jan_Service_Division) = x(LGEUsage.Service_Division)
            Case "02"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Feb_Service_Division) = x(LGEUsage.Service_Division)
            Case "03"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Mar_Service_Division) = x(LGEUsage.Service_Division)
            Case "04"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Apr_Service_Division) = x(LGEUsage.Service_Division)
            Case "05"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_May_Service_Division) = x(LGEUsage.Service_Division)
            Case "06"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jun_Service_Division) = x(LGEUsage.Service_Division)
            Case "07"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Jul_Service_Division) = x(LGEUsage.Service_Division)
            Case "08"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Aug_Service_Division) = x(LGEUsage.Service_Division)
            Case "09"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Sep_Service_Division) = x(LGEUsage.Service_Division)
            Case "10"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Oct_Service_Division) = x(LGEUsage.Service_Division)
            Case "11"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Nov_Service_Division) = x(LGEUsage.Service_Division)
            Case "12"
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Meter_Number) = x(LGEUsage.Meter_Number)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Rate_Category_Text) = x(LGEUsage.Rate_Category_Text)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Billing_Date) = x(LGEUsage.Billing_Date)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Billed_Amount) = x(LGEUsage.Billed_Amount)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Taxes_and_Fees) = x(LGEUsage.Taxes_and_Fees)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Energy_Consumption) = x(LGEUsage.Energy_Consumption)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Average_Temperature) = x(LGEUsage.Average_Temperature)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Heating_degree_days) = x(LGEUsage.Heating_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Cooling_degree_days) = x(LGEUsage.Cooling_degree_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_No_of_billing_days) = x(LGEUsage.No_of_billing_days)
                Worksheets(ImportSheetName).Cells(ir, NexantEnrollments.Usage_Gas_Dec_Service_Division) = x(LGEUsage.Service_Division)
        End Select
    
    End If
    
    If pn = "ROSA" Then
        Worksheets(PMSheetName).Cells(PMROSAUsageRow, InboundLastReadCol).NumberFormat = "@"
        Worksheets(PMSheetName).Cells(PMROSAUsageRow, InboundLastReadCol) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    Else
        Worksheets(PMSheetName).Cells(PMHEAPUsageRow, InboundLastReadCol).NumberFormat = "@"
        Worksheets(PMSheetName).Cells(PMHEAPUsageRow, InboundLastReadCol) = Format(LocalTimeToET(Now()), "YYYYMMDD") + ":" + Format(LocalTimeToET(Now()), "HHMMSS")
    End If

End Sub

