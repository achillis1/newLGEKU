VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Utility_Data 
   Caption         =   "Utility Data"
   ClientHeight    =   12780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14925
   OleObjectBlob   =   "Utility_Data.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Utility_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
Me.Hide
frmAdmin.Show vbModeless

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub



Private Sub CommandButton1_Click()
Dim TimeandDate As String
Dim Result As Double
Dim Enrollment_ID_ROSA As Double
Dim HeadingOffset






End Sub



Private Sub CommandButton2_Click()
Call UserForm_Initialize

End Sub

Private Sub CommandButton3_Click()
Me.Hide
frmAdmin.Show

End Sub

Private Sub Frame17_Click()

End Sub

Private Sub UserForm_Initialize()
Dim TimeandDate As String
Dim LastRow As String
Dim Results As Double
Dim Enrollment_ID_ROSA As String
Dim ColumnHeadings(46) As String
'Dim currentEnrollment As String
Dim ColumnValues(46) As String
Dim counter As Double
Dim HeadingOffset

HeadingOffset = 10

Enrollment_ID_ROSA = currentEnrollment

If Enrollment_ID_ROSA <> "" Then
LastRow = Cells(Rows.Count, 2).End(xlUp).row

'Results = Application.Match(Enrollment_ID_ROSA, Sheets("Enrollments").Range("B1", "B" & Range("B" & Rows.Count).End(xlUp).Row), 0)
Results = Application.Match(Enrollment_ID_ROSA, Sheets("Enrollments").Range(Cells(11, NexantEnrollments.Enrollment_ID_ROSA), Cells(LastRow, NexantEnrollments.Enrollment_ID_ROSA))) + HeadingOffset

Else

End If

If Results <> 0 Then
Elec_Apr_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Average_Temperature)
Elec_Apr_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Billed_Amount)
Elec_Apr_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Billing_Date)
Elec_Apr_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Cooling_degree_days)
Elec_Apr_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Energy_Consumption)
Elec_Apr_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Heating_degree_days)
Elec_Apr_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_KW_Billed_on_Demand_Electric)
Elec_Apr_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Meter_Number)
Elec_Apr_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_No_of_billing_days)
Elec_Apr_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_PF_On_Peak_Electric)
Elec_Apr_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Power_Factor_on_adjustment_Electric)
Elec_Apr_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Rate_Category_Text)
Elec_Apr_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Service_Division)
Elec_Apr_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Apr_Taxes_and_Fees)
Elec_Aug_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Average_Temperature)
Elec_Aug_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Billed_Amount)
Elec_Aug_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Billing_Date)
Elec_Aug_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Cooling_degree_days)
Elec_Aug_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Energy_Consumption)
Elec_Aug_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Heating_degree_days)
Elec_Aug_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_KW_Billed_on_Demand_Electric)
Elec_Aug_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Meter_Number)
Elec_Aug_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_No_of_billing_days)
Elec_Aug_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_PF_On_Peak_Electric)
Elec_Aug_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Power_Factor_on_adjustment_Electric)
Elec_Aug_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Rate_Category_Text)
Elec_Aug_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Service_Division)
Elec_Aug_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Aug_Taxes_and_Fees)
Elec_Dec_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Average_Temperature)
Elec_Dec_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Billed_Amount)
Elec_Dec_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Billing_Date)
Elec_Dec_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Cooling_degree_days)
Elec_Dec_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Energy_Consumption)
Elec_Dec_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Heating_degree_days)
Elec_Dec_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_KW_Billed_on_Demand_Electric)
Elec_Dec_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Meter_Number)
Elec_Dec_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_No_of_billing_days)
Elec_Dec_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_PF_On_Peak_Electric)
Elec_Dec_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Power_Factor_on_adjustment_Electric)
Elec_Dec_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Rate_Category_Text)
Elec_Dec_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Service_Division)
Elec_Dec_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Dec_Taxes_and_Fees)
Elec_Feb_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Average_Temperature)
Elec_Feb_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Billed_Amount)
Elec_Feb_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Billing_Date)
Elec_Feb_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Cooling_degree_days)
Elec_Feb_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Energy_Consumption)
Elec_Feb_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Heating_degree_days)
Elec_Feb_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_KW_Billed_on_Demand_Electric)
Elec_Feb_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Meter_Number)
Elec_Feb_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_No_of_billing_days)
Elec_Feb_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_PF_On_Peak_Electric)
Elec_Feb_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Power_Factor_on_adjustment_Electric)
Elec_Feb_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Rate_Category_Text)
Elec_Feb_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Service_Division)
Elec_Feb_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Feb_Taxes_and_Fees)
Elec_Jan_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Average_Temperature)
Elec_Jan_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Billed_Amount)
Elec_Jan_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Billing_Date)
Elec_Jan_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Cooling_degree_days)
Elec_Jan_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Energy_Consumption)
Elec_Jan_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Heating_degree_days)
Elec_Jan_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_KW_Billed_on_Demand_Electric)
Elec_Jan_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Meter_Number)
Elec_Jan_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_No_of_billing_days)
Elec_Jan_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_PF_On_Peak_Electric)
Elec_Jan_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Power_Factor_on_adjustment_Electric)
Elec_Jan_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Rate_Category_Text)
Elec_Jan_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Service_Division)
Elec_Jan_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jan_Taxes_and_Fees)
Elec_Jul_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Average_Temperature)
Elec_Jul_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Billed_Amount)
Elec_Jul_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Billing_Date)
Elec_Jul_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Cooling_degree_days)
Elec_Jul_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Energy_Consumption)
Elec_Jul_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Heating_degree_days)
Elec_Jul_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_KW_Billed_on_Demand_Electric)
Elec_Jul_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Meter_Number)
Elec_Jul_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_No_of_billing_days)
Elec_Jul_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_PF_On_Peak_Electric)
Elec_Jul_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Power_Factor_on_adjustment_Electric)
Elec_Jul_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Rate_Category_Text)
Elec_Jul_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Service_Division)
Elec_Jul_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jul_Taxes_and_Fees)
Elec_Jun_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Average_Temperature)
Elec_Jun_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Billed_Amount)
Elec_Jun_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Billing_Date)
Elec_Jun_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Cooling_degree_days)
Elec_Jun_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Energy_Consumption)
Elec_Jun_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Heating_degree_days)
Elec_Jun_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_KW_Billed_on_Demand_Electric)
Elec_Jun_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Meter_Number)
Elec_Jun_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_No_of_billing_days)
Elec_Jun_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_PF_On_Peak_Electric)
Elec_Jun_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Power_Factor_on_adjustment_Electric)
Elec_Jun_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Rate_Category_Text)
Elec_Jun_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Service_Division)
Elec_Jun_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Jun_Taxes_and_Fees)
Elec_Mar_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Average_Temperature)
Elec_Mar_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Billed_Amount)
Elec_Mar_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Billing_Date)
Elec_Mar_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Cooling_degree_days)
Elec_Mar_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Energy_Consumption)
Elec_Mar_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Heating_degree_days)
Elec_Mar_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_KW_Billed_on_Demand_Electric)
Elec_Mar_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Meter_Number)
Elec_Mar_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_No_of_billing_days)
Elec_Mar_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_PF_On_Peak_Electric)
Elec_Mar_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Power_Factor_on_adjustment_Electric)
Elec_Mar_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Rate_Category_Text)
Elec_Mar_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Service_Division)
Elec_Mar_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Mar_Taxes_and_Fees)
Elec_May_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Average_Temperature)
Elec_May_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Billed_Amount)
Elec_May_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Billing_Date)
Elec_May_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Cooling_degree_days)
Elec_May_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Energy_Consumption)
Elec_May_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Heating_degree_days)
Elec_May_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_KW_Billed_on_Demand_Electric)
Elec_May_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Meter_Number)
Elec_May_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_No_of_billing_days)
Elec_May_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_PF_On_Peak_Electric)
Elec_May_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Power_Factor_on_adjustment_Electric)
Elec_May_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Rate_Category_Text)
Elec_May_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Service_Division)
Elec_May_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_May_Taxes_and_Fees)
Elec_Nov_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Average_Temperature)
Elec_Nov_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Billed_Amount)
Elec_Nov_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Billing_Date)
Elec_Nov_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Cooling_degree_days)
Elec_Nov_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Energy_Consumption)
Elec_Nov_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Heating_degree_days)
Elec_Nov_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_KW_Billed_on_Demand_Electric)
Elec_Nov_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Meter_Number)
Elec_Nov_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_No_of_billing_days)
Elec_Nov_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_PF_On_Peak_Electric)
Elec_Nov_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Power_Factor_on_adjustment_Electric)
Elec_Nov_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Rate_Category_Text)
Elec_Nov_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Service_Division)
Elec_Nov_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Nov_Taxes_and_Fees)
Elec_Oct_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Average_Temperature)
Elec_Oct_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Billed_Amount)
Elec_Oct_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Billing_Date)
Elec_Oct_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Cooling_degree_days)
Elec_Oct_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Energy_Consumption)
Elec_Oct_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Heating_degree_days)
Elec_Oct_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_KW_Billed_on_Demand_Electric)
Elec_Oct_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Meter_Number)
Elec_Oct_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_No_of_billing_days)
Elec_Oct_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_PF_On_Peak_Electric)
Elec_Oct_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Power_Factor_on_adjustment_Electric)
Elec_Oct_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Rate_Category_Text)
Elec_Oct_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Service_Division)
Elec_Oct_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Oct_Taxes_and_Fees)
Elec_Sep_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Average_Temperature)
Elec_Sep_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Billed_Amount)
Elec_Sep_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Billing_Date)
Elec_Sep_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Cooling_degree_days)
Elec_Sep_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Energy_Consumption)
Elec_Sep_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Heating_degree_days)
Elec_Sep_KW_Billed_on_Demand.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_KW_Billed_on_Demand_Electric)
Elec_Sep_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Meter_Number)
Elec_Sep_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_No_of_billing_days)
Elec_Sep_PF_On_Peak.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_PF_On_Peak_Electric)
Elec_Sep_Power_Factor_on_adjustment.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Power_Factor_on_adjustment_Electric)
Elec_Sep_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Rate_Category_Text)
Elec_Sep_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Service_Division)
Elec_Sep_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Electricity_Sep_Taxes_and_Fees)
Gas_Apr_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Average_Temperature)
Gas_Apr_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Billed_Amount)
Gas_Apr_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Billing_Date)
Gas_Apr_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Cooling_degree_days)
Gas_Apr_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Energy_Consumption)
Gas_Apr_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Heating_degree_days)
Gas_Apr_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Meter_Number)
Gas_Apr_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_No_of_billing_days)
Gas_Apr_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Rate_Category_Text)
Gas_Apr_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Service_Division)
Gas_Apr_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Apr_Taxes_and_Fees)
Gas_Aug_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Average_Temperature)
Gas_Aug_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Billed_Amount)
Gas_Aug_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Billing_Date)
Gas_Aug_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Cooling_degree_days)
Gas_Aug_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Energy_Consumption)
Gas_Aug_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Heating_degree_days)
Gas_Aug_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Meter_Number)
Gas_Aug_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_No_of_billing_days)
Gas_Aug_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Rate_Category_Text)
Gas_Aug_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Service_Division)
Gas_Aug_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Aug_Taxes_and_Fees)
Gas_Dec_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Average_Temperature)
Gas_Dec_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Billed_Amount)
Gas_Dec_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Billing_Date)
Gas_Dec_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Cooling_degree_days)
Gas_Dec_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Energy_Consumption)
Gas_Dec_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Heating_degree_days)
Gas_Dec_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Meter_Number)
Gas_Dec_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_No_of_billing_days)
Gas_Dec_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Rate_Category_Text)
Gas_Dec_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Service_Division)
Gas_Dec_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Dec_Taxes_and_Fees)
Gas_Feb_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Average_Temperature)
Gas_Feb_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Billed_Amount)
Gas_Feb_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Billing_Date)
Gas_Feb_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Cooling_degree_days)
Gas_Feb_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Energy_Consumption)
Gas_Feb_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Heating_degree_days)
Gas_Feb_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Meter_Number)
Gas_Feb_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_No_of_billing_days)
Gas_Feb_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Rate_Category_Text)
Gas_Feb_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Service_Division)
Gas_Feb_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Feb_Taxes_and_Fees)
Gas_Jan_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Average_Temperature)
Gas_Jan_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Billed_Amount)
Gas_Jan_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Billing_Date)
Gas_Jan_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Cooling_degree_days)
Gas_Jan_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Energy_Consumption)
Gas_Jan_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Heating_degree_days)
Gas_Jan_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Meter_Number)
Gas_Jan_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_No_of_billing_days)
Gas_Jan_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Rate_Category_Text)
Gas_Jan_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Service_Division)
Gas_Jan_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jan_Taxes_and_Fees)
Gas_Jul_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Average_Temperature)
Gas_Jul_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Billed_Amount)
Gas_Jul_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Billing_Date)
Gas_Jul_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Cooling_degree_days)
Gas_Jul_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Energy_Consumption)
Gas_Jul_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Heating_degree_days)
Gas_Jul_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Meter_Number)
Gas_Jul_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_No_of_billing_days)
Gas_Jul_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Rate_Category_Text)
Gas_Jul_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Service_Division)
Gas_Jul_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jul_Taxes_and_Fees)
Gas_Jun_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Average_Temperature)
Gas_Jun_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Billed_Amount)
Gas_Jun_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Billing_Date)
Gas_Jun_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Cooling_degree_days)
Gas_Jun_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Energy_Consumption)
Gas_Jun_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Heating_degree_days)
Gas_Jun_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Meter_Number)
Gas_Jun_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_No_of_billing_days)
Gas_Jun_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Rate_Category_Text)
Gas_Jun_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Service_Division)
Gas_Jun_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Jun_Taxes_and_Fees)
Gas_Mar_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Average_Temperature)
Gas_Mar_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Billed_Amount)
Gas_Mar_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Billing_Date)
Gas_Mar_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Cooling_degree_days)
Gas_Mar_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Energy_Consumption)
Gas_Mar_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Heating_degree_days)
Gas_Mar_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Meter_Number)
Gas_Mar_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_No_of_billing_days)
Gas_Mar_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Rate_Category_Text)
Gas_Mar_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Service_Division)
Gas_Mar_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Mar_Taxes_and_Fees)
Gas_May_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Average_Temperature)
Gas_May_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Billed_Amount)
Gas_May_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Billing_Date)
Gas_May_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Cooling_degree_days)
Gas_May_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Energy_Consumption)
Gas_May_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Heating_degree_days)
Gas_May_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Meter_Number)
Gas_May_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_No_of_billing_days)
Gas_May_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Rate_Category_Text)
Gas_May_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Service_Division)
Gas_May_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_May_Taxes_and_Fees)
Gas_Nov_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Average_Temperature)
Gas_Nov_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Billed_Amount)
Gas_Nov_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Billing_Date)
Gas_Nov_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Cooling_degree_days)
Gas_Nov_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Energy_Consumption)
Gas_Nov_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Heating_degree_days)
Gas_Nov_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Meter_Number)
Gas_Nov_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_No_of_billing_days)
Gas_Nov_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Rate_Category_Text)
Gas_Nov_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Service_Division)
Gas_Nov_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Nov_Taxes_and_Fees)
Gas_Oct_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Average_Temperature)
Gas_Oct_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Billed_Amount)
Gas_Oct_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Billing_Date)
Gas_Oct_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Cooling_degree_days)
Gas_Oct_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Energy_Consumption)
Gas_Oct_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Heating_degree_days)
Gas_Oct_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Meter_Number)
Gas_Oct_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_No_of_billing_days)
Gas_Oct_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Rate_Category_Text)
Gas_Oct_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Service_Division)
Gas_Oct_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Oct_Taxes_and_Fees)
Gas_Sep_Average_Temperature.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Average_Temperature)
Gas_Sep_Billed_Amount.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Billed_Amount)
Gas_Sep_Billing_Date.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Billing_Date)
Gas_Sep_Cooling_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Cooling_degree_days)
Gas_Sep_Energy_Consumption.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Energy_Consumption)
Gas_Sep_Heating_degree_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Heating_degree_days)
Gas_Sep_Meter_Number.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Meter_Number)
Gas_Sep_No_of_billing_days.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_No_of_billing_days)
Gas_Sep_Rate_Category_Text.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Rate_Category_Text)
Gas_Sep_Service_Division.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Service_Division)
Gas_Sep_Taxes_and_Fees.Value = Sheets("Enrollments").Cells(Results, NexantEnrollments.Usage_Gas_Sep_Taxes_and_Fees)


'Sheets("Enrollments").Cells(Results, 12) = Format(ConvertLocalToGMT(Now(), True), "yyyymmdd:hhmmss")
Else
MsgBox ("No Enrollment ID found")
Me.Hide

End If


End Sub


