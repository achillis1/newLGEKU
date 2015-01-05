VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OUTUSAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum LGEUsage
    Record_Type = 0
    Transaction_Type
    Enrollment_ID
    Premise_ID
    Meter_Number
    Rate_Category_Text
    Billing_Date
    Billed_Amount
    Taxes_and_Fees
    PF_On_Peak_Electric
    Power_Factor_on_adjustment_Electric
    Energy_Consumption
    KW_Billed_on_Demand_Electric
    Average_Temperature
    Heating_degree_days
    Cooling_degree_days
    No_of_billing_days
    Service_Division
End Enum

Public Function getField(ByRef field As LGEUsage) As String
    Select Case field
        Case LGEUsage.Record_Type: getField = "Record_Type"
        Case LGEUsage.Transaction_Type: getField = "Transaction_Type"
        Case LGEUsage.Enrollment_ID: getField = "Enrollment_ID"
        Case LGEUsage.Premise_ID: getField = "Premise_ID"
        Case LGEUsage.Meter_Number: getField = "Meter_Number"
        Case LGEUsage.Rate_Category_Text: getField = "Rate_Category_Text"
        Case LGEUsage.Billing_Date: getField = "Billing_Date"
        Case LGEUsage.Billed_Amount: getField = "Billed_Amount"
        Case LGEUsage.Taxes_and_Fees: getField = "Taxes_and_Fees"
        Case LGEUsage.PF_On_Peak_Electric: getField = "PF_On_Peak_Electric"
        Case LGEUsage.Power_Factor_on_adjustment_Electric: getField = "Power_Factor_on_adjustment_Electric"
        Case LGEUsage.Energy_Consumption: getField = "Energy_Consumption"
        Case LGEUsage.KW_Billed_on_Demand_Electric: getField = "KW_Billed_on_Demand_Electric"
        Case LGEUsage.Average_Temperature: getField = "Average_Temperature"
        Case LGEUsage.Heating_degree_days: getField = "Heating_degree_days"
        Case LGEUsage.Cooling_degree_days: getField = "Cooling_degree_days"
        Case LGEUsage.No_of_billing_days: getField = "No_of_billing_days"
        Case LGEUsage.Service_Division: getField = "Service_Division"
    End Select
        
End Function
