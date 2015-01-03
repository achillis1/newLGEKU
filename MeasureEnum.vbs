VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeasureEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum MeasureColumns
    Enrollment_ID_ROSA = 2
    Enrollment_ID_HEAP
    Annual_CCF_Savings
    Annual_kWh_Savings
    ECM_ID
    Estimated_annual_savings
    Estimated_contractor_cost
    Estimated_contractor_payback_in_years
    Estimated_DIY_cost
    Estimated_DIY_payback_in_years
    Installation_Status
    Last_Modified_Date_Measure
    Measure_Notes
    Measure_Type
    Notes
    VRM_ID
    VRM_Quantity
End Enum

Public Function getField(ByRef field As MeasureColumns) As String
        Select Case field
            Case MeasureColumns.Enrollment_ID_ROSA: getField = "Enrollment_ID_ROSA"
            Case MeasureColumns.Enrollment_ID_HEAP: getField = "Enrollment_ID_HEAP"
            Case MeasureColumns.Annual_CCF_Savings: getField = "Annual_CCF_Savings"
            Case MeasureColumns.Annual_kWh_Savings: getField = "Annual_kWh_Savings"
            Case MeasureColumns.ECM_ID: getField = "ECM_ID"
            Case MeasureColumns.Estimated_annual_savings: getField = "Estimated_annual_savings"
            Case MeasureColumns.Estimated_contractor_cost: getField = "Estimated_contractor_cost"
            Case MeasureColumns.Estimated_contractor_payback_in_years: getField = "Estimated_contractor_payback_in_years"
            Case MeasureColumns.Estimated_DIY_cost: getField = "Estimated_DIY_cost"
            Case MeasureColumns.Estimated_DIY_payback_in_years: getField = "Estimated_DIY_payback_in_years"
            Case MeasureColumns.Installation_Status: getField = "Installation_Status"
            Case MeasureColumns.Last_Modified_Date_Measure: getField = "Last_Modified_Date_Measure"
            Case MeasureColumns.Measure_Notes: getField = "Measure_Notes"
            Case MeasureColumns.Measure_Type: getField = "Measure_Type"
            Case MeasureColumns.Notes: getField = "Notes"
            Case MeasureColumns.VRM_ID: getField = "VRM_ID"
            Case MeasureColumns.VRM_Quantity: getField = "VRM_Quantity"
        End Select
End Function

