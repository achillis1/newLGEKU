VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnrollmentHeaderEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum NexantEnrollmentsHeader
    Record_Type = 1
    File_Date
    Vendor_Name
    Short_Program_Name
    Miscellaneous
End Enum

Public Function getField(ByRef field As NexantEnrollmentsHeader) As String
        Select Case field
            Case NexantEnrollmentsHeader.Record_Type: getField = "Record_Type"
            Case NexantEnrollmentsHeader.File_Date: getField = "File_Date"
            Case NexantEnrollmentsHeader.Vendor_Name: getField = "Vendor_Name"
            Case NexantEnrollmentsHeader.Short_Program_Name: getField = "Short_Program_Name"
            Case NexantEnrollmentsHeader.Miscellaneous: getField = "Miscellaneous"
        End Select
End Function
