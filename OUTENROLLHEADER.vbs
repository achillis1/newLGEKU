VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OUTENROLLHEADER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum LGEEnrollmentsHeader
    Record_Type = 0
    File_Date
    Vendor_Name
    Short_Program_Name
    Miscellaneous
End Enum

Public Function getField(ByRef field As LGEEnrollmentsHeader) As String
        Select Case field
            Case LGEEnrollmentsHeader.Record_Type: getField = "Record_Type"
            Case LGEEnrollmentsHeader.File_Date: getField = "File_Date"
            Case LGEEnrollmentsHeader.Vendor_Name: getField = "Vendor_Name"
            Case LGEEnrollmentsHeader.Short_Program_Name: getField = "Short_Program_Name"
            Case LGEEnrollmentsHeader.Miscellaneous: getField = "Miscellaneous"
        End Select
End Function


