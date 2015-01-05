VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OUTHEADER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum LGEHeader
    Record_Type = 0
    File_Date
    Vendor_Name
    Short_Program_Name
    Miscellaneous
End Enum

Public Function getField(ByRef field As LGEHeader) As String
        Select Case field
            Case LGEHeader.Record_Type: getField = "Record_Type"
            Case LGEHeader.File_Date: getField = "File_Date"
            Case LGEHeader.Vendor_Name: getField = "Vendor_Name"
            Case LGEHeader.Short_Program_Name: getField = "Short_Program_Name"
            Case LGEHeader.Miscellaneous: getField = "Miscellaneous"
        End Select
End Function


