VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OUTENROLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum LGEEnrollments
    Record_Type = 0
    Transaction_Type
    Enrollment_ID
    Premise_ID
    Company_Code
    Customer_Account
    Main_Account_Flag
    Business_Partner_Number
    Premise_Type
    Service_customer_name
    Customer_Home_Phone
    Customer_mobile_phone
    Service_Street_Address
    Service_City
    Service_State
    Service_Zipcode
    Mailing_Street_Address
    Mailing_City
    Mailing_State
    Mailing_Zipcode
    Customer_Email
    Contact_Name
    Primary_Contact_Address
    Primary_Contact_Address_City
    Primary_Contact_Address_State
    Primary_Contact_Address_Zip
    Primary_Contact_Email
    Primary_Contact_Phone
    Primary_Contact_phone_extension
    Primary_Contact_mobile_phone
    Nbr_Building_Occupants
    Total_conditioned_square_footage
    Send_Report_to_Primary_Contact
    Dog_or_Cat_Flag
    Ownership_Type
    Reason_for_audit
    Short_Program_Name
    Verification_Class
    Baseline_Tier1_vendor_work_order_number
    Remit_to_Contact_Name
    Remit_to_Contact_Address
    Remit_to_Contact_Address_City
    Remit_to_Contact_Address_State
    Remit_to_Contact_Address_Zip
    Remit_to_Contact_Email
    Remit_to_Contact_Phone
    Remit_to_Contact_phone_extension
    Remit_to_Contact_mobile_phone
End Enum

Public Function getField(ByRef field As LGEEnrollments) As String
        Select Case field
            Case LGEEnrollments.Record_Type: getField = "Record_Type"
            Case LGEEnrollments.Transaction_Type: getField = "Transaction_Type"
            Case LGEEnrollments.Enrollment_ID: getField = "Enrollment_ID"
            Case LGEEnrollments.Premise_ID: getField = "Premise_ID"
            Case LGEEnrollments.Company_Code: getField = "Company_Code"
            Case LGEEnrollments.Customer_Account: getField = "Customer_Account"
            Case LGEEnrollments.Main_Account_Flag: getField = "Main_Account_Flag"
            Case LGEEnrollments.Business_Partner_Number: getField = "Business_Partner_Number"
            Case LGEEnrollments.Premise_Type: getField = "Premise_Type"
            Case LGEEnrollments.Service_customer_name: getField = "Service_customer_name"
            Case LGEEnrollments.Customer_Home_Phone: getField = "Customer_Home_Phone"
            Case LGEEnrollments.Customer_mobile_phone: getField = "Customer_mobile_phone"
            Case LGEEnrollments.Service_Street_Address: getField = "Service_Street_Address"
            Case LGEEnrollments.Service_City: getField = "Service_City"
            Case LGEEnrollments.Service_State: getField = "Service_State"
            Case LGEEnrollments.Service_Zipcode: getField = "Service_Zipcode"
            Case LGEEnrollments.Mailing_Street_Address: getField = "Mailing_Street_Address"
            Case LGEEnrollments.Mailing_City: getField = "Mailing_City"
            Case LGEEnrollments.Mailing_State: getField = "Mailing_State"
            Case LGEEnrollments.Mailing__Zipcode: getField = "Mailing_Zipcode"
            Case LGEEnrollments.Customer_Email: getField = "Customer_Email"
            Case LGEEnrollments.Contact_Name: getField = "Contact_Name"
            Case LGEEnrollments.Primary_Contact_Address: getField = "Primary_Contact_Address"
            Case LGEEnrollments.Primary_Contact_Address_City: getField = "Primary_Contact_Address_City"
            Case LGEEnrollments.Primary_Contact_Address_State: getField = "Primary_Contact_Address_State"
            Case LGEEnrollments.Primary_Contact_Address_Zip: getField = "Primary_Contact_Address_Zip"
            Case LGEEnrollments.Primary_Contact_Email: getField = "Primary_Contact_Email"
            Case LGEEnrollments.Primary_Contact_Phone: getField = "Primary_Contact_Phone"
            Case LGEEnrollments.Primary_Contact_phone_extension: getField = "Primary_Contact_phone_extension"
            Case LGEEnrollments.Primary_Contact_mobile_phone: getField = "Primary_Contact_mobile_phone"
            Case LGEEnrollments.Nbr_Building_Occupants: getField = "Nbr_Building_Occupants"
            Case LGEEnrollments.Total_conditioned_square_footage: getField = "Total_conditioned_square_footage"
            Case LGEEnrollments.Send_Report_to_Primary_Contact: getField = "Send_Report_to_Primary_Contact"
            Case LGEEnrollments.Dog_or_Cat_Flag: getField = "Dog_or_Cat_Flag"
            Case LGEEnrollments.Ownership_Type: getField = "Ownership_Type"
            Case LGEEnrollments.Reason_for_audit: getField = "Reason_for_audit"
            Case LGEEnrollments.Short_Program_Name: getField = "Short_Program_Name"
            Case LGEEnrollments.Verification_Class: getField = "Verification_Class"
            Case LGEEnrollments.Baseline_Tier1_vendor_work_order_number: getField = "Baseline_Tier1_vendor_work_order_number"
            Case LGEEnrollments.Remit_to_Contact_Name: getField = "Remit_to_Contact_Name"
            Case LGEEnrollments.Remit_to_Contact_Address: getField = "Remit_to_Contact_Address"
            Case LGEEnrollments.Remit_to_Contact_Address_City: getField = "Remit_to_Contact_Address_City"
            Case LGEEnrollments.Remit_to_Contact_Address_State: getField = "Remit_to_Contact_Address_State"
            Case LGEEnrollments.Remit_to_Contact_Address_Zip: getField = "Remit_to_Contact_Address_Zip"
            Case LGEEnrollments.Remit_to_Contact_Email: getField = "Remit_to_Contact_Email"
            Case LGEEnrollments.Remit_to_Contact_Phone: getField = "Remit_to_Contact_Phone"
            Case LGEEnrollments.Remit_to_Contact_phone_extension: getField = "Remit_to_Contact_phone_extension"
            Case LGEEnrollments.Remit_to_Contact_mobile_phone: getField = "Remit_to_Contact_mobile_phone"
        End Select
End Function


