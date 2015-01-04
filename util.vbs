Attribute VB_Name = "util"

Sub test1()
    frmServiceCenter.Show vbModeless

End Sub

Sub testimport()
    frmImport.Show vbModeless
End Sub


' This is an example to use the Enum classes
Sub test3()
    Dim ic As Integer
    Dim testa As New EnrollmentEnum
    Dim str As String
    
    ic = 2
    str = testa.getField("3")
    str = testa.getField(CStr(ic))
    
    Select Case testa.getField(CStr(ic))
     Case "Enrollment_ID_ROSA"
        Cells(12, ic) = "abc"
     Case "Enrollment_ID_HEAP"
        Cells(12, ic) = "xyz"
    End Select
End Sub

Sub test4()
    Dim ic As Integer
    Dim testb As New MeasureEnum
    Dim str As String

    ic = 5
    str = testb.getField(CStr(ic))


    Select Case testb.getField(CStr(ic))
     Case "Enrollment_ID_ROSA"
        Worksheets("Measures").Cells(12, ic) = "abc"
     Case "Annual_kWh_Savings"
        Worksheets("Measures").Cells(12, ic) = "xyz"
    End Select
End Sub
