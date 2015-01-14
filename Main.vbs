Attribute VB_Name = "Main"
Public EnrollmentFirstDataLine As Integer
Public ImportSheetName As String
Public PMSheetName As String
Public MessageSheetName As String
Public ImportHistorySN As String
Public InboundLastReadCol As Integer
Public currentEnrollment As String

'''' Added by Ding '''
Public anum() As Integer
Public adate() As String
Public arow() As Integer
Public anote() As String
Public atype() As String
Public attemptnum As Integer

Sub Main()
    frmServiceCenter.Show vbModeless
End Sub
'
'Sub showwk()
'        Application.Visible = True
'End Sub

