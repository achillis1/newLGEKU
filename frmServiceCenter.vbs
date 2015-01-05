VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServiceCenter 
   Caption         =   "Service Center"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2295
   OleObjectBlob   =   "frmServiceCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServiceCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()
    Me.Hide
    frmAdmin.Show vbModeless
End Sub

Private Sub cmdExit_Click()
    Application.Visible = True
    ThisWorkbook.Close
End Sub

Private Sub cmdExport_Click()
    Me.Hide
    frmExport.Show vbModeless
End Sub

Private Sub cmdImport_Click()
    Me.Hide
    frmImport.Show vbModeless
End Sub

Private Sub cmdProcessing_Click()
    Me.Hide
    frmProcessing.Show vbModeless
End Sub



Private Sub UserForm_Initialize()
'    Application.Visible = False
    
    EnrollmentFirstDataLine = 11
    ImportSheetName = "Enrollments"
    PMSheetName = "PM"
    InboundLastReadCol = 5
    currentEnrollment = ""
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
    'ThisWorkbook.Close
End Sub
