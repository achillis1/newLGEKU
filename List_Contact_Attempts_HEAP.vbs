VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} List_Contact_Attempts_HEAP 
   Caption         =   "List_Contact_Attempts_HEAP"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15180
   OleObjectBlob   =   "List_Contact_Attempts_HEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "List_Contact_Attempts_HEAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_HEAP_Click()
Me.Hide
frmAdmin.Show

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub Clear_HEAP_Click()

'Clear Fields
Enrollment_Listbox.Clear
        Me.Enrollment_ID_HEAP = ""
        Me.First_Contact_Attempt_Date_HEAP.Value = ""
        Me.First_Contact_Attempt_Notes_HEAP = ""
        Me.First_Contact_Attempt_Type_HEAP = ""
        Me.Second_Contact_Attempt_Date_HEAP = ""
        Me.Second_Contact_Attempt_Notes_HEAP = ""
        Me.Second_Contact_Attempt_Type_HEAP = ""
        Me.Third_Contact_Attempt_Date_HEAP = ""
        Me.Third_Contact_Attempt_Notes_HEAP = ""
        Me.Third_Contact_Attempt_Type_HEAP = ""
        Me.Fourth_Contact_Attempt_Date_HEAP = ""
        Me.Fourth_Contact_Attempt_Notes_HEAP = ""
        Me.Fourth_Contact_Attempt_Type_HEAP = ""
        Me.Fifth_Contact_Attempt_Date_HEAP = ""
        Me.Fifth_Contact_Attempt_Notes_HEAP = ""
        Me.Fifth_Contact_Attempt_Type_HEAP = ""
        Me.Schedule_Date_HEAP = ""
        Me.Schedule_Time_HEAP = ""

Call UserForm_Initialize
End Sub

Private Sub Enrollment_Listbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Set wsdb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Enrollment_Listbox.Value
'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'Retrive values from Database
Me.Enrollment_ID_HEAP = EID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
        'push data from database to form
        'HEAP Scheduling

        Me.First_Contact_Attempt_Date_HEAP = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP)
        Me.First_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP)
        Me.First_Contact_Attempt_Type_HEAP = wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP)
        Me.Second_Contact_Attempt_Date_HEAP = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP)
        Me.Second_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP)
        Me.Second_Contact_Attempt_Type_HEAP = wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP)
        Me.Third_Contact_Attempt_Date_HEAP = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP)
        Me.Third_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP)
        Me.Third_Contact_Attempt_Type_HEAP = wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP)
        Me.Fourth_Contact_Attempt_Date_HEAP = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP)
        Me.Fourth_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP)
        Me.Fourth_Contact_Attempt_Type_HEAP = wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP)
        Me.Fifth_Contact_Attempt_Date_HEAP = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP)
        Me.Fifth_Contact_Attempt_Notes_HEAP = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP)
        Me.Fifth_Contact_Attempt_Type_HEAP = wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP)
        Me.Schedule_Date_HEAP = wsdb.Cells(x, NexantEnrollments.Schedule_Date_HEAP)
        Me.Schedule_Time_HEAP = wsdb.Cells(x, NexantEnrollments.Schedule_Time_HEAP)
        
      

    End If
Next x



End Sub


Private Sub Save_HEAP_Click()
Set wsdb = Worksheets("Enrollments")

'Enrollment_Listbox.Value = EID
EID = Me.Enrollment_ID_HEAP.Value
'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP) = EID Then
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Date_HEAP) = Me.First_Contact_Attempt_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Notes_HEAP) = Me.First_Contact_Attempt_Notes_HEAP
         wsdb.Cells(x, NexantEnrollments.First_Contact_Attempt_Type_HEAP) = Me.First_Contact_Attempt_Type_HEAP
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Date_HEAP) = Me.Second_Contact_Attempt_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Notes_HEAP) = Me.Second_Contact_Attempt_Notes_HEAP
         wsdb.Cells(x, NexantEnrollments.Second_Contact_Attempt_Type_HEAP) = Me.Second_Contact_Attempt_Type_HEAP
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Date_HEAP) = Me.Third_Contact_Attempt_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Notes_HEAP) = Me.Third_Contact_Attempt_Notes_HEAP
         wsdb.Cells(x, NexantEnrollments.Third_Contact_Attempt_Type_HEAP) = Me.Third_Contact_Attempt_Type_HEAP
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Date_HEAP) = Me.Fourth_Contact_Attempt_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Notes_HEAP) = Me.Fourth_Contact_Attempt_Notes_HEAP
         wsdb.Cells(x, NexantEnrollments.Fourth_Contact_Attempt_Type_HEAP) = Me.Fourth_Contact_Attempt_Type_HEAP
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Date_HEAP) = Me.Fifth_Contact_Attempt_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Notes_HEAP) = Me.Fifth_Contact_Attempt_Notes_HEAP
         wsdb.Cells(x, NexantEnrollments.Fifth_Contact_Attempt_Type_HEAP) = Me.Fifth_Contact_Attempt_Type_HEAP
         wsdb.Cells(x, NexantEnrollments.Schedule_Date_HEAP) = Me.Schedule_Date_HEAP
         wsdb.Cells(x, NexantEnrollments.Schedule_Time_HEAP) = Me.Schedule_Time_HEAP
    End If
Next x
End Sub

Private Sub UserForm_Initialize()

Set wsdb = Worksheets("Enrollments")

'last row database
wsDblr = wsdb.Cells(Rows.Count, 2).End(xlUp).Row

'find row in Database for Enrollment ID
For x = 11 To wsDblr
    If wsdb.Cells(x, NexantEnrollments.RECEIVED_AT_VENDOR_date_interfaced_HEAP) <> "" Or wsdb.Cells(x, NexantEnrollments.FIRST_CONTACT_date_interfaced_HEAP) <> "" Or wsdb.Cells(x, NexantEnrollments.PENDING_1_date_interfaced_HEAP) <> "" Then
        'push data from database to form
        'HEAP Scheduling
        With Enrollment_Listbox
            .AddItem wsdb.Cells(x, NexantEnrollments.Enrollment_ID_HEAP)
        End With
              
    End If
Next x
End Sub
