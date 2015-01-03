VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Import"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Me.Hide
    frmServiceCenter.Show
End Sub

Private Sub cmdImport_Click()
    If opbEnrollment Then Call ImportFile("Enrollment")
    If opbUsage Then Call ImportFile("Usage")

    iResult = MsgBox("Import is completed.", vbOKOnly, "Import Done!")
End Sub

Private Sub CommandButton1_Click()
    Call SplitDemo
End Sub

Private Sub UserForm_Terminate()
    frmServiceCenter.Show
End Sub


Sub SplitDemo()
    Dim txt As String
    Dim x As Variant
    Dim i As Long
    txt = "Skeeter,PJ,Purkey,Little Guy"
    x = Split(txt, ",")
    For i = 0 To UBound(x)
       Debug.Print x(i)
    Next i
End Sub

Sub ImportFile(ByVal str As String)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' On Error Resume Next
    Dim thisWB As Workbook
    Dim importWB As Workbook
    Set thisWB = ActiveWorkbook
    
    filetoopen = Application.GetOpenFilename(FileFilter:="OUT Files (*.txt), *.txt", Title:="Select the usage OUT file")
    If filetoopen = False Then
        exit_macro1 = True
        exit_macro_reason = "User canceled during file selection."
        Exit Sub
    End If
    Workbooks.Open Filename:=filetoopen, Format:=1, corruptload:=xlRepairFile
    
    slashloc = InStrRev(filetoopen, "\")
    Filename = Mid(filetoopen, slashloc + 1, Len(filetoopen) - slashloc - 4)
    'C:\share\DriveZ\LGE\documents\HEAP\DSM_HEAP_ENROLL_OUT.TXT"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + filetoopen, Destination:= _
        Cells(1, 1))
        .Name = Filename
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 936
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 2, 2, 2, 1, 2, 1, 2, 1, 1, 2, 2, 1, 1, 1, 2, 1, 1, 1, 2, 1, _
        1, 1, 1, 1, 2, 1, 2, 1, 1, 2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 2, 1, 2, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Set importWB = ActiveWorkbook
    
    lastrow = importWB.Worksheets(Filename).Range("A" & Rows.Count).End(xlUp).Row
    
    Dim existingdata As Range
    Set existingdata = importWB.Worksheets(Filename).Range("a2:AV" & lastrow - 1)
    
    Dim rng As Range
    lastenroll = thisWB.Worksheets(str).Range("A" & Rows.Count).End(xlUp).Row
    Set rng = thisWB.Worksheets(str).Range("A" & lastenroll + 1)
    existingdata.Copy
    rng.PasteSpecial xlValues
    Application.CutCopyMode = False
    
    importWB.Close False
    Set importWB = Nothing
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    thisWB.Worksheets(str).Range("A1").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


