VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSFV 
   Caption         =   "     KnotonSFV 1.2  (Special thanks to Fredrik Qvarfort)"
   ClientHeight    =   2505
   ClientLeft      =   5370
   ClientTop       =   3750
   ClientWidth     =   6090
   Icon            =   "frmSFV.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   6090
   Begin MSComctlLib.ListView lstSFV 
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   6075
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   5280
         TabIndex        =   1
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblTime 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblfile 
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lblRun 
         Height          =   195
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblFail 
         Height          =   195
         Left            =   2220
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6420
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open SFV"
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Create SFV"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuAssociate 
         Caption         =   "Associate SFV"
      End
      Begin VB.Menu mnuUnassociate 
         Caption         =   "Unassociate SFV"
      End
      Begin VB.Menu mnubad 
         Caption         =   "Rename failed files .bad"
      End
   End
End
Attribute VB_Name = "frmSFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CRC As clsCRC
Private CD As CommonDialog
Private FilePath As String
Private ArrCheck() As String
Private DirPath As String
Private Check() As CheckFile
Private OldTick As Single
Private blnRenBad As Boolean

Private Type CheckFile
    Filename As String
    Value As String
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)

'Routine to cancel the current operation
Private Sub cmdCancel_Click()
blnRun = False
Call m_CRC.Clear
mnuOpen.Enabled = True
lblRun.Caption = "Canceled"
End Sub

'Routine to check/get the Checksum, add info to the listview, show various info
'If selected rename bad files .bad
Private Sub CheckSFV(File As String)
Dim i As Integer, x As Integer
Dim ret As String
Dim no As Integer
Dim Fail As Integer
Dim itmx As ListItem
Dim arrBad() As String
ReDim arrBad(0)
x = -1

'clear the listview of current data
Call ClearListView(lstSFV)
blnRun = True

'get the working directory
DirPath = Mid$(File, 1, InStrRev(File, "\"))

'Initialize the timer
OldTick = Timer

'Get the values in the SFV file
Call GetValues(File)
     
'If the values are valid
If Check(0).Filename <> "" Then
    
    'Initialize some info values
    mnuOpen.Enabled = False
    lblRun.Caption = "Running"
    lblFail.Caption = ""
    lblfile.Caption = ""
    lblTime.Caption = "Time:"
    DoEvents
    
    no = UBound(Check)
    
    'Loop the numbers of files to validate the checksum from
    For i = 0 To no
        If blnRun Then
            
            'If the file exist...
            If Len(Dir(DirPath & Check(i).Filename)) <> 0 Then
                'Get the checksum
                ret = m_CRC.CalculateFile(DirPath & Check(i).Filename)
                'add the file to the listview
                Set itmx = lstSFV.ListItems.Add(, , Check(i).Filename)
                'Validate the checksum and show the status
                If UCase(ret) = UCase(Check(i).Value) Then
                    itmx.SubItems(1) = "PASSED"
                ElseIf ret = "0" Or ret = "00000000" Then 'file has 0 size or something is very wrong with it
                    itmx.SubItems(1) = "UNREADABLE"
                    Fail = Fail + 1
                Else ' if the checksum is wrong mark it bad and add it to the badfile array
                    itmx.SubItems(1) = "FAILED"
                    Fail = Fail + 1
                    x = x + 1
                    ReDim Preserve arrBad(x)
                    arrBad(x) = DirPath & Check(i).Filename
                End If
            Else 'The file is missing
                Set itmx = lstSFV.ListItems.Add(, , Check(i).Filename)
                itmx.SubItems(1) = "MISSING"
                Fail = Fail + 1
        End If
            
            itmx.SubItems(2) = UCase(ret) 'show the current checksum value
            itmx.SubItems(3) = UCase(Check(i).Value) 'Show the original checksum value
            'itmx.Selected = True 'Scroll to the bottom
            lstSFV.ListItems(i + 1).EnsureVisible
            Call m_CRC.Clear 'Clear all old checksums
            DoEvents
            'Show some info
            lblTime.Caption = "Time:" & Format$(Timer - OldTick, "#00.00") & " s"
            lblfile.Caption = "File:" & i + 1 & "/" & no + 1
            lblFail.Caption = "Failed:" & Fail
            DoEvents
        End If
    Next
End If
'Show total time
lblTime.Caption = "Time:" & Format$(Timer - OldTick, "#00.00") & " s"
Call m_CRC.Clear

If blnRun Then
    If blnRenBad Then 'If selected rename bad files to .bad
        If arrBad(0) <> "" Then
            For i = 0 To x
                If Len(Dir(arrBad(i) & ".bad")) > 0 Then Kill arrBad(i) & ".bad"
                Name arrBad(i) As arrBad(i) & ".bad"
            Next
            lblTime.Caption = "Time:" & Format$(Timer - OldTick, "#00.00") & " s"
        End If
    End If
lblRun.Caption = "Done"
End If
End Sub

Private Sub Form_Load()
Set m_CRC = New clsCRC
Set CD = New CommonDialog
lstSFV.ColumnHeaders.Add 1, , "Filename"
lstSFV.ColumnHeaders.Add 2, , "Valid"
lstSFV.ColumnHeaders.Add 3, , "Checksum 1"
lstSFV.ColumnHeaders.Add 4, , "Checksum 2"
lstSFV.View = lvwReport
Me.Show

If Command <> "" Then
    Call CheckSFV(Command)
    mnuOpen.Enabled = True
End If

End Sub

'Open sfv file and get values from it
Private Sub GetValues(SFVFile As String)
Dim intFilNr As Integer, i As Integer, x As Integer
Dim strTemp As String
Dim strUnix As Variant
Dim nr As Integer

x = -1
i = -1
ReDim Check(0)
intFilNr = FreeFile
Open SFVFile For Input As intFilNr
While Not EOF(intFilNr)
    Line Input #intFilNr, strTemp
    
    'If the sfv file is in "Unix" style
    If InStr(1, strTemp, vbLf) <> 0 Then
        strUnix = Split(strTemp, vbLf)
        nr = UBound(strUnix) - 1
        
        For i = 0 To nr
            If Mid(strUnix(i), 1, 1) <> ";" Then
                x = x + 1
                ReDim Preserve Check(x)
                Check(x).Filename = Mid$(strUnix(i), 1, Len(strUnix(i)) - 9)
                Check(x).Value = Right$(strUnix(i), 8)
            End If
        Next
        
        Exit Sub
    End If
    'If the sfv file is "Windows" style
    If Not Mid$(Trim(strTemp), 1, 1) = ";" Then
        If Not Mid$(Trim(strTemp), 1, 1) = "" Then
            i = i + 1
            ReDim Preserve Check(i)
            Check(i).Filename = Mid$(strTemp, 1, Len(strTemp) - 9)
            Check(i).Value = Right$(strTemp, 8)
        End If
    End If
Wend
Close #intFilNr
End Sub

Private Sub Form_Resize()
If frmSFV.WindowState <> 1 Then
    lstSFV.Width = frmSFV.ScaleWidth - 30
    lstSFV.Height = frmSFV.ScaleHeight - fraInfo.Height
    fraInfo.Width = lstSFV.Width
    cmdCancel.Left = fraInfo.Left + (fraInfo.Width - cmdCancel.Width)
    fraInfo.Top = lstSFV.Top + lstSFV.Height - 50
    Line1.X2 = frmSFV.Width
    Call AdjustColumns
End If
End Sub


'Enable/disable the option to rename bad files .bad
Private Sub mnubad_Click()
If mnubad.Checked = True Then
    mnubad.Checked = False
    blnRenBad = False
Else
    mnubad.Checked = True
    blnRenBad = True
End If

End Sub

'Initialize the creation of a SFV file
Private Sub mnuCreate_Click()
Dim tmp As String
Dim mFiles As Variant
Dim i As Integer

CD.Filter = "All Files (*.*)|*.*"
CD.DialogTitle = "Select Files"
CD.AllowMultiSelect = True
CD.ShowOpen
mFiles = CD.Filename
If IsArray(mFiles) Then
    mnuCreate.Enabled = False
    tmp = mFiles(1)
    mFiles(1) = mFiles(UBound(mFiles))
    mFiles(UBound(mFiles)) = tmp
    ReDim Check(0)
    DirPath = mFiles(0) & "\"
    
    For i = 1 To UBound(mFiles)
        ReDim Preserve Check(i - 1)
        Check(i - 1).Filename = mFiles(i)
    Next
    
    blnRun = True
Else
    If mFiles <> "" Then
        mnuCreate.Enabled = False
        DirPath = Mid(mFiles, 1, InStrRev(mFiles, "\"))
        ReDim Check(0)
        Check(0).Filename = Mid(mFiles, InStrRev(mFiles, "\") + 1)
        blnRun = True
    End If
End If
   
    
    If blnRun Then
        Call GetChecksumValue
        Call WriteSFV
        lblRun.Caption = "SFV Saved"
    End If
    
    blnRun = False
    mnuCreate.Enabled = True

End Sub

'Create the SFV file
Private Sub WriteSFV()
Dim intFilNr As Integer
Dim SFVFile As Variant
Dim WriteWhat As String
Dim i As Integer

CD.Filter = "SFV Files (*.sfv)|*.sfv"
CD.DialogTitle = "Save SFV File"
CD.InitDir = DirPath
CD.AllowMultiSelect = False
CD.ShowSave
SFVFile = CD.Filename
If SFVFile <> "" Then

WriteWhat = ";Made by KnotonSFV " & Date & vbCrLf & ";" & vbCrLf
For i = 0 To UBound(Check)
    If i <> UBound(Check) Then
        WriteWhat = WriteWhat & Check(i).Filename & " " & Check(i).Value & vbCrLf
    Else
        WriteWhat = WriteWhat & Check(i).Filename & " " & Check(i).Value
    End If
Next

intFilNr = FreeFile
Open SFVFile For Append As intFilNr
Print #intFilNr, WriteWhat
Close #intFilNr
End If
End Sub

'Get Checksum values for the creation of the SFV File
Private Sub GetChecksumValue()
Dim no As Long
Dim ret As String
Dim itmx As ListItem
Dim i As Integer
OldTick = Timer
Call ClearListView(lstSFV)
lblRun.Caption = "Running"
DoEvents

no = UBound(Check)
For i = 0 To no
    If blnRun Then
            If Dir(DirPath & Check(i).Filename) <> "" Then
                ret = m_CRC.CalculateFile(DirPath & Check(i).Filename)
                Check(i).Value = ret
                Set itmx = lstSFV.ListItems.Add(, , Check(i).Filename)
                itmx.SubItems(2) = UCase(ret)
                itmx.Selected = True
                DoEvents
                lblTime.Caption = "Time:" & Format$(Timer - OldTick, "#00.00") & ""
                lblfile.Caption = "File:" & i + 1 & "/" & no + 1
                DoEvents
            End If
    End If
Call m_CRC.Clear
Next

End Sub

'Do some clearing up before closing
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call m_CRC.Clear
Set m_CRC = Nothing
Set CD = Nothing
End Sub

'Clear the listview
Private Sub ClearListView(LV As ListView)
Call SendMessage(LV.hWnd, LVM_DELETEALLITEMS, 0, ByVal 0&)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

'Call the routine to remove association of sfv file
Private Sub mnuUnassociate_Click()
RemoveAssociate App.EXEName, ".sfv"
End Sub

'Call the routine to add association of sfv files
Private Sub mnuAssociate_Click()
Associate App.EXEName, ".sfv", "Checksum File"
End Sub

'Initialize the opening of a sfv file
Private Sub mnuOpen_Click()
Dim tmp As String
CD.Filter = "SFV Files (*.sfv)|*.sfv"
CD.DialogTitle = "Choose file to check"
CD.AllowMultiSelect = False
CD.ShowOpen
tmp = CD.Filename
If tmp <> "" Then
    Call CheckSFV(tmp)
End If
mnuOpen.Enabled = True

End Sub

Sub AdjustColumns()
Dim i As Integer, x As Double

x = lstSFV.Width / lstSFV.ColumnHeaders.Count '- 300

For i = 1 To lstSFV.ColumnHeaders.Count
    lstSFV.ColumnHeaders(i).Width = x
Next i

End Sub



