VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDailyProcessEntryN 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6201
   ClientLeft      =   6981
   ClientTop       =   2119
   ClientWidth     =   5902
   ControlBox      =   0   'False
   Icon            =   "frmDailyProcessEntryN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6201
   ScaleWidth      =   5902
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4410
      TabIndex        =   15
      Top             =   1230
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtFDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2190
      TabIndex        =   14
      Top             =   1230
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtDays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3570
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5520
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2745
      Left            =   4410
      TabIndex        =   12
      Top             =   1335
      Width           =   1335
      Begin VB.Label Label2 
         Caption         =   "'W' = Weekly            Holiday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "'S' = Present          with No Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   150
         TabIndex        =   6
         Top             =   2190
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "'F' = Full Pay +         2 Hrs OT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   150
         TabIndex        =   5
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "'G' = Govt.              Holiday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "'R' = Regular          Day"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   5520
      Width           =   1155
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1530
      Width           =   975
   End
   Begin VB.ComboBox cboMonth 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Name of Month"
      Top             =   1530
      Width           =   3105
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3225
      Left            =   150
      TabIndex        =   9
      Top             =   1905
      Width           =   4170
      _ExtentX        =   7356
      _ExtentY        =   5679
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   285
      BackColor       =   -2147483634
      ForeColor       =   12582912
      BackColorFixed  =   -2147483632
      ForeColorFixed  =   16776960
      BackColorSel    =   4210752
      ForeColorSel    =   16776960
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483630
      GridColorFixed  =   -2147483630
      AllowBigSelection=   -1  'True
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Daily Process Entry (N)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   180
      TabIndex        =   18
      Top             =   240
      Width           =   3405
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   5280
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   0
      Left            =   -240
      TabIndex        =   16
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label29 
      Caption         =   "Name of Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   1530
      TabIndex        =   10
      Top             =   1260
      Width           =   255
   End
End
Attribute VB_Name = "frmDailyProcessEntryN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

Dim dtFirstDate As Date
Dim strInterval As String * 1
Dim intNumber As Integer
Dim dtCurrDate As Date

If Len(txtFDate) = 0 Then
  MsgBox "Type First Date", vbInformation, cnstMsgInfo
  txtFDate.SetFocus
  Exit Sub
End If
If Len(txtDays) = 0 Then
  MsgBox "Type No. of Days", vbInformation, cnstMsgInfo
  txtDays.SetFocus
  Exit Sub
End If

''strInterval = "d"
''dtFirstDate = InputBox("Enter the first date", "Message")
''intNumber = InputBox("Enter number of days", "Message")

strInterval = "d"
dtFirstDate = CDate(txtFDate.Text)
intNumber = CInt(txtDays.Text)

With flexDetails
  .Rows = .Rows + 1
  .row = .Rows - 1
  .col = 0
  .Text = Format(dtFirstDate, cnstDtFrmtI)
  .col = 1
  .Text = Format(dtFirstDate, "dddd")
  .col = 2
  If Format(dtFirstDate, "ddd") = "Fri" Then .Text = "Y" Else .Text = "N"
End With

For intI = 1 To intNumber - 1
  With flexDetails
    .Rows = .Rows + 1
    .row = .Rows - 1
    .col = 0
    .Text = Format(DateAdd(strInterval, intI, dtFirstDate), cnstDtFrmtI)
    dtCurrDate = .Text
    .col = 1
    .Text = Format(dtCurrDate, "dddd")
    .col = 2
    If Format(dtCurrDate, "ddd") = "Fri" Then .Text = "Y" Else .Text = "N"
  End With
  intNumber = intI + 1
Next intI

txtFDate.Text = ""
txtDays.Text = ""

MsgBox "Successfully Add", vbInformation, cnstMsgInfo
End Sub
Private Sub cmdUpdate_Click()

MsgBox "Required Update"

Exit Sub

Dim strDate As String, strStatus As String
Dim strDay As String, bytHoliday As Byte
        
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If flexDetails.Rows = 1 Then
  MsgBox "Data not found to Update", vbInformation, cnstMsgInfo
  Exit Sub
End If

If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

Screen.MousePointer = vbHourglass
Set r = New ADODB.Recordset
StrRecord = "SELECT DeptName FROM Salary_And_Wages WHERE (([Month]='" & UCase(cboMonth.Text) & "')" _
          & " AND ([Year]='" & txtYear.Text & "') AND (EntryStatus='A') AND (ComID=" & CompanyID & ")) GROUP BY DeptName;"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  MsgBox "Attendance already adjusted with Salary", vbCritical, cnstMsgInfo
  Screen.MousePointer = vbDefault
  Exit Sub
End If
r.Close
Screen.MousePointer = vbDefault

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500
DeleteRow = 2
MainConn.BeginTrans

'' Generate Calender

'''=== Add to Daily Process Status ===' Calender
'''==================================='
'''With flexDetails
'''  For intI = 1 To .Rows - 1
'''    .row = intI
'''    .col = 0:       strDate = Format(.Text, cnstDtFrmtQ)
'''    .col = 1:       strDay = Format(strDate, "ddd")
'''    .col = 2:       If .Text = "Y" Then bytHoliday = 1 Else bytHoliday = 0
'''    If bytHoliday = 1 Then strStatus = "W" Else strStatus = "R"
'''
'''    MainComm.CommandText = "INSERT INTO Daily_Process_Status(ComID,prDate,prDay,Holiday,Status,prMonth,prYear,DayHide)VALUES(" & CompanyID & ",'" & strDate & "'," _
'''          & " '" & strDay & "'," & bytHoliday & ",'" & strStatus & "','" & Format(strDate, "mm") & "'," & CInt(Format(strDate, "yyyy")) & ",'N');"
'''    MainComm.Execute
'''  Next intI
'''End With
'''==================================='

With flexDetails
  For intI = 1 To .Rows - 1
    .row = intI
    .col = 0
    strDate = Format(.Text, cnstDtFrmtQ)
    .col = 3
    strStatus = .Text
    If Len(strStatus) <> 0 Then
    
    MainComm.CommandText = "SET NOCOUNT ON Insert Into Daily_Process_Status_DelEdit Select ComID,prDate,prDay,Holiday,Status,prMonth,prYear,DayHide," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() from Daily_Process_Status WHERE ((prDate='" & strDate & "') AND (ComID=" & CompanyID & "));"
    MainComm.Execute
      
      MainComm.CommandText = "SET NOCOUNT ON UPDATE Daily_Process_Status SET Status='" & strStatus & "' WHERE ((prDate='" & strDate & "') AND (ComID=" & CompanyID & "));"
      MainComm.Execute
    End If
  Next intI
End With

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  Err.Number = 0
  Exit Sub
End If

flexDetails.Rows = 1
Screen.MousePointer = vbDefault
Set r = Nothing
End Sub
Private Sub cboMonth_Click()

  If Len(cboMonth) = 0 Then Exit Sub
  Dim bytMonth As Byte

  Set r = New ADODB.Recordset
  StrRecord = "SELECT SLNo FROM Month_Name WHERE [Month]='" & cboMonth.Text & "';"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then bytMonth = r![SLNO]
  r.Close

  flexDetails.Rows = 1
  StrRecord = "SELECT prDate,prDay,Holiday,Status From Daily_Process_Status" _
            & " WHERE (((Month([prDate]))=" & bytMonth & ") AND ((Year([prDate]))='" & txtYear.Text & "') AND (ComID=" & CompanyID & "));"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    Do Until r.EOF
      With flexDetails
        .Rows = .Rows + 1
        .row = .Rows - 1
        .col = 0
        .Text = Format(r![prDate], cnstDtFrmtI)
        .col = 1
        .Text = Format(r![prDate], "dddd")
        .col = 2
        If r![Holiday] = True Then .Text = "Y" Else .Text = "N"
        .col = 3
        If Not IsNull(r![Status]) Then .Text = r![Status]
      End With
      r.MoveNext
    Loop
  End If
  r.Close
  Set r = Nothing
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
With flexDetails
  If .RowSel = 0 Then Exit Sub
  If KeyCode = 13 And .col = 2 Then
    If .Text = "N" Then .Text = "Y" Else .Text = "N"
  End If
  If .col = 3 Then
    If KeyCode = 70 Then .Text = "F"
    If KeyCode = 71 Then .Text = "G"
    If KeyCode = 82 Then .Text = "R"
    If KeyCode = 83 Then .Text = "S"
    If KeyCode = 87 Then .Text = "W"
  End If
End With
End Sub
Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Daily Process Entry"
  
  Set r = New ADODB.Recordset
StrRecord = "SELECT Editstatus FROM Sys_User_Name WHERE UserId='" & bytUserID & "';"
r.Open StrRecord, MainConn, adOpenStatic
If r![Editstatus] = 1 Then
  cmdUpdate.Enabled = True
  Else
   cmdUpdate.Enabled = False
   End If
   
  With flexDetails
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .row = 0
    .col = 0
    .Text = "Date"
    .col = 1
    .Text = "Day"
    .col = 2
    .Text = "Holiday"
    .col = 3
    .Text = "Status"
  End With
  
  prcMonthName cboMonth
  txtYear.Text = PeriodYear
  cboMonth.Text = UCase(Format(Date, "MMMM"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmDailyProcessEntry = Nothing
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtFDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtFDate_LostFocus()
  txtFDate.Text = FormatDate(txtFDate)
End Sub

