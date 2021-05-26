VERSION 5.00
Begin VB.Form frmBonusEntry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6084
   ClientLeft      =   1820
   ClientTop       =   1872
   ClientWidth     =   8216
   ControlBox      =   0   'False
   Icon            =   "frmBonusEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   8216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMLeave 
      BackColor       =   &H00C0C000&
      Caption         =   "Add Emp. &M. Leave "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   150
      TabIndex        =   30
      Top             =   3540
      Width           =   1005
   End
   Begin VB.CommandButton cmdRemoveAll 
      BackColor       =   &H00C0C000&
      Caption         =   "Remove &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1290
      TabIndex        =   7
      Top             =   3630
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C0C000&
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3630
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add to List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   3630
      Width           =   1095
   End
   Begin VB.TextBox txtToDate 
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
      Left            =   2190
      TabIndex        =   4
      Top             =   3150
      Width           =   2415
   End
   Begin VB.TextBox txtFromDate 
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
      Left            =   2190
      TabIndex        =   3
      Top             =   2670
      Width           =   2415
   End
   Begin VB.ComboBox cboDrawnMonth 
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
      Left            =   2190
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Month"
      Top             =   1230
      Width           =   2415
   End
   Begin VB.TextBox txtHeader 
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
      Left            =   2190
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4710
      Width           =   5895
   End
   Begin VB.ComboBox cboSubDept 
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
      Left            =   2190
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select Sub Department"
      Top             =   2190
      Width           =   2415
   End
   Begin VB.ComboBox cboDept 
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
      Left            =   2190
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Name of Department"
      Top             =   1710
      Width           =   2415
   End
   Begin VB.TextBox txtTotal 
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
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   6630
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   15
      Top             =   4305
      Width           =   1455
   End
   Begin VB.ListBox CodeList 
      BackColor       =   &H8000000F&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2964
      Left            =   4830
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1230
      Width           =   3255
   End
   Begin VB.ComboBox cboMonthName 
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
      Left            =   2190
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Select Name of Month"
      Top             =   4230
      Width           =   2415
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
      Left            =   2430
      TabIndex        =   11
      Top             =   5430
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto &Update"
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
      Left            =   4230
      TabIndex        =   10
      Top             =   5430
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Bonus Entry"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   32
      Top             =   5160
      Width           =   11775
   End
   Begin VB.Label Label84 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label lblmonth 
      Caption         =   "Month of Drawing"
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
      Height          =   255
      Left            =   150
      TabIndex        =   29
      Top             =   1230
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   1950
      TabIndex        =   28
      Top             =   1230
      Width           =   255
   End
   Begin VB.Label Label26 
      Caption         =   "Name of Bonus"
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
      Height          =   255
      Left            =   150
      TabIndex        =   27
      Top             =   4710
      Width           =   1335
   End
   Begin VB.Label Label25 
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
      Height          =   255
      Left            =   1950
      TabIndex        =   26
      Top             =   4230
      Width           =   255
   End
   Begin VB.Label lblFrom 
      Caption         =   "From Joining Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   25
      Top             =   2670
      Width           =   1695
   End
   Begin VB.Label lblFrom1 
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
      Height          =   255
      Index           =   0
      Left            =   1950
      TabIndex        =   24
      Top             =   2670
      Width           =   255
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   1950
      TabIndex        =   23
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "To Joining Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   22
      Top             =   3150
      Width           =   1455
   End
   Begin VB.Label Label60 
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
      Height          =   255
      Left            =   1950
      TabIndex        =   21
      Top             =   2190
      Width           =   255
   End
   Begin VB.Label Label59 
      Caption         =   "Sub Department"
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
      Height          =   255
      Left            =   150
      TabIndex        =   20
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label58 
      Caption         =   "Department"
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
      Height          =   255
      Left            =   150
      TabIndex        =   19
      Top             =   1710
      Width           =   1215
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
      Height          =   255
      Index           =   2
      Left            =   1950
      TabIndex        =   18
      Top             =   1710
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4830
      TabIndex        =   17
      Top             =   4305
      Width           =   495
   End
   Begin VB.Label lblFrom1 
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
      Height          =   255
      Index           =   1
      Left            =   6390
      TabIndex        =   16
      Top             =   4305
      Width           =   255
   End
   Begin VB.Label Label29 
      Caption         =   "Month of Payment"
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
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   4230
      Width           =   1575
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
      Height          =   255
      Index           =   0
      Left            =   1950
      TabIndex        =   12
      Top             =   4710
      Width           =   255
   End
End
Attribute VB_Name = "frmBonusEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDept_Click()
    If Len(cboDept) = 0 Then Exit Sub
    Call SubDeptName(cboSubDept, cboDept)
End Sub

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboDrawnMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboMonthName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSubDept_Click()
    CodeList.Clear
    txtTotal.Text = ""
End Sub

Private Sub cboSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()

Dim strFromDate As String, strToDate As String

If Len(cboDrawnMonth) = 0 Then
  MsgBox "Select Month of Drawing Salary", vbInformation, cnstMsgInfo
  cboDrawnMonth.SetFocus
  Exit Sub
End If
If Len(cboDept) = 0 Then
  MsgBox "Select Name of Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If
If Len(txtFromDate) = 0 Then
  MsgBox "Type From Joining Date", vbInformation, cnstMsgInfo
  txtFromDate.SetFocus
  Exit Sub
End If
If Len(txtToDate) = 0 Then
  MsgBox "Type To Joining Date", vbInformation, cnstMsgInfo
  txtToDate.SetFocus
  Exit Sub
End If

Screen.MousePointer = vbHourglass
Set r = New ADODB.Recordset

strFromDate = Format(txtFromDate, cnstDtFrmtQ)
strToDate = Format(txtToDate, cnstDtFrmtQ)

If Len(cboDept) <> 0 And Len(cboSubDept) = 0 Then
  If boJRMgt = True Then
    StrRecord = "SELECT Salary_And_Wages.EmployeeCode FROM Employee_Information INNER JOIN Salary_And_Wages ON (Employee_Information.EmployeeCode = Salary_And_Wages.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Salary_And_Wages.ComID) WHERE (((Salary_And_Wages.Month)='" & cboDrawnMonth.Text & "') AND ((Salary_And_Wages.Year)='" & PeriodYear & "')" _
              & " AND ((Salary_And_Wages.ComID)=" & CompanyID & ") AND ((Employee_Information.JoiningDate) BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND ((Employee_Information.WStatus)='J') AND ((Salary_And_Wages.DeptName)='" & cboDept.Text & "'));"
  Else
    StrRecord = "SELECT Salary_And_Wages.EmployeeCode FROM Employee_Information INNER JOIN Salary_And_Wages ON (Employee_Information.EmployeeCode = Salary_And_Wages.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Salary_And_Wages.ComID) WHERE (((Salary_And_Wages.Month)='" & cboDrawnMonth.Text & "') AND ((Salary_And_Wages.Year)='" & PeriodYear & "')" _
              & " AND ((Salary_And_Wages.ComID)=" & CompanyID & ") AND ((Employee_Information.JoiningDate) BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND ((Employee_Information.WStatus)='W') AND ((Salary_And_Wages.DeptName)='" & cboDept.Text & "'));"
  End If
ElseIf Len(cboDept) <> 0 And Len(cboSubDept) <> 0 Then
  If boJRMgt = True Then
    StrRecord = "SELECT Salary_And_Wages.EmployeeCode FROM Employee_Information INNER JOIN Salary_And_Wages ON (Employee_Information.EmployeeCode = Salary_And_Wages.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Salary_And_Wages.ComID) WHERE (((Salary_And_Wages.Month)='" & cboDrawnMonth.Text & "') AND ((Salary_And_Wages.Year)='" & PeriodYear & "')" _
              & " AND ((Salary_And_Wages.ComID)= " & CompanyID & ") AND ((Employee_Information.JoiningDate) BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND ((Employee_Information.WStatus)='J') AND ((Salary_And_Wages.DeptName)='" & cboDept.Text & "') AND ((Salary_And_Wages.SubDeptName)='" & cboSubDept.Text & "'));"
  Else
    StrRecord = "SELECT Salary_And_Wages.EmployeeCode FROM Employee_Information INNER JOIN Salary_And_Wages ON (Employee_Information.EmployeeCode = Salary_And_Wages.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Salary_And_Wages.ComID) WHERE (((Salary_And_Wages.Month)='" & cboDrawnMonth.Text & "') AND ((Salary_And_Wages.Year)='" & PeriodYear & "')" _
              & " AND ((Salary_And_Wages.ComID)= " & CompanyID & ") AND ((Employee_Information.JoiningDate) BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND ((Employee_Information.WStatus)='W') AND ((Salary_And_Wages.DeptName)='" & cboDept.Text & "') AND ((Salary_And_Wages.SubDeptName)='" & cboSubDept.Text & "'));"
  End If
End If
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  Do Until r.EOF
    CodeList.AddItem r![EmployeeCode]
    txtTotal.Text = CStr(CodeList.ListCount)
    r.MoveNext
  Loop
End If
r.Close

Screen.MousePointer = vbDefault
Set r = Nothing
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdMLeave_Click()

Dim strFromDate As String, strToDate As String

If Len(cboDept) = 0 Then
  MsgBox "Select Name of Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If
If Len(txtFromDate) = 0 Then
  MsgBox "Type From Joining Date", vbInformation, cnstMsgInfo
  txtFromDate.SetFocus
  Exit Sub
End If
If Len(txtToDate) = 0 Then
  MsgBox "Type To Joining Date", vbInformation, cnstMsgInfo
  txtToDate.SetFocus
  Exit Sub
End If

Screen.MousePointer = vbHourglass
Set r = New ADODB.Recordset

strFromDate = Format(txtFromDate, cnstDtFrmtQ)
strToDate = Format(txtToDate, cnstDtFrmtQ)

''StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (JoiningDate BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND (WStatus='W') AND (MLeave=1) AND (MLeft=0));"
If Len(cboDept) <> 0 And Len(cboSubDept) = 0 Then
  If boJRMgt = True Then
    StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (JoiningDate BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND (WStatus='J') AND (MLeave=1) AND (MLeft=0) AND (Department='" & cboDept.Text & "'));"
  Else
    StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (JoiningDate BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND (WStatus='W') AND (MLeave=1) AND (MLeft=0) AND (Department='" & cboDept.Text & "'));"
  End If

ElseIf Len(cboDept) <> 0 And Len(cboSubDept) <> 0 Then
  If boJRMgt = True Then
    StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (JoiningDate BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND (WStatus='J') AND (MLeave=1) AND (MLeft=0) AND (Department='" & cboDept.Text & "') AND ([Section]='" & cboSubDept.Text & "'));"
  Else
    StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (JoiningDate BETWEEN '" & strFromDate & "' AND '" & strToDate & "')" _
              & " AND (WStatus='W') AND (MLeave=1) AND (MLeft=0) AND (Department='" & cboDept.Text & "') AND ([Section]='" & cboSubDept.Text & "'));"
  End If
End If
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  Do Until r.EOF
    CodeList.AddItem r![EmployeeCode]
    txtTotal.Text = CStr(CodeList.ListCount)
    r.MoveNext
  Loop
End If
r.Close
Screen.MousePointer = vbDefault
Set r = Nothing
End Sub
Private Sub cmdRemove_Click()
If CodeList.ListIndex >= 0 Then
  CodeList.RemoveItem (CodeList.ListIndex)
  txtTotal.Text = CStr(CodeList.ListCount)
End If
End Sub
Private Sub cmdRemoveAll_Click()
  CodeList.Clear
  txtTotal.Text = ""
End Sub
Private Sub cmdSave_Click()

Dim sngRateBonus As Single, curAmount As Currency, curFracPayD As Currency
If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If Len(txtHeader) = 0 Then
  MsgBox "Type Name of Bonus", vbInformation, cnstMsgInfo
  txtHeader.SetFocus
  Exit Sub
End If
If Len(cboMonthName) = 0 Then
  MsgBox "Select Payment Month", vbInformation, cnstMsgInfo
  cboMonthName.SetFocus
  Exit Sub
End If
If Len(cboDept) = 0 Then
  MsgBox "Select Name of Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If
If CodeList.ListCount = 0 Then
  MsgBox "Record not found to Update", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If

If MsgBox("Are you sure to Update Bonus", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
sngRateBonus = rateBonus

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set r = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500
DeleteRow = 2
MainConn.BeginTrans

curFracPayD = 0
        
For intI = 1 To CodeList.ListCount
  CodeList.ListIndex = CodeList.ListIndex + 1
  
  'MainComm.CommandText = "SET NOCOUNT ON Insert Into Bonus_Entry_DelEdit Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "') AND ([Month]='" & cboMonthName.Text & "') AND ([Year]='" & PeriodYear & "'));"
  'MainComm.Execute
  
  MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "') AND ([Month]='" & cboMonthName.Text & "') AND ([Year]='" & PeriodYear & "'));"
  MainComm.Execute
  
  StrRecord = "SELECT CurrRate,Department,[Section] FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "'));"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    If boJRMgt = True Then
      curAmount = CLng((r![CurrRate] / 100 * 50) * sngRateBonus)
    Else
      curAmount = CLng((r![CurrRate] / 100 * V_rateBasic) * sngRateBonus)
    End If
    If curAmount > 0 Then curFracPayD = curAmount Mod 10
    MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Bonus_Entry VALUES(" & CompanyID & ",'" & CodeList.Text & "','" & cboMonthName.Text & "','" & PeriodYear & " '," _
          & " " & r![CurrRate] & "," & curAmount - curFracPayD & ",0,0,'" & r![Department] & "','" & r![Section] & "','" & txtHeader.Text & "'," & curFracPayD & ");"
    MainComm.Execute
  End If
  r.Close
Next intI

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  Err.Number = 0
  Exit Sub
End If

MsgBox "Bonus Update Successfully", vbInformation, cnstMsgInfo

Screen.MousePointer = vbDefault

CodeList.Clear
txtTotal.Text = ""
Set r = Nothing
End Sub
Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Festival Bonus Entry"
  If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"

  prcMonthName cboDrawnMonth
  cboDrawnMonth.Text = UCase(Format(Date, "MMMM"))
  prcMonthName cboMonthName
  cboMonthName.Text = UCase(Format(Date, "MMMM"))
  DeptName cboDept
End Sub
Private Sub Form_Unload(Cancel As Integer)
  boJRMgt = False
  MDIfrmMain.lblCaption.Caption = ""
  Set frmBonusEntry = Nothing
End Sub
Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtFromDate_LostFocus()
  txtFromDate.Text = FormatDate(txtFromDate)
End Sub
Private Sub txtHeader_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtHeader_LostFocus()
  txtHeader.Text = UCase(txtHeader.Text)
End Sub
Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtToDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtToDate_LostFocus()
  txtToDate.Text = FormatDate(txtToDate)
End Sub
