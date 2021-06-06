VERSION 5.00
Begin VB.Form frmBonusPaidEntry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8333
   ClientLeft      =   2041
   ClientTop       =   2431
   ClientWidth     =   11726
   ControlBox      =   0   'False
   Icon            =   "frmBonusPaidEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8333
   ScaleWidth      =   11726
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChKALL 
      Caption         =   "All Employee"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   810
      TabIndex        =   31
      Top             =   6180
      Width           =   4035
      Begin VB.CommandButton cmdAddPresent 
         Caption         =   "Add Present Emp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2070
         TabIndex        =   33
         Top             =   330
         Width           =   1755
      End
      Begin VB.TextBox txtDate 
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
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "Present Date  : "
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
         Left            =   120
         TabIndex        =   34
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.Frame FrameCategory 
      Caption         =   " Employee Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2340
      TabIndex        =   22
      Tag             =   "W"
      Top             =   1110
      Width           =   2655
      Begin VB.OptionButton optStaff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optWorker 
         Caption         =   "Workers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Tag             =   "W"
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Emp Code Department wise"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   2460
      TabIndex        =   21
      Top             =   4110
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   2790
      TabIndex        =   19
      Top             =   4500
      Width           =   2055
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0C000&
         Caption         =   "&Add to List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   180
         TabIndex        =   30
         Top             =   180
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C0C000&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   180
         TabIndex        =   29
         Top             =   630
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveAll 
         BackColor       =   &H00C0C000&
         Caption         =   "Remove &All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   180
         TabIndex        =   28
         Top             =   1110
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1665
      Left            =   810
      TabIndex        =   20
      Top             =   4500
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000013&
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
         Height          =   415
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1110
         Width           =   1695
      End
      Begin VB.CommandButton cmdUnPaid 
         BackColor       =   &H80000013&
         Caption         =   "&Unpaid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   630
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000013&
         Caption         =   "&Paid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.ComboBox cboFromCode 
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
      Left            =   2460
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Select Employee Code No."
      Top             =   3270
      Width           =   2415
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
      Left            =   2460
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select Sub Department"
      Top             =   2790
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
      Left            =   2460
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Name of Department"
      Top             =   2310
      Width           =   2415
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   10
      Top             =   7080
      Width           =   1935
   End
   Begin VB.ComboBox cboEmpCode 
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
      Left            =   2460
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Select Employee Code No."
      Top             =   3750
      Width           =   2415
   End
   Begin VB.ListBox CodeList 
      BackColor       =   &H80000016&
      Columns         =   5
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5668
      ItemData        =   "frmBonusPaidEntry.frx":0442
      Left            =   5040
      List            =   "frmBonusPaidEntry.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   6495
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
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Month"
      Top             =   1830
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18.34
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   36
      Top             =   7440
      Width           =   11775
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11775
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
      Left            =   2220
      TabIndex        =   18
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label lblFrom 
      Caption         =   "From Emp Code"
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
      Left            =   420
      TabIndex        =   17
      Top             =   3270
      Width           =   1335
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
      Left            =   420
      TabIndex        =   16
      Top             =   2790
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
      Left            =   2220
      TabIndex        =   15
      Top             =   2790
      Width           =   255
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
      Left            =   2220
      TabIndex        =   14
      Top             =   2310
      Width           =   255
   End
   Begin VB.Label Label58 
      Caption         =   "Name of Department"
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
      Left            =   420
      TabIndex        =   13
      Top             =   2310
      Width           =   1815
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
      Left            =   8520
      TabIndex        =   12
      Top             =   7080
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
      Left            =   9360
      TabIndex        =   11
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "To Emp Code"
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
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   3750
      Width           =   1335
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
      Left            =   2220
      TabIndex        =   8
      Top             =   3750
      Width           =   255
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
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   1830
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
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   5
      Top             =   1830
      Width           =   255
   End
End
Attribute VB_Name = "frmBonusPaidEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDept_Click()
  If Len(cboDept) = 0 Then Exit Sub
  SubDeptName cboSubDept, cboDept
  chkShow.Enabled = True
End Sub
Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cboEmpCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdAdd.SetFocus
End Sub
Private Sub cboEmpCode_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub cboFromCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cboFromCode_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub cboFromCode_LostFocus()
  cboEmpCode.Text = cboFromCode.Text
End Sub
Private Sub cboMonthName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cboSubDept_Click()
  cboFromCode.Clear
  cboEmpCode.Clear
End Sub
Private Sub cboSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cboSubDept_LostFocus()
    If Len(cboMonthName) <> 0 And Len(cboDept) <> 0 And Len(cboSubDept) <> 0 Then
        If boJRMgt = True And boSRMgt = True Then
            empcode cboFromCode, cboDept.Text, cboSubDept.Text, ""
            empcode cboEmpCode, cboDept.Text, cboSubDept.Text, ""
''        ElseIf boJRMgt = True Then
''            EmpCode cboFromCode, cboDept.Text, cboSubDept.Text, "J"
''            EmpCode cboEmpCode, cboDept.Text, cboSubDept.Text, "J"
        Else
            empcode cboFromCode, cboDept.Text, cboSubDept.Text, FrameCategory.Tag
            empcode cboEmpCode, cboDept.Text, cboSubDept.Text, FrameCategory.Tag
        End If
    End If
End Sub
Private Sub chkShow_Click()
  If chkShow.Value = 0 Then Exit Sub
  cboFromCode.Clear
  cboEmpCode.Clear
  Screen.MousePointer = vbHourglass
  Set r = New ADODB.Recordset
  If boJRMgt = True Then
    StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Employee_Information INNER JOIN Bonus_Entry ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
              & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "') AND ((Employee_Information.WStatus)='J')) GROUP BY Bonus_Entry.EmployeeCode;"
  Else
    StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Employee_Information INNER JOIN Bonus_Entry ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
              & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
              & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "') AND ((Employee_Information.WStatus)='" & FrameCategory.Tag & "')) GROUP BY Bonus_Entry.EmployeeCode;"
  End If
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    Do Until r.EOF
      cboFromCode.AddItem r![EmployeeCode]
      cboEmpCode.AddItem r![EmployeeCode]
      r.MoveNext
    Loop
  End If
  r.Close
  Screen.MousePointer = vbDefault
  Set r = Nothing
  chkShow.Value = 0
End Sub
Private Sub cmdAdd_Click()
Screen.MousePointer = vbHourglass
Set r = New ADODB.Recordset

If ChKALL.Value = 1 Then

StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Bonus_Entry.EmployeeCode = Employee_Information.EmployeeCode)" _
          & " AND (Bonus_Entry.ComID = Employee_Information.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
          & " AND ((Bonus_Entry.Year)='" & PeriodYear & "') AND ((Bonus_Entry.HeldUp)=0) AND ((Employee_Information.WStatus)='" & FrameCategory.Tag & "'))" _
          & " GROUP BY Bonus_Entry.EmployeeCode ;"

Else

StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Bonus_Entry.EmployeeCode = Employee_Information.EmployeeCode)" _
          & " AND (Bonus_Entry.ComID = Employee_Information.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
          & " AND ((Bonus_Entry.Year)='" & PeriodYear & "') AND ((Bonus_Entry.HeldUp)=0) AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "')AND ((Bonus_Entry.SubDeptName)='" & cboSubDept.Text & "') AND ((Employee_Information.WStatus)='" & FrameCategory.Tag & "'))" _
          & " GROUP BY Bonus_Entry.EmployeeCode HAVING (((Bonus_Entry.EmployeeCode) Between '" & cboFromCode.Text & "' And '" & cboEmpCode.Text & "'));"


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
Private Sub cmdAddPresent_Click()
If Len(txtDate) = 0 Then
  MsgBox "Enter Date", vbInformation, cnstMsgInfo
  txtDate.SetFocus
  Exit Sub
End If

'If Len(cboSubDept) = 0 Then
'  MsgBox "Select Sub Dept.", vbInformation, cnstMsgInfo
'  cboSubDept.SetFocus
'  Exit Sub
'End If

CodeList.Clear
txtTotal.Text = ""
Screen.MousePointer = vbHourglass
Set r = New ADODB.Recordset
If boJRMgt = True Then
  StrRecord = "SELECT Emp_Daily_Attendance.EmpCode FROM Employee_Information INNER JOIN Emp_Daily_Attendance ON Employee_Information.ComID = Emp_Daily_Attendance.ComID" _
            & " AND Employee_Information.EmployeeCode = Emp_Daily_Attendance.EmpCode WHERE ((Employee_Information.ComID =" & CompanyID & ") AND (Employee_Information.MLeft = 0)" _
            & " AND (Employee_Information.WStatus = 'J') AND (Employee_Information.MLeave = 0) " _
            & "  AND (Emp_Daily_Attendance.DailyDate = '" & Format(txtDate.Text, cnstDtFrmtQ) & "') AND (Emp_Daily_Attendance.TimeIn <> '00:00'));"
Else
  StrRecord = "SELECT Emp_Daily_Attendance.EmpCode FROM Employee_Information INNER JOIN Emp_Daily_Attendance ON Employee_Information.ComID = Emp_Daily_Attendance.ComID" _
            & " AND Employee_Information.EmployeeCode = Emp_Daily_Attendance.EmpCode WHERE ((Employee_Information.ComID =" & CompanyID & ") AND (Employee_Information.MLeft = 0)" _
            & " AND (Employee_Information.WStatus = '" & FrameCategory.Tag & "') AND (Employee_Information.MLeave = 0) " _
            & "  AND (Emp_Daily_Attendance.DailyDate = '" & Format(txtDate.Text, cnstDtFrmtQ) & "') AND (Emp_Daily_Attendance.TimeIn <> '00:00'));"
End If
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  Do Until r.EOF
    CodeList.AddItem r![empcode]
    txtTotal.Text = CStr(CodeList.ListCount)
    r.MoveNext
  Loop
Else
  MsgBox "No record found", vbInformation, cnstMsgInfo
End If
r.Close
Set r = Nothing
Screen.MousePointer = vbDefault
End Sub
Private Sub cmdClose_Click()
  Unload Me
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
    DeleteRow = 2
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
        MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
        Exit Sub
    End If
If CodeList.ListCount = 0 Then
  MsgBox "Record not found to Paid", vbInformation, cnstMsgInfo
  cboEmpCode.SetFocus
  Exit Sub
End If

If MsgBox("Are you sure to Paid the bonus of existing employee", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

CodeList.ListIndex = -1
For intI = 1 To CodeList.ListCount
  CodeList.ListIndex = CodeList.ListIndex + 1

'  ComID=" & CompanyID & " AND
MainComm.CommandText = "SET NOCOUNT ON INSERT INTO  Bonus_Entry_DelEdit " _
                            & " Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus, " & DeleteRow & " ," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE()From Bonus_Entry " _
                            & " WHERE ComID=" & CompanyID & " AND EmployeeCode='" & CodeList.Text & "'AND Month ='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND Paid = 0;"
        
        MainComm.Execute
  
  MainComm.CommandText = "SET NOCOUNT ON UPDATE Bonus_Entry SET Paid = 1,DeleteRow=" & DeleteRow & ",UserId=" & bytUserID & ",IDNo=" & getComVolID & ",EntryDate='" & Date & "',EntryTime=GETDATE() WHERE ComID=" & CompanyID & "" _
      & " AND EmployeeCode='" & CodeList.Text & "'AND Month ='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND Paid = 0;"
  MainComm.Execute
Next intI

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  Err.Number = 0
  Exit Sub
End If

CodeList.Clear
txtTotal.Text = ""

Screen.MousePointer = vbDefault
End Sub
Private Sub cmdUnPaid_Click()
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
    Exit Sub
    End If
If CodeList.ListCount = 0 Then
  MsgBox "Record not found to Unpaid", vbInformation, cnstMsgInfo
  cboEmpCode.SetFocus
  Exit Sub
End If

If MsgBox("Are you sure to Unpaid the bonus of existing employee", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

CodeList.ListIndex = -1
For intI = 1 To CodeList.ListCount
  CodeList.ListIndex = CodeList.ListIndex + 1
  
MainComm.CommandText = "SET NOCOUNT ON INSERT INTO  Bonus_Entry_DelEdit " _
                            & " Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus,0," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() From Bonus_Entry " _
                            & " WHERE ComID=" & CompanyID & " AND EmployeeCode='" & CodeList.Text & "'AND Month ='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND Paid = 1;"
        
        MainComm.Execute
  
  MainComm.CommandText = "SET NOCOUNT ON UPDATE Bonus_Entry SET Paid = 0,DeleteRow=" & DeleteRow & ",UserId=" & bytUserID & ",IDNo=" & getComVolID & ",EntryDate='" & Date & "',EntryTime=GETDATE() WHERE ComID=" & CompanyID & "" _
      & " AND EmployeeCode='" & CodeList.Text & "'AND Month ='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND Paid = 1 ;"
  MainComm.Execute
Next intI

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  Err.Number = 0
  Exit Sub
End If

CodeList.Clear
txtTotal.Text = ""

Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
'  Me.Height = 7020
'  Me.width = 11940
  'Me.Left = 0
  'Me.Top = 0
    Top = (frmBonusPaidEntry.Height - Height) / 2
    Left = (frmBonusPaidEntry.width - width) / 2 + 1600
  
  MDIfrmMain.lblCaption.Caption = "Bonus Paid / Unpaid Entry"
  prcMonthName cboMonthName
  cboMonthName.Text = UCase(Format(Date, "MMMM"))
  DeptName cboDept
  
  Set r = New ADODB.Recordset
    StrRecord = "SELECT Worker,JRmgt,SRmgt,Staff FROM Sys_User_Name WHERE Userid='" & bytUserID & "' ;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r![Worker] = 1 Then
    optWorker.Enabled = True
    Else
    optWorker.Enabled = False
    End If
    If r![Staff] = 1 Then
    optStaff.Enabled = True
    Else
    optStaff.Enabled = False
    End If
   r.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmBonusPaidEntry = Nothing
End Sub

Private Sub optStaff_Click()
    FrameCategory.Tag = "F"
End Sub

Private Sub optWorker_Click()
    FrameCategory.Tag = "W"
End Sub

Private Sub txtDate_LostFocus()
    Call FormatDate(txtDate)
End Sub


