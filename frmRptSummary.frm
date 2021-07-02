VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRptSummary 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7852
   ClientLeft      =   3276
   ClientTop       =   3003
   ClientWidth     =   6825
   ControlBox      =   0   'False
   Icon            =   "frmRptSummary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7852
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   240
      TabIndex        =   10
      Top             =   1620
      Width           =   6375
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
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Select Name of Month"
         Top             =   240
         Width           =   3525
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
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   660
         Width           =   2775
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
         Left            =   2070
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2745
      End
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
         Height          =   345
         Left            =   4920
         TabIndex        =   13
         Top             =   1080
         Width           =   675
      End
      Begin VB.CheckBox chkAllSDept 
         Caption         =   "Add all Sub Dept. Under Selected Dept."
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
         Height          =   225
         Left            =   2070
         TabIndex        =   12
         Top             =   1530
         Width           =   3195
      End
      Begin VB.CommandButton cmdAddDept 
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
         Height          =   345
         Left            =   4920
         TabIndex        =   11
         Top             =   660
         Width           =   675
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
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDept 
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
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   19
         Top             =   240
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
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   660
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
         Left            =   1800
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   885
      Left            =   3090
      TabIndex        =   2
      Top             =   720
      Width           =   3525
      Begin VB.OptionButton optSubDept 
         Caption         =   "Sub Department Wise"
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
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   510
         Width           =   2205
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton OptDept 
         Caption         =   "Department Wise"
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
         Left            =   960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame FrameCategory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   240
      TabIndex        =   27
      Tag             =   "W"
      Top             =   990
      Width           =   2895
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
         Left            =   1860
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   825
      End
      Begin VB.OptionButton optWorkForce 
         Caption         =   "Work Force"
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
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   240
      TabIndex        =   23
      Top             =   5837
      Width           =   6375
      Begin VB.CheckBox chkJun16 
         Caption         =   "June 2016"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   3016
         TabIndex        =   33
         Top             =   858
         Visible         =   0   'False
         Width           =   1105
      End
      Begin VB.CheckBox chkPrev 
         Caption         =   "Show ML Employee Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   4134
         TabIndex        =   32
         Top             =   260
         Width           =   2067
      End
      Begin VB.CheckBox chkBankPay 
         Caption         =   "Bank Payment Summary"
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
         Height          =   255
         Left            =   4134
         TabIndex        =   31
         Top             =   858
         Width           =   2028
      End
      Begin VB.CheckBox chkCashPay 
         Caption         =   "Cash Payment Summary"
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
         Height          =   255
         Left            =   4134
         TabIndex        =   30
         Top             =   540
         Width           =   2028
      End
      Begin VB.CheckBox chkRest 
         Caption         =   "Show All Dept. && Sub Dept. without List added."
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
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "Show Dept. Summary (List added)"
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
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   540
         Width           =   2805
      End
      Begin VB.CheckBox chkDesig 
         Caption         =   "Show Designation wise Summary"
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
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "&Preview"
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
      Left            =   2808
      TabIndex        =   0
      Top             =   7319
      Width           =   1305
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
      Height          =   405
      Left            =   4440
      TabIndex        =   1
      Top             =   7319
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   2370
      Left            =   240
      TabIndex        =   6
      Top             =   3510
      Width           =   6375
      _ExtentX        =   11238
      _ExtentY        =   4169
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   285
      BackColor       =   -2147483634
      ForeColor       =   12582912
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   16776960
      BackColorSel    =   -2147483642
      ForeColorSel    =   -2147483633
      BackColorBkg    =   -2147483626
      GridColor       =   -2147483630
      GridColorFixed  =   -2147483630
      AllowBigSelection=   -1  'True
      ScrollBars      =   2
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
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Height          =   702
      Left            =   0
      TabIndex        =   9
      Top             =   7163
      Width           =   7098
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Salary Summary Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   6600
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmRptSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDept_Click()
  If Len(cboDept) = 0 Then Exit Sub
  SubDeptName cboSubDept, cboDept
End Sub
Private Sub cmdAdd_Click()
If Len(cboSubDept) = 0 And chkAllSDept.Value = 0 Then
  MsgBox "Select Sub Department", vbInformation, cnstMsgInfo
  cboSubDept.SetFocus
  Exit Sub
End If

If chkAllSDept.Value = 1 Then
  Set r = New ADODB.Recordset
  StrRecord = "SELECT * FROM Sub_Dept_Name WHERE ComID=" & CompanyID & " AND DeptName='" & cboDept.Text & "';"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    Do Until r.EOF
      With flexDetails
        .Rows = .Rows + 1
        .row = .Rows - 1
        .col = 0
        .Text = cboDept.Text
        .col = 1
        .Text = r![SubDeptName]
      End With
      r.MoveNext
    Loop
  End If
  r.Close
  Set r = Nothing
Else
  With flexDetails
      .Rows = .Rows + 1
      .row = .Rows - 1
      .col = 0
      .Text = cboDept.Text
      .col = 1
      .Text = cboSubDept.Text
  End With
  cboSubDept.SetFocus
End If
End Sub
Private Sub cmdAddDept_Click()
If Len(cboDept) = 0 Then
  MsgBox "Select Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If
With flexDetails
  .Rows = .Rows + 1
  .row = .Rows - 1
  .col = 0
  .Text = cboDept.Text
End With
cboDept.SetFocus
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdPrint_Click()
Dim strDept As String, strSubDept As String, strBankCash As String, tblName As String
    If Len(cboMonthName) = 0 Then
        MsgBox "Select Name of Month", vbInformation, cnstMsgInfo
        cboMonthName.SetFocus
        Exit Sub
    End If
    If Optdept.Value = True And Len(cboDept) = 0 Then
        MsgBox "Select Department", vbInformation, cnstMsgInfo
        cboDept.SetFocus
        Exit Sub
    End If
    If optSubDept.Value = True And flexDetails.Rows = 1 Then
        MsgBox "Add Sub Department", vbInformation, cnstMsgInfo
        cboSubDept.SetFocus
        Exit Sub
    End If
    
    If chkPrev.Value = vbChecked Then
      tblName = "Salary_And_Wages_ML"
    ElseIf chkJun16.Value = vbChecked Then
      tblName = "Salary_And_Wages_June16"
    Else
      tblName = "Salary_And_Wages"
    End If

    If chkCashPay.Value = 1 Then
      strBankCash = " AND " & tblName & ".EmployeeCode NOT IN (SELECT DISTINCT Empcode FROM Emp_Account WHERE (AccNo IS NOT NULL AND AccNo <> '' AND AccNo <> '=N/A=') AND SUBSTRING(HeldUp,1,1) = 'N')"
    ElseIf chkBankPay.Value = 1 Then
      strBankCash = " AND " & tblName & ".EmployeeCode IN (SELECT DISTINCT Empcode FROM Emp_Account WHERE (AccNo IS NOT NULL AND AccNo <> '' AND AccNo <> '=N/A=') AND SUBSTRING(HeldUp,1,1) = 'N')"
    Else
      strBankCash = ""
    End If
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set r = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    MainConn.BeginTrans
    
    MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Salary_And_Wages;"
    MainComm.Execute
    MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Summary;"
    MainComm.Execute
  
    If chkDesig.Value = 1 Then
        If OptAll.Value = True Then
              MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Salary_And_Wages(One,Year,Month,EmployeeCode,Designation,RatePay,Basic,AttenDays,AbsentDay,OTHours,WEarned,OTPay,FoodPay,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,ComName,Address) SELECT " & tblName & ".ComID," & tblName & ".Year," & tblName & ".Month," & tblName & ".EmployeeCode,Employee_Information.Designation," & tblName & ".RatePay," & tblName & ".Basic," _
                & " " & tblName & ".AttenDays," & tblName & ".AbsentDay," & tblName & ".OTHours," & tblName & ".WEarned," & tblName & ".OTPay," & tblName & ".FoodPay," & tblName & ".FoodAllow," & tblName & ".AttenBonus," & tblName & ".OthersE," _
                & " " & tblName & ".AbsAmount," & tblName & ".PIB_Days," & tblName & ".PIB," & tblName & ".TransportD," & tblName & ".OthersD," & tblName & ".NetPayTK," & tblName & ".FracPay," & tblName & ".FracPayD," & tblName & ".TransAllow," & tblName & ".DinnerAllow," & tblName & ".TransDay," & tblName & ".TransPay," & tblName & ".DeptName," _
                & " " & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Employee_Information INNER JOIN (Company_Information INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID)" _
                & " ON (Employee_Information.ComID = " & tblName & ".ComID) AND (Employee_Information.EmployeeCode = " & tblName & ".EmployeeCode) WHERE (((" & tblName & ".Year) = '" & PeriodYear & "') AND ((" & tblName & ".Month) ='" & cboMonthName.Text & "')" _
                & " AND ((" & tblName & ".HeldUp) = 0) AND ((" & tblName & ".ComID) = " & CompanyID & ") AND ((" & tblName & ".WorkDays) = 26)) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " ORDER BY " & tblName & ".EmployeeCode;"
        ElseIf Optdept.Value = True Then
              MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Salary_And_Wages(One,Year,Month,EmployeeCode,Designation,RatePay,Basic,AttenDays,AbsentDay,OTHours,WEarned,OTPay,FoodPay,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,ComName,Address) SELECT " & tblName & ".ComID," & tblName & ".Year," & tblName & ".Month," & tblName & ".EmployeeCode,Employee_Information.Designation," & tblName & ".RatePay," & tblName & ".Basic," _
                & " " & tblName & ".AttenDays," & tblName & ".AbsentDay," & tblName & ".OTHours," & tblName & ".WEarned," & tblName & ".OTPay," & tblName & ".FoodPay," & tblName & ".FoodAllow," & tblName & ".AttenBonus," & tblName & ".OthersE," _
                & " " & tblName & ".AbsAmount," & tblName & ".PIB_Days," & tblName & ".PIB," & tblName & ".TransportD," & tblName & ".OthersD," & tblName & ".NetPayTK," & tblName & ".FracPay," & tblName & ".FracPayD," & tblName & ".TransAllow," & tblName & ".DinnerAllow," & tblName & ".TransDay," & tblName & ".TransPay," & tblName & ".DeptName," _
                & " " & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Employee_Information INNER JOIN (Company_Information INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID)" _
                & " ON (Employee_Information.ComID = " & tblName & ".ComID) AND (Employee_Information.EmployeeCode = " & tblName & ".EmployeeCode) WHERE (((" & tblName & ".Year) = '" & PeriodYear & "') AND ((" & tblName & ".Month) ='" & cboMonthName.Text & "')" _
                & " AND ((" & tblName & ".DeptName) ='" & cboDept.Text & "') AND ((" & tblName & ".HeldUp) = 0) AND ((" & tblName & ".ComID) = " & CompanyID & ") AND ((" & tblName & ".WorkDays) = 26)) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " ORDER BY " & tblName & ".EmployeeCode;"
        End If
        MainComm.Execute
        
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Salary_And_Wages SET Temp_Salary_And_Wages.EmpName = Designation.Category FROM Temp_Salary_And_Wages INNER JOIN Designation" _
              & " ON Temp_Salary_And_Wages.Designation = Designation.Designation AND Temp_Salary_And_Wages.One = Designation.ComID WHERE Designation.Category IS NOT NULL;"
        MainComm.Execute
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Salary_And_Wages SET Designation=[EmpName] WHERE EmpName IS NOT NULL;"
        MainComm.Execute
      
        If OptAll.Value = True Then
            'MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Salary_And_Wages WHERE DeptName='" & cboDept.Text & "';" 'For WET PROCESSING
            'MainComm.Execute
            MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Salary_And_Wages SET DeptName = 'ALL DEPT. ';" '(WITHOUT WET PROCESSING)';"
            MainComm.Execute
        End If
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Salary_And_Wages SET SubDeptName = 'ALL SUB DEPT. (WORKFORCE)';"
        MainComm.Execute
        
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,SubSubDept,ComName,Address)SELECT Count(EmployeeCode) AS TotalEmp,[Month],[Year],Sum(RatePay) AS TotalGross,Sum(Basic) AS TotalBasic,Sum(AttenDays) AS TotalAttenDays,Sum(AbsentDay) AS TotalAbsDays," _
                & " Sum(OTHours) AS TotalOTHours,Sum(WEarned) AS TotalWEarned,Sum(OTPay) AS TotalOTPay,Sum(FoodPay) AS TotalFoodPay,Sum(FoodAllow) AS TotalFoodAllow,Sum(AttenBonus) AS TotalAttenBonus,Sum(OthersE) AS TotalOthersE,Sum(AbsAmount) AS TotalAbsAmount," _
                & " Sum(TransportD) AS TotalTransportD,Sum(OthersD) AS TotalOthersD,Sum(NetPayTK) AS TotalNetPayTK,Sum(FracPay) AS TotalFracPay,Sum(FracPayD) AS TotalFracPayD,Sum(TransAllow) AS TransAll,Sum(DinnerAllow) AS DinnerAll,Sum(TransDay) AS TTransDay,Sum(TransPay) AS TTransPay,DeptName,SubDeptName," _
                & " Designation,ComName,Address FROM Temp_Salary_And_Wages GROUP BY [Month],[Year],DeptName,SubDeptName,Designation,ComName,Address;"
        MainComm.Execute
        
        StrRecord = "SELECT DeptName,SubDeptName,SubSubDept,([WEarned]+[OTAmount]+[FoodAmount]+[FoodAllow]+[AttenBonus]+[TransAllow]+[DinnerAllow]+[OthersE]+[FracPay]+[AbsAmount]+[PIB]+[TransPay]) AS TotalAmount FROM Temp_Summary;"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET TotalAmount= " & r![TotalAmount] & " WHERE DeptName='" & r![DeptName] & "' AND SubDeptName='" & r![SubDeptName] & "' AND SubSubDept='" & r![SubSubDept] & "';"
                MainComm.Execute
                r.MoveNext
            Loop
        End If
        r.Close
    
        GoTo goCommit
    End If
    
    If OptAll.Value = True Then
        If chkDept.Value = 1 Then
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,ComName,Address)" _
                & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays," _
                & " Sum(" & tblName & ".AbsentDay) AS TotalAbsDays,Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay," _
                & " Sum(" & tblName & ".FoodPay) AS TotalFoodPay,Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE," _
                & " Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB,Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan," _
                & " Sum(" & tblName & ".Interest) AS TotalInterest,Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay," _
                & " Sum(" & tblName & ".FracPayD) AS TotalFracPayD,Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH," _
                & " Sum(" & tblName & ".OneH) AS TotalOneH,Sum(" & tblName & ".Fifty) AS TotalFifty,Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive," _
                & " Sum(" & tblName & ".Two) AS TotalTwo,Sum(" & tblName & ".One) AS TotalOne," & tblName & ".DeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN " & tblName & "" _
                & " ON Company_Information.ID = " & tblName & ".ComID WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".HeldUp) =0)" _
                & " AND ((" & tblName & ".ComID) =" & CompanyID & ")) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
            MainComm.Execute
        Else
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays," _
                & " Sum(" & tblName & ".AbsentDay) AS TotalAbsDays,Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay," _
                & " Sum(" & tblName & ".FoodPay) AS TotalFoodPay,Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE," _
                & " Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB,Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan," _
                & " Sum(" & tblName & ".Interest) AS TotalInterest,Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay," _
                & " Sum(" & tblName & ".FracPayD) AS TotalFracPayD,Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH," _
                & " Sum(" & tblName & ".OneH) AS TotalOneH,Sum(" & tblName & ".Fifty) AS TotalFifty,Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive," _
                & " Sum(" & tblName & ".Two) AS TotalTwo,Sum(" & tblName & ".One) AS TotalOne," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information" _
                & " INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "')" _
                & " AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & "))AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName," & tblName & ".SubDeptName," _
                & " Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
            MainComm.Execute
        End If
    ElseIf Optdept.Value = True Then
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays,Sum(" & tblName & ".AbsentDay) AS TotalAbsDays," _
                & " Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay,Sum(" & tblName & ".FoodPay) AS TotalFoodPay," _
                & " Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE,Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB," _
                & " Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan,Sum(" & tblName & ".Interest) AS TotalInterest," _
                & " Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay,Sum(" & tblName & ".FracPayD) AS TotalFracPayD," _
                & " Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH,Sum(" & tblName & ".OneH) AS TotalOneH,Sum(" & tblName & ".Fifty) AS TotalFifty," _
                & " Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive,Sum(" & tblName & ".Two) AS TotalTwo,Sum(" & tblName & ".One) AS TotalOne," _
                & " " & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID" _
                & " WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".DeptName) ='" & cboDept.Text & "') AND ((" & tblName & ".HeldUp)=0)" _
                & " AND ((" & tblName & ".ComID) =" & CompanyID & ")) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
        MainComm.Execute
    End If

    If optSubDept.Value = True And flexDetails.Rows = 1 Then
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
            & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays,Sum(" & tblName & ".AbsentDay) AS TotalAbsDays," _
            & " Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay,Sum(" & tblName & ".FoodPay) AS TotalFoodPay," _
            & " Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE,Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB," _
            & " Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan,Sum(" & tblName & ".Interest) AS TotalInterest," _
            & " Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay,Sum(" & tblName & ".FracPayD) AS TotalFracPayD," _
            & " Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH,Sum(" & tblName & ".OneH) AS TotalOneH,Sum(" & tblName & ".Fifty) AS TotalFifty," _
            & " Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive,Sum(" & tblName & ".Two) AS TotalTwo,Sum(" & tblName & ".One) AS TotalOne," _
            & " " & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID" _
            & " WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".DeptName) ='" & cboDept.Text & "') AND ((" & tblName & ".SubDeptName) ='" & cboSubDept.Text & "')" _
            & " AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & ")) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
        MainComm.Execute
    ElseIf optSubDept.Value = True And flexDetails.Rows > 1 Then
        If chkRest.Value = 0 Then
            With flexDetails
                For intI = 1 To .Rows - 1
                    .row = intI
                    .col = 0
                    strDept = .Text
                    .col = 1
                    strSubDept = .Text
                    If Len(strSubDept) = 0 Then
                        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,ComName,Address)" _
                            & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays,Sum(" & tblName & ".AbsentDay) AS TotalAbsDays," _
                            & " Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay,Sum(" & tblName & ".FoodPay) AS TotalFoodPay," _
                            & " Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE,Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB," _
                            & " Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan,Sum(" & tblName & ".Interest) AS TotalInterest," _
                            & " Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay,Sum(" & tblName & ".FracPayD) AS TotalFracPayD," _
                            & " Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH,Sum(" & tblName & ".OneH) AS TotalOneH," _
                            & " Sum(" & tblName & ".Fifty) AS TotalFifty,Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive,Sum(" & tblName & ".Two) AS TotalTwo," _
                            & " Sum(" & tblName & ".One) AS TotalOne," & tblName & ".DeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN " & tblName & "" _
                            & " ON Company_Information.ID = " & tblName & ".ComID WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "')" _
                            & " AND ((" & tblName & ".DeptName) ='" & strDept & "') AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & ")) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "' " _
                            & " " & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
                    Else
                        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                            & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays,Sum(" & tblName & ".AbsentDay) AS TotalAbsDays," _
                            & " Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay,Sum(" & tblName & ".FoodPay) AS TotalFoodPay," _
                            & " Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE,Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB," _
                            & " Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan,Sum(" & tblName & ".Interest) AS TotalInterest," _
                            & " Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay,Sum(" & tblName & ".FracPayD) AS TotalFracPayD," _
                            & " Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH,Sum(" & tblName & ".OneH) AS TotalOneH," _
                            & " Sum(" & tblName & ".Fifty) AS TotalFifty,Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive,Sum(" & tblName & ".Two) AS TotalTwo," _
                            & " Sum(" & tblName & ".One) AS TotalOne," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information" _
                            & " INNER JOIN " & tblName & " ON Company_Information.ID = " & tblName & ".ComID WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "')" _
                            & " AND ((" & tblName & ".DeptName) ='" & strDept & "') AND ((" & tblName & ".SubDeptName) ='" & strSubDept & "') AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & "))AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "' " _
                            & " " & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
                    End If
                    MainComm.Execute
                Next intI
            End With
        ElseIf chkRest.Value = 1 Then
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,NotPayD,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                  & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                  & " SELECT Count(" & tblName & ".EmployeeCode) AS TotalEmp," & tblName & ".Month," & tblName & ".Year,Sum(" & tblName & ".RatePay) AS TotalGross,Sum(" & tblName & ".Basic) AS TotalBasic,Sum(" & tblName & ".AttenDays) AS TotalAttenDays,Sum(" & tblName & ".AbsentDay) AS TotalAbsDays," _
                  & " Sum(" & tblName & ".OTHours) AS TotalOTHours,Sum(" & tblName & ".WEarned) AS TotalWEarned,Sum(" & tblName & ".OTPay) AS TotalOTPay,Sum(" & tblName & ".FoodPay) AS TotalFoodPay," _
                  & " Sum(" & tblName & ".FoodAllow) AS TotalFoodAllow,Sum(" & tblName & ".AttenBonus) AS TotalAttenBonus,Sum(" & tblName & ".OthersE) AS TotalOthersE,Sum(" & tblName & ".AbsAmount) AS TotalAbsAmount,Sum(" & tblName & ".PIB_Days) As TotalPIB_Days,Sum(" & tblName & ".PIB) As TotalPIB," _
                  & " Sum(" & tblName & ".TransportD) AS TotalTransportD,Sum(" & tblName & ".Subs) AS TotalSubs,Sum(" & tblName & ".Loan) AS TotalLoan,Sum(" & tblName & ".Interest) AS TotalInterest," _
                  & " Sum(" & tblName & ".OthersD) AS TotalOthersD,Sum(" & tblName & ".NetPayTK) AS TotalNetPayTK,Sum(" & tblName & ".FracPay) AS TotalFracPay,Sum(" & tblName & ".FracPayD) AS TotalFracPayD," _
                  & " Sum(" & tblName & ".TransAllow) AS TransAll,Sum(" & tblName & ".DinnerAllow) AS DinnerAll,Sum(" & tblName & ".TransDay) AS TTransDay,Sum(" & tblName & ".TransPay) AS TTransPay,Sum(" & tblName & ".OneT) AS TotalOneT,Sum(" & tblName & ".FiveH) AS TotalFiveH,Sum(" & tblName & ".OneH) AS TotalOneH," _
                  & " Sum(" & tblName & ".Fifty) AS TotalFifty,Sum(" & tblName & ".Twenty) AS TotalTwenty,Sum(" & tblName & ".Ten) AS TotalTen,Sum(" & tblName & ".Five) AS TotalFive,Sum(" & tblName & ".Two) AS TotalTwo," _
                  & " Sum(" & tblName & ".One) AS TotalOne," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN " & tblName & "" _
                  & " ON Company_Information.ID = " & tblName & ".ComID WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".HeldUp) =0)" _
                  & " AND ((" & tblName & ".ComID) =" & CompanyID & ")) AND " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'" & strBankCash & " GROUP BY " & tblName & ".Month," & tblName & ".Year," & tblName & ".DeptName," & tblName & ".SubDeptName,Company_Information.Name,Company_Information.Address," & tblName & ".ComID;"
            MainComm.Execute
            With flexDetails
                For intI = 1 To .Rows - 1
                    .row = intI
                    .col = 0
                    strDept = .Text
                    .col = 1
                    strSubDept = .Text
                    If Len(strSubDept) = 0 Then
                        MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Summary WHERE DeptName='" & strDept & "';"
                    Else
                        MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Summary WHERE ((DeptName='" & strDept & "') AND (SubDeptName='" & strSubDept & "'));"
                    End If
                    MainComm.Execute
                Next intI
            End With
        End If
    End If
    StrRecord = "SELECT DeptName FROM Temp_Summary;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.RecordCount = 0 Then
        MsgBox "Record not found", vbCritical, cnstMsgInfo
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    r.Close
    If chkDept.Value = 1 Then
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET SubDeptName='ALL SUB DEPT';"
        MainComm.Execute
    End If
    StrRecord = "SELECT DeptName,SubDeptName,([WEarned]+[OTAmount]+[FoodAmount]+[FoodAllow]+[AttenBonus]+[TransAllow]+[DinnerAllow]+[OthersE]+[FracPay]+[AbsAmount]+[PIB]+[TransPay]) AS TotalAmount FROM Temp_Summary;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        Do Until r.EOF
            MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET TotalAmount = " & r![TotalAmount] & " WHERE DeptName='" & r![DeptName] & "'AND SubDeptName='" & r![SubDeptName] & "';"
            MainComm.Execute
            r.MoveNext
        Loop
    End If
    r.Close
    
    MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = 0;"
    MainComm.Execute
    
    If Optdept.Value = True Then
        StrRecord = "SELECT DeptName FROM Temp_Summary;"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                ''MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = (SELECT Count(" & tblName & ".NetPayTK) AS EmpWithoutStamp From " & tblName & " " _
                ''    & " WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & ") And ((" & tblName & ".WSTATUS ) ='" & FrameCategory.Tag & "') AND ((" & tblName & ".NETPAYTK) <=  1000) " _
                ''    & " AND ((" & tblName & ".DeptName) = '" & r![DeptName] & "'))) " _
                ''    & " WHERE DeptName='" & r![DeptName] & "';"
                MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = (SELECT Count(" & tblName & ".NetPayTK) AS EmpWithoutStamp From " & tblName & "" _
                    & " WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND (" & tblName & ".ComID) =" & CompanyID & ") And ((" & tblName & ".WSTATUS ) ='" & FrameCategory.Tag & "') AND ((" & tblName & ".Stamp) = 0)" _
                    & " AND ((" & tblName & ".DeptName) = '" & r![DeptName] & "'))" & strBankCash & ") " _
                    & " WHERE DeptName='" & r![DeptName] & "';"
                MainComm.Execute
                r.MoveNext
            Loop
        End If
        r.Close
    Else
        StrRecord = "SELECT DeptName,SubDeptName FROM Temp_Summary;"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = (SELECT Count(" & tblName & ".NetPayTK) AS EmpWithoutStamp From " & tblName & " " _
                    & " WHERE (((" & tblName & ".Month) ='" & cboMonthName.Text & "') AND ((" & tblName & ".Year) ='" & PeriodYear & "') AND ((" & tblName & ".HeldUp) =0) AND ((" & tblName & ".ComID) =" & CompanyID & ") And " & tblName & ".WSTATUS='" & FrameCategory.Tag & "'AND ((" & tblName & ".Stamp) = 0)" _
                    & " AND ((" & tblName & ".DeptName) = '" & r![DeptName] & "') And ((" & tblName & ".SubDeptName)='" & r![SubDeptName] & "'))" & strBankCash & ") " _
                    & " WHERE DeptName='" & r![DeptName] & "' AND SubDeptName='" & r![SubDeptName] & "';"
                MainComm.Execute
                r.MoveNext
            Loop
        End If
        r.Close
    End If
    
    If cboMonthName.Text = "AUGUST" And PeriodYear = 2005 Then
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET FracPayD = ([FracPayD] + 76) WHERE DeptName='STAFF-A' AND SubDeptName='WET PROCESSING';"
        MainComm.Execute
    End If
    
    If chkPrev.Value = vbChecked Then
      MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET SubSubDept = ' (ML EMPLOYEE)';"
    Else
      MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET SubSubDept = ' ';"
    End If
    MainComm.Execute

goCommit:

MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Err.Number = 0
        Exit Sub
    End If
    '-----Dim Appl As New CRAXDRT.Application
    '-----Dim Report As New CRAXDRT.Report
    
    If chkDept.Value = 1 Then

    Set Report = Appl.OpenReport(ReportPath & "salary summary dept.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "salary summary dept.rpt"
    ElseIf chkDesig.Value = 1 Then
    Set Report = Appl.OpenReport(ReportPath & "salary summary desig.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "salary summary desig.rpt"
    Else
    Set Report = Appl.OpenReport(ReportPath & "salary summary.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "salary summary.rpt"
    End If
    
    If optWorkForce.Value = True Then
      If chkCashPay.Value = 1 Then
        Report.FormulaFields(5).Text = "'Worker Wages (Cash)'"
      ElseIf chkBankPay.Value = 1 Then
        Report.FormulaFields(5).Text = "'Worker Wages (Bank)'"
      ElseIf chkPrev.Value = 1 Then
        Report.FormulaFields(5).Text = "'ML Worker Wages'"
      Else
        Report.FormulaFields(5).Text = "'Worker Wages'"
'       CrystalReport1report.FormulaFields(0).Text = "Param = 'Worker'"
      End If
    Else
      If chkCashPay.Value = 1 Then
        Report.FormulaFields(5).Text = "'Staff Salary (Cash)'"
      ElseIf chkBankPay.Value = 1 Then
        Report.FormulaFields(5).Text = "'Staff Salary (Bank)'"
      ElseIf chkPrev.Value = 1 Then
        Report.FormulaFields(5).Text = "'ML Staff Salary'"
      Else
        Report.FormulaFields(5).Text = "'Staff Salary'"
      End If
'       CrystalReport1report.FormulaFields(0).Text = "Param = 'Staff'"
    End If
    
    Report.DiscardSavedData
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1
'    CrystalReport1.PrintReport
    
    Screen.MousePointer = vbDefault
    Set r = Nothing
    'MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Salary_And_Wages;"
    'MainComm.Execute
    'MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Summary;"
    'MainComm.Execute
End Sub
Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
  With flexDetails
    If .RowSel = 0 Then
      Beep
    Else
      If MsgBox("Are you sure to Remove", vbYesNo + vbQuestion + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
      If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
    End If
  End With
End If
End Sub
Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Report - Salary Summary"
  If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"

  DeptName cboDept
  prcMonthName cboMonthName
  cboMonthName.Text = UCase(Format(DateAdd("M", -1, Date), "MMMM"))
  With flexDetails
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColWidth(0) = 2500
    .ColWidth(1) = 2500
    .row = 0
    .col = 0
    .Text = "Department"
    .col = 1
    .Text = "Sub Department"
  End With
  ' prcConReport CrystalReport1
End Sub
Private Sub Form_Unload(Cancel As Integer)
  boJRMgt = False
  MDIfrmMain.lblCaption.Caption = ""
  Set frmRptSummary = Nothing
End Sub

Private Sub optStaff_Click()
    FrameCategory.Tag = "F"
End Sub

Private Sub optWorkForce_Click()
    FrameCategory.Tag = "W"
End Sub
