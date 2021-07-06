VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRptSummaryC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8190
   ClientLeft      =   3270
   ClientTop       =   3000
   ClientWidth     =   5745
   ControlBox      =   0   'False
   Icon            =   "frmRptSummaryC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCategory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Tag             =   "W"
      Top             =   1050
      Width           =   5355
      Begin VB.OptionButton optWorkForce 
         Caption         =   "Work Force"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   170
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optStaff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   170
         Width           =   825
      End
   End
   Begin VB.CheckBox chkDesig 
      Caption         =   "Show Designation wise Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   6810
      Width           =   2715
   End
   Begin VB.CommandButton cmdAddDept 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   20
      Top             =   2640
      Width           =   675
   End
   Begin VB.CheckBox chkDept 
      Caption         =   "Show Dept. Summary (List added)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6510
      Width           =   2805
   End
   Begin VB.CheckBox chkAllSDept 
      Caption         =   "Add all Sub Dept. Under Selected Dept."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   2070
      TabIndex        =   18
      Top             =   3480
      Width           =   3195
   End
   Begin VB.CheckBox chkRest 
      Caption         =   "Show All Dept. && Sub Dept. without List added."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6210
      Width           =   3735
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   3
      Top             =   3060
      Width           =   675
   End
   Begin VB.ComboBox cboSubDept 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   2
      Top             =   3060
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   5355
      Begin VB.OptionButton optSubDept 
         Caption         =   "Sub Department Wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         Width           =   2205
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton OptDept 
         Caption         =   "Department Wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.ComboBox cboDept 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.ComboBox cboMonthName 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Month"
      Top             =   2220
      Width           =   3525
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Top             =   7560
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   5
      Top             =   7560
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   2370
      Left            =   240
      TabIndex        =   13
      Top             =   3780
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4180
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   285
      BackColor       =   -2147483626
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Salary Summary (C)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   7200
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   16
      Top             =   3060
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   15
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   14
      Top             =   2220
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label lblDept 
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label29 
      Caption         =   "Name of Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2220
      Width           =   1335
   End
End
Attribute VB_Name = "frmRptSummaryC"
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
Dim strDept As String, strSubDept As String
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
              MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Salary_And_Wages(One,Year,Month,EmployeeCode,Designation,RatePay,AttenDays,AbsentDay,OTHours,WEarned,OTPay,FoodPay,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,ComName,Address) SELECT Salary_And_WagesN.ComID,Salary_And_WagesN.Year,Salary_And_WagesN.Month,Salary_And_WagesN.EmployeeCode,Employee_Information.Designation,Salary_And_WagesN.RatePay," _
                & " Salary_And_WagesN.AttenDays,Salary_And_WagesN.AbsentDay,Salary_And_WagesN.OTHours,Salary_And_WagesN.WEarned,Salary_And_WagesN.OTPay,Salary_And_WagesN.FoodPay,Salary_And_WagesN.FoodAllow,Salary_And_WagesN.AttenBonus,Salary_And_WagesN.OthersE," _
                & " Salary_And_WagesN.AbsAmount,Salary_And_WagesN.PIB_Days,Salary_And_WagesN.PIB,Salary_And_WagesN.TransportD,Salary_And_WagesN.OthersD,Salary_And_WagesN.NetPayTK,Salary_And_WagesN.FracPay,Salary_And_WagesN.FracPayD,Salary_And_WagesN.TransAllow,Salary_And_WagesN.DinnerAllow,Salary_And_WagesN.TransDay,Salary_And_WagesN.TransPay,Salary_And_WagesN.DeptName," _
                & " Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Employee_Information INNER JOIN (Company_Information INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID)" _
                & " ON (Employee_Information.ComID = Salary_And_WagesN.ComID) AND (Employee_Information.EmployeeCode = Salary_And_WagesN.EmployeeCode) WHERE (((Salary_And_WagesN.Year) = '" & PeriodYear & "') AND ((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "')" _
                & " AND ((Salary_And_WagesN.HeldUp) = 0) AND ((Salary_And_WagesN.ComID) = " & CompanyID & ") AND ((Salary_And_WagesN.WorkDays) = 26)) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' ORDER BY Salary_And_WagesN.EmployeeCode;"
        ElseIf Optdept.Value = True Then
              MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Salary_And_Wages(One,Year,Month,EmployeeCode,Designation,RatePay,AttenDays,AbsentDay,OTHours,WEarned,OTPay,FoodPay,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,ComName,Address) SELECT Salary_And_WagesN.ComID,Salary_And_WagesN.Year,Salary_And_WagesN.Month,Salary_And_WagesN.EmployeeCode,Employee_Information.Designation,Salary_And_WagesN.RatePay," _
                & " Salary_And_WagesN.AttenDays,Salary_And_WagesN.AbsentDay,Salary_And_WagesN.OTHours,Salary_And_WagesN.WEarned,Salary_And_WagesN.OTPay,Salary_And_WagesN.FoodPay,Salary_And_WagesN.FoodAllow,Salary_And_WagesN.AttenBonus,Salary_And_WagesN.OthersE," _
                & " Salary_And_WagesN.AbsAmount,Salary_And_WagesN.PIB_Days,Salary_And_WagesN.PIB,Salary_And_WagesN.TransportD,Salary_And_WagesN.OthersD,Salary_And_WagesN.NetPayTK,Salary_And_WagesN.FracPay,Salary_And_WagesN.FracPayD,Salary_And_WagesN.TransAllow,Salary_And_WagesN.DinnerAllow,Salary_And_WagesN.TransDay,Salary_And_WagesN.TransPay,Salary_And_WagesN.DeptName," _
                & " Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Employee_Information INNER JOIN (Company_Information INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID)" _
                & " ON (Employee_Information.ComID = Salary_And_WagesN.ComID) AND (Employee_Information.EmployeeCode = Salary_And_WagesN.EmployeeCode) WHERE (((Salary_And_WagesN.Year) = '" & PeriodYear & "') AND ((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "')" _
                & " AND ((Salary_And_WagesN.DeptName) ='" & cboDept.Text & "') AND ((Salary_And_WagesN.HeldUp) = 0) AND ((Salary_And_WagesN.ComID) = " & CompanyID & ") AND ((Salary_And_WagesN.WorkDays) = 26)) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' ORDER BY Salary_And_WagesN.EmployeeCode;"
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
        
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus,OthersE,AbsAmount,PIB_Days,PIB,TransportD,OthersD,NetPayTK,FracPay," _
                & " FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,DeptName,SubDeptName,SubSubDept,ComName,Address)SELECT Count(EmployeeCode) AS TotalEmp,[Month],[Year],Sum(RatePay) AS TotalGross,Sum(AttenDays) AS TotalAttenDays,Sum(AbsentDay) AS TotalAbsDays," _
                & " Sum(OTHours) AS TotalOTHours,Sum(WEarned) AS TotalWEarned,Sum(OTPay) AS TotalOTPay,Sum(FoodPay) AS TotalFoodPay,Sum(FoodAllow) AS TotalFoodAllow,Sum(AttenBonus) AS TotalAttenBonus,Sum(OthersE) AS TotalOthersE,Sum(AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
                & " Sum(TransportD) AS TotalTransportD,Sum(OthersD) AS TotalOthersD,Sum(NetPayTK) AS TotalNetPayTK,Sum(FracPay) AS TotalFracPay,Sum(FracPayD) AS TotalFracPayD,Sum(TransAllow) AS TransAll,Sum(DinnerAllow) AS DinnerAll,Sum(TransDay) AS TTransDay,Sum(TransPay) AS TTransPay,DeptName,SubDeptName," _
                & " Designation,ComName,Address FROM Temp_Salary_And_Wages GROUP BY [Month],[Year],DeptName,SubDeptName,Designation,ComName,Address;"
        MainComm.Execute
        
        StrRecord = "SELECT DeptName,SubDeptName,SubSubDept,([WEarned]+[OTAmount]+[FoodAmount]+[FoodAllow]+[AttenBonus]+[TransAllow]+[DinnerAllow]+[OthersE]+[FracPay]+[TransPay]) AS TotalAmount FROM Temp_Summary;"
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
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,ComName,Address)" _
                & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays," _
                & " Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays,Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay," _
                & " Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay,Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE," _
                & " Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB,Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan," _
                & " Sum(Salary_And_WagesN.Interest) AS TotalInterest,Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay," _
                & " Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD,Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH," _
                & " Sum(Salary_And_WagesN.OneH) AS TotalOneH,Sum(Salary_And_WagesN.Fifty) AS TotalFifty,Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive," _
                & " Sum(Salary_And_WagesN.Two) AS TotalTwo,Sum(Salary_And_WagesN.One) AS TotalOne,Salary_And_WagesN.DeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN Salary_And_WagesN" _
                & " ON Company_Information.ID = Salary_And_WagesN.ComID WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.HeldUp) =0)" _
                & " AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
'            MainComm.Execute
'            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(EmpWithoutStamp) SELECT Count(Salary_And_Wages.NetPayTK) AS EmpWithoutStamp FROM Salary_And_Wages " _
'                & " Inner Join Temp_Summary On Salary_And_Wages.ComID = Temp_Summary.ComID And Salary_And_Wages.DeptName = Temp_Summary.DeptName " _
'                & " WHERE (((Salary_And_Wages.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_Wages.Year) ='" & PeriodYear & "') AND ((Salary_And_Wages.HeldUp) =0)" _
'                & " AND ((Salary_And_Wages.ComID) =" & CompanyID & "))And Salary_And_Wages.NetPayTK < 500 " _
'                & " GROUP BY Salary_And_Wages.Month,Salary_And_Wages.Year,Salary_And_Wages.DeptName,Salary_And_Wages.ComID;"
'            MainComm.Execute
        Else
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays," _
                & " Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays,Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay," _
                & " Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay,Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE," _
                & " Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB,Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan," _
                & " Sum(Salary_And_WagesN.Interest) AS TotalInterest,Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay," _
                & " Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD,Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH," _
                & " Sum(Salary_And_WagesN.OneH) AS TotalOneH,Sum(Salary_And_WagesN.Fifty) AS TotalFifty,Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive," _
                & " Sum(Salary_And_WagesN.Two) AS TotalTwo,Sum(Salary_And_WagesN.One) AS TotalOne,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information" _
                & " INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "')" _
                & " AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName," _
                & " Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
        End If
        MainComm.Execute
    ElseIf Optdept.Value = True Then
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays,Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays," _
                & " Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay,Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay," _
                & " Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE,Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
                & " Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan,Sum(Salary_And_WagesN.Interest) AS TotalInterest," _
                & " Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay,Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD," _
                & " Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH,Sum(Salary_And_WagesN.OneH) AS TotalOneH,Sum(Salary_And_WagesN.Fifty) AS TotalFifty," _
                & " Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive,Sum(Salary_And_WagesN.Two) AS TotalTwo,Sum(Salary_And_WagesN.One) AS TotalOne," _
                & " Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID" _
                & " WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.DeptName) ='" & cboDept.Text & "') AND ((Salary_And_WagesN.HeldUp)=0)" _
                & " AND ((Salary_And_WagesN.ComID) =" & CompanyID & "))  AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
        MainComm.Execute
    End If

    If optSubDept.Value = True And flexDetails.Rows = 1 Then
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
            & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays,Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays," _
            & " Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay,Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay," _
            & " Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE,Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
            & " Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan,Sum(Salary_And_WagesN.Interest) AS TotalInterest," _
            & " Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay,Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD," _
            & " Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH,Sum(Salary_And_WagesN.OneH) AS TotalOneH,Sum(Salary_And_WagesN.Fifty) AS TotalFifty," _
            & " Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive,Sum(Salary_And_WagesN.Two) AS TotalTwo,Sum(Salary_And_WagesN.One) AS TotalOne," _
            & " Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID" _
            & " WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.DeptName) ='" & cboDept.Text & "') AND ((Salary_And_WagesN.SubDeptName) ='" & cboSubDept.Text & "')" _
            & " AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
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
                        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,ComName,Address)" _
                            & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays,Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays," _
                            & " Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay,Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay," _
                            & " Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE,Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
                            & " Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan,Sum(Salary_And_WagesN.Interest) AS TotalInterest," _
                            & " Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay,Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD," _
                            & " Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH,Sum(Salary_And_WagesN.OneH) AS TotalOneH," _
                            & " Sum(Salary_And_WagesN.Fifty) AS TotalFifty,Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive,Sum(Salary_And_WagesN.Two) AS TotalTwo," _
                            & " Sum(Salary_And_WagesN.One) AS TotalOne,Salary_And_WagesN.DeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN Salary_And_WagesN" _
                            & " ON Company_Information.ID = Salary_And_WagesN.ComID WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "')" _
                            & " AND ((Salary_And_WagesN.DeptName) ='" & strDept & "') AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "'" _
                            & " GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
                    Else
                        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                            & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                            & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays,Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays," _
                            & " Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay,Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay," _
                            & " Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE,Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
                            & " Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan,Sum(Salary_And_WagesN.Interest) AS TotalInterest," _
                            & " Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay,Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD," _
                            & " Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH,Sum(Salary_And_WagesN.OneH) AS TotalOneH," _
                            & " Sum(Salary_And_WagesN.Fifty) AS TotalFifty,Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive,Sum(Salary_And_WagesN.Two) AS TotalTwo," _
                            & " Sum(Salary_And_WagesN.One) AS TotalOne,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information" _
                            & " INNER JOIN Salary_And_WagesN ON Company_Information.ID = Salary_And_WagesN.ComID WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "')" _
                            & " AND ((Salary_And_WagesN.DeptName) ='" & strDept & "') AND ((Salary_And_WagesN.SubDeptName) ='" & strSubDept & "') AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "'" _
                            & " GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
                    End If
                    MainComm.Execute
                Next intI
            End With
        ElseIf chkRest.Value = 1 Then
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Temp_Summary(TotalEmp,Month,Year,Cont,TotalWorkDay,TotalAbsDays,TotalOTHours,WEarned,OTAmount,FoodAmount,FoodAllow,AttenBonus," _
                  & " OthersE,AbsAmount,PIB_Days,PIB,TransportD,Subs,Loan,Interest,OthersD,NetPayTK,FracPay,FracPayD,TransAllow,DinnerAllow,TransDay,TransPay,OneT,FiveH,OneH,Fifty,Twenty,Ten,Five,Two,One,DeptName,SubDeptName,ComName,Address)" _
                  & " SELECT Count(Salary_And_WagesN.EmployeeCode) AS TotalEmp,Salary_And_WagesN.Month,Salary_And_WagesN.Year,Sum(Salary_And_WagesN.RatePay) AS TotalGross,Sum(Salary_And_WagesN.AttenDays) AS TotalAttenDays,Sum(Salary_And_WagesN.AbsentDay) AS TotalAbsDays," _
                  & " Sum(Salary_And_WagesN.OTHours) AS TotalOTHours,Sum(Salary_And_WagesN.WEarned) AS TotalWEarned,Sum(Salary_And_WagesN.OTPay) AS TotalOTPay,Sum(Salary_And_WagesN.FoodPay) AS TotalFoodPay," _
                  & " Sum(Salary_And_WagesN.FoodAllow) AS TotalFoodAllow,Sum(Salary_And_WagesN.AttenBonus) AS TotalAttenBonus,Sum(Salary_And_WagesN.OthersE) AS TotalOthersE,Sum(Salary_And_WagesN.AbsAmount) AS TotalAbsAmount,Sum(PIB_Days) As TotalPIB_Days,Sum(PIB) As TotalPIB," _
                  & " Sum(Salary_And_WagesN.TransportD) AS TotalTransportD,Sum(Salary_And_WagesN.Subs) AS TotalSubs,Sum(Salary_And_WagesN.Loan) AS TotalLoan,Sum(Salary_And_WagesN.Interest) AS TotalInterest," _
                  & " Sum(Salary_And_WagesN.OthersD) AS TotalOthersD,Sum(Salary_And_WagesN.NetPayTK) AS TotalNetPayTK,Sum(Salary_And_WagesN.FracPay) AS TotalFracPay,Sum(Salary_And_WagesN.FracPayD) AS TotalFracPayD," _
                  & " Sum(Salary_And_WagesN.TransAllow) AS TransAll,Sum(Salary_And_WagesN.DinnerAllow) AS DinnerAll,Sum(Salary_And_WagesN.TransDay) AS TTransDay,Sum(Salary_And_WagesN.TransPay) AS TTransPay,Sum(Salary_And_WagesN.OneT) AS TotalOneT,Sum(Salary_And_WagesN.FiveH) AS TotalFiveH,Sum(Salary_And_WagesN.OneH) AS TotalOneH," _
                  & " Sum(Salary_And_WagesN.Fifty) AS TotalFifty,Sum(Salary_And_WagesN.Twenty) AS TotalTwenty,Sum(Salary_And_WagesN.Ten) AS TotalTen,Sum(Salary_And_WagesN.Five) AS TotalFive,Sum(Salary_And_WagesN.Two) AS TotalTwo," _
                  & " Sum(Salary_And_WagesN.One) AS TotalOne,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address FROM Company_Information INNER JOIN Salary_And_WagesN" _
                  & " ON Company_Information.ID = Salary_And_WagesN.ComID WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.HeldUp) =0)" _
                  & " AND ((Salary_And_WagesN.ComID) =" & CompanyID & ")) AND Salary_And_WagesN.WSTATUS='" & FrameCategory.Tag & "' GROUP BY Salary_And_WagesN.Month,Salary_And_WagesN.Year,Salary_And_WagesN.DeptName,Salary_And_WagesN.SubDeptName,Company_Information.Name,Company_Information.Address,Salary_And_WagesN.ComID;"
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
    StrRecord = "SELECT DeptName,SubDeptName,ISNULL(([WEarned]+[OTAmount]+[FoodAmount]+[FoodAllow]+[AttenBonus]+[TransAllow]+[DinnerAllow]+[OthersE]+[FracPay]+[TransPay]),0) AS TotalAmount FROM Temp_Summary;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        Do Until r.EOF
            MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET TotalAmount= " & r![TotalAmount] & "  WHERE  DeptName='" & r![DeptName] & "'AND SubDeptName='" & r![SubDeptName] & "';"
            MainComm.Execute
            r.MoveNext
        Loop
    End If
    r.Close
    If Optdept.Value = True Then
        StrRecord = "SELECT DeptName FROM Temp_Summary;"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = (SELECT Count(Salary_And_WagesN.NetPayTK) AS EmpWithoutStamp From Salary_And_WagesN " _
                    & " WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ") And ((Salary_And_WagesN.NetPayTK +  Salary_And_WagesN.FracPayD + Salary_And_WagesN.STAMP) < 500) " _
                    & " AND ((Salary_And_WagesN.DeptName) = '" & r![DeptName] & "'))) " _
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
                MainComm.CommandText = "SET NOCOUNT ON UPDATE Temp_Summary SET EmpWithoutStamp = (SELECT Count(Salary_And_WagesN.NetPayTK) AS EmpWithoutStamp From Salary_And_WagesN " _
                    & " WHERE (((Salary_And_WagesN.Month) ='" & cboMonthName.Text & "') AND ((Salary_And_WagesN.Year) ='" & PeriodYear & "') AND ((Salary_And_WagesN.HeldUp) =0) AND ((Salary_And_WagesN.ComID) =" & CompanyID & ") And ((Salary_And_WagesN.NetPayTK +  Salary_And_WagesN.FracPayD + Salary_And_WagesN.STAMP) < 500) " _
                    & " AND ((Salary_And_WagesN.DeptName) = '" & r![DeptName] & "') And ((Salary_And_WagesN.SubDeptName)='" & r![SubDeptName] & "'))) " _
                    & " WHERE DeptName='" & r![DeptName] & "'AND SubDeptName='" & r![SubDeptName] & "';"
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
    Set Report = Appl.OpenReport(ReportPath & "orgin\salary summary dept.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "orgin\salary summary dept.rpt"
    ElseIf chkDesig.Value = 1 Then
    Set Report = Appl.OpenReport(ReportPath & "orgin\salary summary desig.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "orgin\salary summary desig.rpt"
    Else
    Set Report = Appl.OpenReport(ReportPath & "orgin\salary summary.rpt")
'        CrystalReport1.ReportFileName = ReportPath + "orgin\salary summary.rpt"
    End If
    If optWorkForce.Value = True Then
    Report.FormulaFields(5).Text = "'Worker'"
'       CrystalReport1report.FormulaFields(0).Text = "Param = 'Worker'"
    Else
    Report.FormulaFields(5).Text = "'Staff'"
'       CrystalReport1report.FormulaFields(0).Text = "Param = 'Staff'"
    End If
    
    Report.DiscardSavedData
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1
    
'    CrystalReport1.PrintReport
'
'    Screen.MousePointer = vbDefault
    Set r = Nothing
    'MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Temp_Salary_And_WagesN;"
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
'  ' prcConReport CrystalReport1
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
