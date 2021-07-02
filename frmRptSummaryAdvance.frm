VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRptSummaryAdvance 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7865
   ClientLeft      =   3276
   ClientTop       =   3003
   ClientWidth     =   5876
   ControlBox      =   0   'False
   Icon            =   "frmRptSummaryAdvance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7865
   ScaleWidth      =   5876
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2834
      TabIndex        =   29
      Top             =   6786
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
      Left            =   312
      TabIndex        =   28
      Top             =   6786
      Width           =   2028
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear &List"
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
      Left            =   2756
      TabIndex        =   24
      Top             =   7345
      Width           =   1095
   End
   Begin VB.Frame FrameCategory 
      Caption         =   "Show Report"
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
      Height          =   637
      Left            =   240
      TabIndex        =   20
      Tag             =   "W"
      Top             =   1196
      Width           =   5355
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
         Left            =   210
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.OptionButton optJRMgt 
         Caption         =   "JR Mgt."
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
         Left            =   3000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Paid"
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
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   3750
      Width           =   1125
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
      TabIndex        =   17
      Top             =   2910
      Width           =   675
   End
   Begin VB.CheckBox chkAllSDept 
      Caption         =   "Add all Sub Dept. Under Selected Dept."
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
      Height          =   225
      Left            =   2040
      TabIndex        =   16
      Top             =   3780
      Width           =   3645
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
      TabIndex        =   3
      Top             =   3330
      Width           =   675
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
      TabIndex        =   2
      Top             =   3330
      Width           =   2775
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
      Height          =   525
      Left            =   240
      TabIndex        =   8
      Top             =   1703
      Width           =   5355
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
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
         Left            =   3000
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   2205
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
      TabIndex        =   1
      Top             =   2910
      Width           =   2775
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
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Month"
      Top             =   2490
      Width           =   3525
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
      Left            =   4290
      TabIndex        =   4
      Top             =   7345
      Width           =   1095
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
      Left            =   312
      TabIndex        =   5
      Top             =   7345
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   2730
      Left            =   240
      TabIndex        =   12
      Top             =   4110
      Width           =   5355
      _ExtentX        =   9441
      _ExtentY        =   4816
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
      Height          =   741
      Left            =   0
      TabIndex        =   27
      Top             =   7163
      Width           =   5889
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Advance Salary Summary"
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
      Left            =   165
      TabIndex        =   26
      Top             =   240
      Width           =   3795
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   5895
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
      TabIndex        =   15
      Top             =   3330
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
      TabIndex        =   14
      Top             =   2910
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
      Index           =   0
      Left            =   1800
      TabIndex        =   13
      Top             =   2490
      Width           =   255
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
      TabIndex        =   11
      Top             =   3330
      Width           =   1455
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
      TabIndex        =   7
      Top             =   2910
      Width           =   1095
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
      TabIndex        =   6
      Top             =   2490
      Width           =   1335
   End
End
Attribute VB_Name = "frmRptSummaryAdvance"
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
Dim strDept As String, strSubDept As String, strBankCash As String
Dim strLastDept As String, strLastSubDept As String

    If Len(cboMonthName) = 0 Then
        MsgBox "Select Name of Month", vbInformation, cnstMsgInfo
        cboMonthName.SetFocus
        Exit Sub
    End If
    If OptDept.Value = True And Len(cboDept) = 0 Then
        MsgBox "Select Department", vbInformation, cnstMsgInfo
        cboDept.SetFocus
        Exit Sub
    End If
    If optSubDept.Value = True And flexDetails.Rows = 1 Then
        MsgBox "Add Sub Department", vbInformation, cnstMsgInfo
        cboSubDept.SetFocus
        Exit Sub
    End If
    
    Set r = New ADODB.Recordset
    Set Rs_rpt = New ADODB.Recordset
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    With flexDetails
        For intI = 1 To .Rows - 1
            .row = intI
            .col = 0
            If strLastDept = .Text Then
            Else
                strDept = strDept & "'" & Trim(.Text) & "',"
            End If
            
            strLastDept = .Text
            .col = 1
            If strLastSubDept = .Text Then
            Else
                strSubDept = strSubDept & "'" & Trim(.Text) & "',"
            End If
            strLastSubDept = .Text
        Next intI
        
        If strDept = "" Then
        Else
            strDept = Mid(strDept, 1, (Len(strDept) - 1))
        End If
        If strSubDept = "" Then
        Else
            strSubDept = Mid(strSubDept, 1, (Len(strSubDept) - 1))
        End If
    End With
    
    If chkCashPay.Value = 1 Then
      strBankCash = " AND EmpCode NOT IN (SELECT DISTINCT Empcode FROM Emp_Account WHERE (AccNo IS NOT NULL AND AccNo <> '' AND AccNo <> '=N/A=') AND SUBSTRING(HeldUp,1,1) = 'N')"
    ElseIf chkBankPay.Value = 1 Then
      strBankCash = " AND EmpCode IN (SELECT DISTINCT Empcode FROM Emp_Account WHERE (AccNo IS NOT NULL AND AccNo <> '' AND AccNo <> '=N/A=') AND SUBSTRING(HeldUp,1,1) = 'N')"
    Else
      strBankCash = ""
    End If
    
    Set Report = Appl.OpenReport(ReportPath & "rptAdvanceSalarySum.rpt")
       
        Report.FormulaFields(1).Text = ""
        If chkStatus.Value = 0 Then
            If Len(strSubDept) > 0 Then
                Mysql = " Select * From VIEW_EMP_ADVANCE_SALARY where comid = " & CompanyID & " and Department IN (" & strDept & ") and Section IN (" & strSubDept & ") and Status = 0 And Wstatus = '" & FrameCategory.Tag & "' And Year = '" & PeriodYear & "' and Month = '" & cboMonthName.Text & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            ElseIf Len(strDept) > 0 Then
                Mysql = " Select * from VIEW_EMP_ADVANCE_SALARY where comid = " & CompanyID & " and Department IN (" & strDept & ") and Status = 0 And Wstatus = '" & FrameCategory.Tag & "' And Year = '" & PeriodYear & "' and Month = '" & cboMonthName.Text & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            Else
                Mysql = " Select * from VIEW_EMP_ADVANCE_SALARY where comid = " & CompanyID & " and  Status= 0 And Wstatus = '" & FrameCategory.Tag & "' And Year = '" & PeriodYear & "' and Month = '" & cboMonthName.Text & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            End If
            ''Report.FormulaFields(1).Text = "'UNPAID'"
        Else
            If Len(strSubDept) > 0 Then
                Mysql = " Select * from VIEW_EMP_ADVANCE_SALARY where comid = " & CompanyID & " and Department = '" & Trim(cboDept.Text) & "' And Wstatus = '" & FrameCategory.Tag & "' and Section = '" & Trim(cboSubDept.Text) & "' and Status= 1 And Year = '" & PeriodYear & "' and Month = '" & cboMonthName.Text & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            ElseIf Len(strDept) > 0 Then
                Mysql = " Select * from VIEW_EMP_ADVANCE_SALARY where comid = " & CompanyID & " and Department = '" & Trim(cboDept.Text) & "' And Wstatus = '" & FrameCategory.Tag & "' and Status=1 And Year = '" & PeriodYear & "' and Month = '" & cboMonthName.Text & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            Else
                Mysql = " Select * from VIEW_EMP_ADVANCE_SALARY  where Status = 1 And Wstatus = '" & FrameCategory.Tag & "'" & strBankCash & " ORDER BY DEPARTMENT, SECTION"
            End If
            ''Report.FormulaFields(1).Text = "'PAID'"
        End If
        
        If chkCashPay.Value = 1 Then
          Report.FormulaFields(1).Text = "'(Only Cash)'"
        ElseIf chkBankPay.Value = 1 Then
          Report.FormulaFields(1).Text = "'(Only Bank)'"
        Else
          Report.FormulaFields(1).Text = "'(Cash & Bank)'"
        End If

        Rs_rpt.Open Mysql, MainConn, 3, 3
        Report.DiscardSavedData
        
        Report.Database.SetDataSource Rs_rpt
     
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
    
    Screen.MousePointer = vbDefault
    Set r = Nothing
End Sub

Private Sub cmdClearList_Click()
    flexDetails.Rows = 1
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
    MDIfrmMain.lblCaption.Caption = "Report - Advance Salary Summary"
    If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"
    optJRMgt.Enabled = boJRMgt
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
'    ' prcConReport CrystalReport1
    Set r = New ADODB.Recordset
    StrRecord = "SELECT Worker,JRmgt,SRmgt,Staff FROM Sys_User_Name WHERE UserId='" & bytUserID & "' ;"
    r.Open StrRecord, MainConn, adOpenStatic
    
    If r![JRmgt] = 1 Then
    optJRMgt.Enabled = True
    Else
    optJRMgt.Enabled = False
    End If
    
    If r![Staff] = 1 Then
    optStaff.Enabled = True
    Else
    optStaff.Enabled = False
    End If
   r.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
    boJRMgt = False
    MDIfrmMain.lblCaption.Caption = ""
    Set frmRptSummary = Nothing
End Sub

Private Sub optJrMgt_Click()
    FrameCategory.Tag = "J"
End Sub

Private Sub optStaff_Click()
    FrameCategory.Tag = "F"
End Sub

Private Sub optWorkForce_Click()
    FrameCategory.Tag = "W"
End Sub
