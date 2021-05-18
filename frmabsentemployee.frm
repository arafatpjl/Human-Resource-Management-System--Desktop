VERSION 5.00
Begin VB.Form frmrabsentemployee 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6448
   ClientLeft      =   3211
   ClientTop       =   2366
   ClientWidth     =   5746
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6448
   ScaleWidth      =   5746
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   4320
      Width           =   5460
      Begin VB.TextBox txtTo 
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
         Height          =   330
         Left            =   2625
         TabIndex        =   8
         ToolTipText     =   "Type Date of Joining"
         Top             =   390
         Width           =   2085
      End
      Begin VB.TextBox txtFrom 
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
         Height          =   330
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Type Date of Joining"
         Top             =   390
         Width           =   2085
      End
      Begin VB.Label Label12 
         Caption         =   "Day(s)"
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbldays 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "From Date  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label Label9 
         Caption         =   "To Date  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2625
         TabIndex        =   31
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absent Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   5460
      Begin VB.OptionButton Optother 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Optfp 
         Caption         =   "Family Problem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton OptSick 
         Caption         =   "Sickness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame FrmDept 
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   5460
      Begin VB.ComboBox cboempcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2895
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Emp Code      :"
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
         Left            =   315
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   1440
         TabIndex        =   24
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblSubDept 
         Caption         =   "Sub Dept."
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
         Left            =   315
         TabIndex        =   23
         Top             =   720
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   315
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Left            =   1455
         TabIndex        =   21
         Top             =   360
         Width           =   255
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
      Height          =   570
      Left            =   120
      TabIndex        =   19
      Tag             =   "W"
      Top             =   1080
      Width           =   5445
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
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optJrMgt 
         Caption         =   "JR. Mgt."
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
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optSrMgt 
         Caption         =   "SR. Mgt."
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
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   1035
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
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "&z"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   15
      Top             =   8610
      Width           =   1155
   End
   Begin VB.ComboBox cmbMonth 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   8100
      Width           =   2265
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Type Employee Code"
      Top             =   8010
      Width           =   2235
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00C0C000&
      Caption         =   "Generate Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3330
      TabIndex        =   12
      Top             =   8610
      Width           =   1425
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   10
      Top             =   5880
      Width           =   1050
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
      Height          =   390
      Left            =   4560
      TabIndex        =   11
      Top             =   5880
      Width           =   1035
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   27
      Top             =   5520
      Width           =   5895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Absent Employee Information"
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
      Left            =   660
      TabIndex        =   26
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   30
      TabIndex        =   18
      Top             =   5490
      Width           =   5745
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Year  : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   8160
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Month : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   8040
      Width           =   1245
   End
End
Attribute VB_Name = "frmrabsentemployee"
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

Private Sub cboSubDept_Click()
cboempcode.Clear
End Sub

Private Sub cboSubDept_LostFocus()
    cboempcode.Clear
    'cboToCode.Clear

    If Len(cboDept) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set r = New ADODB.Recordset
            StrRecord = "SELECT EmployeeCode FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (Department='" & cboDept.Text & "')" _
                & " AND ([Section]='" & cboSubDept.Text & "') AND (MLeave = 0) AND (MLeft = 0) AND (WStatus = '" & FrameCategory.Tag & "'));"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
        Do Until r.EOF
            cboempcode.AddItem r![EmployeeCode]
            'cboToCode.AddItem r![EmployeeCode]
        r.MoveNext
        Loop
        End If
        r.Close
    Set r = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdSave_Click()

    If Len(cboDept) = 0 Then
        MsgBox "Select Department Name", vbCritical
        cboDept.SetFocus
    End If
    
    If Len(cboSubDept) = 0 Then
        MsgBox "Select Sub Department Name", vbCritical
        cboSubDept.SetFocus
    End If
    
    If Len(cboempcode) = 0 Then
        MsgBox "Select Employee Code", vbCritical
        cboempcode.SetFocus
    End If
    If Len(txtFrom.Text) = 0 Then
        MsgBox "Enter From Date", vbCritical
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If Len(txtTo.Text) = 0 Then
        MsgBox "Enter To Date", vbCritical
        txtTo.SetFocus
        Exit Sub
    End If

Dim s As Integer
Dim f As Integer
Dim o As Integer


Set r = New ADODB.Recordset

On Error GoTo Err
Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    If OptSick.Value = True Then
    s = lbldays.Caption
    End If
    If Optfp.Value = True Then
    f = lbldays.Caption
    End If
    If Optother.Value = True Then
    o = lbldays.Caption
    End If
    
    Set r = New ADODB.Recordset
    StrRecord = "Select Name from Employee_Information where Employeecode='" & cboempcode & "'"
    r.Open StrRecord, MainConn, adOpenStatic
    
    
    MainConn.BeginTrans

            Mysql = "Insert Into Emp_Absent_Status(CompID,Sickdate,Empcode,EmpName,Department,Section,Sickness,FamilyProb,Others,DeleteRow,UserName,UserPass,ComName,EntryDate,EntryTime) " _
                    & " values(" & CompanyID & ",'" & txtFrom & "','" & cboempcode & "','" & r![Name] & "','" & cboDept & "','" & cboSubDept & "'," & s & "," & f & "," & o & "," & 0 & ",'" & strUser & "','" & StrPass & "','" & getComName & "','" & Date & "',GETDATE());"

            MainComm.CommandText = Mysql

    MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Inserted Successfully", vbInformation
    cboDept.SetFocus
    
    r.Close
    Set r = Nothing
    
    cboDept.ListIndex = -1
    cboSubDept.ListIndex = -1
    cboempcode.ListIndex = -1
    OptSick.Value = 0
    Optfp.Value = 0
    Optother.Value = 0
    txtFrom.Text = ""
    txtTo.Text = ""
    
 Exit Sub
Err:
 MsgBox Err.Description

 
End Sub

Private Sub ChkLeftEmp_Click()
FrmCatLeft.Tag = 1
End Sub

Private Sub ChkRegularEmp_Click()
FrmCatLeft.Tag = 0
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()

    If Len(cboDept) = 0 Then
        MsgBox "Select Department Name", vbCritical
        cboDept.SetFocus
    End If
    
    If Len(cboSubDept) = 0 Then
        MsgBox "Select Department Name", vbCritical
        cboSubDept.SetFocus
    End If
    

    If Len(txtFrom.Text) = 0 Then
        MsgBox "Enter The Start Date..", vbCritical
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If Len(txtTo.Text) = 0 Then
        MsgBox "Enter The End Date..", vbCritical
        txtTo.SetFocus
        Exit Sub
    End If
  
    
    
   '===== For Crystal Report Viewer =============
'    With CrystalReport1
'-----Dim Appl As New CRAXDRT.Application
'-----Dim Report As New CRAXDRT.Report
    
    Set Report = Appl.OpenReport(ReportPath & "Attendance\rptabsentemployee.rpt")
'            .ReportFileName = ReportPath & "Attendance\rptabsentemployee.rpt"
'            .PrintFileName = ReportPath & "Attendance\rptabsentemployee.rpt"
            Report.SQLQueryString = "Select * From Emp_Absent_Status Where (Sickdate between '" & txtFrom & "'and '" & txtTo & "') And Department='" & cboDept & "' and Section='" & cboSubDept & "' and CompID=" & CompanyID & " "
            Report.FormulaFields(1).Text = "'" & txtFrom.Text & "'"
            Report.FormulaFields(2).Text = "'" & txtTo.Text & "'"
'            .Destination = crptToWindow
'            .PrintReport
    Report.DiscardSavedData
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1
'End With

End Sub

Private Sub Command2_Click()
    If strUserlogvar <> cnstUserlogvar Then Exit Sub
    
    If Height = 4950 Then
        Height = 2955
        cmdGenerate.Enabled = False
        cmbMonth.SetFocus
    Else
        Height = 4950
        cmdGenerate.Enabled = True
        txtFrom.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim a As Integer
    Call prcMakeCenter(Me)
    Call DeptName(cboDept)
    MDIfrmMain.lblCaption.Caption = "Employee Absent Information"
    If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"

    txtYear.Text = PeriodYear
    Call prcMonthName(cmbMonth)
    cmbMonth.ListIndex = Month(Date) - 1
    
    ' Call prcConReport(CrystalReport1)
    Set r = New ADODB.Recordset
    StrRecord = "SELECT Worker,JRmgt,SRmgt,Staff FROM Sys_User_Name WHERE UserId='" & bytUserID & "' ;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r![Worker] = 1 Then
    optWorker.Enabled = True
    Else
    optWorker.Enabled = False
    End If
    If r![JRmgt] = 1 Then
    optJrMgt.Enabled = True
    Else
    optJrMgt.Enabled = False
    End If
    If r![SRmgt] = 1 Then
    optSrMgt.Enabled = True
    Else
    optSrMgt.Enabled = False
    End If
    If r![Staff] = 1 Then
    optStaff.Enabled = True
    Else
    optStaff.Enabled = False
    End If
   r.Close
   
'   txtFrom.Enabled = False
'   txtTo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    boJRMgt = False
    MDIfrmMain.lblCaption.Caption = ""
    Set frmrptInsurance = Nothing
End Sub

Private Sub optJrMgt_Click()
    FrameCategory.Tag = "J"
End Sub

Private Sub optSrMgt_Click()
    FrameCategory.Tag = "S"
End Sub

Private Sub optStaff_Click()
    FrameCategory.Tag = "F"
End Sub

Private Sub optWorker_Click()
    FrameCategory.Tag = "W"
End Sub

Private Sub txtFrom_GotFocus()
    Call fncGotFocus(txtFrom)
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFrom_LostFocus()
    Call FormatDate(txtFrom)
End Sub

Private Sub txtTo_GotFocus()
    Call fncGotFocus(txtTo)
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtTo_LostFocus()
    If Len(txtTo.Text) = 0 Then
        MsgBox "Enter To Date", vbCritical
        txtTo.SetFocus
        Exit Sub
    End If
    
    Call FormatDate(txtTo)
    a = DateDiff("d", txtFrom, txtTo)
    lbldays.Caption = a + 1
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub


