VERSION 5.00
Begin VB.Form frmAdvanceSalaryPaidUnpaid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8565
   ClientLeft      =   2040
   ClientTop       =   2430
   ClientWidth     =   11730
   ControlBox      =   0   'False
   Icon            =   "frmAdvanceSalaryPaidUnpaid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCategory 
      Caption         =   " Employee Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2160
      TabIndex        =   34
      Tag             =   "W"
      Top             =   1140
      Width           =   2655
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
         Left            =   1560
         TabIndex        =   1
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optWorker 
         Caption         =   "Workers"
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
         TabIndex        =   0
         Tag             =   "W"
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   840
      TabIndex        =   30
      Top             =   6270
      Width           =   4035
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   345
         Left            =   150
         TabIndex        =   32
         Top             =   390
         Width           =   1665
      End
      Begin VB.CommandButton cmdAddPresent 
         Caption         =   "Add Present Emp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2190
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Present Date  : "
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
         Left            =   150
         TabIndex        =   33
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Emp Code Department wise"
      Enabled         =   0   'False
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
      Height          =   405
      Left            =   2430
      TabIndex        =   23
      Top             =   3810
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   3000
      TabIndex        =   21
      Top             =   4590
      Width           =   1875
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0C000&
         Caption         =   "&Add to List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   90
         TabIndex        =   29
         Top             =   180
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C0C000&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   90
         TabIndex        =   28
         Top             =   660
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveAll 
         BackColor       =   &H00C0C000&
         Caption         =   "Remove &All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   90
         TabIndex        =   27
         Top             =   1140
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   840
      TabIndex        =   22
      Top             =   4590
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000013&
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
         Height          =   415
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1110
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "&Unpaid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   630
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000013&
         Caption         =   "&Paid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   150
         Width           =   1695
      End
   End
   Begin VB.ComboBox cboFromCode 
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
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Select Employee Code No."
      Top             =   3030
      Width           =   2415
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
      Left            =   2400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select Sub Department"
      Top             =   2610
      Width           =   2415
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
      Left            =   2400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Name of Department"
      Top             =   2190
      Width           =   2415
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
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
      TabIndex        =   12
      Top             =   7200
      Width           =   1935
   End
   Begin VB.ComboBox cboEmpCode 
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
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Select Employee Code No."
      Top             =   3450
      Width           =   2415
   End
   Begin VB.ListBox CodeList 
      BackColor       =   &H8000000E&
      Columns         =   5
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5730
      ItemData        =   "frmAdvanceSalaryPaidUnpaid.frx":000C
      Left            =   5040
      List            =   "frmAdvanceSalaryPaidUnpaid.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   6495
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select Name of Month"
      Top             =   1770
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   37
      Top             =   7680
      Width           =   11775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Paid / Unpaid Advance Salary"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   -975
      TabIndex        =   36
      Top             =   240
      Width           =   6870
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
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
      Left            =   2160
      TabIndex        =   20
      Top             =   3030
      Width           =   255
   End
   Begin VB.Label lblFrom 
      Caption         =   "From Emp Code"
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
      Left            =   360
      TabIndex        =   19
      Top             =   3030
      Width           =   1335
   End
   Begin VB.Label Label59 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2610
      Width           =   1455
   End
   Begin VB.Label Label60 
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
      Left            =   2160
      TabIndex        =   17
      Top             =   2610
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
      Index           =   2
      Left            =   2160
      TabIndex        =   16
      Top             =   2190
      Width           =   255
   End
   Begin VB.Label Label58 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2190
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label lblFrom1 
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
      Left            =   9360
      TabIndex        =   13
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "To Emp Code"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3450
      Width           =   1335
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   3450
      Width           =   255
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
      Left            =   360
      TabIndex        =   8
      Top             =   1770
      Width           =   1335
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
      Left            =   2160
      TabIndex        =   7
      Top             =   1770
      Width           =   255
   End
End
Attribute VB_Name = "frmAdvanceSalaryPaidUnpaid"
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
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboEmpCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
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
    If cboMonthName.Text <> "" And cboDept.Text <> "" And cboSubDept.Text <> "" Then
        Call EmpCodeAdvanceSalary(cboFromCode, cboDept.Text, cboSubDept.Text, cboMonthName.Text, FrameCategory.Tag)
        Call EmpCodeAdvanceSalary(cboEmpCode, cboDept.Text, cboSubDept.Text, cboMonthName.Text, FrameCategory.Tag)
    End If
End Sub

Private Sub chkShow_Click()
    If chkShow.Value = 0 Then Exit Sub
      
    cboFromCode.Clear
    cboEmpCode.Clear
      
    Screen.MousePointer = vbHourglass
    Set r = New ADODB.Recordset
      
    If boJRMgt = True Then
        StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "' " _
            & " AND Department = '" & cboDept.Text & "' AND Year='" & PeriodYear & "' AND WStatus = 'J' GROUP BY EmpCode;"
    Else
        StrRecord = "SELECT empcode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "' " _
            & " AND Department = '" & cboDept.Text & "' AND Year = '" & PeriodYear & "' AND WStatus = '" & FrameCategory.Tag & "' GROUP BY EmpCode;"
    End If
    
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        Do Until r.EOF
            cboFromCode.AddItem r![empcode]
            cboEmpCode.AddItem r![empcode]
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
        If optStaff.Value = True And Len(cboDept.Text) = 0 And Len(cboSubDept.Text) = 0 Then
         StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "'" _
                    & " AND Year = '" & PeriodYear & "' AND WStatus =  '" & FrameCategory.Tag & "' AND MLeft = 0" _
                    & " GROUP BY EmpCode;"
        
        Else
        If boJRMgt = True Then
            If Len(cboSubDept.Text) = 0 Then
                StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "'" _
                    & " AND Year = '" & PeriodYear & "' AND Department = '" & cboDept.Text & "' AND WStatus = 'J' AND MLeft = 0" _
                    & " GROUP BY EmpCode HAVING EmpCode Between '" & cboFromCode.Text & "' And '" & cboEmpCode.Text & "';"
            Else
                StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "'" _
                    & " AND Year = '" & PeriodYear & "' AND Department = '" & cboDept.Text & "' and Section = '" & cboSubDept.Text & "'  AND WStatus = 'J' AND MLeft = 0" _
                    & " GROUP BY EmpCode HAVING EmpCode Between '" & cboFromCode.Text & "' And '" & cboEmpCode.Text & "';"
            End If
        Else
            If Len(cboSubDept.Text) = 0 Then
                StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "'" _
                    & " AND Year = '" & PeriodYear & "' AND Department = '" & cboDept.Text & "' AND WStatus =  '" & FrameCategory.Tag & "' AND MLeft = 0" _
                    & " GROUP BY EmpCode;"
            Else
                StrRecord = "SELECT EmpCode FROM VIEW_EMP_ADVANCE_SALARY WHERE ComID = " & CompanyID & " AND Month = '" & cboMonthName.Text & "'" _
                    & " AND Year = '" & PeriodYear & "' AND Department = '" & cboDept.Text & "' and Section = '" & cboSubDept.Text & "'  AND WStatus =  '" & FrameCategory.Tag & "' AND MLeft = 0" _
                    & " GROUP BY EmpCode HAVING EmpCode Between '" & cboFromCode.Text & "' And '" & cboEmpCode.Text & "';"
            End If
        End If
        End If
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                CodeList.AddItem r![empcode]
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

'    If Len(cboSubDept) = 0 Then
'        MsgBox "Select Sub Dept.", vbInformation, cnstMsgInfo
'        cboSubDept.SetFocus
'    Exit Sub
'    End If

    CodeList.Clear
    txtTotal.Text = ""

    Screen.MousePointer = vbHourglass
        Set r = New ADODB.Recordset
        If boJRMgt = True Then
            StrRecord = "SELECT Emp_Daily_Attendance.EmpCode FROM Employee_Information INNER JOIN Emp_Daily_Attendance ON Employee_Information.ComID = Emp_Daily_Attendance.ComID" _
                & " AND Employee_Information.EmployeeCode = Emp_Daily_Attendance.EmpCode WHERE ((Employee_Information.ComID =" & CompanyID & ") AND (Employee_Information.MLeft = 0)" _
                & " AND (Employee_Information.WStatus = 'J') AND (Employee_Information.MLeave = 0) " _
                & " AND (Emp_Daily_Attendance.DailyDate = '" & Format(txtDate.Text, cnstDtFrmtQ) & "') AND (Emp_Daily_Attendance.TimeIn <> '00:00'));"
        Else
            StrRecord = "SELECT Emp_Daily_Attendance.EmpCode FROM Employee_Information INNER JOIN Emp_Daily_Attendance ON Employee_Information.ComID = Emp_Daily_Attendance.ComID" _
                & " AND Employee_Information.EmployeeCode = Emp_Daily_Attendance.EmpCode WHERE ((Employee_Information.ComID =" & CompanyID & ") AND (Employee_Information.MLeft = 0)" _
                & " AND (Employee_Information.WStatus =  '" & FrameCategory.Tag & "') AND (Employee_Information.MLeave = 0) " _
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
    If CodeList.ListCount = 0 Then
        MsgBox "Record not found to Paid", vbInformation, cnstMsgInfo
        cboEmpCode.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to Paid the advance salary for existing employee", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command

        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500

        MainConn.BeginTrans
            For intI = 0 To CodeList.ListCount - 1
                CodeList.Selected(intI) = True
                MainComm.CommandText = "UPDATE Temp_Emp_Advance_Salary SET Status = 1 WHERE ComID=" & CompanyID & " AND EmpCode='" & CodeList.Text & "';"
                MainComm.Execute
            Next
        MainConn.CommitTrans
    Screen.MousePointer = vbDefault
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

Private Sub Command1_Click()
    If CodeList.ListCount = 0 Then
        MsgBox "Record not found to Unpaid", vbInformation, cnstMsgInfo
        cboEmpCode.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to Unpaid the salary of existing employee", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500

        MainConn.BeginTrans
            For intI = 0 To CodeList.ListCount - 1
                CodeList.Selected(intI) = True
                
                MainComm.CommandText = "UPDATE Temp_Emp_Advance_Salary SET Status = 0 WHERE ComID=" & CompanyID & " AND EmpCode='" & CodeList.Text & "' ;"
                MainComm.Execute
            Next
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
    Top = (frmAdvanceSalaryPaidUnpaid.Height - Height) / 2
  Left = (frmAdvanceSalaryPaidUnpaid.width - width) / 2 + 1600
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    MDIfrmMain.lblCaption.Caption = "Advance Salary Paid / Unpaid Entry"
    
    Call DeptName(cboDept)
    Call prcMonthName(cboMonthName)
    cboMonthName.Text = UCase(Format(Date, "MMMM"))
    Set r = New ADODB.Recordset
    StrRecord = "SELECT Worker,JRmgt,SRmgt,Staff FROM Sys_User_Name WHERE UserId='" & bytUserID & "' ;"
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
    Set frmAdvanceSalaryPaidUnpaid = Nothing
End Sub

Private Sub optStaff_Click()
    FrameCategory.Tag = "F"
End Sub

Private Sub optStaff_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optWorker_Click()
    FrameCategory.Tag = "W"
End Sub

Private Sub optWorker_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub

Private Sub txtdate_LostFocus()
  If Len(txtDate) = 0 Then Exit Sub
  txtDate.Text = FormatDate(txtDate)
End Sub
