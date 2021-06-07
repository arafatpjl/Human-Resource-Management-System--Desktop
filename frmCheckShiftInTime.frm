VERSION 5.00
Begin VB.Form frmCheckShiftInTime 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   3195
   ClientTop       =   2805
   ClientWidth     =   5880
   ControlBox      =   0   'False
   Icon            =   "frmCheckShiftInTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtShiftName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2865
   End
   Begin VB.TextBox txtInTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2865
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   1140
      Width           =   1815
   End
   Begin VB.ComboBox cboDept 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1590
      Width           =   2895
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
      Height          =   390
      Left            =   4470
      TabIndex        =   4
      Top             =   4650
      Width           =   1275
   End
   Begin VB.ComboBox cboEmpCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Name of Month"
      Top             =   2550
      Width           =   2895
   End
   Begin VB.ComboBox cboSubDept 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2070
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   4440
      Width           =   5895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Shift Time Report"
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
      Left            =   195
      TabIndex        =   21
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -60
      TabIndex        =   19
      Top             =   4290
      Width           =   5985
   End
   Begin VB.Label Label7 
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
      Left            =   2010
      TabIndex        =   18
      Top             =   3780
      Width           =   75
   End
   Begin VB.Label Label6 
      Caption         =   "Shift Start Time "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   17
      Top             =   3780
      Width           =   1545
   End
   Begin VB.Label Label5 
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
      Left            =   2040
      TabIndex        =   14
      Top             =   1170
      Width           =   105
   End
   Begin VB.Label Label4 
      Caption         =   "Working Date"
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
      Left            =   480
      TabIndex        =   13
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label Label3 
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
      Left            =   2010
      TabIndex        =   12
      Top             =   3300
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Shift Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Top             =   3270
      Width           =   1545
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
      Left            =   480
      TabIndex        =   10
      Top             =   1620
      Width           =   1545
   End
   Begin VB.Label Label29 
      Caption         =   "Emp Code"
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
      Height          =   225
      Left            =   480
      TabIndex        =   9
      Top             =   2580
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Dept."
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
      Left            =   480
      TabIndex        =   8
      Top             =   2070
      Width           =   1545
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
      Index           =   5
      Left            =   1920
      TabIndex        =   7
      Top             =   1620
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
      Left            =   1920
      TabIndex        =   6
      Top             =   2100
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
      Left            =   1920
      TabIndex        =   5
      Top             =   2610
      Width           =   255
   End
End
Attribute VB_Name = "frmCheckShiftInTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDept_Click()
  If Len(cboDept) = 0 Then Exit Sub
  Call SubDeptName(cboSubDept, cboDept)
  
  cboEmpCode.Clear
  Call prcBlankField
End Sub

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboEmpCode_Click()
    Call prcBlankField
    Call prcLoadData
End Sub

Private Sub prcBlankField()
    txtShiftName.Text = ""
    txtInTime.Text = ""
End Sub

Private Sub prcCheckBlank()
    If Len(txtDate) = 0 Then
        MsgBox "Plese Type Working Date"
        txtDate.SetFocus
        Exit Sub
    End If
    If Len(cboDept) = 0 Then
        MsgBox "Plese Check the Department Name"
        cboDept.SetFocus
        Exit Sub
    End If
    If Len(cboSubDept) = 0 Then
        MsgBox "Plese Check the Section Name"
        cboSubDept.SetFocus
        Exit Sub
    End If
    If Len(cboEmpCode) = 0 Then
        MsgBox "Plese Check the Employee Code"
        cboEmpCode.SetFocus
        Exit Sub
    End If
End Sub

Private Sub prcLoadData()
    Call prcCheckBlank
    
    Set r = New ADODB.Recordset
        StrRecord = "Select ShiftName, ShiftIn From viewCheckShiftInTime Where Department= '" & cboDept.Text & "' And Section = '" & cboSubDept.Text & "' And EmployeeCode = '" & cboEmpCode.Text & "' And DailyDate = '" & txtDate.Text & "'"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
          txtShiftName.Text = r.Fields(0)
          txtInTime.Text = Format(r.Fields(1), "HH-MM-SS")
        End If
        r.Close
        Screen.MousePointer = vbDefault
    Set r = Nothing
End Sub

Private Sub cboEmpCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSubDept_Click()
    cboEmpCode.Clear
    Call prcBlankField
End Sub

Private Sub cboSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSubDept_LostFocus()
  If Len(cboSubDept) = 0 Then Exit Sub
  Set r = New ADODB.Recordset
    StrRecord = "Select EmployeeCode From viewCheckShiftInTime Where Department= '" & cboDept.Text & "' And Section = '" & cboSubDept.Text & "' Group By EmployeeCode Order By EmployeeCode "
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    Do Until r.EOF
      cboEmpCode.AddItem r![EmployeeCode]
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

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    
    Call prcBlankField
    Call DeptName(cboDept)
    'AddShiftName cboShiftName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIfrmMain.lblCaption.Caption = ""
    Set frmCheckShiftInTime = Nothing
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
    Call prcBlankField
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtDate_LostFocus()
    Call FormatDate(txtDate)
End Sub

Private Sub txtInTime_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtShiftName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

