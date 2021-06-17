VERSION 5.00
Begin VB.Form frmCreateDesignation 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4290
   ClientLeft      =   3795
   ClientTop       =   2955
   ClientWidth     =   4350
   ControlBox      =   0   'False
   Icon            =   "frmCreateDesignation.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCategory 
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
      Height          =   315
      Left            =   300
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2220
      Width           =   3855
   End
   Begin VB.TextBox txtDeptName 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   300
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1440
      Width           =   3825
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2610
      TabIndex        =   2
      Top             =   3630
      Width           =   1455
   End
   Begin VB.CommandButton cmdclose 
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
      Height          =   375
      Left            =   1050
      TabIndex        =   3
      Top             =   3630
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Designation"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   45
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   2910
      Width           =   4905
   End
   Begin VB.Label Label3 
      Caption         =   "Select Category :"
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
      Left            =   300
      TabIndex        =   5
      Top             =   1950
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Designation"
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
      Left            =   300
      TabIndex        =   4
      Top             =   1170
      Width           =   1935
   End
End
Attribute VB_Name = "frmCreateDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Len(txtDeptName) = 0 Then
        MsgBox "Type Name of Designation", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    Exit Sub
    End If

    If Len(cboCategory.Text) = 0 Then
        MsgBox "Select designation category.", vbInformation, cnstMsgInfo
        cboCategory.SetFocus
    Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Designation(ComID, Designation, Category, Status) VALUES(" & CompanyID & ",'" & txtDeptName.Text & "', '" & cboCategory.Text & "', 0);"
        MainComm.Execute
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Err.Number = 0
    Exit Sub
    End If

    Call Designation(frmEmpInformation.cboDesignation)
    frmEmpInformation.cboDesignation.Text = txtDeptName.Text

    Unload Me

    frmEmpInformation.txtDepartment.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    Call prcDesigCategory(cboCategory)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateDesignation = Nothing
End Sub

Private Sub txtDeptName_GotFocus()
    Call fncGotFocus(txtDeptName)
End Sub

Private Sub txtDeptName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtDeptName_LostFocus()
    If Len(txtDeptName) = 0 Then Exit Sub
    txtDeptName.Text = Trim(UCase(txtDeptName.Text))

    Set r = New ADODB.Recordset
    StrRecord = "SELECT Designation FROM Designation WHERE ComID=" & CompanyID & " AND Designation='" & txtDeptName.Text & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Name of Designation", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    End If
    Set r = Nothing
End Sub
