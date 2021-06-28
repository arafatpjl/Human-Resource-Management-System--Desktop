VERSION 5.00
Begin VB.Form frmCreateUser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5940
   ClientLeft      =   2205
   ClientTop       =   2040
   ClientWidth     =   7935
   ControlBox      =   0   'False
   Icon            =   "frmCreateUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
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
      Left            =   5670
      TabIndex        =   27
      Top             =   5310
      Width           =   1905
   End
   Begin VB.TextBox txtUserName 
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
      Height          =   315
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.ComboBox cboUserName 
      BackColor       =   &H80000016&
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
      ItemData        =   "frmCreateUser.frx":000C
      Left            =   2160
      List            =   "frmCreateUser.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdSave 
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
      Height          =   415
      Left            =   3690
      TabIndex        =   8
      Top             =   5310
      Width           =   1905
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
      Left            =   270
      TabIndex        =   9
      Top             =   5280
      Width           =   1905
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00C0C000&
      Caption         =   "&Show Record"
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
      Left            =   6000
      TabIndex        =   26
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtVPassword 
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
      Height          =   315
      Left            =   5760
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   315
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2655
      Left            =   360
      TabIndex        =   19
      Top             =   2160
      Width           =   2175
      Begin VB.OptionButton optSysAdmin 
         BackColor       =   &H80000016&
         Caption         =   "System Admin"
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
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optExeOfficer 
         BackColor       =   &H80000016&
         Caption         =   "Executive Officer"
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
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optOperator 
         BackColor       =   &H80000016&
         Caption         =   "Operator"
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
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optOfficer 
         BackColor       =   &H80000016&
         Caption         =   "Officer"
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
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Working Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2655
      Left            =   2640
      TabIndex        =   10
      Top             =   2160
      Width           =   4935
      Begin VB.ListBox List2 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2010
         ItemData        =   "frmCreateUser.frx":0010
         Left            =   2760
         List            =   "frmCreateUser.frx":0012
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
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
         Left            =   2270
         TabIndex        =   15
         ToolTipText     =   "Add"
         Top             =   600
         Width           =   420
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   ">>"
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
         Left            =   2270
         TabIndex        =   14
         ToolTipText     =   "Add All"
         Top             =   960
         Width           =   420
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
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
         Left            =   2270
         TabIndex        =   13
         ToolTipText     =   "Remove "
         Top             =   1440
         Width           =   420
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "<>"
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
         Left            =   2270
         TabIndex        =   12
         ToolTipText     =   "Remove All"
         Top             =   1800
         Width           =   420
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2010
         ItemData        =   "frmCreateUser.frx":0014
         Left            =   240
         List            =   "frmCreateUser.frx":0016
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Full Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Permitted Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   30
      Top             =   5040
      Width           =   7935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Or Modify User"
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
      Left            =   75
      TabIndex        =   29
      Top             =   240
      Width           =   3585
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
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
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   1920
      TabIndex        =   24
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Verify Password"
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
      Left            =   4080
      TabIndex        =   23
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Password"
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
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   5520
      TabIndex        =   21
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label6 
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
      Index           =   3
      Left            =   1920
      TabIndex        =   20
      Top             =   1680
      Width           =   255
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboUserName_Click()
    If cboUserName.Text = "" Then Exit Sub

    Set r = New ADODB.Recordset
    StrRecord = "SELECT * FROM User_Log_On WHERE UserNM='" & EncryptIt(cboUserName, 11) & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        txtPassword.Text = DecryptIt(r![Password], 11)
        txtVPassword.Text = DecryptIt(r![Password], 11)

        If DecryptIt(r![Status], 11) = "System Admin" Then
            optSysAdmin.Value = True
        ElseIf DecryptIt(r![Status], 11) = "Executive Officer" Then
            optExeOfficer.Value = True
        ElseIf DecryptIt(r![Status], 11) = "Operator" Then
            optOperator.Value = True
        ElseIf DecryptIt(r![Status], 11) = "Officer" Then
            optOfficer.Value = True
        End If

        List2.Clear
        If r![NewEntryEmp] = 1 Then
            List2.AddItem "New Employee Entry"
        End If
        If r![IncreEntry] = 1 Then
            List2.AddItem "Increment Entry"
        End If
        If r![MonEntrySal] = 1 Then
            List2.AddItem "Monthly Entry - Salary"
        End If
        If r![MonEntryOT] = 1 Then
            List2.AddItem "Monthly Entry - OT"
        End If
        If r![PaidUnpaidSal] = 1 Then
            List2.AddItem "Paid/Unpaid - Salary"
        End If
        If r![PaidUnpaidOT] = 1 Then
            List2.AddItem "Paid/Unpaid - OT"
        End If
        If r![DataEditSal] = 1 Then
            List2.AddItem "Data Edit - Salary"
        End If
        If r![DataEditOT] = 1 Then
            List2.AddItem "Data Edit - OT"
        End If
        If r![ReportSal] = 1 Then
            List2.AddItem "Report - Salary"
        End If
        If r![ReportOT] = 1 Then
            List2.AddItem "Report - OT"
        End If
        If r![JRMgt] = 1 Then
            List2.AddItem "JR Management"
        End If
        If r![Option1] = 1 Then
            List2.AddItem "Option"
        End If
        If r![LockDataEdit] = 1 Then
            List2.AddItem "Lock Data Edit"
        End If
    End If
    r.Close
    Set r = Nothing
End Sub

Private Sub cboUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAdd_Click()
    If List1.Text <> "" Then
        List2.AddItem List1.Text
    End If
End Sub
Private Sub cmdAddAll_Click()
    List2.Clear
    List1.ListIndex = -1

    For intI = 1 To List1.ListCount
        List1.ListIndex = List1.ListIndex + 1
        List2.AddItem List1.Text
    Next intI
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cboUserName.Text = "" Then Exit Sub

    If MsgBox("Are you sure to Delete this record", vbInformation + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON DELETE FROM User_Log_On WHERE UserNM='" & EncryptIt(cboUserName, 11) & "';"
        MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

    txtPassword.Text = ""
    txtVPassword.Text = ""

    List2.Clear
    optSysAdmin.Value = False
    optExeOfficer.Value = False
    optOfficer.Value = False
    optOperator.Value = False

    Call AddUserName(cboUserName)
    cboUserName.SetFocus
End Sub

Private Sub cmdModify_Click()
    txtUserName.Visible = False
    cboUserName.Visible = True
    txtPassword.Text = ""
    txtVPassword.Text = ""
    
    List2.Clear
    
    optSysAdmin.Value = False
    optExeOfficer.Value = False
    optOfficer.Value = False
    optOperator.Value = False
    cboUserName.SetFocus
    
    cmdSave.Caption = "&Update"
End Sub

Private Sub cmdRemove_Click()
    If List2.ListIndex >= 0 Then
        List2.RemoveItem List2.ListIndex
    End If
End Sub

Private Sub cmdRemoveAll_Click()
    List2.Clear
End Sub

Private Sub cmdSave_Click()
    Dim UserName As String, UserPassword As String, Status As String

    Dim V_NewEntry As Byte
    Dim V_MonEntrySal As Byte
    Dim V_MonEntryOT As Byte
    Dim V_PaidSal As Byte
    Dim V_PaidOT As Byte
    Dim V_IncrEntry As Byte
    Dim V_EditSal As Byte
    Dim V_EditOT As Byte
    Dim V_ReportSal As Byte
    Dim V_ReportOT As Byte
    Dim V_JRMgt As Byte
    Dim V_Option As Byte
    Dim V_LockDataEdit As Byte

    If txtUserName.Text = "" And cmdSave.Caption = "&Save" Then
        MsgBox "Type User Name", vbInformation, cnstMsgInfo
        txtUserName.SetFocus
    Exit Sub
    End If
    
    If txtPassword.Text = "" Then
        MsgBox "Type User Password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
    Exit Sub
    End If

    If txtVPassword.Text = "" Then
        MsgBox "Type Verify Password", vbInformation, cnstMsgInfo
        txtVerifyPassword.SetFocus
    Exit Sub
    End If

    If Trim(txtPassword.Text) <> Trim(txtVPassword.Text) Then
        MsgBox "User password does not match with Verify Password", vbInformation, cnstMsgInfo
        txtVPassword.SetFocus
        
        SendKeys "{Home}+{End}"
    Exit Sub
    End If

    If optSysAdmin.Value = False And optExeOfficer.Value = False And optOperator.Value = False And optOfficer.Value = False Then
        MsgBox "Select any Status", vbInformation, cnstMsgInfo
    Exit Sub
    End If

    If optSysAdmin.Value = False And List2.ListCount = 0 Then
        MsgBox "Permission not found", vbInformation, cnstMsgInfo
    Exit Sub
    End If

    Set r = New ADODB.Recordset

    If optSysAdmin.Value = True Then
        Status = optSysAdmin.Caption
  
        StrRecord = "SELECT Status FROM User_Log_On WHERE Status='" & EncryptIt(optSysAdmin.Caption, 11) & "';"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            MsgBox "Status not allowed", vbInformation, cnstMsgInfo
        Exit Sub
        End If
        r.Close
    ElseIf optExeOfficer.Value = True Then
        Status = optExeOfficer.Caption
    ElseIf optOperator.Value = True Then
        Status = optOperator.Caption
    ElseIf optOfficer.Value = True Then
        Status = optOfficer.Caption
    End If

    If cmdSave.Caption = "&Save" Then
        UserName = EncryptIt(Trim(txtUserName.Text), 11)
    Else
        UserName = EncryptIt(Trim(cboUserName.Text), 11)
    End If
    UserPassword = EncryptIt(Trim(txtPassword.Text), 11)
    Status = EncryptIt(Status, 11)

    If cmdSave.Caption = "&Save" Then
        StrRecord = "SELECT UserNM FROM User_Log_On WHERE UserNM='" & UserName & "';"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            MsgBox "Duplicate User Name", vbInformation, cnstMsgInfo
            txtUserName.SetFocus
            SendKeys "{Home}+{End}"
        Exit Sub
        End If
        r.Close
    End If
''    StrRecord = "SELECT Password FROM User_Log_On WHERE Password='" & UserPassword & "';"
''    r.Open StrRecord, MainConn, adOpenStatic
''    If r.EOF = False And r.BOF = False Then
''        MsgBox "Duplicate password", vbInformation, cnstMsgInfo
''        txtPassword.SetFocus
''        SendKeys "{Home}+{End}"
''    Exit Sub
''    End If
''    r.Close

    V_NewEntry = 0
    V_MonEntrySal = 0
    V_MonEntryOT = 0
    V_PaidSal = 0
    V_PaidOT = 0
    V_IncrEntry = 0
    V_EditSal = 0
    V_EditOT = 0
    V_ReportSal = 0
    V_ReportOT = 0
    V_JRMgt = 0
    V_Option = 0
    V_LockDataEdit = 0

    List2.ListIndex = -1

    For intI = 1 To List2.ListCount
        List2.ListIndex = List2.ListIndex + 1

        If List2.Text = "New Employee Entry" Then
            V_NewEntry = 1
        ElseIf List2.Text = "Increment Entry" Then
            V_IncrEntry = 1
        ElseIf List2.Text = "Monthly Entry - Salary" Then
            V_MonEntrySal = 1
        ElseIf List2.Text = "Monthly Entry - OT" Then
            V_MonEntryOT = 1
        ElseIf List2.Text = "Paid/Unpaid - Salary" Then
            V_PaidSal = 1
        ElseIf List2.Text = "Paid/Unpaid - OT" Then
            V_PaidOT = 1
        ElseIf List2.Text = "Data Edit - Salary" Then
            V_EditSal = 1
        ElseIf List2.Text = "Data Edit - OT" Then
            V_EditOT = 1
        ElseIf List2.Text = "Report - Salary" Then
            V_ReportSal = 1
        ElseIf List2.Text = "Report - OT" Then
            V_ReportOT = 1
        ElseIf List2.Text = "JR Management" Then
            V_JRMgt = 1
        ElseIf List2.Text = "Option" Then
            V_Option = 1
        ElseIf List2.Text = "Lock Data Edit" Then
            V_LockDataEdit = 1
        End If
    Next intI

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
        
    If cmdSave.Caption = "&Save" Then
        MainConn.BeginTrans
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO User_Log_On VALUES('" & UserName & "','" & UserPassword & "','" & Status & "'," & V_NewEntry & "," & V_MonEntrySal & "," & V_MonEntryOT & "," _
                & " " & V_PaidSal & "," & V_PaidOT & "," & V_IncrEntry & "," & V_EditSal & "," & V_EditOT & "," & V_ReportSal & "," & V_ReportOT & "," & V_Option & "," & V_JRMgt & ", " & V_LockDataEdit & ",0);"
            MainComm.Execute
        MainConn.CommitTrans
    ElseIf cmdSave.Caption = "&Update" Then
        MainConn.BeginTrans
            MainComm.CommandText = " Update User_Log_On Set Password = '" & UserPassword & "', Status = '" & Status & "', NewEntryEmp = " & V_NewEntry & ", MonEntrySal = " & V_MonEntrySal & ", MonEntryOT =" & V_MonEntryOT & ", " & _
                " PaidUnpaidSal = " & V_PaidSal & ", PaidUnpaidOT = " & V_PaidOT & ", IncreEntry =" & V_IncrEntry & ", DataEditSal = " & V_EditSal & ", DataEditOT = " & V_EditOT & ", ReportSal = " & V_ReportSal & ", ReportOT = " & V_ReportOT & ", " & _
                " Option1 = " & V_Option & ", JRMgt = " & V_JRMgt & ", LockDataEdit = " & V_LockDataEdit & "" & _
                " Where UserNM = '" & UserName & "';"
            MainComm.Execute
        MainConn.CommitTrans
    End If
ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  Err.Number = 0
  Exit Sub
End If

MsgBox "Successfully Saved", vbInformation, cnstMsgInfo

    txtUserName.Text = ""
    txtPassword.Text = ""
    txtVPassword.Text = ""
    optSysAdmin.Value = False
    optExeOfficer.Value = False
    optOperator.Value = False
    optOfficer.Value = False
    List2.Clear
    List1.ListIndex = -1
If cmdSave.Caption = "&Save" Then
    txtUserName.SetFocus
Else
    cboUserName.SetFocus
End If

Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Load()
    MDIfrmMain.lblCaption.Caption = "Create or Modify User"

    List1.AddItem "New Employee Entry"
    List1.AddItem "Increment Entry"
    
    List1.AddItem "Monthly Entry - Salary"
    List1.AddItem "Monthly Entry - OT"
    List1.AddItem "Paid/Unpaid - Salary"
    List1.AddItem "Paid/Unpaid - OT"
    List1.AddItem "Data Edit - Salary"
    List1.AddItem "Data Edit - OT"
    List1.AddItem "Report - Salary"
    List1.AddItem "Report - OT"
    List1.AddItem "JR Management"
    List1.AddItem "Option"
    List1.AddItem "Lock Data Edit"

    Call AddUserName(cboUserName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmCreateUser = Nothing
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub
