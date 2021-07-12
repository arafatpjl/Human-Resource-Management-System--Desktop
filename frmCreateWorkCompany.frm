VERSION 5.00
Begin VB.Form frmCreateWorkCompany 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   7020
   ClientTop       =   2385
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Icon            =   "frmCreateWorkCompany.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPreComName 
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
      Left            =   480
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1650
      Width           =   2775
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
      Left            =   2340
      TabIndex        =   1
      Top             =   2610
      Width           =   1335
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
      Left            =   930
      TabIndex        =   2
      Top             =   2610
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Create Previous Worked Company"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label2 
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
      Height          =   45
      Left            =   0
      TabIndex        =   4
      Top             =   2250
      Width           =   3915
   End
   Begin VB.Label Label1 
      Caption         =   "Previous Company"
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
      TabIndex        =   3
      Top             =   1290
      Width           =   1935
   End
End
Attribute VB_Name = "frmCreateWorkCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    If Len(txtPreComName) = 0 Then
        MsgBox "Type Name of Company", vbInformation, cnstMsgInfo
        txtPreComName.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    'intMaxComID
    MainConn.BeginTrans
        'MainComm.CommandText = "SET NOCOUNT ON INSERT INTO New_Wrok_Company VALUES(" & intMaxComID & ",'" & txtPreComName.Text & "');"
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO New_Wrok_Company(ComName) VALUES('" & txtPreComName.Text & "');"
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

    Call deptname(frmCreateSubDept.cboDept)
    Call addWorkComName(frmEmpInformation.cmbWorkComName)
    frmEmpInformation.cmbWorkComName.Text = txtPreComName.Text

    Unload Me

'    frmCreateSubDept.Left = (MDIfrmMain.width - frmCreateSubDept.width) / 2
    frmEmpInformation.cmbWorkComName.SetFocus

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateWorkCompany = Nothing
End Sub

Private Sub txtPreComName_GotFocus()
    Call fncGotFocus(txtPreComName)
End Sub

Private Sub txtPreComName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPreComName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtPreComName_LostFocus()
    If Len(txtPreComName) = 0 Then Exit Sub
    txtPreComName.Text = Trim(UCase(txtPreComName.Text))
    
    Set r = New ADODB.Recordset
    StrRecord = "SELECT Company FROM Emp_Company WHERE Company='" & txtPreComName.Text & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Name of Company", vbInformation, cnstMsgInfo
        txtPreComName.SetFocus
    End If
    Set r = Nothing
End Sub
