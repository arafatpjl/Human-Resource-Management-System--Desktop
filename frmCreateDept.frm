VERSION 5.00
Begin VB.Form frmCreateDept 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3250
   ClientLeft      =   7020
   ClientTop       =   2379
   ClientWidth     =   3718
   ControlBox      =   0   'False
   Icon            =   "frmCreateDept.frx":0000
   LinkTopic       =   "frmledgercreate"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3250
   ScaleWidth      =   3718
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDeptName 
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
         Size            =   8.15
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
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
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
      Height          =   375
      Left            =   1050
      TabIndex        =   2
      Top             =   2610
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Department"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   165
      TabIndex        =   7
      Top             =   240
      Width           =   2805
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   4
      Top             =   2250
      Width           =   3915
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Department :"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1290
      Width           =   1935
   End
End
Attribute VB_Name = "frmCreateDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler

    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
    End If
        
    If Len(txtDeptName) = 0 Then
        MsgBox "Type Name of Department", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    MainConn.BeginTrans

        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Dept_Name(ComID,DeptName,Status,DeleteRow,UserID,IDNo,EntryDate,EntryTime) VALUES(" & CompanyID & ",'" & txtDeptName.Text & "',0,0," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE());"
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

    Call DeptName(frmCreateSubDept.cboDept)
    frmCreateSubDept.cboDept.Text = txtDeptName.Text

    Unload Me

'    frmCreateSubDept.Left = (MDIfrmMain.width - frmCreateSubDept.width) / 2
    frmCreateSubDept.txtSubDept.SetFocus

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateDept = Nothing
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
    StrRecord = "SELECT DeptName FROM Dept_Name WHERE ComID=" & CompanyID & "AND DeptName='" & txtDeptName.Text & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Name of Department", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    End If
    Set r = Nothing
End Sub
