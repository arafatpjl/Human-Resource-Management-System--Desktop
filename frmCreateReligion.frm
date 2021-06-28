VERSION 5.00
Begin VB.Form frmCreateReligion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   7020
   ClientTop       =   2385
   ClientWidth     =   3795
   ControlBox      =   0   'False
   Icon            =   "frmCreateReligion.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   360
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
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
      Height          =   415
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
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
      Height          =   415
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Religion"
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
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Religion"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmCreateReligion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler

    If txtDeptName.Text = "" Then
        MsgBox "Type Name of Religion", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Religion VALUES(" & CompanyID & ",'" & Trim(txtDeptName) & "');"
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

    Call Religion(frmEmpInformation.txtReligion)
    frmEmpInformation.txtReligion.Text = txtDeptName.Text
    Unload Me

    frmEmpInformation.txtReligion.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateReligion = Nothing
End Sub

Private Sub txtDeptName_GotFocus()
    Call fncGotFocus(txtDeptName)
End Sub

Private Sub txtDeptName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtDeptName_LostFocus()
    If txtDeptName.Text <> "" Then
        txtDeptName.Text = UCase(txtDeptName.Text)
        Set r = New ADODB.Recordset
  
        StrRecord = "SELECT Religion FROM Religion WHERE ComID=" & CompanyID & "AND Religion='" & Trim(txtDeptName) & "';"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            MsgBox "Duplicate Name of Religion", vbInformation, cnstMsgInfo
            txtDeptName.SetFocus
        End If
    End If
End Sub
