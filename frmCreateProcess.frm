VERSION 5.00
Begin VB.Form frmCreateProcess 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2990
   ClientLeft      =   2925
   ClientTop       =   3536
   ClientWidth     =   4914
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCreateProcess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2990
   ScaleWidth      =   4914
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
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
      Height          =   415
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
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
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtProcess 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   -840
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   -840
      X2              =   5040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "  Create Process"
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
      Left            =   1110
      TabIndex        =   5
      Top             =   120
      Width           =   2265
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
      Left            =   -120
      TabIndex        =   4
      Top             =   2730
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Process Name :"
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
      TabIndex        =   3
      Top             =   1050
      Width           =   1545
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmCreateProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreSubGroup_Click()
    frmCreateDept.Show 1
End Sub

Private Sub cmdSave_Click()
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                    Exit Sub
    End If

    If txtProcess.Text = "" Then
        MsgBox "Type Name of Process", vbInformation, cnstMsgInfo
        txtProcess.SetFocus
    Exit Sub
    End If
    

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
        Set r = New ADODB.Recordset
        StrRecord = "SELECT ProcessName FROM New_Process WHERE ProcessName='" & Trim(txtProcess) & "';"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            MsgBox "Duplicate Process Name", vbInformation, cnstMsgInfo
            txtProcess.SetFocus
        Exit Sub
        End If

        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500

        MainConn.BeginTrans
                     
                    MainComm.CommandText = "SET NOCOUNT ON INSERT INTO New_Process(ProcessName) VALUES('" & Trim(txtProcess) & "');"
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
    
    MsgBox "Successfully Added", vbInformation

    txtProcess.Text = ""
    txtProcess.SetFocus
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    MDIfrmMain.lblCaption.Caption = "Create Process Name"
    Call prcMakeCenter(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIfrmMain.lblCaption.Caption = ""
    Set frmCreateProcess = Nothing
End Sub

Private Sub txtSubDept_GotFocus()
'    Call fncGotFocus(txtSubDept)
End Sub

Private Sub txtSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtSubDept_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtSubDept_LostFocus()
    txtSubDept.Text = UCase(txtSubDept.Text)
End Sub
