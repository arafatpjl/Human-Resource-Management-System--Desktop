VERSION 5.00
Begin VB.Form frmDailyStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3094
   ClientLeft      =   39
   ClientTop       =   39
   ClientWidth     =   4745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3094
   ScaleWidth      =   4745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1155
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Enter Date"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtEmpCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Enter Employee Code"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Mgt.Late Update"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   150
      TabIndex        =   10
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   1065
      TabIndex        =   5
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EmpCode"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   780
   End
End
Attribute VB_Name = "frmDailyStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
DeleteRow = 2
If Len(txtEmpCode.Text) = 0 Then
    MsgBox "Enter Employee Code..", vbInformation, cnstMsgInfo
    txtEmpCode.SetFocus
    Exit Sub
End If

If Len(txtDate.Text) = 0 Then
    MsgBox "Enter Date..", vbInformation, cnstMsgInfo
    txtDate.SetFocus
    Exit Sub
End If

    Set MainComm = New ADODB.Command
On Error GoTo ErrGotoLabel
    MainComm.ActiveConnection = MainConn
    Screen.MousePointer = vbHourglass
    MainConn.BeginTrans
If txtEmpCode.Text = "008" Or txtEmpCode.Text = "021" Or txtEmpCode.Text = "022" Then
    MainComm.CommandText = "Update Emp_Daily_Attendance Set EmpStatus='P',UserName='" & strUser & "',UserPass='" & StrPass & "',ComputerName='" & getComName & "',EntryDate='" & Date & "',EntryTime=GETDATE() Where ComID=" & CompanyID & " And  EmpCode='" & txtEmpCode.Text & "' And DailyDate='" & txtDate.Text & "' And EmpStatus='L' ;"
    MainComm.Execute
Else
    MsgBox "Enter the proper Employee Code", vbInformation, cnstMsgInfo
    txtEmpCode.SetFocus
    Exit Sub
End If
    MainConn.CommitTrans
MsgBox "Successfully Updated.", vbOKOnly, cnstMsgInfo
ErrGotoLabel:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, cnstMsgErDB
    Screen.MousePointer = vbDefault
    MainConn.RollbackTrans
    Err.Number = 0
    Exit Sub
End If
Screen.MousePointer = vbDefault
txtEmpCode.Text = ""
txtDate.Text = ""
End Sub

Private Sub txtDate_GotFocus()
    Call fncGotFocus(txtDate)
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtDate_LostFocus()
    txtDate.Text = FormatDate(txtDate)
End Sub

Private Sub txtEmpCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
