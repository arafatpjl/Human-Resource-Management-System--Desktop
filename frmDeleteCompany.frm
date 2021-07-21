VERSION 5.00
Begin VB.Form frmDeleteCompany 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   3015
   ClientTop       =   2565
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmDeleteCompany.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00800080&
      Height          =   1155
      IMEMode         =   3  'DISABLE
      Left            =   480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmDeleteCompany.frx":0442
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtCompany 
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
      ForeColor       =   &H00800080&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
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
      Height          =   415
      Left            =   480
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete  Company"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Delete Company"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Company"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "frmDeleteCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdOpen_Click()

Dim V_DeleteCompany As String

If MsgBox("Are you sure to Delete the Company of " & txtCompany.Text, vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

Set r = New ADODB.Recordset
StrRecord = "SELECT DeptName FROM Dept_Name WHERE ComID=" & CompanyID & ";"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then

  Open App.Path + "\CDelete.sys" For Input As #1

  Input #1, V_DeleteCompany
  Close #1

  If LCase(V_DeleteCompany) <> "deletecomppacpay" Then
    MsgBox "You are not a right person to Delete Company," + vbCrLf + "Please contact the right person", vbInformation, cnstMsgInfo
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
End If
r.Close
Set r = Nothing

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Company_Information WHERE ID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Dept_Name WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Designation WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM District WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Dolllar_Rate_Entry WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Emp_Edu_Qualification WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Employee_Information WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Increment_Information WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM OverTime_Entry WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Religion WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Salary_And_Wages WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Sub_Dept_Name WHERE ComID=" & CompanyID & ";"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Bonus_Entry WHERE ComID=" & CompanyID & ";"
MainComm.Execute

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  If Err.Number = 53 Or Err.Number = 62 Then
    MsgBox "You are not a right person to Delete Company," + vbCrLf + "Please contact the right person", vbInformation, cnstMsgInfo
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  Err.Number = 0
  Exit Sub
End If

Screen.MousePointer = vbDefault

MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

End

End Sub
Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Delete Company"
  txtCompany.Text = MDIfrmMain.lblCompanyName.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmDeleteCompany = Nothing
End Sub
