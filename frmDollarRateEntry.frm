VERSION 5.00
Begin VB.Form frmDollarRateEntry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2912
   ClientLeft      =   3432
   ClientTop       =   3042
   ClientWidth     =   4836
   ControlBox      =   0   'False
   Icon            =   "frmDollarRateEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1719.324
   ScaleMode       =   0  'User
   ScaleWidth      =   4535.108
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
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
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Check Password"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox cboMonthName 
      BackColor       =   &H8000000E&
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Month"
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Check Password"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Form Close"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Set Dollar Rate"
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
      Left            =   210
      TabIndex        =   10
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label43 
      Caption         =   "Rate Per US $"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label29 
      Caption         =   "Name of Month"
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
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmDollarRateEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboMonthName_Click()

Set r = New ADODB.Recordset

If cboMonthName.Text <> "" Then
  
  StrRecord = "SELECT Rate FROM Dolllar_Rate_Entry WHERE Month='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND ComID=" & CompanyID & ";"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    txtRate.Text = Format(r![Rate], "##0.00")
    cmdOK.Visible = False
    cmdUpdate.Visible = True
    cmdUpdate.Top = cmdOK.Top
  Else
    txtRate.Text = ""
    cmdUpdate.Visible = False
    cmdOK.Visible = True
  End If
  
End If

End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If cboMonthName.Text = "" Then
  MsgBox "Select Name of Month", vbInformation, cnstMsgInfo
  cboMonthName.SetFocus
  Exit Sub
End If
If txtRate.Text = "" Then
  MsgBox "Type Rate per US$", vbInformation, cnstMsgInfo
  txtRate.SetFocus
  Exit Sub
End If

Set MainComm = New ADODB.Command

MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans
DeleteRow = 0

MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Dolllar_Rate_Entry VALUES(" & CompanyID & ",'" & PeriodYear & "','" & cboMonthName.Text & "'," & txtRate.Text & "," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE());"
MainComm.Execute

MainConn.CommitTrans

txtRate.Text = ""
cboMonthName.SetFocus

End Sub
Private Sub cmdUpdate_Click()
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If cboMonthName.Text = "" Then
  MsgBox "Select Name of Month", vbInformation, cnstMsgInfo
  cboMonthName.SetFocus
  Exit Sub
End If
If txtRate.Text = "" Then
  MsgBox "Type Rate per US$", vbInformation, cnstMsgInfo
  txtRate.SetFocus
  Exit Sub
End If
  
If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
  Exit Sub
End If

Set MainComm = New ADODB.Command

MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500
'''''
DeleteRow = 2
MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Dolllar_Rate_Entry_DelEdit VALUES(" & CompanyID & ",'" & PeriodYear & "','" & cboMonthName.Text & "'," & txtRate.Text & "," & DeleteRow & " ," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE());"
MainComm.Execute
''''''
MainComm.CommandText = "SET NOCOUNT ON UPDATE Dolllar_Rate_Entry SET Rate=" & txtRate.Text & ",DeleteRow=" & DeleteRow & ",UserId=" & bytUserID & ",IDNo=" & getComVolID & ",EntryDate='" & Date & "',EntryTime='" & Time & "" _
      & " WHERE Month='" & cboMonthName.Text & "'AND Year='" & PeriodYear & "'AND ComID=" & CompanyID & ";"

MainComm.Execute

MsgBox "Successfully Updated", vbInformation, cnstMsgInfo

txtRate.Text = ""
cmdUpdate.Visible = False
cmdOK.Visible = True
cboMonthName.SetFocus

End Sub
Private Sub Form_Load()

MDIfrmMain.lblCaption.Caption = "Dollar Rate Entry"

prcMonthName cboMonthName
cboMonthName.Text = UCase(Format(Date, "MMMM"))

End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmDollarRateEntry = Nothing
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub

