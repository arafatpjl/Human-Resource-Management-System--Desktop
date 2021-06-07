VERSION 5.00
Begin VB.Form frmChangePrincipal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   3195
   ClientTop       =   2265
   ClientWidth     =   5505
   ControlBox      =   0   'False
   Icon            =   "frmChangePrincipal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Salary Principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   300
      TabIndex        =   15
      Top             =   1140
      Width           =   4815
      Begin VB.CheckBox chkPF 
         Caption         =   "Manual"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtBasic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtHRent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtFoodPay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkBasic 
         Caption         =   "Manual"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkHRent 
         Caption         =   "Manual"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkFoodPay 
         Caption         =   "Manual"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   32
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "%   of Basic"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Pro. Fund"
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
         Index           =   0
         Left            =   480
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Basic"
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
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "%   of Gross"
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
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "House Rent"
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
         Index           =   8
         Left            =   480
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "%   of Basic"
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
         Left            =   2160
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Food Pay @"
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
         Index           =   11
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Tk   Per Day"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   12
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Height          =   255
         Index           =   13
         Left            =   3360
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Height          =   255
         Index           =   15
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Height          =   255
         Index           =   17
         Left            =   3360
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bonus Principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   300
      TabIndex        =   10
      Top             =   3060
      Width           =   4815
      Begin VB.CheckBox chkBonus 
         Caption         =   " Manual"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtBonus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   5
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   16
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Basic of Select Month"
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
         Index           =   18
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
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
      Left            =   1860
      TabIndex        =   0
      ToolTipText     =   "Click for Form Close"
      Top             =   4260
      Width           =   1575
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
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
      Left            =   3540
      TabIndex        =   11
      Top             =   4260
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   35
      Top             =   4080
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   " Change Principal"
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
      Left            =   135
      TabIndex        =   34
      Top             =   240
      Width           =   2745
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmChangePrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()

Dim bytBasicM As Byte, bytHRentM As Byte, bytFoodPayM As Byte
Dim bytBonusM As Byte, bytEmpPFM As Byte
Dim sngRateBasic As Single, sngRateHRent As Single, sngRateFPay As Single
Dim sngRateBonus As Single, sngRatePF As Single

If MsgBox("Are you sure to Update Salary Principal", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

If Len(txtBasic) = 0 Then sngRateBasic = 0 Else sngRateBasic = CSng(txtBasic.Text)
If chkBasic.Value = vbChecked Then bytBasicM = 1 Else bytBasicM = 0
If Len(txtHRent) = 0 Then sngRateHRent = 0 Else sngRateHRent = CSng(txtHRent.Text)
If chkHRent.Value = vbChecked Then bytHRentM = 1 Else bytHRentM = 0
If Len(txtFoodPay) = 0 Then sngRateFPay = 0 Else sngRateFPay = CSng(txtFoodPay.Text)
If chkFoodPay.Value = vbChecked Then bytFoodPayM = 1 Else bytFoodPayM = 0
If Len(txtPF) = 0 Then sngRatePF = 0 Else sngRatePF = CSng(txtPF.Text)
If chkPF.Value = vbChecked Then bytEmpPFM = 1 Else bytEmpPFM = 0
If Len(txtBonus) = 0 Then sngRateBonus = 0 Else sngRateBonus = CSng(txtBonus.Text)
If chkBonus.Value = vbChecked Then bytBonusM = 1 Else bytBonusM = 0

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON UPDATE Company_Information SET Basic=" & sngRateBasic & ",BasicM=" & bytBasicM & ",HRent=" & sngRateHRent & ",HRentM=" & bytHRentM & "," _
      & " FoodPay=" & sngRateFPay & ",FoodPayM=" & bytFoodPayM & ",PFund=" & sngRatePF & ",PFundM=" & bytEmpPFM & ",Bonus=" & sngRateBonus & ",BonusM=" & bytBonusM & " WHERE ID=" & CompanyID & ";"
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

Screen.MousePointer = vbDefault

MsgBox "Successfully Updated", vbInformation, cnstMsgInfo

txtBasic.Text = ""
chkBasic.Value = vbUnchecked
txtHRent.Text = ""
chkHRent.Value = vbUnchecked
txtFoodPay.Text = ""
chkFoodPay.Value = vbUnchecked
txtBonus.Text = ""
chkBonus.Value = vbUnchecked
txtPF.Text = ""
chkPF.Value = vbUnchecked

End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub Form_Load()
MDIfrmMain.lblCaption.Caption = "Change Salary Principal"

Set r = New ADODB.Recordset
StrRecord = "SELECT * FROM Company_Information WHERE ID=" & CompanyID & ";"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  If r![BASIC] > 0 Then txtBasic.Text = Format(r![BASIC], "##,##0.00") Else txtBasic.Text = "0.00"
  If r![BasicM] = True Then chkBasic.Value = vbChecked Else chkBasic.Value = cbunchecked

  If r![HRent] > 0 Then txtHRent.Text = Format(r![HRent], "##,##0.00") Else txtHRent.Text = "0.00"
  If r![HRentM] = True Then chkHRent.Value = vbChecked Else chkHRent.Value = cbunchecked

  If r![FoodPay] > 0 Then txtFoodPay.Text = Format(r![FoodPay], "##,##0.00") Else txtFoodPay.Text = "0.00"
  If r![FoodPayM] = True Then chkFoodPay.Value = vbChecked Else chkFoodPay.Value = cbunchecked

  If r![PFund] > 0 Then txtPF.Text = Format(r![PFund], "##,##0.00") Else txtPF.Text = "0.00"
  If r![PFundM] = True Then chkPF.Value = vbChecked Else chkPF.Value = cbunchecked

  If r![Bonus] > 0 Then txtBonus.Text = Format(r![Bonus], "##,##0.00") Else txtBonus.Text = "0.00"
  If r![BonusM] = True Then chkBonus.Value = vbChecked Else chkBonus.Value = cbunchecked
End If
r.Close
Set r = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmChangePrincipal = Nothing
End Sub
Private Sub txtBasic_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtBasic_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtBonus_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtBonus_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtFoodPay_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtFoodPay_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtHRent_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtHRent_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtPF_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtPF_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
