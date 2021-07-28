VERSION 5.00
Begin VB.Form frmEditBusStop 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6825
   ClientLeft      =   6150
   ClientTop       =   1305
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmEditBusStop.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Rate for Worker "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   150
      TabIndex        =   23
      Top             =   2460
      Width           =   3225
      Begin VB.TextBox txtMinW 
         Alignment       =   2  'Center
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
         Left            =   690
         TabIndex        =   2
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtMaxW 
         Alignment       =   2  'Center
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
         Left            =   2220
         TabIndex        =   3
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label10 
         Caption         =   "Min. : "
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
         Left            =   150
         TabIndex        =   25
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Max. : "
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
         Left            =   1650
         TabIndex        =   24
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rate for Staff "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   150
      TabIndex        =   20
      Top             =   3270
      Width           =   3225
      Begin VB.TextBox txtMaxS 
         Alignment       =   2  'Center
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
         Left            =   2220
         MaxLength       =   25
         TabIndex        =   5
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtMinS 
         Alignment       =   2  'Center
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
         Left            =   690
         MaxLength       =   25
         TabIndex        =   4
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Max. : "
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
         Left            =   1650
         TabIndex        =   22
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Min. : "
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
         Left            =   150
         TabIndex        =   21
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rate for JR Mgt. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   150
      TabIndex        =   17
      Top             =   4080
      Width           =   3225
      Begin VB.TextBox txtMaxJR 
         Alignment       =   2  'Center
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
         Left            =   2220
         MaxLength       =   25
         TabIndex        =   7
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtMinJR 
         Alignment       =   2  'Center
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
         Left            =   690
         MaxLength       =   25
         TabIndex        =   6
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "Max. : "
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
         Left            =   1650
         TabIndex        =   19
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label7 
         Caption         =   "Min. : "
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
         Left            =   150
         TabIndex        =   18
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Rate for SR Mgt. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   150
      TabIndex        =   14
      Top             =   4890
      Width           =   3225
      Begin VB.TextBox txtMaxSR 
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
         Left            =   2220
         MaxLength       =   25
         TabIndex        =   9
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtMinSR 
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
         Left            =   690
         MaxLength       =   25
         TabIndex        =   8
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "Max. : "
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
         Left            =   1650
         TabIndex        =   16
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label9 
         Caption         =   "Min. : "
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
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   465
      End
   End
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
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   6120
      Width           =   1005
   End
   Begin VB.ComboBox cboDistrict 
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
      Left            =   150
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtNewName 
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
      Left            =   150
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2040
      Width           =   3225
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "  Edit Bus Stoppage"
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
      Left            =   15
      TabIndex        =   27
      Top             =   240
      Width           =   2820
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Type New Name"
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
      Left            =   150
      TabIndex        =   13
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Stoppage"
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
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   1050
      Width           =   1485
   End
End
Attribute VB_Name = "frmEditBusStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDistrict_Click()
  If Len(cboDistrict) = 0 Then Exit Sub
  Set r = New ADODB.Recordset
  StrRecord = "SELECT * FROM new_Bus_Stop WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & cboDistrict.Text & "'));"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    txtNewName.Text = r![BusStop]
    txtMinW.Text = r![MinRateW]
    txtMaxW.Text = r![MaxRateW]
    txtMinS.Text = r![MinRateST]
    txtMaxS.Text = r![MaxRateST]
    txtMinJR.Text = r![MinRateJR]
    txtMaxJR.Text = r![MaxRateJR]
    txtMinSR.Text = r![MinRateSR]
    txtMaxSR.Text = r![MaxRateSR]
  End If
  r.Close
  Set r = Nothing
End Sub
Private Sub cmdDelete_Click()

If Len(cboDistrict) = 0 Then
  MsgBox "Select Name of Stoppage", vbInformation, cnstMsgInfo
  cboDistrict.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

Set r = New ADODB.Recordset
StrRecord = "SELECT BusStop FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & cboDistrict.Text & "'));"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  MsgBox "This record cannot be Deleted", vbCritical, cnstMsgInfo
  Exit Sub
End If
r.Close

Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON DELETE FROM new_Bus_Stop WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & cboDistrict.Text & "'));"
MainComm.Execute

MainConn.CommitTrans

Screen.MousePointer = vbDefault

MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

txtNewName.Text = ""
txtMinW.Text = "0"
txtMaxW.Text = "0"
txtMinS.Text = "0"
txtMaxS.Text = "0"
txtMinJR.Text = "0"
txtMaxJR.Text = "0"
txtMinSR.Text = "0"
txtMaxSR.Text = "0"

addBusStop cboDistrict
cboDistrict.SetFocus
End Sub
Private Sub cmdSave_Click()

If Len(txtNewName) = 0 Then
  MsgBox "Type New Name of Stoppage", vbInformation, cnstMsgInfo
  txtNewName.SetFocus
  Exit Sub
End If

If Len(txtMinW) = 0 Then txtMinW.Text = "0"
If Len(txtMaxW) = 0 Then txtMaxW.Text = "0"
If Len(txtMinS) = 0 Then txtMinS.Text = "0"
If Len(txtMaxS) = 0 Then txtMaxS.Text = "0"
If Len(txtMinJR) = 0 Then txtMinJR.Text = "0"
If Len(txtMaxJR) = 0 Then txtMaxJR.Text = "0"
If Len(txtMinSR) = 0 Then txtMinSR.Text = "0"
If Len(txtMaxSR) = 0 Then txtMaxSR.Text = "0"

If MsgBox("Are You sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON UPDATE new_Bus_Stop SET BusStop='" & txtNewName.Text & "',MinRateW=" & CSng(txtMinW.Text) & ",MaxRateW=" & CSng(txtMaxW.Text) & ",MinRateST=" & CSng(txtMinS.Text) & ",MaxRateST=" & CSng(txtMaxS.Text) & "," _
      & " MinRateJR=" & CSng(txtMinJR.Text) & ",MaxRateJR=" & CSng(txtMaxJR.Text) & ",MinRateSR=" & CSng(txtMinSR.Text) & ",MaxRateSR=" & CSng(txtMaxSR.Text) & " WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & cboDistrict.Text & "'));"
MainComm.Execute

MainComm.CommandText = "SET NOCOUNT ON UPDATE Employee_Information SET BusStop='" & txtNewName.Text & "' WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & cboDistrict.Text & "'));"
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

txtNewName.Text = ""
txtMinW.Text = "0"
txtMaxW.Text = "0"
txtMinS.Text = "0"
txtMaxS.Text = "0"
txtMinJR.Text = "0"
txtMaxJR.Text = "0"
txtMinSR.Text = "0"
txtMaxSR.Text = "0"

addBusStop cboDistrict
cboDistrict.SetFocus
End Sub
Private Sub Form_Load()
  addBusStop cboDistrict
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmEditBusStop = Nothing
End Sub

Private Sub txtMaxJR_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMaxS_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMaxSR_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMaxW_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMinJR_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMinS_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMinSR_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtMinW_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtNewName_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtNewName_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtNewName_LostFocus()
  txtNewName.Text = Trim(UCase(txtNewName.Text))
  If Len(txtNewName) = 0 Or UCase(cboDistrict.Text) = UCase(txtNewName.Text) Then Exit Sub

  Set r = New ADODB.Recordset
  StrRecord = "SELECT BusStop FROM new_Bus_Stop WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & txtNewName.Text & "'));"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    MsgBox "Duplicate name of Stoppage", vbInformation, cnstMsgInfo
    txtNewName.SetFocus
  End If
  r.Close
  Set r = Nothing
End Sub
Private Sub txtMaxJR_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMaxS_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMaxSR_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMaxW_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMinJR_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMinS_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMinSR_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtMinW_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub

