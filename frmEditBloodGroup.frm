VERSION 5.00
Begin VB.Form frmEditBloodGroup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3840
   ClientLeft      =   6150
   ClientTop       =   1305
   ClientWidth     =   4170
   ControlBox      =   0   'False
   Icon            =   "frmEditBloodGroup.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4170
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
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
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
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   3495
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
      Left            =   240
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
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
      Height          =   415
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   " Edit Blood Group"
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
      Left            =   135
      TabIndex        =   7
      Top             =   240
      Width           =   2685
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4215
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
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Blood Group What You want to Modify"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3795
   End
End
Attribute VB_Name = "frmEditBloodGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()

If Len(cboDistrict) = 0 Then
  MsgBox "Select Name of Blood Group", vbInformation, cnstMsgInfo
  cboDistrict.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub


Set r = New ADODB.Recordset
StrRecord = "SELECT BloodGroup FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (BloodGroup='" & cboDistrict.Text & "'));"
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

MainComm.CommandText = "SET NOCOUNT ON DELETE FROM new_Blood_Group WHERE ((ComID=" & CompanyID & ") AND (BloodGroup='" & cboDistrict.Text & "'));"
MainComm.Execute

MainConn.CommitTrans

Screen.MousePointer = vbDefault

MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

addBloodGroup cboDistrict
cboDistrict.SetFocus

End Sub
Private Sub cmdSave_Click()

If Len(txtNewName) = 0 Then
  MsgBox "Type New Name of Blood Group", vbInformation, cnstMsgInfo
  txtNewName.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON UPDATE new_Blood_Group SET BloodGroup='" & txtNewName.Text & "'" _
           & " WHERE ((ComID=" & CompanyID & ") AND (BloodGroup='" & cboDistrict.Text & "'));"
MainComm.Execute

MainComm.CommandText = "SET NOCOUNT ON UPDATE Employee_Information SET BloodGroup='" & txtNewName.Text & "'" _
           & " WHERE ((ComID=" & CompanyID & ") AND (BloodGroup='" & cboDistrict.Text & "'));"
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
addBloodGroup cboDistrict
cboDistrict.SetFocus

End Sub
Private Sub Form_Load()
  addBloodGroup cboDistrict
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmEditBloodGroup = Nothing
End Sub

Private Sub txtNewName_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtNewName_LostFocus()
  txtNewName.Text = Trim(UCase(txtNewName.Text))
  If Len(txtNewName) = 0 Then Exit Sub

  Set r = New ADODB.Recordset
  StrRecord = "SELECT BloodGroup FROM new_Blood_Group WHERE ((ComID=" & CompanyID & ") AND (BloodGroup='" & txtNewName.Text & "'));"
  r.Open StrRecord, MainConn, adOpenStatic
  If r.EOF = False And r.BOF = False Then
    MsgBox "Duplicate name of Blood Group", vbInformation, cnstMsgInfo
    txtNewName.SetFocus
  End If
  r.Close
  Set r = Nothing
End Sub
