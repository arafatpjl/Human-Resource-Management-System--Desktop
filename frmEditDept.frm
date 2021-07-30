VERSION 5.00
Begin VB.Form frmEditDept 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3965
   ClientLeft      =   6149
   ClientTop       =   1300
   ClientWidth     =   4537
   ControlBox      =   0   'False
   Icon            =   "frmEditDept.frx":0000
   LinkTopic       =   "frmledgercreate"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3965
   ScaleWidth      =   4537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox cboDept 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   3855
   End
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   415
      Left            =   1680
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Edit Department"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Type New Name"
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
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Dept Name What You want to Modify"
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
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
End
Attribute VB_Name = "frmEditDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Dim boDelete As Boolean
boDelete = True
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If Len(cboDept) = 0 Then
  MsgBox "Select Name of Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Delete Department", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

Set r = New ADODB.Recordset
StrRecord = "SELECT DeptName FROM Sub_Dept_Name WHERE ComID=" & CompanyID & " AND DeptName='" & cboDept.Text & "';"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  boDelete = False
End If
r.Close
  
StrRecord = "SELECT Department FROM Employee_Information WHERE ComID=" & CompanyID & " AND Department='" & cboDept.Text & "';"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  boDelete = False
End If
r.Close

If boDelete = True Then
  MsgBox "This record cannot be Deleted", vbInformation, cnstMsgInfo
  Exit Sub
End If

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500
DeleteRow = 1
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON Insert Into Dept_Name_DelEdit Select ComID,DeptName,Status," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() FROM Dept_Name WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute

MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Dept_Name WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
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
Set r = Nothing

MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

DeptName cboDept
cboDept.SetFocus

End Sub
Private Sub cmdSave_Click()
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
If Len(txtDeptName) = 0 Then
  MsgBox "Type Name of Department", vbInformation, cnstMsgInfo
  txtDeptName.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Modify Name of Department", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgInfo) = vbNo Then
  Exit Sub
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500
DeleteRow = 2
MainConn.BeginTrans

MainComm.CommandText = "SET NOCOUNT ON Insert Into Dept_Name_DelEdit Select ComID,DeptName,Status," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() from Dept_Name" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute

MainComm.CommandText = "SET NOCOUNT ON UPDATE Dept_Name SET DeptName='" & Trim(txtDeptName) & "',DeleteRow=" & DeleteRow & ",UserId=" & bytUserID & ",IDNo=" & getComVolID & ",EntryDate='" & Date & "',EntryTime=GETDATE()" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE Sub_Dept_Name SET DeptName='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE Employee_Information SET Department='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND Department='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE Salary_And_Wages SET DeptName='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE OverTime_Entry SET DeptName='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE Bonus_Entry SET DeptName='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND DeptName='" & cboDept.Text & "';"
MainComm.Execute
MainComm.CommandText = "SET NOCOUNT ON UPDATE Pay_Advice_Officer SET Dept='" & Trim(txtDeptName) & "'" _
           & " WHERE ComID=" & CompanyID & "AND Dept='" & cboDept.Text & "';"
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

txtDeptName.Text = ""
DeptName cboDept
cboDept.SetFocus

End Sub

Private Sub Command1_Click()
Unload Me
Set frmEditDept = Nothing
End Sub

Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Edit Department"
  DeptName cboDept
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MDIfrmMain.lblCaption.Caption = ""
  Set frmEditDept = Nothing
End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtDeptName_LostFocus()
If Len(txtDeptName) = 0 Then Exit Sub
txtDeptName.Text = Trim(UCase(txtDeptName.Text))
Set r = New ADODB.Recordset
StrRecord = "SELECT DeptName FROM Dept_Name WHERE ComID=" & CompanyID & " AND DeptName='" & txtDeptName.Text & "';"
r.Open StrRecord, MainConn, adOpenStatic
If r.EOF = False And r.BOF = False Then
  MsgBox "Duplicate Name of Department ", vbInformation, cnstMsgInfo
  txtDeptName.SetFocus
End If
r.Close
Set r = Nothing
End Sub
