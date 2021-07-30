VERSION 5.00
Begin VB.Form frmEditDesignation 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4560
   ClientLeft      =   6150
   ClientTop       =   1305
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmEditDesignation.frx":0000
   LinkTopic       =   "frmledgercreate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCategory 
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
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3060
      Width           =   3855
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
      Height          =   415
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cboDesignation 
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
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1500
      Width           =   3855
   End
   Begin VB.TextBox txtNewName 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Edit Designation"
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
      TabIndex        =   10
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   45
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   4905
   End
   Begin VB.Label Label3 
      Caption         =   "Select Category :"
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
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2790
      Width           =   1545
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
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Designation What You want to Modify"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1140
      Width           =   3735
   End
End
Attribute VB_Name = "frmEditDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboCategory_LostFocus()
    If Len(cboCategory.Text) > 0 Then cboCategory.Text = UCase(Trim(cboCategory.Text))
End Sub

Private Sub cboDesignation_Click()
    txtNewName.Text = cboDesignation.Text
    cboCategory.ListIndex = -1
    
    
    Set r = New ADODB.Recordset

    StrRecord = "SELECT Category FROM Designation WHERE ComID=" & CompanyID & " AND designation='" & cboDesignation.Text & "' and Category is not null;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        cboCategory.Text = r("Category").Value
    Exit Sub
    End If
    r.Close
    Set r = Nothing
End Sub

Private Sub cmdDelete_Click()
    If cboDesignation.Text = "" Then
        MsgBox "Select Name of designation", vbInformation, cnstMsgInfo
        cboDesignation.SetFocus
    Exit Sub
    End If
    
    If MsgBox("Are You sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
        Exit Sub
    End If
    
    Set r = New ADODB.Recordset

    StrRecord = "SELECT Designation FROM Employee_Information WHERE ComID=" & CompanyID & " AND designation='" & cboDesignation.Text & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "This record cannot be Deleted", vbInformation, cnstMsgInfo
        r.Close
    Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Designation WHERE ComID=" & CompanyID & " AND Designation='" & cboDesignation.Text & "';"
        MainComm.Execute
    MainConn.CommitTrans

    Screen.MousePointer = vbDefault
    MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo

    Call Designation(cboDesignation)
    Call prcDesigCategory(cboCategory)
    cboDesignation.SetFocus
End Sub

Private Sub cmdSave_Click()
    If Len(txtNewName.Text) = 0 Then
        MsgBox "Type New Name of designation", vbInformation, cnstMsgInfo
        txtNewName.SetFocus
    Exit Sub
    End If
    
    If Len(cboCategory.Text) = 0 Then
        MsgBox "Select designation category.", vbInformation, cnstMsgInfo
        cboCategory.SetFocus
    Exit Sub
    End If
    
    If MsgBox("Are You sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Designation SET Designation='" & Trim(txtNewName.Text) & "', Category = '" & Trim(cboCategory.Text) & "'" _
            & " WHERE ComID=" & CompanyID & "AND Designation='" & cboDesignation.Text & "';"
        MainComm.Execute
    
        MainComm.CommandText = "SET NOCOUNT ON UPDATE Employee_Information SET Designation='" & Trim(txtNewName) & "'" _
            & " WHERE ComID=" & CompanyID & "AND Designation='" & cboDesignation.Text & "';"
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
    Call Designation(cboDesignation)
    Call prcDesigCategory(cboCategory)
    cboDesignation.SetFocus
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    Call Designation(cboDesignation)
    Call prcDesigCategory(cboCategory)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditDesignation = Nothing
End Sub

Private Sub txtNewName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtNewName_LostFocus()
    Screen.MousePointer = vbHourglass

    If txtNewName.Text <> "" Then
        txtNewName.Text = UCase(txtNewName.Text)
    
        Set r = New ADODB.Recordset

        StrRecord = "SELECT Designation FROM Designation WHERE ComID=" & CompanyID & " AND Designation='" & Trim(txtNewName) & "';"
        r.Open StrRecord, MainConn, adOpenStatic
  
        If r.EOF = False And r.BOF = False Then
            MsgBox "Duplicate Name", vbInformation, cnstMsgInfo
            txtNewName.SetFocus
        End If
    End If

    Screen.MousePointer = vbDefault
End Sub
