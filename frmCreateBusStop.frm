VERSION 5.00
Begin VB.Form frmCreateBusStop 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6390
   ClientLeft      =   7020
   ClientTop       =   2385
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "frmCreateBusStop.frx":0000
   LinkTopic       =   "frmledgercreate"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   210
      TabIndex        =   21
      Top             =   4380
      Width           =   3225
      Begin VB.TextBox txtMinSR 
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
         TabIndex        =   7
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtMaxSR 
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
         TabIndex        =   8
         Top             =   330
         Width           =   825
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
         TabIndex        =   23
         Top             =   330
         Width           =   465
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
         TabIndex        =   22
         Top             =   330
         Width           =   525
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
      Left            =   210
      TabIndex        =   18
      Top             =   3570
      Width           =   3225
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
         TabIndex        =   5
         Top             =   330
         Width           =   825
      End
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
         TabIndex        =   6
         Top             =   330
         Width           =   825
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
         TabIndex        =   20
         Top             =   330
         Width           =   465
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
      Left            =   210
      TabIndex        =   15
      Top             =   2760
      Width           =   3225
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
         TabIndex        =   3
         Top             =   330
         Width           =   825
      End
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
         TabIndex        =   4
         Top             =   330
         Width           =   825
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
         TabIndex        =   17
         Top             =   330
         Width           =   465
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
         TabIndex        =   16
         Top             =   330
         Width           =   525
      End
   End
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
      Left            =   210
      TabIndex        =   12
      Top             =   1950
      Width           =   3225
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
         TabIndex        =   2
         Top             =   330
         Width           =   825
      End
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
         TabIndex        =   1
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label3 
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
         TabIndex        =   14
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label2 
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
         TabIndex        =   13
         Top             =   330
         Width           =   465
      End
   End
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
      Left            =   210
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1500
      Width           =   3225
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
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   5640
      Width           =   1125
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
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   26
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Bus Stoppage"
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
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   3045
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Bus Stoppage"
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
      Left            =   210
      TabIndex        =   11
      Top             =   1230
      Width           =   1995
   End
End
Attribute VB_Name = "frmCreateBusStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    If Len(txtDeptName) = 0 Then
        MsgBox "Type Name of Stoppage", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
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

    Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500

        MainConn.BeginTrans
            MainComm.CommandText = "SET NOCOUNT ON INSERT INTO new_Bus_Stop VALUES(" & CompanyID & ",'" & txtDeptName.Text & "'," & CSng(txtMinW.Text) & "," & CSng(txtMaxW.Text) & "," & CSng(txtMinS.Text) & "," & CSng(txtMaxS.Text) & "," & CSng(txtMinJR.Text) & "," & CSng(txtMaxJR.Text) & "," & CSng(txtMinSR.Text) & "," & CSng(txtMaxSR.Text) & ");"
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

    Call addBusStop(frmEmpInformation.cboBusStop)
    frmEmpInformation.cboBusStop.Text = txtDeptName.Text

    Unload Me

    frmEmpInformation.cmdSave.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateBusStop = Nothing
End Sub

Private Sub txtDeptName_GotFocus()
    Call fncGotFocus(txtDeptName)
End Sub

Private Sub txtDeptName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtDeptName_LostFocus()
    txtDeptName.Text = Trim(UCase(txtDeptName.Text))
    If Len(txtDeptName) = 0 Then Exit Sub

    Set r = New ADODB.Recordset
    StrRecord = "SELECT BusStop FROM new_Bus_Stop WHERE ((ComID=" & CompanyID & ") AND (BusStop='" & txtDeptName.Text & "'));"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Name of Stoppage", vbInformation, cnstMsgInfo
        txtDeptName.SetFocus
    End If
    r.Close
    Set r = Nothing
End Sub

Private Sub txtMaxJR_GotFocus()
    Call fncGotFocus(txtMaxJR)
End Sub

Private Sub txtMaxJR_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMaxS_GotFocus()
    Call fncGotFocus(txtMaxS)
End Sub

Private Sub txtMaxS_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMaxSR_GotFocus()
    Call fncGotFocus(txtMaxSR)
End Sub

Private Sub txtMaxSR_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMaxW_GotFocus()
    Call fncGotFocus(txtMaxW)
End Sub

Private Sub txtMaxW_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMinJR_GotFocus()
    Call fncGotFocus(txtMinJR)
End Sub

Private Sub txtMinJR_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMinS_GotFocus()
    Call fncGotFocus(txtMinS)
End Sub

Private Sub txtMinS_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMinSR_GotFocus()
    Call fncGotFocus(txtMinSR)
End Sub

Private Sub txtMinSR_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMinW_GotFocus()
    Call fncGotFocus(txtMinW)
End Sub

Private Sub txtMinW_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub
