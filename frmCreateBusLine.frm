VERSION 5.00
Begin VB.Form frmCreateBusLine 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3744
   ClientLeft      =   39
   ClientTop       =   -65
   ClientWidth     =   5291
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3744
   ScaleWidth      =   5291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      TabIndex        =   0
      Top             =   3000
      Width           =   885
   End
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
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
      Begin VB.ComboBox cboBusStopage 
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
         Height          =   315
         Left            =   2130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select Stopage Name"
         Top             =   900
         Width           =   1965
      End
      Begin VB.ComboBox cboBusLine 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2130
         Sorted          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Type Stopage Line"
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Stopage Name   :"
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
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Stopage Line  :"
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
         Left            =   540
         TabIndex        =   6
         Top             =   540
         Width           =   1395
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Create Bus Line"
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
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   5295
   End
End
Attribute VB_Name = "frmCreateBusLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBusLine_GotFocus()
    Call addBusStopLine(cboBusLine, cboBusStopage)
End Sub

Private Sub cboBusLine_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboBusLine_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub cboBusLine_LostFocus()
    cboBusLine.Text = Trim(UCase(cboBusLine.Text))
'    Call FindDuplicate
End Sub

Private Sub cboBusStopage_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Len(cboBusLine.Text) = 0 Then
        MsgBox "You should provide Bus Line.", vbCritical, App.Title
        cboBusLine.SetFocus
    Exit Sub
    End If
    
    ''--Find Duplicate Record..
    Set r = New ADODB.Recordset
    StrRecord = "SELECT BusLine,Busstation FROM Busline Where BusLine= '" & cboBusLine.Text & "' And Busstation='" & cboBusStopage.Text & "' ;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Bus Line...", vbInformation, cnstMsgInfo
        cboBusLine.SetFocus
        Exit Sub
    End If
    r.Close
    Set r = Nothing
    '''''''''''''''''''''''''''
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    MainConn.BeginTrans
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Else
    MainComm.CommandText = "Set Nocount On Insert Into Busline(Busline,Busstation,DeleteRow,UserID,IDNo,EntryDate,EntryTime) Values('" & cboBusLine.Text & "','" & cboBusStopage.Text & "',0," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE());"
    MainComm.Execute
    MainConn.CommitTrans
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Err.Number = 0
    Exit Sub
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Successfully Save..", vbOKOnly, App.Title
End Sub

Private Sub Form_Load()
'    Me.Height = 2370
    MDIfrmMain.lblCaption.Caption = "Create-Bus Line Information"
    Call addBusStop(cboBusStopage)
    Call addBusStopLine(cboBusLine, cboBusStopage)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIfrmMain.lblCaption.Caption = ""
    Set frmCreateBusLine = Nothing
End Sub


