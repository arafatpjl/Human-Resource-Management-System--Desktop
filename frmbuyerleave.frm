VERSION 5.00
Begin VB.Form frmbuyerleave 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3495
   ClientLeft      =   3210
   ClientTop       =   2370
   ClientWidth     =   4845
   ControlBox      =   0   'False
   Icon            =   "frmbuyerleave.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   1980
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Year  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   495
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "&z"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2130
      TabIndex        =   4
      Top             =   4350
      Width           =   1155
   End
   Begin VB.ComboBox cmbMonth 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3840
      Width           =   2265
   End
   Begin VB.TextBox txtyear 
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
      Height          =   345
      Index           =   0
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Type Employee Code"
      Top             =   3510
      Width           =   2235
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3060
      TabIndex        =   1
      Top             =   2820
      Width           =   1545
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
      Height          =   450
      Left            =   720
      TabIndex        =   3
      Top             =   2820
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Report: Buyer Leave Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   225
      TabIndex        =   11
      Top             =   240
      Width           =   4485
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Year  : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   510
      TabIndex        =   6
      Top             =   3900
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Month : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   510
      TabIndex        =   5
      Top             =   3450
      Width           =   1245
   End
End
Attribute VB_Name = "frmbuyerleave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrint_Click()




          If Len(Text1.Text) = 0 Then
            MsgBox "Enter Year", vbCritical
            Text1.SetFocus
        Exit Sub
        End If


On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500
        
     
        MainConn.BeginTrans
        
'            Mysql = "DELETE FROM rptfor"
'            MainComm.CommandText = Mysql
'            MainComm.Execute
'
        
            Mysql = "Exec prcbuyerleave '" & Text1.Text & "'"
            MainComm.CommandText = Mysql
            MainComm.Execute
       
        
        
            Mysql = "UPDATE TempBuyerleave1 SET TempBuyerleave1.name=Company_Information.Name,TempBuyerleave1.Address=Company_Information.address From Company_Information Where Company_Information.id = " & CompanyID & " "
            
            
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            Mysql = "Update TempBuyerleave1 set year='" & Text1.Text & "'"
            MainComm.CommandText = Mysql
            MainComm.Execute
       
            
           
            
      MainConn.CommitTrans
Screen.MousePointer = vbDefault
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Err.Number = 0
    Exit Sub
    End If
 
'-----Dim Appl As New CRAXDRT.Application
'-----Dim Report As New CRAXDRT.Report

'With CrystalReport1
 Set Report = Appl.OpenReport(ReportPath & "BuyerLeave.rpt")
'            .ReportFileName = ReportPath & "BuyerLeave.rpt"
'            .PrintFileName = ReportPath & "BuyerLeave.rp"
'             ReportTitle = "Buyer Leave Summary"
'            .Destination = crptToWindow
'            .PrintReport
'End With
    Report.DiscardSavedData
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1

Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
 
    Call prcMakeCenter(Me)
    ' Call prcConReport(CrystalReport1)
   
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIfrmMain.lblCaption.Caption = ""
    Set frmbuyerleave = Nothing
End Sub

