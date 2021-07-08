VERSION 5.00
Begin VB.Form rptEmpFingerPrint 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   3210
   ClientTop       =   2370
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "rptEmpFingerPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   4695
      Begin VB.ComboBox cboDept 
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
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   250
         Width           =   3135
      End
      Begin VB.ComboBox cboSubDept 
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
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   700
         Width           =   3135
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
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         Top             =   250
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Department"
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
         Left            =   120
         TabIndex        =   14
         Top             =   250
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Sub Dept."
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
         Left            =   120
         TabIndex        =   13
         Top             =   700
         Width           =   975
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
         Left            =   1140
         TabIndex        =   12
         Top             =   700
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
      Begin VB.OptionButton OptBoth 
         Caption         =   "Both FingerPrint && Picture Remaining"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
      Begin VB.OptionButton OptFinger 
         Caption         =   "Only FingerPrint Remaining"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton OptPicture 
         Caption         =   "Only Picture Remaining"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdPrintPreview 
      BackColor       =   &H00C0C000&
      Caption         =   "Preview Report"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
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
      Height          =   390
      Left            =   330
      TabIndex        =   0
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Employee FingerPrint Report"
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
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4395
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "rptEmpFingerPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub cboDept_Click()
Call SubDeptName(cboSubDept, cboDept)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrintPreview_Click()
  If Len(cboDept) = 0 Then
   MsgBox "Please, Select Department", vbOKOnly
   cboDept.SetFocus
   Exit Sub
  End If
'   If Len(cboSubDept) = 0 Then
'   MsgBox "Please, Select SubDepartment", vbOKOnly
'   cboSubDept.SetFocus
'   Exit Sub
'   End If
   
  On Error GoTo ErrorHandler
'          Set MainComm = New ADODB.Command
'        Set r = New ADODB.Recordset
'        Set R1 = New ADODB.Recordset
'
'        MainComm.ActiveConnection = MainConn
'        MainComm.CommandTimeout = 500
'
'MainComm.CommandText = "EXEC FRPRINT " & CompanyID & "," & cboDept.Text & "," & cboSubDept.Text & "," & Frame1.Tag & ";"
'MainComm.Execute

Set Report = Appl.OpenReport(ReportPath & "RptFingerremainingemployee.rpt")
Report.DiscardSavedData
    frmMainReport.CRVIEWER.ReportSource = Report
'    Report.ParameterFields.GetItemByName("@COMID").AddDefaultValue CInt(CompanyID)
'    Report.ParameterFields.GetItemByName("@DEPT").AddDefaultValue cboDept.Text
'    Report.ParameterFields.GetItemByName("@SECTION").AddDefaultValue cboSubDept.Text
'    Report.ParameterFields.GetItemByName("@STATUS").AddDefaultValue CInt(Frame1.Tag)

    Report.ParameterFields.GetItemByName("@COMID").AddCurrentValue CInt(CompanyID)
    Report.ParameterFields.GetItemByName("@DEPT").AddCurrentValue cboDept.Text
    Report.ParameterFields.GetItemByName("@SECTION").AddCurrentValue cboSubDept.Text
    Report.ParameterFields.GetItemByName("@STATUS").AddCurrentValue CInt(Frame1.Tag)


    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Err.Number = 0
    Exit Sub
    End If
  End Sub

Private Sub Form_Load()
   OptBoth.Value = True
   Frame1.Tag = 1
   DeptName cboDept
End Sub

Private Sub OptBoth_Click()
Frame1.Tag = 1
End Sub

Private Sub OptFinger_Click()
  Frame1.Tag = 2
End Sub

Private Sub OptPicture_Click()
  Frame1.Tag = 3
End Sub
