VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDinnerEntry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8749
   ClientLeft      =   39
   ClientTop       =   39
   ClientWidth     =   15002
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8749
   ScaleWidth      =   15002
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCategory 
      Caption         =   " Employee Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4560
      TabIndex        =   35
      Tag             =   "W"
      Top             =   970
      Width           =   5055
      Begin VB.OptionButton optStaff 
         Caption         =   "Staff"
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
         Height          =   255
         Left            =   1770
         TabIndex        =   39
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optSrMgt 
         Caption         =   "SR. Mgt."
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
         Height          =   255
         Left            =   3930
         TabIndex        =   38
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optJrMgt 
         Caption         =   "JR. Mgt."
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
         Height          =   255
         Left            =   2790
         TabIndex        =   37
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optWorker 
         Caption         =   "Workers"
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
         Height          =   255
         Left            =   540
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6840
      TabIndex        =   33
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Addition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   6975
      Begin VB.CheckBox Chksection 
         Caption         =   "Section off "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbdeptadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   2835
      End
      Begin VB.ComboBox cmbsubdeptadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   670
         Width           =   2835
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Delete Previous Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkEmpCode 
         Caption         =   "Emp Code Wise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddList 
         Caption         =   "ADD LIST"
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   1560
         Width           =   1110
      End
      Begin VB.TextBox txtDateAdd 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   1550
         Width           =   2775
      End
      Begin VB.TextBox txtEmpCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
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
         Left            =   1800
         TabIndex        =   15
         Top             =   1100
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Department "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   670
         Width           =   1305
      End
      Begin VB.Label Label6 
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
         Index           =   6
         Left            =   1440
         TabIndex        =   26
         Top             =   670
         Width           =   255
      End
      Begin VB.Label Label6 
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
         Index           =   5
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
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
         Index           =   3
         Left            =   1440
         TabIndex        =   22
         Top             =   1550
         Width           =   255
      End
      Begin VB.Label Label6 
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
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   1100
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1550
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1100
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deletion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7200
      TabIndex        =   2
      Top             =   1560
      Width           =   7695
      Begin VB.CommandButton Cmdprint 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbsubdeptname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2625
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2835
      End
      Begin VB.ComboBox cmbdeptname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2835
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2610
         TabIndex        =   4
         Top             =   960
         Width           =   2835
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
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
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label6 
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
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   1005
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Department "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
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
         Left            =   840
         TabIndex        =   7
         Top             =   1005
         Width           =   525
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   7200
      Width           =   765
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3555
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   14625
      _ExtentX        =   25807
      _ExtentY        =   6278
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   325
      Left            =   2041
      Top             =   7202
      Visible         =   0   'False
      Width           =   3042
      _ExtentX        =   5368
      _ExtentY        =   575
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmDinnerEntry.frx":0000
      OLEDBString     =   $"frmDinnerEntry.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Dinner Allowance Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   7800
      Width           =   15375
   End
End
Attribute VB_Name = "frmDinnerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dtmonth As String
Dim dtMth As Integer
Dim dtyear As Integer
Private Sub chkEmpCode_Click()
If chkEmpCode.Value = 1 Then
    txtEmpCode.Enabled = True
    txtEmpCode.Text = ""
    txtEmpCode.SetFocus
    Else
        txtEmpCode.Enabled = False
End If
End Sub

Private Sub cmbdeptadd_Click()
Call SubDeptName(cmbsubdeptadd, cmbdeptadd)
End Sub

Private Sub cmbdeptname_Click()
    Call SubDeptName(cmbsubdeptname, cmbdeptname)
End Sub

Private Sub cmbdeptname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbsubdeptname_Click()
Text1.Text = ""
'Set r = New ADODB.Recordset
'    If Len(cmbsubdeptname.Text) > 0 Then
'        If Adodc1.Recordset.State = adStateOpen Then: Adodc1.Recordset.CancelBatch: Adodc1.Recordset.Close
''        If boJRMgt = False Then
'            Mysql = "Select EmployeeCode, Name,Department,Section,Designation,DailyDate,TimeIn,TimeOut,Hours,Amount,Sign From EmpDinner Where Comid = " & CompanyID & " and Department='" & cmbdeptname.Text & "' and Section='" & cmbsubdeptname.Text & "'  Order By Employeecode"
'            Adodc1.Recordset.Open Mysql
'            r.Open Mysql, MainConn, adOpenStatic
'            If r.RecordCount > 0 Then
'            Text1.Text = r.RecordCount
'            End If
'
'        Set DataGrid1.DataSource = Adodc1
'
'        Call prcGridSetting
'
'
'    End If
'r.Close
'Set r = Nothing
End Sub

Private Sub cmbsubdeptname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        DataGrid1.SetFocus
    End If
End Sub

Private Sub cmdAddList_Click()


If Len(cmbdeptadd) = 0 Then
        MsgBox "Select Department", vbInformation, cnstMsgInfo
        cboDept.SetFocus
        Exit Sub
    End If
    If Len(cmbsubdeptadd) = 0 Then
        MsgBox "Select Section", vbInformation, cnstMsgInfo
        cboSubDept.SetFocus
        Exit Sub
    End If


If chkEmpCode.Value = 1 Then
    If Len(txtEmpCode.Text) = 0 Then
        MsgBox "Enter the Employee Code", vbOKCancel, cnstMsgInfo
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If

    If Len(txtDateAdd.Text) = 0 Then
        MsgBox "Enter the Date", vbOKCancel, cnstMsgInfo
        txtDateAdd.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call Validity(txtDateAdd, True)
    If VAREXIT = 1 Then Exit Sub
    
    dtMth = DatePart("m", txtDateAdd.Text)
    dtmonth = MonthName(dtMth)
    dtyear = Year(txtDateAdd.Text)
    Set r = New ADODB.Recordset
    StrRecord = "Select * From Salary_And_Wages Where ComId=" & CompanyID & " And Month='" & dtmonth & "' And Year=" & dtyear & " AND (EntryStatus='A');"
    r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then
            MsgBox "Attendance already adjusted with Salary", vbCritical, cnstMsgInfo
           Exit Sub
        End If
    r.Close
    
    On Error GoTo ErrorHand
    Screen.MousePointer = vbHourglass
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
    
    
If Chksection.Value = 0 Then

    
    
    If chkEmpCode.Value = 1 And chkDelete.Value = 0 Then
    
        Set r = New ADODB.Recordset
            StrRecord = "Select * From EmpDinner Where ComId=" & CompanyID & "  And  DailyDate='" & txtDateAdd.Text & "' And EmployeeCode='" & txtEmpCode.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
            r.Open StrRecord, MainConn, adOpenStatic
            If r.EOF = False And r.BOF = False Then
                MsgBox "Same Record Already inserted", vbCritical, cnstMsgInfo
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            r.Close
            MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5) and  EmployeeCode='" & txtEmpCode.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "' ;"
            MainComm.Execute
    ElseIf chkEmpCode.Value = 0 And chkDelete.Value = 0 Then
    
            Set r = New ADODB.Recordset
            StrRecord = "Select * From EmpDinner Where ComId=" & CompanyID & "    And  DailyDate='" & txtDateAdd.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "' ;"
            r.Open StrRecord, MainConn, adOpenStatic
            If r.EOF = False And r.BOF = False Then
                MsgBox "Same Record Already inserted", vbCritical, cnstMsgInfo
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            r.Close
            MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5)  AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
     MainComm.Execute
    End If
   

If chkDelete.Value = 1 And chkEmpCode.Value = 0 Then
        
    Set r = New ADODB.Recordset
    StrRecord = "Select * From VIEWDINNER Where ComId=" & CompanyID & "  And  DailyDate='" & txtDateAdd.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.RecordCount = 0 Then
        MsgBox "Record Not Found", vbCritical, cnstMsgInfo
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    r.Close
    If MsgBox("Are you sure, you want to delete the records ", vbYesNo, cnstMsgInfo) = vbNo Then Exit Sub
        'MainConn.BeginTrans
        MainComm.CommandText = "Delete From EmpDinner Where ComId=" & CompanyID & " And DailyDate='" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
        MainComm.Execute
        
        MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5)  AND DEPARTMENT='" & cmbdeptadd.Text & "' AND SECTION='" & cmbsubdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
         MainComm.Execute
         
         
txtDateAdd.Text = ""
End If
''''''''''''''''''
ElseIf Chksection.Value = 1 Then
 
    If chkEmpCode.Value = 1 And chkDelete.Value = 0 Then
    
        Set r = New ADODB.Recordset
            StrRecord = "Select * From EmpDinner Where ComId=" & CompanyID & "  And  DailyDate='" & txtDateAdd.Text & "' And EmployeeCode='" & txtEmpCode.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
            r.Open StrRecord, MainConn, adOpenStatic
            If r.EOF = False And r.BOF = False Then
                MsgBox "Same Record Already inserted", vbCritical, cnstMsgInfo
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            r.Close
            MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5) and  EmployeeCode='" & txtEmpCode.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "' ;"
            MainComm.Execute
    ElseIf chkEmpCode.Value = 0 And chkDelete.Value = 0 Then
    
            Set r = New ADODB.Recordset
            StrRecord = "Select * From EmpDinner Where ComId=" & CompanyID & "    And  DailyDate='" & txtDateAdd.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "' ;"
            r.Open StrRecord, MainConn, adOpenStatic
            If r.EOF = False And r.BOF = False Then
                MsgBox "Same Record Already inserted", vbCritical, cnstMsgInfo
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            r.Close
            MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5)  AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
     MainComm.Execute
    End If
   

If chkDelete.Value = 1 And chkEmpCode.Value = 0 Then
        
    Set r = New ADODB.Recordset
    StrRecord = "Select * From VIEWDINNER Where ComId=" & CompanyID & "  And  DailyDate='" & txtDateAdd.Text & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.RecordCount = 0 Then
        MsgBox "Record Not Found", vbCritical, cnstMsgInfo
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    r.Close
    If MsgBox("Are you sure, you want to delete the records ", vbYesNo, cnstMsgInfo) = vbNo Then Exit Sub
        'MainConn.BeginTrans
        MainComm.CommandText = "Delete From EmpDinner Where ComId=" & CompanyID & " And DailyDate='" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "' AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
        MainComm.Execute
        
        MainComm.CommandText = "Insert Into EmpDinner(Comid,EmployeeCode,Name,Department,Section,Designation,Dailydate,Timein,Timeout,Hours,Wstatus,Amount) " _
            & " SELECT  ComID, EmployeeCode, Name, Department, [Section], Designation, DailyDate, TimeIn, TimeOut, " _
            & " CAST(DATEDIFF(n, TimeIn, TimeOut) AS Float)/ 60, WStatus, " _
            & " (CASE WHEN WSTATUS='W'THEN 25 WHEN WSTATUS='F'THEN 25 WHEN WSTATUS='J'THEN 45 WHEN WSTATUS='S'THEN 60 END)  FROM VIEWDINNER " _
            & " Where ComId=" & CompanyID & "  And DailyDate = '" & Format(txtDateAdd.Text, cnstDtFrmtQ) & "'  " _
            & " AND NOT ((TimeIn <> '1900-01-01' AND Timeout = '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout <> '1900-01-01') OR (TimeIn = '1900-01-01' AND Timeout = '1900-01-01')) " _
            & " AND (TimeOut BETWEEN '1900-01-01 22:00:00 PM' AND '1900-01-01 23:59:59 PM') " _
            & " AND (CAST(DATEDIFF(n, TimeIn,TimeOut) AS Float) / 60 >= 14.5)  AND DEPARTMENT='" & cmbdeptadd.Text & "' AND WSTATUS='" & FrameCategory.Tag & "';"
         MainComm.Execute
         
         
txtDateAdd.Text = ""
End If

End If


MainConn.CommitTrans
MsgBox "Data Successfully Inserted", vbOKOnly, cnstMsgInfo
ErrorHand:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Exit Sub
     End If
     
    MsgBox "Successfully Updated", vbInformation, cnstMsgInfo
    Screen.MousePointer = vbDefault
   Set r = Nothing
End Sub

Private Sub cmdClose_Click()
    MDIfrmMain.lblCaption.Caption = " "
    
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans

'---------
'With CrystalReport1
'-----Dim Appl As New CRAXDRT.Application
'-----Dim Report As New CRAXDRT.Report
    
    Set Report = Appl.OpenReport(ReportPath & "RptDinnerBillDay.rpt")

' .ReportFileName = ReportPath + "RptDinnerBillDay.rpt"
' .PrintFileName = ReportPath + "RptDinnerBillDay.rpt"




   If cmbsubdeptname.Text = "" Then

   Mysql = " Select * from empDinner where  department= '" & cmbdeptname.Text & "' AND  comid=" & CompanyID & "  AND (DAILYDATE = '" & txtDate.Text & "') AND (SIGN > 0) and (Wstatus = '" & FrameCategory.Tag & "') ORDER BY DEPARTMENT,SECTION,EMPLOYEECODE ;"


   Else

   Mysql = " Select * from empDinner where  department= '" & cmbdeptname.Text & "' and section= '" & cmbsubdeptname.Text & "'  AND  comid=" & CompanyID & " and  (DAILYDATE = '" & txtDate.Text & "') AND (SIGN> 0) and (Wstatus = '" & FrameCategory.Tag & "') ORDER BY DEPARTMENT,SECTION,EMPLOYEECODE ;"


   End If


    MainComm.CommandText = "dELETE  FROM Temp_empDinner_2 "
    MainComm.Execute
    


    MainComm.CommandText = "INSERT INTO  Temp_empDinner_2 " & Mysql & ""
    MainComm.Execute


MainConn.CommitTrans

    Report.DiscardSavedData
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1


'       .Destination = crptToWindow
'    .PrintReport

'    End With

'CrystalReport1.ReportFileName = ReportPath + "RptDinnerBillDay.rpt"
'CrystalReport1.PrintFileName = ReportPath + "RptDinnerBillDay.rpt"
'CrystalReport1.PrintReport




ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  Err.Number = 0
  Exit Sub
End If
'CrystalReport1.ReportFileName = ReportPath + "RptDinnerBillDay.rpt"
'CrystalReport1.PrintFileName = ReportPath + "RptDinnerBillDay.rpt"
'CrystalReport1.PrintReport
'

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdate_Click()


        On Error GoTo ErrorHandler
        Screen.MousePointer = vbHourglass
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
                    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
                Exit Sub
        End If
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500
        DeleteRow = 2
            MainConn.BeginTrans
            
            MainComm.CommandText = "SET NOCOUNT ON Insert Into EmpDinner_DelEdit Select ComID,EmployeeCode,Name,Department,[Section],Designation,DailyDate,TimeIn,TimeOut,Hours,Amount,WStatus,Sign," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() from EmpDinner Where Comid = " & CompanyID & " and Department='" & cmbdeptname.Text & "' and Section='" & cmbsubdeptname.Text & "' And DailyDate='" & Format(txtDate.Text, cnstDtFrmtQ) & "' and wstatus='" & FrameCategory.Tag & "' ; "
            MainComm.Execute
            
            MainComm.CommandText = "SET NOCOUNT ON Update EmpDinner Set Sign=2 Where Comid = " & CompanyID & " and Department='" & cmbdeptname.Text & "' and Section='" & cmbsubdeptname.Text & "' And DailyDate='" & Format(txtDate.Text, cnstDtFrmtQ) & "'  and wstatus='" & FrameCategory.Tag & "' "
            MainComm.Execute
            
            MainConn.CommitTrans
        




ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
        Screen.MousePointer = vbDefault
        Err.Number = 0
    Exit Sub
    End If
    
    Set r = Nothing
    Screen.MousePointer = vbDefault

           
    If Len(txtDate) = 0 Then Exit Sub
    txtDate.Text = FormatDate(txtDate)
    If Len(cmbsubdeptname.Text) > 0 Then
        If Adodc1.Recordset.State = adStateOpen Then: Adodc1.Recordset.CancelBatch: Adodc1.Recordset.Close
'        If boJRMgt = False Then
            Mysql = "Select EmployeeCode, Name,Department,Section,Designation,DailyDate,TimeIn,TimeOut,Hours,Amount,WStatus,Sign From EmpDinner Where Comid = " & CompanyID & " and Department='" & cmbdeptname.Text & "' and Section='" & cmbsubdeptname.Text & "' And DailyDate='" & Format(txtDate.Text, cnstDtFrmtQ) & "'  and wstatus='" & FrameCategory.Tag & "'  Order By Employeecode"
'        End If
        Adodc1.Recordset.Open Mysql
        Set DataGrid1.DataSource = Adodc1
        
        Call prcGridSetting
    End If



End Sub


Private Sub DataGrid1_GotFocus()
    DataGrid1.col = 4
    DataGrid1.SetFocus

End Sub

Private Sub DataGrid1_LostFocus()
    Text1.Text = DataGrid1.VisibleRows
End Sub

Private Sub Form_Load()
      ' Call prcConReport(CrystalReport1)
    MDIfrmMain.lblCaption.Caption = "Dinner Bill Entry"
    Move (Screen.width - width) / 2, (Screen.Height - Height) / 2
    'Top = (frmDinnerEntry.Height - Height) / 2 + 900
    'Left = (frmDinnerEntry.width - width) / 2
'    Set r = New ADODB.Recordset
'    StrRecord = "SELECT Editstatus FROM Sys_User_Name WHERE Userid='" & bytUserID & "';"
'    r.Open StrRecord, MainConn, adOpenStatic
''    If r![Editstatus] = 1 Then
'    cmdUpdate.Enabled = True
'    Else
'    cmdUpdate.Enabled = False
'    End If
    
    
     Set r = New ADODB.Recordset
    StrRecord = "SELECT Worker,JRmgt,SRmgt,Staff FROM Sys_User_Name WHERE UserId='" & bytUserID & "' ;"
    r.Open StrRecord, MainConn, adOpenStatic
    If r![Worker] = 1 Then
    optWorker.Enabled = True
    Else
    optWorker.Enabled = False
    End If
    If r![JRmgt] = 1 Then
    optJrMgt.Enabled = True
    Else
    optJrMgt.Enabled = False
    End If
    If r![SRmgt] = 1 Then
    optSrMgt.Enabled = True
    Else
    optSrMgt.Enabled = False
    End If
    If r![Staff] = 1 Then
    optStaff.Enabled = True
    Else
    optStaff.Enabled = False
    End If
   r.Close

    Mysql = "Select 1,EmployeeCode, Name,Department,Section,Designation,DailyDate,TimeIn,TimeOut,Hours,Amount,WStatus,Sign From EmpDinner Where Comid = 0"
    Call prcAdodc(Adodc1, Mysql)
    Set DataGrid1.DataSource = Adodc1
    Call prcGridSetting
    Call DeptName(cmbdeptname)
    Call DeptName(cmbdeptadd)
    cmbsubdeptname.Clear
    cmbsubdeptadd.Clear
End Sub

Public Sub prcGridSetting()
    With DataGrid1
        .Columns(0).width = 1000
        .Columns(1).width = 1000
        .Columns(2).width = 1600
        .Columns(3).width = 1000
        .Columns(4).width = 1200
        .Columns(5).width = 1200
        .Columns(6).width = 1800
        .Columns(7).width = 1800
        .Columns(8).width = 500
        .Columns(9).width = 600
        .Columns(10).width = 600
        .Columns(11).width = 500
        
'        .Columns(0).Caption = "SL.NO"
        .Columns(0).Caption = "Emp Code"
        .Columns(1).Caption = "Name"
        .Columns(2).Caption = "Department"
        .Columns(3).Caption = "Section"
        .Columns(4).Caption = "Designation"
        .Columns(5).Caption = "DailyDate"
        .Columns(6).Caption = "TimeIn"
        .Columns(7).Caption = "TimeOut"
        .Columns(8).Caption = "Hours"
        .Columns(9).Caption = "Amount"
        .Columns(10).Caption = "Status"
        .Columns(11).Caption = "Sign"
        
        
        .Columns(0).Locked = True
        .Columns(1).Locked = True
        .Columns(2).Locked = True
        .Columns(3).Locked = True
        .Columns(4).Locked = True
        .Columns(5).Locked = True
        .Columns(6).Locked = True
        .Columns(7).Locked = True
        .Columns(8).Locked = True
        .Columns(9).Locked = True
        .Columns(10).Locked = True
        .Columns(11).Locked = True
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDinnerEntry = Nothing
End Sub

Private Sub optJrMgt_Click()
FrameCategory.Tag = "J"
End Sub

Private Sub optSrMgt_Click()
FrameCategory.Tag = "S"
End Sub

Private Sub optStaff_Click()
FrameCategory.Tag = "F"
End Sub

Private Sub optWorker_Click()

FrameCategory.Tag = "W"

End Sub

Private Sub txtDate_GotFocus()
'    SendKeys "{Home}+{End}"
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtDate_LostFocus()
Text1.Text = ""

If Len(txtDate) = 0 Then Exit Sub
   txtDate.Text = FormatDate(txtDate)
'Dim dtmonth As String
'
'
'Dim dtyear As Integer

dtmonth = Month(txtDate.Text)
dtyear = Year(txtDate.Text)

    Set r = New ADODB.Recordset
    StrRecord = "SELECT MONTH FROM Month_Name WHERE SLNO = " & dtmonth & " ;"

    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        Do While Not r.EOF
           dtmonth = r![Month]
        r.MoveNext
        Loop
    End If
    r.Close
    Set r = Nothing



Text1.Text = ""

'Set r = New ADODB.Recordset
'
'Screen.MousePointer = vbHourglass
'        StrRecord = "SELECT DeptName,SubDeptName FROM Salary_And_Wages WHERE (([Month]='" & dtmonth & "') AND ([Year]=" & dtyear & ")" _
'            & " AND (ComID=" & CompanyID & ")) GROUP BY DeptName,SubDeptName HAVING ((DeptName='" & cmbdeptname.Text & "') AND (SubDeptName='" & cmbsubdeptname.Text & "'));"
'        r.Open StrRecord, MainConn, adOpenStatic
'        If r.EOF = False And r.BOF = False Then
'            MsgBox "Attendance already adjusted with Salary", vbCritical, cnstMsgInfo
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        r.Close
        Screen.MousePointer = vbDefault



Set r = New ADODB.Recordset
    If Len(txtDate) = 0 Then Exit Sub
   'txtDate.Text = FormatDate(txtDate)
    If Len(cmbsubdeptname.Text) > 0 Then
        If Adodc1.Recordset.State = adStateOpen Then: Adodc1.Recordset.CancelBatch: Adodc1.Recordset.Close
'        If boJRMgt = False Then
            Mysql = "Select EmployeeCode, Name,Department,Section,Designation,DailyDate,TimeIn,TimeOut,Hours,Amount,WStatus,Sign From EmpDinner Where Comid = " & CompanyID & " and Department='" & cmbdeptname.Text & "' and Section='" & cmbsubdeptname.Text & "' And DailyDate='" & Format(txtDate.Text, cnstDtFrmtQ) & "'  AND WSTATUS='" & FrameCategory.Tag & "' Order By Employeecode"
'        End If
        Adodc1.Recordset.Open Mysql
        r.Open Mysql, MainConn, adOpenStatic
        If r.RecordCount > 0 Then
            Text1.Text = r.RecordCount
        End If
         r.Close
        Set DataGrid1.DataSource = Adodc1

        Call prcGridSetting
        End If
   
    Set r = Nothing
End Sub

Private Sub txtDateAdd_GotFocus()
'    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDateAdd_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtDateAdd_LostFocus()
    
    If Len(txtDateAdd.Text) = 0 Then Exit Sub
        txtDateAdd.Text = FormatDate(txtDateAdd)
End Sub

Private Sub txtEmpCode_LostFocus()
txtEmpCode.Text = UCase(txtEmpCode.Text)
'If Len(txtEmpcode.Text) = 0 Then Exit Sub
'    Set r = New ADODB.Recordset
'    StrRecord = "SELECT WStatus FROM Employee_Information WHERE EMPLOYEECODE = '" & txtEmpcode.Text & "' AND MLEFT=0;"
'
'    r.Open StrRecord, MainConn, adOpenStatic
'    If r.RecordCount > 0 Then
'      If r![WSTATUS] = "W" Then
'        optWorker.Enabled = True
'      ElseIf r![WSTATUS] = "F" Then
'        optStaff.Enabled = True
'      ElseIf r![WSTATUS] = "J" Then
'        optJrMgt.Enabled = True
'      ElseIf r![WSTATUS] = "S" Then
'        optSrMgt.Enabled = True
'      End If
'    Else
'      MsgBox "EMPLOYEE CODE NOT FOUND!!!", vbOKCancel, cnstMsgInfo
'      txtEmpcode.SetFocus
'    End If
End Sub


