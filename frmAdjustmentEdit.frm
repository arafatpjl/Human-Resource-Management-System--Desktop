VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmadjustmentEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdjustmentEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   300
      Top             =   7200
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   7800
      Width           =   1635
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5235
      Left            =   180
      TabIndex        =   2
      Top             =   1860
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   9234
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.ComboBox cmbsubdeptname 
      Height          =   315
      Left            =   7553
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1380
      Width           =   2475
   End
   Begin VB.ComboBox cmbdeptname 
      Height          =   315
      Left            =   2723
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   2955
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Temporary Adjustment Update"
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Width           =   12375
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12375
   End
   Begin VB.Label Label2 
      Caption         =   "Sub-Department :"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   1410
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Department :"
      Height          =   255
      Left            =   1590
      TabIndex        =   4
      Top             =   1380
      Width           =   1275
   End
End
Attribute VB_Name = "frmadjustmentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbdeptname_Click()
    Call SubDeptName(cmbsubdeptname, cmbdeptname)
End Sub

Private Sub cmbdeptname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbsubdeptname_Click()
    If Len(cmbsubdeptname.Text) > 0 Then
        If Adodc1.Recordset.State = adStateOpen Then: Adodc1.Recordset.CancelBatch: Adodc1.Recordset.Close
        If boJRMgt = False Then
            'Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, ProposedGross, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = " & CompanyID & " and WorkingYear = " & PeriodYear & " and Deptname='" & cmbdeptname.Text & "' and subdeptname='" & cmbsubdeptname.Text & "' and left(empCode, 2) <>'JR' Order By Empcode"
            Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, CurrRate, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = " & CompanyID & " and WorkingYear = " & PeriodYear & " and Deptname='" & cmbdeptname.Text & "' and subdeptname='" & cmbsubdeptname.Text & "' and left(empCode, 2) <>'JR' Order By Empcode"
        Else
            'Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, ProposedGross, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = " & CompanyID & " and WorkingYear = " & PeriodYear & " and Deptname='" & cmbdeptname.Text & "' and subdeptname='" & cmbsubdeptname.Text & "' Order By Empcode"
            Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, CurrRate, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = " & CompanyID & " and WorkingYear = " & PeriodYear & " and Deptname='" & cmbdeptname.Text & "' and subdeptname='" & cmbsubdeptname.Text & "' Order By Empcode"
        End If
        Adodc1.Recordset.Open Mysql
        Set DataGrid1.DataSource = Adodc1
        
        Call prcGridSetting
    End If
End Sub

Private Sub cmbsubdeptname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        DataGrid1.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim DIFFAMT As Currency
    
    With DataGrid1
        If .Columns(7).Value = 0 Then
            If ColIndex = 4 Then
                DIFFAMT = .Columns(4).Value - .Columns(3).Value
                .Columns(7).Value = DIFFAMT
                '.Columns(6).Value = .Columns(6).Value + DIFFAMT
                .Columns(6).Value = Val(.Columns(5).Value) + Val(.Columns(4).Value)
            Else
                DIFFAMT = .Columns(6).Value - .Columns(5).Value
                .Columns(7).Value = DIFFAMT
                .Columns(4).Value = .Columns(4).Value + DIFFAMT
            End If
        Else
            
            If ColIndex = 4 Then
                DIFFAMT = (.Columns(4).Value - .Columns(3).Value) + .Columns(7).Value
                .Columns(7).Value = .Columns(7).Value - DIFFAMT
                '.Columns(6).Value = .Columns(6).Value + DIFFAMT
                .Columns(6).Value = Val(.Columns(5).Value) + Val(.Columns(4).Value)
            Else
                DIFFAMT = (.Columns(6).Value - .Columns(5).Value) + .Columns(7).Value
                .Columns(7).Value = .Columns(7).Value - DIFFAMT
                .Columns(4).Value = .Columns(4).Value + DIFFAMT
            End If
        End If
    End With
End Sub

Private Sub DataGrid1_GotFocus()
    DataGrid1.col = 4
    DataGrid1.SetFocus
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo x
''    If Shift = 0 Then
''        If KeyCode = 9 Or KeyCode = 13 Then
''            With DataGrid1
''                If .Col = 4 Then
''                    .Col = 6
''                Else
''                    .Row = .Row + 1: .Col = 4
''                End If
''                .SetFocus
''            End With
''        End If
''    Else
''        If KeyCode = 9 Or KeyCode = 13 Then
''            With DataGrid1
''                If .Col = 6 Then
''                    .Col = 4
''                Else
''                    .Row = .Row + 1: .Col = 6
''                End If
''                .SetFocus
''            End With
''        End If
''    End If
Exit Sub
x:
    If Err.Number = 6148 Then
        Adodc1.Recordset.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    'Move (Screen.width - width) / 2, (Screen.Height - Height) / 2
    Top = (frmadjustmentEdit.Height - Height) / 2
  Left = (frmadjustmentEdit.width - width) / 2 + 1600
    'Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, ProposedGross, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = 0"
    Mysql = "Select EmpCode, EmpName, Designation, ProposedAmount, IncrementAmount, CurrRate, GrossWillBe, Difference, NewDesignation From Temp_adjustment_2010dec Where Comid = 0"
    Call prcAdodc(Adodc1, Mysql)
    Set DataGrid1.DataSource = Adodc1
    Call prcGridSetting
    
    Call DeptName(cmbdeptname)
    cmbsubdeptname.Clear
End Sub

Public Sub prcGridSetting()
    With DataGrid1
        .Columns(0).width = 1000
        .Columns(1).width = 2000
        .Columns(2).width = 1500
        .Columns(3).width = 1200
        .Columns(4).width = 1200
        .Columns(5).width = 1200
        .Columns(6).width = 1200
        .Columns(7).width = 0
        .Columns(8).width = 0
        
        .Columns(0).Caption = "Emp Code"
        .Columns(1).Caption = "Name"
        .Columns(2).Caption = "Designation"
        .Columns(3).Caption = "Prop. Amount"
        .Columns(4).Caption = "Adjust. Amount"
        .Columns(5).Caption = "Curr. Rate"
        .Columns(6).Caption = "Gross Will Be"
        .Columns(7).Caption = "Diff Amount"
        .Columns(8).Caption = "New Designation"
        
        .Columns(0).Locked = True
        .Columns(1).Locked = True
        .Columns(2).Locked = True
        .Columns(3).Locked = True
        .Columns(4).Locked = False
        .Columns(5).Locked = True
        .Columns(6).Locked = False
        .Columns(7).Locked = True
        .Columns(8).Locked = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTempIncrementEdit = Nothing
End Sub
