VERSION 5.00
Begin VB.Form frmEditBonus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6513
   ClientLeft      =   3042
   ClientTop       =   2041
   ClientWidth     =   5252
   ControlBox      =   0   'False
   Icon            =   "frmEditBonus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6513
   ScaleWidth      =   5252
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
      Left            =   120
      TabIndex        =   34
      Tag             =   "W"
      Top             =   1080
      Width           =   5055
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
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   1035
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
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
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
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox txtFracD 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Fraction Pay"
      Top             =   4230
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
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
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Click for Form Close"
      Top             =   5880
      Width           =   1215
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
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Click for Delete Record"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.OptionButton optHeldNo 
      Caption         =   "No"
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
      Height          =   195
      Left            =   2970
      TabIndex        =   26
      Top             =   5070
      Width           =   615
   End
   Begin VB.OptionButton optHeldYes 
      Caption         =   "Yes"
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
      Height          =   195
      Left            =   2130
      TabIndex        =   25
      Top             =   5070
      Width           =   615
   End
   Begin VB.TextBox txtAmount 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2010
      TabIndex        =   7
      Top             =   4620
      Width           =   1575
   End
   Begin VB.ComboBox cboSubDept 
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
      Height          =   315
      Left            =   2010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2610
      Width           =   3015
   End
   Begin VB.ComboBox cboDept 
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
      Height          =   315
      Left            =   2010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2205
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
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
      Height          =   375
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "Click for Save Record"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtRatePay 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3825
      Width           =   1575
   End
   Begin VB.ComboBox cboMonthName 
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
      Height          =   315
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ComboBox cboEmpCode 
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
      Height          =   315
      Left            =   2010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3015
      Width           =   3015
   End
   Begin VB.TextBox txtEmpName 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3420
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Edit Bonus"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.63
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   32
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1770
      TabIndex        =   30
      Top             =   4230
      Width           =   255
   End
   Begin VB.Label Label83 
      Caption         =   "Fraction"
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
      Left            =   210
      TabIndex        =   29
      Top             =   4230
      Width           =   1335
   End
   Begin VB.Label Label56 
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
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   1770
      TabIndex        =   28
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label57 
      Caption         =   "Held Up"
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
      Left            =   210
      TabIndex        =   27
      Top             =   5055
      Width           =   975
   End
   Begin VB.Label Label44 
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
      Left            =   1770
      TabIndex        =   24
      Top             =   4620
      Width           =   255
   End
   Begin VB.Label Label45 
      Caption         =   "Bonus Amount"
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
      Left            =   210
      TabIndex        =   23
      Top             =   4620
      Width           =   1335
   End
   Begin VB.Label Label60 
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
      Left            =   1770
      TabIndex        =   22
      Top             =   2610
      Width           =   255
   End
   Begin VB.Label Label59 
      Caption         =   "Sub Department"
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
      Left            =   210
      TabIndex        =   21
      Top             =   2610
      Width           =   1455
   End
   Begin VB.Label Label58 
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   20
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1770
      TabIndex        =   19
      Top             =   2205
      Width           =   255
   End
   Begin VB.Label Label31 
      Caption         =   "Rate of Pay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   18
      Top             =   3825
      Width           =   1335
   End
   Begin VB.Label Label30 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1770
      TabIndex        =   17
      Top             =   3825
      Width           =   255
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1770
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label29 
      Caption         =   "Payment Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   14
      Top             =   3420
      Width           =   495
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1770
      TabIndex        =   13
      Top             =   3420
      Width           =   255
   End
   Begin VB.Label Label11 
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
      Left            =   1770
      TabIndex        =   12
      Top             =   3015
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Employee Code"
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
      Left            =   210
      TabIndex        =   11
      Top             =   3015
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDept_Click()
    If Len(cboDept) = 0 Then Exit Sub
    SubDeptName cboSubDept, cboDept
End Sub
Private Sub cboEmpCode_Click()
    txtEmpName.Text = ""
    txtRatePay.Text = ""
    txtFracD.Text = ""
    txtAmount.Text = ""
    optHeldNo.Value = True
    optHeldYes.Value = False
    If Len(cboEmpCode) = 0 Then Exit Sub
    
    Set r = New ADODB.Recordset
        StrRecord = "SELECT Name FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode & "'));"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.EOF = False And r.BOF = False Then txtEmpName.Text = " " & r![Name]
        r.Close
        
        StrRecord = "SELECT RatePay,Basic,HeldUp,FracPayD FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode & "')" _
                  & " AND ([Month]='" & cboMonthName.Text & "') AND ([Year]='" & PeriodYear & "') AND (DeptName='" & cboDept.Text & "') AND (SubDeptName='" & cboSubDept.Text & "'));"
        r.Open StrRecord, MainConn, adOpenStatic
        If r.RecordCount > 0 Then
            txtRatePay.Text = Format(r![RatePay], "##,##0.00")
            txtFracD.Text = Format(r![FracPayD], "##0.00")
            txtAmount.Text = Format(r![BASIC], "##,##0.00")
            If r![HeldUp] = True Then optHeldYes.Value = True
        End If
        r.Close
    Set r = Nothing
End Sub
Private Sub cboSubDept_Click()
    cboEmpCode.Clear
End Sub
Private Sub cboSubDept_LostFocus()
    If Len(cboSubDept) = 0 Then Exit Sub
    Set r = New ADODB.Recordset
'    If boJRMgt = True And boSRMgt = True Then
'        StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
'                & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
'                & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.SubDeptName)='" & cboSubDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "')" _
'                & " AND ((Employee_Information.WStatus)<>'S')) GROUP BY Bonus_Entry.EmployeeCode;"
'    ElseIf boJRMgt = True Then
'        StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
'                & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
'                & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.SubDeptName)='" & cboSubDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "')" _
'                & " AND ((Employee_Information.WStatus)='J')) GROUP BY Bonus_Entry.EmployeeCode;"
'
'    Else
'        If Mid(cboDept.Text, 1, 5) = "STAFF" Then
            StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
                & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
                & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.SubDeptName)='" & cboSubDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "')" _
                & " AND ((Employee_Information.WStatus)='" & FrameCategory.Tag & "')) GROUP BY Bonus_Entry.EmployeeCode;"
'        Else
'            StrRecord = "SELECT Bonus_Entry.EmployeeCode FROM Bonus_Entry INNER JOIN Employee_Information ON (Employee_Information.EmployeeCode = Bonus_Entry.EmployeeCode)" _
'                & " AND (Employee_Information.ComID = Bonus_Entry.ComID) WHERE (((Bonus_Entry.ComID)=" & CompanyID & ") AND ((Bonus_Entry.Month)='" & cboMonthName.Text & "')" _
'                & " AND ((Bonus_Entry.DeptName)='" & cboDept.Text & "') AND ((Bonus_Entry.SubDeptName)='" & cboSubDept.Text & "') AND ((Bonus_Entry.Year)='" & PeriodYear & "')" _
'                & " AND ((Employee_Information.WStatus)='W')) GROUP BY Bonus_Entry.EmployeeCode;"
'        End If
'    End If
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        Do Until r.EOF
            cboEmpCode.AddItem r![EmployeeCode]
            r.MoveNext
        Loop
    End If
    r.Close
    Screen.MousePointer = vbDefault
    Set r = Nothing
End Sub
Private Sub cmdClose_Click()
    If MsgBox("Are you sure to Close", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    Unload Me
End Sub
Private Sub cmdDelete_Click()
         DeleteRow = 1
        If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
        MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
        Exit Sub
        End If

    If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    MainConn.BeginTrans
    
      MainComm.CommandText = "SET NOCOUNT ON INSERT INTO  Bonus_Entry_DelEdit " _
                            & " Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus, " & DeleteRow & " ," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE()From Bonus_Entry " _
                            & " WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode.Text & "')" _
          & " AND ([Month]='" & cboMonthName.Text & "') AND ([Year]='" & PeriodYear & "') AND (DeptName='" & cboDept.Text & "') AND (SubDeptName='" & cboSubDept.Text & "'));"
        MainComm.Execute
    MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode.Text & "')" _
          & " AND ([Month]='" & cboMonthName.Text & "') AND ([Year]='" & PeriodYear & "') AND (DeptName='" & cboDept.Text & "') AND (SubDeptName='" & cboSubDept.Text & "'));"
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
    
    MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo
    
    cboEmpCode.ListIndex = -1
    txtEmpName.Text = ""
    txtRatePay.Text = ""
    txtFracD.Text = ""
    txtAmount.Text = ""
    optHeldNo.Value = True
    cboEmpCode.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSave_Click()

Dim TypeHeldUp As Byte
    DeleteRow = 2
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
    Exit Sub
    End If

If optHeldNo.Value = True Then TypeHeldUp = 0 Else TypeHeldUp = 1

If cboMonthName.Text = "" Then
  MsgBox "Select Name of Month", vbInformation, cnstMsgInfo
  cboMonthName.SetFocus
  Exit Sub
End If
If cboDept.Text = "" Then
  MsgBox "Select Name of Department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If
If cboSubDept.Text = "" Then
  MsgBox "Select Sub Department", vbInformation, cnstMsgInfo
  cboSubDept.SetFocus
  Exit Sub
End If
If cboEmpCode.Text = "" Then
  MsgBox "Select Employee Code", vbInformation, cnstMsgInfo
  cboEmpCode.SetFocus
  Exit Sub
End If

If MsgBox("Are you sure you want to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then
  cboEmpCode.SetFocus
  Exit Sub
End If

If CInt(txtFracD.Text) < 0 Then txtFracD.Text = "0"
txtFracD.Refresh

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainComm.CommandTimeout = 500

MainConn.BeginTrans
   MainComm.CommandText = "SET NOCOUNT ON INSERT INTO  Bonus_Entry_DelEdit " _
                            & " Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus, " & DeleteRow & " ," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() From Bonus_Entry " _
                             & " WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode.Text & "') AND ([Month]='" & cboMonthName.Text & "')" _
        & " AND ([Year]='" & PeriodYear & "') AND (DeptName='" & cboDept.Text & "') AND (SubDeptName='" & cboSubDept.Text & "'));"
        MainComm.Execute


MainComm.CommandText = "SET NOCOUNT ON UPDATE Bonus_Entry SET Basic=" & CLng(txtAmount) - CInt(txtFracD.Text) & ",FracPayD=" & CInt(txtFracD.Text) & ",HeldUp=" & TypeHeldUp & ",DeleteRow=" & DeleteRow & ",UserId=" & bytUserID & ",IDNo=" & getComVolID & ",EntryDate='" & Date & "',EntryTime=GETDATE()" _
        & " WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & cboEmpCode.Text & "') AND ([Month]='" & cboMonthName.Text & "')" _
        & " AND ([Year]='" & PeriodYear & "') AND (DeptName='" & cboDept.Text & "') AND (SubDeptName='" & cboSubDept.Text & "'));"
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

MsgBox "Successfully Updated", vbInformation, cnstMsgInfo

Screen.MousePointer = vbDefault
cboEmpCode.ListIndex = -1
txtEmpName.Text = ""
txtRatePay.Text = ""
txtAmount.Text = ""

optHeldNo.Value = True
cboEmpCode.SetFocus
End Sub
Private Sub Form_Load()
  MDIfrmMain.lblCaption.Caption = "Edit Festival Bonus"
  If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"
    Set r = New ADODB.Recordset
StrRecord = "SELECT Editstatus,DelStatus FROM Sys_User_Name WHERE UserId='" & bytUserID & "';"
r.Open StrRecord, MainConn, adOpenStatic
If r![Editstatus] = 1 Then
  cmdSave.Enabled = True
  Else
   cmdSave.Enabled = False
End If
If r![DelStatus] = 1 Then
  cmdDelete.Enabled = True
  Else
   cmdDelete.Enabled = False
End If
   r.Close
   
  DeptName cboDept
  prcMonthName cboMonthName
  cboMonthName.Text = UCase(Format(Date, "MMMM"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
  boJRMgt = False
  MDIfrmMain.lblCaption.Caption = ""
  Set frmEditBonus = Nothing
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

Private Sub txtAmount_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtAmount_LostFocus()
  If txtAmount.Text <> "" Then
    txtAmount = Format(txtAmount, "##,##0.00")
  Else
    txtAmount = "0.00"
  End If
End Sub
Private Sub txtFracD_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtFracD_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


