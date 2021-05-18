VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAdvance 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   7335
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2325
   End
   Begin VB.ComboBox cmbAdjustmentStartMonth 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2130
      Width           =   2325
   End
   Begin VB.TextBox txtAdvanceDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   2730
      Width           =   2325
   End
   Begin VB.TextBox txtAdvanceAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2670
      TabIndex        =   7
      Top             =   2730
      Width           =   2265
   End
   Begin VB.TextBox txtAdjustmentStartYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2670
      TabIndex        =   4
      Top             =   2130
      Width           =   2265
   End
   Begin VB.CommandButton cmdAdjustment 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Adjustment"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Save"
      Height          =   405
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D8E9EC&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   405
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   1365
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Refresh"
      Height          =   405
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   1365
   End
   Begin VB.TextBox txtAdjustmentAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5040
      TabIndex        =   5
      Top             =   2130
      Width           =   2085
   End
   Begin VB.ComboBox cmbEmployeeCode 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   2085
   End
   Begin VB.ComboBox cmbSubDepartment 
      Height          =   315
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   2265
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   3150
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   300
      AllowUserResizing=   3
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   " Employee Advance Entry"
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
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   6480
      Width           =   9015
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      Caption         =   "Department :"
      Height          =   255
      Left            =   210
      TabIndex        =   20
      Top             =   1170
      Width           =   1485
   End
   Begin VB.Label Label9 
      Caption         =   "Monthly Adj. Amount :"
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   19
      Top             =   1860
      Width           =   2085
   End
   Begin VB.Label Label7 
      Caption         =   "Adj. Start Year :"
      Height          =   195
      Left            =   2670
      TabIndex        =   18
      Top             =   1860
      Width           =   2265
   End
   Begin VB.Label Label6 
      Caption         =   "Adj. Start Month  :"
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   1860
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "Sub Department :"
      Height          =   255
      Left            =   2670
      TabIndex        =   16
      Top             =   1170
      Width           =   2265
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Code :"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   1170
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Advance Date :"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   2490
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Total Advance Amount :"
      Height          =   195
      Left            =   2670
      TabIndex        =   13
      Top             =   2490
      Width           =   2265
   End
End
Attribute VB_Name = "frmAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intMonthID As Integer, intYear As Integer

Private Sub cmbAdjustmentStartMonth_Click()
    cmbAdjustmentStartMonth.Tag = 0
    If cmbAdjustmentStartMonth.ListIndex >= 0 Then cmbAdjustmentStartMonth.Tag = cmbAdjustmentStartMonth.ItemData(cmbAdjustmentStartMonth.ListIndex)
End Sub

Private Sub cmbAdjustmentStartMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbDepartment_Click()
    Call prcAddSubDepartment
    cmbEmployeeCode.Clear
End Sub

Private Sub cmbDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbEmployeeCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbSubDepartment_Click()
    Call prcAddEmployeCode
End Sub

Private Sub cmbsubDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAdjustment_Click()
    If fncBlank(1) = True Then Exit Sub
    Call prcAdjustAmount
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    
    cmbDepartment.ListIndex = -1: cmbSubDepartment.Clear: cmbEmployeeCode.Clear
    
    txtAdvanceDate.Text = ""
    txtAdvanceAmount.Text = ""
    cmbAdjustmentStartMonth.ListIndex = -1
    txtAdjustmentStartYear.Text = ""
    txtAdjustmentAmount.Text = ""
    
    flexDetails.Rows = 1
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    Dim NewID As Integer
    Dim sMonth As String, iYear As Integer, cAmount As Currency
    
    If flexDetails.Rows = 1 Then Exit Sub
    
    If MsgBox("Do You Want To Save.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    Set r = New Recordset
    Mysql = "SELECT COUNT(*) FROM ADVANCEMAIN WHERE COMID =" & CompanyID & " AND DeptName='" & Trim(cmbDepartment.Text) & "' AND SubDeptName='" & Trim(cmbSubDepartment.Text) & "' AND empcode = '" & Trim(cmbEmployeeCode.Text) & "' AND DATE = '" & Format(txtAdvanceDate.Text, cnstDtFrmtQ) & "'"
    MainComm.CommandText = Mysql
    Set r = MainComm.Execute
    If r.Fields(0).Value > 1 Then
        MsgBox "DUPLICATE ENTRY.", vbCritical, App.Title
        Exit Sub
    End If
    
    MainConn.BeginTrans
        NewID = fncNewID
        
        Mysql = "Insert AdvanceMain (AdvanceID, Comid, EmpCode, DeptName, SubDeptName, Amount, Date, AdjustStartMonth, AdjustStartYear, AdjustAmount)"
        Mysql = Mysql + " Values (" & NewID & ", " & CompanyID & ", '" & cmbEmployeeCode.Text & "', '" & cmbDepartment.Text & "', '" & cmbSubDepartment.Text & "', " & CCur(txtAdvanceAmount.Text) & ",  '" & Format(txtAdvanceDate.Text, cnstDtFrmtQ) & "', '" & cmbAdjustmentStartMonth.Text & "', " & CDbl(txtAdjustmentStartYear.Text) & ", " & CCur(txtAdjustmentAmount.Text) & ") "
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        With flexDetails
            For i = 1 To .Rows - 1
                .row = i
                .col = 1: iYear = .Text
                .col = 2: sMonth = .Text
                .col = 3: cAmount = .Text
                
                Mysql = " Insert Into AdvanceSub (AdvanceID, Month, Year, Amount)"
                Mysql = Mysql + " Values (" & NewID & ", '" & sMonth & "', " & iYear & ", " & cAmount & ")"
                MainComm.CommandText = Mysql
                MainComm.Execute
            Next
        End With
    MainConn.CommitTrans
    MsgBox "Data Saved Successfully.", vbInformation, App.Title
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
MainConn.RollbackTrans
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.width - width) / 2, (Screen.Height - Height) / 2
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    '' SETTING GRID FLEXDETAILS
    Call prcGridSetting
    
    '' LOAD DEPARTMENT
    Call prcAddDepartment
    
    '' LOAD MONTH NAME
    Call prcMonthName(cmbAdjustmentStartMonth)
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Public Sub prcGridSetting()
    With flexDetails
        .Cols = 4
        .Rows = 1
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        
        .row = 0
        .col = 0:   .Text = "SL#"
        .col = 1:   .Text = "YEAR"
        .col = 2:   .Text = "MONTH"
        .col = 3:   .Text = "AMOUNT"
    End With
End Sub

Public Sub prcAddDepartment()
    Dim tmprst As New Recordset
    Mysql = "Select DeptName from Dept_Name Where ComID = " & CompanyID & " Order By DeptName"
    
    MainComm.CommandText = Mysql
    Set tmprst = MainComm.Execute
    
    cmbDepartment.Clear
    If tmprst.EOF = False And tmprst.BOF = False Then
        tmprst.MoveFirst
        Do While Not tmprst.EOF
            cmbDepartment.AddItem tmprst.Fields(0).Value
        tmprst.MoveNext
        Loop
    End If
    Set tmprst = Nothing
End Sub

Public Sub prcAddSubDepartment()
    Dim tmprst As New Recordset
    Mysql = "Select SubDeptName from Sub_Dept_Name Where DeptName = '" & cmbDepartment.Text & "' and ComID=" & CompanyID & " Order By SubDeptName"
    MainComm.CommandText = Mysql
    Set tmprst = MainComm.Execute
    
    cmbSubDepartment.Clear
    If tmprst.EOF = False And tmprst.BOF = False Then
        tmprst.MoveFirst
        Do While Not tmprst.EOF
            cmbSubDepartment.AddItem tmprst.Fields(0).Value
        tmprst.MoveNext
        Loop
    End If
    Set tmprst = Nothing
End Sub

Public Sub prcAddEmployeCode()
    Dim tmprst As New Recordset
    Mysql = "Select EmployeeCode from Employee_Information Where ComID = " & CompanyID & " and  Department = '" & cmbDepartment.Text & "' and Section = '" & Trim(cmbSubDepartment.Text) & "' and MLeave=0 and MLeft=0 Order By EmployeeCode"
    MainComm.CommandText = Mysql
    Set tmprst = MainComm.Execute
    
    cmbEmployeeCode.Clear
    If tmprst.EOF = False And tmprst.BOF = False Then
        tmprst.MoveFirst
        Do While Not tmprst.EOF
            cmbEmployeeCode.AddItem tmprst.Fields(0).Value
        tmprst.MoveNext
        Loop
    End If
    Set tmprst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAdvance = Nothing
End Sub
Private Sub txtAdjustmentAmount_GotFocus()
    txtAdjustmentAmount.SelStart = 0
    txtAdjustmentAmount.SelLength = Len(txtAdjustmentAmount.Text)
End Sub

Private Sub txtAdjustmentAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAdjustmentAmount_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
    If KeyAscii = 46 Then
        If InStr(1, txtAdjustmentAmount.Text, ".") > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAdjustmentAmount_LostFocus()
    If Len(txtAdjustmentAmount.Text) = 0 Then txtAdjustmentAmount.Text = "0.00"
    
    txtAdjustmentAmount.Text = FormatNumber(txtAdjustmentAmount.Text, 2)
End Sub

Private Sub txtAdjustmentStartYear_GotFocus()
    txtAdjustmentStartYear.SelStart = 0
    txtAdjustmentStartYear.SelLength = Len(txtAdjustmentStartYear.Text)
End Sub

Private Sub txtAdjustmentStartYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAdjustmentStartYear_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtAdjustmentStartYear_LostFocus()
    If Len(txtAdjustmentStartYear.Text) = 0 Then txtAdjustmentStartYear.Text = Year(Date)
End Sub

Private Sub txtAdvanceAmount_GotFocus()
    txtAdvanceAmount.SelStart = 0
    txtAdvanceAmount.SelLength = Len(txtAdvanceAmount.Text)
End Sub

Private Sub txtAdvanceAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAdvanceAmount_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
    If KeyAscii = 46 Then
        If InStr(1, txtAdjustmentAmount.Text, ".") > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAdvanceAmount_LostFocus()
    If Len(txtAdvanceAmount.Text) = 0 Then txtAdvanceAmount.Text = "0.00"
    txtAdvanceAmount.Text = FormatNumber(txtAdvanceAmount.Text, 2)
End Sub

Private Sub txtAdvanceDate_GotFocus()
    txtAdvanceDate.SelStart = 0
    txtAdvanceDate.SelLength = Len(txtAdvanceDate.Text)
End Sub

Private Sub txtAdvanceDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAdvanceDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtAdvanceDate_LostFocus()
    Call FormatDate(txtAdvanceDate)
End Sub

Public Function fncBlank(intval As Integer) As Boolean
fncBlank = True
    If intval = 1 Then
        If Len(txtAdvanceAmount.Text) = 0 Then txtAdvanceAmount.Text = "0.00"
        If Len(txtAdjustmentAmount.Text) = 0 Then txtAdjustmentAmount.Text = "0.00"
    
        If Len(cmbEmployeeCode.Text) = 0 Then
            MsgBox "You should provide employee code", vbCritical, App.Title
            cmbEmployeeCode.SetFocus
            Exit Function
        End If
        
        If Len(txtAdvanceDate.Text) = 0 Then
            MsgBox "You should provide date of advance.", vbCritical, App.Title
            txtAdvanceDate.SetFocus
            Exit Function
        End If
        If CCur(txtAdvanceAmount.Text) <= 0 Then
            MsgBox "You should provide advance amount.", vbCritical, App.Title
            txtAdvanceAmount.SetFocus
            Exit Function
        End If
        
        If Len(cmbAdjustmentStartMonth.Text) = 0 Then
            MsgBox "You should provide month to start adjust.", vbCritical, App.Title
            cmbAdjustmentStartMonth.SetFocus
            Exit Function
        End If
        If Len(txtAdjustmentStartYear.Text) = 0 Then
            MsgBox "You should provide year to adjust advance amount.", vbCritical, App.Title
            txtAdjustmentStartYear.SetFocus
            Exit Function
        End If
    
        If CCur(txtAdjustmentAmount.Text) <= 0 Then
            MsgBox "You should provide adjustment amount.", vbCritical, App.Title
            txtAdjustmentAmount.SetFocus
            Exit Function
        End If
    
        If CCur(txtAdvanceAmount.Text) < CCur(txtAdjustmentAmount.Text) Then
            MsgBox "Advance amount should be greater or equal to adjustment amount.", vbCritical, App.Title
            txtAdjustmentAmount.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Public Sub prcAdjustAmount()
    Dim AdvAmount As Currency, AdjAmount As Currency
    Dim RTBalance As Currency, SLNO As Integer
    
    AdvAmount = CCur(txtAdvanceAmount.Text)
    AdjAmount = CCur(txtAdjustmentAmount.Text)
    RTBalance = AdvAmount
    
    flexDetails.Rows = 1:     SLNO = 1
    
    intMonthID = cmbAdjustmentStartMonth.Tag
    intYear = txtAdjustmentStartYear.Text
    
    While RTBalance > 0
        If RTBalance >= AdjAmount Then
            With flexDetails
                .Rows = .Rows + 1
                .row = .Rows - 1
               
                .col = 0:   .Text = SLNO
                .col = 1:   .Text = intYear
                .col = 2:   .Text = MonthName(intMonthID)
                .col = 3:   .Text = FormatNumber(AdjAmount, 2)
            End With
            
            RTBalance = RTBalance - AdjAmount
            SLNO = SLNO + 1
            
            If intMonthID = 12 Then
                intMonthID = 1
                intYear = intYear + 1
            Else
                intMonthID = intMonthID + 1
            End If
        
        ElseIf RTBalance > 0 Then
            With flexDetails
                .Rows = .Rows + 1
                .row = .Rows - 1
               
                .col = 0:   .Text = SLNO
                .col = 1:   .Text = intYear
                .col = 2:   .Text = MonthName(intMonthID)
                .col = 3:   .Text = FormatNumber(RTBalance, 2)
            End With
            RTBalance = 0
        End If
    Wend
End Sub

Public Function fncNewID() As Integer
    Mysql = "Select isnull(Max(AdvanceID),0)+1 as MaxID from AdvanceMain Where Comid = " & CompanyID & ""
    MainComm.CommandText = Mysql
    Set r = MainComm.Execute
    fncNewID = r("MaxID")
End Function
