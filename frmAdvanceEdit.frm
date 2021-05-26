VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAdvanceEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7566
   ClientLeft      =   39
   ClientTop       =   -65
   ClientWidth     =   7345
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
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
   ScaleHeight     =   7566
   ScaleWidth      =   7345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   90
      TabIndex        =   23
      Top             =   1890
      Width           =   7185
      Begin MSFlexGridLib.MSFlexGrid flexHistory 
         Height          =   4335
         Left            =   60
         TabIndex        =   3
         Top             =   210
         Width           =   7035
         _ExtentX        =   12412
         _ExtentY        =   7644
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   300
         AllowUserResizing=   3
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Delete"
      Height          =   405
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6870
      Width           =   1365
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   2115
   End
   Begin VB.ComboBox cmbAdjustmentStartMonth 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2250
      Width           =   2115
   End
   Begin VB.TextBox txtAdvanceDate 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   2850
      Width           =   2115
   End
   Begin VB.TextBox txtAdvanceAmount 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   2850
      Width           =   2265
   End
   Begin VB.TextBox txtAdjustmentStartYear 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   2250
      Width           =   2265
   End
   Begin VB.CommandButton cmdAdjustment 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Adjustment"
      Height          =   375
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Update"
      Height          =   405
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6870
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D8E9EC&
      Caption         =   "E&xit"
      Height          =   405
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6870
      Width           =   1365
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Refresh"
      Height          =   405
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6870
      Width           =   1365
   End
   Begin VB.TextBox txtAdjustmentAmount 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4890
      TabIndex        =   6
      Top             =   2250
      Width           =   2085
   End
   Begin VB.ComboBox cmbEmployeeCode 
      Height          =   315
      Left            =   4890
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   2085
   End
   Begin VB.ComboBox cmbSubDepartment 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2265
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3255
      Left            =   330
      TabIndex        =   10
      Top             =   3270
      Width           =   6705
      _ExtentX        =   11837
      _ExtentY        =   5751
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   300
      AllowUserResizing=   3
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   26
      Top             =   6600
      Width           =   7455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Edit Advance Entry"
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
      TabIndex        =   25
      Top             =   240
      Width           =   2835
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Department :"
      Height          =   255
      Left            =   390
      TabIndex        =   21
      Top             =   1290
      Width           =   1425
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Adj. Amount :"
      Height          =   195
      Left            =   4890
      TabIndex        =   20
      Top             =   1980
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Adj. Start Year :"
      Height          =   195
      Left            =   2520
      TabIndex        =   19
      Top             =   1980
      Width           =   2265
   End
   Begin VB.Label Label6 
      Caption         =   "Adj. Start Month  :"
      Height          =   195
      Left            =   390
      TabIndex        =   18
      Top             =   1980
      Width           =   1425
   End
   Begin VB.Label Label5 
      Caption         =   "Sub Department :"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   1290
      Width           =   2265
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Code :"
      Height          =   255
      Left            =   4890
      TabIndex        =   16
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Advance Date :"
      Height          =   195
      Left            =   390
      TabIndex        =   15
      Top             =   2610
      Width           =   1425
   End
   Begin VB.Label Label2 
      Caption         =   "Amount :"
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   2610
      Width           =   2265
   End
End
Attribute VB_Name = "frmAdvanceEdit"
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
    
    flexHistory.Rows = 1
    Frame1.Visible = True
End Sub

Private Sub cmbDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbEmployeeCode_Click()
    Dim SLNO As Integer
    Set r = New Recordset
    
     ''Amount, Date, AdjustStartMonth, AdjustStartYear
    Mysql = "Select Amount, Date, AdjustStartMonth, AdjustStartYear, AdjustAmount From AdvanceMain Where comid = " & CompanyID & " and EmpCode = '" & Trim(cmbEmployeeCode.Text) & "' and DeptName = '" & Trim(CmbDepartment.Text) & "' and SubDeptName='" & Trim(cmbSubDepartment.Text) & "'"
    MainComm.CommandText = Mysql
    Set r = MainComm.Execute
    
    If r.EOF = False And r.BOF = False Then
        SLNO = 1
        Do While Not r.EOF
            With flexHistory
                .Rows = .Rows + 1
                .row = .Rows - 1
               
                .col = 0:   .Text = SLNO
                .col = 1:   .Text = r("Amount").Value
                .col = 2:   .Text = r("Date").Value
                .col = 3:   .Text = r("AdjustStartMonth").Value
                .col = 4:   .Text = r("AdjustStartYear").Value
                .col = 5:   .Text = r("AdjustAmount").Value
            End With
            SLNO = SLNO + 1
        r.MoveNext
        Loop
    Else
        flexHistory.Rows = 1
    End If
    Set r = Nothing
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

Private Sub cmdDelete_Click()
    DeleteRow = 1
    If IsNull(bytUserID) = True And IsNull(getComVolID) = True Then
    MsgBox "You have not sufficient Privilege", vbInformation, cnstMsgInfo
    Exit Sub
    End If

On Error GoTo x
    If flexDetails.Rows = 1 Then Exit Sub
    
    If MsgBox("Do you want to delete advance information", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    Dim AdvanceID As Integer
    Set r = New Recordset
    
    Mysql = "Select AdvanceId from AdvanceMain where comid = " & CompanyID & " and empcode = '" & Trim(cmbEmployeeCode.Text) & "' and Date = '" & Format(txtAdvanceDate.Text, cnstDtFrmtI) & "'"
    MainComm.CommandText = Mysql
    Set r = MainComm.Execute
    
    If r.EOF = False And r.BOF = False Then
        AdvanceID = r("AdvanceID")
    Else
        MsgBox "Data Not Found To Delete.", vbCritical, App.Title
        Exit Sub
    End If
    
    MainConn.BeginTrans
    
      MainComm.CommandText = "SET NOCOUNT ON INSERT INTO  AdvanceMain_DelEdit " _
                            & " Select AdvanceID,ComID,EmpCode,DeptName,SubDeptName,Amount,[Date],AdjustStartMont,AdjustStartYear,AdjustAmount," & DeleteRow & " ," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() From AdvanceMain " _
                            & " Where AdvanceID = " & AdvanceID & ""
          MainComm.Execute
    
        Mysql = "Delete Advancesub Where AdvanceID = " & AdvanceID & ""
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = "Delete AdvanceMain Where AdvanceID = " & AdvanceID & ""
        MainComm.CommandText = Mysql
        MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Data Deleted Successfully.", vbInformation, App.Title
    
    Set r = Nothing
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
MainConn.RollbackTrans
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    CmbDepartment.ListIndex = -1: cmbSubDepartment.Clear: cmbEmployeeCode.Clear
    
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
    MainConn.BeginTrans
        NewID = fncFindID
        
        Mysql = "Delete AdvanceSub Where AdvanceID = " & NewID & ""
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Update AdvanceMain "
        Mysql = Mysql + " Set Amount = " & CCur(txtAdvanceAmount.Text) & ", Date = '" & Format(txtAdvanceDate.Text, cnstDtFrmtQ) & "', "
        Mysql = Mysql + " AdjustStartMonth = '" & Trim(cmbAdjustmentStartMonth.Text) & "', AdjustStartYear = " & txtAdjustmentStartYear.Text & " "
        Mysql = Mysql + " Where comid = " & CompanyID & " and  AdvanceId = " & NewID & ""
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


Private Sub flexHistory_DblClick()
    Call flexHistory_KeyDown(13, 0)
End Sub

Private Sub flexHistory_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Dim dt As Date, SLNO As Integer
        With flexHistory
            .col = 2: dt = Format(.Text, cnstDtFrmtQ)
        End With
        Set r = New Recordset
        
        Mysql = " SELECT M.AdvanceID, M.DeptName, M.SubDeptName, M.EmpCode, M.Amount AS AdvAmount, M.Date, M.AdjustStartMonth, M.AdjustStartYear, M.AdjustAmount, S.Month, S.Year, S.Amount FROM AdvanceMain AS M INNER JOIN AdvanceSub AS S ON M.AdvanceID = S.AdvanceID "
        Mysql = Mysql + " Where M.Date = '" & dt & "' and DeptName = '" & Trim(CmbDepartment.Text) & "' and  SubDeptName = '" & Trim(cmbSubDepartment.Text) & "' and  EmpCode = '" & Trim(cmbEmployeeCode.Text) & "' and Comid = " & CompanyID & " Order By S.SLNO"
        MainComm.CommandText = Mysql
        Set r = MainComm.Execute
        
        If r.EOF = False And r.BOF = False Then
            cmbAdjustmentStartMonth.Text = r("AdjustStartMonth")
            txtAdjustmentStartYear.Text = r("AdjustStartYear")
            txtAdjustmentAmount.Text = FormatNumber(r("AdjustAmount"), 2)
            txtAdvanceDate.Text = Format(r("Date"), cnstDtFrmtI)
            txtAdvanceAmount.Text = FormatNumber(r("AdvAmount"), 2)
            
            SLNO = 1
            Do While Not r.EOF
                With flexDetails
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                   
                    .col = 0:   .Text = SLNO
                    .col = 1:   .Text = r("Year")
                    .col = 2:   .Text = r("Month")
                    .col = 3:   .Text = FormatNumber(r("Amount"), 2)
                End With
                SLNO = SLNO + 1
            r.MoveNext
            Loop
            Frame1.Visible = False
        End If
        Set r = Nothing
    End If
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.width - width) / 2, (Screen.Height - Height) / 2
        
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

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
        .SelectionMode = flexSelectionByRow
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
    
    ''Amount, Date, AdjustStartMonth, AdjustStartYear
    With flexHistory
        .SelectionMode = flexSelectionByRow
        .Cols = 6
        .Rows = 1
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        .ColWidth(5) = 1200
        
        .row = 0
        .col = 0:   .Text = "SL#"
        .col = 1:   .Text = "AMOUNT"
        .col = 2:   .Text = "ADV. DATE"
        .col = 3:   .Text = "ADJ. START MONTH"
        .col = 4:   .Text = "ADJ. START YEAR"
        .col = 5:   .Text = "ADJ. AMOUNT"
    End With
End Sub

Public Sub prcAddDepartment()
    Dim tmprst As New Recordset
    Mysql = "Select DeptName from Dept_Name Where ComID = " & CompanyID & " Order By DeptName"
    
    MainComm.CommandText = Mysql
    Set tmprst = MainComm.Execute
    
    CmbDepartment.Clear
    If tmprst.EOF = False And tmprst.BOF = False Then
        tmprst.MoveFirst
        Do While Not tmprst.EOF
            CmbDepartment.AddItem tmprst.Fields(0).Value
        tmprst.MoveNext
        Loop
    End If
    Set tmprst = Nothing
End Sub

Public Sub prcAddSubDepartment()
    Dim tmprst As New Recordset
    Mysql = "Select SubDeptName from Sub_Dept_Name Where DeptName = '" & CmbDepartment.Text & "' and ComID=" & CompanyID & " Order By SubDeptName"
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
    Mysql = "Select EmployeeCode from Employee_Information Where ComID = " & CompanyID & " and  Department = '" & CmbDepartment.Text & "' and Section = '" & Trim(cmbSubDepartment.Text) & "' and MLeave=0 and MLeft=0 Order By EmployeeCode"
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

Public Function fncFindID() As Integer
    Set r = New Recordset
        Mysql = "Select AdvanceId from AdvanceMain where comid = " & CompanyID & " and empcode = '" & Trim(cmbEmployeeCode.Text) & "' and Date = '" & Format(txtAdvanceDate.Text, cnstDtFrmtQ) & "'"
        MainComm.CommandText = Mysql
        Set r = MainComm.Execute
        
        If r.EOF = False And r.BOF = False Then fncFindID = r("AdvanceID")
    Set r = Nothing
End Function


