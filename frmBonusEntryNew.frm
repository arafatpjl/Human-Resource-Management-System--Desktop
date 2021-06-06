VERSION 5.00
Begin VB.Form frmBonusEntryNew 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9139
   ClientLeft      =   1820
   ClientTop       =   1872
   ClientWidth     =   9334
   ControlBox      =   0   'False
   Icon            =   "frmBonusEntryNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9139
   ScaleWidth      =   9334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Alldept 
      Caption         =   "All Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdReportForm 
      BackColor       =   &H00C0C000&
      Caption         =   "&Report Form"
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
      Left            =   7530
      TabIndex        =   19
      Top             =   5550
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteLeftData 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete Left Data"
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
      Left            =   7530
      TabIndex        =   18
      Top             =   5100
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemoveAll 
      BackColor       =   &H00C0C000&
      Caption         =   "Remove &All"
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
      Left            =   7530
      TabIndex        =   5
      Top             =   2250
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C0C000&
      Caption         =   "&Remove"
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
      Left            =   7530
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add to List"
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
      Left            =   7530
      TabIndex        =   3
      Top             =   3270
      Width           =   1695
   End
   Begin VB.TextBox txtToDate 
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
      Height          =   315
      Left            =   5820
      TabIndex        =   2
      Top             =   1320
      Width           =   1605
   End
   Begin VB.TextBox txtHeader 
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
      Height          =   315
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1830
      Width           =   5775
   End
   Begin VB.ComboBox cmbSubDepartment 
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
      Left            =   2940
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Sub Department"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ComboBox cmbDepartment 
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
      Left            =   180
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Name of Department"
      Top             =   1320
      Width           =   2625
   End
   Begin VB.TextBox txtTotal 
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
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   12
      Top             =   7830
      Width           =   1455
   End
   Begin VB.ListBox CodeList 
      BackColor       =   &H8000000F&
      Columns         =   5
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5512
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2250
      Width           =   7305
   End
   Begin VB.ComboBox cmbMonthName 
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
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select Name of Month"
      Top             =   7830
      Width           =   2415
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
      Left            =   3960
      TabIndex        =   9
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto &Update"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   8160
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Bonus Entry New"
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
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   2820
   End
   Begin VB.Label Label84 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9375
   End
   Begin VB.Label Label26 
      Caption         =   "Name of Bonus  :"
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
      Left            =   180
      TabIndex        =   17
      Top             =   1860
      Width           =   1395
   End
   Begin VB.Label Label9 
      Caption         =   "To Joining Date :"
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
      Left            =   5820
      TabIndex        =   16
      Top             =   1050
      Width           =   1605
   End
   Begin VB.Label Label59 
      Caption         =   "Sub Department :"
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
      Left            =   2910
      TabIndex        =   15
      Top             =   1050
      Width           =   1605
   End
   Begin VB.Label Label58 
      Caption         =   "Department :"
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
      Left            =   180
      TabIndex        =   14
      Top             =   1050
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5250
      TabIndex        =   13
      Top             =   7830
      Width           =   645
   End
   Begin VB.Label Label29 
      Caption         =   "Month of Payment :"
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
      Left            =   180
      TabIndex        =   10
      Top             =   7830
      Width           =   1815
   End
End
Attribute VB_Name = "frmBonusEntryNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbDepartment_Click()
    If Len(cmbDepartment) = 0 Then Exit Sub
    Call SubDeptName(cmbSubDepartment, cmbDepartment)
End Sub

Private Sub cmbDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMonthName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbSubDepartment_Click()
    CodeList.Clear
    txtTotal.Text = ""
End Sub

Private Sub cmbsubDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    Dim strFromDate As String, strToDate As String

    If Alldept.Value = 0 Then
    If fncBlank(1) = True Then Exit Sub
    txtTotal.Text = "0"
    End If
    
    Set r = New Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    Screen.MousePointer = vbHourglass
        
        If Alldept.Value = 0 Then
        
        If Len(cmbDepartment) <> 0 And Len(cmbSubDepartment) = 0 Then
            If boJRMgt = True Then
                StrRecord = "Select EmployeeCode From Employee_Information " & _
                    " Where ((Comid = " & CompanyID & ") and (Department = '" & cmbDepartment.Text & "') and (Wstatus = 'J') and (JoiningDate <='" & txtToDate.Text & "') and (Mleft = 0))"
            Else
                StrRecord = "Select EmployeeCode From Employee_Information " & _
                    " Where ((Comid = " & CompanyID & ") and (Department = '" & cmbDepartment.Text & "') and (Wstatus In ('W', 'F')) and (JoiningDate <='" & txtToDate.Text & "') and (Mleft = 0))"
            End If
        ElseIf Len(cmbDepartment) <> 0 And Len(cmbSubDepartment) <> 0 Then
            If boJRMgt = True Then
                StrRecord = "Select EmployeeCode From Employee_Information " & _
                    " Where ((Comid = " & CompanyID & ") and (Department = '" & cmbDepartment.Text & "') and (Section ='" & cmbSubDepartment.Text & "') and (Wstatus = 'J') and (JoiningDate <='" & txtToDate.Text & "') and (Mleft = 0))"
            Else
                    StrRecord = "Select EmployeeCode From Employee_Information " & _
                    " Where ((Comid = " & CompanyID & ") and (Department = '" & cmbDepartment.Text & "') and (Section ='" & cmbSubDepartment.Text & "') and (Wstatus In ('W', 'F')) and (JoiningDate <='" & txtToDate.Text & "') and (Mleft = 0))"
            End If
        End If
        
        
        Else
        
                   StrRecord = "Select EmployeeCode From Employee_Information " & _
                   " Where ((Comid = " & CompanyID & ") and (Wstatus In ('W','F')) and (JoiningDate <='" & txtToDate.Text & "') and (Mleft = 0))"
        
        End If
        
        
        
        
        MainComm.CommandText = StrRecord
        Set r = MainComm.Execute
        If r.EOF = False And r.BOF = False Then
            Do Until r.EOF
                CodeList.AddItem r![EmployeeCode]
            r.MoveNext
            Loop
            txtTotal.Text = CStr(CodeList.ListCount)
        End If
        r.Close
    Set r = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
x:
Screen.MousePointer = vbDefault
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteLeftData_Click()
On Error GoTo x
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
    
    If MsgBox("Do you want to delete bonus information data for left employee", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    MainConn.BeginTrans
        MainComm.CommandText = "Delete Bonus_Entry From Bonus_Entry AS A " & _
            " Inner join Employee_Information As B On A.Comid = B.Comid and A.EmployeeCode = B.EmployeeCode " & _
            " where ((A.comid = " & CompanyID & ") and (A.Month = '" & cmbMonthName & "') and (A.Year =" & PeriodYear & ") and (B.MLeft = 1))"
        MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Data deletion complete successfully", vbInformation, App.Title
Exit Sub
x:
MainConn.RollbackTrans
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdRemove_Click()
    If CodeList.ListIndex >= 0 Then
        CodeList.RemoveItem (CodeList.ListIndex)
    End If
    txtTotal.Text = CStr(CodeList.ListCount)
End Sub

Private Sub cmdRemoveAll_Click()
    CodeList.Clear
    txtTotal.Text = "0"
End Sub

Private Sub cmdReportForm_Click()
    boJRMgt = True
    frmRptBonusSheet.Show 1
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    Dim sngRateBonus As Single, curAmount1 As Currency, curAmount As Currency, curFracPayD As Currency, curstamp As Currency
    If fncBlank(2) = True Then Exit Sub

    If MsgBox("Are you sure to Update Bonus", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    sngRateBonus = rateBonus

    Screen.MousePointer = vbHourglass
        Set r = New ADODB.Recordset
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainComm.CommandTimeout = 500
        DeleteRow = 2
        MainConn.BeginTrans
            curFracPayD = 0
       

            For intI = 1 To CodeList.ListCount
                CodeList.ListIndex = CodeList.ListIndex + 1
                
                MainComm.CommandText = "SET NOCOUNT ON Insert Into Bonus_Entry_DelEdit Select ComID,EmployeeCode,[Month],[Year],RatePay,Basic,Paid,HeldUp,DeptName,SubDeptName,ReportHeader,FracPayD,Wstatus," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE() FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "') AND ([Month]='" & cmbMonthName.Text & "') AND ([Year]='" & PeriodYear & "'));"
                MainComm.Execute
                
                MainComm.CommandText = "SET NOCOUNT ON DELETE FROM Bonus_Entry WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "') AND ([Month]='" & cmbMonthName.Text & "') AND ([Year]='" & PeriodYear & "'));"
                MainComm.Execute

                StrRecord = "SELECT CurrRate, Department, [Section], Wstatus FROM Employee_Information WHERE ((ComID=" & CompanyID & ") AND (EmployeeCode='" & CodeList.Text & "'));"
                r.Open StrRecord, MainConn, adOpenStatic
                
                If r.EOF = False And r.BOF = False Then
                    If r![WSTATUS] = "J" Then
                      curFracPayD = 0
                      curstamp = 0
                      curAmount = CLng((r![CurrRate] / 100 * 50) * sngRateBonus)
                    ElseIf r![WSTATUS] = "F" Then
                        curFracPayD = 0
                        curstamp = 0
                        curAmount = CLng((r![CurrRate] / 100 * V_rateBasic) * sngRateBonus)
                    Else
                        curFracPayD = 0
                        curstamp = 10
                        curAmount = CLng((((r![CurrRate] - VW_rateMedical) * VW_rateBasic) * sngRateBonus))
''                        - VW_rateMedical) * VW_rateBasic
                        curAmount1 = (curAmount - curstamp) '5 TK FOR STAMP
                        ''If curAmount1 > 0 Then curFracPayD = curAmount1 Mod 50
                        If curAmount1 > 0 Then curFracPayD = curAmount1 Mod 10
                    End If
                    
                                  
                    MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Bonus_Entry VALUES(" & CompanyID & ", '" & CodeList.Text & "', '" & cmbMonthName.Text & "', '" & PeriodYear & " ', " _
                        & " " & r![CurrRate] & ", " & curAmount - curFracPayD & ", 0, 0, '" & r![Department] & "', '" & r![Section] & "', '" & txtHeader.Text & "', " & curFracPayD & ", " & curstamp & ",'" & r![WSTATUS] & "'," & DeleteRow & "," & bytUserID & "," & getComVolID & ",'" & Date & "',GETDATE());"
                    MainComm.Execute
                End If
                r.Close
            Next intI
        MainConn.CommitTrans
        MsgBox "Bonus Update Successfully", vbInformation, cnstMsgInfo

        Screen.MousePointer = vbDefault

        CodeList.Clear
        txtTotal.Text = ""
        Set r = Nothing
Exit Sub
x:
Screen.MousePointer = vbDefault
MainConn.RollbackTrans
Set r = Nothing
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
    MDIfrmMain.lblCaption.Caption = "Festival Bonus Entry"
    If boJRMgt = True Then MDIfrmMain.lblCaption.Caption = MDIfrmMain.lblCaption.Caption + " - JR"

    prcMonthName cmbMonthName
    cmbMonthName.Text = UCase(Format(Date, "MMMM"))
    Call DeptName(cmbDepartment)
    
    If strUserlogvar <> cnstUserlogvar Then
        cmdDeleteLeftData.Visible = False
        cmdReportForm.Visible = False
    Else
        cmdDeleteLeftData.Visible = True
        cmdReportForm.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    boJRMgt = False
    MDIfrmMain.lblCaption.Caption = ""
    
    Set frmBonusEntryNew = Nothing
End Sub

Private Sub txtHeader_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtHeader_LostFocus()
    txtHeader.Text = UCase(txtHeader.Text)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate.Text = FormatDate(txtToDate)
End Sub

Private Function fncBlank(param) As Boolean
fncBlank = True
    If param = 1 Then   ''1 execute when try to Add to list
        If Len(cmbDepartment.Text) = 0 Then
            MsgBox "You should select department name.", vbCritical, App.Title
            cmbDepartment.SetFocus
        Exit Function
        End If
    
        If Len(txtToDate.Text) = 0 Then
            MsgBox "You should provide to joining date.", vbCritical, App.Title
            txtToDate.SetFocus
        Exit Function
        End If
    ElseIf param = 2 Then   ''2 execute when try to Update Data
        If Len(txtHeader.Text) = 0 Then
            MsgBox "You should provide bonus name.", vbCritical, App.Title
            txtHeader.SetFocus
        Exit Function
        End If
    
        If Len(cmbMonthName.Text) = 0 Then
            MsgBox "You should select payment month.", vbCritical, App.Title
            cmbMonthName.SetFocus
        Exit Function
        End If
    End If
fncBlank = False
End Function

