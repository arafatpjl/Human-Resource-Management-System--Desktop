VERSION 5.00
Begin VB.Form frmCreateCompany 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6487
   ClientLeft      =   2210
   ClientTop       =   1794
   ClientWidth     =   7384
   ControlBox      =   0   'False
   Icon            =   "frmCreateCompany.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6487
   ScaleWidth      =   7384
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00800080&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "Clo&se"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   415
      Left            =   1920
      TabIndex        =   11
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Create  Company"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   415
      Left            =   4080
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bonus Principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   5145
      TabIndex        =   29
      Top             =   2880
      Width           =   1950
      Begin VB.TextBox txtBonus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   5
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkBonus 
         Caption         =   " Manual"
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
         Left            =   480
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Select   Month"
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
         Index           =   19
         Left            =   480
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Index           =   14
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Basic of"
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
         Index           =   18
         Left            =   1080
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   16
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Salary Principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   360
      TabIndex        =   19
      Top             =   2880
      Width           =   4815
      Begin VB.TextBox txtPF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkPF 
         Caption         =   "Manual"
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
         Left            =   3720
         TabIndex        =   41
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkFoodPay 
         Caption         =   "Manual"
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
         Left            =   3720
         TabIndex        =   39
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkHRent 
         Caption         =   "Manual"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkBasic 
         Caption         =   "Manual"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtFoodPay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtHRent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtBasic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Pro. Fund"
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
         Index           =   3
         Left            =   480
         TabIndex        =   45
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "%   of Basic"
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
         Left            =   2160
         TabIndex        =   44
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   42
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Index           =   17
         Left            =   3360
         TabIndex        =   38
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Index           =   15
         Left            =   3360
         TabIndex        =   36
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Or"
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
         Index           =   13
         Left            =   3360
         TabIndex        =   32
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   12
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Tk   Per Day"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Food Pay @"
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
         Index           =   11
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "**"
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
         Height          =   135
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "%   of Basic"
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
         Left            =   2160
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "House Rent"
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
         Index           =   8
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "%   of Gross"
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
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Basic"
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
         Index           =   7
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set Company Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   3840
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtVerrify 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Verrify"
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
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   615
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
         Left            =   1200
         TabIndex        =   16
         Top             =   360
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
         Index           =   6
         Left            =   1200
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.TextBox txtAddress 
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
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   360
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   360
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   50
      Top             =   5520
      Width           =   7455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Create Company"
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
      Left            =   270
      TabIndex        =   49
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label6 
      Caption         =   "Data Year"
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
      Index           =   20
      Left            =   360
      TabIndex        =   47
      Top             =   5040
      Width           =   855
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
      Left            =   1200
      TabIndex        =   46
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Address "
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
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Company Name"
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
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmCreateCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim V_Password As String
    Dim DBID As Byte
    
    Dim ManualBasic As Byte, ManualHRent As Byte, ManualFoodPay As Byte
    Dim ManualBonus As Byte, ManualPF As Byte
    
    Dim V_rateBasic As Single, V_rateHRent As Single, V_rateFoodPay As Single
    Dim V_rateBonus As Single, V_ratePF As Single
    
    If Len(txtName) = 0 Then
        MsgBox "Type Name of Company", vbInformation, cnstMsgInfo
        txtName.SetFocus
    Exit Sub
    End If
    
    If Len(txtAddress) = 0 Then
        MsgBox "Type Company Address", vbInformation, cnstMsgInfo
        txtAddress.SetFocus
    Exit Sub
    End If
    
    If Len(txtYear) = 0 Then
        MsgBox "Type Data Year", vbInformation, cnstMsgInfo
        txtYear.SetFocus
    Exit Sub
    End If
    
    If Len(txtPassword) = 0 Then
        If MsgBox("You don't set any Password." + vbCrLf + "Are you sure to continue !", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
            txtPassword.SetFocus
        Exit Sub
        End If
    Else
        If txtVerrify.Text <> txtPassword.Text Then
            MsgBox "Invalid Varrify Password", vbInformation, cnstMsgInfo
            txtVerrify.SetFocus
        Exit Sub
        End If
    End If

    If Len(txtPassword) <> 0 Then V_Password = EncryptIt(txtPassword.Text, 11) Else V_Password = ""
    
    If Len(txtBasic) = 0 Then V_rateBasic = 0 Else V_rateBasic = CSng(txtBasic.Text)
    If chkBasic.Value = vbChecked Then ManualBasic = 1 Else ManualBasic = 0
    
    If Len(txtHRent) = 0 Then V_rateHRent = 0 Else V_rateHRent = CSng(txtHRent.Text)
    If chkHRent.Value = vbChecked Then ManualHRent = 1 Else ManualHRent = 0
    
    If Len(txtFoodPay) = 0 Then V_rateFoodPay = 0 Else V_rateFoodPay = CSng(txtFoodPay.Text)
    If chkFoodPay.Value = vbChecked Then ManualFoodPay = 1 Else ManualFoodPay = 0
    
    If Len(txtPF) = 0 Then V_ratePF = 0 Else V_ratePF = CSng(txtPF.Text)
    If chkPF.Value = vbChecked Then ManualPF = 1 Else ManualPF = 0
    
    If Len(txtBonus) = 0 Then V_rateBonus = 0 Else V_rateBonus = CSng(txtBonus.Text)
    If chkBonus.Value = vbChecked Then ManualBonus = 1 Else ManualBonus = 0

    PeriodYear = txtYear.Text

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set r = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500

    DBID = 1
    StrRecord = "SELECT Max(ID) AS LastID FROM Company_Information HAVING ((Max(ID) Is Not Null));"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then DBID = DBID + r![LastID]
    r.Close

    MainConn.BeginTrans
        MainComm.CommandText = "SET NOCOUNT ON INSERT INTO Company_Information(ID,Name,Address,[Password],Basic,BasicM,HRent,HRentM,FoodPay,FoodPayM,PFund,PFundM,Bonus,BonusM,NZN)" _
              & " VALUES(" & DBID & ",'" & txtName.Text & "','" & txtAddress.Text & "','" & V_Password & "'," & V_rateBasic & "," & ManualBasic & "," & V_rateHRent & "," _
              & " " & ManualHRent & "," & V_rateFoodPay & "," & ManualFoodPay & "," & V_ratePF & "," & ManualPF & "," & V_rateBonus & "," & ManualBonus & ",0);"
        MainComm.Execute
    MainConn.CommitTrans

    MDIfrmMain.lblCompanyName.Caption = txtName.Text & " (" & txtYear.Text & ")"
    MDIfrmMain.lblName.Caption = DBID
    
    MDIfrmMain.mnucomInfor.Visible = False
    MDIfrmMain.mnuDataEntry.Visible = True
    MDIfrmMain.mnuDataEdit.Visible = True
    MDIfrmMain.mnuReport.Visible = True
    MDIfrmMain.mnuTools.Visible = True
'    MDIfrmMain.mnuCloseCompany.Visible = True
'    MDIfrmMain.mnuNext58.Visible = True

    Unload Me

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Err.Number = 0
    Exit Sub
    End If

    Screen.MousePointer = vbDefault
    Set r = Nothing
End Sub

Private Sub Form_Load()
    MDIfrmMain.lblCaption.Caption = "Create Company"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIfrmMain.lblCaption.Caption = ""
    Set frmCreateCompany = Nothing
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.Text = Trim(UCase(txtAddress.Text))
End Sub

Private Sub txtBasic_GotFocus()
    Call fncGotFocus(txtBasic)
End Sub

Private Sub txtBasic_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtFoodPay_GotFocus()
    Call fncGotFocus(txtFoodPay)
End Sub

Private Sub txtFoodPay_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtHRent_GotFocus()
    Call fncGotFocus(txtHRent)
End Sub

Private Sub txtHRent_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtName_GotFocus()
    Call fncGotFocus(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtName_LostFocus()
    If Len(txtName) = 0 Then Exit Sub
    txtName.Text = Trim(UCase(txtName.Text))

    On Error GoTo ErrorHandler

    Set r = New ADODB.Recordset
    StrRecord = "SELECT Name FROM Company_Information WHERE Name='" & txtName.Text & "';"
    r.Open StrRecord, MainConn, adOpenStatic
    If r.EOF = False And r.BOF = False Then
        MsgBox "Duplicate Name of Company", vbInformation, cnstMsgInfo
        txtName.SetFocus
    Exit Sub
    End If
    r.Close

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Err.Number = 0
    End
    End If
    Set r = Nothing
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.Text = Trim(txtPassword.Text)
End Sub

Private Sub txtPF_GotFocus()
    Call fncGotFocus(txtPF)
End Sub

Private Sub txtPF_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtVerrify_GotFocus()
    Call fncGotFocus(txtVerrify)
End Sub

Private Sub txtVerrify_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVerrify_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtVerrify_LostFocus()
    If Len(txtPassword) = 0 And Len(txtVerrify) <> 0 Then
        MsgBox "Type Password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
    Else
        txtVerrify.Text = Trim(txtVerrify.Text)
    End If
End Sub

Private Sub txtYear_GotFocus()
    Call fncGotFocus(txtYear)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtYear_LostFocus()
    If Len(txtYear) = 0 Then Exit Sub
    txtYear.Text = Trim(txtYear.Text)
''    If CInt(txtYear.Text) < 2001 Or CInt(txtYear.Text) > 2015 Then
''        MsgBox "Data Year must be between '2001' and '2015'", vbInformation, cnstMsgInfo
''        txtYear.SetFocus
''    End If
End Sub
