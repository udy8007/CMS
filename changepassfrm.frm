VERSION 5.00
Begin VB.Form changepassfrm 
   BackColor       =   &H00000000&
   Caption         =   "Change Password"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox usetxt 
      DataField       =   "C_PHONE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5865
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "phone number"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox oldtxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5880
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3150
      Width           =   2415
   End
   Begin VB.TextBox contxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5895
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4605
      Width           =   2415
   End
   Begin VB.TextBox newtxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5895
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3900
      Width           =   2415
   End
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6360
      Picture         =   "changepassfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5835
      Width           =   1440
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3870
      Picture         =   "changepassfrm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5850
      Width           =   1440
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   465
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   10995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sri Devi Arts and Science College"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   420
      TabIndex        =   11
      Top             =   480
      Width           =   10995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   2805
      TabIndex        =   10
      Top             =   1170
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Index           =   8
      Left            =   3210
      TabIndex        =   9
      Top             =   2400
      Width           =   2115
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Index           =   2
      Left            =   3210
      TabIndex        =   8
      Top             =   3855
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Left            =   3210
      TabIndex        =   7
      Top             =   3150
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Left            =   3210
      TabIndex        =   6
      Top             =   4605
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3540
      Left            =   443
      Shape           =   4  'Rounded Rectangle
      Top             =   1995
      Width           =   10995
   End
End
Attribute VB_Name = "changepassfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset

Private Sub cmdcan_Click()
Unload Me

End Sub

Private Sub cmdnew_Click()
Set rs = New adodb.Recordset
rs.Open "select count(username) from login where username = '" & usetxt.Text & "'", adodc, adOpenKeyset, adLockOptimistic
If rs.Fields(0) = 0 Then
    MsgBox "User Does not Exist . . . !!!", vbCritical, "User"
    usetxt.Text = ""
    oldtxt.Text = ""
    newtxt.Text = ""
    contxt.Text = ""
    usetxt.SetFocus
Else
    Set rs = New adodb.Recordset
    rs.Open "select * from login where username = '" & usetxt.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(1) = oldtxt.Text Then
        If newtxt.Text = contxt.Text Then
            Dim inqry As String
            inqry = "update login set password = '" & newtxt.Text & "' where username = '" & usetxt.Text & "'"
            adodc.Execute (inqry)
            MsgBox "Password Changed . . . !!!", vbInformation, "User Create"
            Unload Me
        Else
            MsgBox "Invalid Password . . . !!!", vbCritical, "Password"
            newtxt.Text = ""
            contxt.Text = ""
            newtxt.SetFocus
        End If
    Else
        MsgBox "Invalid Password . . . !!!", vbCritical, "Password"
        oldtxt.Text = ""
        oldtxt.SetFocus
    End If
    
End If
End Sub

Private Sub Form_Load()
Call connection
End Sub

