VERSION 5.00
Begin VB.Form deluserfrm 
   BackColor       =   &H00000000&
   Caption         =   "Delete User"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox passtxt 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5445
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "phone number"
      Top             =   3210
      Width           =   2415
   End
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
      Height          =   375
      Left            =   5445
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "phone number"
      Top             =   2490
      Width           =   2415
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   3990
      Picture         =   "deleteuserfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4590
      Width           =   1440
   End
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   6120
      Picture         =   "deleteuserfrm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4605
      Width           =   1440
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   465
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   10995
   End
   Begin VB.Label Label1 
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
      TabIndex        =   7
      Top             =   510
      Width           =   10995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete User"
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
      Height          =   615
      Left            =   2805
      TabIndex        =   6
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Height          =   375
      Index           =   2
      Left            =   4005
      TabIndex        =   5
      Top             =   3225
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      Height          =   375
      Index           =   8
      Left            =   4005
      TabIndex        =   4
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1995
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   2070
      Width           =   10995
   End
End
Attribute VB_Name = "deluserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset

Private Sub cmdcan_Click()
Unload Me
End Sub

Private Sub cmddel_Click()
Set rs = New adodb.Recordset
rs.Open "select count(username) from login where username = '" & usetxt.Text & "'", adodc, adOpenKeyset, adLockPessimistic
If rs.Fields(0) = 1 Then
    Set rs = New adodb.Recordset
    rs.Open "select * from login where username = '" & usetxt.Text & "'", adodc, adOpenKeyset, adLockPessimistic
    If rs.Fields(1) = passtxt.Text Then
        Dim inqry As String
        inqry = "delete login where username = '" & usetxt.Text & "'"
        adodc.Execute (inqry)
        MsgBox "User Deleted . . . !!!", vbCritical, "Delete User"
        Unload Me
    Else
        MsgBox "Invalid password . . . !!!", vbCritical, "Invalid Password"
        passtxt.Text = ""
        passtxt.SetFocus
    End If
Else
    MsgBox "User Does not Exist . . . !!!", vbCritical, "User"
    usetxt.Text = ""
    passtxt.Text = ""
    usetxt.SetFocus
End If
End Sub

Private Sub Form_Load()
Call connection
End Sub
