VERSION 5.00
Begin VB.Form adduserfrm 
   BackColor       =   &H00000000&
   Caption         =   "Add New User"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "adduserfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbdesg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "adduserfrm.frx":000C
      Left            =   5385
      List            =   "adduserfrm.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4740
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
      Left            =   5385
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "phone number"
      Top             =   2610
      Width           =   2415
   End
   Begin VB.TextBox cpasstxt 
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
      Left            =   5385
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "phone number"
      Top             =   4020
      Width           =   2415
   End
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
      Left            =   5385
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "phone number"
      Top             =   3300
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
      Height          =   810
      Left            =   6300
      Picture         =   "adduserfrm.frx":0053
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6180
      Width           =   1441
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   810
      Left            =   4290
      Picture         =   "adduserfrm.frx":0495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6180
      Width           =   1441
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
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
      Index           =   0
      Left            =   3075
      TabIndex        =   11
      Top             =   4740
      Width           =   2055
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
      Left            =   3075
      TabIndex        =   10
      Top             =   3300
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      Height          =   375
      Index           =   1
      Left            =   3075
      TabIndex        =   9
      Top             =   4020
      Width           =   2055
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
      Left            =   3075
      TabIndex        =   8
      Top             =   2580
      Width           =   2055
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3630
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
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
      Top             =   480
      Width           =   10995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add New User"
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
      Top             =   1170
      Width           =   6375
   End
End
Attribute VB_Name = "adduserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset

Private Sub cmdcan_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()

Set rs = New adodb.Recordset
rs.Open "select count(*) from login where username = '" & usetxt.Text & "'", adodc, adOpenKeyset, adLockOptimistic
If rs.Fields(0) = 1 Then
    MsgBox "User Name already Exist . . . !!!", vbCritical, "Invalid User Name"
    usetxt.SetFocus
Else
    If passtxt.Text = cpasstxt.Text Then
        Dim inqry As String
        Set rs = New adodb.Recordset
        rs.Open "select * from login", adodc, adOpenKeyset, adLockOptimistic
       inqry = "insert into login values('" & usetxt.Text & "','" & passtxt.Text & "','" & cmbdesg.Text & "')"
       
       adodc.Execute (inqry)
        MsgBox "User Created . . . !!!", vbInformation, "User Create"
        Unload Me
    Else
        MsgBox "Check Your Password . . . !!!", vbCritical, "Password"
        passtxt.Text = ""
        cpasstxt.Text = ""
        passtxt.SetFocus
    End If
End If

End Sub

Private Sub Form_Load()
Call connection

End Sub

Private Sub usetxt_LostFocus()
    Set rs = New adodb.Recordset
   Dim sql As String
   sql = "select count(userNAME) from login where userNAME = '" & usetxt.Text & "'"
  
    rs.Open sql, adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
    
    Else
        MsgBox "Username already exits . . . !!!", vbCritical, "Username"
        usetxt.SetFocus
        usetxt.SelStart = 0
        usetxt.SelLength = Len(usetxt.Text)
    End If
    
End Sub
