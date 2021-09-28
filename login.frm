VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2445
   ClientLeft      =   3795
   ClientTop       =   2790
   ClientWidth     =   6135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlogincancel 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   870
      Width           =   1215
   End
   Begin VB.TextBox txtpass 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1470
      Width           =   2175
   End
   Begin VB.TextBox txtuserid 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   0
      Top             =   870
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   4290
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   4035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sri Devi College"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   945
      TabIndex        =   6
      Top             =   90
      Width           =   4245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      ForeColor       =   &H00FF80FF&
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      ForeColor       =   &H00FF80FF&
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   870
      Width           =   1335
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As adodb.connection
Dim rs As adodb.Recordset
Dim Str As String, Counter As Integer
Dim Q As String
Dim Q1 As Integer


Private Sub Form_Load()
    Set Conn = New adodb.connection
    Set rs = New adodb.Recordset
End Sub

Private Sub cmdlogincancel_Click()
    Unload Me
End
End Sub

Private Sub cmdlogin_Click()
    
    Counter = Counter + 1
    If Counter > 3 Then
        MsgBox "Too may trys, Sorry............ ", vbExclamation, "Sorry..........."
        
        Unload Me
        Exit Sub
    End If
    Dim Qry As String
 
Conn.Open "Provider = sqloledb;" & _
          "Data Source=UDY8007\UDYSERVER;" & _
          "Initial Catalog=SDC;" & _
          "User ID=SDCUser;" & _
          "Password=SDCpwd;"
          
MsgBox "Welcome to SriDevi College Management System"
Qry = " select * from login where userNAME ='" & txtuserid.Text & "' and password = '" & txtpass.Text & "'"
rs.Open Qry, Conn, adOpenDynamic, adLockPessimistic
 

If rs.EOF = False Then
    If Len(txtpass.Text) <> Len(rs.Fields("password").Value) Then
        MsgBox "Invalid Userid or Password, Please retype . . . !!!", vbCritical, "Invalid UserName/Password"
    Else
        DESIG = rs.Fields(2)
        
        'College.Show
        tree_main.Show
        Unload Login
        
    End If
End If
If rs.EOF = True Then
    MsgBox "Invalid Userid or Password, Please retype . . . !!!", vbExclamation, "User/Password Error"
End If
rs.Close

Conn.Close
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdlogin_Click
End Sub
