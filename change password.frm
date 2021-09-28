VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form changepass 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "change password.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtretype 
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
      IMEMode         =   3  'DISABLE
      Left            =   2865
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2400
      Width           =   2235
   End
   Begin VB.TextBox txtnewpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2865
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1980
      Width           =   2235
   End
   Begin VB.TextBox txtoldpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2865
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1560
      Width           =   2235
   End
   Begin VB.TextBox txtuserid 
      BackColor       =   &H00C0C0FF&
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
      Left            =   2865
      MaxLength       =   15
      TabIndex        =   9
      Top             =   1140
      Width           =   2235
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Return"
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
      Left            =   3960
      Picture         =   "change password.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3210
      Width           =   1441
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Change"
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
      Left            =   630
      Picture         =   "change password.frx":2D1E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3180
      Width           =   1441
   End
   Begin VB.CommandButton cmdrefresh 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Refresh"
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
      Left            =   2295
      Picture         =   "change password.frx":35E8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3180
      Width           =   1441
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   375
      Left            =   4320
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "ty02"
      Password        =   "ty02"
      RecordSource    =   "SELECT * FROM Login"
      Caption         =   "Attendance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Height          =   240
      Left            =   525
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Height          =   240
      Left            =   525
      TabIndex        =   7
      Top             =   1980
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Height          =   240
      Left            =   525
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
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
      Height          =   240
      Left            =   525
      TabIndex        =   5
      Top             =   1140
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   1845
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1050
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   690
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   555
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   5655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "A M C O S T "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   180
      Width           =   6015
   End
End
Attribute VB_Name = "changepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As ADODB.connection
Dim rs As ADODB.Recordset
Dim Str As String
Dim Qry As String

Private Sub cmdCancel_Click()
    Unload Me
    
    College.Show
End Sub

Private Sub cmdchange_Click()
    CollegeADODC.Recordset.Find "USERNAME='" & txtuserid.Text & "'", 0, adSearchForward, 1
    If CollegeADODC.Recordset.EOF = True Or CollegeADODC.Recordset.BOF = True Then
        MsgBox "Invalid userid and password", vbQuestion, "Invalid UserName/Password"
        Exit Sub
    End If
    If txtnewpass.Text = txtretype.Text Then
            Qry = "update login set password = '" & txtnewpass.Text & "' where userNAME = '" & txtuserid.Text & "'"
            Dim CN As New ADODB.connection
            CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
            CN.Execute Qry
            blankbox
            MsgBox "THE PASSWORD HAS BEEN CHANGE", vbInformation
    Else
            MsgBox "INVALID CONFIRMATION PASSWORD", vbExclamation
    End If
End Sub

Private Sub cmdrefresh_Click()
    CollegeADODC.Recordset.CancelUpdate
End Sub

Private Sub Form_Load()

Set rs = New ADODB.Recordset
Str = "select * from login"
rs.Open Str, CollegeADODC.ConnectionString, adOpenDynamic, adLockOptimistic
End Sub

Function blankbox()
    txtuserid.Text = ""
    txtnewpass.Text = ""
    txtoldpass.Text = ""
    txtretype.Text = ""
End Function

