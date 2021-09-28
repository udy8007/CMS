VERSION 5.00
Begin VB.Form staffreport 
   BackColor       =   &H80000007&
   Caption         =   "Staff Report"
   ClientHeight    =   3135
   ClientLeft      =   2610
   ClientTop       =   2850
   ClientWidth     =   7380
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   7380
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
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
      Left            =   2055
      Picture         =   "staffreport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2100
      Width           =   1441
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3885
      Picture         =   "staffreport.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2100
      Width           =   1441
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   1890
      TabIndex        =   4
      Top             =   1290
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "A M C O S T"
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
      Left            =   1575
      TabIndex        =   0
      Top             =   330
      Width           =   4245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   675
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   6795
   End
End
Attribute VB_Name = "staffreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim staffmas As ADODB.Recordset

Private Sub Command1_Click()
   If Combo1.Text = "" Then
        MsgBox "Select the value", vbCritical
        Combo1.SetFocus
        Exit Sub
   End If
   
   
   Unload envCollege
    
    envCollege.Staff_Grouping Combo1.Text
     Unload Me
'
''     DataReport1.Show
    StaffAttendance.Show
'
    
End Sub

Private Sub Command2_Click()
    Unload Me
    'query.Show
    tree_main.Show
End Sub

Private Sub Form_Load()
Call connection

Set staffmas = New ADODB.Recordset
staffmas.Open "select empcode from staffmaster order by empcode", adodc, adOpenKeyset, adLockBatchOptimistic

Do While Not staffmas.EOF
    Combo1.AddItem staffmas.Fields(0)
    staffmas.MoveNext
Loop

End Sub
