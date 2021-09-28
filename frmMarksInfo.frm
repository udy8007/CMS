VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMarksInfo 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marks Entry"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMarksInfo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox SubjectCode 
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
      Left            =   7425
      TabIndex        =   15
      Top             =   2700
      Width           =   1695
   End
   Begin VB.CommandButton cmdmarks 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Marks Entry"
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
      Left            =   4365
      Picture         =   "frmMarksInfo.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5010
      Width           =   1441
   End
   Begin VB.CommandButton cmdReturn 
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
      Left            =   6105
      Picture         =   "frmMarksInfo.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5010
      Width           =   1441
   End
   Begin VB.ComboBox cbounit 
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
      ItemData        =   "frmMarksInfo.frx":3028
      Left            =   4065
      List            =   "frmMarksInfo.frx":3047
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3270
      Width           =   1215
   End
   Begin VB.ComboBox cmbyear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "frmMarksInfo.frx":306F
      Left            =   4065
      List            =   "frmMarksInfo.frx":307C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2700
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker ExamDate 
      Bindings        =   "frmMarksInfo.frx":3089
      DataField       =   "TEACHER"
      DataMember      =   "ResultMarks"
      DataSource      =   "envCollege"
      Height          =   345
      Left            =   4065
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24576001
      CurrentDate     =   37588
   End
   Begin MSAdodcLib.Adodc Collegeadodc 
      Height          =   375
      Left            =   390
      Top             =   4500
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   "select * from result"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmMarksInfo.frx":309F
      DataField       =   "COURCECODE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   7425
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "COURCECODE"
      BoundColumn     =   "COURCECODE"
      Text            =   ""
      Object.DataMember      =   "CourceInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo teachercode 
      Bindings        =   "frmMarksInfo.frx":30B8
      DataField       =   "EMPCODE"
      DataMember      =   "TeacherInfo"
      DataSource      =   "envCollege"
      Height          =   360
      Left            =   7425
      TabIndex        =   9
      Top             =   3270
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "EMPCODE"
      BoundColumn     =   "EMPCODE"
      Text            =   ""
      Object.DataMember      =   "TeacherInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1875
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   1950
      Width           =   10995
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anand Mercantile College Of Science, Management and Computer Technology"
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
      Height          =   855
      Left            =   450
      TabIndex        =   12
      Top             =   390
      Width           =   10995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   10995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marks Entry"
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
      Height          =   735
      Left            =   2775
      TabIndex        =   11
      Top             =   1110
      Width           =   6375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2205
      TabIndex        =   6
      Top             =   2700
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5715
      TabIndex        =   5
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5715
      TabIndex        =   4
      Top             =   2730
      Width           =   1560
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5715
      TabIndex        =   3
      Top             =   3270
      Width           =   1560
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2205
      TabIndex        =   2
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2205
      TabIndex        =   1
      Top             =   3270
      Width           =   1560
   End
End
Attribute VB_Name = "frmMarksInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rs As adodb.Recordset
Private Sub cmdMarks_Click()
    If cmbyear.Text = "" Then
        MsgBox "Select then press ok  . . . !!!", vbInformation, "select properly"
        cmbyear.SetFocus
        Exit Sub
    End If
    If CourceCode.Text = "" Then
        MsgBox "Select then press ok . . . !!!", vbInformation, "select properly"
        CourceCode.SetFocus
        Exit Sub
    End If
    If SubjectCode.Text = "" Then
        MsgBox "Select then press ok . . . !!!", vbInformation, "select properly"
        SubjectCode.SetFocus
        Exit Sub
    End If
    If teachercode.Text = "" Then
        MsgBox "Select then press ok . . . !!!", vbInformation, "select properly"
        teachercode.SetFocus
        Exit Sub
    End If
    If cbounit.Text = "" Then
        MsgBox "Select then press ok . . . !!!", vbInformation, "select properly"
        cbounit.SetFocus
        Exit Sub
    End If
    mYear = cmbyear.Text
    mCourceCode = CourceCode.Text
    mTeacherCode = teachercode.Text
    mSubjectCode = SubjectCode.Text
    mUnits = cbounit.Text
    mExamDate = ExamDate

 
    Me.Hide
    Load frmMarksEntry
    frmMarksEntry.Show
End Sub
Private Sub cmdreturn_Click()
    Unload Me
    College.Show
End Sub

Private Sub CourceCode_lostFocus()
SubjectCode.Clear
Call connection
Set rs = New adodb.Recordset
rs.Open "select SUBJECTCODE from subjectmaster where COURCECODE = '" & CourceCode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    SubjectCode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
End Sub

Private Sub Form_Load()
    ExamDate = Date
End Sub
