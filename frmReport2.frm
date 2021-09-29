VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReport2 
   BackColor       =   &H80000012&
   Caption         =   "Marks Report"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRollNo 
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
      Left            =   2550
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtTop 
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
      Left            =   3090
      TabIndex        =   3
      Top             =   3070
      Width           =   555
   End
   Begin VB.ComboBox cboYear 
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
      ItemData        =   "frmReport2.frx":27A2
      Left            =   2550
      List            =   "frmReport2.frx":27AF
      TabIndex        =   1
      Top             =   2600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Search"
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
      Left            =   5205
      Picture         =   "frmReport2.frx":27BC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7500
      Width           =   1441
   End
   Begin VB.CommandButton cmdShowAll 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show &All"
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
      Left            =   1815
      Picture         =   "frmReport2.frx":3086
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7500
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
      Left            =   8625
      Picture         =   "frmReport2.frx":3950
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7500
      Width           =   1441
   End
   Begin VB.CommandButton cmdClearAll 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear All"
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
      Left            =   6915
      Picture         =   "frmReport2.frx":3ECC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7500
      Width           =   1441
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "R&eport"
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
      Left            =   3495
      Picture         =   "frmReport2.frx":430E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7500
      Width           =   1441
   End
   Begin VB.CheckBox chkOrderMarks 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Order By Marks :"
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
      Left            =   3870
      TabIndex        =   8
      Top             =   3540
      Width           =   2025
   End
   Begin VB.OptionButton optTop 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Top :"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   3070
      Width           =   945
   End
   Begin VB.OptionButton optOverAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "OverAll List :"
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
      Left            =   1170
      TabIndex        =   4
      Top             =   3540
      Width           =   1695
   End
   Begin VB.ComboBox cboUnitCode 
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
      ItemData        =   "frmReport2.frx":4618
      Left            =   5700
      List            =   "frmReport2.frx":462E
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2600
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmReport2.frx":464A
      Height          =   2865
      Left            =   1193
      TabIndex        =   9
      Top             =   4320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5054
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16761024
      ForeColor       =   8388736
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Student's Information"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "COURCECODE"
         Caption         =   "Course Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EXAMDATE"
         Caption         =   "Exam Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "RESULTID"
         Caption         =   "Result ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ROLLNO"
         Caption         =   "Roll No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SUBJECTCODE"
         Caption         =   "Subject Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CHECKBY"
         Caption         =   "Check By"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "UNITS"
         Caption         =   "Unit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "MARKS"
         Caption         =   "Marks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "REMARK"
         Caption         =   "Remark"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmReport2.frx":4665
      Height          =   360
      Left            =   5700
      TabIndex        =   5
      Top             =   2130
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   495
      Left            =   30
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM RESULT"
      Caption         =   "Student"
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
   Begin MSDataListLib.DataCombo SubjectCode 
      Bindings        =   "frmReport2.frx":467E
      Height          =   360
      Left            =   5700
      TabIndex        =   7
      Top             =   3070
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "SUBJECTCODE"
      BoundColumn     =   "COURCECODE"
      Text            =   ""
      Object.DataMember      =   "SubjectInfo"
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
   Begin MSAdodcLib.Adodc Temp 
      Height          =   495
      Left            =   9270
      Top             =   3630
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM RESULT"
      Caption         =   "Student"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   2925
      Left            =   1163
      Top             =   4290
      Width           =   9555
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   2025
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   1950
      Width           =   10995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marks Report"
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
      Left            =   2768
      TabIndex        =   21
      Top             =   1170
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   443
      Shape           =   4  'Rounded Rectangle
      Top             =   210
      Width           =   10995
   End
   Begin VB.Label Label7 
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
      Height          =   855
      Left            =   450
      TabIndex        =   20
      Top             =   480
      Width           =   10995
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   780
      TabIndex        =   19
      Top             =   2600
      Width           =   1665
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Code :"
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
      Left            =   4020
      TabIndex        =   18
      Top             =   2600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
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
      Left            =   4020
      TabIndex        =   17
      Top             =   3075
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course :"
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
      Left            =   4020
      TabIndex        =   16
      Top             =   2130
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No. :"
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
      Left            =   780
      TabIndex        =   15
      Top             =   2130
      Width           =   1665
   End
End
Attribute VB_Name = "frmReport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String

Private Sub cmdClearAll_Click()
    txtTop.Text = ""
    txtRollNo.Text = ""
    
    cboUnitCode.Text = ""
    SubjectCode.Text = ""
    CourceCode.Text = ""
    optOverAll.Value = True
End Sub

Private Sub cmdReport_Click()
    
    If txtRollNo.Text = "" And CourceCode.Text = "" And cboUnitCode.Text = "" And SubjectCode.Text = "" And chkOrderMarks.Value = 0 Then
        cmdShowAll_Click
    Else
        cmdSearch_Click
    End If
    If CollegeADODC.Recordset.RecordCount = 0 Then
        MsgBox "No Students in this Criteria . . . !!!", vbExclamation, "No Students"
    Else
        Report2.Show
    End If
End Sub

Private Sub cmdreturn_Click()
    Unload Me
    'frmDetailReportOptions.Show
'    query.Show
End Sub

Private Sub cmdSearch_Click()
    If optOverAll.Value = False And optTop.Value = False Then
        MsgBox "Please, Select Option  . . . !!!", vbExclamation, "Option Select Error"
    ElseIf optOverAll.Value = True Then
        ShowRecords
    ElseIf optTop.Value = True Then
        If Val(txtTop.Text) = 0 Then
            MsgBox "Please, Fill Top Students Value . . . !!!", vbExclamation, "Student Select Error"
        Else
            TopStudents
        End If
    End If
End Sub

Private Sub cmdShowAll_Click()
    cmdClearAll_Click
    sql = "SELECT ROWNUM R, COURCECODE, EXAMDATE, RESULTID,    ROLLNO, SUBJECTCODE, CHECKBY, UNITS, MARKS,   REMARK From RESULT "
    OrderRecords
    
    CollegeADODC.Refresh
End Sub

Private Sub Form_Load()
    optOverAll.Value = True
End Sub


Public Sub ShowData()
    sql = "SELECT ROWNUM R, COURCECODE, EXAMDATE, RESULTID, ROLLNO, SUBJECTCODE, CHECKBY, UNITS, MARKS,   REMARK From RESULT"
    If txtRollNo.Text <> "" Then sql = sql & " WHERE ROLLNO=" & txtRollNo.Text
    If cboUnitCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(unitS) LIKE '%" & UCase(cboUnitCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(UNITS) LIKE '%" & UCase(cboUnitCode.Text) & "%' "
        End If
    End If
      If cboYear.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND YEAR =" & Val(cboYear.Text)
        Else
            sql = sql & " WHERE YEAR =" & Val(cboYear.Text)
        End If
    End If
    If SubjectCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(subjectcode) LIKE '%" & UCase(SubjectCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(subjectcode) LIKE '%" & UCase(SubjectCode.Text) & "%' "
        End If
    End If
     If CourceCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(COURCEcode) LIKE '%" & UCase(CourceCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(COURCEcode) LIKE '%" & UCase(CourceCode.Text) & "%' "
        End If
    End If
     
    OrderRecords
End Sub

Public Sub TopStudents()
On Error GoTo DataError
    ShowData
    OrderRecords
    Temp.RecordSource = sql
    Temp.Refresh
    
    Dim mRowNum() As Long, i As Integer
    Dim J As Integer, mStop As Integer
    mStop = IIf(Val(txtTop.Text) > Temp.Recordset.RecordCount, Temp.Recordset.RecordCount, Val(txtTop.Text))
    
    With Temp.Recordset
        Do Until .EOF
            i = i + 1
            If i > mStop Then Exit Do
            ReDim Preserve mRowNum(i)
            mRowNum(i) = .Fields("R")
            .MoveNext
        Loop
    End With
    sql = Trim(sql)
    If i <> mStop Then
        If chkOrderMarks.Value = 1 Then
            sql = left(sql, (InStr(1, sql, " ORDER BY")))
        End If
        
            If InStr(1, sql, "WHERE") > 0 Then
                sql = sql & " AND "
            Else
                sql = sql & " WHERE "
            End If
        sql = sql & " ROWNUM IN ("
        J = 1
        Do While J <= mStop
            If J > 1 Then sql = sql & ","
            sql = sql & mRowNum(J)
            J = J + 1
        Loop
        sql = sql & ") "
        OrderRecords
    End If
    
    CollegeADODC.RecordSource = sql
    CollegeADODC.Refresh
    MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Student's Marks Information"
         
         Dim CN As New adodb.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT2 AS " & sql

    Exit Sub
DataError:
    'MsgBox Err.Description, vbCritical, "Data Show Error"
End Sub


Public Sub ShowRecords()
On Error GoTo DataError
    ShowData
    
    Debug.Print sql
    OrderRecords
    CollegeADODC.Refresh
    MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Student's Marks Information"
         Dim CN As New adodb.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT2 AS " & sql

        Exit Sub
DataError:
    MsgBox Err.Description, vbCritical, "Data Show Error"
End Sub

Public Sub OrderRecords()
    If chkOrderMarks.Value = 1 And InStr(1, sql, " ORDER BY") = 0 Then
        sql = sql & " ORDER BY MARKS DESC"
    End If
    CollegeADODC.RecordSource = sql
    CollegeADODC.Refresh
End Sub

Private Sub optOverAll_Click()
    txtTop.Text = ""
End Sub

Private Sub txtTop_Change()
    optTop.Value = True
End Sub

Private Sub txtROLLno_GotFocus()
    txtRollNo.SelStart = 0
    txtRollNo.SelLength = Len(txtRollNo.Text)
End Sub
Private Sub txtROLLno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtTOP_GotFocus()
    txtRollNo.SelStart = 0
    txtRollNo.SelLength = Len(txtRollNo.Text)
End Sub
Private Sub txtTOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

