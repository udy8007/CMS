VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReport3 
   BackColor       =   &H80000012&
   Caption         =   "Student Attendance Report"
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
   Icon            =   "frmReport3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Left            =   4860
      Picture         =   "frmReport3.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
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
      Left            =   1470
      Picture         =   "frmReport3.frx":306C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
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
      Left            =   8280
      Picture         =   "frmReport3.frx":3936
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
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
      Left            =   6570
      Picture         =   "frmReport3.frx":3EB2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
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
      Left            =   3150
      Picture         =   "frmReport3.frx":42F4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1441
   End
   Begin MSComCtl2.DTPicker StartDate 
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      Top             =   3030
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   635
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   49545219
      CurrentDate     =   37680
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
      ItemData        =   "frmReport3.frx":45FE
      Left            =   2520
      List            =   "frmReport3.frx":460B
      TabIndex        =   1
      Top             =   2580
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmReport3.frx":4618
      Height          =   3105
      Left            =   1553
      TabIndex        =   6
      Top             =   4080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5477
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "PRESENTDAYS"
         Caption         =   "Present Day"
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
         DataField       =   "ABSCENTDAYS"
         Caption         =   "Absent Day"
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
         DataField       =   "TOTAL"
         Caption         =   "Total"
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
         DataField       =   "AVERAGEPRESENT"
         Caption         =   "Average Present"
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
         DataField       =   "YEAR"
         Caption         =   "Year"
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
         DataField       =   "ATTENDANCEDATE"
         Caption         =   "Attendence Date"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
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
      Left            =   2520
      TabIndex        =   0
      Top             =   2115
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmReport3.frx":4633
      Height          =   360
      Left            =   5910
      TabIndex        =   3
      Top             =   2115
      Width           =   2145
      _ExtentX        =   3784
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
      Left            =   120
      Top             =   5940
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
      RecordSource    =   "SELECT * FROM attendinfo"
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
      Bindings        =   "frmReport3.frx":464C
      Height          =   360
      Left            =   5910
      TabIndex        =   4
      Top             =   2580
      Width           =   2145
      _ExtentX        =   3784
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
      Left            =   330
      Top             =   6600
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
      RecordSource    =   "SELECT * FROM attendinfo"
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
   Begin MSComCtl2.DTPicker EndDate 
      Height          =   360
      Left            =   5910
      TabIndex        =   5
      Top             =   3030
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   635
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   49545219
      CurrentDate     =   37680
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3165
      Left            =   1523
      Top             =   4050
      Width           =   8835
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1605
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   1980
      Width           =   10995
   End
   Begin VB.Label Label8 
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
      Left            =   443
      TabIndex        =   19
      Top             =   480
      Width           =   10995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   443
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   10995
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attedence Entry"
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
      TabIndex        =   18
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3930
      TabIndex        =   17
      Top             =   3030
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Left            =   840
      TabIndex        =   16
      Top             =   3030
      Width           =   1455
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
      Left            =   840
      TabIndex        =   15
      Top             =   2580
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
      Left            =   4290
      TabIndex        =   14
      Top             =   2580
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
      Left            =   4290
      TabIndex        =   13
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No :"
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
      Left            =   840
      TabIndex        =   12
      Top             =   2115
      Width           =   1455
   End
End
Attribute VB_Name = "frmReport3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearAll_Click()
    txtRollNo.Text = ""
    cboYear.Text = ""
    SubjectCode.Text = ""
    CourceCode.Text = ""
End Sub

Private Sub cmdReport_Click()
    If txtRollNo.Text = "" And CourceCode.Text = "" And SubjectCode.Text = "" And cboYear.Text = "" And StartDate.Value = Null And EndDate.Value = Null Then
        cmdShowAll_Click
    Else
        cmdSearch_Click
    End If
    If CollegeADODC.Recordset.RecordCount = 0 Then
        MsgBox "No Students in this Criteria . . . !!!", vbExclamation, "No Students"
    Else
        Report3.Show
    End If

End Sub

Private Sub cmdreturn_Click()
    Unload Me
    'frmDetailReportOptions.Show
'    query.Show
End Sub


Private Sub cmdSearch_Click()
    ShowData
End Sub

Private Sub cmdShowAll_Click()
    cmdClearAll_Click
    Dim sql As String
    sql = "SELECT * FROM ATTENDINFO "
    CollegeADODC.RecordSource = sql
    CollegeADODC.Refresh
    MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Student's Information"

    Dim CN As New adodb.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT3 AS " & sql
    
End Sub

Public Sub ShowData()
On Error GoTo DataError
    Dim sql As String
    sql = "SELECT * from attendinfo "
    If txtRollNo.Text <> "" Then sql = sql & " WHERE ROLLNO=" & txtRollNo.Text
    If CourceCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        End If
    End If
    
    If cboYear.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND year =" & Val(cboYear.Text)
        Else
            sql = sql & " WHERE year=" & Val(cboYear.Text)
        End If
    End If
    
    If SubjectCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(subjectcode) LIKE '%" & UCase(SubjectCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(subjectcode) LIKE '%" & UCase(SubjectCode.Text) & "%' "
        End If
    End If
    
    If StartDate.Value = Null And EndDate.Value = Null Then
    Else
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND "
        Else
            sql = sql & " WHERE "
        End If
        sql = sql & " ATTENDANCEDATE BETWEEN '" & Format(StartDate.Value, "dd-MMM-yyyy") & "' AND '" & Format(EndDate.Value, "dd-MMM-yyyy") & "'"
    End If
    CollegeADODC.RecordSource = sql
    CollegeADODC.Refresh
    MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Attendance Information"
    
    Dim CN As New adodb.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT3 AS " & sql
    
    Exit Sub
DataError:
    MsgBox Err.Description, vbCritical, "Show Error"
End Sub


Private Sub Form_Load()
    StartDate.Value = Date
    EndDate.Value = Date
End Sub

Private Sub txtROLLno_GotFocus()
    txtRollNo.SelStart = 0
    txtRollNo.SelLength = Len(txtRollNo.Text)
End Sub
Private Sub txtROLLno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

