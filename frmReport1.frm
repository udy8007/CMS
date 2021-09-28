VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReport1 
   BackColor       =   &H80000012&
   Caption         =   "Student detail Report"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmReport1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmReport1.frx":27A2
      Height          =   3315
      Left            =   1080
      TabIndex        =   13
      Top             =   3900
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5847
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   8388736
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   5
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "COURCECODE"
         Caption         =   "Course Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "JOINDATE"
         Caption         =   "Join Date"
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
         DataField       =   "LASTNAME"
         Caption         =   "Last Name"
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
         DataField       =   "FIRSTNAME"
         Caption         =   "First Name"
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
         DataField       =   "PARENTNAME"
         Caption         =   "Father's Name"
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
         DataField       =   "ADDRESS"
         Caption         =   "Address"
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
         DataField       =   "PHONE"
         Caption         =   "Phone"
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
      BeginProperty Column09 
         DataField       =   "PADDRESS"
         Caption         =   "Permenent Address"
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
      BeginProperty Column10 
         DataField       =   "PPHONE"
         Caption         =   "Phone"
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
      BeginProperty Column11 
         DataField       =   "GENDER"
         Caption         =   "Sex"
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
      BeginProperty Column12 
         DataField       =   "STREAM"
         Caption         =   "Stream"
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
      BeginProperty Column13 
         DataField       =   "BIRTHDATE"
         Caption         =   "Brith Date"
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
      BeginProperty Column14 
         DataField       =   "CAST"
         Caption         =   "Cast"
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
      BeginProperty Column15 
         DataField       =   "OCCUPATION"
         Caption         =   "Occupation"
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
      BeginProperty Column16 
         DataField       =   "ANNUALINCOME"
         Caption         =   "Annual Income"
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
      BeginProperty Column17 
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "Re&port"
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
      Picture         =   "frmReport1.frx":27BD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7410
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
      Picture         =   "frmReport1.frx":2AC7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7410
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
      Picture         =   "frmReport1.frx":2F09
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7410
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
      Picture         =   "frmReport1.frx":3485
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7410
      Width           =   1441
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
      Picture         =   "frmReport1.frx":3D4F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7410
      Width           =   1441
   End
   Begin VB.TextBox txtCity 
      Enabled         =   0   'False
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
      Left            =   7658
      TabIndex        =   14
      Top             =   2580
      Visible         =   0   'False
      Width           =   2445
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
      ItemData        =   "frmReport1.frx":4619
      Left            =   7658
      List            =   "frmReport1.frx":4626
      TabIndex        =   3
      Top             =   3060
      Width           =   1335
   End
   Begin VB.TextBox txtName 
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
      Left            =   3428
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2580
      Width           =   3135
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
      Height          =   360
      Left            =   3428
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2100
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmReport1.frx":4633
      Height          =   360
      Left            =   3450
      TabIndex        =   2
      Top             =   3090
      Width           =   3135
      _ExtentX        =   5530
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
      Left            =   90
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "SELECT * FROM StudentMaster"
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
      Height          =   3375
      Left            =   1050
      Top             =   3870
      Width           =   10035
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1635
      Left            =   510
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
      Left            =   480
      TabIndex        =   17
      Top             =   450
      Width           =   10995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   210
      Width           =   10995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Report"
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
      Left            =   2805
      TabIndex        =   16
      Top             =   1170
      Width           =   6375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      Enabled         =   0   'False
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
      Left            =   6218
      TabIndex        =   15
      Top             =   2580
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   6218
      TabIndex        =   12
      Top             =   3060
      Width           =   1215
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
      Left            =   1778
      TabIndex        =   11
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   1778
      TabIndex        =   10
      Top             =   2580
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
      Left            =   1778
      TabIndex        =   9
      Top             =   2100
      Width           =   1455
   End
End
Attribute VB_Name = "frmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearAll_Click()
    txtRollNo.Text = ""
    txtName.Text = ""
    txtCity.Text = ""
    CourceCode.Text = ""
    cboYear.Text = ""
End Sub

Private Sub cmdReport_Click()
    If txtRollNo.Text = "" And txtName.Text = "" And txtCity.Text = "" And CourceCode.Text = "" And cboYear.Text = "" Then
        cmdShowAll_Click
    Else
        cmdSearch_Click
    End If
     If CollegeADODC.Recordset.RecordCount = 0 Then
        MsgBox "No Students in this Criteria . . . !!!", vbExclamation, "No Students"
    Else
        Report1.Show
    End If

End Sub

Private Sub cmdreturn_Click()
    Unload Me
    'frmDetailReportOptions.Show
    'query.Show
End Sub

Private Sub cmdSearch_Click()
    Dim sql As String
    sql = "SELECT * FROM STUDENTMASTER"
    If txtRollNo.Text <> "" Then sql = sql & " WHERE ROLLNO=" & "'" & txtRollNo.Text & "'"
    
    If txtName.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(FIRSTNAME) LIKE '%" & UCase(txtName.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(FIRSTNAME) LIKE '%" & UCase(txtName.Text) & "%' "
        End If
    End If
    
    If txtCity.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(CITY) LIKE '%" & UCase(txtCity.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(CITY) LIKE '%" & UCase(txtCity.Text) & "%' "
        End If
    End If
    If CourceCode.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        Else
            sql = sql & " WHERE UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        End If
    End If
    If cboYear.Text <> "" Then
        If InStr(1, sql, "WHERE") > 0 Then
            sql = sql & " AND YEAR =" & Val(cboYear.Text)
        Else
            sql = sql & " WHERE YEAR =" & Val(cboYear.Text)
        End If
    End If
        
    Debug.Print sql
    CollegeADODC.RecordSource = sql
    
    CollegeADODC.Refresh
    MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Student's Information"
    Dim CN As New ADODB.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT1 AS " & sql
End Sub

Private Sub cmdShowAll_Click()
    cmdClearAll_Click
    Dim sql As String
    sql = "SELECT * FROM STUDENTMASTER "
            CollegeADODC.RecordSource = sql
        CollegeADODC.Refresh
        MsgBox "Total Records :  . . . " & CollegeADODC.Recordset.RecordCount, vbInformation, "Student's Information"
        Dim CN As New ADODB.connection
    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
    CN.Execute "CREATE OR REPLACE VIEW STUDENTREPORT1 AS " & sql
    
End Sub

Private Sub txtROLLno_GotFocus()
    txtRollNo.SelStart = 0
    txtRollNo.SelLength = Len(txtRollNo.Text)
End Sub
Private Sub txtROLLno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
        
End Sub

Private Sub txtname_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

