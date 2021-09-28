VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form attendance 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "= ExamDate"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "attendance.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpresent 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Present"
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
      Left            =   8190
      Picture         =   "attendance.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7620
      Width           =   1441
   End
   Begin VB.CommandButton cmdShowAll 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Show All"
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
      Left            =   9900
      Picture         =   "attendance.frx":306C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7620
      Width           =   1441
   End
   Begin VB.CommandButton cmdabsent 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Absent"
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
      Left            =   6480
      Picture         =   "attendance.frx":3936
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7620
      Width           =   1441
   End
   Begin VB.CommandButton CMDDONE 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
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
      Left            =   1920
      Picture         =   "attendance.frx":4778
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6810
      Width           =   1441
   End
   Begin VB.CommandButton cmdCancelUpdate 
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
      Left            =   3900
      Picture         =   "attendance.frx":4A82
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7680
      Width           =   1441
   End
   Begin VB.CommandButton CMDCANCEL 
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
      Left            =   1200
      Picture         =   "attendance.frx":4EC4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7770
      Width           =   1441
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Add New"
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
      Left            =   180
      Picture         =   "attendance.frx":5440
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6810
      Width           =   1441
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Report"
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
      Left            =   3600
      Picture         =   "attendance.frx":5D0A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6810
      Width           =   1441
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   330
      Left            =   5670
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "SELECT * FROM Attendance "
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
   Begin MSAdodcLib.Adodc Temp 
      Height          =   330
      Left            =   4680
      Top             =   4170
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin VB.Frame Farame1 
      BackColor       =   &H00000000&
      Caption         =   "Attedence Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   4755
      Left            =   210
      TabIndex        =   5
      Top             =   1350
      Width           =   6195
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
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "attendance.frx":6014
         Left            =   900
         List            =   "attendance.frx":6021
         Sorted          =   -1  'True
         TabIndex        =   30
         Text            =   "1"
         Top             =   810
         Width           =   1215
      End
      Begin VB.ComboBox cboRollNo 
         DataField       =   "ROLLNO"
         DataSource      =   "CollegeADODC"
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
         ItemData        =   "attendance.frx":602E
         Left            =   3390
         List            =   "attendance.frx":6030
         TabIndex        =   1
         Top             =   810
         Width           =   1875
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid StudentInfo 
         Height          =   855
         Left            =   60
         TabIndex        =   14
         Top             =   1800
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   1508
         _Version        =   393216
         BackColor       =   16761024
         FixedCols       =   0
         BackColorFixed  =   16711680
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox Text1 
         DataField       =   "REMARK"
         DataSource      =   "CollegeADODC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3300
         Width           =   2625
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         Caption         =   "Present/Absent :"
         DataField       =   "PRESENT"
         DataSource      =   "CollegeADODC"
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
         Left            =   330
         TabIndex        =   2
         Top             =   3690
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "ATTENDANCEDATE"
         DataSource      =   "CollegeADODC"
         Height          =   360
         Left            =   2190
         TabIndex        =   3
         Top             =   4050
         Width           =   3885
         _ExtentX        =   6853
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
         CheckBox        =   -1  'True
         Format          =   24576000
         CurrentDate     =   37588
      End
      Begin MSDataListLib.DataCombo CourceCode 
         Bindings        =   "attendance.frx":6032
         DataField       =   "COURCECODE"
         DataSource      =   "CollegeADODC"
         Height          =   360
         Left            =   3390
         TabIndex        =   0
         Top             =   330
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "attendance.frx":604B
         DataField       =   "TEACHERCODE"
         DataSource      =   "CollegeADODC"
         Height          =   360
         Left            =   1740
         TabIndex        =   31
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "EMPCODE"
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
      Begin MSDataListLib.DataCombo SubjectCode 
         Bindings        =   "attendance.frx":6064
         DataField       =   "SUBJECTCODE"
         DataSource      =   "CollegeADODC"
         Height          =   360
         Left            =   1740
         TabIndex        =   32
         Top             =   3300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "SUBJECTCODE"
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
      Begin VB.Label Label11 
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
         Left            =   2160
         TabIndex        =   20
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         Left            =   30
         TabIndex        =   18
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Student's Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   60
         TabIndex        =   15
         Top             =   1440
         Width           =   6105
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Note :"
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
         Height          =   255
         Left            =   2370
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID :"
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
         Left            =   30
         TabIndex        =   11
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblId 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ATTENDANCEID"
         DataSource      =   "CollegeADODC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   930
         TabIndex        =   10
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         TabIndex        =   9
         Top             =   2880
         Width           =   1575
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
         Left            =   2160
         TabIndex        =   8
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         TabIndex        =   7
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Date :         "
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
         Left            =   210
         TabIndex        =   6
         Top             =   4080
         Width           =   1695
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid StudentList 
      Height          =   4815
      Left            =   6480
      TabIndex        =   16
      Top             =   1650
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8493
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10995
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   420
      TabIndex        =   19
      Top             =   750
      Width           =   6015
   End
   Begin VB.Label lblList 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Student's Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   1260
      Width           =   4935
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataField       =   "HOLIDAY"
      DataMember      =   "HolidayInfo"
      DataSource      =   "envCollege"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anand Mercantile College Of Science , Management and Computer Technology"
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
      Left            =   435
      TabIndex        =   29
      Top             =   360
      Width           =   10995
   End
End
Attribute VB_Name = "attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Str As String

Private Sub cmbyear_Click()
cmdShowAll.Enabled = True
End Sub

Private Sub cmbyear_Validate(Cancel As Boolean)
    If cmbyear.Text = "" Then
        MsgBox "Please Select Year", vbExclamation, "Year SElection Error"
        Cancel = True
        Exit Sub
    End If
    RollNoShow
End Sub

Private Sub cmdabsent_Click()
    ShowAbscent
End Sub

Private Sub cmdCancelUpdate_Click()
    CollegeADODC.Recordset.CancelUpdate
    CollegeADODC.Refresh
End Sub

Private Sub cmdNew_Click()
    Dim md As Long
    CollegeADODC.Recordset.MoveLast
    md = CollegeADODC.Recordset.Fields(0) + 1
    
    CollegeADODC.Recordset.AddNew
    lblId.Caption = md
    cmbyear.SetFocus
End Sub

Private Sub cmdpresent_Click()
    ShowPresent
End Sub

Private Sub cmdShowAll_Click()
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM STUDENTmaster", CollegeADODC.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    StudentList.ColWidth(0) = 2500
    StudentList.ColWidth(1) = 3500
    Set StudentList.DataSource = rs
    
    lblList.Caption = "Student List"
    rs.Close
End Sub

Private Sub CollegeADODC_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    StudentShow
    YearShow
End Sub

Private Sub CourceCode_Validate(Cancel As Boolean)
    Dim sql As String
    sql = Temp.RecordSource
        
    Temp.RecordSource = "SELECT * FROM STUDENTMASTER WHERE COURCECODE='" & CourceCode.Text & "'"
    Temp.Refresh
    cboRollNo.Clear
    Do Until Temp.Recordset.EOF
        cboRollNo.AddItem Temp.Recordset.Fields("ROLLNO")
        Temp.Recordset.MoveNext
    Loop
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
        With envCollege.rsHolidayInfo
        Do Until .EOF
            If .Fields("HolidayDate") = DTPicker1.Value Then
                MsgBox DTPicker1.Value & " is  HOLIDAY", vbQuestion, "Holiday!!!!!!!!!!"
                MsgBox "Please, Select Another Date", vbInformation, "Date SElect Error"
                Cancel = True
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
End Sub


Private Sub Form_Activate()
    cmdShowAll_Click
    
    StudentShow
    YearShow
End Sub

Private Sub Form_Load()
'With CollegeADODC.Recordset
'     .MoveFirst
'      While Not CollegeADODC.Recordset.EOF
'        CourceCode.ReFill (CollegeADODC.Recordset.Fields(0).Value)
'        CollegeADODC.Recordset.MoveNext
'      Wend
'   End With
DTPicker1.Value = Date
End Sub
Private Sub cmdCancel_Click()
    Unload Me
    College.Show
End Sub

Private Sub cmddone_Click()
    CollegeADODC.Recordset.Update
    MsgBox "ATTENDACE ENTRY OVER ...", vbInformation, "Over"
End Sub

Public Sub StudentShow()
    If cboRollNo.Text = "" Then Exit Sub
    Dim rs1 As New ADODB.Recordset
    rs1.Open "SELECT * FROM STUDENTmaster WHERE ROLLNO=" & cboRollNo.Text, CollegeADODC.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    StudentInfo.ColWidth(0) = 2500
    StudentInfo.ColWidth(1) = 3500
    Set StudentInfo.DataSource = rs1
    rs1.Close
End Sub

Private Sub RollNo_Click(Area As Integer)
    StudentShow
End Sub

Public Sub ShowAbscent()
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM STUDENTinfo WHERE present=0", CollegeADODC.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    StudentList.ColWidth(0) = 2500
    StudentList.ColWidth(1) = 3500
    Set StudentList.DataSource = rs
    
    lblList.Caption = "Abscent Student List"
    rs.Close
End Sub

Public Sub ShowPresent()
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM STUDENTinfo WHERE present=1", CollegeADODC.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    StudentList.ColWidth(0) = 2500
    StudentList.ColWidth(1) = 3500
     lblList.Caption = "Present Student List"
    Set StudentList.DataSource = rs
      
    rs.Close
End Sub


Public Sub RollNoShow()
 Dim sql As String
    If cmbyear.Text = "" Then Exit Sub
    sql = "SELECT * FROM STUDENTMASTER WHERE YEAR=" & cmbyear.Text
    Debug.Print sql
    Temp.RecordSource = sql
    Temp.Refresh
    If Temp.Recordset.RecordCount = 0 Then
        LockCtl
        Exit Sub
    Else
        UNLockCtl
    End If
    cboRollNo.Clear
    With Temp.Recordset
        Do Until .EOF
            cboRollNo.AddItem .Fields("RollNo")
            .MoveNext
        Loop
    End With
    cboRollNo.Text = CollegeADODC.Recordset.Fields("RollNo") & ""
End Sub

Public Sub LockCtl()
    Dim CTL As Control
    For Each CTL In Me.Controls
        If TypeOf CTL Is TextBox Or TypeOf CTL Is CheckBox _
            Or TypeOf CTL Is ComboBox Or TypeOf CTL Is DTPicker Or TypeOf CTL Is MSHFlexGrid _
           Or TypeOf CTL Is DataCombo Or TypeOf CTL Is CommandButton Then
            CTL.Enabled = False
        End If
    Next CTL
    cmbyear.Enabled = True
    lblMessage = "No Such Student"
End Sub

Public Sub UNLockCtl()
    Dim CTL As Control
    lblMessage = ""
    For Each CTL In Me.Controls
            CTL.Enabled = True
    Next CTL
End Sub

Public Sub YearShow()
    If cboRollNo.Text = "" Then Exit Sub
 Dim sql As String
    sql = "SELECT * FROM STUDENTMASTER WHERE RollNo=" & cboRollNo.Text
    Debug.Print sql
    Temp.RecordSource = sql
    Temp.Refresh
    If Temp.Recordset.RecordCount = 0 Then Exit Sub
    
    cmbyear.Text = Temp.Recordset.Fields("YEAR")
End Sub


