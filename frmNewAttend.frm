VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNewAttend 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attendance Entry"
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
   Icon            =   "frmNewAttend.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Height          =   960
      Left            =   6060
      Picture         =   "frmNewAttend.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1530
   End
   Begin VB.CommandButton cmdNewEntry 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Attendence Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4320
      Picture         =   "frmNewAttend.frx":2D1E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1530
   End
   Begin VB.ComboBox cmbyear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      ItemData        =   "frmNewAttend.frx":3028
      Left            =   5175
      List            =   "frmNewAttend.frx":3035
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "3"
      Top             =   2220
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   495
      Left            =   8010
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "ATTENDANCE"
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
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmNewAttend.frx":3042
      DataField       =   "COURCECODE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   5175
      TabIndex        =   1
      Top             =   2715
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmNewAttend.frx":305B
      DataField       =   "TEACHERCODE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   5175
      TabIndex        =   2
      Top             =   3210
      Width           =   1335
      _ExtentX        =   2355
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
      Bindings        =   "frmNewAttend.frx":3074
      DataField       =   "SUBJECTCODE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   5175
      TabIndex        =   3
      Top             =   3705
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "ATTENDANCEDATE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   5175
      TabIndex        =   4
      Top             =   4200
      Width           =   3495
      _ExtentX        =   6165
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
      Format          =   24510464
      CurrentDate     =   37588
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   2865
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
      Width           =   10995
   End
   Begin VB.Label Label6 
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
      Left            =   2745
      TabIndex        =   11
      Top             =   1140
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   180
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
      Left            =   420
      TabIndex        =   10
      Top             =   420
      Width           =   10995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Date :"
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
      Left            =   3075
      TabIndex        =   9
      Top             =   4200
      Width           =   1875
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
      Left            =   3075
      TabIndex        =   8
      Top             =   2715
      Width           =   1875
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
      Left            =   3075
      TabIndex        =   7
      Top             =   2220
      Width           =   1875
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   3105
      TabIndex        =   6
      Top             =   3210
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Code  :"
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
      Left            =   3075
      TabIndex        =   5
      Top             =   3705
      Width           =   1875
   End
End
Attribute VB_Name = "frmNewAttend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNewEntry_Click()
        
     tcode = DataCombo1.Text
     year = cmbyear
     attendancedate = DTPicker1.Value
     CCode = CourceCode.Text
     scode = SubjectCode.Text
     
'      With frmAttendance
'        .CollegeADODC.RecordSource = "select * from studentmaster where courcecode='" & CourceCode.Text & "' and year=" & cmbyear.Text
'        Debug.Print .CollegeADODC.RecordSource
'        .CollegeADODC.Refresh
'        .Temp.RecordSource = "select * from studentmaster where courcecode='" & CourceCode.Text & "' and year=" & cmbyear.Text
'        .Temp.Refresh
'
'        MySQL = "select * from studentmaster where courcecode='" & CourceCode.Text & "' and year=" & cmbyear.Text
'        AddRollNo
'        Unload Me
'    End With
'
    Unload Me
    frmAttendance.Show
End Sub

'Public Sub AddRollNo()
'On Error Resume Next
'With frmAttendance
'    .CollegeADODC.RecordSource = "select * from attendance "
'    .CollegeADODC.Refresh
'    .CollegeADODC.Recordset.MoveLast
'
'    Dim md As Long
'    md = .CollegeADODC.Recordset.Fields(0) + 1
'    .CollegeADODC.RecordSource = "select * from attendance " _
'        & " where courcecode='" & CourceCode.Text & "' and year=" & cmbyear.Text _
'        & " and subjectcode='" & SubjectCode.Text & "' and to_char(attendancedate,'DD/MM/YYYY')='" & Format(DTPicker1.Value, "DD/MM/YYYY") & "'"
'    .CollegeADODC.Refresh
'
'            .MoveFirst
'
'        Do Until .EOF
'
'            Debug.Print frmAttendance.CollegeADODC.RecordSource
'            frmAttendance.CollegeADODC.Recordset.Find "RollNo=" & .Fields("Rollno"), 0, adSearchForward, 1
'
'            If frmAttendance.CollegeADODC.Recordset.EOF Then frmAttendance.CollegeADODC.Recordset.AddNew
'
'            frmAttendance.CollegeADODC.Recordset.Fields("AttendanceId") = md
'            frmAttendance.CollegeADODC.Recordset.Fields("Rollno") = .Fields("Rollno")
'            frmAttendance.CollegeADODC.Recordset.Fields("TeacherCode") = DataCombo1.Text
'            frmAttendance.CollegeADODC.Recordset.Fields("SubjectCode") = SubjectCode.Text
'            frmAttendance.CollegeADODC.Recordset.Fields("AttendanceDate") = DTPicker1.Value
'            frmAttendance.CollegeADODC.Recordset.Fields("courcecode") = CourceCode.Text
'            frmAttendance.CollegeADODC.Recordset.Fields("year") = cmbyear.Text
'            frmAttendance.CollegeADODC.Recordset.Update
'            md = md + 1
'            .MoveNext
'        Loop
'        frmAttendance.CollegeADODC.RecordSource = "select * from attendance where teachercode='" & DataCombo1.Text & "' and subjectcode='" & SubjectCode.Text & "' and attendancedate='" & Format(DTPicker1.Value, "dd-MMM-YY") & "' AND COURCECODE='" & CourceCode.Text & "'"
'        frmAttendance.CollegeADODC.Refresh
'    End With
'    End With
'End Sub
'
'
'Private Sub Command1_Click()
'    Unload Me
'    College.Show
'End Sub

Private Sub Command1_Click()
    Unload Me
    College.Show
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
End Sub


