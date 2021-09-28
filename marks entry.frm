VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form marksentry 
   BackColor       =   &H00C0E0FF&
   Caption         =   "MARKS ENTRY FORM"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "marks entry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   1335
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   6255
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   600
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo RollNoList 
         Bindings        =   "marks entry.frx":27A2
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ROLLNO"
         Text            =   ""
         Object.DataMember      =   "StudentInfo"
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ROLL NO."
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "NAME "
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   17
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "&Update"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MARKS ENTRY "
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   6255
      Begin MSDataListLib.DataCombo RollNo 
         Bindings        =   "marks entry.frx":27BB
         DataField       =   "ROLLNO"
         DataSource      =   "CollegeADODC"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ROLLNO"
         Text            =   ""
         Object.DataMember      =   "StudentInfo"
      End
      Begin MSDataListLib.DataCombo SubjectCode 
         Bindings        =   "marks entry.frx":27D4
         DataField       =   "SUBJECTCODE"
         DataSource      =   "CollegeADODC"
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SUBJECTCODE"
         Text            =   ""
         Object.DataMember      =   "SubjectInfo"
      End
      Begin VB.ComboBox cmbicode 
         DataField       =   "units"
         DataSource      =   "CollegeADODC"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "marks entry.frx":27ED
         Left            =   5040
         List            =   "marks entry.frx":27FA
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkinternal 
         BackColor       =   &H00C0E0FF&
         Caption         =   "INTERNAL"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3885
         TabIndex        =   19
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox maxmarktxt 
         DataField       =   "marks"
         DataSource      =   "CollegeADODC"
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   3720
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "examdate"
         DataSource      =   "CollegeADODC"
         Height          =   420
         Left            =   2520
         TabIndex        =   11
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   741
         _Version        =   393216
         Format          =   24510464
         CurrentDate     =   37588
      End
      Begin VB.ComboBox cmbunit 
         DataField       =   "UNITS"
         DataSource      =   "CollegeADODC"
         Height          =   315
         ItemData        =   "marks entry.frx":280A
         Left            =   2520
         List            =   "marks entry.frx":2820
         TabIndex        =   9
         Text            =   "UNIT "
         Top             =   2625
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "marks entry.frx":283C
         DataField       =   "CHECKBY"
         DataSource      =   "CollegeADODC"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "EMPCODE"
         Text            =   ""
         Object.DataMember      =   "TeacherInfo"
      End
      Begin MSDataListLib.DataCombo CourceCode 
         Bindings        =   "marks entry.frx":2855
         DataField       =   "COURCECODE"
         DataSource      =   "CollegeADODC"
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "COURCECODE"
         BoundColumn     =   "COURCECODE"
         Text            =   ""
         Object.DataMember      =   "CourceInfo"
      End
      Begin VB.Label lblId 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RESULTID"
         DataSource      =   "CollegeADODC"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Roll No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "TEACHER CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Marks "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "TEST DATE "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3120
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "UNIT CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "SUBJECT CODE "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "COURSE CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   720
         Width           =   1500
      End
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   495
      Left            =   4680
      Top             =   6120
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
      RecordSource    =   "SELECT * FROM Result"
      Caption         =   "Result"
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
End
Attribute VB_Name = "marksentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkinternal_Click()
    If chkinternal.Value = 1 Then
        cmbicode.Enabled = True
        cmbunit.Enabled = False
    Else
        cmbicode.Enabled = False
        cmbunit.Enabled = True
    End If
End Sub
   
Private Sub cmdCancel_Click()
    Unload Me
    
    College.Show
End Sub

Private Sub cmddone_Click()
    CollegeADODC.Recordset.Update
    MsgBox " MARKS ENTRY PROCESS COMPELETE ", vbInformation, "Record Save"
End Sub

Private Sub cmdNew_Click()
    With CollegeADODC.Recordset
        .MoveLast
        Dim md As Long
        md = .Fields("ResultId") + 1
        .AddNew
        lblId.Caption = md
    End With
    txtName.Text = ""
End Sub

Private Sub cmdSearch_Click()
    
    If RollNoList.Text = "" Then
        MsgBox "Please, select RollNo", vbCritical, "Roll No Error"
        Exit Sub
    End If
    With envCollege.rsStudentInfo
        .Find "RollNO=" & RollNoList.Text, 0, adSearchForward, 1
        txtName.Text = .Fields("LastName") & " " & .Fields("Firstname") & " " & .Fields("ParentName")
    End With
End Sub

Private Sub Form_Activate()
    cmdNew_Click
    
End Sub

Private Sub Form_Load()
    
    DTPicker1.Value = Date
    
End Sub

Private Sub RollNo_Change()
    ShowName
End Sub

Private Sub RollNo_Click(Area As Integer)
    ShowName
End Sub

Public Sub ShowName()
On Error Resume Next
    With envCollege.rsStudentInfo
        .Find "RollNO=" & RollNo.Text, 0, adSearchForward, 1
        txtName.Text = .Fields("LastName") & " " & .Fields("Firstname") & " " & .Fields("ParentName")
    End With
End Sub
