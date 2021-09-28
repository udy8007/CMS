VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Height          =   975
      Left            =   4560
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Default         =   -1  'True
      Height          =   975
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
      Begin VB.OptionButton optReport 
         Caption         =   "&Detail Report"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton optReport 
         Caption         =   "&Vendorand Item Report"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Student &Attendance"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo RollNo 
         Bindings        =   "frmReport.frx":27A2
         Height          =   330
         Left            =   7200
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "ROLLNO"
         Text            =   ""
         Object.DataMember      =   "StudentInfo"
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Staff At&tendance Report "
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2355
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Student &Result"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   4
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Student &Information"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo EmpCode 
         Bindings        =   "frmReport.frx":27BB
         Height          =   330
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "EMPCODE"
         Text            =   ""
         Object.DataMember      =   "TeacherInfo"
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   735
         Left            =   120
         Top             =   2400
         Width           =   3900
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   735
         Left            =   120
         Top             =   480
         Width           =   3900
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   735
         Left            =   120
         Top             =   1440
         Width           =   3900
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "frmReport.frx":27D4
      Top             =   0
      Width           =   7350
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mIndex As Integer

Private Sub cmdReturn_Click()
    Unload Me
    College.Show
End Sub

Private Sub cmdShow_Click()
    On Error GoTo ShowError
    If mIndex >= 0 And mIndex <= 2 And Val(RollNo.Text) = 0 Then
        MsgBox "Please, Select RollNo", vbExclamation, "Selection Error"
        Exit Sub
    ElseIf mIndex = 3 And Val(EmpCode.Text) = 0 Then
        MsgBox "Please, Select Employee Code", vbExclamation, "Employee Code Selection Error"
        Exit Sub
    End If
    If mIndex >= 0 And mIndex <= 3 Then Unload envCollege
    'If mIndex <> 4 Then Unload envCollege
    Select Case mIndex
        Case 0
            
            envCollege.Student_Grouping RollNo.Text
            StudentInfo.Show
        Case 1
            
            envCollege.StudentAttendance_Grouping RollNo.Text
            StudentAttendance.Show
        Case 2
            
            envCollege.StudentResult RollNo.Text
            StudentResult.Show
        Case 3
            
            envCollege.Staff_Grouping EmpCode.Text
            StaffAttendance.Show
        Case 4
            VendorItems.Show
        Case 5
            Unload Me
            frmDetailReportOptions.Show
    End Select
    Exit Sub
ShowError:
    MsgBox Err.Description, vbCritical, "Report Error"
End Sub

Private Sub Form_Load()
    optReport(5).Value = True
End Sub

Private Sub optReport_Click(Index As Integer)
    mIndex = Index
End Sub
