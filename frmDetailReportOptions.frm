VERSION 5.00
Begin VB.Form frmDetailReportOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Report Options"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetailReportOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4695
      Begin VB.OptionButton optOptions 
         Caption         =   "Query Regarding &Attendance of a Student"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Query Regarding &Marks of a Student"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "&Information Of  a Student"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please, select specified option && click on  OK button"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmDetailReportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mIndex As Integer
Private Sub cmdOk_Click()
    If mIndex >= 0 And mIndex <= 3 Then Unload Me
    Select Case mIndex
        Case 0
            frmReport1.Show
        Case 1
            frmReport2.Show
        Case 2
            frmReport3.Show
        Case 3
            
            frmStudentUpdate.Show
        Case Else
            MsgBox "Please, Select Your option", vbExclamation, "Option Select Error"
    End Select
End Sub

Private Sub cmdReturn_Click()
    Unload Me
    College.Show
End Sub

Private Sub optOptions_Click(Index As Integer)
    mIndex = Index
End Sub
