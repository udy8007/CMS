VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStudentUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Information Update Screen"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStudentUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtRollNo 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboYear 
      Height          =   330
      ItemData        =   "frmStudentUpdate.frx":27A2
      Left            =   1560
      List            =   "frmStudentUpdate.frx":27AF
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin MSDataListLib.DataCombo CourceCode 
      Bindings        =   "frmStudentUpdate.frx":27BC
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "COURCECODE"
      BoundColumn     =   "COURCECODE"
      Text            =   ""
      Object.DataMember      =   "CourceInfo"
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   375
      Left            =   960
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "&Name"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Student Update Information By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Roll No"
      Height          =   210
      Left            =   780
      TabIndex        =   8
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Cource"
      Height          =   210
      Left            =   765
      TabIndex        =   7
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Year"
      Height          =   210
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "frmStudentUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    frmDetailReportOptions.Show
End Sub

Private Sub cmdOk_Click()
    Dim SQL As String
    SQL = "SELECT * FROM STUDENTMASTER "
    If txtRollNo.Text <> "" Then SQL = SQL & " WHERE ROLLNO=" & txtRollNo.Text
    If txtName.Text <> "" Then
        If InStr(1, SQL, "WHERE") > 0 Then
            SQL = SQL & " AND UPPER(FIRSTNAME) LIKE '%" & UCase(txtName.Text) & "%' "
        Else
            SQL = SQL & " WHERE UPPER(FIRSTNAME) LIKE '%" & UCase(txtName.Text) & "%' "
        End If
    End If
    
    If CourceCode.Text <> "" Then
        If InStr(1, SQL, "WHERE") > 0 Then
            SQL = SQL & " AND UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        Else
            SQL = SQL & " WHERE UPPER(COURCECODE) LIKE '%" & UCase(CourceCode.Text) & "%' "
        End If
    End If
    If cboYear.Text <> "" Then
        If InStr(1, SQL, "WHERE") > 0 Then
            SQL = SQL & " AND YEAR =" & Val(cboYear.Text)
        Else
            SQL = SQL & " WHERE YEAR =" & Val(cboYear.Text)
        End If
    End If
        
    Debug.Print SQL


        CollegeADODC.RecordSource = SQL
        CollegeADODC.Refresh
        If CollegeADODC.Recordset.RecordCount = 0 Then
            MsgBox "No student List of this criteria", vbExclamation, "No Students"
        Else
            fyaddmission.Show
            fyaddmission.CollegeADODC.RecordSource = SQL
            fyaddmission.CollegeADODC.Refresh
            
            Unload frmStudentUpdate
        End If
    
End Sub

