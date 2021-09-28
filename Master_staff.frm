VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Master_staff 
   BackColor       =   &H80000008&
   Caption         =   "Staff Master"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Edit"
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
      Left            =   3810
      Picture         =   "Master_staff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7410
      Width           =   1441
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LASTNAME"
      DataSource      =   "Staff"
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
      Left            =   3330
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2535
      Width           =   2415
   End
   Begin VB.TextBox txtphone 
      DataField       =   "PHONE"
      DataSource      =   "Staff"
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
      Left            =   3330
      TabIndex        =   4
      Text            =   "0"
      Top             =   5130
      Width           =   2415
   End
   Begin VB.TextBox txtadd 
      DataField       =   "ADDRESS"
      DataSource      =   "Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4028
      Width           =   2415
   End
   Begin VB.TextBox txtEmpCode 
      BackColor       =   &H00C0C0FF&
      DataField       =   "EMPCODE"
      DataSource      =   "Staff"
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
      Left            =   3330
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2040
      Width           =   1425
   End
   Begin VB.TextBox txtsecname 
      DataField       =   "PARENTNAME"
      DataSource      =   "Staff"
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
      Left            =   3330
      TabIndex        =   2
      Top             =   3540
      Width           =   2415
   End
   Begin VB.TextBox txtfirstname 
      DataField       =   "FIRSTNAME"
      DataSource      =   "Staff"
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
      Left            =   3300
      TabIndex        =   1
      Top             =   3030
      Width           =   2415
   End
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Previous"
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
      Left            =   4665
      Picture         =   "Master_staff.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6450
      Width           =   1441
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Next"
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
      Left            =   6375
      Picture         =   "Master_staff.frx":0B44
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6450
      Width           =   1441
   End
   Begin VB.CommandButton cmdsave 
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
      Left            =   3810
      Picture         =   "Master_staff.frx":1486
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7380
      Width           =   1441
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   7230
      Picture         =   "Master_staff.frx":1790
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7380
      Width           =   1441
   End
   Begin VB.CommandButton cmdreturn 
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
      Left            =   8940
      Picture         =   "Master_staff.frx":1BD2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7380
      Width           =   1441
   End
   Begin VB.CommandButton cmdnew 
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
      Left            =   2100
      Picture         =   "Master_staff.frx":214E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7380
      Width           =   1441
   End
   Begin VB.CommandButton cmdrefresh 
      BackColor       =   &H0080FFFF&
      Caption         =   "Re&fresh"
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
      Left            =   5520
      Picture         =   "Master_staff.frx":2A18
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7380
      Width           =   1441
   End
   Begin VB.TextBox txtsal 
      DataField       =   "SALARY"
      DataSource      =   "Staff"
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
      Left            =   7710
      TabIndex        =   8
      Top             =   3839
      Width           =   2415
   End
   Begin VB.ComboBox cmbdesg 
      DataField       =   "DESIGNATION"
      DataSource      =   "Staff"
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
      ItemData        =   "Master_staff.frx":32E2
      Left            =   7710
      List            =   "Master_staff.frx":32F8
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3247
      Width           =   2415
   End
   Begin VB.TextBox txtqualification 
      DataField       =   "QUALIFICATION"
      DataSource      =   "Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtremark 
      DataField       =   "REMARK"
      DataSource      =   "Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Master_staff.frx":3336
      Top             =   4431
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpjoin 
      Bindings        =   "Master_staff.frx":333B
      DataField       =   "JOINDATE"
      DataSource      =   "Staff"
      Height          =   360
      Left            =   7710
      TabIndex        =   11
      Top             =   5640
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   142737409
      CurrentDate     =   37673
   End
   Begin MSComCtl2.DTPicker dtpbirth 
      Bindings        =   "Master_staff.frx":335D
      DataField       =   "BIRTHDATE"
      DataSource      =   "Staff"
      Height          =   360
      Left            =   3330
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   142737409
      CurrentDate     =   37673
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sur Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   32
      Top             =   2535
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   4365
      Left            =   465
      Shape           =   4  'Rounded Rectangle
      Top             =   1830
      Width           =   10995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Master"
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
      Height          =   615
      Left            =   2550
      TabIndex        =   30
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   465
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10995
   End
   Begin VB.Label Label2 
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
      Left            =   465
      TabIndex        =   29
      Top             =   360
      Width           =   10995
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   28
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   27
      Top             =   3045
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Second Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   26
      Top             =   3540
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   25
      Top             =   4035
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   24
      Top             =   5145
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Salary :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   5670
      TabIndex        =   23
      Top             =   3839
      Width           =   1815
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   5670
      TabIndex        =   22
      Top             =   3247
      Width           =   1815
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   1260
      TabIndex        =   21
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Join Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   5670
      TabIndex        =   20
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   5670
      TabIndex        =   19
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remark :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   5670
      TabIndex        =   10
      Top             =   4431
      Width           =   1815
   End
End
Attribute VB_Name = "Master_staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim sql As String
Dim cnt As Integer

Private Sub cmdDelete_Click()
On Error GoTo errhandler
If txtLastName.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtLastName.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete staffmaster where empcode = " & txtEmpCode.Text & ""
    adodc.Execute sql
    rs1.Requery
    MsgBox " Record Deleted . . . !!!", vbInformation, "Delete"
    txtEmpCode.Text = ""
    txtLastName.Text = ""
    txtsecname.Text = ""
    txtfirstname.Text = ""
    txtadd.Text = ""
    txtphone.Text = ""
    txtqualification.Text = ""
    cmbdesg.Text = ""
    txtsal.Text = ""
    txtremark.Text = ""
    dtpjoin.Value = Now
    dtpbirth.Value = Now
    
    Call Form_Load
End If
Exit Sub
errhandler:
'MsgBox Err.Number
'MsgBox Err.Description
If Err.Number = -2147217873 Then
    MsgBox "You can not delete this record . . .!!!", vbCritical
    MsgBox "First delete all the reference record . . . !!!", vbInformation
Else
    MsgBox "You can not perform this Operation . . . !!!", vbCritical
End If

End Sub

Private Sub cmdedit_Click()
Call enable_true
End Sub

Private Sub cmdnew_Click()
Call enable_true
Set rs = New ADODB.Recordset
rs.Open "select * from staffmaster", adodc, adOpenKeyset, adLockOptimistic
cnt = 0
Do While Not rs.EOF
    If cnt < rs.Fields("empcode") Then cnt = rs.Fields("empcode")
    rs.MoveNext
Loop
cnt = cnt + 1
txtEmpCode.Text = cnt
txtLastName.Text = ""
txtsecname.Text = ""
txtfirstname.Text = ""
txtadd.Text = ""
txtphone.Text = ""
txtqualification.Text = ""
cmbdesg.Text = ""
txtsal.Text = ""
txtremark.Text = ""
dtpjoin.Value = Now
dtpbirth.Value = Now
End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtEmpCode.Text = rs1.Fields("empcode")
    txtLastName.Text = rs1.Fields("lastname")
    txtsecname.Text = rs1.Fields("parentname")
    txtfirstname.Text = rs1.Fields("firstname")
    txtadd.Text = rs1.Fields("ADDRESS")
    txtphone.Text = rs1.Fields("PHONE")
    txtqualification.Text = rs1.Fields("QUALIFICATION")
    cmbdesg.Text = rs1.Fields("DESIGNATION")
    txtsal.Text = rs1.Fields("SALARY")
    txtremark.Text = rs1.Fields("REMARK")
    dtpjoin.Value = rs1.Fields("JOINDATE")
    dtpbirth.Value = rs1.Fields("BIRTHDATE")
End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
    txtEmpCode.Text = rs1.Fields("empcode")
    txtLastName.Text = rs1.Fields("lastname")
    txtfirstname.Text = rs1.Fields("firstname")
    txtsecname.Text = rs1.Fields("parentname")
    txtadd.Text = rs1.Fields("ADDRESS")
    txtphone.Text = rs1.Fields("PHONE")
    txtqualification.Text = rs1.Fields("QUALIFICATION")
    cmbdesg.Text = rs1.Fields("DESIGNATION")
    txtsal.Text = rs1.Fields("SALARY")
    txtremark.Text = rs1.Fields("REMARK")
    dtpjoin.Value = rs1.Fields("JOINDATE")
    dtpbirth.Value = rs1.Fields("BIRTHDATE")
End If
End If
End Sub
Private Sub cmdrefresh_Click()
Call Form_Load
End Sub

Private Sub cmdreturn_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
cnt = 0
If txtLastName.Text = "" Then
    cnt = 1
End If
If txtfirstname.Text = "" Then
    cnt = 1
End If
If txtsecname.Text = "" Then
    cnt = 1
End If
If txtadd.Text = "" Then
    cnt = 1
End If
If txtphone.Text = "" Then
    cnt = 1
End If
If txtqualification.Text = "" Then
    cnt = 1
End If
If cmbdesg.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If
If txtsal.Text = "" Then
    cnt = 1
End If
If cnt = 0 Then
    Set rs = New ADODB.Recordset
    Dim Yn As String
    rs.Open "select count(empcode) from staffmaster where empcode = '" & txtEmpCode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
     
    If rs.Fields(0) = 0 Then
       sql = "insert into staffmaster values (" & txtEmpCode.Text & ",'" & txtLastName.Text & "','" & txtfirstname.Text & "','" & txtsecname.Text & "','" & txtadd.Text & "'," & CDec(txtphone.Text) & ",'" & txtqualification.Text & "','" & cmbdesg.Text & "'," & CDec(txtsal.Text) & ",'" & txtremark.Text & "','" & dtpjoin.Value & "','" & dtpbirth.Value & "')"
       ' MsgBox sql, vbInformation, "Save"
         adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        
   Else
      
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
      Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
           sql = "Update staffmaster set LASTNAME = '" & txtLastName.Text & "',FIRSTNAME = '" & txtfirstname.Text & "', PARENTNAME = '" & txtsecname.Text & "',ADDRESS = '" & txtadd.Text & "',DESIGNATION = '" & cmbdesg.Text & "', PHONE = " & txtphone.Text & ", JOINDATE ='" & dtpjoin.Value & "', BIRTHDATE = '" & dtpbirth.Value & "',QUALIFICATION = '" & txtqualification.Text & "',SALARY = " & txtsal.Text & ",REMARK = '" & txtremark.Text & "' where empcode = " & txtEmpCode.Text & ""
            adodc.Execute sql
           ' MsgBox sql
           MsgBox "Record Updated . . . !!!", vbInformation, "Update"
            
        End If
        
   End If
Else
    MsgBox "Can not Insert Null Value . . . !!!", vbCritical, "Invalid Data"
End If
 
rs1.Requery
rs1.MoveFirst
txtEmpCode.Text = rs1.Fields("empcode")
txtLastName.Text = rs1.Fields("lastname")
txtsecname.Text = rs1.Fields("parentname")
txtfirstname.Text = rs1.Fields("firstname")
txtadd.Text = rs1.Fields("ADDRESS")
txtphone.Text = rs1.Fields("PHONE")
txtqualification.Text = rs1.Fields("QUALIFICATION")
cmbdesg.Text = rs1.Fields("DESIGNATION")
txtsal.Text = rs1.Fields("SALARY")
txtremark.Text = rs1.Fields("REMARK")
dtpjoin.Value = rs1.Fields("JOINDATE")
dtpbirth.Value = rs1.Fields("BIRTHDATE")

cnt = 0
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0
Set rs1 = New ADODB.Recordset
rs1.Open "select * from staffmaster order by  empcode", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New ADODB.Recordset

rs.Open "select * from staffmaster order by  empcode", adodc, adOpenKeyset, adLockOptimistic
'rs.MoveFirst
rs.Requery
If rs.EOF = False Then
    txtEmpCode.Text = rs.Fields("empcode")
    txtLastName.Text = rs.Fields("lastname")
    txtsecname.Text = rs.Fields("parentname")
    txtfirstname.Text = rs.Fields("firstname")
    txtadd.Text = rs.Fields("ADDRESS")
    txtphone.Text = rs.Fields("PHONE")
    txtqualification.Text = rs.Fields("QUALIFICATION")
    cmbdesg.Text = rs.Fields("DESIGNATION")
    txtsal.Text = rs.Fields("SALARY")
    txtremark.Text = rs.Fields("REMARK")
    dtpjoin.Value = rs.Fields("JOINDATE")
    dtpbirth.Value = rs.Fields("BIRTHDATE")
   
End If
End Sub



Private Sub txtfirstname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtsal_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtsecname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub enable_true()

txtEmpCode.Enabled = True
txtLastName.Enabled = True
txtsecname.Enabled = True
txtfirstname.Enabled = True
txtadd.Enabled = True
txtphone.Enabled = True
txtqualification.Enabled = True
cmbdesg.Enabled = True
txtsal.Enabled = True
txtremark.Enabled = True
dtpjoin.Enabled = True
dtpbirth.Enabled = True
cmdsave.Visible = True
cmdedit.Visible = False
txtLastName.SetFocus

End Sub

Private Sub enable_false()

txtEmpCode.Enabled = False
txtLastName.Enabled = False
txtsecname.Enabled = False
txtfirstname.Enabled = False
txtadd.Enabled = False
txtphone.Enabled = False
txtqualification.Enabled = False
cmbdesg.Enabled = False
txtsal.Enabled = False
txtremark.Enabled = False
dtpjoin.Enabled = False
dtpbirth.Enabled = False
cmdsave.Visible = False
cmdedit.Visible = True

End Sub
