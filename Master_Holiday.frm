VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Master_Holiday 
   BackColor       =   &H80000012&
   Caption         =   "Holiday Master"
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
      Left            =   3510
      Picture         =   "Master_Holiday.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1441
   End
   Begin VB.TextBox txtRemark 
      DataField       =   "REMARK"
      DataSource      =   "Holiday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4853
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Master_Holiday.frx":0442
      Top             =   3870
      Width           =   3855
   End
   Begin VB.TextBox txtholiday 
      DataField       =   "HOLIDAY"
      DataSource      =   "Holiday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4853
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3135
      Width           =   3855
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
      Left            =   4365
      Picture         =   "Master_Holiday.frx":0447
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
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
      Left            =   6090
      Picture         =   "Master_Holiday.frx":0B49
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
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
      Left            =   3510
      Picture         =   "Master_Holiday.frx":148B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
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
      Left            =   6930
      Picture         =   "Master_Holiday.frx":1795
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
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
      Left            =   8640
      Picture         =   "Master_Holiday.frx":1BD7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
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
      Left            =   1800
      Picture         =   "Master_Holiday.frx":2153
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
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
      Left            =   5220
      Picture         =   "Master_Holiday.frx":2A1D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1441
   End
   Begin MSComCtl2.DTPicker HolidayDate 
      DataField       =   "HOLIDAYDATE"
      DataSource      =   "Holiday"
      Height          =   435
      Left            =   4860
      TabIndex        =   0
      Top             =   2370
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   767
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
      Format          =   142802945
      CurrentDate     =   37609
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3255
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
      Width           =   10995
   End
   Begin VB.Label Label28 
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
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2880
      TabIndex        =   14
      Top             =   3870
      Width           =   1875
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday Name :"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   3135
      Width           =   1875
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday Date :"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   2370
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday Master"
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
      Left            =   2565
      TabIndex        =   11
      Top             =   1140
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   210
      Width           =   10995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sri Devi Arts and Science Collegece College"
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
      TabIndex        =   10
      Top             =   450
      Width           =   10995
   End
End
Attribute VB_Name = "Master_Holiday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset
Dim rs1 As adodb.Recordset
Dim sql As String
Dim cnt As Integer

Private Sub cmdDelete_Click()
If txtholiday.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtholiday.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete holidaymaster where holiday = '" & txtholiday.Text & "' and holidaydate = '" & HolidayDate.Value & "'"
    adodc.Execute sql
    rs1.Requery
    MsgBox " Record Deleted . . . !!!", vbInformation, "Delete"
    Call Form_Load
End If
End Sub

Private Sub cmdedit_Click()
Call enable_true
End Sub

Private Sub cmdnew_Click()
Call enable_true
txtholiday.Text = ""
HolidayDate.Value = Date
txtremark.Text = ""
End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtholiday.Text = rs1.Fields("Holiday")
    HolidayDate.Value = rs1.Fields("HolidayDate")
    txtremark.Text = rs1.Fields("REMARK")
End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
txtholiday.Text = rs1.Fields("Holiday")
    HolidayDate.Value = rs1.Fields("HolidayDate")
    txtremark.Text = rs1.Fields("REMARK")
End If
End If
End Sub

Private Sub cmdpre_Click()

End Sub

Private Sub cmdrefresh_Click()
Call Form_Load
End Sub

Private Sub cmdreturn_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
cnt = 0
If txtholiday.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If
If cnt = 0 Then
    Dim Yn As String
    Set rs = New adodb.Recordset
    rs.Open "select count(holiday) from holidaymaster where HolidayDate = '" & HolidayDate.Value & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into holidaymaster values ('" & txtholiday.Text & "','" & HolidayDate.Value & "','" & txtremark.Text & "')"
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update holidaymaster set holiday = '" & txtholiday.Text & "',remark = '" & txtremark & "'where holidaydate = '" & HolidayDate.Value & "'"
            adodc.Execute sql
            MsgBox "Record Updated . . . !!!", vbInformation, "Update"
        End If
    End If
Else
    MsgBox "Can not Insert Null Value . . . !!!", vbCritical, "Invalid Data"
End If
cnt = 0
rs1.Requery
rs1.MoveFirst
txtholiday.Text = rs1.Fields("Holiday")
HolidayDate.Value = rs1.Fields("HolidayDate")
txtremark.Text = rs1.Fields("REMARK")
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0
Set rs1 = New adodb.Recordset
rs1.Open "select * from holidaymaster order by holidaydate", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New adodb.Recordset

rs.Open "select * from holidaymaster order by holidaydate", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    txtholiday.Text = rs.Fields("Holiday")
    HolidayDate.Value = rs.Fields("HolidayDate")
    txtremark.Text = rs.Fields("REMARK")
End If
End Sub

Private Sub txtholiday_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub


Private Sub enable_true()
    txtholiday.Enabled = True
    HolidayDate.Enabled = True
    txtremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    txtholiday.SetFocus
End Sub

Private Sub enable_false()
    txtholiday.Enabled = False
    HolidayDate.Enabled = False
    txtremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
End Sub
