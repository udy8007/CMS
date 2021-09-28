VERSION 5.00
Begin VB.Form Master_Subject 
   BackColor       =   &H80000012&
   Caption         =   "Subject Master"
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
      Picture         =   "Master_Subject.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7200
      Width           =   1441
   End
   Begin VB.ComboBox cmbcoursecode 
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
      Left            =   4958
      TabIndex        =   0
      Text            =   "cmbcoursecode"
      Top             =   2070
      Width           =   1965
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
      Picture         =   "Master_Subject.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6180
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
      Left            =   6075
      Picture         =   "Master_Subject.frx":0B44
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6180
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
      Picture         =   "Master_Subject.frx":1486
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
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
      Picture         =   "Master_Subject.frx":1790
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
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
      Picture         =   "Master_Subject.frx":1BD2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
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
      Picture         =   "Master_Subject.frx":214E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
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
      Picture         =   "Master_Subject.frx":2A18
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1441
   End
   Begin VB.TextBox subnmtxt 
      DataField       =   "SUBJECTNAME"
      DataSource      =   "Subject"
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
      Left            =   4958
      TabIndex        =   2
      Top             =   2955
      Width           =   3855
   End
   Begin VB.TextBox unittxt 
      DataField       =   "NOOFUNIT"
      DataSource      =   "Subject"
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
      Left            =   4958
      TabIndex        =   3
      Top             =   3390
      Width           =   1935
   End
   Begin VB.TextBox txtremark 
      DataField       =   "REMARK"
      DataSource      =   "Subject"
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
      Left            =   4958
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Master_Subject.frx":32E2
      Top             =   4740
      Width           =   3855
   End
   Begin VB.TextBox txtmin 
      DataField       =   "MINMARKS"
      DataSource      =   "Subject"
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
      Left            =   4958
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtmax 
      DataField       =   "MAXMARKS"
      DataSource      =   "Subject"
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
      Left            =   4958
      TabIndex        =   5
      Top             =   4290
      Width           =   1935
   End
   Begin VB.TextBox txtSubjectCode 
      DataField       =   "SUBJECTCODE"
      DataSource      =   "Subject"
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
      Left            =   4958
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   4125
      Left            =   443
      Shape           =   4  'Rounded Rectangle
      Top             =   1860
      Width           =   10995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Master"
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
      Height          =   795
      Index           =   1
      Left            =   2340
      TabIndex        =   22
      Top             =   1050
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   420
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
      Index           =   1
      Left            =   465
      TabIndex        =   21
      Top             =   360
      Width           =   10995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Index           =   0
      Left            =   3075
      TabIndex        =   20
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Name :"
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
      Index           =   0
      Left            =   3075
      TabIndex        =   19
      Top             =   2955
      Width           =   1740
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Of Unit :"
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
      Index           =   0
      Left            =   3075
      TabIndex        =   18
      Top             =   3390
      Width           =   1740
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code :"
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
      TabIndex        =   17
      Top             =   2085
      Width           =   1740
   End
   Begin VB.Label Label29 
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
      Left            =   3075
      TabIndex        =   16
      Top             =   4740
      Width           =   1740
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Mark :"
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
      TabIndex        =   15
      Top             =   3840
      Width           =   1740
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Mark :"
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
      TabIndex        =   14
      Top             =   4305
      Width           =   1740
   End
End
Attribute VB_Name = "Master_Subject"
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
If txtSubjectCode.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtSubjectCode.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete subjectmaster where courcecode = '" & cmbcoursecode.Text & "' and subjectcode = '" & txtSubjectCode.Text & "'"
    adodc.Execute sql
    rs1.Requery
    MsgBox " Record Deleted . . . !!!", vbInformation, "Delete"
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

txtSubjectCode.Text = ""
subnmtxt.Text = ""
unittxt.Text = ""
txtmin.Text = ""
txtmax.Text = ""
txtremark.Text = ""
Call enable_true

End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    cmbcoursecode.Text = rs1.Fields("courcecode")
    txtSubjectCode.Text = rs1.Fields("SubjectCode")
    subnmtxt.Text = rs1.Fields("SUBJECTNAME")
    unittxt.Text = rs1.Fields("NOOFUNIT")
    txtmin.Text = rs1.Fields("MINMARKS")
    txtmax.Text = rs1.Fields("MAXMARKS")
    txtremark.Text = rs1.Fields("REMARK")
    
End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
    If rs1.BOF = False Then
        cmbcoursecode.Text = rs1.Fields("courcecode")
        txtSubjectCode.Text = rs1.Fields("SubjectCode")
        subnmtxt.Text = rs1.Fields("SUBJECTNAME")
        unittxt.Text = rs1.Fields("NOOFUNIT")
        txtmin.Text = rs1.Fields("MINMARKS")
        txtmax.Text = rs1.Fields("MAXMARKS")
        txtremark.Text = rs1.Fields("REMARK")
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

If cmbcoursecode.Text = "" Then
    cnt = 1
End If
If txtSubjectCode.Text = "" Then
    cnt = 1
End If
If subnmtxt.Text = "" Then
    cnt = 1
End If
If unittxt.Text = "" Then
    cnt = 1
End If
If txtmin.Text = "" Then
    cnt = 1
End If
If txtmax.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If
If cnt = 0 Then
    Set rs = New ADODB.Recordset
    Dim Yn As String
    rs.Open "select count(*) from subjectmaster where courcecode = '" & cmbcoursecode.Text & "' and subjectcode = '" & txtSubjectCode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into subjectmaster values ('" & cmbcoursecode & "','" & txtSubjectCode.Text & "','" & subnmtxt.Text & "','" & unittxt.Text & "','" & txtmin.Text & "','" & txtmax.Text & "','" & txtremark.Text & "')"
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update subjectmaster set SUBJECTNAME  = '" & subnmtxt.Text & "',NOOFUNIT = '" & unittxt.Text & "', MINMARKS = '" & txtmin.Text & "', MAXMARKS = '" & txtmax.Text & "', REMARK = '" & txtremark.Text & "' where CourceCode = '" & cmbcoursecode & "' and SubjectCode = '" & txtSubjectCode.Text & "'"
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
cmbcoursecode.Text = rs1.Fields("courcecode")
txtSubjectCode.Text = rs1.Fields("SubjectCode")
subnmtxt.Text = rs1.Fields("SUBJECTNAME")
unittxt.Text = rs1.Fields("NOOFUNIT")
txtmin.Text = rs1.Fields("MINMARKS")
txtmax.Text = rs1.Fields("MAXMARKS")
txtremark.Text = rs1.Fields("REMARK")
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0

Call connection
Set rs = New ADODB.Recordset
rs.Open "select courcecode from courcemaster order by courcecode", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    cmbcoursecode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
Set rs1 = New ADODB.Recordset
rs1.Open "select * from subjectmaster order by SubjectCode", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New ADODB.Recordset

rs.Open "select * from subjectmaster order by SubjectCode", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    cmbcoursecode.Text = rs.Fields("courcecode")
    txtSubjectCode.Text = rs.Fields("SubjectCode")
    subnmtxt.Text = rs.Fields("SUBJECTNAME")
    unittxt.Text = rs.Fields("NOOFUNIT")
    txtmin.Text = rs.Fields("MINMARKS")
    txtmax.Text = rs.Fields("MAXMARKS")
    txtremark.Text = rs.Fields("REMARK")
    
End If
End Sub

Private Sub subnmtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub txtmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub unittxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub enable_true()
    cmbcoursecode.Enabled = True
    txtSubjectCode.Enabled = True
    subnmtxt.Enabled = True
    unittxt.Enabled = True
    txtmin.Enabled = True
    txtmax.Enabled = True
    txtremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    txtSubjectCode.SetFocus
    
End Sub

Private Sub enable_false()
    cmbcoursecode.Enabled = False
    txtSubjectCode.Enabled = False
    subnmtxt.Enabled = False
    unittxt.Enabled = False
    txtmin.Enabled = False
    txtmax.Enabled = False
    txtremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
End Sub
