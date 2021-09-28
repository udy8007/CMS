VERSION 5.00
Begin VB.Form Master_Course 
   BackColor       =   &H00000000&
   Caption         =   "Course Master"
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
      Picture         =   "Master_Course.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
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
      Picture         =   "Master_Course.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
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
      Picture         =   "Master_Course.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "Master_Course.frx":15D6
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
      Picture         =   "Master_Course.frx":1B52
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
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
      Picture         =   "Master_Course.frx":1F94
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
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
      Picture         =   "Master_Course.frx":229E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1441
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
      Left            =   4350
      Picture         =   "Master_Course.frx":2BE0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1441
   End
   Begin VB.TextBox txtcoursecode 
      DataField       =   "COURCECODE"
      DataSource      =   "Cource"
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
      Left            =   4853
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2460
      Width           =   2415
   End
   Begin VB.TextBox txtcoursename 
      DataField       =   "COURCENAME"
      DataSource      =   "Cource"
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
      Left            =   4853
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2985
      Width           =   4575
   End
   Begin VB.TextBox txtduration 
      DataField       =   "DURATION"
      DataSource      =   "Cource"
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
      Left            =   4853
      MaxLength       =   15
      TabIndex        =   2
      Top             =   3495
      Width           =   2415
   End
   Begin VB.TextBox txtremark 
      DataField       =   "REMARK"
      DataSource      =   "Cource"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4853
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Master_Course.frx":32E2
      Top             =   4020
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
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
      Left            =   420
      TabIndex        =   16
      Top             =   360
      Width           =   10995
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Course Master"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   2340
      TabIndex        =   14
      Top             =   2460
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name :"
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
      Left            =   2340
      TabIndex        =   13
      Top             =   2985
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duration :"
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
      Left            =   2340
      TabIndex        =   12
      Top             =   3495
      Width           =   1575
   End
   Begin VB.Label Label32 
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
      Left            =   2340
      TabIndex        =   11
      Top             =   4020
      Width           =   1575
   End
End
Attribute VB_Name = "Master_Course"
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

If txtcoursecode.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtcoursecode.SetFocus
    Exit Sub
End If

sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete courcemaster where courcecode = '" & txtcoursecode.Text & "'"
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
Call enable_true
txtcoursecode.Text = ""
txtcoursename.Text = ""
txtduration.Text = ""
txtremark.Text = ""

End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtcoursecode.Text = rs1.Fields(0)
    txtcoursename.Text = rs1.Fields(1)
    txtduration.Text = rs1.Fields(2)
    txtremark.Text = rs1.Fields(3)
End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
    txtcoursecode.Text = rs1.Fields(0)
    txtcoursename.Text = rs1.Fields(1)
    txtduration.Text = rs1.Fields(2)
    txtremark.Text = rs1.Fields(3)
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
Dim Yn As String
If txtcoursecode.Text = "" Then
    cnt = 1
End If
If txtcoursename.Text = "" Then
    cnt = 1
End If
If txtduration.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If
If cnt = 0 Then
    Set rs = New ADODB.Recordset
    rs.Open "select count(courcecode) from courcemaster where courcecode = '" & txtcoursecode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into courcemaster values ('" & txtcoursecode.Text & "','" & txtcoursename.Text & "','" & txtduration.Text & "','" & txtremark.Text & "')"
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update courcemaster set COURCENAME = '" & txtcoursename.Text & "',DURATION = '" & txtduration & "',remark = '" & txtremark.Text & "' where COURCECODE = '" & txtcoursecode.Text & "'"
            adodc.Execute sql
            MsgBox "Record Updated", vbInformation, "Update"
           
        End If
    End If
Else
    MsgBox "Can not Insert Null Value . . . !!!", vbCritical, "Invalid Data"
End If
cnt = 0
rs1.Requery
rs1.MoveFirst
txtcoursecode.Text = rs1.Fields(0)
txtcoursename.Text = rs1.Fields(1)
txtduration.Text = rs1.Fields(2)
txtremark.Text = rs1.Fields(3)
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0
Set rs1 = New ADODB.Recordset
rs1.Open "select * from courcemaster order by courcecode", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New ADODB.Recordset

rs.Open "select * from courcemaster order by courcecode", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    txtcoursecode.Text = rs.Fields(0)
    txtcoursename.Text = rs.Fields(1)
    txtduration.Text = rs.Fields(2)
    txtremark.Text = rs.Fields(3)
End If
End Sub

Private Sub txtcoursecode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub txtcoursename_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub txtduration_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Then Else KeyAscii = 0
End Sub

Private Sub enable_true()

    txtcoursecode.Enabled = True
    txtcoursename.Enabled = True
    txtduration.Enabled = True
    txtremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    txtcoursecode.SetFocus
    

End Sub

Private Sub enable_false()
    txtcoursecode.Enabled = False
    txtcoursename.Enabled = False
    txtduration.Enabled = False
    txtremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
End Sub
