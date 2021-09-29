VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fyaddmission 
   BackColor       =   &H80000007&
   Caption         =   "Admission Process"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fyaddmission.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Edit"
      Height          =   720
      Left            =   2100
      Picture         =   "fyaddmission.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7470
      Width           =   1441
   End
   Begin VB.ComboBox cmbcast 
      Height          =   360
      ItemData        =   "fyaddmission.frx":2BE4
      Left            =   7890
      List            =   "fyaddmission.frx":2BF4
      TabIndex        =   12
      Top             =   3435
      Width           =   1950
   End
   Begin VB.TextBox TXTOCCU 
      DataField       =   "OCCUPATION"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   7890
      MaxLength       =   250
      TabIndex        =   13
      Top             =   3840
      Width           =   1950
   End
   Begin VB.TextBox txtRemark 
      DataField       =   "REMARK"
      DataSource      =   "CollegeADODC"
      Height          =   1005
      Left            =   7890
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "fyaddmission.frx":2C09
      Top             =   6120
      Width           =   3105
   End
   Begin VB.TextBox TXTPADD 
      DataField       =   "PADDRESS"
      DataSource      =   "CollegeADODC"
      Height          =   1005
      Left            =   7890
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4245
      Width           =   3105
   End
   Begin VB.TextBox TXTPPHONE 
      DataField       =   "PPHONE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   7890
      MaxLength       =   15
      TabIndex        =   15
      Top             =   5310
      Width           =   1950
   End
   Begin VB.TextBox TXTANNINC 
      DataField       =   "ANNUALINCOME"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   7890
      MaxLength       =   14
      TabIndex        =   16
      Top             =   5715
      Width           =   1950
   End
   Begin VB.TextBox TXTLADD 
      DataField       =   "ADDRESS"
      DataSource      =   "CollegeADODC"
      Height          =   1005
      Left            =   2550
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4245
      Width           =   3105
   End
   Begin VB.ComboBox cmbgender 
      Height          =   360
      ItemData        =   "fyaddmission.frx":2C0E
      Left            =   2550
      List            =   "fyaddmission.frx":2C18
      TabIndex        =   10
      Top             =   6120
      Width           =   1950
   End
   Begin VB.TextBox TXTROLLNO 
      BackColor       =   &H00C0C0FF&
      DataField       =   "ROLLNO"
      Height          =   360
      Left            =   7890
      TabIndex        =   3
      Top             =   2460
      Width           =   1950
   End
   Begin VB.ComboBox cmbcourse 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   2550
      TabIndex        =   0
      Top             =   2040
      Width           =   2430
   End
   Begin VB.ComboBox cmby 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      ItemData        =   "fyaddmission.frx":2C2A
      Left            =   2550
      List            =   "fyaddmission.frx":2C37
      TabIndex        =   1
      Top             =   2460
      Width           =   2430
   End
   Begin VB.CommandButton cmdrefresh 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Refresh"
      Height          =   720
      Left            =   5286
      Picture         =   "fyaddmission.frx":2C44
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Add New"
      Height          =   720
      Left            =   510
      Picture         =   "fyaddmission.frx":350E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Return"
      Height          =   720
      Left            =   10064
      Picture         =   "fyaddmission.frx":3DD8
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
      Height          =   720
      Left            =   8470
      Picture         =   "fyaddmission.frx":4354
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
      Height          =   720
      Left            =   2102
      Picture         =   "fyaddmission.frx":4796
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Next"
      Height          =   720
      Left            =   6878
      Picture         =   "fyaddmission.frx":4AA0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7470
      Width           =   1441
   End
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Previous"
      Height          =   720
      Left            =   3694
      Picture         =   "fyaddmission.frx":53E2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7470
      Width           =   1441
   End
   Begin VB.ComboBox cmbstream 
      Height          =   360
      ItemData        =   "fyaddmission.frx":5AE4
      Left            =   2550
      List            =   "fyaddmission.frx":5AF4
      TabIndex        =   9
      Top             =   5730
      Width           =   1950
   End
   Begin VB.TextBox TXTLPHONE 
      DataField       =   "PPHONE"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   2550
      MaxLength       =   15
      TabIndex        =   8
      Top             =   5310
      Width           =   1950
   End
   Begin VB.TextBox TXTLNAME 
      DataField       =   "LASTNAME"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   4
      Top             =   3030
      Width           =   2430
   End
   Begin VB.TextBox TXTSNAME 
      DataField       =   "FIRSTNAME"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   6
      Top             =   3840
      Width           =   2430
   End
   Begin VB.TextBox TXTFNAME 
      DataField       =   "PARENTNAME"
      DataSource      =   "CollegeADODC"
      Height          =   360
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   5
      Top             =   3435
      Width           =   2430
   End
   Begin MSComCtl2.DTPicker dtjoin 
      Height          =   360
      Left            =   7890
      TabIndex        =   2
      Top             =   2040
      Width           =   1950
      _ExtentX        =   3440
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
      CalendarBackColor=   12632319
      Format          =   36962305
      CurrentDate     =   38427
   End
   Begin MSComCtl2.DTPicker dtbirth 
      Height          =   360
      Left            =   7890
      TabIndex        =   11
      Top             =   3030
      Width           =   1950
      _ExtentX        =   3440
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
      Format          =   36962305
      CurrentDate     =   38425
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cast :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   44
      Top             =   3435
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   0
      Left            =   390
      TabIndex        =   43
      Top             =   3435
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Second Name :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   390
      TabIndex        =   42
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   390
      TabIndex        =   41
      Top             =   3030
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Address :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   0
      Left            =   390
      TabIndex        =   40
      Top             =   4245
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   390
      TabIndex        =   39
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stream :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   390
      TabIndex        =   38
      Top             =   5715
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   37
      Top             =   3030
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Permenent Address :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   36
      Top             =   4245
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Phone :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   390
      TabIndex        =   35
      Top             =   5310
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Permenent Phone :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   34
      Top             =   5310
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   33
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Annual Income :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   32
      Top             =   5715
      Width           =   2055
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remark :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   5700
      TabIndex        =   31
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   4335
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   2910
      Width           =   10995
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   630
      TabIndex        =   30
      Top             =   2070
      Width           =   1695
   End
   Begin VB.Label LBLROLLNO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No. :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   6060
      TabIndex        =   29
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   630
      TabIndex        =   28
      Top             =   2430
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Join Date :"
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   6060
      TabIndex        =   27
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1005
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   10995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admission"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   26
      Top             =   1110
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   465
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   10995
   End
   Begin VB.Label Label17 
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
      TabIndex        =   25
      Top             =   390
      Width           =   10995
   End
End
Attribute VB_Name = "fyaddmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
'
'Private Sub cmddel_Click()
'    If MsgBox("Delete", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Confirmation") = vbYes Then
'        CollegeADODC.Recordset.Delete
'        MsgBox "Record has been deleted", vbExclamation, "Delete Record"
'        CollegeADODC.Refresh
'    End If
'End Sub
'
'Private Sub cmdhome_Click()
'      Unload Me
'      College.Show
'End Sub
'
'Private Sub cmdnext_Click()
'    CollegeADODC.Recordset.MoveNext
'    If CollegeADODC.Recordset.EOF = True Then
'        MsgBox "You are already present on last record", vbExclamation, "Last Record"
'        CollegeADODC.Recordset.MoveLast
'    End If
'End Sub
'Private Sub cmdpre_Click()
'    If Not CollegeADODC.Recordset.EOF Then
'        CollegeADODC.Recordset.MovePrevious
'    End If
'    If CollegeADODC.Recordset.BOF = True Then
'        MsgBox "You are alread present on First record", vbExclamation, "First Record"
'        CollegeADODC.Recordset.MoveFirst
'    End If
'End Sub
'Private Sub cmdrefresh_Click()
'    CollegeADODC.Recordset.CancelUpdate
'    CollegeADODC.Refresh
'End Sub
'
'Private Sub CMDUPDATE_Click()
'On Error GoTo ERR1
'    CollegeADODC.Recordset.Update
'    MsgBox "Record Has Been Change", vbInformation, "Save Record"
'    Exit Sub
'ERR1:
'    MsgBox "Information missing", vbExclamation, "Update Error"
'End Sub
'
'Private Sub CourceCode_Click(Area As Integer)
'
'End Sub
'
'Private Sub cmdsave_Click()
'End Sub
'
'Private Sub Form_Load()
'    Call connection
'    Set rs = New ADODB.Recordset
'    rs.Open "select courcecode from courcemaster", adodc, adOpenKeyset, adLockOptimistic
'    While rs.EOF = False
'        cmbcourse.AddItem rs.Fields(0)
'        rs.MoveNext
'    Wend
'    dtbirth.Value = Date
'    'CMBCAST.Text = "GEN"
'End Sub
'Private Sub cmdAdd_Click()
'    CollegeADODC.Recordset.AddNew
'    DTPicker1.Value = Date - 17 * 365
'    cmbcast.Text = "GEN"
'    CourceCode.SetFocus
'End Sub
'
'Private Sub OPTNEW_Click()
'
'    cmbyear.Text = "YEAR"
'    cmdAdd.Enabled = True
'    cmdUpdate.Enabled = False
'    LBLROLLNO.Visible = False
'    TXTROLLNO.Visible = False
'End Sub
'
'Private Sub OPTOLD_Click()
'    If cmbyear.Text = "2" Or cmbyear.Text = "3" Then
'        LBLROLLNO.Visible = True
'        TXTROLLNO.Visible = True
'    End If
'    cmdUpdate.Enabled = True
'End Sub
'
'Function blankbox()
'    TXTFNAME.Text = ""
'    TXTSNAME.Text = ""
'    TXTLNAME.Text = ""
'    TXTLADD.Text = ""
'    TXTLPHONE.Text = ""
'    TXTPADD.Text = ""
'    TXTPPHONE.Text = ""
'    cmbstream.Text = "STREAM"
'    cmbcast.Text = "GENERAL"
'    TXTOCCU.Text = ""
'    TXTANNINC.Text = ""
'End Function
'
'
'Private Sub txtROLLno_KeyPress(KeyAscii As Integer)
'    Dim m As Integer
'    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Then
'       m = 1
'    Else
'        KeyAscii = 0
'    End If
'
'End Sub
'
'Private Sub txtLname_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
'End Sub
'
'Private Sub txtSname_GotFocus()
'    TXTSNAME.SelStart = 0
'    TXTSNAME.SelLength = Len(TXTSNAME.Text)
'End Sub
'Private Sub txtSname_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
'End Sub
'
'Private Sub txtFname_GotFocus()
'    TXTFNAME.SelStart = 0
'    TXTFNAME.SelLength = Len(TXTFNAME.Text)
'End Sub
'Private Sub txtFname_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
'End Sub
'
Dim rs As adodb.Recordset
Dim rs1 As adodb.Recordset
Dim sql As String
Dim cnt As Integer

Private Sub cmdDelete_Click()
On Error GoTo errhandler

If TXTLNAME.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    TXTLNAME.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete studentmaster where courcecode = '" & cmbcourse.Text & "' and year = '" & cmby.Text & "' and rollno = '" & CDec(txtRollNo.Text) & "'"
    adodc.Execute sql
    rs1.Requery
    MsgBox " Record Deleted", vbInformation, "Delete"
    dtjoin.Value = Date
    txtRollNo.Text = ""
    TXTLNAME.Text = ""
    TXTFNAME.Text = ""
    TXTSNAME.Text = ""
    TXTLADD.Text = ""
    TXTLPHONE.Text = ""
    TXTPADD.Text = ""
    TXTPPHONE.Text = ""
    dtbirth.Value = Date
    TXTOCCU.Text = ""
    TXTANNINC.Text = ""
    txtremark.Text = ""
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
Set rs = New adodb.Recordset
rs.Open "select * from studentmaster", adodc, adOpenKeyset, adLockOptimistic

    dtjoin.Value = Now
    
    txtRollNo.Text = ""
    TXTLNAME.Text = ""
    TXTFNAME.Text = ""
    TXTSNAME.Text = ""
    TXTLADD.Text = ""
    TXTLPHONE.Text = ""
    TXTPADD.Text = ""
    TXTPPHONE.Text = ""
    dtbirth.Value = Now
    TXTOCCU.Text = ""
    TXTANNINC.Text = ""
    txtremark.Text = ""

End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    cmbcourse.Text = rs1.Fields("CourceCode")
    dtjoin.Value = rs1.Fields("JOINDATE")
    cmby.Text = rs1.Fields("year")
    txtRollNo.Text = rs1.Fields("RollNo")
    TXTLNAME.Text = rs1.Fields("LASTNAME")
    TXTFNAME.Text = rs1.Fields("FIRSTNAME")
    TXTSNAME.Text = rs1.Fields("PARENTNAME")
    TXTLADD.Text = rs1.Fields("ADDRESS")
    TXTLPHONE.Text = rs1.Fields("PHONE")
    TXTPADD.Text = rs1.Fields("PADDRESS")
    TXTPPHONE.Text = rs1.Fields("PPHONE")
    cmbgender.Text = rs1.Fields("GENDER")
    cmbstream.Text = rs1.Fields("STREAM")
    dtbirth.Value = rs1.Fields("bIRTHDATE")
    cmbcast.Text = rs1.Fields("CAST")
    TXTOCCU.Text = rs1.Fields("OCCUPATION")
    TXTANNINC.Text = rs1.Fields("ANNUALINCOME")
    txtremark.Text = rs1.Fields("REMARK")
    End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
        cmbcourse.Text = rs1.Fields("CourceCode")
    dtjoin.Value = rs1.Fields("JOINDATE")
    cmby.Text = rs1.Fields("year")
    txtRollNo.Text = rs1.Fields("RollNo")
    TXTLNAME.Text = rs1.Fields("LASTNAME")
    TXTFNAME.Text = rs1.Fields("FIRSTNAME")
    TXTSNAME.Text = rs1.Fields("PARENTNAME")
    TXTLADD.Text = rs1.Fields("ADDRESS")
    TXTLPHONE.Text = rs1.Fields("PHONE")
    TXTPADD.Text = rs1.Fields("PADDRESS")
    TXTPPHONE.Text = rs1.Fields("PPHONE")
    cmbgender.Text = rs1.Fields("GENDER")
    cmbstream.Text = rs1.Fields("STREAM")
    dtbirth.Value = rs1.Fields("bIRTHDATE")
    cmbcast.Text = rs1.Fields("CAST")
    TXTOCCU.Text = rs1.Fields("OCCUPATION")
    TXTANNINC.Text = rs1.Fields("ANNUALINCOME")
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
If cmbcourse.Text = "" Then
    cnt = 1
End If

If cmby.Text = "" Then
    cnt = 1
End If

If txtRollNo.Text = "" Then
    cnt = 1
End If
If TXTLNAME.Text = "" Then
    cnt = 1
End If
If TXTFNAME.Text = "" Then
    cnt = 1
End If
If TXTSNAME.Text = "" Then
    cnt = 1
End If
If TXTLADD.Text = "" Then
    cnt = 1
End If
If TXTLPHONE.Text = "" Then
    cnt = 1
End If
If TXTPADD.Text = "" Then
    cnt = 1
End If
If TXTPPHONE.Text = "" Then
    cnt = 1
End If
If cmbgender.Text = "" Then
    cnt = 1
End If
If cmbstream.Text = "" Then
    cnt = 1
End If
If dtbirth.Value = "" Then
    cnt = 1
End If
If cmbcast.Text = "" Then
    cnt = 1
End If
If TXTOCCU.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If
If cnt = 0 Then
    Set rs = New adodb.Recordset
    rs.Open "select count(*) from studentmaster where courcecode = '" & cmbcourse.Text & "' and rollno = " & CDec(txtRollNo.Text) & " and year = '" & cmby.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into studentmaster values ('" & cmbcourse & "','" & dtjoin.Value & "','" & cmby.Text & "'," & CDec(txtRollNo.Text) & ",'" & TXTLNAME.Text & "','" & TXTFNAME.Text & "','" & TXTSNAME.Text & "','" & TXTLADD.Text & "'," & CDec(TXTLPHONE.Text) & ",'" & TXTPADD.Text & "'," & CDec(TXTPPHONE.Text) & ",'" & cmbgender.Text & "','" & cmbstream.Text & "','" & dtbirth.Value & "','" & cmbcast.Text & "','" & TXTOCCU.Text & "', " & CDec(TXTANNINC.Text) & ",'" & txtremark & "')"
        'MsgBox sql
        
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        
        
    Else
    Dim Yn As String
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update STUDENTMASTER set JOINDATE = '" & dtjoin.Value & "', ROLLNO = " & CDec(txtRollNo.Text) & ", LASTNAME = '" & TXTLNAME.Text & "',FIRSTNAME = '" & TXTFNAME.Text & "',PARENTNAME = '" & TXTSNAME.Text & "', ADDRESS = '" & TXTLADD.Text & "',PHONE = " & CDec(TXTLPHONE.Text) & ",PADDRESS = '" & TXTPADD.Text & "',PPHONE = " & CDec(TXTPPHONE.Text) & ",GENDER = '" & cmbgender.Text & "',STREAM = '" & cmbstream.Text & "', BIRTHDATE = '" & dtbirth.Value & "', CAST = '" & cmbcast.Text & "',OCCUPATION = '" & TXTOCCU.Text & "', ANNUALINCOME = " & CDec(TXTANNINC.Text) & ",REMARK = '" & txtremark & "' where COURCECODE = '" & cmbcourse.Text & "' and YEAR = '" & cmby.Text & "' and rollno = " & CDec(txtRollNo.Text) & ""
           ' MsgBox sql
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
        cmbcourse.Text = rs1.Fields("CourceCode")
    dtjoin.Value = rs1.Fields("JOINDATE")
    cmby.Text = rs1.Fields("year")
    txtRollNo.Text = rs1.Fields("RollNo")
    TXTLNAME.Text = rs1.Fields("LASTNAME")
    TXTFNAME.Text = rs1.Fields("FIRSTNAME")
    TXTSNAME.Text = rs1.Fields("PARENTNAME")
    TXTLADD.Text = rs1.Fields("ADDRESS")
    TXTLPHONE.Text = rs1.Fields("PHONE")
    TXTPADD.Text = rs1.Fields("PADDRESS")
    TXTPPHONE.Text = rs1.Fields("PPHONE")
    cmbgender.Text = rs1.Fields("GENDER")
    cmbstream.Text = rs1.Fields("STREAM")
    dtbirth.Value = rs1.Fields("bIRTHDATE")
    cmbcast.Text = rs1.Fields("CAST")
    TXTOCCU.Text = rs1.Fields("OCCUPATION")
    TXTANNINC.Text = rs1.Fields("ANNUALINCOME")
    txtremark.Text = rs1.Fields("REMARK")
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0
Set rs = New adodb.Recordset
rs.Open "select courcecode from courcemaster", adodc, adOpenKeyset, adLockOptimistic
If rs.EOF = False Then
    While rs.EOF = False
        cmbcourse.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
End If
rs.Close

Set rs1 = New adodb.Recordset
rs1.Open "select * from studentmaster order by courcecode,year ", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New adodb.Recordset
rs.Open "select * from studentmaster order by courcecode", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    cmbcourse.Text = rs.Fields("CourceCode")
    dtjoin.Value = rs.Fields("JOINDATE")
    cmby.Text = rs.Fields("year")
    txtRollNo.Text = rs.Fields("RollNo")
    TXTLNAME.Text = rs.Fields("LASTNAME")
    TXTFNAME.Text = rs.Fields("FIRSTNAME")
    TXTSNAME.Text = rs.Fields("PARENTNAME")
    TXTLADD.Text = rs.Fields("ADDRESS")
    TXTLPHONE.Text = rs.Fields("PHONE")
    TXTPADD.Text = rs.Fields("PADDRESS")
    TXTPPHONE.Text = rs.Fields("PPHONE")
    cmbgender.Text = rs.Fields("GENDER")
    cmbstream.Text = rs.Fields("STREAM")
    dtbirth.Value = rs.Fields("bIRTHDATE")
    cmbcast.Text = rs.Fields("CAST")
    TXTOCCU.Text = rs.Fields("OCCUPATION")
    TXTANNINC.Text = rs.Fields("ANNUALINCOME")
    txtremark.Text = rs.Fields("REMARK")
End If
End Sub

Private Sub TXTANNINC_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub TXTFNAME_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub TXTLNAME_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub TXTLPHONE_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub TXTOCCU_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub TXTPPHONE_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtROLLno_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Then Else KeyAscii = 0
End Sub

Private Sub TXTSNAME_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub enable_false()
    
    cmbcourse.Enabled = False
    dtjoin.Enabled = False
    cmby.Enabled = False
    txtRollNo.Enabled = False
    TXTLNAME.Enabled = False
    TXTFNAME.Enabled = False
    TXTSNAME.Enabled = False
    TXTLADD.Enabled = False
    TXTLPHONE.Enabled = False
    TXTPADD.Enabled = False
    TXTPPHONE.Enabled = False
    cmbgender.Enabled = False
    cmbstream.Enabled = False
    dtbirth.Enabled = False
    cmbcast.Enabled = False
    TXTOCCU.Enabled = False
    TXTANNINC.Enabled = False
    txtremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
    
End Sub

Private Sub enable_true()

    cmbcourse.Enabled = True
    dtjoin.Enabled = True
    cmby.Enabled = True
    txtRollNo.Enabled = True
    TXTLNAME.Enabled = True
    TXTFNAME.Enabled = True
    TXTSNAME.Enabled = True
    TXTLADD.Enabled = True
    TXTLPHONE.Enabled = True
    TXTPADD.Enabled = True
    TXTPPHONE.Enabled = True
    cmbgender.Enabled = True
    cmbstream.Enabled = True
    dtbirth.Enabled = True
    cmbcast.Enabled = True
    TXTOCCU.Enabled = True
    TXTANNINC.Enabled = True
    txtremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    cmbcourse.SetFocus
End Sub
