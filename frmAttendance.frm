VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAttendance 
   BackColor       =   &H80000008&
   Caption         =   "Student Attendance"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAttendance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbsubcode 
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
      Left            =   5850
      TabIndex        =   3
      Top             =   3690
      Width           =   1965
   End
   Begin VB.ComboBox cmbtcode 
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
      Left            =   5850
      TabIndex        =   2
      Top             =   3270
      Width           =   1965
   End
   Begin VB.ComboBox CourceCode 
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
      Left            =   5850
      TabIndex        =   0
      Top             =   2430
      Width           =   1965
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   4320
      Picture         =   "frmAttendance.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1441
   End
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
      Height          =   720
      Left            =   6135
      Picture         =   "frmAttendance.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1441
   End
   Begin VB.ComboBox cmbyear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      ItemData        =   "frmAttendance.frx":3028
      Left            =   5850
      List            =   "frmAttendance.frx":3035
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   2850
      Width           =   1965
   End
   Begin VB.ListBox lstAttendance 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   2490
      Left            =   2685
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   4230
      Width           =   6555
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   495
      Left            =   8730
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "SELECT * FROM Attendance "
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   5850
      TabIndex        =   7
      Top             =   1980
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   688
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
      Format          =   24903681
      CurrentDate     =   37588
   End
   Begin MSAdodcLib.Adodc Temp 
      Height          =   495
      Left            =   8790
      Top             =   2100
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
      RecordSource    =   "SELECT * FROM Attendance "
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   5205
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   10995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Attendance"
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
      Left            =   2768
      TabIndex        =   14
      Top             =   1050
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   435
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10995
   End
   Begin VB.Label Label1 
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
      Height          =   885
      Left            =   390
      TabIndex        =   13
      Top             =   360
      Width           =   10995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Date :             "
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
      Height          =   255
      Left            =   3930
      TabIndex        =   12
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Height          =   255
      Left            =   3930
      TabIndex        =   11
      Top             =   3690
      Width           =   1695
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
      Height          =   255
      Left            =   3930
      TabIndex        =   10
      Top             =   3270
      Width           =   1695
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
      Height          =   255
      Left            =   3930
      TabIndex        =   9
      Top             =   2840
      Width           =   1695
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
      Height          =   255
      Left            =   3930
      TabIndex        =   8
      Top             =   2410
      Width           =   1695
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim sql As String
Dim rs As ADODB.Recordset
Dim mI(200) As Long
Dim md As Integer
Dim mrollno As Integer

Private Sub cmbsubcode_LostFocus()
Set rs = New ADODB.Recordset
rs.Open " select * from studentmaster where courcecode = '" & CourceCode.Text & "' and year = " & cmbyear.Text & " order by rollno ", adodc, adOpenKeyset, adLockOptimistic
If rs.EOF = True Then
    MsgBox "No Student Found . . . !!!", vbInformation, "Student Record"
    Unload Me
    Exit Sub
End If
lstAttendance.Clear
While rs.EOF = False
    lstAttendance.AddItem rs("rollNo") & " | " & rs("LastName") & " " & rs("FirstName")
    rs.MoveNext
Wend

End Sub

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub CourceCode_lostFocus()
Set rs = New ADODB.Recordset
rs.Open "select subjectcode from subjectmaster where courcecode = '" & CourceCode.Text & "' order by subjectcode", adodc, adOpenKeyset, adLockOptimistic
cmbsubcode.Clear
While rs.EOF = False
    cmbsubcode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
lstAttendance.Clear
End Sub

Private Sub Form_Load()
Call connection
DTPicker1.Value = Date
Set rs = New ADODB.Recordset
rs.Open "select courcecode from courcemaster", adodc, adOpenKeyset, adLockOptimistic
CourceCode.Clear
While rs.EOF = False
    CourceCode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
Set rs = New ADODB.Recordset
rs.Open "select empcode from staffmaster", adodc, adOpenKeyset, adLockOptimistic
cmbtcode.Clear
While rs.EOF = False
    cmbtcode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close

End Sub

Private Sub lstAttendance_Click()
    If lstAttendance.Selected(lstAttendance.ListIndex) = True Then
        mI(lstAttendance.ListIndex) = 1
    Else
        mI(lstAttendance.ListIndex) = 0
    End If
End Sub


Private Sub cmdsave_Click()
tt = lstAttendance.ListCount

Dim Temp As String
Dim temp1 As String
Dim ch As String
Dim aid As String

For i = 0 To tt - 1
    temp1 = ""
    Temp = lstAttendance.List(i)
 
    For J = 1 To Len(Temp)
        ch = Mid(Temp, J, 1)
        If ch = "|" Then
            Exit For
        Else
            temp1 = temp1 & ch
        End If
    Next
    temp1 = Trim(temp1)
    Set rs = New ADODB.Recordset
    rs.Open "select COUNT(ATTENDANCEID) from ATTENDANCE", adodc, adOpenKeyset, adLockOptimistic
    cnt = rs.Fields(0)
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "select max(ATTENDANCEID) from ATTENDANCE", adodc, adOpenKeyset, adLockOptimistic
    
    
    If cnt <> 0 Then
        aid = rs.Fields(0)
        aid = aid + 1
    Else
        aid = 1
    End If
    rs.Close
    If lstAttendance.Selected(i) = True Then
        pre = 1
    Else
        pre = o
    End If
    Set rs = New ADODB.Recordset
    rs.Open "select count(*) from ATTENDANCE where CourceCode = '" & CourceCode.Text & "' and year = " & cmbyear.Text & " and ROLLNO = " & temp1 & " and teachercode = " & cmbtcode.Text & " and subjectCode = '" & cmbsubcode.Text & "' and attendancedate = '" & UCase(Format(DTPicker1.Value, "dd-mmm-yy")) & "'", adodc, adOpenKeyset, adLockOptimistic
    ct = rs.Fields(0)
    If ct <> 0 Then
        MsgBox Temp & " - - Record Already Exist . . . !!!", vbInformation, "Already Exist"
        rs.Close
    Else
        sql = "insert into ATTENDANCE values (" & CDec(aid) & ",'" & CourceCode.Text & "'," & CDec(cmbyear.Text) & "," & CDec(temp1) & "," & CDec(cmbtcode.Text) & ",'" & cmbsubcode.Text & "','" & UCase(Format(DTPicker1.Value, "dd-mmm-yy")) & "'," & CDec(pre) & ",'REMARK')"
        adodc.Execute sql
        
        rs.Close
    End If
Next
MsgBox "Record saved . . . !!!", vbInformation, "Save Attendance"
lstAttendance.Clear
End Sub


