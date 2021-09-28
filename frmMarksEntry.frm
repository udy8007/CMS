VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMarksEntry 
   BackColor       =   &H80000007&
   Caption         =   "Marks Entry"
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
   Icon            =   "frmMarksEntry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbunit 
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
      ItemData        =   "frmMarksEntry.frx":27A2
      Left            =   6060
      List            =   "frmMarksEntry.frx":27C1
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2520
      Width           =   1905
   End
   Begin VB.ComboBox cmbteacher 
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
      Left            =   6060
      TabIndex        =   4
      Top             =   2970
      Width           =   1905
   End
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
      Left            =   2310
      TabIndex        =   2
      Top             =   2970
      Width           =   1905
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
      Left            =   2310
      TabIndex        =   1
      Top             =   2520
      Width           =   1905
   End
   Begin VB.PictureBox vsFlexArray1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3060
      ScaleHeight     =   3195
      ScaleWidth      =   5655
      TabIndex        =   7
      Top             =   3690
      Width           =   5715
   End
   Begin VB.CommandButton cmdReturn 
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
      Picture         =   "frmMarksEntry.frx":27E9
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7290
      Width           =   1441
   End
   Begin VB.CommandButton cmdUpdate 
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
      Picture         =   "frmMarksEntry.frx":2D65
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7290
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
      ItemData        =   "frmMarksEntry.frx":306F
      Left            =   9180
      List            =   "frmMarksEntry.frx":307C
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "3"
      Top             =   2520
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker dtExam 
      Height          =   360
      Left            =   9180
      TabIndex        =   6
      Top             =   2970
      Width           =   1890
      _ExtentX        =   3334
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
      Format          =   142802945
      CurrentDate     =   37716
   End
   Begin MSDataGridLib.DataGrid MarksGrid 
      Bindings        =   "frmMarksEntry.frx":3089
      Height          =   3255
      Left            =   3060
      TabIndex        =   13
      Top             =   3690
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16761024
      ForeColor       =   8388736
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Marks Entry "
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "COURCECODE"
         Caption         =   "COURCECODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EXAMDATE"
         Caption         =   "EXAMDATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "RESULTID"
         Caption         =   "RESULTID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ROLLNO"
         Caption         =   "ROLLNO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SUBJECTCODE"
         Caption         =   "SUBJECTCODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CHECKBY"
         Caption         =   "CHECKBY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "UNITS"
         Caption         =   "UNITS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "MARKS"
         Caption         =   "MARKS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "REMARK"
         Caption         =   "REMARK"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3315
      Left            =   3030
      Top             =   3660
      Width           =   5760
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marks Entry"
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
      Left            =   2775
      TabIndex        =   19
      Top             =   1170
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   1575
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   10995
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Height          =   360
      Left            =   8070
      TabIndex        =   17
      Top             =   2535
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
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
      Left            =   360
      TabIndex        =   16
      Top             =   2070
      Width           =   1740
   End
   Begin VB.Label lblId 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   2310
      TabIndex        =   15
      Top             =   2070
      Width           =   1890
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit :"
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
      Left            =   4410
      TabIndex        =   14
      Top             =   2535
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Date :"
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
      Left            =   8070
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Height          =   360
      Left            =   4410
      TabIndex        =   11
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Left            =   360
      TabIndex        =   0
      Top             =   2535
      Width           =   1740
   End
   Begin VB.Label Label7 
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
      Height          =   855
      Left            =   450
      TabIndex        =   18
      Top             =   450
      Width           =   10995
   End
End
Attribute VB_Name = "frmMarksEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim sql As String
Dim r1 As Integer

Private Sub cmbcoursecode_LostFocus()
Set rs = New adodb.Recordset
rs.Open "select  SUBJECTCODE from subjectmaster where COURCECODE = '" & cmbcoursecode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    cmbsubcode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub cmdreturn_Click()
    Unload Me
    'College.Show
    tree_main.Show
End Sub

Private Sub cmdUpdate_Click()
Dim i As Integer
Dim c1 As Integer
Dim c2 As Integer
Dim c3 As String
Dim sql As String
r1 = vsFlexArray1.Rows
For i = 1 To r1 - 2
    c1 = vsFlexArray1.TextMatrix(i, 1)
    c2 = vsFlexArray1.TextMatrix(i, 2)
    c3 = vsFlexArray1.TextMatrix(i, 3)
    Set rs = New adodb.Recordset
    rs.Open "select * from result where CourceCode = '" & cmbcoursecode.Text & "' and ExamDate = '" & UCase(Format(dtExam.Value, "dd-mmm-yy")) & "' and year = " & CDec(cmbyear.Text) & " and ROLLNO = " & CDec(c1) & " and SubjectCode = '" & cmbsubcode.Text & "' and CHECKBY = " & CDec(cmbteacher.Text) & " and UNITS = '" & cmbunit.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.EOF = True Then
        If c3 = "" Then
            c3 = "--"
        End If
        sql = "insert into result values ('" & cmbcoursecode.Text & "','" & UCase(Format(dtExam.Value, "dd-mmm-yy")) & "'," & CDec(lblId.Caption) & "," & CDec(cmbyear.Text) & "," & CDec(c1) & ",'" & cmbsubcode.Text & "'," & CDec(cmbteacher.Text) & ",'" & cmbunit.Text & "'," & CDec(c2) & ",'" & c3 & "')"
        adodc.Execute sql
    Else
        MsgBox "Record Exist . . . !!!", vbCritical, "Record Exist"
    End If
Next
Unload Me
End Sub

Private Sub Form_Load()
Call connection
vsFlexArray1.Clear
Set rs = New adodb.Recordset
rs.Open "select * from result", adodc, adOpenKeyset, adLockOptimistic
cnt = 0
Do While Not rs.EOF
    If cnt < rs.Fields("resultid") Then cnt = rs.Fields("resultid")
    rs.MoveNext
Loop
cnt = cnt + 1
lblId.Caption = cnt
rs.Close
Set rs = New adodb.Recordset
rs.Open "select courcecode from courcemaster order by courcecode", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    cmbcoursecode.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
Set rs = New adodb.Recordset
rs.Open "select empcode from staffmaster order by empcode", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    cmbteacher.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close

vsFlexArray1.ColWidth(0) = 1000
vsFlexArray1.ColWidth(1) = 1000
vsFlexArray1.ColWidth(2) = 1000
vsFlexArray1.ColWidth(3) = 2500

vsFlexArray1.TextMatrix(0, 1) = "Roll No."
vsFlexArray1.TextMatrix(0, 2) = "Marks"
vsFlexArray1.TextMatrix(0, 3) = "Remarks"

End Sub

Private Sub vsFlexArray1_GotFocus()
Dim i As Integer

i = 1
Set rs = New adodb.Recordset
rs.Open "select * from studentmaster where courcecode = '" & cmbcoursecode.Text & "' and year = '" & cmbyear.Text & "' order by rollno", adodc, adOpenKeyset, adLockOptimistic
While rs.EOF = False
    Temp = rs("rollno")
    vsFlexArray1.TextMatrix(i, 1) = Temp
    vsFlexArray1.Rows = vsFlexArray1.Rows + 1
    rs.MoveNext
    i = i + 1
Wend
End Sub

Private Sub vsFlexArray1_KeyPress(KeyAscii As Integer)
'If vsFlexArray1.Col = 3 And KeyAscii = 13 Then
'    vsFlexArray1.Rows = vsFlexArray1.Rows + 1
'    vsFlexArray1.Col = 1
'    vsFlexArray1.Row = vsFlexArray1.Rows - 1
'    vsFlexArray1.FocusRect = flexFocusHeavy
'End If
End Sub

Private Sub vsFlexArray1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If vsFlexArray1.Col = 1 Or vsFlexArray1.Col = 2 Then
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Then Else KeyAscii = 0
End If
End Sub
