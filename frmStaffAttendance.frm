VERSION 5.00
Begin VB.Form frmStaffAttendance 
   BackColor       =   &H00000000&
   Caption         =   "Staff Attendance"
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
   Icon            =   "frmStaffAttendance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
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
      Left            =   6075
      Picture         =   "frmStaffAttendance.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
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
      Left            =   4365
      Picture         =   "frmStaffAttendance.frx":2D1E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1441
   End
   Begin VB.CheckBox chkPresent 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   2715
      Width           =   1750
   End
   Begin VB.TextBox txtEmpCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   390
      TabIndex        =   5
      Top             =   2700
      Width           =   975
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1530
      TabIndex        =   4
      Top             =   2700
      Width           =   3795
   End
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7500
      MaxLength       =   255
      TabIndex        =   1
      Text            =   "--"
      Top             =   2700
      Width           =   4110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Present / Absent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Left            =   5385
      TabIndex        =   13
      Top             =   2340
      Width           =   1785
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Attendance"
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
      Left            =   2475
      TabIndex        =   12
      Top             =   1110
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   855
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   10995
   End
   Begin VB.Label Label5 
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
      Left            =   390
      TabIndex        =   11
      Top             =   390
      Width           =   10995
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00000000&
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
      Height          =   375
      Left            =   2190
      TabIndex        =   10
      Top             =   1890
      Width           =   1965
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Date :"
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
      Left            =   390
      TabIndex        =   9
      Top             =   1890
      Width           =   2265
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Emp Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   315
      Left            =   390
      TabIndex        =   8
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   315
      Left            =   1530
      TabIndex        =   7
      Top             =   2340
      Width           =   3795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   315
      Left            =   7275
      TabIndex        =   6
      Top             =   2340
      Width           =   4110
   End
End
Attribute VB_Name = "frmStaffAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim i As Integer
Dim Temp As Integer
Dim TotalRecord As Integer
Dim rs1 As New ADODB.Recordset
Dim sql As String
Dim tmp As Integer
Dim rs As ADODB.Recordset
Dim rem1 As String
Dim cnt As Integer
Dim tt As Integer
Dim sv As Integer
'
'
'
'
'Public Sub AddInfo()
'    Set RS = New ADODB.Recordset
'
'
'    SQL = StaffAttendView.RecordSource
'    StaffAttendView.RecordSource = "SELECT * FROM STAFFATTENDANCE"
'    StaffAttendView.Refresh
'
'    TotalRecord = StaffAttendView.Recordset.RecordCount
'    StaffAttendView.RecordSource = SQL
'    StaffAttendView.Refresh
'
'    CollegeADODC.RecordSource = "SELECT * FROM  staffattendance "
'    CollegeADODC.Refresh
'
'    tmp = 0
'    Do While Not CollegeADODC.Recordset.EOF
'            If tmp < CollegeADODC.Recordset.Fields(0) Then tmp = CollegeADODC.Recordset.Fields(0)
'            CollegeADODC.Recordset.MoveNext
'    Loop
'    'RS2.Open "SELECT * FROM STAFFATTENDANCE", CollegeADODC.ConnectionString, adOpenDynamic, adLockPessimistic
'    RS1.Open "SELECT * FROM EMPCODEVIEW", CollegeADODC.ConnectionString, adOpenDynamic, adLockPessimistic
'
'    CollegeADODC.Refresh
'    CollegeADODC.Recordset.MoveFirst
'
'    Do Until RS1.EOF
'        With CollegeADODC.Recordset
'            Debug.Print .RecordCount
'            TotalRecord = TotalRecord + 1
'            .Find .Fields(1).Name & "=" & RS1.Fields(0), 0, adSearchForward, 1
'            If .EOF Then
'                .AddNew
'                .Fields(0) = tmp
'                Temp = RS1.Fields(0)
'                .Fields(1) = Temp
'
'                .Fields(2) = lblDate.Caption
'                .Fields("PRESENT") = IIf(chkPresent(I).Value = True, 1, 0)
'                .Update
'            End If
'
'            RS1.MoveNext
'        End With
'    Loop
'End Sub
'
'Private Sub chkPresent_Click(Index As Integer)
'On Error GoTo PresentError
'    Dim CN As New ADODB.connection, mSQL As String, mPresent As Integer
'
'    If txtEmpCode(Index).Text = "" Then Exit Sub
'    If chkPresent(Index).Value = 1 Then
'        mPresent = 1
'    Else
'        mPresent = 0
'    End If
'        mSQL = "UPDATE STAFFATTENDANCE Set PRESENT = " & mPresent & " WHERE (TO_CHAR(ATTENDANCEDATE, 'MM/DD/YYYY') " _
'          & " = TO_CHAR(SYSDATE, 'MM/DD/YYYY')) AND staffcode = " & txtEmpCode(Index).Text

Private Sub Command1_Click()

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmdsave_Click()
tt = 0
sv = 0
Set rs = New ADODB.Recordset
rs.Open " select count(*) from staffattendance", adodc, adOpenKeyset, adLockOptimistic
'RS.Requery
If rs.EOF = True Then
    scode = 0
Else
    scode = rs.Fields(0)
End If
For i = 0 To Temp - 1
    scode = scode + 1
    stcode = txtEmpCode(i).Text
    dt = Date
    If chkPresent(i).Value = 1 Then
        chk = 1
    Else
        chk = 0
    End If
    
    rem1 = txtRemark(i).Text
    Set rs1 = New ADODB.Recordset
    rs1.Open "select count(*) from STAFFATTENDANCE where ATTENDANCEDATE  = '" & Format(dt, "dd/mmm/yy") & "' and staffcode = " & Val(txtEmpCode(i).Text) & "", adodc, adOpenKeyset, adLockOptimistic
    If rs1.Fields(0) > 0 Then
        MsgBox "Already Exist . . . !!!", vbCritical, "Can Not Save"
        Exit Sub
    Else
        If tt = 0 Then
            Temp = MsgBox("Do You Want To save . . . ???", vbYesNo, "Save")
            tt = 1
        End If
        If Temp = 6 Then
            sql = "insert into staffattendance values (" & scode & "," & stcode & ",'" & Format(dt, "dd/mmm/yy") & "'," & chk & ",'" & rem1 & "')"
            adodc.Execute sql
            sv = 1
            'MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        Else
            Exit Sub
        End If
        
    End If
    If sv = 0 Then
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        sv = 1
    End If
Next
rs.Requery
rs.Close
Unload Me
Load tree_main
tree_main.Show
End Sub

'
'    CN.Open CollegeADODC.ConnectionString, "ty02", "ty02"
'    CN.Execute mSQL
'    If mPresent = 1 Then
'        MsgBox txtName(Index).Text & "'s Attendance :-  PRESENT", vbInformation, "Present"
'    Else
'        MsgBox txtName(Index).Text & "'s Attendance :-  ABSCENT", vbInformation, "Abscent"
'    End If
'    Exit Sub
'PresentError:
'    'MsgBox Err.Description, vbCritical, "Save Error"
'End Sub
'Private Sub cmdClose_Click()
'    Unload Me
'    College.Show
'End Sub
'
'Public Sub ShowCheckBox()
'On Error GoTo ShowError
'    With CollegeADODC.Recordset
'        .MoveFirst
'
'    Do Until .EOF
'        If I <> 0 Then NewPresent
'
'        chkPresent(I).Value = .Fields(3) & ""
'
'        If I <> 0 Then NewEmpCode
'
'        txtEmpCode(I).Text = .Fields(1) & ""
'
'        If I <> 0 Then NewName
'
'        Debug.Print .Fields(0)
'        StaffAttendView.Recordset.Find StaffAttendView.Recordset.Fields(0).Name & "=" & .Fields(1)
'
'        txtName(I).Text = StaffAttendView.Recordset.Fields(1)
'
'        If I <> 0 Then NewRemark
'
'
'        txtREMARK(I).Text = .Fields(4) & ""
'        Debug.Print CollegeADODC.RecordSource
'        CollegeADODC.Recordset.MoveNext
'        I = I + 1
'    Loop
'    End With
'Exit Sub
'ShowError:
'        MsgBox Err.Description, vbCritical, "Record Show Error"
'End Sub
'
'Public Sub NewName()
'    Load txtName(I)
'    txtName(I).Top = txtName(I - 1).Top + txtName(I - 1).Height + 10
'    txtName(I).Visible = True
'End Sub
'
'Public Sub NewRemark()
'    Load txtREMARK(I)
'    txtREMARK(I).Top = txtREMARK(I - 1).Top + txtREMARK(I - 1).Height + 10
'    txtREMARK(I).Visible = True
'End Sub
'
'Public Sub NewPresent()
'Load chkPresent(I)
'    chkPresent(I).Top = chkPresent(I - 1).Top + chkPresent(I - 1).Height + 10
'    chkPresent(I).Visible = True
'End Sub
'
'Public Sub NewEmpCode()
'    Load txtEmpCode(I)
'    txtEmpCode(I).Top = txtEmpCode(I - 1).Top + txtEmpCode(I - 1).Height + 10
'    txtEmpCode(I).Visible = True
'End Sub
'
'Public Sub UnLoadCheckBox()
'    Dim J As Integer
'    txtEmpCode(0) = ""
'    txtName(0) = ""
'    chkPresent(0).Value = False
'    txtREMARK(0).Text = ""
'    For J = txtEmpCode.LBound + 1 To txtEmpCode.UBound
'        Unload txtEmpCode(J)
'        Unload txtName(J)
'        Unload chkPresent(J)
'        Unload txtREMARK(J)
'    Next J
'End Sub
'
'Private Sub Command1_Click()
'    MsgBox "Record has been save"
'End Sub
'
'Private Sub Form_Load()
'Call connection
'    lblDate.Caption = Date
'    I = 0
'    AddInfo
'    ShowCheckBox
'End Sub
'
'
Private Sub Form_Load()
lblDate.Caption = Date
Call connection
Set rs = New ADODB.Recordset
rs.Open "select count(*) from staffmaster order by empcode", adodc, adOpenKeyset, adLockOptimistic
Temp = rs.Fields(0)
If cnt = 0 Then
    For i = 1 To Temp - 1
        Load txtEmpCode(i)
        Load txtName(i)
        Load txtRemark(i)
        Load chkPresent(i)
        txtEmpCode(i).tOp = txtEmpCode(i - 1).tOp + txtEmpCode(i - 1).Height + 10
        txtName(i).tOp = txtName(i - 1).tOp + txtName(i - 1).Height + 10
        txtRemark(i).tOp = txtRemark(i - 1).tOp + txtRemark(i - 1).Height + 10
        chkPresent(i).tOp = (chkPresent(i - 1).tOp + chkPresent(i - 1).Height + 30) + 15
        txtEmpCode(i).Visible = True
        txtName(i).Visible = True
        txtRemark(i).Visible = True
        chkPresent(i).Visible = True
    Next
End If
    rs.Close
    Set rs = New ADODB.Recordset
rs.Open " SELECT * FROM STAFFMASTER order by empcode", adodc, adOpenKeyset, adLockOptimistic
Dim tt As Integer
tt = 0
While rs.EOF = False
    txtEmpCode(tt).Text = rs.Fields("empcode")
    txtName(tt).Text = rs.Fields("lastname") & " " & rs.Fields("parentname") & " " & rs.Fields("firstname")
    If rs.Fields("remark") <> "" Then
        txtRemark(tt).Text = rs.Fields("remark")
        
    Else
        txtRemark(tt).Text = " "
    End If
    tt = tt + 1
    rs.MoveNext
Wend
rs.Close
End Sub


