VERSION 5.00
Begin VB.Form Master_vendor 
   BackColor       =   &H80000012&
   Caption         =   "Vendor Master"
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
      Picture         =   "Master_vendor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
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
      Left            =   5220
      Picture         =   "Master_vendor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Left            =   1800
      Picture         =   "Master_vendor.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   8670
      Picture         =   "Master_vendor.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Left            =   6930
      Picture         =   "Master_vendor.frx":1B52
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7380
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
      Picture         =   "Master_vendor.frx":1F94
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7380
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
      Picture         =   "Master_vendor.frx":229E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
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
      Left            =   4365
      Picture         =   "Master_vendor.frx":2BE0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1441
   End
   Begin VB.TextBox txtVendorName 
      DataField       =   "VENDORNAME"
      DataSource      =   "Vendor"
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
      Left            =   4455
      TabIndex        =   0
      Top             =   2409
      Width           =   3495
   End
   Begin VB.TextBox txtvendormail 
      DataField       =   "EMAIL"
      DataSource      =   "Vendor"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   4461
      Width           =   3495
   End
   Begin VB.TextBox txtvendoradd 
      DataField       =   "ADDRESS"
      DataSource      =   "Vendor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   4455
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2838
      Width           =   3495
   End
   Begin VB.TextBox txtvendorphone 
      DataField       =   "PHONE"
      DataSource      =   "Vendor"
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
      Left            =   4455
      TabIndex        =   2
      Top             =   4032
      Width           =   1695
   End
   Begin VB.TextBox txtVendorCode 
      BackColor       =   &H00C0C0FF&
      DataField       =   "VENDORCODE"
      DataSource      =   "Vendor"
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
      Left            =   4455
      TabIndex        =   4
      Top             =   1980
      Width           =   1695
   End
   Begin VB.TextBox txtvendorremark 
      DataField       =   "REMARK"
      DataSource      =   "Vendor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   4455
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "Master_vendor.frx":32E2
      Top             =   4890
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   4425
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   1860
      Width           =   10995
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Master"
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
      Left            =   2753
      TabIndex        =   19
      Top             =   1050
      Width           =   6375
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Code :"
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
      Left            =   2820
      TabIndex        =   18
      Top             =   1980
      Width           =   1500
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   2790
      TabIndex        =   17
      Top             =   2415
      Width           =   1500
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
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
      Left            =   2820
      TabIndex        =   16
      Top             =   4455
      Width           =   1500
   End
   Begin VB.Label Label23 
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
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2790
      TabIndex        =   15
      Top             =   2835
      Width           =   1500
   End
   Begin VB.Label Label24 
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
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   2790
      TabIndex        =   13
      Top             =   4035
      Width           =   1500
   End
   Begin VB.Label Label38 
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
      Left            =   2820
      TabIndex        =   12
      Top             =   4890
      Width           =   1500
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
      Left            =   480
      TabIndex        =   20
      Top             =   360
      Width           =   10995
   End
End
Attribute VB_Name = "Master_vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset
Dim rs1 As adodb.Recordset
Dim sql As String
Dim cnt As Integer

Private Sub cmdDelete_Click()
On Error GoTo errhandler
If txtVendorName.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtVendorName.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete vendormaster where vendorcode = '" & txtVendorCode.Text & "'"
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
Set rs = New adodb.Recordset
rs.Open "select * from vendormaster", adodc, adOpenKeyset, adLockOptimistic
cnt = 0
Do While Not rs.EOF
    If cnt < rs.Fields("vendorcode") Then cnt = rs.Fields("vendorcode")
    rs.MoveNext
Loop
cnt = cnt + 1

txtVendorCode.Text = cnt
txtVendorName.Text = ""
txtvendoradd.Text = ""
txtvendorphone.Text = ""
txtvendormail.Text = ""
txtvendorremark.Text = ""
End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtVendorCode.Text = rs1.Fields("VENDORCODE")
    txtVendorName.Text = rs1.Fields("VENDORNAME")
    txtvendoradd.Text = rs1.Fields("ADDRESS")
    txtvendorphone.Text = rs1.Fields("phone")
    txtvendormail.Text = rs1.Fields("email")
    txtvendorremark.Text = rs1.Fields("remark")
    End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
    txtVendorCode.Text = rs1.Fields("VENDORCODE")
    txtVendorName.Text = rs1.Fields("VENDORNAME")
    txtvendoradd.Text = rs1.Fields("ADDRESS")
    txtvendorphone.Text = rs1.Fields("phone")
    txtvendormail.Text = rs1.Fields("email")
    txtvendorremark.Text = rs1.Fields("remark")
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
If txtVendorName.Text = "" Then
    cnt = 1
End If
If txtvendoradd.Text = "" Then
    cnt = 1
End If
If txtvendorphone.Text = "" Then
    cnt = 1
End If
If txtvendorremark.Text = "" Then
    cnt = 1
End If
If txtvendormail.Text = "" Then
    cnt = 1
End If

If cnt = 0 Then
    Dim Yn As String
    Set rs = New adodb.Recordset
    rs.Open "select count(vendorcode) from vendormaster where vendorcode = '" & txtVendorCode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into vendormaster values ('" & txtVendorCode.Text & "','" & txtVendorName.Text & "','" & txtvendoradd.Text & "','" & txtvendorphone.Text & "','" & txtvendormail.Text & "','" & txtvendorremark.Text & "')"
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        
        
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update vendormaster set VENDORNAME = '" & txtVendorName.Text & "',ADDRESS = '" & txtvendoradd.Text & "', PHONE = '" & txtvendorphone.Text & "',EMAIL = '" & txtvendormail.Text & "',REMARK = '" & txtvendorremark.Text & "' where VENDORCODE = '" & txtVendorCode.Text & "'"
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
txtVendorCode.Text = rs1.Fields("VENDORCODE")
txtVendorName.Text = rs1.Fields("VENDORNAME")
txtvendoradd.Text = rs1.Fields("ADDRESS")
txtvendorphone.Text = rs1.Fields("phone")
txtvendormail.Text = rs1.Fields("email")
txtvendorremark.Text = rs1.Fields("remark")
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0
Set rs1 = New adodb.Recordset
rs1.Open "select * from vendormaster order by vendorcode", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New adodb.Recordset

rs.Open "select * from vendormaster order by vendorcode", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    txtVendorCode.Text = rs.Fields("VENDORCODE")
    txtVendorName.Text = rs.Fields("VENDORNAME")
    txtvendoradd.Text = rs.Fields("ADDRESS")
    txtvendorphone.Text = rs.Fields("phone")
    txtvendormail.Text = rs.Fields("email")
    txtvendorremark.Text = rs.Fields("remark")
    
End If
End Sub

Private Sub txtVendorName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub txtvendorphone_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub enable_true()
    
    txtVendorCode.Enabled = True
    txtVendorName.Enabled = True
    txtvendoradd.Enabled = True
    txtvendorphone.Enabled = True
    txtvendormail.Enabled = True
    txtvendorremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    txtVendorName.SetFocus

End Sub

Private Sub enable_false()
       
    txtVendorCode.Enabled = False
    txtVendorName.Enabled = False
    txtvendoradd.Enabled = False
    txtvendorphone.Enabled = False
    txtvendormail.Enabled = False
    txtvendorremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
    
End Sub
