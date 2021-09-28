VERSION 5.00
Begin VB.Form master_item 
   BackColor       =   &H00000000&
   Caption         =   "Item Master"
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
      Left            =   3630
      Picture         =   "master_item.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6840
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
      Left            =   5340
      Picture         =   "master_item.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6810
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
      Left            =   1920
      Picture         =   "master_item.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6810
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
      Left            =   8760
      Picture         =   "master_item.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6810
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
      Left            =   7050
      Picture         =   "master_item.frx":1B52
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6810
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
      Left            =   3630
      Picture         =   "master_item.frx":1F94
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
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
      Left            =   6195
      Picture         =   "master_item.frx":229E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
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
      Left            =   4500
      Picture         =   "master_item.frx":2BE0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1441
   End
   Begin VB.TextBox txtItemCode 
      BackColor       =   &H00C0C0FF&
      DataField       =   "ITEMCODE"
      DataSource      =   "ItemMaster"
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
      Left            =   4260
      TabIndex        =   13
      Top             =   2010
      Width           =   1335
   End
   Begin VB.TextBox txtItemName 
      DataField       =   "ITEMNAME"
      DataSource      =   "ItemMaster"
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
      Left            =   4260
      MaxLength       =   35
      TabIndex        =   0
      Top             =   2580
      Width           =   4485
   End
   Begin VB.TextBox txtprice 
      DataField       =   "PRICE"
      DataSource      =   "ItemMaster"
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
      Left            =   4260
      MaxLength       =   14
      TabIndex        =   1
      Top             =   3090
      Width           =   1335
   End
   Begin VB.TextBox txtmin 
      DataField       =   "MINQTY"
      DataSource      =   "ItemMaster"
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
      Left            =   5820
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtmax 
      DataField       =   "MAXQTY"
      DataSource      =   "ItemMaster"
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
      Left            =   7380
      MaxLength       =   4
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtqty 
      DataField       =   "QTYONHAND"
      DataSource      =   "ItemMaster"
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
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3870
      Width           =   1335
   End
   Begin VB.TextBox txtremark 
      DataField       =   "REMARK"
      DataSource      =   "ItemMaster"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4260
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "master_item.frx":32E2
      Top             =   4410
      Width           =   4485
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   3705
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   1830
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
      Left            =   390
      TabIndex        =   23
      Top             =   390
      Width           =   10995
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Master"
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
      Left            =   2505
      TabIndex        =   22
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Min :"
      DataSource      =   "ItemMaster"
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
      Left            =   5820
      TabIndex        =   21
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Max :"
      DataSource      =   "ItemMaster"
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
      Left            =   7380
      TabIndex        =   20
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "On Hand :"
      DataSource      =   "ItemMaster"
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
      Left            =   4260
      TabIndex        =   19
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code :"
      DataSource      =   "ItemMaster"
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
      Left            =   2610
      TabIndex        =   18
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
      DataSource      =   "ItemMaster"
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
      Left            =   2610
      TabIndex        =   17
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
      DataSource      =   "ItemMaster"
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
      Left            =   2610
      TabIndex        =   16
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label Label40 
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
      Left            =   2610
      TabIndex        =   15
      Top             =   4410
      Width           =   1455
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Quantity :"
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
      Left            =   2610
      TabIndex        =   14
      Top             =   3540
      Width           =   1455
   End
End
Attribute VB_Name = "master_item"
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
If txtItemName.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    txtItemName.SetFocus
    Exit Sub
End If
sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete itemmaster where itemcode = '" & txtItemCode.Text & "'"
    adodc.Execute sql
    rs1.Requery
    MsgBox " Record Deleted . . . !!!", vbInformation, "Delete"
    Call Form_Load
End If
errhandler:
MsgBox Err.Number
End Sub

Private Sub cmdedit_Click()
Call enable_true
End Sub

Private Sub cmdnew_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from itemmaster", adodc, adOpenKeyset, adLockOptimistic
cnt = 0
Do While Not rs.EOF
    If cnt < rs.Fields("itemcode") Then cnt = rs.Fields("itemcode")
    rs.MoveNext
Loop
cnt = cnt + 1
txtItemCode.Text = cnt
txtItemName.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtmin.Text = ""
txtmax.Text = ""
txtremark.Text = ""
Call enable_true

End Sub

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtItemCode.Text = rs1.Fields("ITEMCODE")
    txtItemName.Text = rs1.Fields("ITEMNAME")
    txtprice.Text = rs1.Fields("PRICE")
    txtqty.Text = rs1.Fields("QTYONHAND")
    txtmin.Text = rs1.Fields("MINQTY")
    txtmax.Text = rs1.Fields("MAXQTY")
    txtremark.Text = rs1.Fields("REMARK")

End If
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
    If rs1.BOF = False Then
    txtItemCode.Text = rs1.Fields("ITEMCODE")
    txtItemName.Text = rs1.Fields("ITEMNAME")
    txtprice.Text = rs1.Fields("PRICE")
    txtqty.Text = rs1.Fields("QTYONHAND")
    txtmin.Text = rs1.Fields("MINQTY")
    txtmax.Text = rs1.Fields("MAXQTY")
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

If txtItemName.Text = "" Then
    cnt = 1
End If
If txtprice.Text = "" Then
    cnt = 1
End If
If txtqty.Text = "" Then
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
    Dim Yn As String
    Set rs = New ADODB.Recordset
    rs.Open "select count(*) from itemmaster where itemcode = '" & txtItemCode.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
    sql = "insert into itemmaster values ('" & txtItemCode & "','" & txtItemName.Text & "','" & txtprice.Text & "','" & txtqty.Text & "', '" & txtmin.Text & "','" & txtmax.Text & "','" & txtremark.Text & "')"
    
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        
        
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update itemmaster set ITEMNAME =  '" & txtItemName.Text & "',price = '" & CDec(txtprice.Text) & "',QTYONHAND = '" & CDec(txtqty.Text) & "', MINQTY ='" & CDec(txtmin.Text) & "',MAXQTY ='" & CDec(txtmax.Text) & "',REMARK = '" & txtremark.Text & "' where ITEMCODE = '" & txtItemCode.Text & "'"
           MsgBox sql
           ' adodc.Execute sql
            MsgBox "Record Updated . . . !!!", vbInformation, "Update"
        End If
    End If
Else
    MsgBox "Can not Insert Null Value . . . !!!", vbCritical, "Invalid Data"
End If
cnt = 0
rs1.Requery
rs1.MoveFirst
txtItemCode.Text = rs1.Fields("ITEMCODE")
txtItemName.Text = rs1.Fields("ITEMNAME")
txtprice.Text = rs1.Fields("PRICE")
txtqty.Text = rs1.Fields("QTYONHAND")
txtmin.Text = rs1.Fields("MINQTY")
txtmax.Text = rs1.Fields("MAXQTY")
txtremark.Text = rs1.Fields("REMARK")
Call enable_false
End Sub

Private Sub Form_Load()
Call connection
Call enable_false
cnt = 0

Call connection
Set rs1 = New ADODB.Recordset
rs1.Open "select * from itemmaster order by itemcode", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs = New ADODB.Recordset

rs.Open "select * from itemmaster order by itemcode", adodc, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If rs.EOF = False Then
    txtItemCode.Text = rs.Fields("ITEMCODE")
    txtItemName.Text = rs.Fields("ITEMNAME")
    txtprice.Text = rs.Fields("PRICE")
    txtqty.Text = rs.Fields("QTYONHAND")
    txtmin.Text = rs.Fields("MINQTY")
    txtmax.Text = rs.Fields("MAXQTY")
    txtremark.Text = rs.Fields("REMARK")
    
End If
End Sub



Private Sub txtItemName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0

End Sub

Private Sub txtmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 And (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub


Private Sub enable_true()
    txtItemCode.Enabled = True
    txtItemName.Enabled = True
    txtprice.Enabled = True
    txtqty.Enabled = True
    txtmin.Enabled = True
    txtmax.Enabled = True
    txtremark.Enabled = True
    cmdsave.Visible = True
    cmdedit.Visible = False
    txtItemName.SetFocus

End Sub

Private Sub enable_false()
    txtItemCode.Enabled = False
    txtItemName.Enabled = False
    txtprice.Enabled = False
    txtqty.Enabled = False
    txtmin.Enabled = False
    txtmax.Enabled = False
    txtremark.Enabled = False
    cmdsave.Visible = False
    cmdedit.Visible = True
End Sub
