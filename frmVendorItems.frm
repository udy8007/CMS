VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVendorItems 
   BackColor       =   &H80000012&
   Caption         =   "Purchase Master"
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
   Icon            =   "frmVendorItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      Left            =   3540
      Picture         =   "frmVendorItems.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7170
      Width           =   1441
   End
   Begin VB.ComboBox VendorCodeList 
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
      Left            =   4320
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox ItemCodeList 
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
      Left            =   4320
      TabIndex        =   0
      Top             =   3030
      Width           =   1455
   End
   Begin VB.TextBox txtPRICE 
      DataField       =   "PRICE"
      DataSource      =   "CollegeADODC"
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
      Left            =   4320
      MaxLength       =   14
      TabIndex        =   2
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrev 
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
      Left            =   4395
      Picture         =   "frmVendorItems.frx":2BE4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1441
   End
   Begin VB.CommandButton cmdNext 
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
      Picture         =   "frmVendorItems.frx":32E6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
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
      Left            =   3540
      Picture         =   "frmVendorItems.frx":3C28
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7170
      Width           =   1441
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   6945
      Picture         =   "frmVendorItems.frx":3F32
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7170
      Width           =   1441
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
      Left            =   8640
      Picture         =   "frmVendorItems.frx":4374
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7170
      Width           =   1441
   End
   Begin VB.CommandButton cmdAdd 
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
      Left            =   1845
      Picture         =   "frmVendorItems.frx":48F0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7170
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
      Picture         =   "frmVendorItems.frx":51BA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7170
      Width           =   1441
   End
   Begin VB.TextBox txtREMARK 
      DataField       =   "REMARK"
      DataSource      =   "CollegeADODC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4320
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmVendorItems.frx":5A84
      Top             =   4860
      Width           =   4965
   End
   Begin VB.TextBox txtPURCHASEID 
      BackColor       =   &H00C0C0FF&
      DataField       =   "PURCHASEID"
      DataSource      =   "CollegeADODC"
      Enabled         =   0   'False
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2100
      Width           =   1455
   End
   Begin VB.TextBox txtQTY 
      DataField       =   "QTY"
      DataSource      =   "CollegeADODC"
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
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   4400
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker PurchaseDate 
      Height          =   360
      Left            =   4320
      TabIndex        =   22
      Top             =   2560
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   143851521
      CurrentDate     =   38427
   End
   Begin VB.Label lblFieldLabel 
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
      Index           =   6
      Left            =   2535
      TabIndex        =   24
      Top             =   4860
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   3
      Height          =   4035
      Left            =   660
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
      Left            =   480
      TabIndex        =   21
      Top             =   450
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Master"
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
      Left            =   2753
      TabIndex        =   20
      Top             =   1170
      Width           =   6375
   End
   Begin VB.Label lblVendor 
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   5910
      TabIndex        =   19
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   5910
      TabIndex        =   18
      Top             =   3015
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   5
      Left            =   2535
      TabIndex        =   17
      Top             =   4400
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
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
      Index           =   4
      Left            =   2535
      TabIndex        =   16
      Top             =   3940
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
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
      Index           =   3
      Left            =   2535
      TabIndex        =   15
      Top             =   3480
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code :"
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
      Index           =   2
      Left            =   2535
      TabIndex        =   14
      Top             =   3020
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date :"
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
      Index           =   1
      Left            =   2535
      TabIndex        =   13
      Top             =   2560
      Width           =   1635
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase ID :"
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
      Left            =   2535
      TabIndex        =   12
      Top             =   2100
      Width           =   1635
   End
End
Attribute VB_Name = "frmVendorItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset
Dim rs1 As adodb.Recordset
Dim rs2 As adodb.Recordset
Dim sql As String
Dim cnt As Integer

Private Sub cmdAdd_Click()
Call enable_true
Set rs = New adodb.Recordset
rs.Open "select * from purchaseitems", adodc, adOpenKeyset, adLockOptimistic
cnt = 0
Do While Not rs.EOF
    If cnt < rs.Fields("purchaseid") Then cnt = rs.Fields("purchaseid")
    rs.MoveNext
Loop
cnt = cnt + 1
txtPURCHASEID.Text = cnt

ItemCodeList.Text = ""
VendorCodeList.Text = ""
txtprice.Text = ""
txtqty.Text = ""
txtremark = ""
lblItem.Caption = ""
lblVendor.Caption = ""
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errhandler
If VendorCodeList.Enabled = True Then
    MsgBox "First save record and after delete the record...", vbCritical
    VendorCodeList.SetFocus
    Exit Sub
End If

sql = MsgBox("Do You Want To Delete Record . . . ???", vbYesNo, "Delete Record")
If sql = 6 Then
    sql = "delete purchaseitems where purchaseid = '" & txtPURCHASEID.Text & "'"
    adodc.Execute sql
    rs1.Requery
    MsgBox "Record Deleted . . . !!!", vbInformation, "Delete"
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

Private Sub cmdnext_Click()
If rs1.EOF = False Then
    rs1.MoveNext
If rs1.EOF = False Then
    txtPURCHASEID.Text = rs1("PURCHASEID")
    PurchaseDate.Value = rs1("PurchaseDate")
    ItemCodeList.Text = rs1("ITEMCODE")
    VendorCodeList.Text = rs1("VENDORCODE")
    txtprice.Text = rs1("PRICE")
    txtqty.Text = rs1("QTY")
    txtremark = rs1("REMARK")
    Set rs2 = New adodb.Recordset
    rs2.Open "select itemname from itemmaster where itemcode = " & ItemCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblItem.Caption = rs2.Fields(0)
    rs2.Close

    Set rs2 = New adodb.Recordset
    rs2.Open "select vendorname from vendormaster where vendorcode = " & VendorCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblVendor.Caption = rs2.Fields(0)
    rs2.Close
    End If
End If
End Sub
Private Sub cmdPrev_Click()
If rs1.BOF = False Then
    rs1.MovePrevious
If rs1.BOF = False Then
    txtPURCHASEID.Text = rs1("PURCHASEID")
    PurchaseDate.Value = rs1("PurchaseDate")
    ItemCodeList.Text = rs1("ITEMCODE")
    VendorCodeList.Text = rs1("VENDORCODE")
    txtprice.Text = rs1("PRICE")
    txtqty.Text = rs1("QTY")
    txtremark = rs1("REMARK")
    Set rs2 = New adodb.Recordset
    rs2.Open "select itemname from itemmaster where itemcode = " & ItemCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblItem.Caption = rs2.Fields(0)
    rs2.Close

    Set rs2 = New adodb.Recordset
    rs2.Open "select vendorname from vendormaster where vendorcode = " & VendorCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblVendor.Caption = rs2.Fields(0)
    rs2.Close
    End If
End If

End Sub

Private Sub cmdrefresh_Click()
Call Form_Load
End Sub

Private Sub cmdreturn_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
cnt = 0
If txtPURCHASEID.Text = "" Then
    cnt = 1
End If
If ItemCodeList.Text = "" Then
    cnt = 1
End If
If VendorCodeList.Text = "" Then
    cnt = 1
End If
If txtprice.Text = "" Then
    cnt = 1
End If
If txtqty.Text = "" Then
    cnt = 1
End If
If txtremark.Text = "" Then
    cnt = 1
End If

If cnt = 0 Then
    Set rs = New adodb.Recordset
    rs.Open "select count(purchaseid) from PURCHASEITEMS where purchaseid = '" & txtPURCHASEID.Text & "'", adodc, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        sql = "insert into PURCHASEITEMS values (" & CDec(txtPURCHASEID.Text) & ",'" & UCase(Format(PurchaseDate.Value, "dd-mmm-yy")) & "'," & CDec(ItemCodeList.Text) & "," & CDec(VendorCodeList.Text) & "," & CDec(txtprice.Text) & "," & CDec(txtqty.Text) & ",'" & txtremark.Text & "')"
       ' MsgBox sql
        adodc.Execute sql
        MsgBox "Record Saved . . . !!!", vbInformation, "Save"
        
        
    Else
        MsgBox "Record Already Exist . . . !!!", vbInformation, "Save"
        Yn = MsgBox("Do You Want To Update Record . . . ???", vbYesNo, "Save")
        If Yn = 6 Then
            sql = "Update purchaseitems set PurchaseDate = '" & UCase(Format(PurchaseDate.Value, "dd-mmm-yy")) & "', ITEMCODE = " & CDec(ItemCodeList.Text) & ",VENDORCODE = " & CDec(VendorCodeList.Text) & ",PRICE = " & CDec(txtprice.Text) & ",QTY = " & CDec(txtqty.Text) & ",REMARK = '" & txtremark.Text & "' where purchaseid = " & txtPURCHASEID.Text & ""
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
txtPURCHASEID.Text = rs1("PURCHASEID")
PurchaseDate.Value = rs1("PurchaseDate")
ItemCodeList.Text = rs1("ITEMCODE")
VendorCodeList.Text = rs1("VENDORCODE")
txtprice.Text = rs1("PRICE")
txtqty.Text = rs1("QTY")
txtremark = rs1("REMARK")
Set rs2 = New adodb.Recordset
rs2.Open "select itemname from itemmaster where itemcode = " & ItemCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
lblItem.Caption = rs2.Fields(0)
rs2.Close

Set rs2 = New adodb.Recordset
rs2.Open "select vendorname from vendormaster where vendorcode = " & VendorCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
lblVendor.Caption = rs2.Fields(0)
rs2.Close

Call enable_false
End Sub

Private Sub Form_Load()

Call connection
Call enable_false
cnt = 0
Set rs1 = New adodb.Recordset
rs1.Open "select * from PURCHASEITEMS order by PURCHASEID", adodc, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    rs1.MoveFirst
End If

Set rs2 = New adodb.Recordset
rs2.Open " select itemcode from itemmaster order by itemcode", adodc, adOpenKeyset, adLockOptimistic
While rs2.EOF = False
    ItemCodeList.AddItem rs2("Itemcode")
    rs2.MoveNext
Wend
rs2.Close

Set rs2 = New adodb.Recordset
rs2.Open " select vendorcode from vendormaster order by vendorcode", adodc, adOpenKeyset, adLockOptimistic
While rs2.EOF = False
    VendorCodeList.AddItem rs2("vendorcode")
    rs2.MoveNext
Wend
rs2.Close

Set rs = New adodb.Recordset
rs.Open "select * from PURCHASEITEMS order by PURCHASEID", adodc, adOpenKeyset, adLockOptimistic

If rs.EOF = False Then
    rs.MoveFirst
    txtPURCHASEID.Text = rs("PURCHASEID")
    PurchaseDate.Value = rs("PurchaseDate")
    ItemCodeList.Text = rs("ITEMCODE")
    VendorCodeList.Text = rs("VENDORCODE")
    txtprice.Text = rs("PRICE")
    txtqty.Text = rs("QTY")
    txtremark = rs("REMARK")
   
End If
End Sub

Private Sub ItemCodeList_LostFocus()
Set rs2 = New adodb.Recordset
    rs2.Open "select itemname from itemmaster where itemcode = " & ItemCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblItem.Caption = rs2.Fields(0)
    rs2.Close
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 46 And (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 And (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub VendorCodeList_LostFocus()
Set rs2 = New adodb.Recordset
    rs2.Open "select vendorname from vendormaster where vendorcode = " & VendorCodeList.Text & "", adodc, adOpenKeyset, adLockOptimistic
    lblVendor.Caption = rs2.Fields(0)
    rs2.Close
End Sub

Private Sub enable_true()
    PurchaseDate.Enabled = True
    ItemCodeList.Enabled = True
    lblItem.Enabled = True
    VendorCodeList.Enabled = True
    lblVendor.Enabled = True
    txtprice.Enabled = True
    txtqty.Enabled = True
    txtremark.Enabled = True
    cmdUpdate.Visible = True
    cmdedit.Visible = False
        
End Sub

Private Sub enable_false()
    PurchaseDate.Enabled = False
    ItemCodeList.Enabled = False
    lblItem.Enabled = False
    VendorCodeList.Enabled = False
    lblVendor.Enabled = False
    txtprice.Enabled = False
    txtqty.Enabled = False
    txtremark.Enabled = False
    cmdUpdate.Visible = False
    cmdedit.Visible = True
    
End Sub
