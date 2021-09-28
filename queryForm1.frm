VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form query 
   Caption         =   "Query Report"
   ClientHeight    =   5310
   ClientLeft      =   2355
   ClientTop       =   2850
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7950
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4650
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   763
      ImageHeight     =   510
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "queryForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "queryForm1.frx":DFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "queryForm1.frx":12C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "queryForm1.frx":222C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   1110
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5847
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "queryForm1.frx":33DB0
      Top             =   0
      Width           =   7350
   End
End
Attribute VB_Name = "query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
    'College.Show
    tree_main.Show
    
End Sub

Private Sub Form_Load()
TreeView1.ImageList = ImageList1
With TreeView1.Nodes
 
    .Add , 0, "r1", "Report"
    .Add "r1", tvwChild, "r11", "Detail"
    .Add "r11", tvwChild, "r111", "Student Information"
    .Add "r11", tvwChild, "r112", "Student Marks"
    .Add "r11", tvwChild, "r113", "Student Attendance"
    .Add "r1", tvwChild, "r12", "Staff Attendance"
    .Add "r1", tvwChild, "r13", "Vendor Item"
End With
End Sub

Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Key = "r111" Then frmReport1.Show
If TreeView1.SelectedItem.Key = "r112" Then frmReport2.Show
If TreeView1.SelectedItem.Key = "r113" Then frmReport3.Show
If TreeView1.SelectedItem.Key = "r12" Then staffreport.Show
If TreeView1.SelectedItem.Key = "r13" Then VendorItems.Show
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
    If TreeView1.SelectedItem.Key = "r111" Then frmReport1.Show
    If TreeView1.SelectedItem.Key = "r112" Then frmReport2.Show
    If TreeView1.SelectedItem.Key = "r113" Then frmReport3.Show
    If TreeView1.SelectedItem.Key = "r12" Then staffreport.Show
    If TreeView1.SelectedItem.Key = "r13" Then VendorItems.Show
End If
End Sub
