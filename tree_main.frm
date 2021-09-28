VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tree_main 
   BackColor       =   &H00000000&
   Caption         =   "College Management System     "
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1560
      Top             =   7080
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8205
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   14473
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8220
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   15293
            Picture         =   "tree_main.frx":0000
            TextSave        =   "14:26"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "28-09-2021"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   5220
      Left            =   3480
      Picture         =   "tree_main.frx":0752
      Top             =   2520
      Width           =   9900
   End
   Begin VB.Image Image2 
      Height          =   1530
      Left            =   5400
      Picture         =   "tree_main.frx":1D8F4
      Top             =   360
      Width           =   5445
   End
End
Attribute VB_Name = "tree_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp

Private Sub Form_Load()

With TreeView1.Nodes
    .Add , 0, "r1", "Master"
    .Add "r1", tvwChild, "r11", "Course Master"
    .Add "r1", tvwChild, "r12", "Subject Master"
    .Add "r1", tvwChild, "r13", "Staff Master"
    '.Add "r1", tvwChild, "r14", "Item Master"
    .Add "r1", tvwChild, "r15", "Vendor Master"
    .Add "r1", tvwChild, "r16", "Holiday Master"
    .Add , 0, "r2", "Attendance"
    .Add "r2", tvwChild, "r21", "Student Attendance"
    .Add "r2", tvwChild, "r22", "Staff Attendance"
    .Add , 0, "r3", "Transaction"
    .Add "r3", tvwChild, "r31", "Admission"
    .Add "r3", tvwChild, "r32", "Marks Entry"
    .Add "r3", tvwChild, "r33", "Purchase Master"
    .Add , 0, "r4", "Report"
    .Add "r4", tvwChild, "r41", "Detail"
    .Add "r41", tvwChild, "r411", "Student Information"
    .Add "r41", tvwChild, "r412", "Student Marks"
    .Add "r41", tvwChild, "r413", "Student Attendance"
    .Add "r4", tvwChild, "r42", "Staff Attendance"
    .Add "r4", tvwChild, "r43", "Vendor Item"
    .Add , 0, "r5", "Tools"
    .Add "r5", tvwChild, "r51", "Calculator"
    .Add "r5", tvwChild, "r52", "Notepad"
    .Add , 0, "r6", "Login"
    .Add "r6", tvwChild, "r61", "Add User"
    .Add "r6", tvwChild, "r62", "Delete User"
    .Add "r6", tvwChild, "r63", "Change Password"
    .Add , 0, "r7", "About CMS"
    
    .Add , 0, "r8", "Exit"
    
End With
End Sub



Private Sub Timer1_Timer()
'tree_main.Caption = Mid(tree_main.Caption, 2) + Mid(tree_main.Caption, 1, 1)

End Sub

Private Sub TreeView1_DblClick()
If DESIG = "ADMIN" Then
If TreeView1.SelectedItem.Key = "r11" Then Master_Course.Show
If TreeView1.SelectedItem.Key = "r12" Then Master_Subject.Show
If TreeView1.SelectedItem.Key = "r13" Then Master_staff.Show
If TreeView1.SelectedItem.Key = "r14" Then master_item.Show
If TreeView1.SelectedItem.Key = "r15" Then Master_vendor.Show
If TreeView1.SelectedItem.Key = "r16" Then Master_Holiday.Show

If TreeView1.SelectedItem.Key = "r21" Then frmAttendance.Show
If TreeView1.SelectedItem.Key = "r22" Then frmStaffAttendance.Show

If TreeView1.SelectedItem.Key = "r31" Then fyaddmission.Show
If TreeView1.SelectedItem.Key = "r32" Then frmMarksEntry.Show
If TreeView1.SelectedItem.Key = "r33" Then frmVendorItems.Show

If TreeView1.SelectedItem.Key = "r411" Then frmReport1.Show
If TreeView1.SelectedItem.Key = "r412" Then frmReport2.Show
If TreeView1.SelectedItem.Key = "r413" Then frmReport3.Show
If TreeView1.SelectedItem.Key = "r42" Then staffreport.Show
If TreeView1.SelectedItem.Key = "r43" Then VendorItems.Show

If TreeView1.SelectedItem.Key = "r51" Then Temp = Shell("C:\Windows\Calc.Exe", 1)
If TreeView1.SelectedItem.Key = "r52" Then Temp = Shell("C:\Windows\notepad.Exe", 1)

If TreeView1.SelectedItem.Key = "r61" Then adduserfrm.Show
If TreeView1.SelectedItem.Key = "r62" Then deluserfrm.Show
If TreeView1.SelectedItem.Key = "r63" Then changepassfrm.Show

If TreeView1.SelectedItem.Key = "r7" Then frmTextEffect.Show

If TreeView1.SelectedItem.Key = "r8" Then End
End If


If DESIG = "TEACHER" Then
If TreeView1.SelectedItem.Key = "r11" Then Master_Course.Show
If TreeView1.SelectedItem.Key = "r12" Then Master_Subject.Show
If TreeView1.SelectedItem.Key = "r13" Then Master_staff.Show
If TreeView1.SelectedItem.Key = "r14" Then master_item.Show
If TreeView1.SelectedItem.Key = "r15" Then Master_vendor.Show
If TreeView1.SelectedItem.Key = "r16" Then Master_Holiday.Show

If TreeView1.SelectedItem.Key = "r21" Then frmAttendance.Show
If TreeView1.SelectedItem.Key = "r22" Then
    MsgBox "You can not use this facility", vbCritical
    Exit Sub
End If
If TreeView1.SelectedItem.Key = "r31" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r32" Then frmMarksEntry.Show
If TreeView1.SelectedItem.Key = "r33" Then frmVendorItems.Show

If TreeView1.SelectedItem.Key = "r411" Then frmReport1.Show
If TreeView1.SelectedItem.Key = "r412" Then frmReport2.Show
If TreeView1.SelectedItem.Key = "r413" Then frmReport3.Show
If TreeView1.SelectedItem.Key = "r42" Then staffreport.Show
If TreeView1.SelectedItem.Key = "r43" Then VendorItems.Show

If TreeView1.SelectedItem.Key = "r51" Then Temp = Shell("C:\Windows\Calc.Exe", 1)
If TreeView1.SelectedItem.Key = "r52" Then Temp = Shell("C:\Windows\notepad.Exe", 1)

If TreeView1.SelectedItem.Key = "r61" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r62" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r63" Then changepassfrm.Show

If TreeView1.SelectedItem.Key = "r7" Then frmTextEffect.Show

If TreeView1.SelectedItem.Key = "r8" Then End
End If

End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
If DESIG = "ADMIN" Then
If KeyAscii = 13 Or KeyAscii = 32 Then
    If TreeView1.SelectedItem.Key = "r11" Then Master_Course.Show
    If TreeView1.SelectedItem.Key = "r12" Then Master_Subject.Show
    If TreeView1.SelectedItem.Key = "r13" Then Master_staff.Show
    If TreeView1.SelectedItem.Key = "r14" Then master_item.Show
    If TreeView1.SelectedItem.Key = "r15" Then Master_vendor.Show
    If TreeView1.SelectedItem.Key = "r16" Then Master_Holiday.Show
    
    If TreeView1.SelectedItem.Key = "r21" Then frmAttendance.Show
    If TreeView1.SelectedItem.Key = "r22" Then frmStaffAttendance.Show
    
    If TreeView1.SelectedItem.Key = "r31" Then fyaddmission.Show
    If TreeView1.SelectedItem.Key = "r32" Then frmMarksEntry.Show
    If TreeView1.SelectedItem.Key = "r33" Then frmVendorItems.Show
    
    If TreeView1.SelectedItem.Key = "r411" Then frmReport1.Show
    If TreeView1.SelectedItem.Key = "r412" Then frmReport2.Show
    If TreeView1.SelectedItem.Key = "r413" Then frmReport3.Show
    If TreeView1.SelectedItem.Key = "r42" Then staffreport.Show
    If TreeView1.SelectedItem.Key = "r43" Then VendorItems.Show
    
    If TreeView1.SelectedItem.Key = "r51" Then Temp = Shell("C:\Windows\Calc.Exe", 1)
    If TreeView1.SelectedItem.Key = "r52" Then Temp = Shell("C:\Windows\notepad.Exe", 1)
    
    If TreeView1.SelectedItem.Key = "r61" Then adduserfrm.Show
    If TreeView1.SelectedItem.Key = "r62" Then deluserfrm.Show
    If TreeView1.SelectedItem.Key = "r63" Then changepassfrm.Show
    
    If TreeView1.SelectedItem.Key = "r7" Then frmTextEffect.Show
    
    If TreeView1.SelectedItem.Key = "r8" Then End
End If
End If


If DESIG = "TEACHER" Then
If KeyAscii = 13 Or KeyAscii = 32 Then
If TreeView1.SelectedItem.Key = "r11" Then Master_Course.Show
If TreeView1.SelectedItem.Key = "r12" Then Master_Subject.Show
If TreeView1.SelectedItem.Key = "r13" Then Master_staff.Show
If TreeView1.SelectedItem.Key = "r14" Then master_item.Show
If TreeView1.SelectedItem.Key = "r15" Then Master_vendor.Show
If TreeView1.SelectedItem.Key = "r16" Then Master_Holiday.Show

If TreeView1.SelectedItem.Key = "r21" Then frmAttendance.Show
If TreeView1.SelectedItem.Key = "r22" Then MsgBox "You can not use this facility", vbCritical

If TreeView1.SelectedItem.Key = "r31" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r32" Then frmMarksEntry.Show
If TreeView1.SelectedItem.Key = "r33" Then frmVendorItems.Show

If TreeView1.SelectedItem.Key = "r411" Then frmReport1.Show
If TreeView1.SelectedItem.Key = "r412" Then frmReport2.Show
If TreeView1.SelectedItem.Key = "r413" Then frmReport3.Show
If TreeView1.SelectedItem.Key = "r42" Then staffreport.Show
If TreeView1.SelectedItem.Key = "r43" Then VendorItems.Show

If TreeView1.SelectedItem.Key = "r51" Then Temp = Shell("C:\Windows\Calc.Exe", 1)
If TreeView1.SelectedItem.Key = "r52" Then Temp = Shell("C:\Windows\notepad.Exe", 1)

If TreeView1.SelectedItem.Key = "r61" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r62" Then MsgBox "You can not use this facility", vbCritical
If TreeView1.SelectedItem.Key = "r63" Then changepassfrm.Show

If TreeView1.SelectedItem.Key = "r7" Then frmTextEffect.Show

If TreeView1.SelectedItem.Key = "r8" Then End
End If
End If
End Sub
