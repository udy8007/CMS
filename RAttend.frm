VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm College 
   BackColor       =   &H8000000C&
   Caption         =   "College Management System"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "RAttend.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:40"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "29-09-2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   12726
            Text            =   "Developed by AKASH KACHHIA"
            TextSave        =   "Developed by AKASH KACHHIA"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnulogin 
      Caption         =   "&Login"
      Begin VB.Menu mnuchangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnucreate 
         Caption         =   "C&reate User"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete User"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnustaff 
         Caption         =   "&Staff Master"
      End
      Begin VB.Menu mnucourse 
         Caption         =   "C&ourse Master"
      End
      Begin VB.Menu mnusubject 
         Caption         =   "S&ubject Master"
      End
      Begin VB.Menu mnuholiday 
         Caption         =   "&Holiday Master"
      End
      Begin VB.Menu mnuitem 
         Caption         =   "&Item Master"
      End
      Begin VB.Menu mnuvendor 
         Caption         =   "&Vendor Master"
      End
   End
   Begin VB.Menu mnuattendence 
      Caption         =   "&Attendence"
      Begin VB.Menu mnustudentattendence 
         Caption         =   "&Student Attendence"
      End
      Begin VB.Menu mnustaffattndence 
         Caption         =   "Staff &Attendence"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuadmission 
         Caption         =   "&Admission"
      End
      Begin VB.Menu mnuMarksentry 
         Caption         =   "&Marks Entry"
      End
      Begin VB.Menu mnupurchaseitem 
         Caption         =   "&Purchase Item"
      End
   End
   Begin VB.Menu mnuquery 
      Caption         =   "&Query"
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnucalculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "&NotePad"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
      Begin VB.Menu mnulogout 
         Caption         =   "Logout"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "College"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuadmission_Click()
Call close_form
fyaddmission.Show
End Sub

Private Sub mnucalculator_Click()
'Temp = Shell("D:\college_12_3\project_12_3\Calculator.swf", 1)
'Calculator.Show
Temp = Shell("C:\Windows\Calc.Exe", 1)
End Sub

Private Sub mnuchangePassword_Click()
Call close_form
changepassfrm.Show

End Sub

Private Sub mnucourse_Click()
Call close_form
Master_Course.Show
End Sub

Private Sub mnucreatedelete_Click()
Call close_form
newlogin.Show
End Sub

Private Sub mnucreate_Click()
Call close_form
adduserfrm.Show
End Sub

Private Sub mnudelete_Click()
Call close_form
deluserfrm.Show
End Sub

Private Sub mnuholiday_Click()
Call close_form
Master_Holiday.Show
End Sub

Private Sub mnuitem_Click()
Call close_form
master_item.Show
End Sub

Private Sub mnulogout_Click()
Call close_form
Unload Me
End Sub

Private Sub mnuMarksentry_Click()
Call close_form
frmMarksEntry.Show
'frmMarksInfo.Show
End Sub

Private Sub mnuNotepad_Click()
Temp = Shell("C:\Windows\notepad.Exe", 1)

End Sub

Private Sub mnupurchaseitem_Click()
Call close_form
frmVendorItems.Show
End Sub

Private Sub mnuquery_Click()
Call close_form
query.Show
End Sub

'Option Explicit
'
'Private Sub cmd_Click()
'    Unload Me
'    frmStaffAttendance.Show vbModal
'End Sub
'
'Private Sub cmdAbout_Click()
'    frmsplash.Show vbModal
'    Unload Me
'End Sub
'
'Private Sub cmdAdmission_Click()
'    Unload Me
'    fyaddmission.Show
'End Sub
'
'Private Sub cmdChangePassword_Click()
'    Unload Me
'    changepass.Show
'End Sub
'
'Private Sub cmdExit_Click()
'
'End
'
'End Sub
'
'Private Sub cmdMarks_Click()
'Unload Me
'    frmMarksInfo.Show
'End Sub
'
''Private Sub cmdMaster_Click()
''    Unload Me
''
''    masters.Show vbModal
''End Sub
'
'Private Sub cmdNewLogin_Click()
'    Unload Me
'    newlogin.Show vbModal
'End Sub
'
'Private Sub cmdPurchaseItems_Click()
'    Unload Me
'    frmVendorItems.Show vbModal
'End Sub
'
'Private Sub cmdQuery_Click()
'    Unload Me
'    'frmReport.Show
'    query.Show
'
'End Sub
'
'Private Sub cmdStudent_Click()
'    Unload Me
'    frmNewAttend.Show
'End Sub
'
'Private Sub cmdTeachers_Click()
'    Unload Me
'    frmStaffAttendance.Show
'End Sub
'
'Private Sub MDIForm_Load()
'    If DESIG = "CLERK" Then
'        cmdNewLogin.Enabled = False
'        cmdMaster.Enabled = True
'        cmdAdmission.Enabled = True
'        cmdmarks.Enabled = True
'        cmdPurchaseItems.Enabled = True
'        cmdStudent.Enabled = True
'        cmdTeachers.Enabled = False
'        cmdChangePassword.Enabled = True
'        cmdQuery.Enabled = True
'    End If
'    If DESIG = "TEACHER" Then
'        cmdNewLogin.Enabled = False
'        cmdMaster.Enabled = True
'        cmdAdmission.Enabled = False
'        cmdmarks.Enabled = True
'        cmdPurchaseItems.Enabled = True
'        cmdStudent.Enabled = True
'        cmdTeachers.Enabled = False
'        cmdChangePassword.Enabled = True
'        cmdQuery.Enabled = True
'    End If
'    If DESIG = "LAB INCHARGE" Then
'        cmdNewLogin.Enabled = False
'        cmdMaster.Enabled = True
'        cmdAdmission.Enabled = False
'        cmdmarks.Enabled = True
'        cmdPurchaseItems.Enabled = True
'        cmdStudent.Enabled = False
'        cmdTeachers.Enabled = False
'        cmdChangePassword.Enabled = True
'        cmdQuery.Enabled = True
'    End If
'    If DESIG = "ADMIN" Then
'        cmdNewLogin.Enabled = True
'        cmdMaster.Enabled = True
'        cmdAdmission.Enabled = True
'        cmdmarks.Enabled = True
'        cmdPurchaseItems.Enabled = True
'        cmdStudent.Enabled = True
'        cmdTeachers.Enabled = True
'        cmdChangePassword.Enabled = True
'        cmdQuery.Enabled = True
'    End If
'    If DESIG = "PRINCIPAL" Then
'        cmdNewLogin.Enabled = True
'        cmdMaster.Enabled = True
'        cmdAdmission.Enabled = True
'        cmdmarks.Enabled = True
'        cmdPurchaseItems.Enabled = True
'        cmdStudent.Enabled = True
'        cmdTeachers.Enabled = True
'        cmdChangePassword.Enabled = True
'        cmdQuery.Enabled = True
'    End If
'End Sub
'
Private Sub mnustaff_Click()
Call close_form
Master_staff.Show
End Sub

Private Sub mnustaffattndence_Click()
Call close_form
frmStaffAttendance.Show
End Sub

Private Sub mnustudentattendence_Click()
Call close_form
frmAttendance.Show
End Sub

Private Sub mnusubject_Click()
Call close_form
Master_Subject.Show

End Sub

Private Sub mnuvendor_Click()
Call close_form
Master_vendor.Show
End Sub
