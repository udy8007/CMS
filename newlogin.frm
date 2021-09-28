VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form newlogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create & Delete Login"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "newlogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   6360
      Picture         =   "newlogin.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2550
      Width           =   1441
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   6360
      Picture         =   "newlogin.frx":306C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1441
   End
   Begin VB.CommandButton cmdhome 
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
      Left            =   6360
      Picture         =   "newlogin.frx":3936
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1441
   End
   Begin VB.CommandButton cmddel 
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
      Left            =   6360
      Picture         =   "newlogin.frx":3EB2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1740
      Width           =   1441
   End
   Begin VB.CommandButton CMDOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
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
      Left            =   6360
      Picture         =   "newlogin.frx":42F4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   930
      Width           =   1441
   End
   Begin VB.TextBox txtuserid 
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
      Left            =   2760
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2260
      Width           =   2175
   End
   Begin VB.ComboBox cmbdesig 
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
      ItemData        =   "newlogin.frx":45FE
      Left            =   2760
      List            =   "newlogin.frx":4611
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "DESIGNATION"
      Top             =   3420
      Width           =   2175
   End
   Begin VB.TextBox txtretype 
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
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2840
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc CollegeADODC 
      Height          =   330
      Left            =   4080
      Top             =   270
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      RecordSource    =   "SELECT * FROM Login"
      Caption         =   "New Login"
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
      BorderWidth     =   2
      Height          =   2655
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1425
      Width           =   5655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "A M C O S T "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   270
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   555
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Create and Delete Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   30
      TabIndex        =   8
      Top             =   870
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Id :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   360
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   360
      Left            =   600
      TabIndex        =   6
      Top             =   2260
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   360
      Left            =   600
      TabIndex        =   5
      Top             =   3420
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   360
      Left            =   600
      TabIndex        =   4
      Top             =   2840
      Width           =   1815
   End
End
Attribute VB_Name = "newlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Str As String
Dim flag As Boolean
Dim tmp As String
Dim i
Dim ans
Dim ch, ch1, dif

Private Sub cmdcan_Click()
    flag = False
    CollegeADODC.Recordset.Cancel
    txtuserid.Text = ""
    txtpass.Text = ""
    txtretype.Text = ""
    cmbdesig.Text = ""
         txtuserid.Visible = False
        txtpass.Visible = False
        txtretype.Visible = False
        cmbdesig.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        cmdcan.Visible = False
        CMDOK.Visible = False
        cmdAdd.Visible = True
        cmddel.Visible = True
        cmdhome.Visible = True
End Sub

Private Sub cmdAdd_Click()
    flag = True
    txtuserid.Visible = True
    txtpass.Visible = True
    txtretype.Visible = True
    cmbdesig.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    cmdAdd.Visible = False
    cmddel.Visible = False
    cmdhome.Visible = False
    CMDOK.Visible = True
    cmdcan.Visible = True
End Sub

Private Sub cmdOk_Click()
On Error GoTo AddError
If flag = True Then
    CollegeADODC.Recordset.MoveFirst
    Do While Not CollegeADODC.Recordset.EOF
        If CollegeADODC.Recordset.Fields(0) = UCase(Trim(txtuserid.Text)) Then
            MsgBox "User already exists."
            txtuserid.Text = ""
            txtuserid.SetFocus
            Call cmdAdd_Click
            Exit Sub
        Else
            CollegeADODC.Recordset.MoveNext
        End If
    Loop
    If txtpass.Text = txtretype.Text Then
    
        CollegeADODC.Recordset.AddNew
        CollegeADODC.Recordset.Fields(0) = UCase(Trim(txtuserid.Text))
        CollegeADODC.Recordset.Fields(1) = UCase(Trim(txtpass.Text))
        CollegeADODC.Recordset.Fields(2) = cmbdesig.Text
        CollegeADODC.Recordset.Update
        MsgBox "New Login has been added", vbInformation, "New Login"
        txtuserid.Text = ""
        txtpass.Text = ""
        txtretype.Text = ""
        cmbdesig.Text = ""
         txtuserid.Visible = False
        txtpass.Visible = False
        txtretype.Visible = False
        cmbdesig.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        cmdcan.Visible = False
        CMDOK.Visible = False
        cmdAdd.Visible = True
        cmddel.Visible = True
        cmdhome.Visible = True
        
    Else
        MsgBox "INVALID CONFIRMATION PASSWORD", vbCritical, "Error"
    End If
    Exit Sub
Else
    If flag = False Then
        If txtuserid.Text = "" Then
            MsgBox "Please, Enter UserName", vbExclamation, "Blank UserName"
            Exit Sub
        End If
        CollegeADODC.Recordset.MoveFirst
        Do While Not CollegeADODC.Recordset.EOF
            If txtuserid.Text = CollegeADODC.Recordset.Fields(0) Then
                If txtpass.Text = CollegeADODC.Recordset.Fields(1) Then
                    ans = MsgBox("Do you want to delete", vbYesNo)
                    If ans = vbYes Then
                        CollegeADODC.Recordset.Delete
                        MsgBox "Login deleted"
                        txtuserid.Text = ""
                        txtpass.Text = ""
                        txtretype.Text = ""
                        cmbdesig.Text = ""
                        txtuserid.Visible = False
                        txtpass.Visible = False
                        Label1.Visible = False
                        Label2.Visible = False
                        cmdcan.Visible = False
                        CMDOK.Visible = False
                        cmdAdd.Visible = True
                        cmddel.Visible = True
                        cmdhome.Visible = True
                        Exit Sub
                    Else
                        txtuserid.Text = ""
                        txtpass.Text = ""
                        txtretype.Text = ""
                        cmbdesig.Text = ""
                        txtuserid.Visible = False
                        txtpass.Visible = False
                        Label1.Visible = False
                        Label2.Visible = False
                        cmdcan.Visible = False
                        CMDOK.Visible = False
                        cmdAdd.Visible = True
                        cmddel.Visible = True
                        cmdhome.Visible = True
                        Exit Sub
                    End If
                Else
                    MsgBox "Invalid Password.:"
                    txtpass.Text = ""
                    txtpass.SetFocus
                    Call cmddel_Click
                    Exit Sub
                 End If
            Else
                CollegeADODC.Recordset.MoveNext
            End If
        Loop
        MsgBox "userid not found"
                    txtuserid.Visible = False
                    txtpass.Visible = False
                    Label1.Visible = False
                    Label2.Visible = False
                    cmdcan.Visible = False
                    CMDOK.Visible = False
                    cmdAdd.Visible = True
                    cmddel.Visible = True
                    cmdhome.Visible = True
    End If
End If
flag = False
AddError:
        MsgBox Err.Description, vbCritical, "Add Record Error"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmddel_Click()
    flag = False
    txtuserid.Visible = True
    txtpass.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    cmdAdd.Visible = False
    cmddel.Visible = False
    cmdhome.Visible = False
    CMDOK.Visible = True
    cmdcan.Visible = True
    
    
End Sub

Private Sub cmdhome_Click()
    Unload Me
    
    College.Show
End Sub

Private Sub Form_Activate()
    txtuserid.Text = ""
    txtpass.Text = ""
    txtretype.Text = ""
    cmbdesig.Text = ""
     txtuserid.Visible = False
    txtpass.Visible = False
    txtretype.Visible = False
    cmbdesig.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    cmdcan.Visible = False
    CMDOK.Visible = False
'    CollegeADODC.Recordset.AddNew
End Sub

