VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   240
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   6360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image3 
      Height          =   2100
      Left            =   2520
      Picture         =   "splash.frx":000C
      Top             =   3240
      Width           =   3885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AKASH KACHHIA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   3480
      TabIndex        =   3
      Top             =   5160
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   5520
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Image Image2 
      Height          =   1950
      Left            =   150
      Picture         =   "splash.frx":4F8B
      Top             =   1440
      Width           =   8580
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   240
      Picture         =   "splash.frx":3B735
      Top             =   120
      Width           =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   -720
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Timer1_Timer()
    Unload Me
    Load frmlogin
    frmlogin.Show
    
End Sub

Private Sub Timer2_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = 100 Then
        Unload Me
        Load frmlogin
        frmlogin.Show
    End If
End Sub
