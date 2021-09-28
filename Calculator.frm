VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6060
   Begin VB.PictureBox ShockwaveFlash2 
      Height          =   855
      Left            =   1680
      ScaleHeight     =   795
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox ShockwaveFlash1 
      Height          =   5295
      Left            =   30
      ScaleHeight     =   5235
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   30
      Width           =   5955
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ShockwaveFlash1.Play
End Sub

Private Sub ShockwaveFlash1_OnReadyStateChange(newState As Long)

End Sub
