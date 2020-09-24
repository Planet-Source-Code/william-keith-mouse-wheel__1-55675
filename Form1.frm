VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Roll the mouse wheel back and forth while hovering over this form and watch the immediate window to see the effect of this code."
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Hook Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook
End Sub
