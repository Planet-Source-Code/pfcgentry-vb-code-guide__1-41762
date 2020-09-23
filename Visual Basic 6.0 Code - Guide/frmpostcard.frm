VERSION 5.00
Begin VB.Form frmpostcard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome..."
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "This app is PostCard Ware.. If you would like please send a Post Crad to our Location tell us what you think about our Software..."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmpostcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    If frmMain.mnuontop.Checked = False Then
        Call FormOnTop(Me.hwnd, True)
        Exit Sub
    End If
    
    If frmMain.mnuontop.Checked = True Then
        Call FormOnTop(Me.hwnd, False)
        Exit Sub
    End If
    
End Sub
