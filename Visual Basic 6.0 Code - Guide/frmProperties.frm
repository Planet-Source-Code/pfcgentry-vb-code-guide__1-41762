VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Properties Still In the Works"
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Display Splash at Startup."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
    
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
    
