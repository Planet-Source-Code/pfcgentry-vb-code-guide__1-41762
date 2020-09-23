VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Window"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmCodeWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   6855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   600
         Width           =   5175
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6800
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmCodeWindow.frx":0CCA
      End
      Begin VB.Label Label1 
         Caption         =   "Function  Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Min. OS Required:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   6615
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6588
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmCodeWindow.frx":0D4C
      End
      Begin VB.Label Label4 
         Caption         =   "Declarations:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Parameters:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox3 
         Height          =   4935
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmCodeWindow.frx":0DCE
      End
      Begin VB.Label Label6 
         Caption         =   "Your Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5535
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox4 
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmCodeWindow.frx":0E50
      End
      Begin VB.Label Label8 
         Caption         =   "Example:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5535
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox5 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmCodeWindow.frx":0ED2
      End
      Begin VB.Label Label7 
         Caption         =   "Example:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6615
      End
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Information"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "(---) Parameters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Example(s)"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   ".NET"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCodeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TabStrip2_Click()
    
    If TabStrip2.Tabs(1).Selected = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
    End If
    
    
    
    If TabStrip2.Tabs(2).Selected = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
    End If
    
    
    
    If TabStrip2.Tabs(3).Selected = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
        Frame4.Visible = False
        Frame5.Visible = False
    End If
    
    
    
    If TabStrip2.Tabs(4).Selected = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = True
        Frame5.Visible = False
    End If
    
    
    
    If TabStrip2.Tabs(5).Selected = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = True
    End If
End Sub
    
