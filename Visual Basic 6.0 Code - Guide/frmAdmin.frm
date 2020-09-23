VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text7 
      Height          =   1335
      Left            =   4920
      TabIndex        =   20
      Top             =   3240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAdmin.frx":0CCA
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      ItemData        =   "frmAdmin.frx":0D4C
      Left            =   120
      List            =   "frmAdmin.frx":0D6B
      TabIndex        =   14
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Text9 
      Height          =   1365
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3840
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "rstInfo"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   1275
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Height          =   1035
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      Height          =   1035
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Text            =   "Not Code as of Yet"
      Top             =   4920
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Declarations"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   ".NET Source"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Source"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Notes"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Parameters"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Code Operating System"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Code Information"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Code Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Code ID"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdmin"
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
    
Private Sub Command1_Click()
    With frmMain.rstInfo
        .AddNew
        
        If Trim(Text1.Text) <> "" Then
            !CodeID = Text1.Text
        Else
            MsgBox "You Must Enter A Name!", vbInformation, "Error Adding Record"
            Exit Sub
        End If
        
        If Trim(Text2.Text) <> "" Then
            !CodeName = Text2.Text
        Else
            !CodeName = "No Entry"
        End If
        
        If Trim(Text3.Text) <> "" Then
            !CodeInfo = Text3.Text
        Else
            !CodeInfo = "No Entry"
        End If
        
        If Trim(Text4.Text) <> "" Then
            !CodeOS = Text4.Text
        Else
            !CodeOS = "No Entry"
        End If
        
        If Trim(Text5.Text) <> "" Then
            !Parameters = Text5.Text
        Else
            !Parameters = "No Entry"
        End If
        
        If Trim(Text6.Text) <> "" Then
            !Notes = Text6.Text
        Else
            !Notes = "No Entry"
        End If
        
        If Trim(Text7.Text) <> "" Then
            !Source = Text7.TextRTF
        Else
            !Source = "No Entry"
        End If
        
        If Trim(Text8.Text) <> "" Then
            !Net = Text8.Text
        Else
            !Net = "No Entry"
        End If
        
        If Trim(Text9.Text) <> "" Then
            !Declarations = Text9.Text
        Else
            !Declarations = "No Entry"
        End If
        
        
        .Update
    End With
    
    frmMain.Refresh
    
End Sub
    
Private Sub Command4_Click()
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
