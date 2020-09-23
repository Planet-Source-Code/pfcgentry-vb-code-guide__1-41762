VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Visual Basic 6.0 Code-Guide"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0CCA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3534
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":420E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   5535
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox5 
         Height          =   4935
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":4EE8
      End
      Begin VB.Label Label7 
         Caption         =   "Example:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5535
      Left            =   3360
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox4 
         Height          =   4935
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMain.frx":4F7C
      End
      Begin VB.Label Label8 
         Caption         =   "Example:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   3360
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox3 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":4FFE
      End
      Begin VB.Label Label6 
         Caption         =   "Your Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   3735
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6588
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":5080
      End
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label5 
         Caption         =   "Parameters:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Declarations:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   6855
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6800
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":5102
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Min. OS Required:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Code  Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   6015
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10610
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Information"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "(---) Parameters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Example(s)"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   ".NET"
            ImageVarType    =   2
            ImageIndex      =   4
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10610
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grouped"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alphabetical"
            ImageVarType    =   2
            ImageIndex      =   5
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6315
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuontop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCodeBar 
         Caption         =   "Code-Guide Bar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlphaBlend 
         Caption         =   "Alpha Blend"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Syntax Coloring"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administration"
      Visible         =   0   'False
      Begin VB.Menu mnuAdminToolBar 
         Caption         =   "Admin Tool Bar"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "Read Me"
      End
      Begin VB.Menu mnuPostcard 
         Caption         =   "PostCardWare"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send an Example"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuABout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorize 
         Caption         =   "Syntax Code"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim as Public, which allows these variables to be used by all forms and modules within
'this VB project.
Public con As ADODB.Connection
Public rstInfo As ADODB.Recordset
Public sql As String
'Const CON_DB1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.bcg;Persist Security Info=False;Jet OLEDB:Database Password=gentry"



Private Sub Form_Load()


    Dim sTmp As String, CON_DB1 As String
    
    
    
    Frame1.BorderStyle = 0
    Frame2.BorderStyle = 0
    Frame3.BorderStyle = 0
    Frame4.BorderStyle = 0
    Frame5.BorderStyle = 0
    
    CON_DB1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.bcg;Persist Security Info=False;Jet OLEDB:Database Password=gentry"
    
    On Error Resume Next
    
    Set con = New ADODB.Connection
    Set rstInfo = New ADODB.Recordset
    rstInfo.CursorType = adOpenKeyset
    rstInfo.LockType = adLockOptimistic
    con.Open CON_DB1
    
    
    rstInfo.Open "Select * from [Code]", con
    
    frmMain.Caption = frmMain.Caption & " - " & rstInfo.RecordCount & " Record(s) available"
    If rstInfo.RecordCount = 0 Then Exit Sub
    
    Set Text2.DataSource = rstInfo
    Set Text3.DataSource = rstInfo
    Set Text4.DataSource = rstInfo
    Set RichTextBox1.DataSource = rstInfo
    Set RichTextBox2.DataSource = rstInfo
    Set RichTextBox3.DataSource = rstInfo
    Set RichTextBox4.DataSource = rstInfo
    Set RichTextBox5.DataSource = rstInfo
    
    'Text1.DataField = "CodeID"
    Text2.DataField = "CodeName"
    RichTextBox1.DataField = "CodeInfo"
    Text3.DataField = "CodeOS"
    Text4.DataField = "Declarations"
    RichTextBox2.DataField = "Parameters"
    RichTextBox3.DataField = "Notes"
    RichTextBox4.DataField = "Source"
    RichTextBox5.DataField = "Net"
    
    Dim i As Integer
    
    With rstInfo
        For i = 0 To rstInfo.RecordCount
            Call AddParent(!CodeID, !CodeID)
            Call AddNode(!CodeID, !CodeName)
            Call List1.AddItem(!CodeName)
            rstInfo.MoveNext
        Next i
    End With
    rstInfo.MoveFirst
    
    
    With frmMain
        .Text1.ToolTipText = "Input search name"
        .List1.ToolTipText = "Contains possible search Data"
        .Text1.MaxLength = 20
    End With
    
    
    
    
    
End Sub
    
Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.Height <= 7380 Then
        Me.Height = 7380
    End If
    
    
    If Me.Width <= 10545 Then
        Me.Width = 10545
    End If
    
    TabStrip1.Left = 120
    TabStrip2.Left = 3240
    TabStrip1.Top = 480
    TabStrip2.Top = 480
    
    TabStrip1.Height = frmMain.Height - 1400
    TabStrip2.Height = frmMain.Height - 1400
    TabStrip2.Width = frmMain.Width - 3500
    
    TreeView1.Top = 960
    List1.Top = 1320
    TreeView1.Height = frmMain.Height - 2080
    List1.Height = frmMain.Height - 2350
    
    Text2.Left = 1560
    Text3.Left = 1560
    Text4.Left = 120
    
    RichTextBox1.Left = 120
    RichTextBox2.Left = 120
    RichTextBox3.Left = 120
    RichTextBox4.Left = 120
    RichTextBox5.Left = 120
    
    
    
    RichTextBox1.Top = 1560
    RichTextBox2.Top = 1680
    RichTextBox3.Top = 480
    RichTextBox4.Top = 480
    RichTextBox5.Top = 480
    
    
    Frame1.Left = 3360
    
    
    Frame1.Width = TabStrip2.Width - 230
    Frame2.Width = TabStrip2.Width - 230
    Frame3.Width = TabStrip2.Width - 230
    Frame4.Width = TabStrip2.Width - 230
    Frame5.Width = TabStrip2.Width - 230
    
    Frame1.Height = TabStrip2.Height - 530
    Frame2.Height = TabStrip2.Height - 530
    Frame3.Height = TabStrip2.Height - 530
    Frame4.Height = TabStrip2.Height - 530
    Frame5.Height = TabStrip2.Height - 530
    
    
    RichTextBox1.Width = TabStrip2.Width - 470
    RichTextBox2.Width = TabStrip2.Width - 470
    RichTextBox3.Width = TabStrip2.Width - 470
    RichTextBox4.Width = TabStrip2.Width - 470
    RichTextBox5.Width = TabStrip2.Width - 470
    
    RichTextBox1.Height = TabStrip2.Height - 2250
    RichTextBox2.Height = TabStrip2.Height - 2300
    RichTextBox3.Height = TabStrip2.Height - 1150
    RichTextBox4.Height = TabStrip2.Height - 1150
    RichTextBox5.Height = TabStrip2.Height - 1150
    
    Text2.Width = frmMain.Width - 5500
    Text3.Width = frmMain.Width - 5500
    Text4.Width = TabStrip2.Width - 470
    
End Sub
    
Private Sub List1_Click()
    
    Dim sNode As String
    'On Error Resume Next
    
    sNode = List1.Text
    
    If sNode <> "" Then
        rstInfo.MoveFirst
        rstInfo.Find "[CodeName] = '" & sNode & "'"
        
    If mnuColor.Checked = False Then
        mnuColor.Checked = True
        'Call Color_All
        Exit Sub
    End If
    
    If mnuColor.Checked = True Then
        mnuColor.Checked = False
       ' Call Color_All
        Exit Sub
    End If

        
    End If
    
End Sub
    
Private Sub mnuABout_Click()
frmAbout.Show
End Sub

Private Sub mnuAdminToolBar_Click()
    frmAdmin.Show
End Sub
    
Private Sub mnuAlphablend_Click()
    
    Dim answer As Integer
    
    answer = MsgBox("This will only work for NT/2000/XP do you wish to Continue?", vbQuestion + vbYesNo, "Continue")
    
    If answer = vbYes Then
        
    End If
    
    If answer = vbNo Then
        Exit Sub
    End If
    
    
    If mnuAlphaBlend.Checked = False Then
        mnuAlphaBlend.Checked = True
        MakeTransparent Me.hwnd, 160
        Exit Sub
    End If
    
    If mnuAlphaBlend.Checked = True Then
        mnuAlphaBlend.Checked = False
        MakeOpaque Me.hwnd
        Exit Sub
    End If
    
    
    
End Sub

    
Private Sub mnuColor_Click()
    
    If mnuColor.Checked = False Then

    End If
    
    If mnuColor.Checked = True Then

    End If
  
    
    
End Sub

    
Private Sub mnuExit_Click()
    
    Dim answer As Integer
    answer = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit?")
    
    If answer = vbYes Then
    
    Set rstInfo = Nothing
    Set con = Nothing
    
    Unload Me
    
    End
    
    End If
    
    If answer = vbNo Then
    
    Exit Sub
    
    End If
    
    

End Sub
    
Private Sub mnuontop_Click()
    
    If mnuontop.Checked = False Then
        mnuontop.Checked = True
        Call FormOnTop(Me.hwnd, True)
        Exit Sub
    End If
    
    If mnuontop.Checked = True Then
        mnuontop.Checked = False
        Call FormOnTop(Me.hwnd, False)
        Exit Sub
    End If
    
End Sub
    
    
Private Sub mnuPostcard_Click()
frmpostcard.Show
End Sub

Private Sub mnuProperties_Click()
    frmProperties.Show
End Sub
    
Private Sub mnuReadme_Click()
MsgBox "No read me as of Yet"
End Sub

Private Sub mnuSend_Click()
MsgBox "No send example as of Yet"
End Sub

Private Sub RichTextBox4_Change()
    
    If mnuColor.Checked = True Then
        'Call Color_All
        Exit Sub
    End If
    
End Sub
    
Private Sub RichTextBox4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ''If not right-clicking then exit
    If Button <> 2 Then Exit Sub
    
    ''Pop up the menu
    Call Me.PopupMenu(mnuPopup, vbAlignNone)
    
End Sub
    
Private Sub TabStrip1_Click()
    
    If TabStrip1.Tabs(1).Selected = True Then
        
        TreeView1.Visible = True
        List1.Visible = False
        Text1.Visible = False
        
    End If
    
    If TabStrip1.Tabs(2).Selected = True Then
        
        TreeView1.Visible = False
        List1.Visible = True
        Text1.Visible = True
        
    End If
    
End Sub
    
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
    
Private Sub Text1_Change()
    
    intPlacement = Len(Text1) '(1)
    strFind = UCase(Text1.Text) '(2)
    Text1.SetFocus '$$ Just some House Keeping
    For intResult = 0 To 21 '(3)
        intSearchResult = InStr(UCase(Left(List1.List(intResult), intPlacement)), strFind) '(4, 5, 6)
        If intSearchResult > 0 Then '(7)
        Text1.Text = List1.List(intResult)
    
    Dim sNode As String
    'On Error Resume Next
    
    sNode = List1.List(intResult)
    
    If sNode <> "" Then
        rstInfo.MoveFirst
        rstInfo.Find "[CodeName] = '" & sNode & "'"
    End If
        
        
        
        
        Else 'Do Nothing
    End If
Next intResult

'Text1.SelStart = Len(Text1.Text) '=== Error Cant back space after done... needs to be fixed

End Sub
    
    
Public Function AddNode(ByVal ParentName As String, ByVal ChildText As String)
    TreeView1.Nodes.Add ParentName, tvwChild, , ChildText
End Function
    
    
Public Function NodeIndex(ByVal ChildName As String) As Integer
    Dim x As Integer
    
    For x = 1 To TreeView1.Nodes.Count - 1
        If TreeView1.Nodes.Item(x).Text = ChildName Then NodeIndex = x
    Next x
End Function
    
    
Public Function AddParent(ByVal ParentName, ParentText As String)
    TreeView1.Nodes.Add , , ParentName, ParentText
End Function
    
    
Public Function DeleteNode(ByVal NodeText As String)
    TreeView1.Nodes.Remove NodeIndex(NodeText)
End Function
    
    
Private Sub TreeView1_Click()
    
    Dim sParent As String, sNode As String
    On Error Resume Next
    
    sParent = TreeView1.SelectedItem.Parent.Text
    sNode = TreeView1.SelectedItem.Text
    
    If sParent <> "" Then
        rstInfo.MoveFirst
        rstInfo.Find "[CodeID] = '" & sParent & "'"
        rstInfo.Find "[CodeName] = '" & sNode & "'"
        
        
    If mnuColor.Checked = False Then
        mnuColor.Checked = True
        'Call Color_All
        Exit Sub
    End If
    
    If mnuColor.Checked = True Then
        mnuColor.Checked = False
        'Call Color_All
        Exit Sub
    End If
    
        
        
    End If
    
End Sub
    
