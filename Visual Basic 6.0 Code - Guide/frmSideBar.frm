VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSideBar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "frmSideBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grouped"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alphabetical"
            ImageVarType    =   2
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
   Begin VB.Menu mnuGuid 
      Caption         =   "File"
      Begin VB.Menu mnuShowMain 
         Caption         =   "Show Main Window"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuAlphablend 
         Caption         =   "Alpha Blend"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyntax 
         Caption         =   "Syntax Coloring"
      End
   End
End
Attribute VB_Name = "frmSideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lHeight As Long 'Initial form height
Dim lWidth As Long 'Initial form width
Dim AppBar As APPBARDATA 'Represents your docked form



Private Sub Form_Activate()
    Dock dsRight, Me, AppBar
    
    
End Sub
    
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
    
Private Sub Form_Load()
    
    lHeight = Me.Height
    lWidth = Me.Width
    
    
    
    
    If frmMain.rstInfo.RecordCount = 0 Then Exit Sub
    
    
    
    Dim i As Integer
    With frmMain.rstInfo
        For i = 0 To frmMain.rstInfo.RecordCount
            Call AddParent(!CodeID, !CodeID)
            Call AddNode(!CodeID, !CodeName)
            Call List1.AddItem(!CodeName)
            frmMain.rstInfo.MoveNext
        Next i
    End With
    frmMain.rstInfo.MoveFirst
    
    
    
    
    
    
End Sub
    
Public Sub ResetSize()
    ''''''''''''''''''''''''''''''''''
    ' This is just to bring the form '
    ' back to the size it was before '
    ' and center it on the screen.   '
    ''''''''''''''''''''''''''''''''''
    
    Me.Height = lHeight
    Me.Width = lWidth
    Me.Left = (Screen.Width - lWidth) / 2
    Me.Top = (Screen.Height - lHeight) / 2
End Sub
    
Private Sub Form_Resize()
    TabStrip1.Height = Me.Height - 650
    TreeView1.Height = TabStrip1.Height - 870
    List1.Height = TabStrip1.Height - 900
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    UnDock AppBar
End Sub
    
Private Sub mnuAlphablend_Click()
    
    Dim answer As Integer
    
    answer = MsgBox("This will only work for NT/2000/XP do you wish to Continue?", vbQuestion + vbYesNo, "Continue")
    
    If answer = vbYes Then
        
    End If
    
    If answer = vbNo Then
        Exit Sub
    End If
    
    
    If mnuAlphablend.Checked = False Then
        mnuAlphablend.Checked = True
        MakeTransparent Me.hwnd, 160
        Exit Sub
    End If
    
    If mnuAlphablend.Checked = True Then
        mnuAlphablend.Checked = False
        MakeOpaque Me.hwnd
        Exit Sub
    End If
    
End Sub
    
Private Sub mnuExit_Click()
    UnDock AppBar
    frmMain.Show
    Unload Me
End Sub
    
Private Sub mnuShowMain_Click()
    frmMain.Show
    Unload Me
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
    
    
    
Private Sub Text1_Change()
    
    intPlacement = Len(Text1) '(1)
    strFind = UCase(Text1.Text) '(2)
    Text1.SetFocus '$$ Just some House Keeping
    For intResult = 0 To 21 '(3)
        intSearchResult = InStr(UCase(Left(List1.List(intResult), intPlacement)), strFind) '(4, 5, 6)
        If intSearchResult > 0 Then '(7)
        MsgBox List1.List(intResult)
        Else 'Do Nothing
    End If
Next intResult

End Sub
    
    
Public Function AddNode(ByVal ParentName As String, ByVal ChildText As String)
    Me.TreeView1.Nodes.Add ParentName, tvwChild, , ChildText
End Function
    
    
Public Function NodeIndex(ByVal ChildName As String) As Integer
    
    
    For X = 1 To tV.Nodes.Count - 1
        If Me.TreeView1.Nodes.Item(X).Text = ChildName Then NodeIndex = X
    Next X
End Function
    
    
Public Function AddParent(ByVal ParentName, ParentText As String)
    On Error Resume Next
    Me.TreeView1.Nodes.Add , , ParentName, ParentText
End Function
    
    
Public Function DeleteNode(ByVal NodeText As String)
    Me.TreeView1.Nodes.Remove NodeIndex(NodeText)
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
    End If
    
End Sub
    
    
    
    
    
    
    
    
    
    
