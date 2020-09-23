Attribute VB_Name = "modDockBar"
Public Declare Function SHAppBarMessage Lib "Shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Public Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Long

Public Const ABM_NEW = &H0
Public Const ABM_REMOVE = &H1
Public Const ABM_SETPOS = &H3
Public Const ABE_BOTTOM = 3
Public Const ABE_TOP = 1
Public Const ABE_LEFT = 0
Public Const ABE_RIGHT = 2
Public Const WM_MOUSEMOVE = &H200
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40

Public Const SPI_GETWORKAREA = 48

Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Type APPBARDATA
    cbSize           As Long
    hwnd             As Long
    uCallbackMessage As Long
    uEdge            As Long
    rc               As RECT
    lParam           As Long
End Type

Public Enum DockSide
    dsTop = 1
    dsBottom = 3
    dsLeft = 0
    dsRight = 2
End Enum

Public Function Dock(ByVal ScreenSide As DockSide, ByRef xForm As Form, ByRef AppBar As APPBARDATA) As Boolean
    Dim lScreenWidth As Long   ' Holds the width of the screen in pixles
    Dim lScreenHeight As Long  ' Holds the height of the screen in pixles
    Dim lxHeight As Long       ' Holds the height of the form in pixles
    Dim lxWidth As Long        ' Holds the width of the form in pixles
    Dim lTaskBarHeight As Long ' Holds the height of the taskbar
    Dim bResult As Boolean     ' Holds the API calls results
    
    Dim WorkArea As RECT  ' These hold the area of the screen
    Dim xWorkArea As RECT ' not currently reserved for a docked
    Dim yWorkArea As RECT ' program.
    
    AppBar.hwnd = xForm.hwnd               ' Handle to the form to be docked
    AppBar.cbSize = Len(AppBar)            ' Size of the AppBar Variable
    AppBar.uCallbackMessage = WM_MOUSEMOVE ' Call back function for any system messages
    
    GetWorkArea WorkArea ' Get the current area of the screen not reserved
    
    lScreenWidth = Screen.Width / Screen.TwipsPerPixelX   ' Get the screen
    lScreenHeight = Screen.Height / Screen.TwipsPerPixelY ' dimensions.
    
    lxHeight = xForm.Height / Screen.TwipsPerPixelY ' Get the form
    lxWidth = xForm.Width / Screen.TwipsPerPixelX   ' dimensions
    
    bResult = SHAppBarMessage(ABM_REMOVE, AppBar) ' Undock the form if it is docked
    DoEvents ' Let the system catch up
    bResult = SHAppBarMessage(ABM_NEW, AppBar) ' Register the program with windows for docking
    
    Select Case ScreenSide ' Find where you want to dock it and set the placement dimensions
        Case ABE_TOP ' Top of screen
            AppBar.uEdge = ABE_TOP
            AppBar.rc.Top = WorkArea.Top
            AppBar.rc.Left = 0
            AppBar.rc.Right = lScreenWidth
            AppBar.rc.Bottom = lxHeight
        Case ABE_BOTTOM ' Bottom of screen
            AppBar.uEdge = ABE_BOTTOM
            AppBar.rc.Top = WorkArea.Bottom - lxHeight
            AppBar.rc.Left = 0
            AppBar.rc.Right = lScreenWidth
            AppBar.rc.Bottom = WorkArea.Bottom
        Case ABE_LEFT ' Left side of screen
            AppBar.uEdge = ABE_LEFT
            AppBar.rc.Top = 0
            AppBar.rc.Left = WorkArea.Left
            AppBar.rc.Right = WorkArea.Left + lxWidth
            AppBar.rc.Bottom = lScreenHeight
        Case ABE_RIGHT ' Right side of screen
            AppBar.uEdge = ABE_RIGHT
            AppBar.rc.Top = 0
            AppBar.rc.Left = WorkArea.Right - lxWidth
            AppBar.rc.Right = WorkArea.Right
            AppBar.rc.Bottom = lScreenHeight
    End Select
    
    GetWorkArea xWorkArea ' Find the area of the screen not reserved
    
    bResult = SHAppBarMessage(ABM_SETPOS, AppBar) ' Reserve screen space for the form
    DoEvents ' This can take a second so give the
    DoEvents ' system time to register the space
    
    With AppBar.rc
        ' This If...Then chunk keeps checking until the unreserved
        ' screen area before we reserved space is different then
        ' the unreserved area after we reserved space depending on
        ' where we have it docked.  Just waiting for the system to
        ' catch up basicly.
        If ScreenSide = dsTop Then
            GetWorkArea yWorkArea
            Do Until yWorkArea.Top > xWorkArea.Top
                GetWorkArea yWorkArea
                DoEvents
            Loop
        ElseIf ScreenSide = dsBottom Then
            GetWorkArea yWorkArea
            Do Until xWorkArea.Bottom > yWorkArea.Bottom
                GetWorkArea yWorkArea
                DoEvents
            Loop
        ElseIf ScreenSide = dsLeft Then
            GetWorkArea yWorkArea
            Do Until yWorkArea.Left > xWorkArea.Left
                GetWorkArea yWorkArea
                DoEvents
            Loop
        ElseIf ScreenSide = dsRight Then
            GetWorkArea yWorkArea
            Do Until xWorkArea.Right > yWorkArea.Right
                GetWorkArea yWorkArea
                DoEvents
            Loop
        End If
        
        ' This next line will put the form on top of all
        ' other windows
        bResult = SetWindowPos(xForm.hwnd, HWND_TOPMOST, .Top, .Left, .Right, .Bottom, SWP_SHOWWINDOW)
        
        ' This last chunk is just to resize the form
        ' to fit the space we reserved for it.
        xForm.Top = .Top * Screen.TwipsPerPixelY
        xForm.Left = .Left * Screen.TwipsPerPixelX
        xForm.Height = (.Bottom - .Top) * Screen.TwipsPerPixelY
        xForm.Width = (.Right - .Left) * Screen.TwipsPerPixelX
    End With
End Function

Public Sub UnDock(ByRef AppBar As APPBARDATA)
    Call SHAppBarMessage(ABM_REMOVE, AppBar) ' unreserve the space we reserved and unregister the form
    DoEvents ' Wait for the system to catch up
End Sub

Public Function TaskBarHeight()
    ' Returns the height of the task bar
    Dim lHeight1 As Long
    Dim lHeight2 As Long
    Dim bResult As Boolean
    Dim rcArea As RECT
    Dim luiParam As Long
    Dim lScreenHeight As Long
    
    ' Find the area of the screen not reserved
    bResult = SystemParametersInfo(SPI_GETWORKAREA, 0, rcArea, 0)
    
    ' Find the total height of the screen
    lScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    
    TaskBarHeight = lScreenHeight - rcArea.Bottom
End Function

Public Sub GetWorkArea(ByRef WndRgn As RECT)
    ' Returns the area of the screen not reserved
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, WndRgn, 0)
End Sub
