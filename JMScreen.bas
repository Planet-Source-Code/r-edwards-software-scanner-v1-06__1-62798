Attribute VB_Name = "JMScreenSubs"
'------------------------------------------------------------
' Module    : JMScreen.bas
' Purpose   : Windows has a taskbar, which although being
'             admirably useful, can be a right pain in the
'             proverbial, when it comes to deciding how much
'             space you have on the screen for forms.
'             To make matters worse, it can be placed on any
'             edge of the screen.
'
' I have written four functions:
'------------------------------------------------------------
'   JMScreenLeft   - returns the Left value of the Screen
'   JMScreenTop    - returns the Top value of the Screen
'   JMScreenWidth  - returns the Widtht value of the Screen
'   JMScreenHeight - returns the Height value of the Screen
'
' which replace the normal calls of Screen.Height and Screen.Width
'------------------------------------------------------------
'
Option Explicit

'  Scructure Definitions
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

'  API Definitions
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Private Const ABS_ALWAYSONTOP = &H2
Private Const ABS_AUTOHIDE = &H1
Private Const ABM_GETSTATE = &H4
Private Const ABM_GETTASKBARPOS = &H5

'------------------------------------------------------------
' Purpose   : This routine is called by all of the other routines
'             and tests if the Windows Taskbar is shown
' Arguments : none
' Returns   : True if taskbar is visible
'             False if taskbar is not visible
'------------------------------------------------------------
'
Private Function JMTaskbarExists() As Boolean
    
    Dim wrkBar As APPBARDATA
    
    On Error Resume Next

    '--- Set Size of Structure
    wrkBar.cbSize = 36

    '--- Get Status of Taskbar
    Select Case SHAppBarMessage(ABM_GETSTATE, wrkBar)
        '--- Taskbar exists
        Case ABS_ALWAYSONTOP, ABS_AUTOHIDE
            JMTaskbarExists = True
            Exit Function
    End Select

    '--- Taskbar does not
    JMTaskbarExists = False

End Function

'------------------------------------------------------------
' Purpose   : This routine returns the height value of the screen
' Arguments : none
'------------------------------------------------------------
'
Public Function JMScreenHeight() As Long
    
    Dim wrkBar As APPBARDATA
    Dim wrkHeight As Long
    Dim wrkTop As Long
    Dim wrkBottom As Long
    
    On Error GoTo JMScreenHeightError

    '--- Set Default Height
    JMScreenHeight = Screen.Height

    '--- Test for a Taskbar
    If (JMTaskbarExists() = False) Then Exit Function

    '--- Set Size of Structure
    wrkBar.cbSize = 36

    '--- Get Size and Position of Taskbar
    wrkHeight = Screen.Height / Screen.TwipsPerPixelY
    If (SHAppBarMessage(ABM_GETTASKBARPOS, wrkBar) = False) Then Exit Function

    '--- Extract Top and Bottom
    wrkTop = wrkBar.rc.Top
    wrkBottom = wrkBar.rc.Bottom

    '--- Set if Bar is Vertical
    If (wrkTop <= 0 And wrkBottom >= wrkHeight) Then
        wrkHeight = Screen.Height

    '--- Set if Bar is at Top
    ElseIf (wrkTop < 0) Then
        wrkHeight = (wrkHeight - wrkBottom) * Screen.TwipsPerPixelY

    '--- Set if Bar is at Bottom
    ElseIf (wrkBottom >= wrkHeight) Then
        wrkHeight = wrkTop * Screen.TwipsPerPixelY

    '--- Set if Anywhere Else (Shouldn't be!)
    Else
        wrkHeight = Screen.Height
    End If

    '--- Set Height
    JMScreenHeight = wrkHeight
    Exit Function

'==========
JMScreenHeightError:    'error
'==========
    JMScreenHeight = Screen.Height
    Exit Function

End Function

'------------------------------------------------------------
' Purpose   : This routine returns the width value of the screen
' Arguments : none
'------------------------------------------------------------
'
Public Function JMScreenWidth() As Long
    
    Dim wrkBar As APPBARDATA
    Dim wrkWidth As Long
    Dim wrkLeft As Long
    Dim wrkRight As Long
    
    On Error GoTo JMScreenWidthError

    '--- Set Default Width
    JMScreenWidth = Screen.Width

    '--- Test for a Taskbar
    If (JMTaskbarExists() = False) Then Exit Function

    '--- Set Size of Structure
    wrkBar.cbSize = 36

    '--- Get Size and Position of Taskbar
    wrkWidth = Screen.Width / Screen.TwipsPerPixelX
    If (SHAppBarMessage(ABM_GETTASKBARPOS, wrkBar) = False) Then Exit Function

    '--- Extract Left and Right
    wrkLeft = wrkBar.rc.Left
    wrkRight = wrkBar.rc.Right

    '--- Set if Bar is Horizontal
    If (wrkLeft <= 0 And wrkRight >= wrkWidth) Then
        wrkWidth = Screen.Width

    '--- Set if Bar is at Left
    ElseIf (wrkLeft < 0) Then
        wrkWidth = (wrkWidth - wrkRight) * Screen.TwipsPerPixelX

    '--- Set if Bar is at Right
    ElseIf (wrkRight >= wrkWidth) Then
        wrkWidth = wrkLeft * Screen.TwipsPerPixelY

    '--- Set if Anywhere Else (Shouldn't be!)
    Else
        wrkWidth = Screen.Width
    End If

    '--- Set Width
    JMScreenWidth = wrkWidth
    Exit Function

'==========
JMScreenWidthError: 'error
'==========
    JMScreenWidth = Screen.Width
    Exit Function

End Function

'------------------------------------------------------------
' Purpose   : This routine returns the top value of the screen
' Arguments : none
'------------------------------------------------------------
'
Public Function JMScreenTop() As Long
    
    Dim wrkBar As APPBARDATA
    Dim wrkScreenTop As Long
    Dim wrkHeight As Long
    Dim wrkTop As Long
    Dim wrkBottom As Long
    
    On Error GoTo JMScreenTopError
    
    '--- Set Default Top
    JMScreenTop = 0

    '--- Test for a Taskbar
    If (JMTaskbarExists() = False) Then Exit Function

    '--- Set Size of Structure
    wrkBar.cbSize = 36

    '--- Get Size and Position of Taskbar
    If (SHAppBarMessage(ABM_GETTASKBARPOS, wrkBar) = False) Then Exit Function

    '--- Extract Top and Bottom
    wrkTop = wrkBar.rc.Top
    wrkBottom = wrkBar.rc.Bottom

    '--- Set Screen Height
    wrkHeight = Screen.Height / Screen.TwipsPerPixelY

    '--- Set if Bar is at Top
    If (wrkTop < 0 And wrkBottom < wrkHeight) Then
        wrkScreenTop = wrkBottom * Screen.TwipsPerPixelY

    '--- Set if Anywhere Else
    Else
        wrkScreenTop = 0
    End If

    '--- Set Top
    JMScreenTop = wrkScreenTop
    Exit Function

'==========
JMScreenTopError:   'error
'==========
    JMScreenTop = 0
    Exit Function

End Function

'------------------------------------------------------------
' Purpose   : This routine returns the left value of the screen
' Arguments : none
'------------------------------------------------------------
'
Public Function JMScreenLeft() As Long
    
    Dim wrkBar As APPBARDATA
    Dim wrkScreenLeft As Long
    Dim wrkWidth As Long
    Dim wrkLeft As Long
    Dim wrkRight As Long
    
    On Error GoTo JMScreenLeftError

    '--- Set Default Top
    JMScreenLeft = 0

    '--- Test for a Taskbar
    If (JMTaskbarExists() = False) Then Exit Function

    '--- Set Size of Structure
    wrkBar.cbSize = 36

    '--- Get Size and Position of Taskbar
    If (SHAppBarMessage(ABM_GETTASKBARPOS, wrkBar) = False) Then Exit Function

    '--- Extract Left and Right
    wrkLeft = wrkBar.rc.Left
    wrkRight = wrkBar.rc.Right

    '--- Set Screen Height
    wrkWidth = Screen.Width / Screen.TwipsPerPixelX

    '--- Set if Bar is at Left
    If (wrkLeft < 0 And wrkRight < wrkWidth) Then
        wrkScreenLeft = wrkRight * Screen.TwipsPerPixelX

    '--- Set if Anywhere Else
    Else
        wrkScreenLeft = 0
    End If

    '--- Set Left
    JMScreenLeft = wrkScreenLeft
    Exit Function

'==========
JMScreenLeftError:  'error
'==========
    JMScreenLeft = 0
    Exit Function

End Function

'------------------------------------------------------------
' Purpose   : Positions a form on the screen. Also makes sure entire form
'             is within the screen boundries
' Arguments : fName   - form name
'             argTop  - top position of form
'             argLeft - left position of form
'------------------------------------------------------------
'
Public Sub SetFormPosition(fName As Form, argTop As Long, argLeft As Long)
    
    On Error Resume Next
    
    '--- Position Form
    fName.Left = argLeft
    fName.Top = argTop

    '--- Check not too far right
    If ((fName.Left + fName.Width) > (JMScreenLeft() + JMScreenWidth())) Then
        fName.Left = JMScreenLeft() + JMScreenWidth() - fName.Width
    End If

    '--- Check not too far down
    If ((fName.Top + fName.Height) > (JMScreenTop() + JMScreenHeight())) Then
        fName.Top = JMScreenTop() + JMScreenHeight() - fName.Height
    End If
    
    '--- Check not too far left
    If (fName.Left < JMScreenLeft()) Then fName.Left = JMScreenLeft()

    '--- Check not too far up
    If (fName.Top < JMScreenTop()) Then fName.Top = JMScreenTop()

End Sub

