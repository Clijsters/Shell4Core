Imports System
Imports System.Drawing
Imports System.Drawing.Rectangle
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Linq.Queryable
Namespace Shell4Core

    ''' <summary>Main AppBar for Shell4Core</summary>
    Public Class AppBar
        Inherits System.Windows.Forms.Form

        Public ABD As AppBarData
        Private uCallBack As Integer
        Private BarIsDocked As Boolean = False
        Private ScreenToDock As Integer

        Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
            Get
                Dim cp As CreateParams = MyBase.CreateParams
                cp.Style = cp.Style And Not &HC00000 ' WS_CAPTION   <=Covered by FormBorderStyle.None
                cp.Style = cp.Style And Not &H800000 ' WS_BORDER    <=Covered by FormBorderStyle.None
                cp.ExStyle = &H80 Or &H8 ' WS_EX_TOOLWINDOW | WS_EX_TOPMOST
                Return cp
            End Get
        End Property

        ''' <summary>Creates a new AppBar Instance.</summary>
        ''' <param name="Edge">AppBar.ABEdge: Defines the Edge to dock on.</param>
        ''' <param name="Screen">Integer: The Screen to appear on. (Default is 0)</param>
        Public Sub New(ByVal Edge As ABEdge, Optional ByVal Screen As Integer = 0)
            Me.ABD.uEdge = Edge
            Me.ABD.cbSize = Marshal.SizeOf(ABD)
            Me.ABD.hWnd = Me.Handle
            Me.ScreenToDock = Screen
        End Sub

        Public Sub RegisterBar()
            If Not BarIsDocked Then 'Dock it
                uCallBack = RegisterWindowMessage("AppBarMessage")
                ABD.uCallbackMessage = uCallBack
                If SHAppBarMessage(CInt(ABMsg.ABM_NEW), ABD) Then
					ABSetPos()
					BarIsDocked = True
                Else
					
                End If
			Else 'Undock it
                SHAppBarMessage(CInt(ABMsg.ABM_REMOVE), ABD)
                BarIsDocked = False
            End If
        End Sub

        'https://msdn.microsoft.com/library/en-us/shellcc/platform/Shell/programmersguide/shell_int/shell_int_programming/appbars.asp
        'http://stackoverflow.com/questions/14698755/appbar-multi-monitor
        Private Sub ABSetPos()
            Dim scrReference As Rectangle
            'TODO: Let the user select the preferred screen
            Dim scr As Screen = Screen.AllScreens(Me.ScreenToDock)
            scrReference = scr.Bounds

            If ABD.uEdge = CInt(ABEdge.ABE_LEFT) Or ABD.uEdge = CInt(ABEdge.ABE_RIGHT) Then
                Throw New ArgumentOutOfRangeException("Docking to the left and right Edges is currently not supported.")
            Else
                ABD.rc.left = scrReference.Left
                ABD.rc.right = scrReference.Right
                If ABD.uEdge = CInt(ABEdge.ABE_TOP) Then
                    ABD.rc.top = scrReference.Top
                    ABD.rc.bottom = Size.Height
                Else
                    ABD.rc.bottom = scrReference.Bottom
                    ABD.rc.top = ABD.rc.bottom - Size.Height
                End If
            End If

            ' Query the system for an approved size and position. 
            SHAppBarMessage(CInt(ABMsg.ABM_QUERYPOS), ABD)

            Select Case ABD.uEdge
                Case CInt(ABEdge.ABE_LEFT)
                    ABD.rc.right = ABD.rc.left + Size.Width
                Case CInt(ABEdge.ABE_RIGHT)
                    ABD.rc.left = ABD.rc.right - Size.Width
                Case CInt(ABEdge.ABE_TOP)
                    ABD.rc.bottom = ABD.rc.top + Size.Height
                Case CInt(ABEdge.ABE_BOTTOM)
                    ABD.rc.top = ABD.rc.bottom - Size.Height
            End Select

            'Pass the final bounding rectangle to the system
            SHAppBarMessage(CInt(ABMsg.ABM_SETPOS), ABD)
            'Now, let's move and resize to it
            MoveWindow(ABD.hWnd, ABD.rc.left, ABD.rc.top, ABD.rc.right - ABD.rc.left, ABD.rc.bottom - ABD.rc.top, True)
        End Sub

        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            If m.Msg = uCallBack Then
                'https://msdn.microsoft.com/en-us/library/windows/desktop/cc144177%28v=vs.85%29.aspx
                Select Case m.WParam.ToInt32()
                    Case CInt(ABNotify.ABN_POSCHANGED)
                        ABSetPos()
                End Select
            End If
            MyBase.WndProc(m)
        End Sub
#Region "Win32-API"

        <DllImport("User32.dll", EntryPoint:="IsWindowVisible")>
        Public Shared Function IsWindowVisible(ByVal hwnd As IntPtr) As Boolean
        End Function
        <DllImport("User32.dll", EntryPoint:="GetWindowText")>
        Public Shared Function GetWindowText(ByVal hwnd As Integer, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
        End Function
        <DllImport("User32.dll", EntryPoint:="EnumWindows")> _
        Public Shared Function EnumWindows(ByVal Adress As EnumWindowProc, ByVal y As Integer) As Integer
        End Function
        Public Delegate Function EnumWindowProc(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        'http://www.pinvoke.net/default.aspx/user32/GetWindowLongPtr.html
        Public Shared Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
            If IntPtr.Size = 8 Then
                Return GetWindowLongPtr64(hWnd, nIndex)
            Else
                Return GetWindowLongPtr32(hWnd, nIndex)
            End If
        End Function
        Private Declare Auto Function RegisterWindowMessage Lib "User32.dll" (ByVal msg As String) As Integer
        'https://msdn.microsoft.com/de-de/library/windows/desktop/bb762108%28v=vs.85%29.aspx
        Public Declare Function SHAppBarMessage Lib "Shell32.dll" Alias "SHAppBarMessage" (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As System.UInt32
        Public Declare Function MoveWindow Lib "User32.dll" Alias "MoveWindow" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal repaint As Boolean) As Boolean
        Public Declare Function ShowWindowAsync Lib "User32.dll" Alias "ShowWindowAsync" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        <DllImport("User32.dll", EntryPoint:="GetWindowLong")> _
        Private Shared Function GetWindowLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        End Function
        <DllImport("User32.dll", EntryPoint:="GetWindowLongPtr")> _
        Private Shared Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        End Function
#End Region

#Region "AppBarStructs"
        Structure RECT
            Public left As Integer
            Public top As Integer
            Public right As Integer
            Public bottom As Integer
        End Structure 'RECT

        Structure APPBARDATA
            Public cbSize As Integer
            Public hWnd As IntPtr
            Public uCallbackMessage As Integer
            Public uEdge As ABEdge
            Public rc As RECT
            Public lParam As IntPtr
        End Structure

        Enum ABEdge
            ABE_LEFT = 0
            ABE_TOP = 1
            ABE_RIGHT = 2
            ABE_BOTTOM = 3
        End Enum

        Enum ABMsg
            ABM_NEW = 0
            ABM_REMOVE = 1
            ABM_QUERYPOS = 2
            ABM_SETPOS = 3
            ABM_GETSTATE = 4
            ABM_GETTASKBARPOS = 5
            ABM_ACTIVATE = 6
            ABM_GETAUTOHIDEBAR = 7
            ABM_SETAUTOHIDEBAR = 8
            ABM_WINDOWPOSCHANGED = 9
            ABM_SETSTATE = 10
        End Enum

        Enum ABNotify
            ABN_STATECHANGE = 0
            ABN_POSCHANGED = 1
            ABN_FULLSCREENAPP = 2
            ABN_WINDOWARRANGE = 3
        End Enum

        'http://www.pinvoke.net/default.aspx/Enums/WindowLongFlags.html
        Enum GWL As Integer
            GWL_EXSTYLE = -20
            GWLP_HINSTANCE = -6
            GWLP_HWNDPARENT = -8
            GWL_ID = -12
            GWL_STYLE = -16
            GWL_USERDATA = -21
            GWL_WNDPROC = -4
            DWLP_USER = &H8
            DWLP_MSGRESULT = &H0
            DWLP_DLGPROC = &H4
        End Enum

        Enum WindowStates
            HIDE = 0
            SHOWNORMAL = 1
            SHOWMINIMIZED = 2
            MAXIMIZE = 3
            SHOWMAXIMIZED = 3
            SHOWNOACTIVATE = 4
            SHOW = 5
            MINIMIZE = 6
            SHOWMINNOACTIVE = 7
            SHOWNA = 8
            RESTORE = 9
            SHOWDEFAULT = 10
            FORCEMINIMIZE = 11
        End Enum
#End Region
    End Class

    Public Class AppBarIcon
        Inherits System.Windows.Forms.Button

        'Window Styles
        Const WS_OVERLAPPED As UInt32 = 0
        Const WS_POPUP As UInt32 = &H80000000&
        Const WS_CHILD As UInt32 = &H40000000
        Const WS_MINIMIZE As UInt32 = &H20000000
        Const WS_VISIBLE As UInt32 = &H10000000
        Const WS_DISABLED As UInt32 = &H8000000
        Const WS_CLIPSIBLINGS As UInt32 = &H4000000
        Const WS_CLIPCHILDREN As UInt32 = &H2000000
        Const WS_MAXIMIZE As UInt32 = &H1000000
        Const WS_CAPTION As UInt32 = &HC00000      ' WS_BORDER or WS_DLGFRAME  
        Const WS_BORDER As UInt32 = &H800000
        Const WS_DLGFRAME As UInt32 = &H400000
        Const WS_VSCROLL As UInt32 = &H200000
        Const WS_HSCROLL As UInt32 = &H100000
        Const WS_SYSMENU As UInt32 = &H80000
        Const WS_THICKFRAME As UInt32 = &H40000
        Const WS_GROUP As UInt32 = &H20000
        Const WS_TABSTOP As UInt32 = &H10000
        Const WS_MINIMIZEBOX As UInt32 = &H20000
        Const WS_MAXIMIZEBOX As UInt32 = &H10000
        Const WS_TILED As UInt32 = WS_OVERLAPPED
        Const WS_ICONIC As UInt32 = WS_MINIMIZE
        Const WS_SIZEBOX As UInt32 = WS_THICKFRAME

        Public PtrhWnd As IntPtr
        Public IsRestored As Boolean = False

        Sub New(hwnd As IntPtr)
            Me.PtrhWnd = hwnd
            Me.Height = 45
            Me.Width = 120
            Dim state As Long = AppBar.GetWindowLongPtr(hwnd, AppBar.GWL.GWL_STYLE)

            If (state And WS_MINIMIZE) = WS_MINIMIZE Then
                Me.BackColor = System.Drawing.SystemColors.ControlDark
                Me.IsRestored = False
            ElseIf (state And WS_VISIBLE) = WS_VISIBLE Then
                Me.BackColor = System.Drawing.SystemColors.ControlLight
                Me.IsRestored = True
            Else
                'Debug.WriteLine("Unknown State")
            End If
        End Sub

        Sub AppBarIcon_Click(sender As System.Object, e As System.EventArgs) Handles Me.Click
            'TODO: Use state
            If Me.IsRestored Then
                AppBar.ShowWindowAsync(Me.PtrhWnd, AppBar.WindowStates.MINIMIZE)
                Me.BackColor = System.Drawing.SystemColors.ControlDark
                Me.IsRestored = False
            Else
                AppBar.ShowWindowAsync(Me.PtrhWnd, AppBar.WindowStates.RESTORE)
                Me.BackColor = System.Drawing.SystemColors.ControlLight
                Me.IsRestored = True
            End If
        End Sub

        'Adapt Slot() from StartMenu to this
        Sub ArrangeMe(ByRef intLeft)
            Me.Left = intLeft
            intLeft += Me.Width + 5
        End Sub

    End Class
End Namespace
