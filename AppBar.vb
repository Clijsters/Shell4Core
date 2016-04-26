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

''' <summary>Main AppBar for Shell4Core</summary>
Public Class AppBar
    Inherits System.Windows.Forms.Form

    Public ABE As ABEdge
    Private uCallBack As Integer
    Private BarIsDocked As Boolean = False
    Private ScreenToDock As Integer = 0
    ''' <summary>Time in milliseconds, AppBar will enumerate open Windows</summary>
    Const enumWinTimer = 2000
    ''' <summary>Holds a List to order AppBarIcons by their hwnd. This omits changes in order, when WindowStates change.</summary>
    Dim openWindows As New List(Of AppBarIcon)
    'TODO: Rename this. It's just for AppBarIcons
    Dim intLeft As Integer
    'Needed for appbar allocation/fixing
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style And Not &HC00000 ' WS_CAPTION
            cp.Style = cp.Style And Not &H800000 ' WS_BORDER
            cp.ExStyle = &H80 Or &H8 ' WS_EX_TOOLWINDOW | WS_EX_TOPMOST
            Return cp
        End Get
    End Property

    ''' <summary>Creates a new AppBar Instance.</summary>
    ''' <param name="Edge">AppBar.ABEdge: Defines the Edge to dock on</param>
    Public Sub New(ByVal Edge As ABEdge)
        ABE = Edge
        InitializeComponent()
    End Sub

#Region "Designer"

    Public WithEvents startButton As System.Windows.Forms.Button
    Public WithEvents Panel1 As System.Windows.Forms.Panel
    Public WithEvents TimeLabel As System.Windows.Forms.Label

    Private Sub InitializeComponent()
        console.writeline("InitializeComponent")
        Me.SuspendLayout()

        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Location = New System.Drawing.Point(70, 10)
        Me.Panel1.Name = "RunningApps"
        Me.Panel1.Size = New System.Drawing.Size(834, 50)

        'TODO: Create DateLabel, group them in panel, center them.
        Me.TimeLabel = New System.Windows.Forms.Label()
        Me.TimeLabel.Text = "XX:XX" 'Is changed with every WindowEnumeration
        Me.TimeLabel.Size = New System.Drawing.Size(38, 13)
        Me.TimeLabel.Location = New System.Drawing.Point(1890, 10)
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        Me.startButton = New System.Windows.Forms.Button()
        Me.startButton.Location = New System.Drawing.Point(10, 10)
        Me.startButton.Name = "StartButton"
        Me.startButton.Text = "Start"
        Me.startButton.Size = New System.Drawing.Size(50, 50)
        Me.startButton.TabIndex = 1

        Me.ClientSize = New System.Drawing.Size(1940, 70)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "AppBar"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Controls.AddRange(New Control() {Me.startButton, Me.Panel1, Me.TimeLabel})

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
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
        Public uEdge As Integer
        Public rc As RECT
        Public lParam As IntPtr
    End Structure

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
        ABN_POSCHANGED
        ABN_FULLSCREENAPP
        ABN_WINDOWARRANGE
    End Enum

    Enum ABEdge
        ABE_LEFT = 0
        ABE_TOP
        ABE_RIGHT
        ABE_BOTTOM
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
#End Region

#Region "AppBarAPIs"
    '''TODO: Convert all declares to DLLImport, it's more readable for non-VB programmers.
    <DllImport("User32.dll", EntryPoint:="GetWindowLong")> _
    Private Shared Function GetWindowLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function
    <DllImport("User32.dll", EntryPoint:="GetWindowLongPtr")> _
    Private Shared Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    Private Declare Auto Function RegisterWindowMessage Lib "User32.dll" (ByVal msg As String) As Integer
    Public Declare Function SHAppBarMessage Lib "Shell32.dll" Alias "SHAppBarMessage" (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As System.UInt32
    Public Declare Function MoveWindow Lib "User32.dll" Alias "MoveWindow" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal repaint As Boolean) As Boolean
    Public Declare Function ShowWindowAsync Lib "User32.dll" Alias "ShowWindowAsync" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    Public Delegate Function EnumWindowProc(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    Public Declare Function EnumWindows Lib "User32.dll" (ByVal Adress As EnumWindowProc, ByVal y As Integer) As Integer
    Public Declare Function IsWindowVisible Lib "User32.dll" (ByVal hwnd As IntPtr) As Boolean

    'It has to be shared for AppBarIcon.New()
    <DllImport("User32.dll", EntryPoint:="GetWindowText")>
    Public Shared Function GetWindowText(ByVal hwnd As Integer, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    End Function
#End Region

    'http://www.pinvoke.net/default.aspx/user32/GetWindowLongPtr.html
    Public Shared Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 8 Then
            Return GetWindowLongPtr64(hWnd, nIndex)
        Else
            Return GetWindowLongPtr32(hWnd, nIndex)
        End If
    End Function

    Public Sub RegisterBar()
        Dim abd As New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle

        If Not BarIsDocked Then
            uCallBack = RegisterWindowMessage("AppBarMessage")
            abd.uCallbackMessage = uCallBack

            Dim ret As System.UInt32 = SHAppBarMessage(CInt(ABMsg.ABM_NEW), abd) 'Minding the unsigned Integer
            BarIsDocked = True 'Are you sure?
            ABSetPos()
        Else
            SHAppBarMessage(CInt(ABMsg.ABM_REMOVE), abd)
            BarIsDocked = False
        End If
    End Sub

    'https://msdn.microsoft.com/library/en-us/shellcc/platform/Shell/programmersguide/shell_int/shell_int_programming/appbars.asp
    'http://stackoverflow.com/questions/14698755/appbar-multi-monitor
    Private Sub ABSetPos()
        Dim scrReference As Rectangle
        Dim abd As New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        abd.uEdge = CInt(ABE)

        'TODO: Let the user select the preferred screen
        Dim scr As Screen = Screen.AllScreens(Me.ScreenToDock)
        scrReference = scr.Bounds

        If abd.uEdge = CInt(ABEdge.ABE_LEFT) Or abd.uEdge = CInt(ABEdge.ABE_RIGHT) Then
            'Saved it for later. Currently we are only supporting TOP and BOTTOM
            'abd.rc.top = 0
            'abd.rc.bottom = scrReference.Height
            'If abd.uEdge = CInt(ABEdge.ABE_LEFT) Then
            'abd.rc.left = 0
            'abd.rc.right = Size.Width
            'Else
            'abd.rc.right = scrReference.Width
            'abd.rc.left = abd.rc.right - Size.Width
            'End If
        Else
            abd.rc.left = scrReference.Left
            abd.rc.right = scrReference.Right
            If abd.uEdge = CInt(ABEdge.ABE_TOP) Then
                abd.rc.top = scrReference.Top
                abd.rc.bottom = Size.Height
            Else
                abd.rc.bottom = scrReference.Bottom
                abd.rc.top = abd.rc.bottom - Size.Height
            End If
        End If

        ' Query the system for an approved size and position. 
        SHAppBarMessage(CInt(ABMsg.ABM_QUERYPOS), abd)
        ' Adjust the rectangle, depending on the edge to which the appbar is (going to be) anchored.
        Select Case abd.uEdge
            Case CInt(ABEdge.ABE_LEFT)
                abd.rc.right = abd.rc.left + Size.Width
            Case CInt(ABEdge.ABE_RIGHT)
                abd.rc.left = abd.rc.right - Size.Width
            Case CInt(ABEdge.ABE_TOP)
                abd.rc.bottom = abd.rc.top + Size.Height
            Case CInt(ABEdge.ABE_BOTTOM)
                abd.rc.top = abd.rc.bottom - Size.Height
        End Select

        'Pass the final bounding rectangle to the system
        SHAppBarMessage(CInt(ABMsg.ABM_SETPOS), abd)
        'Now, let's move and resize to it
        MoveWindow(abd.hWnd, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, True)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = uCallBack Then
            Select Case m.WParam.ToInt32()
                Case CInt(ABNotify.ABN_POSCHANGED)
                    ABSetPos()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Public Sub ticked(sender As Object, e As system.eventargs)
        Me.OpenWindows.Clear()
        intLeft = 0
        Me.Panel1.SuspendLayout()
        Me.Panel1.Controls.Clear()
        If EnumWindows(AddressOf EnumCallBack, 0) Then
            Dim q As IEnumerable(Of AppBarIcon) = From i In Me.openWindows
                    Order By i.PtrhWnd.toString() 'String implements IComparable
            console.writeline("Refreshing open windows")
            For Each ItemSMI As AppBarIcon In q
                ItemSMI.ArrangeMe(intLeft)
                Me.Panel1.Controls.Add(ItemSMI)
            Next
        Else
            'Throw New Exception
        End If
        Me.Panel1.ResumeLayout(False)
    End Sub


    ''' <summary>Is called by EnumWindows(). Adds new AppBarIcon based on hWnd</summary>
    ''' <param name="hWnd">Handle to Window found.</param>
    ''' <param name="lParam">Not in use</param>
    Function EnumCallBack(ByVal hWnd As IntPtr, ByVal lParam As System.IntPtr) As Boolean
        Try
            Dim sbTitle As New System.Text.StringBuilder("", 256)
            If IsWindowVisible(hwnd) Then
                GetWindowText(hwnd, sbTitle, 256)
                Dim strTitle As String = sbTitle.ToString()
                If Not (String.IsNullOrEmpty(strTitle) Or strTitle.startsWith("Start")) Then
                    Dim newBtn As New AppBarIcon(hwnd)
                    newBtn.text = strTitle
                    Me.openWindows.Add(newBtn)
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class

Class AppBarIcon
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
    Private IsRestored As Boolean = False

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
        switchState()
    End Sub

    'Currently, intLeft is Public and held by AooBar...
    Sub ArrangeMe(ByRef intLeft)
        Me.Left = intLeft
        intLeft += Me.Width + 5
    End Sub

    Sub switchState()
        'TODO: Use state
        If Me.IsRestored Then
            Debug.WriteLine("Minimizing" & VbTab & Me.PtrhWnd.toString() & VbTab & Me.Text)
            AppBar.ShowWindowAsync(Me.PtrhWnd, AppBar.WindowStates.MINIMIZE)
            Me.BackColor = System.Drawing.SystemColors.ControlDark
            Me.IsRestored = False
        Else
            Debug.WriteLine("Restoring" & VbTab & Me.PtrhWnd.toString() & VbTab & Me.Text)
            AppBar.ShowWindowAsync(Me.PtrhWnd, AppBar.WindowStates.RESTORE)
            Me.BackColor = System.Drawing.SystemColors.ControlLight
            Me.IsRestored = True
        End If
    End Sub

End Class
