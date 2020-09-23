Attribute VB_Name = "Module1"
'**********************************************************
'These API are used for Filter effects
'**********************************************************
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'**********************************************************
'ImagePixels will hold the image[0-2] - Hold Red, Green & Blue
'Pixel values, 0 to 800 - Rows, 0 to 800 - Cols
Public ImagePixels(0 To 2, 0 To 800, 0 To 800) As Integer
''**********************************************************
Public X As Integer, Y As Integer
Public FilterNorm As Integer, FilterBias As Integer
Public CustomFilter(4, 4) As Single
Public FilterCancel As Boolean
Public picIndex As Integer
Option Explicit
'
' Declarations
'

Public Const SW_HIDE = 0    ' Hide Window
Public Const SW_SHOW = 5    ' Show Window

Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
   
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Sub DisplayTaskBar(ByVal bShow As Boolean)
    Dim lTaskBarHWND As Long
    Dim lRet As Long
    Dim lFlags As Long
'
' Show / hide the taskbar
'
    On Error GoTo vbErrorHandler
    
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    
    
    lTaskBarHWND = FindWindow("Shell_TrayWnd", "")
    lRet = ShowWindow(lTaskBarHWND, lFlags)
    
    If lRet < 0 Then
    '
    ' Handle error from api
    '
    End If

    Exit Sub
    
vbErrorHandler:
'
' Handle Errors here
'
End Sub

Public Sub DisplayDeskTopIcons(ByVal bShow As Boolean)
    Dim lDesktopHwnd As Long
    Dim lFlags As Long
'
' Show / Hide the Desktop Icons
'
    On Error Resume Next
    lDesktopHwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    If lDesktopHwnd = 0 Then
        ' raise an error ! You have no desktop !!!
        Exit Sub
    End If
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    ShowWindow lDesktopHwnd, lFlags
End Sub


