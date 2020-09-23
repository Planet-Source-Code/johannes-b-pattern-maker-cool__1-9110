VERSION 5.00
Begin VB.Form Effects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effects"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Invert"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Emboss"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lighter/Darker"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Move pixels"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sharpen"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Blur"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Effects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'invert
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Dim tempPic As PictureTypeConstants
Public Sub ReadPixels()
X = Form1.Picture1.ScaleWidth
Y = Form1.Picture1.ScaleHeight
If X > 800 Or Y > 800 Then

    MsgBox "Please use a picture smaller then 800x800!"
    X = 0
    Y = 0
    Exit Sub
End If

'Form2.Width = Form2.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
'Form2.Height = Form2.ScaleY(Picture1.Height + 30, vbPixels, vbTwips)

    For i = 0 To Y - 1
        For j = 0 To X - 1
            pixel = GetPixel(Form1.Picture1.hdc, j, i)
            Red = pixel Mod 256
            Green = ((pixel And &HFF00) / 256&) Mod 256&
            Blue = (pixel And &HFF0000) / 65536
            ImagePixels(0, i, j) = Red
            ImagePixels(1, i, j) = Green
            ImagePixels(2, i, j) = Blue
        Next
        
        DoEvents
    Next
   
    Exit Sub
    
BadImageType:
    MsgBox Err.Description
    X = 0
    Y = 0
    Exit Sub

End Sub

Private Sub Command1_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next
Dim i As Long, j As Long
Dim Dx As Integer, Dy As Integer
Dim Red As Integer, Green As Integer, Blue As Integer

ReadPixels

    Dx = 1: Dy = 1
    T1 = Timer

    
    hBMP = CreateCompatibleBitmap(Form1.Picture1.hdc, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Form1.Picture1.hdc)
    SelectObject hDestDC, hBMP
    
    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImagePixels(0, i, j) + 0.5 * (ImagePixels(0, i, j) - ImagePixels(0, i - Dx, j - Dy))
            Green = ImagePixels(1, i, j) + 0.5 * (ImagePixels(1, i, j) - ImagePixels(1, i - Dx, j - Dy))
            Blue = ImagePixels(2, i, j) + 0.5 * (ImagePixels(2, i, j) - ImagePixels(2, i - Dx, j - Dy))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            SetPixelV hDestDC, j, i, RGB(Red, Green, Blue)
        Next
        
        DoEvents
    Next
   
    BitBlt Form1.Picture1.hdc, 1, 1, Form1.Picture1.ScaleWidth - 2, Form1.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Form1.Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
    Effects.Caption = "Effects"
End Sub


Private Sub Command2_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next

Dim i As Long, j As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim div As Integer

div = InputBox("Type a value for Blur", "Blur", 9)

ReadPixels


    T1 = Timer
    hBMP = CreateCompatibleBitmap(Form1.Picture1.hdc, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Form1.Picture1.hdc)
    SelectObject hDestDC, hBMP

    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImagePixels(0, i - 1, j - 1) + ImagePixels(0, i - 1, j) + ImagePixels(0, i - 1, j + 1) + _
            ImagePixels(0, i, j - 1) + ImagePixels(0, i, j) + ImagePixels(0, i, j + 1) + _
            ImagePixels(0, i + 1, j - 1) + ImagePixels(0, i + 1, j) + ImagePixels(0, i + 1, j + 1)
            
            Green = ImagePixels(1, i - 1, j - 1) + ImagePixels(1, i - 1, j) + ImagePixels(1, i - 1, j + 1) + _
            ImagePixels(1, i, j - 1) + ImagePixels(1, i, j) + ImagePixels(1, i, j + 1) + _
            ImagePixels(1, i + 1, j - 1) + ImagePixels(1, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            
            Blue = ImagePixels(2, i - 1, j - 1) + ImagePixels(2, i - 1, j) + ImagePixels(2, i - 1, j + 1) + _
            ImagePixels(2, i, j - 1) + ImagePixels(2, i, j) + ImagePixels(2, i, j + 1) + _
            ImagePixels(2, i + 1, j - 1) + ImagePixels(2, i + 1, j) + ImagePixels(2, i + 1, j + 1)
            
            SetPixelV hDestDC, j, i, RGB(Red / div, Green / div, Blue / div)
        Next
       
        DoEvents
    Next
   
    BitBlt Form1.Picture1.hdc, 1, 1, Form1.Picture1.ScaleWidth - 2, Form1.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Form1.Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
Effects.Caption = "Effects"

End Sub

Private Sub Command3_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next
Dim i As Long, j As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim Rx As Integer, Ry As Integer

ReadPixels

    T1 = Timer


    hBMP = CreateCompatibleBitmap(Form1.Picture1.hdc, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Form1.Picture1.hdc)
    SelectObject hDestDC, hBMP
    For i = 2 To Y - 3
        For j = 2 To X - 3
            Rx = Rnd * 4 - 2
            Ry = Rnd * 4 - 2
            Red = ImagePixels(0, i + Rx, j + Ry)
            Green = ImagePixels(1, i + Rx, j + Ry)
            Blue = ImagePixels(2, i + Rx, j + Ry)
            SetPixelV hDestDC, j, i, RGB(Red, Green, Blue)
        Next
       
        DoEvents
    Next
    
    BitBlt Form1.Picture1.hdc, 1, 1, Form1.Picture1.ScaleWidth - 2, Form1.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Form1.Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
   
    Form1.Picture1.Refresh
    
     tempPic = Form1.Picture1.Picture
     Effects.Caption = "Effects"
End Sub


Private Sub Command4_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next
Dim Brightness As Single
Dim NewColor As Long
Dim X, Y As Integer
Dim R, G, b As Integer
Dim dive As Integer
'change the brightness to a percent
dive = InputBox("Input a value. > 100 = lighter, < 100 = darker", "Lighter/Darker", 120)
Brightness = dive / 100
'run a loop through the picture to change every pixel
For X = 0 To Form1.Picture1.ScaleWidth
For Y = 0 To Form1.Picture1.ScaleHeight
'get the current color value
NewColor = GetPixel(Form1.Picture1.hdc, X, Y)
'extract the R,G,B values from the long returned by GetPixel
R = (NewColor Mod 256)
b = (Int(NewColor / 65536))
G = ((NewColor - (b * 65536) - R) / 256)
'change the RGB settings to their appropriate brightness
R = R * Brightness
b = b * Brightness
G = G * Brightness
'make sure the new variables aren't too high or too low
If R > 255 Then R = 255
If R < 0 Then R = 0
If b > 255 Then b = 255
If b < 0 Then b = 0
If G > 255 Then G = 255
If G < 0 Then G = 0
'set the new pixel
SetPixelV Form1.Picture1.hdc, X, Y, RGB(R, G, b)
'continue through the loop
Next Y
'refresh the picture box every 10 lines (a nice progress bar effect)
If X Mod 10 = 0 Then Form1.Picture1.Refresh
Next X
'final picture refresh
Form1.Picture1.Refresh
Effects.Caption = "Effects"
End Sub


Private Sub Command5_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next
Dim i As Long, j As Long
Dim Dx As Integer, Dy As Integer
Dim Red As Integer, Green As Integer, Blue As Integer

ReadPixels

    Dx = 1
    Dy = 1
    


    hBMP = CreateCompatibleBitmap(Form1.Picture1.hdc, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Form1.Picture1.hdc)
    SelectObject hDestDC, hBMP
    
    T1 = Timer
    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = Abs(ImagePixels(0, i, j) - ImagePixels(0, i + Dx, j + Dy) + 128)
            Green = Abs(ImagePixels(1, i, j) - ImagePixels(1, i + Dx, j + Dy) + 128)
            Blue = Abs(ImagePixels(2, i, j) - ImagePixels(2, i + Dx, j + Dy) + 128)
            SetPixelV hDestDC, j, i, RGB(Red, Green, Blue)
        Next
       
        DoEvents
    Next
    
    BitBlt Form1.Picture1.hdc, 1, 1, Form1.Picture1.ScaleWidth - 2, Form1.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Form1.Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
    Effects.Caption = "Effects"
End Sub


Private Sub Command6_Click()
Effects.Caption = "PLEASE WAIT..."
On Error Resume Next
For X = 0 To Form1.Picture1.ScaleWidth
For Y = 0 To Form1.Picture1.ScaleHeight
SetPixel Form1.Picture1.hdc, X, Y, 16777215 - GetPixel(Form1.Picture1.hdc, X, Y)
Next Y
Next X
Form1.Picture1.Refresh
Effects.Caption = "Effects"
End Sub


