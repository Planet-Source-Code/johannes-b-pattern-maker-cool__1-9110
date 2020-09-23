VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Pattern maker (Johannes B.  Email: JB_Rulez_54@hotmail.com)"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "Save as..."
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW!"
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Effects"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "20"
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "RND*"
         Height          =   195
         Left            =   5280
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CLS"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Gradient"
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   2880
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000C0&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back color (CLS)"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pattern color"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "5"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "6"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "5"
         Top             =   120
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   4200
         X2              =   4320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Size:"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "+ Height:"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "+ Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, b As Integer

Private Sub Check2_Click()
If Check2.Value = 1 Then
Command3.Enabled = False
Text4.Enabled = True
MsgBox "Gradient colors = Black to Red"
Else
Command3.Enabled = True
Text4.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Text5.Enabled = True
Else
Text5.Enabled = False
End If
End Sub

Private Sub Command1_Click()
On Error GoTo JB
If Check1.Value = 1 Then
Picture1.Cls
End If

If Check3.Value = 1 Then
1:
Text1.Text = Rnd * Text5.Text
Text2.Text = Rnd * Text5.Text
Text3.Text = Rnd * Text5.Text

If Text1.Text < 3 Or Text2.Text < 3 Or Text3.Text < 3 Then
GoTo 1
End If


End If

Form1.Caption = "DRAWING..."

A = 0
b = 0


Do
A = A + Text1.Text

If A >= Picture1.ScaleWidth + 10 Then
A = 0
b = b + Text2.Text
End If
If Check2.Value = 1 Then
Picture1.Circle (A, b), Text3.Text, b / Text4.Text * 4
Else
Picture1.Circle (A, b), Text3.Text
End If

Loop Until b >= Picture1.ScaleHeight + 10

Form1.Caption = "Pattern maker"
Exit Sub
JB:
MsgBox "ERROR! Check values!"
Form1.Caption = "ERROR!"
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo nisse
CM.Filter = "Windows bitmap (*.BMP)|*.bmp"
CM.ShowSave
SavePicture Picture1.Image, CM.FileName
MsgBox "Picture saved! Width = " & Picture1.Width & " Height = " & Picture1.Height
Exit Sub
nisse:
Exit Sub
End Sub


Private Sub Command3_Click()
On Error Resume Next
CM.ShowColor
Picture1.ForeColor = CM.Color
Picture2.BackColor = CM.Color

End Sub

Private Sub Command4_Click()
On Error Resume Next
CM.ShowColor
Picture1.BackColor = CM.Color
Picture3.BackColor = CM.Color
End Sub


Private Sub Command5_Click()
Effects.Show
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Width = Form1.ScaleWidth
Picture1.Height = Form1.ScaleHeight - 50
End Sub

Private Sub Picture1_Click()
MsgBox "SIZE: Width = " & Picture1.Width & ". Height = " & Picture1.Height
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Text1.Text = 0 Then Text1.Text = 1
End Sub

Private Sub Text2_Change()
On Error Resume Next
If Text2.Text = 0 Then Text2.Text = 1
End Sub

Private Sub Text3_Change()
On Error Resume Next
If Text3.Text = 0 Then Text3.Text = 1
End Sub


Private Sub Text4_Change()
On Error Resume Next
If Text4.Text = 0 Then Text4.Text = 1
End Sub

Private Sub Text5_Change()
On Error Resume Next
If Text5.Text = 0 Then Text5.Text = 1
End Sub


