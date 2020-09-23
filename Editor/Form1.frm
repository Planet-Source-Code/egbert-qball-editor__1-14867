VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qball editor"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "E&rase all"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Status :"
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   6240
      Width           =   1455
      Begin VB.Label Status 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max 150 blocks"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Blocks : "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Grid On"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Open"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select a color :"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1650
      Begin VB.PictureBox Brick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   270
         Index           =   5
         Left            =   120
         Picture         =   "Form1.frx":2012
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   19
         Top             =   960
         Width           =   720
      End
      Begin VB.PictureBox Brick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   270
         Index           =   4
         Left            =   840
         Picture         =   "Form1.frx":278C
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.PictureBox Brick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   270
         Index           =   3
         Left            =   840
         Picture         =   "Form1.frx":2F06
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   8
         Top             =   240
         Width           =   720
      End
      Begin VB.PictureBox Brick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   270
         Index           =   2
         Left            =   120
         Picture         =   "Form1.frx":3680
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   7
         Top             =   600
         Width           =   720
      End
      Begin VB.PictureBox Brick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   270
         Index           =   1
         Left            =   120
         Picture         =   "Form1.frx":3DFA
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mouse right :"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1650
      Begin VB.PictureBox Select2 
         AutoSize        =   -1  'True
         Height          =   270
         Left            =   480
         Picture         =   "Form1.frx":4574
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   4
         Tag             =   "1"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse left :"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1650
      Begin VB.PictureBox Select1 
         AutoSize        =   -1  'True
         Height          =   270
         Left            =   480
         Picture         =   "Form1.frx":4CEE
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   2
         Tag             =   "1"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.PictureBox Work 
      AutoRedraw      =   -1  'True
      Height          =   6705
      Left            =   1920
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   240
      Width           =   5235
      Begin VB.Shape Marker1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   210
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.PictureBox BackGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   2160
      Picture         =   "Form1.frx":5468
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Var() As BrickS

Private Sub Brick_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Select1.Picture = Brick(Index).Picture
Select1.Tag = Index
End If
If Button = 2 Then
Select2.Picture = Brick(Index).Picture
Select2.Tag = Index
End If
End Sub

Private Sub Command1_Click()
On Error GoTo 2
Dim Ret1 As String
Dim Ret2 As VbMsgBoxResult
Dim Free

1:
Ret1 = ShowSave("QBall Files (*.Qba)|*.Qba", "Select save place")

If Ret1 = "" Then Exit Sub
If Dir(Ret1) <> "" Then
Ret2 = MsgBox("File already exist over write ?", vbExclamation + vbYesNo, "Warining!")
If Ret2 = vbNo Then GoTo 1
End If
Free = FreeFile
Open Ret1 For Binary Access Write As #Free
Put Free, , UBound(Var)
Put Free, , Var
Close #Free

Exit Sub
2:
MsgBox "Error by saving level. please try agian on a other name", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub Command2_Click()
On Error GoTo 1
Dim Ret1 As String
Dim Ret2 As Long
Dim Free

Ret1 = ShowOpen("QBall Files (*.Qba)|*.Qba", "Select save place")

If Ret1 = "" Then Exit Sub
If Dir(Ret1) = "" Then
MsgBox "File not found", vbExclamation + vbOKOnly, "Error"
Exit Sub
End If
Free = FreeFile
Open Ret1 For Binary Access Read As #Free
Get Free, , Ret2
ReDim Var(Ret2)
Get Free, , Var
Close #Free
DrawVar
Exit Sub
1:
MsgBox "Error by opening level.", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub Command3_Click()
Dim Ret As VbMsgBoxResult
Ret = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Quit")
If Ret = vbYes Then End
End Sub

Private Sub Command4_Click()
If Command4.Caption = "&Grid On" Then Command4.Caption = "&Grid Off" Else Command4.Caption = "&Grid On"
End Sub

Private Sub Command5_Click()
Dim Ret1 As VbMsgBoxResult
Ret1 = MsgBox("Are you sure you want to clear all?", vbExclamation + vbYesNo, "Warning")

If Ret1 = vbYes Then
Dim A As Integer
Dim X1, Y1 As Long

Erase Var
ReDim Var(1)

'Draw background
Work.Cls
For Y1 = 0 To Me.ScaleHeight Step BackGround.ScaleHeight
For X1 = 0 To Me.ScaleWidth Step BackGround.ScaleWidth
BitBlt Work.hDC, X1, Y1, BackGround.ScaleWidth, BackGround.ScaleHeight, BackGround.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Free
Dim Ret1 As Long

CreatExtension
Dim X1, Y1 As Long
Me.Show
ReDim Var(1)

If Dir(Command) <> "" Then
Free = FreeFile
Open Command For Binary Access Read As #Free
Get Free, , Ret1
ReDim Var(Ret1)
Get Free, , Var
Close #Free
DrawVar
Exit Sub
End If

'Draw background
Work.Cls
For Y1 = 0 To Me.ScaleHeight Step BackGround.ScaleHeight
For X1 = 0 To Me.ScaleWidth Step BackGround.ScaleWidth
BitBlt Work.hDC, X1, Y1, BackGround.ScaleWidth, BackGround.ScaleHeight, BackGround.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Marker1.Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Marker1.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Marker1.Visible = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Marker1.Visible = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Marker1.Visible = False
End Sub

Private Sub Work_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Work_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Work_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X1, Y1 As Long
Dim A As Integer

Marker1.Visible = True

If Command4.Caption = "&Grid On" Then
For X1 = 0 To Work.ScaleWidth Step Marker1.Width
For Y1 = 0 To (Work.ScaleHeight - 150) Step Marker1.Height
If X > X1 And X < X1 + Marker1.Width And Y > Y1 And Y < Y1 + Marker1.Height Then
X = X1
Y = Y1
GoTo 1
End If
Next Y1
Next X1
Exit Sub
End If

1:
If X < Work.ScaleWidth - Marker1.Width Then
Marker1.Left = X
Else
Marker1.Left = Work.ScaleWidth - Marker1.Width
End If
If Y < (Work.ScaleHeight - 150) - Marker1.Height Then
Marker1.Top = Y
Else
Marker1.Top = (Work.ScaleHeight - 150) - Marker1.Height
End If

If Button <> 0 Then
For A = 1 To UBound(Var)
If Var(A).X = X And Var(A).Y = Y Then
GoTo 2
End If
Next A
If Status.Caption > 149 Then MsgBox "Bricks are on the maximum", vbInformation + vbOKOnly, "Error": Exit Sub
Status.Caption = Status.Caption + 1
ReDim Preserve Var(UBound(Var) + 1)
A = UBound(Var)
2:
Var(A).X = X
Var(A).Y = Y
Var(A).Height = Brick(1).ScaleHeight
Var(A).Width = Brick(1).ScaleWidth
Var(A).Visible = True

If Button = 1 Then
BitBlt Work.hDC, X, Y, Select1.ScaleWidth, Select1.ScaleHeight, Select1.hDC, 0, 0, vbSrcCopy ' painting Block
Var(A).Pic = Select1.Tag
End If
If Button = 2 Then
BitBlt Work.hDC, X, Y, Select2.ScaleWidth, Select2.ScaleHeight, Select2.hDC, 0, 0, vbSrcCopy ' painting Block
Var(A).Pic = Select2.Tag
End If
End If
End Sub

Private Sub Work_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Work_MouseMove(Button, Shift, X, Y)
End Sub

Function DrawVar()
Dim A As Integer
Dim X1, Y1 As Long

'Draw background
Work.Cls
For Y1 = 0 To Me.ScaleHeight Step BackGround.ScaleHeight
For X1 = 0 To Me.ScaleWidth Step BackGround.ScaleWidth
BitBlt Work.hDC, X1, Y1, BackGround.ScaleWidth, BackGround.ScaleHeight, BackGround.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1

For A = 1 To UBound(Var)
If Var(A).Visible = True Then
Select1.Picture = Brick(Var(A).Pic).Picture
BitBlt Work.hDC, Var(A).X, Var(A).Y, Select1.ScaleWidth, Select1.ScaleHeight, Select1.hDC, 0, 0, vbSrcCopy ' painting Block
End If
Next A
Status.Caption = UBound(Var)
End Function
