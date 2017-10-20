VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   495
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   120
      Top             =   2280
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vkey As Long) As Integer




Private Sub Command1_Click()
Label1.Caption = ""
MsgBox "Clear Label Enabled", vbOKOnly, "System"

End Sub


Private Sub Command2_Click()
If Check1.Value = 1 Then
Command1.Enabled = True
End If
If Check1.Value = 0 Then
MsgBox "todavia no se ha activado", vbOKCancel, "System"
End If
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Command2.Enabled = False
Check1.Enabled = False
End Sub

Private Sub Timer1_Timer()

If GetAsyncKeyState(65) <> 0 Then
Label1.Caption = Label1.Caption + "a"
End If

If GetAsyncKeyState(66) <> 0 Then
Label1.Caption = Label1.Caption + "b"
End If

If GetAsyncKeyState(67) <> 0 Then
Label1.Caption = Label1.Caption + "c"
End If

If GetAsyncKeyState(68) <> 0 Then
Label1.Caption = Label1.Caption + "d"
End If

If GetAsyncKeyState(69) <> 0 Then
Label1.Caption = Label1.Caption + "e"
End If

If GetAsyncKeyState(70) <> 0 Then
Label1.Caption = Label1.Caption + "f"
End If

If GetAsyncKeyState(71) <> 0 Then
Label1.Caption = Label1.Caption + "g"
End If

If GetAsyncKeyState(72) <> 0 Then
Label1.Caption = Label1.Caption + "h"
End If

If GetAsyncKeyState(73) <> 0 Then
Label1.Caption = Label1.Caption + "i"
End If

If GetAsyncKeyState(74) <> 0 Then
Label1.Caption = Label1.Caption + "j"
End If

If GetAsyncKeyState(75) <> 0 Then
Label1.Caption = Label1.Caption + "k"
End If

If GetAsyncKeyState(76) <> 0 Then
Label1.Caption = Label1.Caption + "l"
End If

If GetAsyncKeyState(77) <> 0 Then
Label1.Caption = Label1.Caption + "m"
End If

If GetAsyncKeyState(78) <> 0 Then
Label1.Caption = Label1.Caption + "n"
End If

If GetAsyncKeyState(79) <> 0 Then
Label1.Caption = Label1.Caption + "o"
End If

If GetAsyncKeyState(80) <> 0 Then
Label1.Caption = Label1.Caption + "p"
End If

If GetAsyncKeyState(81) <> 0 Then
Label1.Caption = Label1.Caption + "q"
End If

If GetAsyncKeyState(82) <> 0 Then
Label1.Caption = Label1.Caption + "r"
End If

If GetAsyncKeyState(83) <> 0 Then
Label1.Caption = Label1.Caption + "s"
End If

If GetAsyncKeyState(84) <> 0 Then
Label1.Caption = Label1.Caption + "t"
End If

If GetAsyncKeyState(85) <> 0 Then
Label1.Caption = Label1.Caption + "u"
End If

If GetAsyncKeyState(86) <> 0 Then
Label1.Caption = Label1.Caption + "v"
End If

If GetAsyncKeyState(87) <> 0 Then
Label1.Caption = Label1.Caption + "w"
End If

If GetAsyncKeyState(88) <> 0 Then
Label1.Caption = Label1.Caption + "x"
End If

If GetAsyncKeyState(89) <> 0 Then
Label1.Caption = Label1.Caption + "y"
End If

If GetAsyncKeyState(90) <> 0 Then
Label1.Caption = Label1.Caption + "z"
End If

If GetAsyncKeyState(32) <> 0 Then
Label1.Caption = Label1.Caption + "space"
End If

End Sub


