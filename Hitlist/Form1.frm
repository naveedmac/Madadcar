VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

If Command1.Caption = "start" Then
startcounting
Command1.Caption = "stop"

Else
stopcounting
Command1.Caption = "start"

End If

End Sub

Private Sub Command2_Click()
MsgBox "time is   " & Hour(total) & "hours" & vbCrLf & Minute(total) & "minutes" & vbCrLf & Second(total) & "second"

End Sub

Private Sub Form_Load()
'startcounting
'If Minute(t1) > 1 Then
't1 = 0
'MsgBox "1 minute"
'End If


End Sub
