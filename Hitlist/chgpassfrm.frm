VERSION 5.00
Begin VB.Form chgpassfrm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Picture         =   "chgpassfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1400
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Picture         =   "chgpassfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1400
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   5160
      TabIndex        =   10
      Top             =   2280
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Image Image5 
      Height          =   540
      Left            =   5040
      Picture         =   "chgpassfrm.frx":5838
      Top             =   5640
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OLD PASSWORD"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   5040
      Picture         =   "chgpassfrm.frx":8454
      Top             =   4200
      Width           =   2160
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   5040
      Picture         =   "chgpassfrm.frx":B070
      Top             =   3480
      Width           =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   5040
      Picture         =   "chgpassfrm.frx":DC8C
      Top             =   4920
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "chgpassfrm.frx":108A8
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "chgpassfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim change As New Users

Private Sub Command1_Click()
MsgBox change.changepasswsord(Text1.Text, Text2.Text, Text3.Text, Text4.Text)

End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
End Sub


Private Sub Recordset()

    
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.LockType = adLockBatchOptimistic
    RS.Source = SQL
    RS.ActiveConnection = CN
    RS.Open
End Sub
Public Sub COMmand()
COM.CommandType = adCmdText
COM.CommandText = SQL
COM.ActiveConnection = CN
COM.Execute

End Sub
