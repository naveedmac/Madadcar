VERSION 5.00
Begin VB.Form adminuserfrm 
   Caption         =   "Madadcar - Admin User Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW USER"
      DisabledPicture =   "adminuserfrm.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4440
      MaskColor       =   &H0080FF80&
      Picture         =   "adminuserfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EDIT USER"
      DisabledPicture =   "adminuserfrm.frx":5838
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4440
      MaskColor       =   &H0080FF80&
      Picture         =   "adminuserfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD USER"
      DisabledPicture =   "adminuserfrm.frx":B070
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4440
      MaskColor       =   &H0080FF80&
      Picture         =   "adminuserfrm.frx":DC8C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIN"
      DisabledPicture =   "adminuserfrm.frx":108A8
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8640
      MaskColor       =   &H0080FF80&
      Picture         =   "adminuserfrm.frx":134C4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER PAGE"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "adminuserfrm.frx":160E0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "adminuserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
edituserfrm.Show

End Sub

Private Sub Command2_Click()
viewusersfrm.Show

End Sub

Private Sub Command4_Click()
adduserfrm.Show

End Sub

Private Sub OK_Click()
'adminmainfrm.Show
If domainstat = 1 Then
adminmainfrm.Show
Else
superusermainfrm.Show
End If
End Sub
