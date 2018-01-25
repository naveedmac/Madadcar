VERSION 5.00
Begin VB.Form adminterminalfrm 
   Caption         =   "Madadcar - Admin Terminal Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "adminterminalfrm.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton LOGOUT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIN PAGE"
      DisabledPicture =   "adminterminalfrm.frx":1D4C44
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
      Left            =   8040
      MaskColor       =   &H0080FF80&
      Picture         =   "adminterminalfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHANGE STATUS"
      DisabledPicture =   "adminterminalfrm.frx":1DA47C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4800
      MaskColor       =   &H0080FF80&
      Picture         =   "adminterminalfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW TERMINAL"
      DisabledPicture =   "adminterminalfrm.frx":1DFCB4
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
      Left            =   4800
      MaskColor       =   &H0080FF80&
      Picture         =   "adminterminalfrm.frx":1E28D0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD TERMINAL"
      DisabledPicture =   "adminterminalfrm.frx":1E54EC
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
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "adminterminalfrm.frx":1E8108
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN TERMINAL PAGE"
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
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   7095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "adminterminalfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
enableterminalfrm.Show

End Sub

Private Sub Command3_Click()
viewterminalfrm.Show

End Sub

Private Sub Command4_Click()
addterminalfrm.Show

End Sub

Private Sub LOGOUT_Click()
If domainstat = 1 Then
adminmainfrm.Show
Else
superusermainfrm.Show
End If

Unload Me
'adminmainfrm.Show

End Sub
