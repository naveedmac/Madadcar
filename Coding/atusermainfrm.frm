VERSION 5.00
Begin VB.Form atusermainfrm 
   Caption         =   "Madadcar - Authenticated User Main Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   Picture         =   "atusermainfrm.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton LOGOUTbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      DisabledPicture =   "atusermainfrm.frx":1D4C44
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
      Left            =   8280
      MaskColor       =   &H0080FF80&
      Picture         =   "atusermainfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton chgpassbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHANGE PASSWORD"
      DisabledPicture =   "atusermainfrm.frx":1DA47C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4920
      MaskColor       =   &H0080FF80&
      Picture         =   "atusermainfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton carbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CARS"
      DisabledPicture =   "atusermainfrm.frx":1DFCB4
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
      Left            =   4920
      MaskColor       =   &H0080FF80&
      Picture         =   "atusermainfrm.frx":1E28D0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHENTICATED USER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHENTICATED USER MAIN PAGE"
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
      Height          =   1455
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "atusermainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub carbtn_Click()
viewhitlistatfrm
End Sub

Private Sub chgpassbtn_Click()
chgpassfrm.Show

End Sub

Private Sub LOGOUTbtn_Click()
logoutfrm.Show

End Sub
