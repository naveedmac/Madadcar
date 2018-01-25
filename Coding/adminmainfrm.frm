VERSION 5.00
Begin VB.Form adminmainfrm 
   Caption         =   "Madadcar - Administrator Main Page"
   ClientHeight    =   8490
   ClientLeft      =   2295
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   Picture         =   "adminmainfrm.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton carbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CARS"
      DisabledPicture =   "adminmainfrm.frx":1D4C44
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
      Left            =   3360
      MaskColor       =   &H0080FF80&
      Picture         =   "adminmainfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton usrbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "USERS"
      DisabledPicture =   "adminmainfrm.frx":1DA47C
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
      Left            =   3360
      MaskColor       =   &H0080FF80&
      Picture         =   "adminmainfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton terminalbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TERMINALS"
      DisabledPicture =   "adminmainfrm.frx":1DFCB4
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
      Left            =   3360
      MaskColor       =   &H0080FF80&
      Picture         =   "adminmainfrm.frx":1E28D0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton chgpassbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GPS "
      DisabledPicture =   "adminmainfrm.frx":1E54EC
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
      Left            =   3360
      MaskColor       =   &H0080FF80&
      Picture         =   "adminmainfrm.frx":1E8108
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton LOGOUTbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      DisabledPicture =   "adminmainfrm.frx":1EAD24
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
      Left            =   7800
      MaskColor       =   &H0080FF80&
      Picture         =   "adminmainfrm.frx":1ED940
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN MAIN PAGE"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "adminmainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub carbtn_Click()
admincarsfrm.Show
Unload Me

End Sub

Private Sub chgpassbtn_Click()
gpsfrm.Show
Unload Me

End Sub

Private Sub LOGOUTbtn_Click()
logoutfrm.Show
Unload Me
End Sub

Private Sub terminalbtn_Click()
adminterminalfrm.Show
Unload Me

End Sub

Private Sub usrbtn_Click()
adminuserfrm.Show
Unload Me
End Sub
