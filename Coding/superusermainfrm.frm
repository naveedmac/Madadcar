VERSION 5.00
Begin VB.Form superusermainfrm 
   AutoRedraw      =   -1  'True
   Caption         =   "Madadcar - Super User Main Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "superusermainfrm.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton LOGOUT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      DisabledPicture =   "superusermainfrm.frx":1D4C44
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
      Picture         =   "superusermainfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton chgpassbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHANGE PASSWORD"
      DisabledPicture =   "superusermainfrm.frx":1DA47C
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
      Left            =   4800
      MaskColor       =   &H0080FF80&
      Picture         =   "superusermainfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton terminalbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TERMINALS"
      DisabledPicture =   "superusermainfrm.frx":1DFCB4
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
      Picture         =   "superusermainfrm.frx":1E28D0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton userbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "USERS"
      DisabledPicture =   "superusermainfrm.frx":1E54EC
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
      Picture         =   "superusermainfrm.frx":1E8108
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton carbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CARS"
      DisabledPicture =   "superusermainfrm.frx":1EAD24
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
      Picture         =   "superusermainfrm.frx":1ED940
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPER USER"
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
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
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
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPER USER MAIN FORM"
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
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   7455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "superusermainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub carbtn_Click()
admincarsfrm.carrecoveredbtn.Visible = False


admincarsfrm.Show

End Sub



Private Sub chgpassbtn_Click()
chgpassfrm.Show

End Sub

Private Sub Form_Load()
admincarsfrm.carrecoveredbtn.Visible = True
adminuserfrm.Command4.Visible = True
adminuserfrm.Command1.Visible = True



End Sub

Private Sub LOGOUT_Click()
Unload Me

End Sub

Private Sub terminalbtn_Click()
superterminalfrm.Show

End Sub

Private Sub userbtn_Click()
adminuserfrm.Command4.Visible = False
adminuserfrm.Command1.Visible = False


adminuserfrm.Show

End Sub
