VERSION 5.00
Begin VB.Form authencarfrm 
   Caption         =   "Madadcar - Authenticated User Car Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "authencarfrm.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIN"
      DisabledPicture =   "authencarfrm.frx":1D4C44
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
      Left            =   7560
      MaskColor       =   &H0080FF80&
      Picture         =   "authencarfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW HITLIST"
      DisabledPicture =   "authencarfrm.frx":1DA47C
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
      Left            =   2640
      MaskColor       =   &H0080FF80&
      Picture         =   "authencarfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton car 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD to HITLIST"
      DisabledPicture =   "authencarfrm.frx":1DFCB4
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
      Left            =   2640
      MaskColor       =   &H0080FF80&
      Picture         =   "authencarfrm.frx":1E28D0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAR (AUTHENTICATED USER"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   8295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "authencarfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
