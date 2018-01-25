VERSION 5.00
Begin VB.Form carsatuserfrm 
   Caption         =   "Madadcar - Authenticatd Users Car Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIN"
      DisabledPicture =   "carsatuser.frx":0000
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
      Picture         =   "carsatuser.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW HITLIST"
      DisabledPicture =   "carsatuser.frx":5838
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
      Picture         =   "carsatuser.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton car 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD TO HITLIST"
      DisabledPicture =   "carsatuser.frx":B070
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
      Picture         =   "carsatuser.frx":DC8C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAR"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "carsatuser.frx":108A8
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "carsatuserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
