VERSION 5.00
Begin VB.Form admincarsfrm 
   Caption         =   "Madadcar - Administrator Cars Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton LOGOUT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIN"
      DisabledPicture =   "carsfrm.frx":0000
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
      Picture         =   "carsfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton carrecoveredbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CAR RECOVERED"
      DisabledPicture =   "carsfrm.frx":5838
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton viewdbbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW DB"
      DisabledPicture =   "carsfrm.frx":B070
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":DC8C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton viewhitlistbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIEW HITLIST"
      DisabledPicture =   "carsfrm.frx":108A8
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":134C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton addhitlistbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD TO DB"
      DisabledPicture =   "carsfrm.frx":160E0
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":18CFC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton updatedbbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "UPDATE DB"
      DisabledPicture =   "carsfrm.frx":1B918
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":1E534
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton updatehitlistbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "UPDATE HITLIST"
      DisabledPicture =   "carsfrm.frx":21150
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "carsfrm.frx":23D6C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
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
      Left            =   5400
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "carsfrm.frx":26988
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "admincarsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addhitlistbtn_Click()
addmdfrm.Show

End Sub

Private Sub carrecoveredbtn_Click()
carrecoveredfrm.Show

End Sub

Private Sub LOGOUT_Click()
Me.Hide
If domainstat = 1 Then
adminmainfrm.Show
Else
superusermainfrm.Show
End If
End Sub

Private Sub updatedbbtn_Click()
updatemdfrm.Show

End Sub

Private Sub updatehitlistbtn_Click()
updatehitlistfrm.Show

End Sub

Private Sub viewdbbtn_Click()
viewmdfrm.Show

End Sub

Private Sub viewhitlistbtn_Click()
viewhitlistfrm.Show

End Sub
