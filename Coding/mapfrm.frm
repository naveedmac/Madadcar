VERSION 5.00
Begin VB.Form mapfrm 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   2040
      Picture         =   "mapfrm.frx":0000
      ScaleHeight     =   6435
      ScaleWidth      =   7275
      TabIndex        =   1
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude "
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
      Left            =   9120
      TabIndex        =   5
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude   "
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
      Left            =   9120
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -120
      Picture         =   "mapfrm.frx":13D2E
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   5400
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Map Car Location"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   8895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "mapfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim latitude As Double
Dim longitude As Double
Private Sub Command2_Click()
Picture1.Circle (Text1.Text, Text2.Text), 1, vbRed

End Sub

Private Sub Command1_Click()
MsgBox gpsfrm.publatitude & " " & gpsfrm.publongitude
Picture1.Refresh
Text3.Text = gpsfrm.publatitude
Text4.Text = gpsfrm.publongitude
Picture1.DrawWidth = 2
Picture1.Circle (CDbl(gpsfrm.publatitude), CDbl(gpsfrm.publongitude)), 0.5, vbRed

End Sub

Private Sub Form_Load()
Picture1.ScaleMode = 0
Picture1.ScaleLeft = 61.92861111
Picture1.ScaleTop = 24.72527778
Picture1.ScaleWidth = 14.60333344
Picture1.ScaleHeight = 0.088055777
'longitude = gpsfrm.txtlong.Text
'latitude = gpsfrm.txtlat.Text
'MsgBox gpsfrm.publatitude & " " & gpsfrm.publongitude



'Picture1.DrawWidth = 2
'Picture1.Circle (gpsfrm.publatitude, gpsfrm.publongitude), 0.5, vbRed
End Sub

Private Sub Picture1_Click()
Picture1.DrawWidth = 2
Picture1.Circle (CSng(Text4.Text), CSng(Text3.Text)), 0.5, vbRed
MsgBox "hi"
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.Text = X
Text1.Text = Y

End Sub

