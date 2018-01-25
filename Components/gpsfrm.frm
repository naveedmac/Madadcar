VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form gpsfrm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Map"
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
      Left            =   7440
      Picture         =   "gpsfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7680
      Width           =   1400
   End
   Begin VB.TextBox txtlat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtdatum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Text            =   "        WGS 84"
      Top             =   5730
      Width           =   1935
   End
   Begin VB.TextBox txtbaud 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Text            =   "          4800"
      Top             =   5130
      Width           =   1935
   End
   Begin VB.TextBox txtport 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Text            =   "      Com Port 1"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtsatellite 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtlong 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Top             =   2730
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
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
      Left            =   9000
      Picture         =   "gpsfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1400
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
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
      Left            =   4560
      Picture         =   "gpsfrm.frx":5838
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1400
   End
   Begin VB.TextBox txtinput 
      Height          =   1815
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   480
      Top             =   3120
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
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
      Left            =   10080
      Picture         =   "gpsfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
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
      Left            =   8040
      Picture         =   "gpsfrm.frx":B070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1400
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
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
      Left            =   1320
      TabIndex        =   18
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Baud Rate  "
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
      Left            =   1320
      TabIndex        =   17
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serail Port "
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
      Left            =   1320
      TabIndex        =   16
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No of Satellites"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   3840
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
      Left            =   1320
      TabIndex        =   14
      Top             =   3240
      Width           =   2655
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
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GPS Data"
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
      Left            =   8040
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Car Location"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "gpsfrm.frx":DC8C
      Top             =   -240
      Width           =   12000
   End
End
Attribute VB_Name = "gpsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public publatitude As Double
Public publongitude As Double
Private Sub Command1_Click()
If MSComm1.PortOpen = True Then
MsgBox "Port already opened..!!"
Exit Sub
End If
    
MSComm1.CommPort = 1
MSComm1.Settings = "4800,n,8,1"
MSComm1.PortOpen = True
End Sub

Private Sub Command2_Click()
If MSComm1.PortOpen = False Then
MsgBox "Port already closed..!!"
Exit Sub
End If
MSComm1.PortOpen = False

End Sub

Private Sub Command3_Click()
txtlong.Text = ""
txtlat.Text = ""
txtsatellite.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Command5_Click()
MSComm1.PortOpen = False
Timer1.Interval = 0

txtlat.SelStart = 0
txtlat.SelLength = 3
deglat = CInt(txtlat.SelText)
MsgBox deglat
txtlat.SelStart = 3
txtlat.SelLength = 2
minlat = CInt(txtlat.SelText) * 60
txtlat.SelStart = InStr(txtlat.Text, ".")
txtlat.SelLength = Len(txtlat.Text) - txtlat.SelStart
seclat = CDbl(txtlat.SelText)
deglat = deglat & "." & (minlat + seclat)
MsgBox deglat

txtlong.SelStart = 0
txtlong.SelLength = 2
'deglong = txtlong.SelText
deglong = CInt(txtlong.SelText)
MsgBox deglong
txtlong.SelStart = 2
txtlong.SelLength = 2
'minlong = txtlong.SelText
minlong = CInt(txtlong.SelText) * 60
MsgBox minlong
txtlong.SelStart = InStr(txtlong.Text, ".")
txtlong.SelLength = Len(txtlong.Text) - txtlong.SelStart
'seclong = txtlong.SelText
seclong = CDbl(txtlong.SelText)
MsgBox seclong
deglong = deglong & "." & (minlong + seclong)
MsgBox deglong
publatitude = deglat
publongitude = deglong

Me.Hide
mapfrm.Show
End Sub

Private Sub Timer1_Timer()
If MSComm1.PortOpen = True Then
txtinput.Text = ""
txtinput.Text = MSComm1.Input
txtinput.SelStart = 0
seekstring = "$GPGLL"

textstart = InStr(txtinput.Text, seekstring)
txtinput.SelStart = textstart - 1
txtinput.SelLength = Len(txtinput.Text) - textstart
'MsgBox txtinput.SelText
CountWords (txtinput.SelText)

End If


End Sub

Function CountWords(strText As String) As Long
   ' This procedure counts the number of words in a string.
    Dim astrsentance() As String
    Dim astrWords() As String
     astrsentance = Split(strText, "$")
     
     astrWords = Split(astrsentance(2), ",")
     txtlong.Text = astrWords(2)
     txtlat.Text = astrWords(4)
     txtsatellite.Text = astrWords(7)
     
     
     
     ' Count number of elements in array -- this will be the
     ' number of words.
   CountWords = UBound(astrWords) - LBound(astrWords)


End Function

