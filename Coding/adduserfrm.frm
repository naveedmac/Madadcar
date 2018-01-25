VERSION 5.00
Begin VB.Form adduserfrm 
   Caption         =   "Madadcar - Add User Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "adduserfrm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox uidtxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   18
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton adduser 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD USER"
      DisabledPicture =   "adduserfrm.frx":28B3E
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
      Picture         =   "adduserfrm.frx":2B75A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      DisabledPicture =   "adduserfrm.frx":2E376
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
      Left            =   6720
      MaskColor       =   &H0080FF80&
      Picture         =   "adduserfrm.frx":30F92
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.ComboBox didcombo 
      Height          =   315
      Left            =   7440
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox addresstxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox nictxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox passwordtxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox unametxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox telephonetxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
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
      Left            =   4920
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID :"
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
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD USER"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   1440
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TELEPHONE :"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DOMAIN :"
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
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS :"
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
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIC :"
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
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME :"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   9000
      Left            =   0
      Picture         =   "adduserfrm.frx":33BAE
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   4800
      Picture         =   "adduserfrm.frx":2087F2
      Top             =   3240
      Width           =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   9000
      Left            =   0
      Picture         =   "adduserfrm.frx":20B40E
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   4800
      Picture         =   "adduserfrm.frx":22DEC2
      Top             =   3240
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD USER"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   6000
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "adduserfrm.frx":230ADE
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image7 
      Height          =   9000
      Left            =   0
      Picture         =   "adduserfrm.frx":253592
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "adduserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newuser As New Users
Dim fetchdomain As New Domain
Dim SQL As String




Private Sub Recordset()

    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.LockType = adLockBatchOptimistic
    RS.Source = SQL
    RS.ActiveConnection = CN
    RS.Open
    
End Sub
Private Sub adduser_Click()
If didcombo.Text = "" Or unametxt.Text = "" Or passwordtxt.Text = "" Or nictxt.Text = "" Or addresstxt.Text = "" Or telephonetxt.Text = "" Then
MsgBox "Input All the values..!!"
'pswrdtxt.SetFocus
SendKeys "{Home}+{End}"
Exit Sub
End If

If uidtxt.Text <> "" Then
MsgBox "You dont have to enter UserID..!!"
uidtxt.Text = ""
SendKeys "{Home}+{End}"
Exit Sub
End If


newuser.domainId = didcombo.ItemData(didcombo.ListIndex)
newuser.userName = unametxt.Text
newuser.password = passwordtxt.Text
newuser.NICofuser = nictxt.Text
newuser.address = addresstxt.Text
newuser.telephone = telephonetxt.Text

Dim var

uidtxt.Text = newuser.adduser()
MsgBox var

End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1

'******************************
SQL = fetchdomain.getdomain
'MsgBox "hi"
Call Recordset
RS.MoveFirst
While RS.EOF = False
didcombo.AddItem (RS(1).Value)
'prodnamecombo.AddItem (RS(1).Value)
didcombo.ItemData(didcombo.NewIndex) = RS(0).Value
RS.MoveNext
Wend
RS.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
CN.Close

End Sub

Private Sub OK_Click()
Unload Me
End Sub

Private Sub telephonetxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
KeyAscii = 8
Else
If KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 8 Then
KeyAscii = 0
End If
End If

End Sub
