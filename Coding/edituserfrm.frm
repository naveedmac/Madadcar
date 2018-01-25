VERSION 5.00
Begin VB.Form edituserfrm 
   Caption         =   "Madadcar - Edit User form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "edituserfrm.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox fonetxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7320
      TabIndex        =   16
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton search 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SEARCH"
      DisabledPicture =   "edituserfrm.frx":1D4C44
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8400
      MaskColor       =   &H0080FF80&
      Picture         =   "edituserfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.ComboBox didcombo 
      Height          =   315
      Left            =   7320
      TabIndex        =   14
      Top             =   6840
      Width           =   3015
   End
   Begin VB.TextBox uidtxt 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox unametxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox passwordtxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7320
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7440
      Picture         =   "edituserfrm.frx":1DA47C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1400
   End
   Begin VB.TextBox nictxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.TextBox addresstxt 
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton updatebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UPDATE"
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
      Left            =   5640
      Picture         =   "edituserfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DOMAIN :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIC :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT USER"
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
      Height          =   975
      Left            =   4320
      TabIndex        =   13
      Top             =   1320
      Width           =   6975
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "edituserfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim listdomain As New Domain
Dim edituser As New Users


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

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
SQL = listdomain.getdomain
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
Private Sub Recordset()

    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.LockType = adLockBatchOptimistic
    RS.Source = SQL
    RS.ActiveConnection = CN
    RS.Open
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
CN.Close

End Sub

Private Sub search_Click()
edituser.userId = uidtxt.Text
edituser.getuser
unametxt.Text = edituser.userName
passwordtxt.Text = edituser.password
nictxt.Text = edituser.NICofuser
addresstxt.Text = edituser.address
didcombo.ListIndex = (edituser.domainId - 1)
If edituser.telephone = 0 Then
fonetxt.Text = ""
Else
fonetxt.Text = edituser.telephone
End If


End Sub

Private Sub updatebtn_Click()

Dim reply
reply = MsgBox("Do You Want to Update User?", vbYesNo)
If reply = vbYes Then

edituser.userId = uidtxt.Text
edituser.password = passwordtxt.Text
edituser.address = addresstxt.Text
edituser.telephone = fonetxt.Text
edituser.domainId = (didcombo.ListIndex + 1)

edituser.updatedb

Else
uidtxt.Text = edituser.userId
 passwordtxt.Text = edituser.password
 addresstxt.Text = edituser.address
 fonetxt.Text = edituser.telephone
didcombo.ListIndex = (edituser.domainId - 1)
Exit Sub
End If

End Sub
