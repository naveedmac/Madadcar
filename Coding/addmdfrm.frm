VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addmdfrm 
   Caption         =   "Madadcar -Add Car To Database Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "addmdfrm.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   7320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox devicetxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton addbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADD"
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
      Left            =   5280
      Picture         =   "addmdfrm.frx":28B3E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   1400
   End
   Begin VB.TextBox nictxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   6480
      Width           =   3015
   End
   Begin VB.TextBox makertxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton okbtn 
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
      Left            =   7560
      Picture         =   "addmdfrm.frx":2B75A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Width           =   1400
   End
   Begin VB.TextBox modeltxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox enginetxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox chasistxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox regtxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CHASIS No :"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DEVICE No. :"
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
      Left            =   4800
      TabIndex        =   14
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MODEL No. :"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MAKER :"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIC OF OWNER :"
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
      TabIndex        =   10
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENGINE No :"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "REG No. :"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD CAR "
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
      Left            =   5400
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -240
      Picture         =   "addmdfrm.frx":2E376
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "addmdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newcar As New car


Private Sub addbtn_Click()

If regtxt.Text = "" Or enginetxt.Text = "" Or chasistxt.Text = "" Or modeltxt.Text = "" Or makertxt.Text = "" Or nictxt.Text = "" Then
MsgBox "Input All the Values..!!", vbExclamation, Error
'pswrdtxt.SetFocus
SendKeys "{Home}+{End}"
Exit Sub
End If

devicetxt.Text = newcar.addnewcar(enginetxt.Text, chasistxt.Text, makertxt.Text, modeltxt.Text, regtxt.Text, nictxt.Text)

'RS.Close


End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
CN.Close

End Sub

Private Sub modeltxt_Change()
If KeyAscii = 8 Then
KeyAscii = 8
Else
If KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 8 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub okbtn_Click()
Unload Me

End Sub
