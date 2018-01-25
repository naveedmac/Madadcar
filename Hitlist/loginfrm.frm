VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form loginfrm 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox domainidc 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox pswrdtxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox useridtxt 
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN"
      DisabledPicture =   "loginfrm.frx":0000
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
      Left            =   5520
      MaskColor       =   &H0080FF80&
      Picture         =   "loginfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCEL"
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
      Picture         =   "loginfrm.frx":5838
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1400
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Madadcar\stolencar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Madadcar\stolencar.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERID"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DOMAIN"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   540
      Left            =   4920
      Picture         =   "loginfrm.frx":8454
      Top             =   3480
      Width           =   2160
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   4920
      Picture         =   "loginfrm.frx":B070
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   4920
      Picture         =   "loginfrm.frx":DC8C
      Top             =   5280
      Width           =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   5160
      TabIndex        =   5
      Top             =   1560
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Picture         =   "loginfrm.frx":108A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim login As New Users
Dim var2
'Dim alldomain As New Domain
Dim SQL As String



Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Madadcar\stolencar.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
    
domainidc.AddItem ("admin")
domainidc.AddItem ("suser")
domainidc.AddItem ("auser")

'*********************************

'SQL = alldomain.getdomain
'MsgBox "hi"
'Call Recordset
'RS.MoveFirst
'While RS.EOF = False
'domainidc.AddItem (RS(1).Value)
'prodnamecombo.AddItem (RS(1).Value)
'domainidc.ItemData(domainidc.NewIndex) = RS(0).Value
'RS.MoveNext
'Wend
'RS.Close


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

Private Sub OK_Click()
login.userId = useridtxt.Text
login.password = pswrdtxt.Text
login.domainId = domainidc.ItemData(domainidc.ListIndex)
login.domainId = login.domainId + 1

'MsgBox domainidc.ListIndex(2)

MsgBox login.verifyuser()
'RS.Close
'MsgBox login.domainId
If login.domainId = 1 Then
domainstat = 1
Else
domainstat = 2
End If
MsgBox domainstat
'Unload Me
'CN.Close
'viewhitlistfrm.Show



End Sub

