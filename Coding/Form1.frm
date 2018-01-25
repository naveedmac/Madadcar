VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form loginfrm 
   BackColor       =   &H8000000E&
   Caption         =   "Madadcar - Login Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "domainn"
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
      Picture         =   "Form1.frx":1D4C44
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1400
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN"
      DisabledPicture =   "Form1.frx":1D7860
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
      Picture         =   "Form1.frx":1DA47C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.TextBox pswrdtxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox useridtxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7440
      TabIndex        =   1
      Top             =   4440
      Width           =   3135
   End
   Begin VB.ComboBox domainidc 
      Height          =   315
      ItemData        =   "Form1.frx":1DD098
      Left            =   7440
      List            =   "Form1.frx":1DD09A
      TabIndex        =   0
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   48
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   5160
      TabIndex        =   8
      Top             =   1560
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim login As New Users
Dim var2
Dim alldomain As New Domain
Dim SQL As String

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1


'*********************************

SQL = alldomain.getdomain
'MsgBox "hi"
Call Recordset
RS.MoveFirst
While RS.EOF = False
domainidc.AddItem (RS(1).Value)
'prodnamecombo.AddItem (RS(1).Value)
domainidc.ItemData(domainidc.NewIndex) = RS(0).Value
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
'Unload Me
End Sub

Private Sub OK_Click()

If pswrdtxt.Text = "" Then
MsgBox "Input Password..!!", vbExclamation, Error
pswrdtxt.SetFocus
SendKeys "{Home}+{End}"
Exit Sub
End If

'If domainidc.ItemData(domainidc.ListIndex) = "" Then
'MsgBox "Select Domain..!!", vbDefaultButton3, Error
'End If

If useridtxt.Text = "" Then
MsgBox "Input UserID..!!", vbExclamation, Error
useridtxt.SetFocus
SendKeys "{Home}+{End}"
Exit Sub
End If

login.userId = useridtxt.Text
login.password = pswrdtxt.Text
login.domainId = domainidc.ItemData(domainidc.ListIndex)


Dim var
 
var = login.verifyuser()
If var = 0 Then
useridtxt.SetFocus
SendKeys "{Home}+{End}"
    'CN.Close
Exit Sub
End If
 
'RS.Close
'MsgBox login.domainId
If login.domainId = 1 Then
domainstat = 1

Else
domainstat = 2

End If
'MsgBox domainstat
'Unload Me


CN.Close
Unload Me

End Sub
