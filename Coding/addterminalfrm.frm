VERSION 5.00
Begin VB.Form addterminalfrm 
   Caption         =   "Madadcar - Add Terminal Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox terminalid 
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox terminalname 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton okbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      DisabledPicture =   "addterminalfrm.frx":0000
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
      Left            =   7320
      MaskColor       =   &H0080FF80&
      Picture         =   "addterminalfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton addbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADD"
      DisabledPicture =   "addterminalfrm.frx":5838
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
      Left            =   5280
      MaskColor       =   &H0080FF80&
      Picture         =   "addterminalfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.ComboBox areacombo 
      Height          =   315
      Left            =   7440
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ComboBox locationcombo 
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TERMINAL ID :"
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
      TabIndex        =   8
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TERMINAL NAME :"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AREA :"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION :"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD TERMINAL"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -120
      Picture         =   "addterminalfrm.frx":B070
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "addterminalfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SQL As String
Dim viewallLoc As New Location


Private Sub Command1_Click()
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

'**********************************
SQL = viewallLoc.getallLocations


'MsgBox "hi"
Call Recordset
RS.MoveFirst
While RS.EOF = False
locationcombo.AddItem (RS(1).Value)
'prodnamecombo.AddItem (RS(1).Value)
locationcombo.ItemData(locationcombo.NewIndex) = RS(0).Value
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
Public Sub COMmand()
COM.CommandType = adCmdText
COM.CommandText = SQL
COM.ActiveConnection = CN
COM.Execute
End Sub
