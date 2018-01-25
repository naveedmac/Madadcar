VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form viewhitlistatfrm 
   Caption         =   "Madadcar - View Hit List Form"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\stolencar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\stolencar.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton chgpassbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHANGE PASSWORD"
      DisabledPicture =   "viewhitlistatfrm.frx":0000
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
      Left            =   1680
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlistatfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      DisabledPicture =   "viewhitlistatfrm.frx":5838
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
      Left            =   7440
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlistatfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADD"
      DisabledPicture =   "viewhitlistatfrm.frx":B070
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
      Left            =   5640
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlistatfrm.frx":DC8C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRIMARY KEY"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   5040
      Picture         =   "viewhitlistatfrm.frx":108A8
      Top             =   6720
      Width           =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW HIT LIST"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "viewhitlistatfrm.frx":134C4
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "viewhitlistatfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim SQL1 As String
Dim reg As String
Dim ch As String
Dim eng As String
Dim prim As Integer
Dim maker1 As String
Dim model1 As String
Dim cstatus As String
Dim nic As String

Private Sub chgpassbtn_Click()
chgpassfrm.Show
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command3_Click()
CN.Open
SQL = "select * from car where carstatus='stolen';"
Call Recordset
'CN.Close


CN1.Open
SQL1 = "delete from stolen"
Call COMmand
CN1.Close


'CN.Open
'CN1.Open

RS.MoveFirst
While Not RS.EOF = True
'CN1.Open
reg = RS.Fields(0)
 ch = RS.Fields(1)
 eng = RS.Fields(2)
 prim = RS.Fields(3)
 maker1 = RS.Fields(4)
 model1 = RS.Fields(5)
 cstatus = RS.Fields(7)
 nic = RS.Fields(6)


SQL1 = "INSERT INTO stolen VALUES( ' " & reg & "'  , '" & ch & "' , '" & eng & "' , " & prim & " , '" & maker1 & "', " & model1 & " , '" & nic & "' , '" & cstatus & "')"
Call COMmand
Wend
'RS.RecordCount

End Sub
Public Sub COMmand()
COM1.CommandType = adCmdText
COM1.CommandText = SQL1
COM1.ActiveConnection = CN1
COM1.Execute

End Sub

Private Sub Form_Load()

'MsgBox "welcome"
'CN.Close
If valid = 1 Then
        Exit Sub
    End If
    
    If CN.Close = True Then
    'RS.Open
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
    CN.Close
    
    
    
    
    
    CN1.CursorLocation = adUseClient
    CN1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\stolencar.mdb;Persist Security Info=False"
    CN1.Open
    CN1.Close
    
    Else
    
    CN1.CursorLocation = adUseClient
    CN1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\stolencar.mdb;Persist Security Info=False"
    CN1.Open
    CN1.Close
    End If
End Sub
Private Sub Recordset()

    
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.LockType = adLockBatchOptimistic
    RS.Source = SQL
    RS.ActiveConnection = CN
    RS.Open
End Sub
Private Sub Recordset1()

    
    RS1.CursorLocation = adUseClient
    RS1.CursorType = adOpenStatic
    RS1.LockType = adLockBatchOptimistic
    RS1.Source = SQL1
    RS1.ActiveConnection = CN1
    RS1.Open
End Sub

