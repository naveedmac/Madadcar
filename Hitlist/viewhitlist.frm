VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form viewhitlistfrm 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update"
      DisabledPicture =   "viewhitlist.frx":0000
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
      Left            =   4080
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlist.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   600
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   6600
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
      CommandType     =   1
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
      RecordSource    =   "select * from stolen;"
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADD"
      DisabledPicture =   "viewhitlist.frx":5838
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
      Left            =   5760
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlist.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      DisabledPicture =   "viewhitlist.frx":B070
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
      Left            =   7560
      MaskColor       =   &H0080FF80&
      Picture         =   "viewhitlist.frx":DC8C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1400
   End
   Begin VB.CommandButton chgpassbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CHANGE PASSWORD"
      DisabledPicture =   "viewhitlist.frx":108A8
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
      Picture         =   "viewhitlist.frx":134C4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   2235
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "viewhitlist.frx":160E0
      Height          =   4095
      Left            =   2280
      TabIndex        =   0
      Top             =   2280
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "regno"
         Caption         =   "regno"
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
         DataField       =   "chno"
         Caption         =   "chno"
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
      BeginProperty Column02 
         DataField       =   "engno"
         Caption         =   "engno"
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
      BeginProperty Column03 
         DataField       =   "pkey"
         Caption         =   "pkey"
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
      BeginProperty Column04 
         DataField       =   "maker"
         Caption         =   "maker"
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
      BeginProperty Column05 
         DataField       =   "model"
         Caption         =   "model"
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
      BeginProperty Column06 
         DataField       =   "nico"
         Caption         =   "nico"
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
      BeginProperty Column07 
         DataField       =   "carstattus"
         Caption         =   "carstattus"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   600
      Width           =   2175
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
      Left            =   3240
      TabIndex        =   6
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   3240
      Picture         =   "viewhitlist.frx":160F5
      Top             =   6720
      Width           =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW HIT LIST"
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
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "viewhitlist.frx":18D11
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "viewhitlistfrm"
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


Private Sub Command3_Click()
'CN.Open
SQL = "select * from car where carstatus='stolen';"
Call Recordset
'CN.Close


'CN1.Open
SQL1 = "delete from stolen"
Call COMmand
'CN1.Close


'CN.Open
'CN1.Open

RS.MoveFirst
While Not RS.EOF = True
'CN1.Open
'reg = RS.Fields(0)
 'ch = RS.Fields(1)
 'eng = RS.Fields(2)
 'prim = RS.Fields(3)
 'maker1 = RS.Fields(4)
 'model1 = RS.Fields(5)
 'cstatus = RS.Fields(7)
 'nic = RS.Fields(6)


SQL1 = "INSERT INTO stolen VALUES( ' " & RS.Fields(0) & "'  , '" & RS.Fields(1) & "' , '" & RS.Fields(2) & "' , " & RS.Fields(3) & " , '" & RS.Fields(4) & "', " & RS.Fields(5) & " , '" & RS.Fields(6) & "' , '" & RS.Fields(7) & "')"
Call COMmand
RS.MoveNext

Wend
'RS.RecordCount

RS.Close
Adodc1.Refresh

'RS1.Close
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
    
    
    
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
    
    
    
    
    
    
    CN1.CursorLocation = adUseClient
    CN1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\stolencar.mdb;Persist Security Info=False"
    CN1.Open
    'CN1.Close
    
    
    
    
    
'CN.Open
SQL = "select * from car where carstatus='stolen';"
Call Recordset
'CN.Close


'CN1.Open
SQL1 = "delete from stolen"
Call COMmand
'CN1.Close


'CN.Open
'CN1.Open

RS.MoveFirst
While Not RS.EOF = True
'CN1.Open
'reg = RS.Fields(0)
 'ch = RS.Fields(1)
 'eng = RS.Fields(2)
 'prim = RS.Fields(3)
 'maker1 = RS.Fields(4)
 'model1 = RS.Fields(5)
 'cstatus = RS.Fields(7)
 'nic = RS.Fields(6)


SQL1 = "INSERT INTO stolen VALUES( ' " & RS.Fields(0) & "'  , '" & RS.Fields(1) & "' , '" & RS.Fields(2) & "' , " & RS.Fields(3) & " , '" & RS.Fields(4) & "', " & RS.Fields(5) & " , '" & RS.Fields(6) & "' , '" & RS.Fields(7) & "')"
Call COMmand
RS.MoveNext

Wend
'RS.RecordCount

RS.Close
'RS1.Close
Adodc1.Refresh


    
    End Sub
    
    
    'RS.Close
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


Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub
