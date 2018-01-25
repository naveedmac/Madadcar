VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form viewmdfrm 
   Caption         =   "Madadcar - View Car DataBase"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox regtxt 
      Height          =   285
      Left            =   7320
      TabIndex        =   8
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton search 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      DisabledPicture =   "viewmdfrm.frx":0000
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
      Left            =   8640
      MaskColor       =   &H0080FF80&
      Picture         =   "viewmdfrm.frx":2C1C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.ComboBox statuscombo 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Top             =   3000
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from car"
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
   Begin VB.ComboBox searchtype 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton LOGOUT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      DisabledPicture =   "viewmdfrm.frx":5838
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
      Left            =   6360
      MaskColor       =   &H0080FF80&
      Picture         =   "viewmdfrm.frx":8454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "viewmdfrm.frx":B070
      Height          =   3135
      Left            =   3120
      TabIndex        =   2
      Top             =   3960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5530
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
         DataField       =   "carstatus"
         Caption         =   "carstatus"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH TYPE"
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
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH TYPE"
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
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW MOTHER DATABASE "
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
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "viewmdfrm.frx":B085
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "viewmdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
searchtype.AddItem ("View All Stolen Cars")
searchtype.AddItem ("Registration No.")
searchtype.AddItem ("Maker")
searchtype.AddItem ("Status")
regtxt.Visible = False
search.Visible = False
Label2.Visible = False
statuscombo.Visible = False
statuscombo.AddItem ("Normal")
statuscombo.AddItem ("Stolen")
statuscombo.AddItem ("Recovered")

End Sub

Private Sub Form_Unload(Cancel As Integer)
'CN.Close

End Sub

Private Sub LOGOUT_Click()
Unload Me

End Sub


Private Sub search_Click()
Select Case searchtype.ListIndex
Case 0
Adodc1.RecordSource = "select * from car"
Adodc1.Refresh


Case 1
Adodc1.RecordSource = "select * from car where regno='" & regtxt.Text & "';"
Adodc1.Refresh

Case 2

Adodc1.RecordSource = "select * from car where maker='" & regtxt.Text & "';"
Adodc1.Refresh


Case 3
Adodc1.RecordSource = "select * from car where carstatus='" & statuscombo.Text & "';"
Adodc1.Refresh


End Select
End Sub

Private Sub searchtype_click()
Select Case searchtype.ListIndex
Case 0
'Adodc1.RecordSource = "select * from car where carstatus ='stolen'"
'Adodc1.Refresh
'Adodc1.RecordSource = "select * from car"
'Adodc1.Refresh

regtxt.Visible = False
search.Visible = True
Label2.Visible = False
statuscombo.Visible = False

'Label2.Caption = "Registration No."
'Image2.Visible = False

Case 1
'Adodc1.RecordSource = stolencar.getstolenCar("abc982")
'"select * from car where regno='" & regtxt.Text & "';"
'Adodc1.Refresh
regtxt.Visible = True
regtxt.Text = ""
search.Visible = True
Label2.Visible = True
Label2.Caption = "Registration No."
statuscombo.Visible = False
'Image2.Visible = True
'stolencar.flag = True

Case 2
regtxt.Visible = True
regtxt.Text = ""

search.Visible = True
Label2.Visible = True
Label2.Caption = "Maker"
statuscombo.Visible = False


'Image2.Visible = True
'stolencar.maker = regtxt.Text

'stolencar.flag = False


Case 3
regtxt.Visible = False
search.Visible = True
Label2.Visible = True
Label2.Caption = "Status"
statuscombo.Visible = True


End Select
End Sub
