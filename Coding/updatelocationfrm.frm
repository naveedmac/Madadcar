VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form updatelocationfrm 
   Caption         =   "Madadcar - Update Location Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "updatelocationfrm.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   7800
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from location"
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
   Begin VB.CommandButton editbtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EDIT"
      DisabledPicture =   "updatelocationfrm.frx":1D4C44
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
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "updatelocationfrm.frx":1D7860
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1875
   End
   Begin VB.TextBox locationidtxt 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox locationtxt 
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton searchbtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEARCH"
      DisabledPicture =   "updatelocationfrm.frx":1DA47C
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
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "updatelocationfrm.frx":1DD098
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1875
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
      Left            =   6600
      Picture         =   "updatelocationfrm.frx":1DFCB4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1400
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "updatelocationfrm.frx":1E28D0
      Height          =   1575
      Left            =   2880
      TabIndex        =   3
      Top             =   6000
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2778
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "lid"
         Caption         =   "Location Id"
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
         DataField       =   "location"
         Caption         =   "Location"
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
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT LOCATION NAME :"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION ID :"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE LOCATION "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   4680
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "updatelocationfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editloc As New Location

Private Sub editbtn_Click()
editloc.locationName = locationtxt.Text
editloc.updatedb
Label1.Visible = False
locationtxt.Visible = False
editbtn.Visible = False
MsgBox "Location Editted"
Adodc1.RecordSource = "select * from location where lid=" & locationidtxt.Text
Adodc1.Refresh
Adodc1.Recordset.MoveLast


End Sub

Private Sub Form_Load()
If valid = 1 Then
        Exit Sub
    End If
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Madadcar\madadcardb.mdb;Persist Security Info=False"
    CN.Open
    valid = 1
Label1.Visible = False
locationtxt.Visible = False
editbtn.Visible = False



End Sub

Private Sub searchbtn_Click()
editloc.locationid = locationidtxt.Text
locationtxt.Text = editloc.getlocationdetails
Adodc1.RecordSource = "select * from location where lid=" & locationidtxt.Text
Adodc1.Refresh

Label1.Visible = True

locationtxt.Visible = True
editbtn.Visible = True
End Sub
