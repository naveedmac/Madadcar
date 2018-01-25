VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Terminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A67AC01B8"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A67B4019A
Public terminalId As String

'##ModelId=3E9A67BB00DC
Public terminalName As String

'##ModelId=3E9A6CA500E6
Public areaId As String
Public terminalstatus As String
'##ModelId=3E9A6D90024E
Private mlClassDebugID As Long

'##ModelId=3E9A69D6024E
Public NewProperty As Collection

'##ModelId=3E9A6D9001E0

Public Function addterminal() As Integer
    Dim max
    'Dim state
    SQL = "SELECT MAX(tid) FROM terminal;"
    Call Recordset
    max = RS.Fields(0)
    terminalId = max + 1
    RS.Close
    SQL = "INSERT INTO terminal VALUES(  " & terminalId & "  , '" & terminalName & "')"
    Call COMmand
    MsgBox "Terminal Added"
    addlocation = terminalId
End Function
Public Function getlocationdetails() As String

    SQL = "select * from location where tid=" & terminalId & ";"
    Call Recordset
    terminalName = RS.Fields(1)
    getlocationdetails = terminalName
End Function
Public Function getallLocations() As String
    getallLocations = "select * from terminal"
End Function
Public Sub updatedb()
    SQL = "update terminal set tname='" & terminalName & ",aid=" & areaId & ",tstatus='" & terminalstatus & "' where tid='" & terminalId & "';"
    'SQL = "update car set nico='" & NICofOwner & "',carstatus='" & carStatus & "' where regno='" & registrationNo & "';"

    Call COMmand
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



Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
