VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A692E015E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A6952021C
Public locationid As String

'##ModelId=3E9A695503B6
Public locationName As String

'##ModelId=3E9A6D920046
Private mlClassDebugID As Long
Private SQL As String

'##ModelId=3E9A6A310032
Public NewProperty As Collection

'##ModelId=3E9A6D9103C0
Public Function addlocation() As Integer
    Dim max
    max = 0
    'Dim state
    SQL = "SELECT MAX(lid) FROM location;"
    Call Recordset
    max = RS.Fields(0)
    locationid = max + 1
    RS.Close
    SQL = "INSERT INTO location VALUES(  " & locationid & "  , '" & locationName & "')"
    Call COMmand
    MsgBox "location Added"
    addlocation = locationid
End Function
Public Function getlocationdetails() As String

    SQL = "select * from location where lid=" & locationid & ";"
    Call Recordset
    locationName = RS.Fields(1)
    getlocationdetails = locationName
End Function
Public Function getallLocations() As String
    getallLocations = "select * from location"
End Function
Public Sub updatedb()
    SQL = "update location set location='" & locationName & "' where lid=" & locationid & ";"
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
