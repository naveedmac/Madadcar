VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A691E01E0"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A692301A4
Public areaId As Integer

'##ModelId=3E9A69260258
Public areaName As String

Public locationid As Integer

'##ModelId=3E9A6D8F02C6
Private mlClassDebugID As Long
Private SQL As String

'##ModelId=3E9A6A1D0370
Public NewProperty As Collection



Public Function addarea() As Integer
    Dim max
    'Dim state
    SQL = "SELECT MAX(aid) FROM area;"
    Call Recordset
    max = RS.Fields(0)
    areaId = max + 1
    RS.Close
    SQL = "INSERT INTO area VALUES(  " & areaId & "  , '" & areaName & "','" & locationid & "')"
    Call COMmand
    MsgBox "area Added"
    addarea = areaId
End Function
Public Function getareadetails() As String

    SQL = "select * from area where aid=" & areaId & ";"
    Call Recordset
    areaName = RS.Fields(1)
    locationid = RS.Fields(2)
    
    getareadetails = areaName


End Function
Public Function getallarea() As String
    getallarea = "select * from area;"
End Function
Public Sub updatedb()
    SQL = "update area set aname='" & areaName & "', lid=" & locationid & " where aid=" & areaId & ";"
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
'##ModelId=3E9A6D8F0258
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
