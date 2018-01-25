VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Owner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A68CE0078"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A68D301C2
Public name As String

'##ModelId=3E9A68D60276
Public NICofOwner As String

'##ModelId=3E9A68DF0172
Public phoneofOwner As Integer
Private SQL As String

'##ModelId=3E9A68EB00AA
Public address As String

'##ModelId=3E9A6D9203B6
Private mlClassDebugID As Long
Private flag As Boolean

'##ModelId=3E9A69A3030C
Public NewProperty As Collection
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
'##ModelId=3E9A690B035C
Public Function checkOwner() As Boolean

    On Error GoTo checkOwnerErr

    '## Your code goes here ...

SQL = "select * from owner where nico= '" & NICofOwner & "';"
     Call Recordset
     
     If RS.RecordCount <> 0 Then
     RS.Close
     checkOwner = True
     Else
     RS.Close
     
     checkOwner = False
     

End If
    
    Exit Function
checkOwnerErr:
    Call RaiseError(MyUnhandledError, "checkOwner Sub")
End Function

'##ModelId=3E9A691202D0
Public Sub createnewOwner()
    On Error GoTo createnewOwnerErr

    '## Your code goes here ...

    Exit Sub
createnewOwnerErr:
    Call RaiseError(MyUnhandledError, "createnewOwner Sub")
End Sub

'##ModelId=3E9A6D920348
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property

