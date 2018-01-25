VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Domain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A66170136"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A661D015E
Public domainId As String

'##ModelId=3E9A662802D0
Private domainName As String

'##ModelId=3E9A6D8B00C8
Private mlClassDebugID As Long

'##ModelId=3E9A696F02F8
Public NewProperty As Collection

'##ModelId=3E9A662F00C8
Public Function getdomain() As String

    On Error GoTo getdomainErr
getdomain = "select * from domainn"

    
    '## Your code goes here ...

    Exit Function
getdomainErr:
    Call RaiseError(MyUnhandledError, "getdomain Sub")
End Function

'##ModelId=3E9A6D8B005A
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
