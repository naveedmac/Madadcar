VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A6709024E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A671302B2
Private userId As String

'##ModelId=3E9A671802EE
Private terminald As String

'##ModelId=3E9A671E023A
Private loginDate As Date

'##ModelId=3E9A6734032A
Private loginTime As Date

'##ModelId=3E9A673E03D4
Private logoff_Time As Date

'##ModelId=3E9A67480136
Private logoffDate As Date

'##ModelId=3E9A6D8E014A
Private mlClassDebugID As Long

'##ModelId=3E9A6751032A
Public Sub createnewuser()
    On Error GoTo createnewuserErr

    '## Your code goes here ...

    Exit Sub
createnewuserErr:
    Call RaiseError(MyUnhandledError, "createnewuser Sub")
End Sub

'##ModelId=3E9A675901F4
Public Sub getactiveuser()
    On Error GoTo getactiveuserErr

    '## Your code goes here ...

    Exit Sub
getactiveuserErr:
    Call RaiseError(MyUnhandledError, "getactiveuser Sub")
End Sub

'##ModelId=3E9A67A300A0
Public Sub getdomain()
    On Error GoTo getdomainErr

    '## Your code goes here ...

    Exit Sub
getdomainErr:
    Call RaiseError(MyUnhandledError, "getdomain Sub")
End Sub

'##ModelId=3E9A6D8E0000
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
