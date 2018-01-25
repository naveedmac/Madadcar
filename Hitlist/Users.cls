VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A66B7005A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit
Private SQL As String

'##ModelId=3E9A66C00398
Public userId As String

'##ModelId=3E9A66C602B2
Public userName As String

'##ModelId=3E9A66CA03CA
Public NICofuser As String

'##ModelId=3E9A66D10190
Public telephone As Integer

'##ModelId=3E9A66D501CC
Public address As String

'##ModelId=3E9A66D90136
Public password As String

'##ModelId=3E9A66DF02DA
Public domainId As String

'##ModelId=3E9A6D8C0276
Private mlClassDebugID As Long
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
Public Sub showform()
     
     If domainId = 1 Then
     adminmainfrm.Show
     End If
     
     If domainId = 2 Then
     superusermainfrm.Show
     End If
     
     If domainId = 3 Then
    viewhitlistatfrm.Show
    
     End If
     
     
End Sub

'##ModelId=3E9A66ED0276
Public Function verifyuser() As Boolean

    'MsgBox name
    'On Error GoTo verifyuserErr

    '## Your code goes here ...
     'domainId = id
     'userName = name
     'password = pass
     
     'SQL = "SELECT * FROM user;"
     MsgBox domainId
     SQL = "select * from userst where uid= " & userId & "  and pasword= '" & password & "' and did= " & domainId & " ;"
     Call Recordset
     
     If RS.RecordCount <> 0 Then
     MsgBox password
     MsgBox "correct pass"
     CN.Close
     
     viewhitlistfrm.Show
   ' Call showform
     verifyuser = True
     
     Else
    MsgBox "wrong password"
         verifyuser = False
         
    'If RS.Fields(4) = pass Then
    
   'superusermainfrm.Show
   
    'Else
    'MsgBox "wrong password"
    End If
    
    'RS.Close
    
    
    
 '   Exit Function
'verifyuserErr:
 '   Call RaiseError(MyUnhandledError, "verifyuser Sub")
End Function

'##ModelId=3E9A66F402D0
Public Sub pop_userattrib(uid As Integer, name As String, nicu As String, tele As Integer, address As String, pass As String, id As Integer)
domainId = id
userName = name
password = pass
userId = uid
telephone = tele
address = address
NICofuser = nicu
End Sub

Public Sub getuser()
    On Error GoTo getuserErr
MsgBox "hi1"
    '## Your code goes here ...
 SQL = "select * from userst where uid=" & userId & ";"
  Call Recordset
  MsgBox "hi"
  Call pop_userattrib(RS.Fields(0), RS.Fields(1), RS.Fields(2), RS.Fields(3), RS.Fields(4), RS.Fields(5), RS.Fields(6))
  

    Exit Sub
getuserErr:
    Call RaiseError(MyUnhandledError, "getuser Sub")
End Sub
'##ModelId=3E9A66F402D0
Public Function adduser() As Integer

Dim max As Integer


    On Error GoTo adduserErr

    '## Your code goes here ...
    SQL = "SELECT MAX(uid) FROM userst;"
    Call Recordset
    
    'MsgBox RS.Fields(0)
    max = RS.Fields(0)
    userId = max + 1
    '
    'userName = name
    'password = pass
    'NICofuser = nic
    'address = add
   ' domainId = id
    'Dim ab As String
     'ab = "d1"
   SQL = "INSERT INTO userst VALUES(  " & userId & "  , '" & userName & "' , '" & NICofuser & "' , " & telephone & " , '" & address & "', '" & password & "'," & domainId & ")"
    'RS.Close
    
    'SQL = "select * from car"
    Call COMmand
    MsgBox " User Added"
    adduser = userId
    RS.Close
    Exit Function
adduserErr:
    Call RaiseError(MyUnhandledError, "adduser Sub")
End Function

'##ModelId=3E9A66FB01D6
Public Sub changepasswsord()
    On Error GoTo changepasswsordErr

    '## Your code goes here ...

    Exit Sub
changepasswsordErr:
    Call RaiseError(MyUnhandledError, "changepasswsord Sub")
End Sub

'##ModelId=3E9A6D8C0208
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
