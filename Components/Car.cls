VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A67D6023A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=3E9A67DD0334
Private primaryKey As Integer

'##ModelId=3E9A67E600DC
Private engineNo As String

'##ModelId=3E9A67EC028A
Private chasisNo As String

'##ModelId=3E9A67F40154
Public maker As String

'##ModelId=3E9A67FC019A
Private modelNo As Integer

'##ModelId=3E9A68590082
Private registrationNo As String

'##ModelId=3E9A685F0262
Public NICofOwner As String
Public flag As Boolean

'##ModelId=3E9A6867012C
Public carStatus As String
Private SQL As String

'##ModelId=3E9A6D94017C
Private mlClassDebugID As Long


Private Function pop_attrib(eng As String, chas As String, make As String, model As Integer, regno As String, nic As String, carstat As String)
engineNo = eng
chasisNo = chas
modelNo = model
maker = make
registrationNo = regno
NICofOwner = nic
carStatus = carstat
End Function

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


'##ModelId=3E9A687500FA
Public Function addnewcar(eng As String, chas As String, make As String, model As Integer, regno As String, nic As String) As Integer

    'On Error GoTo addnewcarErr
    Dim max
    Dim state

    '## Your code goes here ...
    Call pop_attrib(eng, chas, make, model, regno, nic, "normal")
    
    SQL = "SELECT MAX(pkey) FROM car;"
    Call Recordset
    
    'MsgBox RS.Fields(0)
    max = RS.Fields(0)
    primaryKey = max + 1
    state = "normal"
    RS.Close
    SQL = "INSERT INTO car VALUES( ' " & registrationNo & "'  , '" & chasisNo & "' , '" & engineNo & "' , " & primaryKey & " , '" & maker & "', " & modelNo & " , '" & NICofOwner & "' , '" & state & "')"
    'RS.Close
    
    'SQL = "select * from car"
    
    Call COMmand
    'RS.MoveLast
    
    'RS.AddNew
    'RS.Fields(0) = registrationNo
    'RS.Fields(1) = chasisNo
    'RS.Fields(2) = engineNo
    'RS.Fields(3) = primaryKey
    'RS.Fields(4) = maker
    'RS.Fields(5) = modelNo
    'RS.Fields(6) = NICofOwner
    'RS.Fields(7) = "normal"
    'RS.Update
'Call COMmand

    MsgBox "Car added."
    addnewcar = primaryKey
    
    Exit Function
addnewcarErr:
    MsgBox "error"
    Call RaiseError(MyUnhandledError, "addnewcar Sub")
End Function

'##ModelId=3E9A688A0154
Public Function getcardetails(reg As String) As String

    On Error GoTo getcardetailsErr

    '## Your code goes here ...
  SQL = "select * from car where regno='" & reg & "';"
  Call Recordset
  If RS.RecordCount = 0 Then
  MsgBox "Your Car does not exist.!!"
  RS.Close
  getcardetails = ""
  
  ' Exit Function
  Else
       
  Call pop_attrib(RS.Fields(2), RS.Fields(1), RS.Fields(4), RS.Fields(5), RS.Fields(0), RS.Fields(6), RS.Fields(7))
  
  getcardetails = "select * from car where regno='" & reg & "';"
  RS.Close
  End If
    
    Exit Function
getcardetailsErr:
    Call RaiseError(MyUnhandledError, "getcardetails Sub")
End Function

Public Sub updatedb()
SQL = "update car set nico='" & NICofOwner & "',carstatus='" & carStatus & "' where regno='" & registrationNo & "';"
Call COMmand

End Sub
'##ModelId=3E9A688F01C2
Public Sub makeChanges()
    On Error GoTo makeChangesErr

    '## Your code goes here ...

    Exit Sub
makeChangesErr:
    Call RaiseError(MyUnhandledError, "makeChanges Sub")
End Sub

'##ModelId=3E9A68970208
Public Function getstolenCar(reg As String) As String

    On Error GoTo getstolenCarErr
If flag = False Then
getstolenCar = "select * from car where maker='" & reg & "' and carstatus='stolen';"



Else
getstolenCar = "select * from car where carstatus='stolen' and regno='" & reg & "'"

End If
    '## Your code goes here ...

    Exit Function
getstolenCarErr:
    Call RaiseError(MyUnhandledError, "getstolenCar Sub")
End Function

'##ModelId=3E9A68B2014A
Public Sub addnewstolenCar()
    On Error GoTo addnewstolenCarErr

    '## Your code goes here ...

    Exit Sub
addnewstolenCarErr:
    Call RaiseError(MyUnhandledError, "addnewstolenCar Sub")
End Sub

'##ModelId=3E9A68BA02A8
Public Sub checkStatus()
    On Error GoTo checkStatusErr

    '## Your code goes here ...

    Exit Sub
checkStatusErr:
    Call RaiseError(MyUnhandledError, "checkStatus Sub")
End Sub

'##ModelId=3E9A68C3012C
Public Sub changeStatus()
    On Error GoTo changeStatusErr

    '## Your code goes here ...

    Exit Sub
changeStatusErr:
    Call RaiseError(MyUnhandledError, "changeStatus Sub")
End Sub

'##ModelId=3E9A6D94010E
Public Property Get ClassDebugID() As Variant
    On Error GoTo ClassDebugIDErr

    ClassDebugID = mlClassDebugID

    Exit Property
ClassDebugIDErr:
    Call RaiseError(MyUnhandledError, "ClassDebugID Property")
End Property
