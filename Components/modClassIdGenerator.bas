Attribute VB_Name = "modClassIdGenerator"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E9A6D890188"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Debug.ClassIdGenerator"
Option Explicit

'Class ID generator
'##ModelId=3E9A6D96015E
Public Function GetNextClassDebugID() As Long
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
End Function
