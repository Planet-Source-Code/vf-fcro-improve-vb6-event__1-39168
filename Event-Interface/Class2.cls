VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Events
Public Sub Expose(Obj As MyClass2, SomeParam As String)
End Sub
'First Parametar must be an Object who call the event!always!


Public Function CallBackEvent(Obj As MyClass2, SomeParam As String) As Long
End Function
