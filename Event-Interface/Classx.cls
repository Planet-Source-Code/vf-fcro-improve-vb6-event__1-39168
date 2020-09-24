VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyClass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public IEvent As EventInterface



Sub DoExpose()
'Raise Event
IEvent.Expose Me, "The New Approach!"
'First Parametar must be an Object who call the Event!!!!!!!!
End Sub


Sub TestCallback()
Dim ret As Long
'Raise Event
ret = IEvent.CallBackEvent(Me, "Test")
MsgBox "Return Result is " & ret, , "Info"
End Sub
