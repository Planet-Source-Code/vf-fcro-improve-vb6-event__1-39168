VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyClass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Expose(SomeParam As String)
Event CallbackEvent(SomeParam As String)
'Standard Event haven't callback possibility!


Sub DoExpose()
RaiseEvent Expose("Standard Approach!")
End Sub


Sub TestCallback()
RaiseEvent CallbackEvent("Test")
'We dont receive any callback stuff..Like Return results!
MsgBox "We don't have any callback results! Huh..", , "Info"
End Sub


