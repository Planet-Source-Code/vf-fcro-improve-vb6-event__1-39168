VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Callback possiblity"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Event Interface Class does Callback us for Results!!!"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Standard Event Class doesn't have CallBack possibility!"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private NewClassApproach As New MyClass2
Implements EventInterface


Private WithEvents ClassicClass As MyClass1
Attribute ClassicClass.VB_VarHelpID = -1




Private Sub ClassicClass_CallbackEvent(SomeParam As String)
'Couldn't return results!
End Sub



Private Sub Command1_Click()
ClassicClass.TestCallback
End Sub

Private Sub Command2_Click()
NewClassApproach.TestCallback
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Function EventInterface_CallBackEvent(Obj As NewApproach.MyClass2, SomeParam As String) As Long
EventInterface_CallBackEvent = 999
End Function

Private Sub EventInterface_Expose(Obj As NewApproach.MyClass2, SomeParam As String)
End Sub



Private Sub Form_Load()
Set ClassicClass = New MyClass1
Set NewClassApproach.IEvent = Me  'REQUIRED!!!!!FOR EVENT INTERFACE!!!

End Sub
