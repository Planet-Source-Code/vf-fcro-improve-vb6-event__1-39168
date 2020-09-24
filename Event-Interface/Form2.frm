VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "REDIM & TEST EVENT"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   ScaleHeight     =   2595
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Class with Event Interface can be bounded and properly Evented"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Standard Class With Event can be bounded,but cannot be properly Evented"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Easy Redim And Impossible To Expose Event------>>>>>Class With Standard Event"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Intermediate Redim And Easy To Expose Event ------->>>>>Class With Event Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
Private NewClassApproach() As New MyClass2
Implements EventInterface
'CAN BE BOUNDED AND PROPERLY EXPOSED!!!!!
'YES...
'*******************************************


'************************************************
Private WithEvents ClassicClass As MyClass1
Attribute ClassicClass.VB_VarHelpID = -1
'CAN BE BOUNDED AND UNPROPERLY EXPOSED!!!!
'Something Like This:
'Private WithEvents ClassicClass() As MyClass1
'************************************************

Private Function EventInterface_CallBackEvent(Obj As NewApproach.MyClass2, SomeParam As String) As Long
'Must be implemented
End Function

Private Sub EventInterface_Expose(Obj As NewApproach.MyClass2, SomeParam As String)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'OBJ is an Object who called Event!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
MsgBox "Class PTR:" & ObjPtr(Obj), , "Event Expose"
End Sub
Private Sub ClassicClass_Expose(SomeParam As String)
MsgBox "Godot Event"
'Never Exposed Event!!!!!!
'VB wouldn't call this Event ever!!!!!
End Sub

Private Sub Command1_Click()
'There is no way to redim class with event on properly way!

ReDim ClassicClass(4)
For u = 0 To UBound(ClassicClass)
Set ClassicClass(u) = New MyClass1
Next u

MsgBox "Number Of Standard Event Classes:" & UBound(ClassicClass) + 1

For u = 0 To UBound(ClassicClass)
ClassicClass(u).DoExpose
'We Call our test,but nothings happen on Event!
Next u
End Sub

Private Sub Command2_Click()
ReDim NewClassApproach(4)
'***********REUQIRED TO SET UP EVENT INTERFACE WITH FORM!
For u = 0 To UBound(NewClassApproach)
Set NewClassApproach(u) = New MyClass2
Set NewClassApproach(u).IEvent = Me
Next u
'*********************************************************

MsgBox "Number Of Event Interface Classes:" & UBound(NewClassApproach) + 1
For u = 0 To UBound(NewClassApproach)
NewClassApproach(u).DoExpose 'Call To Test Event
Next u
End Sub

Private Sub Command3_Click()
Unload Me
End Sub



