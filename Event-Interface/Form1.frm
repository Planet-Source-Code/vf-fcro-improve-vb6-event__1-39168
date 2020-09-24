VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tests:Author Vanja Fuckar,EMAIL:INGA@VIP.HR"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TEST # 3 (Callback Test!)"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TEST # 2 (Boundary Test!)"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calling Speed Test :Class Event Interface Approach"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Calling Speed Test :Standard Class Event Approach"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Two Different Approach to the Event!!!!!!!BY VANJA FUCKAR,EMAIL INGA@VIP.HR

'TESTING SPEED!!!!!!!!!!!


'******************STANDARD APPROACH**************
Private WithEvents ClassicClass As MyClass1
Attribute ClassicClass.VB_VarHelpID = -1
'*************************************************

'****************NEW APPROACH*********************
Private NewClassApproach As New MyClass2
Implements EventInterface
'*************************************************



Dim Counter As Long

Private Sub Command3_Click()
Form2.Show 1
End Sub

Private Sub Command4_Click()
Form3.Show 1
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Function EventInterface_CallBackEvent(Obj As NewApproach.MyClass2, SomeParam As String) As Long
'Must be implemented
End Function

Private Sub EventInterface_Expose(Obj As NewApproach.MyClass2, SomeParam As String)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'OBJ is an Object who called Event!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Test Calling Time
End Sub

Private Sub ClassicClass_Expose(SomeParam As String)
'Test Calling Time
End Sub
'Look at differences beetween EventInterface and Classic Event!



Private Sub Command1_Click()
'TEST 1

Dim TC As Long
TC = GetTickCount
For u = 0 To 100000
ClassicClass.DoExpose
Next u
MsgBox "Execute Time With Classic Event Method:" & GetTickCount - TC & " Millisec.", , "Test"
End Sub

Private Sub Command2_Click()
'TEST 2

Dim TC As Long
TC = GetTickCount
For u = 0 To 100000
NewClassApproach.DoExpose
Next u
MsgBox "Execute Time With Event Interface Method:" & GetTickCount - TC & " Millisec.", , "Test"
End Sub





Private Sub Form_Load()
MsgBox "This Tests shows you 2 way to expose Event:" & vbCrLf & "Classic Event vs Event Interface" & vbCrLf & _
"Who's 'Better For You' decide on your own.Before you choose Classic Way,consider the tests results...", , "Info"
Set ClassicClass = New MyClass1
Set NewClassApproach.IEvent = Me 'REQUIRED!!!!!FOR EVENT INTERFACE!!!
End Sub


Private Sub Form_Unload(Cancel As Integer)
MsgBox "Finally,What we've got?? Class With Event Interface have this new possiblity:" & vbCrLf & _
"1:Faster Calling Convention" & vbCrLf & "2:Event Boundary Ability" & vbCrLf & "3:Callback possiblity", , "the Conclusion!"
End Sub
