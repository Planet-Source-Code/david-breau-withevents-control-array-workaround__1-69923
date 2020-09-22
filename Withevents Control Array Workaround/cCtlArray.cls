VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCtlArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private cEH() As cEventHandler

Dim bInitialized As Boolean

'----------------------------------------------------
'the key to get around using withevents with a control
'array is create a class (cEventHandler) that has a
'withevents declaration for just one object (in this
'case a textbox) then place code in the desired events.
'Then you create an array of this class (the front end class)
'so in essence you have created a withevents scenario with
'an array.
'In this example I wanted to have textboxes to have the
'behavior of having certain text when it has the focus
'and different text when it loses focus.
'----------------------------------------------------

Public Sub AddTextbox(T As Textbox, _
                       sTextOnFocus As String, _
                       sTextOnLostFocus As String)
 '
 'the purpose of this code is avoid the error
 'that would occur when attempting to get the
 'ubound of an array if it hasnt yet been initialized
 If bInitialized Then
    Dim i As Integer
    i = UBound(cEH) + 1
    
 Else
    bInitialized = True
    i = 0
 End If '---------------------------------------------
 
 '
 'create another instance of the class
 ReDim Preserve cEH(i)
 Set cEH(i) = New cEventHandler
 '
 'set the control reference
 Set cEH(i).Textbox = T
 '
 'set the "code" for the events
 cEH(i).sTextOnFocus = sTextOnFocus
 cEH(i).sTextOnLostFocus = sTextOnLostFocus
End Sub
 
Private Sub Class_Terminate()
Dim i As Integer, icnt As Integer
  
  If bInitialized Then
    icnt = UBound(cEH)
    '
    'destroy reference to the array of classes
    'as well as the control in each array
    For i = 0 To icnt
      Set cEH(i).Textbox = Nothing
      Set cEH(i) = Nothing
    Next i
  End If
End Sub
