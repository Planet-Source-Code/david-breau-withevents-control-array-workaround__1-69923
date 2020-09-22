VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public WithEvents Textbox As Textbox
Attribute Textbox.VB_VarHelpID = -1
Public sTextOnFocus As String
Public sTextOnLostFocus As String
 
 
Private Sub Textbox_GotFocus()
  With Textbox
     If .Text = sTextOnLostFocus Then
        .Text = sTextOnFocus
     End If
  End With
End Sub

Private Sub Textbox_LostFocus()
  With Textbox
     If .Text = sTextOnFocus Then
       .Text = sTextOnLostFocus
     End If
  End With
End Sub
                 

