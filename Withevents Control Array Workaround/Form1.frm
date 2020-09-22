VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "Im text #3"
      Top             =   3420
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "Im text #2"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Im text #1"
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private cCtlArr  As New cCtlArray



Private Sub Form_Initialize()
  Text4 = "This project demonstrates a workaround to the problem of not being able to use the ""Withevents"" keyword with a control array (or any array for that matter). Give each of the textboxes below focus and see what happens.  See the code comments." & vbCrLf & String(30, "¤") & vbCrLf & "www.ip-mask.com" & vbCrLf & "tools and tutorials protecting your internet privacy"
End Sub

Private Sub Form_Load()
  cCtlArr.AddTextbox Text1, "", "Im text #1"
  cCtlArr.AddTextbox Text2, "", "Im text #2"
  cCtlArr.AddTextbox Text3, "", "Im text #3"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set cCtlArr = Nothing
End Sub
