VERSION 5.00
Begin VB.Form FürIcon 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "FürIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Stop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 Stop
End Sub
