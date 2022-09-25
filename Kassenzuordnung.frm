VERSION 5.00
Begin VB.Form Kassenzuordnung 
   Caption         =   "Form1"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Ok 
      Caption         =   "O&k"
      Height          =   375
      Left            =   13320
      TabIndex        =   1
      Top             =   10200
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Kassenzuordnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ok_Click()
 Unload Me
End Sub
