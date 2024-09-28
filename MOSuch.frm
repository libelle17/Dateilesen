VERSION 5.00
Begin VB.Form MOSuch 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox MOServer 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label MOServerCapture 
      Caption         =   "&MOServer"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "MOSuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
