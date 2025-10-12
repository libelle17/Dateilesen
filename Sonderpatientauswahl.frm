VERSION 5.00
Begin VB.Form Sonderpatientauswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sonderpatient-Auswahl"
   ClientHeight    =   8910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   8640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Sonderpatientauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Abbruch%
Public aufRufer As Lese

Private Sub Form_Activate()
 Abbruch = 0
 Me.List1.SetFocus
End Sub ' Form_Activate()

Private Sub Form_Load()
End Sub  ' Form_Load

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub ' DMPArt_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub OKButton_Click()
 Me.Visible = False
End Sub ' OKButton_Click()

Private Sub CancelButton_Click()
 Abbruch = True
 Me.Visible = False
End Sub ' CancelButton_Click()


