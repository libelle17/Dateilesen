VERSION 5.00
Begin VB.Form DMPListenAuswahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Auswahl der DMP-Liste"
   ClientHeight    =   4170
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Dokuzahl 
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox DMPArt 
      Height          =   3570
      ItemData        =   "DMPListenAuswahl.frx":0000
      Left            =   120
      List            =   "DMPListenAuswahl.frx":0002
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label DMPArtLbl 
      Caption         =   "&DMPArt:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label DokuZahlLbl 
      Caption         =   "&Zahl der Dokumentationen zurück:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   160
      Width           =   2535
   End
End
Attribute VB_Name = "DMPListenAuswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Abbruch%

Private Sub Form_Activate()
 Abbruch = 0
 Me.DMPArt.SetFocus
End Sub ' Form_Activate()

Private Sub Form_Load()
 Dim i As DMPartEnum ' [_First] = -1 DMPDm DMPKhk DMPAb DMPCopd DMPBK DMPOsteo DMPRa DMPChi DMPRS DMPAd DMPDep [_Last]
 Me.Dokuzahl = 6
 For i = [_First] + 1 To [_Last] - 1
  Select Case i
   Case DMPDm: Me.DMPArt.AddItem "Diabetes"
   Case DMPKhk: Me.DMPArt.AddItem "KHK"
   Case DMPChi: Me.DMPArt.AddItem "Chron. Herzinsuffizienz"
   Case DMPAb: Me.DMPArt.AddItem "Asthma brochiale"
   Case DMPCopd: Me.DMPArt.AddItem "COPD"
   Case DMPOsteo: Me.DMPArt.AddItem "Osteoporose"
   Case DMPRa: Me.DMPArt.AddItem "Rheumatoide Arthritis"
   Case DMPChi: Me.DMPArt.AddItem "Chronische Herzinsuffizienz"
   Case DMPRa: Me.DMPArt.AddItem "Rückenschmerz"
   Case DMPRS: Me.DMPArt.AddItem "Adipositas"
   Case DMPDep: Me.DMPArt.AddItem "Depression"
  End Select
 Next i
End Sub ' Form_Load()

Private Sub OKButton_Click()
 Me.Hide
End Sub ' OKButton_Click()

Private Sub CancelButton_Click()
 Abbruch = True
 Me.Hide
End Sub ' CancelButton_Click()

Private Sub DMPArt_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub ' DMPArt_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub DokuZahl_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub ' DokuZahl_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub OKButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub ' OKButton_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click
  Case 27
   Call CancelButton_Click
 End Select
End Sub ' CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)
