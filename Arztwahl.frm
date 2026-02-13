VERSION 5.00
Begin VB.Form Arztwahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Arztauswahl"
   ClientHeight    =   8655
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "&Nuber"
      Height          =   375
      Index           =   7
      Left            =   5520
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "K&reis"
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Hoffmann"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Fuchs"
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Bender"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Index           =   3
      Left            =   11760
      TabIndex        =   12
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "n&ä Mitarbeiter"
      Height          =   375
      Index           =   2
      Left            =   10560
      TabIndex        =   11
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Überspringen"
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   10
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Schade"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Kothny"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Zurück"
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   8
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Hammerschmidt"
      Height          =   375
      Index           =   0
      Left            =   -120
      TabIndex        =   0
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Texte 
      Caption         =   "Inhalt"
      Height          =   8175
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "Arztwahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Abbruch%
Public Arzt%

Private Sub Form_Activate()
 Me.OKButton(Arzt).SetFocus
End Sub ' Form_Activate()

Private Sub Form_Load()
 Abbruch = 0
 Me.WindowState = vbMaximized
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then
 DBCn.Close
 DBCn.Open
 Resume
End If ' Err.Number = -2147467259 Then
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Load()

Private Sub CancelButton_Click(Index%)
 Abbruch = IIf(Index, Index, -1)
 Me.Hide
End Sub ' CancelButton_Click()

Private Sub CancelButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click(KeyCode)
  Case 27
   Call CancelButton_Click(3)
 End Select
End Sub ' CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Form_Resize()
 Dim i%, top&, left&
 Const KW% = 1300
 top = Me.Height - 9090 + 8160 + 100
 left = Me.Width - (13455 - 8040) - Me.OKButton.COUNT * (KW + 10)
 For i = 0 To Me.OKButton.COUNT - 1
  Me.OKButton(i).top = top
  Me.OKButton(i).left = left + i * (KW + 10)
  Me.OKButton(i).Height = 375
  Me.OKButton(i).Width = KW
 Next i
 For i = 0 To Me.CancelButton.COUNT - 1
  Me.CancelButton(i).top = top
  Me.CancelButton(i).left = left + (Me.OKButton.COUNT + i) * (KW + 10)
  Me.CancelButton(i).Height = 375
  Me.CancelButton(i).Width = KW
 Next i
End Sub ' Form_Resize()

Private Sub OKButton_Click(Index%)
 Arzt = Index
 Me.Hide
End Sub ' OKButton_Click(Index%)

Private Sub OKButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Call CancelButton_KeyDown(Index, KeyCode, Shift)
End Sub ' OKButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
