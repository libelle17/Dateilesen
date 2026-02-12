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
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Index           =   1
      Left            =   11880
      TabIndex        =   5
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Schade"
      Height          =   375
      Index           =   2
      Left            =   9360
      TabIndex        =   2
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Kothny"
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   1
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Überspringen"
      Height          =   375
      Index           =   0
      Left            =   10680
      TabIndex        =   3
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Hammerschmidt"
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   0
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Texte 
      Caption         =   "Inhalt"
      Height          =   8175
      Left            =   0
      TabIndex        =   4
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
 Abbruch = Index + 1
 Me.Hide
End Sub ' CancelButton_Click()

Private Sub CancelButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13
   Call OKButton_Click(KeyCode)
  Case 27
   Call CancelButton_Click(Index)
 End Select
End Sub ' CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Form_Resize()
 Dim i%, top&, left&
 top = Me.Height - 9090 + 8160 + 100
 left = Me.Width - (13455 - 8040) - 1345
 For i = 0 To 2
  Me.OKButton(i).top = top
  Me.OKButton(i).left = left + i * 1345
 Next i
 For i = 0 To 1
  Me.CancelButton(i).top = top
  Me.CancelButton(i).left = left + (3 + i) * 1345
 Next i
End Sub ' Form_Resize()

Private Sub OKButton_Click(Index As Integer)
 Arzt = Index
 Me.Hide
End Sub

Private Sub OKButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Call CancelButton_KeyDown(Index, KeyCode, Shift)
End Sub ' OKButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
