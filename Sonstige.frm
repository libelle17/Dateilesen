VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Sonstige 
   Caption         =   "Sonstige Optionen"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Sonstige"
   ScaleHeight     =   7785
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox DebugDatei 
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   5775
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox QuelleAnamneseBögen 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog CommonDialogSonstige 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox DateiNachzuholen 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label DebugDateiBez 
      Caption         =   "De&bug-Datei:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label QuelleAnamneseBögenBez 
      Caption         =   "&Quelle zum Abholen der Anamnesebögen:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label DateiNachzuhoelenBez 
      Caption         =   "&Datei für nachzuholende Laborimporte:"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Sonstige"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese

Private Sub DateiNachzuhoelenBez_Click()
 Call DNHDialog
End Sub


Private Sub DebugDatei_GotFocus()
 Call DebugDDialog
End Sub


Private Sub DebugDateiBez_Click()
 Call DebugDDialog
End Sub

Private Sub DebugDateiBez_DblClick()
 Call DebugDDialog
End Sub

Private Sub DebugDateiBez_LinkClose()
 Call DebugDDialog

End Sub

Private Sub DebugDateiBez_LinkOpen(Cancel As Integer)
 Call DebugDDialog

End Sub


Private Sub Form_Load()
' imAufbauSonstige = True
' Call HolReg(hlese)
' imAufbauSonstige = False
 With Me.CommonDialogSonstige
  .Orientation = cdlLandscape
  .flags = 0
' .Flags = .Flags OR FileOpenConstants.cdlOFNExplorer
  .flags = .flags Or FileOpenConstants.cdlOFNHideReadOnly ' Schreibgeschützt-Checkbox entfernen
  .flags = .flags Or FileOpenConstants.cdlOFNLongNames ' '   Lange Dateinamen erlauben (nur sinnvoll bei Nicht-Win95-Design)
  .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
  .flags = .flags Or FileOpenConstants.cdlOFNPathMustExist
 End With
End Sub ' Form_Load()

Private Sub Form_Terminate()
'
End Sub

Private Sub QuelleAnamneseBögen_GotFocus()
 Call QuellDialog
End Sub

Private Sub QuelleAnamneseBögenBez_Click()
 Call QuellDialog
End Sub
Function QuellDialog()
  With Me.CommonDialogSonstige
   .Filename = "*.mdb"
   .DialogTitle = "Datenbank mit Anamnesebögen"
   If FSO.FileExists(Me.QuelleAnamneseBögen) Then
    .initDir = FSO.GetParentFolderName(FSO.GetFile(Me.QuelleAnamneseBögen).path)
   Else
    .initDir = aVerz ' uverz & "tmexport"
   End If
   .flags = .flags Or FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    Me.QuelleAnamneseBögen = .Filename
   End If
  End With
End Function ' QuellDialog

Function DNHDialog()
  With Me.CommonDialogSonstige
   .Filename = "*.txt"
   .DialogTitle = "Dateiname für nachzuholende Laborimporte"
   If FSO.FileExists(Me.DateiNachzuholen) Then
    .initDir = FSO.GetParentFolderName(FSO.GetFile(Me.DateiNachzuholen).path)
   Else
    .initDir = aVerz ' uverz & "Anamnese"
   End If
   .flags = .flags And Not FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    Me.DateiNachzuholen = FSO.GetParentFolderName(.Filename) & "\" & "LaborimporteNachzuholen" & Format$(Now, "yyyymmdd.hhmmss") & ".txt"
   End If
  End With
End Function ' BDTDateiDialog()

Function DebugDDialog()
  With Me.CommonDialogSonstige
   .Filename = "*.txt"
   .DialogTitle = "Dateiname für Debug.Protokoll"
   If FSO.FileExists(Me.DebugDatei) Then
    .initDir = FSO.GetParentFolderName(FSO.GetFile(Me.DebugDatei).path)
   Else
    .initDir = üVerz ' uverz & "Anamnese"
   End If
   .flags = .flags And Not FileOpenConstants.cdlOFNFileMustExist
   .ShowOpen
   If FSO.FileExists(.Filename) Then
    Me.DateiNachzuholen = FSO.GetParentFolderName(.Filename) & "\" & "Dateilese-DebugDatei" & Format$(Now, "yyyymmdd.hhmmss") & ".txt"
   End If
  End With
End Function ' BDTDateiDialog()
