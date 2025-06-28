VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PathLabForm 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Datumsauswahl"
   ClientHeight    =   750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10530
   KeyPreview      =   -1  'True
   LinkTopic       =   "PathLabForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Auswahl 
      Caption         =   "&Auswahl"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Labordatei 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbre&chen"
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label LabordateiLab 
      Caption         =   "&Datum"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PathLabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ePL As New PatListe
Public eLese As Lese
Public Titel$
Public Filter$
Public weiter% ' 0 = nicht weiter, -1 = weiter
Public erg$
Public initDir$

Private Sub Auswahl_Click()
 Dim fileflags As FileOpenConstants
 Dim filefilter$
 Me.CommonDialog1.DialogTitle = Titel
 'Set the default file name AND filter
 CommonDialog1.initDir = initDir
 CommonDialog1.Filename = vNS
 filefilter = Filter
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 CommonDialog1.flags = fileflags
 CommonDialog1.ShowOpen
 Me.Labordatei = CommonDialog1.Filename
 
 If FSO.FileExists(Me.Labordatei) Then
  Call OKButton_Click
 End If
End Sub ' Auswahl_Click

Private Sub CancelButton_Click()
 Unload Me
End Sub ' CancelButton_Click()

Private Sub Form_Initialize()
 Titel = "Labor-Excel-Datei auswählen:"
' Filter = "Excel-Dateien(*.xls)|*.xls|Alle Dateien (*.*)|*.*"
 Filter = "LDT-Dateien(*.ldt)|*.ldt|CSV-Dateien(*.csv)|*.csv|Alle Dateien (*.*)|*.*"
 initDir = pVerz
 weiter = True
End Sub ' Form_Initialize()

Private Sub Form_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case 27
   Unload Me
  Case 13
   Call OKButton_Click
 End Select
End Sub ' Form_KeyPress(KeyAscii As Integer)

Private Sub OKButton_Click() ' -> Pathologische Laborwerte anschauen -> LabordateiAnzeig(Me.LabDat)
 If weiter Then
  If IsDate(Me.Labordatei) Then
   ePL.PLArt = artlab
   Set ePL.hlese = eLese
   ePL.LabDatum = Me.Labordatei
   ePL.Show
  End If
' If FSO.FileExists(Me.Labordatei) Then
'  ePL.PLArt = artlab
'  Set ePL.hlese = eLese
'  ePL.LabDat = Me.Labordatei
'  ePL.Show
' End If
 Else
  erg = Me.Labordatei
  Unload Me
 End If
End Sub ' OKButton_Click

