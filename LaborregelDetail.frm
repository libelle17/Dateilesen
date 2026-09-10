VERSION 5.00
Begin VB.Form LaborregelDetail 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Laborregel"
   ClientHeight    =   6080
   ClientLeft      =   2400
   ClientTop       =   2400
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "LaborregelDetail"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6080
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtLaborparameter 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   150
      Width           =   5800
   End
   Begin VB.Label lblLaborparameter 
      Caption         =   "Anfordern"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   3200
   End
   Begin VB.TextBox txtAbkueRegex 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   520
      Width           =   5800
   End
   Begin VB.Label lblAbkueRegex 
      Caption         =   "Abkü-Regex (Wert/Intervall; leer=wie oben)"
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   520
      Width           =   3200
   End
   Begin VB.TextBox txtICDRegex 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   890
      Width           =   5800
   End
   Begin VB.Label lblICDRegex 
      Caption         =   "ICD-Regex (leer = diagnoseunabhängig)"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   890
      Width           =   3200
   End
   Begin VB.TextBox txtEigenerWertVgl 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   1260
      Width           =   5800
   End
   Begin VB.Label lblEigenerWertVgl 
      Caption         =   "eigener letzter Wert (z.B. >24)"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   1260
      Width           =   3200
   End
   Begin VB.TextBox txtBMIVgl 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   1630
      Width           =   5800
   End
   Begin VB.Label lblBMIVgl 
      Caption         =   "BMI (z.B. <18)"
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   1630
      Width           =   3200
   End
   Begin VB.TextBox txtAlterVgl 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   2000
      Width           =   5800
   End
   Begin VB.Label lblAlterVgl 
      Caption         =   "Alter (z.B. >50)"
      Height          =   285
      Left            =   150
      TabIndex        =   11
      Top             =   2000
      Width           =   3200
   End
   Begin VB.TextBox txtRRVgl 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Top             =   2370
      Width           =   5800
   End
   Begin VB.Label lblRRVgl 
      Caption         =   "RR syst. Ø letzte 6 Mon. (z.B. >140)"
      Height          =   285
      Left            =   150
      TabIndex        =   13
      Top             =   2370
      Width           =   3200
   End
   Begin VB.ComboBox cboVersicherung 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown-Liste
      TabIndex        =   14
      Top             =   2740
      Width           =   5800
   End
   Begin VB.Label lblVersicherung 
      Caption         =   "Versicherung"
      Height          =   285
      Left            =   150
      TabIndex        =   15
      Top             =   2740
      Width           =   3200
   End
   Begin VB.ComboBox cboDMP 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown-Liste
      TabIndex        =   16
      Top             =   3110
      Width           =   5800
   End
   Begin VB.Label lblDMP 
      Caption         =   "DMP"
      Height          =   285
      Left            =   150
      TabIndex        =   17
      Top             =   3110
      Width           =   3200
   End
   Begin VB.TextBox txtMedikamentRegex 
      Height          =   285
      Left            =   3480
      TabIndex        =   18
      Top             =   3480
      Width           =   5800
   End
   Begin VB.Label lblMedikamentRegex 
      Caption         =   "Medikament-Regex"
      Height          =   285
      Left            =   150
      TabIndex        =   19
      Top             =   3480
      Width           =   3200
   End
   Begin VB.TextBox txtIntervallMonate 
      Height          =   285
      Left            =   3480
      TabIndex        =   20
      Top             =   3850
      Width           =   1500
   End
   Begin VB.Label lblIntervallMonate 
      Caption         =   "Intervall (Mon., 0=einmalig)"
      Height          =   285
      Left            =   150
      TabIndex        =   21
      Top             =   3850
      Width           =   3200
   End
   Begin VB.TextBox txtKommentar 
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Top             =   4220
      Width           =   5800
   End
   Begin VB.Label lblKommentar 
      Caption         =   "Kommentar"
      Height          =   285
      Left            =   150
      TabIndex        =   23
      Top             =   4220
      Width           =   3200
   End
   Begin VB.Label lblGueltigAb 
      BorderStyle     =   1  'Fest Einfach
      Height          =   285
      Left            =   3480
      TabIndex        =   24
      Top             =   4590
      Width           =   5800
   End
   Begin VB.Label lblGueltigAbCap 
      Caption         =   "gültig ab"
      Height          =   285
      Left            =   150
      TabIndex        =   25
      Top             =   4590
      Width           =   3200
   End
   Begin VB.TextBox txtGueltigBis 
      Height          =   285
      Left            =   3480
      TabIndex        =   26
      Top             =   4960
      Width           =   5800
   End
   Begin VB.Label lblGueltigBis 
      Caption         =   "gültig bis (leer = unbefristet)"
      Height          =   285
      Left            =   150
      TabIndex        =   27
      Top             =   4960
      Width           =   3200
   End
   Begin VB.CommandButton cmdAbbrechen 
      Caption         =   "Abbre&chen"
      Height          =   400
      Left            =   7880
      TabIndex        =   28
      Top             =   5480
      Width           =   1400
   End
   Begin VB.CommandButton cmdLoeschen 
      Caption         =   "&Löschen"
      Height          =   400
      Left            =   6320
      TabIndex        =   29
      Top             =   5480
      Width           =   1400
   End
   Begin VB.CommandButton cmdSpeichern 
      Caption         =   "&Speichern"
      Height          =   400
      Left            =   4760
      TabIndex        =   30
      Top             =   5480
      Width           =   1400
   End
End
Attribute VB_Name = "LaborregelDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Bearbeitungsformular fuer einen Datensatz der Tabelle laborregel.
' ID = 0 -> neuer Datensatz, sonst wird der Datensatz mit dieser ID geladen.
Public ID As Long

Private Sub Form_Load()
 Me.cboVersicherung.AddItem "egal"
 Me.cboVersicherung.AddItem "privat"
 Me.cboVersicherung.AddItem "kasse"
 Me.cboDMP.AddItem "egal"
 Me.cboDMP.AddItem "ja"
 Me.cboDMP.AddItem "nein"
 If ID <> 0 Then
  Dim r As New ADODB.Recordset
  myFrag r, "SELECT * FROM laborregel WHERE ID = " & ID
  If Not r.BOF Then
   Me.txtLaborparameter = nz(r!Laborparameter, "")
   Me.txtAbkueRegex = nz(r!AbkueRegex, "")
   Me.txtICDRegex = nz(r!ICDRegex, "")
   Me.txtEigenerWertVgl = nz(r!EigenerWertVgl, "")
   Me.txtBMIVgl = nz(r!BMIVgl, "")
   Me.txtAlterVgl = nz(r!AlterVgl, "")
   Me.txtRRVgl = nz(r!RRVgl, "")
   Me.cboVersicherung.Text = IIf(LenB(nz(r!Versicherung, "")) = 0, "egal", r!Versicherung)
   Me.cboDMP.Text = IIf(LenB(nz(r!DMP, "")) = 0, "egal", r!DMP)
   Me.txtMedikamentRegex = nz(r!MedikamentRegex, "")
   Me.txtIntervallMonate = nz(r!IntervallMonate, "")
   Me.txtKommentar = nz(r!Kommentar, "")
   Me.lblGueltigAb = r!GueltigAb
   Me.txtGueltigBis = nz(r!GueltigBis, "")
  End If
  Set r = Nothing
 Else
  Me.cboVersicherung.Text = "egal"
  Me.cboDMP.Text = "egal"
  Me.lblGueltigAb = Now()
  Me.cmdLoeschen.Enabled = False
 End If
End Sub ' Form_Load

' baut aus einem Textfeld entweder NULL (bei leerem Text) oder einen gequoteten SQL-String-Literal
Private Function sqlStr$(ByVal s$)
 s = Trim$(s)
 If LenB(s) = 0 Then
  sqlStr = "NULL"
 Else
  sqlStr = "'" & REPLACE$(s, "'", "''") & "'"
 End If
End Function ' sqlStr$

Private Sub cmdSpeichern_Click()
 On Error GoTo fehler
 If LenB(Trim$(Me.txtLaborparameter)) = 0 Then
  MsgBox "Bitte einen Wert fuer 'Anfordern' angeben.", vbExclamation
  Exit Sub
 End If
 Dim vers$, dmp$, intv$
 vers = IIf(Me.cboVersicherung.Text = "egal", "NULL", "'" & Me.cboVersicherung.Text & "'")
 dmp = IIf(Me.cboDMP.Text = "egal", "NULL", "'" & Me.cboDMP.Text & "'")
 intv = IIf(LenB(Trim$(Me.txtIntervallMonate)) = 0, "NULL", REPLACE$(CStr(MachNumerisch(Me.txtIntervallMonate)), ",", "."))
 Dim sql$
 If ID = 0 Then
  sql = "INSERT INTO laborregel (Laborparameter,AbkueRegex,ICDRegex,EigenerWertVgl,BMIVgl,AlterVgl,RRVgl,Versicherung,DMP,MedikamentRegex,IntervallMonate,Kommentar,GueltigAb,GueltigBis) VALUES (" & _
   sqlStr(Me.txtLaborparameter) & "," & sqlStr(Me.txtAbkueRegex) & "," & sqlStr(Me.txtICDRegex) & "," & sqlStr(Me.txtEigenerWertVgl) & "," & _
   sqlStr(Me.txtBMIVgl) & "," & sqlStr(Me.txtAlterVgl) & "," & sqlStr(Me.txtRRVgl) & "," & vers & "," & dmp & "," & _
   sqlStr(Me.txtMedikamentRegex) & "," & intv & "," & sqlStr(Me.txtKommentar) & ",NOW()," & sqlStr(Me.txtGueltigBis) & ")"
 Else
  sql = "UPDATE laborregel SET Laborparameter=" & sqlStr(Me.txtLaborparameter) & _
   ",AbkueRegex=" & sqlStr(Me.txtAbkueRegex) & _
   ",ICDRegex=" & sqlStr(Me.txtICDRegex) & ",EigenerWertVgl=" & sqlStr(Me.txtEigenerWertVgl) & _
   ",BMIVgl=" & sqlStr(Me.txtBMIVgl) & ",AlterVgl=" & sqlStr(Me.txtAlterVgl) & ",RRVgl=" & sqlStr(Me.txtRRVgl) & _
   ",Versicherung=" & vers & ",DMP=" & dmp & ",MedikamentRegex=" & sqlStr(Me.txtMedikamentRegex) & _
   ",IntervallMonate=" & intv & ",Kommentar=" & sqlStr(Me.txtKommentar) & ",GueltigBis=" & sqlStr(Me.txtGueltigBis) & _
   " WHERE ID=" & ID
 End If
 myEFrag sql
 Unload Me
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdSpeichern_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdSpeichern_Click

Private Sub cmdLoeschen_Click()
 If ID = 0 Then Exit Sub
 If MsgBox("Regel fuer """ & Me.txtLaborparameter & """ wirklich loeschen?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
 On Error GoTo fehler
 myEFrag "DELETE FROM laborregel WHERE ID = " & ID
 Unload Me
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLoeschen_Click/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' cmdLoeschen_Click

Private Sub cmdAbbrechen_Click()
 Unload Me
End Sub ' cmdAbbrechen_Click

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then Unload Me
End Sub ' Form_KeyPress
