VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PatListe 
   Caption         =   "Patientenliste"
   ClientHeight    =   10425
   ClientLeft      =   510
   ClientTop       =   1245
   ClientWidth     =   18255
   FillColor       =   &H00FFFFFF&
   Icon            =   "PatListe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   18255
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   15
      Left            =   12480
      TabIndex        =   20
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   14
      Left            =   11640
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   13
      Left            =   10560
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   14760
      TabIndex        =   17
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   12
      Left            =   9480
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   10
      Left            =   7320
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   9
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.ListBox Li1 
      Height          =   255
      Left            =   17160
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cb1 
      Height          =   315
      Left            =   15960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG 
      Height          =   9975
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   17595
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   14040
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "PatListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hlese As Lese
Public AnfCode%
Const MTBeg% = 1
Const MTMax% = 1
Const DMP_Import = "DMP_Import.BDT"
Const rot& = 10790143
Dim MFGTyp%
Dim Cstumm%
Dim GruSp%, NachNameSp%, PIDSp%, ICDSp%, VorDokuSp%
Dim altFarbe&(MTBeg To MTMax)
Dim cRow&(MTBeg To MTMax), cCol&(MTBeg To MTMax), cRowSel&(MTBeg To MTMax) ' für Tb1 und li1 und spätere Änderungen
Dim obtb%
Dim noenter%
Dim fgespei%(MTBeg To MTMax) ' Farbe gespeichert
Private Declare Sub Sleep Lib "kernel32" (ByVal ms&)
Dim Nachname$
Dim aktr&, altr&, alttop&, altC&
Enum Zeilenart
 typ1neu = 1
 typ2neu
 typ1alt
 typ2alt
End Enum
Enum ArtTyp
 artPat = 1
 artDiag
 artlab
End Enum
Public Art As ArtTyp
Public boga As Zeilenart ' nur für callMachDMPBogen
Public DokuDat As Date ' nur für callMachDMPBogen
Public LabDat$ ' nur für PathLabForm

Public Sub Pause(Millisekunden As Long)
Sleep Millisekunden
End Sub
Private Function getZeilenart() As Zeilenart
 With MFG
'  altC = .col
 If .TextMatrix(.Row, VorDokuSp) = "" Then ' Vorbefunde
  If .TextMatrix(.Row, ICDSp) Like "E10*" Then
   getZeilenart = typ1neu
  Else
   getZeilenart = typ2neu
  End If
 Else
  If .TextMatrix(.Row, ICDSp) Like "E10*" Then
   getZeilenart = typ1alt
  Else
   getZeilenart = typ2alt
  End If
 End If
'.col = altC
 End With
End Function ' getZeilenart() As Zeilenart
Public Sub TopAusricht()
  Dim i%, firstvis%
  With MFG
    On Error Resume Next
    If .Row - altr > (.Height / .CellHeight * 0.9) Then
     .TopRow = .Row - 1
    End If
  End With
  For i = Me.Command1.Count - 1 To 0 Step -1
   If Me.Command1(i).Visible Then
    firstvis = i
    Exit For
   End If
  Next i
  Me.Label1.Left = Me.Command1(firstvis).Left + Me.Command1(firstvis).Width + 100
  Me.Text1.Left = Me.Label1.Left + Me.Label1.Width + 50
End Sub ' TopAusricht
Public Sub verschieb(Ziel$)
 Call Shell("cmd.exe /c move """ & Me.LabDat & """ """ & Ziel & "")
End Sub
Private Sub lösch()
 Dim erg&
 erg = MsgBox("Wollen Sie wirklich die Datei:" & vbCrLf & Me.LabDat & vbCrLf & "löschen?", vbDefaultButton2 + vbYesNo)
 If erg = vbYes Then
  Kill Me.LabDat
 End If
End Sub

Public Sub Command1_Click(index As Integer)
 Select Case Art
  Case artlab
   Select Case index
    Case 0: Call verschieb("P:\Kothny Labor")
    Case 1: Call verschieb("P:\Schade Labor")
    Case 2: Call lösch
   End Select
  Case artPat
   Select Case index
    Case 0: Call dokuErstelle
    Case 1: Call FertigStellen(Me.MFG.Row, True) ' in TurbomedAnzeigen
    Case 2: Call FertigStellen(Me.MFG.Row)
    Case 3: Call Suchen
    Case 4: Call Weitersuchen
    Case 5: Call NächsterGleichartiger
  '  Case 6: Call Typ1ED
  '  Case 7: Call Typ1FD
  '  Case 8: Call Typ2ED
  '  Case 9: Call Typ2FD
    Case 6: Call DokuBeliebig
    Case 7: Call FertigStellenBeliebig
    Case 8: Call Anleitung
    Case 9: Call htmlb
    Case 10: Call zurück
    Case 11: Call Auffrisch
    Case 12: Call alleFS
    Case 13: Call aktFS
   End Select
 End Select
End Sub ' Command1_Click(Index As Integer)
Sub aktFS()
 Call doFS(False)
End Sub
Sub alleFS()
 Dim pCol As New Collection, i&, altCol&, altRow&, Zeile
 With Me.MFG
  .col = 0
  For i = .Row To .Rows - 1
   .Row = i
   If .CellBackColor = rot Then
    pCol.Add i
   End If
  Next i
  .Row = altRow
  .col = altCol
 End With
 For Each Zeile In pCol
  Call FertigStellen((Zeile))
  Pause (100)
  SendKeys "{ESC}", True
 Next Zeile
End Sub
Sub Auffrisch()
 Unload Me
 fgespei(MFGTyp) = 0
 Load Me
End Sub
Private Sub Anleitung()

End Sub ' Anleitung

Private Sub Form_Unload(Cancel As Integer)
 Me.hlese.Show
End Sub

Private Sub li1_Click()
 Dim rAF&, i&
 If Cstumm = 0 Then
  Li1.Visible = False
  For i = cRow(MFGTyp) To cRowSel(MFGTyp)
   MFG.col = 0
   MFG.Row = i
   Call DBCn.Execute("update diagreihe set dg1 = '" & Li1 & "' where icd = '" & MFG.TextMatrix(i, 0) & "'", rAF)
   If rAF <> 1 Then
    MsgBox ("Fehler beim Diagnosenupdate: " & Li1 & " " & MFG.TextMatrix(i, 0))
   End If
   MFG.col = GruSp
   MFG.Text = Li1
  Next
'  Call MFGrefresh
  If Me.MFG.Row < Me.MFG.Rows Then
   Me.MFG.Row = Me.MFG.Row + 1
  End If
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
  Me.Li1.Clear
  Me.Li1.Visible = False
 End If
 Me.MFG.SetFocus
End Sub ' li1_click
Sub FertigStellenBeliebig()
 Dim erg As Variant
 Select Case Art
  Case artDiag
  Case artPat
   erg = InputBox("Bitte Pat-ID eingeben:", "Eingabe der Patientennummer", Me.MFG.TextMatrix(Me.MFG.Row, 1))
   If IsNumeric(erg) Then
    Call FertigStellen(0, , Me.MFG.TextMatrix(Me.MFG.Row, 1))
   End If
 End Select
End Sub ' FertigStellenBeliebig()

Private Sub DokuBeliebig() ' Doku zu beliebigem Patienten
 Dim erg As Variant
 Select Case Art
  Case artDiag
  Case artPat
   erg = InputBox("Bitte Pat-ID eingeben:", "Eingabe der Patientennummer", Me.MFG.TextMatrix(Me.MFG.Row, 1))
   If IsNumeric(erg) Then
    Call callMachDMPBogen(CLng(erg), , True)
   End If
 End Select
End Sub ' DokuBeliebig

Public Sub NächsterGleichartiger()
 Dim aktZA As Zeilenart, i&
 MFG.SetFocus
aktZA = getZeilenart
 With MFG
  alttop = .TopRow
  altr = .Row
  Call MFG_leavecell
  For aktr = .Row + 1 To .Rows - 1
   .Row = aktr
   If getZeilenart = aktZA Then
    Call TopAusricht
    Call mfg_entercell
    Exit Sub
   End If
  Next
  .Row = altr
  Call mfg_entercell
 End With
 
End Sub ' NächsterGleichartiger
Public Sub Weitersuchen() ' Weitersuchen
 MFG.SetFocus
 With MFG
  alttop = .TopRow
'  altr = .Row
'  altC = .col
  Call MFG_leavecell
'  .col = 3
  For aktr = .Row + 1 To .Rows - 1
   .Row = aktr
   If LCase(.TextMatrix(aktr, .col)) Like LCase(Nachname) & "*" Then
    Call TopAusricht
    Call mfg_entercell
    Exit Sub
   End If
  Next aktr
'  .Row = altr
'  .col = altC
  Call mfg_entercell
 End With
End Sub ' Weitersuchen
Public Sub Suchen() ' Suchen
 MFG.SetFocus
 Nachname = InputBox("Bitte Suchstring eingeben:", "Zeilensuche")
 With MFG
  alttop = .TopRow
'  altr = .Row
'  altC = .col
  Call MFG_leavecell
'  .col = 3
  For aktr = .Row To .Rows - 1
   .Row = aktr
   If LCase(.TextMatrix(aktr, .col)) Like LCase(Nachname) & "*" Then
    Call TopAusricht
    Call mfg_entercell
    Exit Sub
   End If
  Next aktr
'  .Row = altr
'  .col = altC
  Call mfg_entercell
 End With
End Sub ' Suchen
Public Sub FertigStellen(Zeile&, Optional nuranzeigen%, Optional PatID&) ' nachdem BDT-Datei(en) manuell importiert wurde(n)
 Const obstumm% = 0
 Dim hnd&, VorDoku$, Pat_id&, dtyp%, rs As New Recordset
 Dim DT As dmptyp, j%
 Dim rTyp As New ADODB.Recordset
 With MFG
  .SetFocus
  alttop = .TopRow
  altr = Zeile
'  altC = .col
  If PatID <> 0 Then
   Pat_id = PatID
   rs.Open "select quartal from faelle where pat_id = " & PatID & " and bhfb < now()- " & Verspätung & " order by bhfb desc", DBCn
  Else
   VorDoku = .TextMatrix(Zeile, VorDokuSp)
   Pat_id = .TextMatrix(Zeile, PIDSp)
  End If
'  .col = altC
 End With
' Call DMPAusgeb(dt, Pat_id, Not obstumm)
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  AppActivate "TurboMed", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ENTER}", True
  Pause (100)
  SendKeys "{F12}", True
  Pause (100)
'  SendKeys "{ESC}", True
'  Pause (100)
'  SendKeys "{ESC}", True
'  Pause (100)
  SendKeys "p", True
  Pause (100)
  SendKeys "" & "{bs}" & Pat_id & "", True
  
  Call doFS(nuranzeigen, True)
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FertigStellen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' FertigStellen
Sub doFS(nuranzeigen%, Optional ohneakt% = False)
  Me.Text1.Enabled = False
  If Not ohneakt Then AppActivate "TurboMed", True
  Pause (100)
  SendKeys "{ENTER}", True
  Pause (100)
  SendKeys "{F3}", True
  Pause (100)
  If nuranzeigen Then
'   AppActivate "TurboMed", True
   SendKeys "%{F1}", True
   Pause (100)
   Exit Sub
  End If
  SendKeys "{F5}", True
  Pause (100)
  SendKeys "{DOWN " & Me.Text1 - 2 & "}", True ' Filter DMP wählen
  Pause (100)
  SendKeys "{Enter}", True
  Pause (100)
  SendKeys "{Enter}", True
  Pause (100)
  SendKeys "%D", True
  Pause (100)
  SendKeys " ", True
  Pause (100)
  SendKeys "%O", True
  Pause (100)
  SendKeys " ", True
  Pause (100)
  SendKeys " " & Format(QEnd(ZQuart(Now() - Verspätung)), "ddmmyyyy") & "{TAB}", True
  Pause (200)
  SendKeys "%O", True
  Pause (100)
  SendKeys " ", True
  Pause (100)
  SendKeys " ", True
  Pause (100)
  SendKeys " ", True
  Pause (100)
'  AppActivate "C:\TurboMed\Tmp", True
'  Pause (100)
'  Dim rs As New ADODB.Recordset
'  rs.Open "select nachname, vorname, gebdat from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
'  SendKeys "%DU", True
'  Pause (100)
'  SendKeys "P:\" & rs!Nachname & " " & rs!Vorname & ", geb. " & rs!GebDat & ", DMP-Doku " & QEnd(ZQuart(Now() - Verspätung))
'  Pause (100)
'  SendKeys "{Enter}", True
  Call htmlb(obstumm:=True)
  Me.Text1.Enabled = True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doFS/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub      ' doFS
Private Sub zurück()
 Dim erg$, erg1$, erg2$
 erg = Dir(hVerz & DMP_Import)
 If LenB(erg) <> 0 Then
  Open hVerz & DMP_Import & "kopie.BDT" For Output As #389
  Open hVerz & DMP_Import For Input As #390
  Do While Not EOF(390)
   Input #390, erg
nächster:
   If Mid$(erg, 4, 4) = "8000" Then
    Input #390, erg1
    Input #390, erg2
    If Mid$(erg2, 4, 4) = "3000" And Mid$(erg2, 8, 4) = MFG.TextMatrix(MFG.Row, 1) Then
'     Stop
     Do While Not EOF(390)
      Input #390, erg
      If Mid$(erg, 4, 4) = "8000" Then GoTo nächster
     Loop
    Else
'     Stop
     Print #389, erg
     Print #389, erg1
     Print #389, erg2
    End If
   Else
    Print #389, erg
   End If
  Loop
  Close #389
  Close #390
  Kill hVerz & DMP_Import
  Name hVerz & DMP_Import & "kopie.BDT" As hVerz & DMP_Import
  Call Auffrisch
 End If
End Sub ' zurück
Private Sub htmlb(Optional obstumm%) ' HTML-Dateien bewahren
  Dim TempDir$, erg$, UVerz$(), vz, lin$, lina$, dmpart$, NN$, VN$, Geb$, aktDat$, Zahl&
  Const Ziel$ = "P:\"
  ReDim UVerz(0)
  TempDir = getDokPfad("Temporärdateien")
  erg = Dir(TempDir & "\", vbDirectory)
  Do
   If LenB(erg) = 0 Then Exit Do
   If erg <> "." And erg <> ".." Then
    If (GetAttr(TempDir & "\" & erg) And vbDirectory) = vbDirectory Then
     UVerz(UBound(UVerz)) = erg
     ReDim Preserve UVerz(UBound(UVerz) + 1)
    End If
   End If
   erg = Dir
  Loop
  For Each vz In UVerz
   erg = Dir(TempDir & "\" & vz & "\")
   Do
    If LenB(erg) = 0 Then Exit Do
    If InStrB(erg, ".html") <> 0 Then
     lina = ""
     dmpart = ""
     NN = ""
     VN = ""
     Geb = ""
     aktDat = TempDir & "\" & vz & "\" & erg
     Open aktDat For Input As #329
     Do While Not EOF(329)
      lina = lin
      Input #329, lin
      If InStrB(lina, """DMP-Header""") <> 0 Then
       dmpart = lin
       dmpart = Mid$(dmpart, InStr(dmpart, "<h1>") + 4)
       dmpart = Left$(dmpart, InStr(dmpart, "</h1>") - 1)
      ElseIf InStrB(lina, "Vorname") <> 0 Then
       NN = lin
       NN = Mid$(NN, InStr(NN, "<div class=""header_value"">") + Len("<div class=""header_value"">"))
       NN = Left$(NN, InStr(InStr(NN, "<br>") + 1, NN, "<br>") - 1)
       NN = Replace$(NN, "<br>", " ")
      ElseIf InStrB(lina, "geb. am") <> 0 Then
       Geb = lin
       Geb = Mid$(Geb, InStr(Geb, "<div class=""header_value"">") + Len("<div class=""header_value"">"))
       Geb = Left$(Geb, InStr(Geb, "</div>") - 1)
       Exit Do
      End If
     Loop
     Close #329
     FileCopy aktDat, Ziel & NN & ", geb. " & Geb & ", DMP-" & dmpart & ".html"
     Zahl = Zahl + 1
     Kill aktDat
     Debug.Print erg & " " & dmpart & " " & NN & " " & Geb
    End If
    erg = Dir
   Loop
  Next vz
  If Not obstumm Then
   MsgBox "Fertig mit HTML bewahren, " & Zahl & " Dateien nach " & Ziel & " kopiert und umbenannt"
  End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in htmlb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub
Public Sub dokuErstelle()
 Dim VorDoku$, Pat_id&
 Select Case Art
  Case artDiag
   Call MFGrefresh
  Case Else
  With MFG
  .SetFocus
  alttop = .TopRow
  altr = .Row
  VorDoku = .TextMatrix(.Row, VorDokuSp)
  Pat_id = .TextMatrix(.Row, PIDSp)
 End With
 Call callMachDMPBogen(Pat_id, VorDoku)
 End Select
End Sub ' dokuErstelle
Public Sub callMachDMPBogen(Pat_id&, Optional VorDoku$, Optional obmitauswahl%)
 Dim rTyp As New ADODB.Recordset, dtyp%
 Dim dmpba As New DMPBogenauswahl
 rTyp.Open "select icd from diagnosen where pat_id = " & Pat_id & " and diagsicherheit <> 'A' and icd like 'E1%' order by icd", DBCn, adOpenStatic, adLockReadOnly
 If Not rTyp.EOF Then dtyp = Mid$(rTyp.Fields(0), 3, 1) + 1
 If dtyp = 2 Then If VorDoku = "" Then boga = typ2neu Else boga = typ2alt Else If VorDoku = "" Then boga = typ1neu Else boga = typ1alt
 DokuDat = QEnd(ZQuart(Now() - Verspätung))
 If obmitauswahl Then
  dmpba.Option1(boga - 1) = True
  Set dmpba.vater = Me
  Dim rs As New ADODB.Recordset
  rs.Open "select nachname n, vorname v from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
  If rs.EOF Then
   MsgBox "Pat_id " & Pat_id & " nicht gefunden!"
   Exit Sub
  Else
   dmpba.Caption = "Erstelle DMP-Bogen zu " & Pat_id & " (" & rs!n & ", " & rs!V & ")"
   dmpba.DokuDatum = DokuDat
   dmpba.Show vbModal, Me
'  boga = dmpba.Option1
   If boga = 0 Then Exit Sub
  End If
 End If
 Call machDMPBogen(Pat_id, boga, DokuDat)
End Sub ' callMachDMPBogen(Pat_id&, Optional VorDoku$, Optional obmitauswahl%)
'Public Sub Typ1ED()
' Dim Pat_id&
' With MFG
''  altC = .col
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
' End With
' Call machDMPBogen(Pat_id, typ1neu)
'End Sub ' command6_click
'Public Sub Typ1FD()
' Dim Pat_id&
' With MFG
''  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
' End With
' Call machDMPBogen(Pat_id, typ1alt)
'End Sub
'Public Sub Typ2ED()
' Dim Pat_id&
' With MFG
'  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
'  End With
'  Call machDMPBogen(Pat_id, typ2neu)
'End Sub
'Public Sub Typ2FD()
' Dim Pat_id&
' With MFG
'  altC = .col
''  .col = 2
'  Pat_id = .TextMatrix(.Row, PIDSp)
''  .col = altC
'  End With
'  Call machDMPBogen(Pat_id, typ2alt)
'End Sub
Public Sub machDMPBogen(Pat_id&, boga As Zeilenart, DokuDat As Date)
 Const obstumm% = 0
 Dim ranam As New ADODB.Recordset
 Dim rnam As New ADODB.Recordset
 Dim rfal As New ADODB.Recordset
 Dim rform As New ADODB.Recordset
 Dim hnd&
 Dim DT As dmptyp, j%, pos%
 Dim RRsyst%, RRdiast%, RRstr$
 Dim SendStr$, SendStr1a$, SendStr2$
 Dim BDT As New BDTSchreib, erg$, DokuDatS$ ' DokuDat As Date
 Dim ErstellDat$, i%, GewichtS$, GrößeS$, HbA1cS$, HypoZS$, KrZS$, KreaS$, RRS$
  
 Dim WagSt$, ArmSt$
 On Error GoTo fehler
 erg = Dir(hVerz & DMP_Import)
 If LenB(erg) <> 0 Then
  Select Case MsgBox("Datei '" & hVerz & DMP_Import & "' vorhanden. Löschen (sonst anhängen)?", vbYesNoCancel + vbDefaultButton2, "Rückfrage")
   Case vbYes: Kill hVerz & DMP_Import
   Case vbNo
   Case vbCancel: Exit Sub
  End Select
 End If
' DokuDat = QEnd(ZQuart(Now() - Verspätung))
 DokuDatS = Format(DokuDat, "dd.mm.yyyy")
 
 Call ranam.Open("select * from anamnesebogen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly)
 Call rnam.Open("select * from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly)
 Call rfal.Open("select " & IIf(Not LVobMySQL, "top 1", "") & " * from faelle where pat_id = " & Pat_id & " and bhfb <= " & DatForm(QEnd(ZQuart(Now() - Verspätung))) & " order by bhfb desc, schgr" & IIf(LVobMySQL, " limit 1", ""), DBCn, adOpenStatic, adLockReadOnly)
 Call rform.Open("select " & IIf(Not LVobMySQL, "top 1", "") & " feldinh from formular where pat_id = " & Pat_id & " and Feld = 'Kasse' and zeitpunkt <= " & DatForm(QEnd(ZQuart(Now - Verspätung))) & " and feldinh like '%'" & " order by zeitpunkt desc" & IIf(LVobMySQL, " limit 1", ""), DBCn, adOpenStatic, adLockReadOnly) ' & rfal!VKNr & "%'"  geht nicht gut: VKNr nicht unbedingt aktuell in Faelle (s.Pat_id 51)
' Ermittlung der 'Kasse' aus Rezepten oder Überweisungen oder vorherigen DMP-Dokus usw.
 Dim Kasse$
 If Not rform.BOF Then
  Kasse = rform!FeldInh ' Trim$(Replace$(rform!FeldInh, rfal!VKNr, ""))
 Else
  Set rform = Nothing
  rform.Open "select kurzname, rname from kassenliste where vk = " & rfal!VKNr, DBCn, adOpenStatic, adLockReadOnly
  If Not rform.BOF Then
   Kasse = rform!kurzname
   If InStrB(Kasse, " ") <> 0 Then
    pos = InStr(Kasse, " ")
    Kasse = Left(Kasse, pos - 1)
   End If
  End If
  Kasse = Left$(Left$(Kasse, 25) & Space$(25), 25) & rfal!VKNr
 End If
 Set rform = Nothing
 
 Dim Postleitzahl$
 Postleitzahl = rnam!Plz & " " & rnam!Ort
 Call rform.Open("select " & IIf(Not LVobMySQL, "top 1", "") & " feldinh from formular where pat_id = " & Pat_id & " and Feld = 'KVKGueltig' and zeitpunkt <= " & DatForm(QEnd(ZQuart(Now - Verspätung))) & " and feldinh like '%/%'" & " order by zeitpunkt desc" & IIf(LVobMySQL, " limit 1", ""), DBCn, adOpenStatic, adLockReadOnly)
 If Not rform.BOF Then
  Postleitzahl = Left$(Postleitzahl & Space$(25), 25) & rform!FeldInh
 Else
  Postleitzahl = Left$(Postleitzahl & Space$(30), 30)
 End If
 Set rform = Nothing
 
 Dim Status$
 Call rform.Open("select " & IIf(Not LVobMySQL, "top 1", "") & " feldinh from formular where pat_id = " & Pat_id & " and Feld = 'Status' and zeitpunkt <= " & DatForm(QEnd(ZQuart(Now - Verspätung))) & " and feldinh like '% %'" & " order by zeitpunkt desc" & IIf(LVobMySQL, " limit 1", ""), DBCn, adOpenStatic, adLockReadOnly)
 If Not rform.BOF Then
  Status = rform!FeldInh
 Else
  Status = rfal!KVKs & " " & rfal!KVKserg
 End If
 Set rform = Nothing
 
 ' QAnf(ZQuart(Now() - verspätung))
 Call GetPrRR(ranam, RRsyst, RRdiast)
 If RRsyst > 0 Then
  RRstr = Format$(RRsyst, "000") & Format$(max(RRdiast, 30), "000")
 Else
  RRstr = "{TAB 6}"
 End If
 Call DmPAusgeb(DT, Pat_id, Not obstumm)
 Call BDT.ImportFolderHerricht(hVerz, DMP_Import)
 erg = Dir(BDT.z)
 If LenB(erg) = 0 Then
  Call BDT.BDTKopf
 End If
 BDT.Satzart "6200" ' Falldaten
 BDT.PatID Pat_id
 If Not IsNull(rnam!NVorsatz) Then
  If rnam!NVorsatz <> "" Then
   BDT.NVorsatz rnam!NVorsatz
  End If
 End If
 BDT.Nachname rnam!Nachname
 BDT.Vorname rnam!Vorname
 BDT.Geb rnam!GebDat
 BDT.DAdd "6200", DokuDatS ' Tagesdatum
 BDT.TAdd "6201", "2000"
 BDT.Add "3635TM#889690003"
 BDT.Add "3636TM#641915300"
 Select Case boga
  Case typ2alt, typ1alt, typ2neu, typ1neu
   Dim dmt$, obneu%
   If boga = typ2alt Or boga = typ2neu Then dmt = "2" Else dmt = "1"
   If boga = typ1neu Or boga = typ2neu Then obneu = True Else obneu = 0
   BDT.Add "6295eDMPDM" & dmt
   BDT.Add "6296" & IIf(obneu, "Erst", "Folge") & "-Dokumentation Diabetes mellitus Typ " & dmt
   BDT.Add "6297$\TurboMed\Formulare\Patientenmenue\eDMP Datenerfassung.tmf "
   BDT.FFAdd "ACEHemmerDM" & dmt & "#" & -(Not DT.obACEH)
   BDT.FIAdd "X"
   If dmt = "2" Then
    If (DT.obSonstAD = adja Or DT.obGlit = adja Or DT.obGlucI = adja Or DT.obSHGlin = adja) Then '
     BDT.FFAdd "AntidiabetischDM2#0"
     BDT.FIAdd "X"
    Else ' das verstehe wer will ...
     BDT.FFAdd "AntidiabetischDM2#1"
     BDT.FIAdd "X"
     BDT.FFAdd "AntidiabetischDM2#2"
     BDT.FIAdd "X"
    End If
   End If
   BDT.FFAdd "AntihypertensivDM" & dmt & "#" & -(Not DT.obAntihyp)
   BDT.FIAdd "X"
   BDT.FFAdd "BetablockerDM" & dmt & "#" & -(Not DT.obBetabl)
   BDT.FIAdd "X"
   BDT.FFAdd "Datum"
   BDT.FIAdd Format(DokuDat, "dd.mm.yy")
   BDT.FFAdd "DiabWahrgenommenDM" & dmt & "#3" ' beim letzten Mal keine Schulung empfohlen
   BDT.FIAdd "X"
   BDT.FFAdd "DintervallDM" & dmt & "#0"
   BDT.FIAdd "X"
   BDT.FFAdd "DokuTyp"
   BDT.FIAdd "DM" & dmt & "." & IIf(obneu, "E", "F") & "D"
   Dim EinschrWg%
   Select Case dmt
    Case 1: EinschrWg = 1
    Case 2: EinschrWg = 0
   End Select
   BDT.FFAdd "Einschreibung#" & EinschrWg ' Einschreibung wegen Diabetes mellitus Typ 2
   BDT.FIAdd "X"
   If Not DT.neuDial And Not DT.neuErbl And Not DT.neuAmp And Not DT.neuMI Then
    BDT.FFAdd "EreignisseDM" & dmt & "#4"  ' 0 = Nierenersatztherapie, 1 = Erblindung, 2 = Amputation, 3 = Herzinfarkt, 4 = keine
    BDT.FIAdd "X"
   Else
    If DT.neuDial Then
     BDT.FFAdd "EreignisseDM" & dmt & "#0"
     BDT.FIAdd "X"
    End If
    If DT.neuErbl Then
     BDT.FFAdd "EreignisseDM" & dmt & "#1"
     BDT.FIAdd "X"
    End If
    If DT.neuAmp Then
     BDT.FFAdd "EreignisseDM" & dmt & "#2"
     BDT.FIAdd "X"
    End If
    If DT.neuMI Then
     BDT.FFAdd "EreignisseDM" & dmt & "#3"
     BDT.FIAdd "X"
    End If
   End If
   ErstellDat = Format(DokuDat, "ddmmyyyy")
   For i = 0 To 7
    BDT.FFAdd "ErstellungsdatumDM" & dmt & "#" & i
    BDT.FIAdd Mid(ErstellDat, i + 1, 1)
   Next i
   Dim testvar$
   testvar = Format(Pat_id, "0000000")
   For i = 7 - Len(CStr(Pat_id)) To 6
    BDT.FFAdd "FallNummerDM" & dmt & "#" & i
    BDT.FIAdd Mid(testvar, i + 1, 1)
   Next i
   If Not (DT.FEn(15) Or DT.FEn(1) Or DT.FEn(5) Or DT.FEn(17) Or DT.FEn(2) Or DT.FEn(6) Or DT.FEn(16) Or DT.FEn(3) Or DT.FEn(4)) Then
    BDT.FFAdd "Folgeerkrankung#0"
    BDT.FIAdd "X"
   Else
    If DT.FEn(15) Then ' COPD
     BDT.FFAdd "Folgeerkrankung#1"
     BDT.FIAdd "X"
    End If
    If DT.FEn(1) Then ' Art. Hypertonie
     BDT.FFAdd "Folgeerkrankung#2"
     BDT.FIAdd "X"
    End If
    If DT.FEn(5) Then ' Schlaganfall
     BDT.FFAdd "Folgeerkrankung#3"
     BDT.FIAdd "X"
    End If
    If DT.FEn(17) Then ' Herzinsuffizienz
     BDT.FFAdd "Folgeerkrankung#4"
     BDT.FIAdd "X"
    End If
    If DT.FEn(2) Then ' Fettstoffwechselstörung
     BDT.FFAdd "Folgeerkrankung#6"
     BDT.FIAdd "X"
    End If
    If DT.FEn(6) Then ' AVK
     BDT.FFAdd "Folgeerkrankung#7"
     BDT.FIAdd "X"
    End If
    If DT.FEn(16) Then ' Asthma
     BDT.FFAdd "Folgeerkrankung#8"
     BDT.FIAdd "X"
    End If
    If DT.FEn(3) Or DT.FEn(4) Then ' KHE
     BDT.FFAdd "Folgeerkrankung#9"
     BDT.FIAdd "X"
    End If
   End If
   Select Case DT.fußst
    Case ndok
     BDT.FFAdd "FussstatusDM" & dmt & "#0"
    Case unauff
     BDT.FFAdd "FussstatusDM" & dmt & "#1"
    Case auff
     BDT.FFAdd "FussstatusDM" & dmt & "#2"
   End Select
   BDT.FIAdd "X"
   GewichtS = Format(DT.gewi, "000")
   For i = 3 - Len(CStr(GewichtS)) To 2
    BDT.FFAdd "Gewicht#" & i
    BDT.FIAdd Mid(GewichtS, i + 1, 1)
   Next i
   If dmt = "2" Then
    BDT.FFAdd "GlibenclamidDM2#" & -(Not (DT.obGlib = adja))
    BDT.FIAdd "X"
   End If
   If DT.fußst = auff Then
    BDT.FFAdd "Grad0bisVDM" & dmt & "#" & Left$(DT.mWA, 1) ' Wagner-Stadium, stimmt zufällig überein
    BDT.FIAdd "X"
    ArmSt = Mid$(DT.mWA, 2, 1)
    If Len(ArmSt) = 1 And ArmSt >= "A" And ArmSt <= "D" Then
     BDT.FFAdd "GradAbisDDM" & dmt & "#" & Asc(ArmSt) - 65
     BDT.FIAdd "X"
    End If
   End If
   GrößeS = Format(DT.kgr, "000")
   For i = 3 - Len(CStr(GrößeS)) To 2
    BDT.FFAdd "Groesse#" & i
    BDT.FIAdd Mid(GrößeS, i + 1, 1)
   Next i
   BDT.FFAdd "HMGDM" & dmt & "#" & -(Not DT.obHMG)
   BDT.FIAdd "X"
   HbA1cS = Format(DT.bekHb * 10, "000")
   For i = 3 - Len(CStr(HbA1cS)) To 2
    BDT.FFAdd "Hba1cDM" & dmt & "#" & i
    BDT.FIAdd Mid(HbA1cS, i + 1, 1)
   Next i
   BDT.FFAdd "HyperWahrgenommenDM" & dmt & "#3"
   BDT.FIAdd "X"
   HypoZS = Format(DT.hypoZKK, "00")
   For i = 2 - Len(CStr(HypoZS)) To 1
    BDT.FFAdd "HypoglykaemenDM" & dmt & "#" & i
    BDT.FIAdd Mid(HypoZS, i + 1, 1)
   Next i
   BDT.FFAdd "Ik"
   BDT.FIAdd rfal!IK
   BDT.FFAdd "InsulinDM" & dmt & "#" & -(Not (DT.obIns Or DT.obAnal))
   BDT.FIAdd "X"
   BDT.FFAdd "KVKGueltig"
   BDT.FIAdd "889690003" ' eigene LANR
   BDT.FFAdd "Kasse"
   BDT.FIAdd Kasse
   KreaS = Format(DT.Crea * 100, "0000")
   For i = 4 - Len(CStr(KreaS)) To 3
    BDT.FFAdd "Kreatinin2#" & i
    BDT.FIAdd Mid(KreaS, i + 1, 1)
   Next i
   If dmt = "2" Then
    BDT.FFAdd "MetforminDM2#" & -(Not (DT.obmetf = adja))
    BDT.FIAdd "X"
   End If
   BDT.FFAdd "Nachname"
   BDT.FIAdd rfal!Nachname
   BDT.FFAdd "NetzhautDM" & dmt & "#" & IIf(DT.aug = durchg, "0", IIf(DT.aug = ndurch, "1", "2"))
   BDT.FIAdd "X"
   BDT.FFAdd "Postleitzahl"
   BDT.FIAdd Postleitzahl
   BDT.FFAdd "PulsstatusDM" & dmt & "#" & IIf(DT.puls = ndok, "0", IIf(DT.puls = unauff, "1", "2"))
   BDT.FIAdd "X"
   RRS = Format(DT.RRdiast, "000")
   For i = 0 To 2 ' hier wirklich 0
    BDT.FFAdd "RRd#" & i
    BDT.FIAdd Mid(RRS, i + 1, 1)
   Next i
   RRS = Format(DT.RRsyst, "000")
   For i = 0 To 2 ' hier wirklich 0
    BDT.FFAdd "RRs#" & i
    BDT.FIAdd Mid(RRS, i + 1, 1)
   Next i
   BDT.FFAdd "Raucher#" & -(Not DT.Tabak)
   BDT.FIAdd "X"
   BDT.FFAdd "SEmpfohlenDM" & dmt & "#2"
   BDT.FIAdd "X"
   BDT.FFAdd "SensibilitaetDM" & dmt & "#" & IIf(DT.sens = ndok, "0", IIf(DT.sens = unauff, "1", "2"))
   BDT.FIAdd "X"
   If DT.FEn(7) Then ' Nephropathie
    BDT.FFAdd "SpaetfolgenDM" & dmt & "#0"
    BDT.FIAdd "X"
   End If
   If DT.FEn(11) Then ' Neuropathie
    BDT.FFAdd "SpaetfolgenDM" & dmt & "#1"
    BDT.FIAdd "X"
   End If
   If DT.FEn(9) Then ' Retinopathie
    BDT.FFAdd "SpaetfolgenDM" & dmt & "#2"
    BDT.FIAdd "X"
   End If
   If dmt = "1" Then
    For i = 0 To 1
     BDT.FFAdd "StaHba1cDM1#" & i
     BDT.FIAdd "0"
    Next i
   End If
   KrZS = Format(DT.krZKK, "00")
   For i = 2 - Len(CStr(KrZS)) To 1
    BDT.FFAdd "StationaereDM" & dmt & "#" & i
    BDT.FIAdd Mid(KrZS, i + 1, 1)
   Next i
   BDT.FFAdd "Status"
   BDT.FIAdd Status
   BDT.FFAdd "Strasse"
   BDT.FIAdd rnam!strasse
   ErstellDat = Format(QEnd(ZQuart(QEnd(ZQuart(Now - Verspätung)) + 1)), "ddmmyyyy")
   For i = 0 To 7
    BDT.FFAdd "TerminDM" & dmt & "#" & i
    BDT.FIAdd Mid(ErstellDat, i + 1, 1)
   Next i
   BDT.FFAdd "ThromDM" & dmt & "#" & -(Not DT.obThro)
   BDT.FIAdd "X"
   BDT.FFAdd "UeberweisungDM" & dmt & "#0" ' zum diabetologisch qual Arzt bzw. Einrichtung
   BDT.FIAdd "X"
   BDT.FFAdd "UrinDM" & dmt & "#" & IIf(DT.mau = ndok, "0", IIf(DT.mau = auff, "1", "2"))
   BDT.FIAdd "X"
   BDT.FFAdd "Versichertennr"
   BDT.FIAdd rnam!Versichertennummer
   BDT.FFAdd "Vertragsarztnr"
   BDT.FIAdd "641915300"
   BDT.FFAdd "Vorname"
   BDT.FIAdd rnam!Vorname
   BDT.FFAdd "ZielHba1cDM" & dmt & "#" & IIf(DT.hbEmpf = senken, "1", "0")
   BDT.FIAdd "X"
   BDT.FFAdd "geboren"
   BDT.FIAdd Format(rnam!GebDat, "dd.mm.yy")
   BDT.FFAdd "lblPatient"
   BDT.FIAdd rnam!Vorname & " " & rnam!Nachname & " geb. am: " & Format(rnam!GebDat, "dd.mm.yyyy")
   
 End Select
 Call BDT.Schreib(anhängen:=True)
 Exit Sub
 
 Select Case boga
  
  Case typ2neu
    SendStr = "%{F2}{TAB}+{TAB}"
    SendStr = SendStr & DT.daseit & IIf(DT.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.dspmed, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & "{TAB} {TAB 3} {TAB 2} {TAB 2}"
    For j = 0 To UBound(DT.FEn)
     SendStr = SendStr & IIf(DT.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(DT.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.kgr < 10, "00", IIf(DT.kgr < 100, "0", "")) & DT.kgr
    SendStr = SendStr & IIf(DT.gewi < 10, "00", IIf(DT.gewi < 100, "0", "")) & DT.gewi
'    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(DT.puls = ndok, " ", "") & "{TAB}" & IIf(DT.puls = unauff, " ", "") & "{TAB}" & IIf(DT.puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(DT.sens = ndok, " ", "") & "{TAB}" & IIf(DT.sens = unauff, " ", "") & "{TAB}" & IIf(DT.sens = auff Or DT.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr = SendStr & RRstr ' Blutdruck
    SendStr = SendStr & Format$(DT.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format$(DT.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(DT.fußst = ndok, " ", "") & "{TAB}" & IIf(DT.fußst = unauff, " ", "") & "{TAB}" & IIf(DT.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.mSei = rE Or DT.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.mSei = li Or DT.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left(DT.mWA, 1)
    ArmSt = Mid$(DT.mWA, 2, 1)
    If ArmSt = "" Then
     If WagSt = "0" Then
      If DT.puls = unauff Then ArmSt = "A" Else ArmSt = "C" ' 0 eher nicht entzündet
     ElseIf WagSt > "0" Then
      If DT.puls = unauff Then ArmSt = "B" Else ArmSt = "D"
     Else
      ' WagSt = ""
     End If
    End If
    SendStr = SendStr & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.oap = rE Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.oap = li Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2}" ' Überweisung/ Einweisung
    SendStr = SendStr & IIf(DT.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & Format$(DT.hypoZAn, "00") & "00" ' stat Aufenthalte wegen schwerer Hypos
    If DT.dspmed Then
     SendStr = SendStr & IIf(DT.obGlib = adja, " {TAB 3}", IIf(DT.obGlib = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(DT.obmetf = adja, " {TAB 3}", IIf(DT.obmetf = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(DT.obGlucI = adja, " {TAB 3}", IIf(DT.obGlucI = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(DT.obSHGlin = adja, " {TAB 3}", IIf(DT.obSHGlin = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(DT.obGlit = adja, " {TAB 3}", IIf(DT.obGlit = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr = SendStr & IIf(DT.obIns, IIf(DT.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
     SendStr = SendStr & IIf(DT.obAnal, IIf(DT.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
    End If
    SendStr = SendStr & IIf(DT.obHMG, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.obAntihyp, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.obThro, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}" ' Schulungen
    SendStr = SendStr & IIf(DT.Tabak, " {TAB 2}", "{TAB}{TAB}") ' 'X' steht bei "nein" schon drin
    SendStr = SendStr & IIf(DT.bmi >= 25, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(DT.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.aug = durchg, " {TAB 3}", IIf(DT.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr = SendStr & " {TAB 14} {TAB 2}"
   
   Case typ2alt
    SendStr = "%{F3}{TAB}+{TAB}"
    SendStr1a = "(%N){TAB}+{TAB}  "
    SendStr2 = " {TAB 2} {TAB}"
    For j = 1 To UBound(DT.FEn)
     SendStr2 = SendStr2 & IIf(DT.FEn(j) And Not DT.FE(j), " ", "") & "{TAB}"
    Next j
    SendStr2 = SendStr2 & IIf(DT.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.gewi < 10, "00", IIf(DT.gewi < 100, "0", "")) & DT.gewi
'    sendstr2 = sendstr2 & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr2 = SendStr2 & IIf(DT.puls = ndok, " ", "") & "{TAB}" & IIf(DT.puls = unauff, " ", "") & "{TAB}" & IIf(DT.puls = auff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(DT.sens = ndok, " ", "") & "{TAB}" & IIf(DT.sens = unauff, " ", "") & "{TAB}" & IIf(DT.sens = auff Or DT.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & RRstr ' Blutdruck
    SendStr2 = SendStr2 & Format$(DT.bekHb * 10, "000") & "062"
    SendStr2 = SendStr2 & "000" & Format$(DT.Crea * 10, "000") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.fußst = ndok, " ", "") & "{TAB}" & IIf(DT.fußst = unauff, " ", "") & "{TAB}" & IIf(DT.fußst = auff, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.mSei = rE Or DT.mSei = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.mSei = li Or DT.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left$(DT.mWA, 1)
    ArmSt = Mid$(DT.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If DT.puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr2 = SendStr2 & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.oap = rE Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.oap = li Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 2}" ' Überweisung/ Einweisung
    SendStr2 = SendStr2 & IIf(DT.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & Format$(DT.hypoZAn, "00") & "0" ' stat Aufenthalte wegen schwerer Hypos
    SendStr2 = SendStr2 & IIf(DT.dspmed, " {TAB 2}", "{TAB} {TAB}")
    If DT.dspmed Then
     SendStr2 = SendStr2 & IIf(DT.obGlib = adja, " {TAB 3}", IIf(DT.obGlib = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(DT.obmetf = adja, " {TAB 3}", IIf(DT.obmetf = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(DT.obGlucI = adja, " {TAB 3}", IIf(DT.obGlucI = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(DT.obSHGlin = adja, " {TAB 3}", IIf(DT.obSHGlin = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(DT.obGlit = adja, " {TAB 3}", IIf(DT.obGlit = adki, "{TAB} {TAB} {TAB}", "{TAB} {TAB 2}"))
     SendStr2 = SendStr2 & IIf(DT.obIns, IIf(DT.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
     SendStr2 = SendStr2 & IIf(DT.obAnal, IIf(DT.insz > 2, "{TAB} {TAB 2}", "{TAB 2} {TAB}"), " {TAB 3}")
    Else
     SendStr2 = SendStr2 & "{TAB 21}"
    End If
    SendStr2 = SendStr2 & IIf(DT.obHMG, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.obAntihyp, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.obThro, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 6} {TAB 3} {TAB}" ' Schulungen
    SendStr2 = SendStr2 & IIf(DT.Tabak, " {TAB 2}", "{TAB}{TAB}") ' 'X' steht bei "nein" in der Folgedoku noch nicht drin
    SendStr2 = SendStr2 & IIf(DT.bmi >= 25, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr2 = SendStr2 & IIf(DT.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.aug = durchg, " {TAB 3}", IIf(DT.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr2 = SendStr2 & " {TAB 14} {TAB 2}"

  Case typ1neu ' Typ 1 neu:
    SendStr = "%(MO){TAB}+{TAB}" & DT.daseit & " {TAB 2}"
    SendStr = SendStr & IIf(DT.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.obSchw, " ", "") & "{TAB} {TAB}" ' Schwangerschaft / keine Folgeerkrankungen
    For j = 1 To UBound(DT.FEn)
     SendStr = SendStr & IIf(DT.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr = SendStr & IIf(DT.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.kgr < 10, "00", IIf(DT.kgr < 100, "0", "")) & DT.kgr
    SendStr = SendStr & IIf(DT.gewi < 10, "00", IIf(DT.gewi < 100, "0", "")) & DT.gewi
    SendStr = SendStr & " {TAB 2}" ' altersgemäße körperliche Entwicklung
    SendStr = SendStr & IIf(DT.puls = ndok, " ", "") & "{TAB}" & IIf(DT.puls = unauff, " ", "") & "{TAB}" & IIf(DT.puls = auff, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(DT.sens = ndok, " ", "") & "{TAB}" & IIf(DT.sens = unauff, " ", "") & "{TAB}" & IIf(DT.sens = auff Or DT.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr = SendStr & IIf(DT.fußst = ndok, " ", "") & "{TAB}" & IIf(DT.fußst = unauff, " ", "") & "{TAB}" & IIf(DT.fußst = auff, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.mSei = rE Or DT.mSei = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.mSei = li Or DT.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left(DT.mWA, 1)
    ArmSt = Mid$(DT.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If DT.puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr = SendStr & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.oap = rE Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & IIf(DT.oap = li Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr = SendStr & "{TAB 2}" & RRstr  ' Überweisung ausgestellt und Blutdruck
    SendStr = SendStr & Format$(DT.bekHb * 10, "000") & "062"
    SendStr = SendStr & "000" & Format$(DT.Crea * 10, "000") & "{TAB}"
    SendStr = SendStr & IIf(DT.mau = ndok, " ", "") & "{TAB}" & IIf(DT.mau = auff, " ", "") & "{TAB}" & IIf(DT.mau = unauff, " ", "") & "{TAB 2}"
    SendStr = SendStr & Format$(DT.hypoZAn, "00") & "0000"
    SendStr = SendStr & IIf(DT.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.obIns, IIf(DT.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(DT.obAnal, IIf(DT.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr = SendStr & IIf(DT.obHMG, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(DT.obAntihyp, " {TAB}", "{TAB}")
    SendStr = SendStr & IIf(DT.obThro, " {TAB}", "{TAB}")
    SendStr = SendStr & "{TAB 2} {TAB}" & "{TAB 2} {TAB}"
    SendStr = SendStr & IIf(DT.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr = SendStr & IIf(DT.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr = SendStr & IIf(DT.aug = durchg, " {TAB 3}", IIf(DT.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr = SendStr & " {TAB 16} {TAB 2}"
   
   Case typ1alt
    SendStr = "%(MU){TAB}+{TAB}"
    SendStr1a = "(%N){TAB}+{TAB}  "
    SendStr2 = IIf(DT.dspsy, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.obSchw, " ", "") & "{TAB} {TAB}" ' Schwangerschaft / keine Folgeerkrankungen
    For j = 1 To UBound(DT.FEn)
     SendStr2 = SendStr2 & IIf(DT.FEn(j), " ", "") & "{TAB}"
    Next j
    SendStr2 = SendStr2 & IIf(DT.Tabak, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & " {TAB 2}" ' altersgerechte Entwicklung
    SendStr2 = SendStr2 & IIf(DT.gewi < 10, "00", IIf(DT.gewi < 100, "0", "")) & DT.gewi
    SendStr2 = SendStr2 & IIf(DT.puls = ndok, " ", "") & "{TAB}" & IIf(DT.puls = unauff, " ", "") & "{TAB}" & IIf(DT.puls = auff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(DT.sens = ndok, " ", "") & "{TAB}" & IIf(DT.sens = unauff, " ", "") & "{TAB}" & IIf(DT.sens = auff Or DT.sens = pathdok, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & IIf(DT.fußst = ndok, " ", "") & "{TAB}" & IIf(DT.fußst = unauff, " ", "") & "{TAB}" & IIf(DT.fußst = auff, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.mSei = rE Or DT.mSei = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.mSei = li Or DT.mSei = gleich, " ", "") & "{TAB}"
    WagSt = Left(DT.mWA, 1)
    ArmSt = Mid$(DT.mWA, 2, 1)
    If ArmSt = "" And WagSt <> "" Then If DT.puls = unauff Then ArmSt = "B" Else ArmSt = "D"
    SendStr2 = SendStr2 & IIf(WagSt = "0", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "1", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "2", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "3", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "4", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(WagSt = "5", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "A", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "B", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "C", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(ArmSt = "D", " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.oap = rE Or DT.oap = gleich, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.oap = li, " ", "") & "{TAB}"
    SendStr2 = SendStr2 & "{TAB 2}" & RRstr   ' Überweisung ausgestellt und Blutdruck
    
    SendStr2 = SendStr2 & Format$(DT.bekHb * 10, "000") & "062"
    SendStr2 = SendStr2 & "000" & Format$(DT.Crea * 10, "000") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.mau = ndok, " ", "") & "{TAB}" & IIf(DT.mau = auff, " ", "") & "{TAB}" & IIf(DT.mau = unauff, " ", "") & "{TAB 2}"
    SendStr2 = SendStr2 & Format$(DT.hypoZAn, "0") & "00"
    SendStr2 = SendStr2 & IIf(DT.oblaser, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.obIns, IIf(DT.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.obAnal, IIf(DT.tart <> csii, "{TAB} {TAB 2}", "{TAB 3} "), " {TAB 3}") & "{TAB}"
    SendStr2 = SendStr2 & IIf(DT.obHMG, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & IIf(DT.obAntihyp, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & IIf(DT.obThro, " {TAB}", "{TAB}")
    SendStr2 = SendStr2 & "{TAB 6} {TAB 3} {TAB}" ' Schulungen
    SendStr2 = SendStr2 & IIf(DT.hbEmpf = halten, " {TAB 3}", "{TAB} {TAB 2}")
    SendStr2 = SendStr2 & IIf(DT.rrEmpf = halten, " {TAB 2}", "{TAB} {TAB}")
    SendStr2 = SendStr2 & IIf(DT.aug = durchg, " {TAB 3}", IIf(DT.aug = ndurch, "{TAB} {TAB 2}", "{TAB 2} {TAB}"))
    SendStr2 = SendStr2 & " {TAB 16} {TAB 2}"
    
 End Select
 hnd = FensterHandle("TurboMed")
 If hnd <> 0 Then
  On Error Resume Next
  Err.Clear
  AppActivate "TurboMed", True
  If Err.Number <> 0 Then
   MsgBox "Fehler beim Aktivieren von Turbomed"
   Exit Sub
  End If
  On Error GoTo fehler
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
'  hnd = FensterHandle("Hinweis")
'  If hnd <> 0 Then
'   On Error Resume Next
'   AppActivate "Hinweis", True
'   If Err.Number <> 0 Then
'    SendKeys "%N", True
'    Pause (100)
'    SendKeys "{ESC}", True
'    Pause (100)
'   End If
'   On Error GoTo fehler
'  End If
  SendKeys "{ESC}", True
  Pause (100)
  SendKeys "{ESC}", True
  Pause (100)
  hnd = FensterHandle("TurboMed - `Kartei")
  If hnd <> 0 Then
   On Error Resume Next
   AppActivate "TurboMed - `Kartei", False
   Pause (100)
   SendKeys "{ESC}", True
   Pause (100)
   hnd = FensterHandle("TurboMed - `Fehler")
   If hnd <> 0 Then
    AppActivate "TurboMed - `Fehler", False
    Pause (100)
    SendKeys "{ENTER}", True
    Pause (100)
    SendKeys "{ESC}", True
    Pause (100)
   End If
   On Error GoTo fehler
  End If
'  MsgBox "Fertig!"
'  Exit Sub
'  SendKeys "{TAB}+{TAB}" & "%p" & "{BS}" & "%p" & "{BS}" & "%p" & "{BS}"
  SendKeys "p" & "{BS}", True
'  SendKeys "{ESC}", True
  Pause (100)
'  SendKeys "p", True
'  Pause (100)
  SendKeys Pat_id & "", True
  Pause (100)
  SendKeys "{ENTER}", True
  Pause (1000)
'  SendKeys "{ENTER}", True
'  Pause (100)
'  SendKeys "{ENTER}", True
'  Pause (100)
 Else
  MsgBox "Fehler beim Anwählen des aufgerufenen Turbomed!"
  Exit Sub
 End If
 If SendStr <> "" Then
    SendKeys SendStr, True
    If SendStr1a <> "" Then
     Pause (100)
     SendKeys SendStr1a, True
    End If
    If SendStr2 <> "" Then
     Pause (100)
     SendKeys SendStr2, True
    End If
'   If 1 = 1 Then
    Dim dmpdat$, dmpdatop$
    dmpdat = Format$(min(QEnd(ZQuart(Now - Verspätung)), Now), "dd.mm.yy")
    dmpdatop = Format$(min(QEnd(ZQuart(Now - Verspätung)), Now), "ddmmyyyy")
    Dim DatName$, DNn$
    DatName = "p:\" & DT.Nachname & " " & DT.Vorname & " (PID" & DT.Pat_id & "), DMP-Formular vom " & dmpdat & ".pdf"
    DNn = Dir(DatName)
    If DNn <> "" Then Kill DatName
    DatName = Replace$(Replace$(DatName, "(", "{(}"), ")", "{)}")
    AppActivate "TurboMed", False
    Pause (100)
    SendKeys dmpdatop & "{DEL 8}" & dmpdat
    Pause (1000)
    SendKeys "{TAB}%P", True
    Pause (200)
    SendKeys "%O%(DD)aa{ENTER}", True
    Pause (200)
    SendKeys DatName & "{ENTER}n{ENTER}{F3}{ENTER}", True
'   End If
 End If
 MFG.SetFocus
 On Error Resume Next
 AppActivate "TurboMed", True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in machDMPBogen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' machDMPBogen

Private Sub MFGrefresh()
 Dim i&
 Dim aicd$
 Dim rs As New ADODB.Recordset
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 Me.Hide
 Me.MFG.ColWidth(0) = 690
 Me.MFG.ColWidth(1) = 1185
   rs.Open "select count(*) as ct from diagreihe", DBCn, adOpenStatic, adLockOptimistic
   MFG.Rows = rs!ct + 30
   Set rs = Nothing
   rs.Open "select max(zahl) as mz from (select icd, count(*) as zahl from (select d.icd as icd,diagtext from diagnosen d left join diagreihe r on d.icd = r.icd group by d.icd, diagtext) as innen group by icd) as innen;", DBCn, adOpenStatic, adLockOptimistic
   MFG.Cols = rs!mz + 3
 For i = 2 To Me.MFG.Cols
  Me.MFG.ColWidth(i) = 3100
 Next i
   Set rs = Nothing
   rs.Open "select d.icd as icd,dg1, diagtext from diagnosen d left join diagreihe r on d.icd = r.icd left join diagnoseng1 g1 on r.dg1 = g1.gruppe group by d.icd, diagtext order by g1.rf ,dg2,icd", DBCn, adOpenStatic, adLockOptimistic
   aicd = ""
   MFG.Row = 1
   Do While Not rs.EOF
    If rs!ICD <> aicd Then MFG.Row = MFG.Row + 1: MFG.col = 0: MFG = rs!ICD: MFG.col = GruSp: MFG = rs!dg1: aicd = rs!ICD
    MFG.col = MFG.col + 1
    MFG = rs!DiagText
    rs.Move 1
   Loop
 Me.Show
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MFGrefresh/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MFGrefresh
Public Sub ExcelLesen(Datei$)
 Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
 Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
 Dim rX As New ADOX.Catalog, sql$, ka%, ke%, runde%, angefangen%, obAnfang%, i&
 Dim XCon As New ADODB.Connection
 Dim rEx As New ADODB.Recordset, rs As New ADODB.Recordset
 
 On Error GoTo fehler
 DBCnS = DBCn
 Set XCon = Nothing
 XCon.Open XStra & Datei & XStrb
 Set rX = Nothing
 rX.ActiveConnection = XCon
 rEx.Open "`" & rX.Tables(rX.Tables.Count - 1).Name & "`", XCon ' Hier Excel, nicht obmysql = 0!
' Dim maxsp&
' maxsp = rEx.Fields.Count
 On Error Resume Next
 DBCn.Execute ("drop table laborxls")
 On Error GoTo 0
 DBCn.Execute ("create table laborxls (id integer(10) auto_increment key, patient varchar(255), fehlerart varchar(3000))")
 Do While Not rEx.EOF
  If obAnfang Then
    If IsNull(rEx.Fields(1)) Then Exit Do ' 8.1.08
    DBCn.Execute ("insert into laborxls(patient,fehlerart) values('" & rEx.Fields(0) & "','" & rEx.Fields(1) & "')")
  ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) And Not IsNull(rEx.Fields(2)) Then
   obAnfang = True
  End If
  runde = runde + 1
  rEx.Move 1
 Loop
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ExcelLesen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ExcelLesen
Sub ExcelAusgeb(Datei$)
 Dim rs As New ADODB.Recordset, i&, vorDp$, pos&, rs1 As New ADODB.Recordset, j&, DateiDatStr$, sql$
 Dim Pat_id&, vorPID&, DateiDat As Date, vorFarbe&
 Const NamSp% = 1
 Const ParSp% = 2
 Const WertSp% = 3
 Const EinhSp% = 4
 Const VorWSp1% = 5
 Const VorWSp2% = 6
 Const NBSp% = 7
 Const MedSp% = 8
 Const PIDSp% = 0
 Const vorTage% = 2
 On Error GoTo fehler
 Me.MFG.Cols = 9
 DateiDatStr = Datei
 Do
  pos = InStr(DateiDatStr, "\")
  If pos <> 0 Then DateiDatStr = Mid(DateiDatStr, pos + 1) Else Exit Do
 Loop
 pos = InStr(DateiDatStr, " ")
 If pos <> 0 Then
  j = InStr(pos + 1, DateiDatStr, ".xls")
  i = InStr(pos + 1, DateiDatStr, " ")
  If i <> 0 And i < j Then j = i
  DateiDatStr = Mid(DateiDatStr, pos, j - pos)
  If IsDate(DateiDatStr) Then
   DateiDat = CDate(DateiDatStr)
  End If
 End If
 rs.Open "select count(0) as ct from laborxls", DBCn, adOpenStatic, adLockReadOnly
 Me.MFG.Rows = rs!ct + 1
 Set rs = Nothing
 rs.Open "select patient, fehlerart from laborxls order by patient, fehlerart", DBCn, adOpenStatic, adLockReadOnly
 i = 1
 Me.MFG.TextMatrix(0, NamSp) = "Patient"
 Me.MFG.TextMatrix(0, ParSp) = "Parameter"
 Me.MFG.TextMatrix(0, WertSp) = "Wert"
 Me.MFG.TextMatrix(0, NBSp) = "Normbereich"
 Me.MFG.TextMatrix(0, PIDSp) = "Pat_ID"
 Me.MFG.TextMatrix(0, EinhSp) = "Einheit"
 Me.MFG.TextMatrix(0, VorWSp1) = "Vorwert 1"
 Me.MFG.TextMatrix(0, VorWSp2) = "Vorwert 2"
 Do While Not rs.EOF
  Me.MFG.TextMatrix(i, WertSp) = rs!fehlerart
  pos = InStr(rs!fehlerart, ":")
  If pos <> 0 Then
   vorDp = Left$(rs!fehlerart, pos - 1)
   Set rs1 = Nothing
   rs1.Open "select * from laborparameter where abkü = '" & vorDp & "'", DBCn, adOpenStatic, adLockReadOnly
   If Not rs1.EOF Then
    Me.MFG.TextMatrix(i, ParSp) = rs1!Langtext
    Me.MFG.TextMatrix(i, WertSp) = Mid(rs!fehlerart, pos + 1)
    Me.MFG.Row = i
    Me.MFG.col = NBSp
    Me.MFG.CellAlignment = 0 'flexAlignLeftTop
   End If
  End If
  pos = InStr(Me.MFG.TextMatrix(i, WertSp), "(")
  If pos <> 0 Then
   Me.MFG.TextMatrix(i, NBSp) = Mid$(Me.MFG.TextMatrix(i, WertSp), pos)
   Me.MFG.TextMatrix(i, WertSp) = Left$(Me.MFG.TextMatrix(i, WertSp), pos - 1)
  End If
  pos = InStr(Me.MFG.TextMatrix(i, WertSp), " ")
  If pos <> 0 Then
   Me.MFG.TextMatrix(i, EinhSp) = Mid$(Me.MFG.TextMatrix(i, WertSp), pos + 1)
   Me.MFG.TextMatrix(i, WertSp) = Left$(Me.MFG.TextMatrix(i, WertSp), pos - 1)
  End If
  Me.MFG.TextMatrix(i, NamSp) = rs!patient
  Set rs1 = Nothing
  sql = "from namen where trim(concat(titel,' ',vorname,' ',nachname)) = '" & rs!patient & "'"
  rs1.Open "select count(0) as ct " & sql, DBCn, adOpenStatic, adLockReadOnly
  If rs1!ct = 1 Then
   Set rs1 = Nothing
   rs1.Open "select pat_id, gebdat " & sql, DBCn, adOpenStatic, adLockReadOnly
   Me.MFG.TextMatrix(i, NamSp) = Me.MFG.TextMatrix(i, NamSp) & " (" & Round((Now() - rs1!GebDat) / 365.24, 0) & " a)"
   Pat_id = rs1!Pat_id
   Me.MFG.TextMatrix(i, PIDSp) = Pat_id
   Me.MFG.Row = i
   Me.MFG.col = PIDSp
   If Pat_id <> vorPID Then vorFarbe = IIf(vorFarbe = &HFFFFFF, &HE0E0E0, &HFFFFFF)
   Me.MFG.CellBackColor = vorFarbe
   Me.MFG.col = EinhSp
   Me.MFG.CellBackColor = vorFarbe
   Me.MFG.col = NBSp
   Me.MFG.CellBackColor = vorFarbe
   vorPID = Pat_id
   Set rs1 = Nothing
   rs1.Open "select zeitpunkt,wert from labor1 where pat_id = " & Pat_id & " and abkü = '" & vorDp & "' " & IIf(DateiDat = 0, "", "and zeitpunkt < " & DatForm(DateiDat - vorTage)) & " order by zeitpunkt desc", DBCn, adOpenStatic, adLockReadOnly
   If Not rs1.BOF Then
    Me.MFG.TextMatrix(i, VorWSp1) = rs1!Wert & " (" & Int(rs1!Zeitpunkt) & ")"
    rs1.Move 1
    If Not rs1.EOF Then
     Me.MFG.TextMatrix(i, VorWSp2) = rs1!Wert & " (" & Int(rs1!Zeitpunkt) & ")"
    End If
   End If
   Dim DT As dmptyp, vordat As Date
   Select Case vorDp
    Case "KREA", "KREA02", "KRE02", "CREAT"
     If CDbl(Replace$(Me.MFG.TextMatrix(i, WertSp), ".", ",")) > 1.3 Then
      Call TherAuskunft(Pat_id, 0, DT.insz, vordat, DT.obIns, DT.obAnal, DT.obGlib, DT.obmetf, DT.obGlucI, DT.obSHGlin, DT.obGlit, DT.obSonstAD, DT.obHMG, DT.obAntihyp, DT.obACEH, DT.obBetabl, DT.obThro)
      If DT.obmetf <> 0 Then
       Me.MFG.TextMatrix(i, MedSp) = "Metformin!"
       Me.MFG.Row = i
       Me.MFG.col = MedSp
       Me.MFG.CellBackColor = vbRed
      End If
     End If
   End Select
   Set rs1 = Nothing
   rs1.Open "select zeitpunkt,art from eintraege where art in ('tk','gs') and pat_id = " & Pat_id & " order by zeitpunkt desc limit 1", DBCn, adOpenStatic, adLockReadOnly
   If Not rs1.BOF Then
    Me.MFG.Row = i
'    Me.MFG.ColSel = Me.MFG.Cols - 1
    For j = NamSp To WertSp ' Me.MFG.Cols - 1
     Me.MFG.col = j
     Select Case rs1!Art
      Case "tk"
       Me.MFG.CellBackColor = IIf(j = NamSp, &H4040FF, &HC0C0FF)        ' rötlich
      Case "gs"
       Me.MFG.CellBackColor = IIf(j = NamSp, &HFFFF&, &HC0FFFF)     'gelblich
'      Case Else
'       Stop
     End Select
    Next j
'    Me.MFG.TextMatrix(i, ZPSp) = rs1!Zeitpunkt
'   Else
'    Stop
   End If
'  Else
'   Stop
  End If
  i = i + 1
  rs.Move 1
 Loop
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ExcelAusgeb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ExcelAusgeb

Private Sub Form_Load()
 Dim sql$, i%, erg$, pCol As New Collection, j&, gehezu&
 Me.WindowState = vbMaximized
 MFGTyp = 1
 Li1.Visible = False
 cb1.Visible = False
 gehezu = 1
 Select Case Me.Art
  Case artlab
   Me.Caption = "Pathologische Laborwerte (" & Me.LabDat & ")"
   Me.Command1(0).Caption = "Datei zu &Kothny verschieben"
   Me.Command1(1).Caption = "Datei zu &Schade verschieben"
   Me.Command1(2).Caption = "Datei &Löschen"
   For i = 3 To Me.Command1.Count - 1
    Me.Command1(i).Visible = False
   Next i
   Me.Label1.Visible = False
   Call ExcelLesen(Me.LabDat)
   Call ExcelAusgeb(Me.LabDat)
   Call SizeColumns(MFG, Me, True, 3200)
  Case Else
 On Error GoTo fehler
 erg = Dir(hVerz & DMP_Import)
 If LenB(erg) <> 0 Then
  Open hVerz & DMP_Import For Input As #388
  Do While Not EOF(388)
  Input #388, erg
  If Mid(erg, 4, 4) = "3000" Then
   pCol.Add Mid(erg, 8)
  End If
  Loop
  Close #388
 End If
 Select Case Me.Art
  Case artDiag
   Me.Caption = "Diagnosen sortieren"
   For i = Me.Command1.LBound To Me.Command1.UBound
    On Error Resume Next
    Me.Command1(i).Visible = False
    On Error GoTo fehler
   Next i
   Me.Command1(0).Caption = "&Refresh"
   DBCn.Execute ("insert into diagreihe (icd) SELECT icd FROM diagnosen where not icd in (select icd from diagreihe) group by icd ")
   Call MFGrefresh
   Me.MFG.Visible = False
   Me.MFG.col = 1
   For i = 2 To MFG.Rows
    Me.MFG.Row = i
    If Me.MFG.Text = "" Then
     Me.MFG.TopRow = Me.MFG.Row - 1
     Me.MFG.LeftCol = 1
     Exit For
    End If
   Next i
   Me.MFG.Visible = True
  Case artPat
 Dim rDPat As New ADODB.Recordset, lfdnr&, cbcol&, k&
 Dim rDok As New ADODB.Recordset
' sql = "select n.*, icd, kurzname from namen n left join diagnosen d on n.pat_id = d.pat_id left join lfaelle f on n.pat_id = f.pat_id left join kassenliste k on k.ik = f.ik and k.vk = f.vknr where notiz like '%DMP hier%' and (icd like 'E1%' and not icd like 'E15%' and not icd like 'E16%') and diagsicherheit not in ('Z','A') and not kurzname like 'AOK%' order by n.pat_id"
 sql = "select f.bhfb, notiz, tkz, n.pat_id, n.nachname, n.vorname, icd, kurzname from namen n left join diagnosen d on n.pat_id = d.pat_id left join (select f.pat_id, f.fid, f.schgr, f.bhfb, f.ik, f.vknr, f.quartal from (select * from (select BhFB AS mbhfb, bhfe1, Pat_ID AS pid, quartal from faelle f where bhfb < " & DatForm(QEnd(ZQuart(Now() - Verspätung))) & " order by pid, bhfb desc) l group by pid desc) l left join faelle f on f.pat_id = l.pid and f.bhfb = l.mbhfb group by pat_id) f on n.pat_id = f.pat_id left join anamnesebogen a on n.pat_id = a.pat_id left join kassenliste k on k.ik = f.ik and k.vk = f.vknr where notiz like '%DMP hier%' and (icd like 'E1%' and not icd like 'E15%' and not icd like 'E16%') and diagsicherheit not in ('Z','A') and not isnull (f.bhfb) group by n.pat_id order by n.pat_id"
' sql = "select n.*, icd, kurzname from namen n left join diagnosen d on n.pat_id = d.pat_id left join lfaelle f on n.pat_id = f.pat_id left join kassenliste k on k.ik = f.ik and k.vk = f.vknr where notiz like '%DMP hier%' and (icd like 'E1%' and not icd like 'E15%' and not icd like 'E16%') and diagsicherheit not in ('Z','A') and kurzname like 'AOK%' order by n.pat_id"
 rDPat.Open "select count(*) from (" & sql & ") as innen", DBCn, adOpenStatic, adLockReadOnly
 With MFG
 .Rows = rDPat.Fields(0) + 2
 .Cols = 21
 .Row = 1
 Set rDPat = Nothing
 lfdnr = 0
 rDPat.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rDPat.EOF
   .col = 0
'   If rDPat!Pat_id = 2901 Then Stop
   If obhierdmp(rDPat!Notiz) And rDPat!Tkz = 0 Then cbcol = rot Else cbcol = vbWhite
   .Text = lfdnr
   .CellBackColor = vbWhite
   For k = 1 To pCol.Count
    If pCol.Item(k) = CStr(rDPat!Pat_id) Then
     .CellBackColor = rot
     If cbcol = rot Then gehezu = .Row
     Exit For
    End If
   Next k
   If GruSp = 0 Then GruSp = .col
'   .CellBackColor = cbcol
   lfdnr = lfdnr + 1
   .col = .col + 1
   If PIDSp = 0 Then PIDSp = .col
   .Text = rDPat!Pat_id
   .CellBackColor = cbcol
   .col = .col + 1
   If NachNameSp = 0 Then NachNameSp = .col
   .Text = rDPat!Nachname
   .CellBackColor = cbcol
   .col = .col + 1
   .Text = rDPat!Vorname
   .CellBackColor = cbcol
   .col = .col + 1
   .Text = IIf(IsNull(rDPat!kurzname), "", rDPat!kurzname)
   .CellBackColor = cbcol
   .col = .col + 1
   .Text = Replace$(Replace$(Replace$(rDPat!Notiz, "DMP", ""), &HA, " "), &HD, " ")
   .CellBackColor = cbcol
   .col = .col + 1
   If ICDSp = 0 Then ICDSp = .col
   .Text = rDPat!ICD
   .CellBackColor = IIf(rDPat!ICD Like "E11*", cbcol, vbYellow)
   .col = .col + 1
   .Text = ZQuart(rDPat!BhFB)
   If ZQuart(rDPat!BhFB) = ZQuart(Now() - Verspätung) Then .CellBackColor = rot Else .CellBackColor = vbWhite
   rDok.Open "select * from dmpreihe where pat_id = " & rDPat!Pat_id & " order by exportiert desc;", DBCn, adOpenDynamic, adLockReadOnly
   Dim begCol%
   If Not rDok.BOF Then
    begCol = .col + 1
    If Now() - rDok!DokuDatum > 120 Then ' 21. des übernä Monats
     begCol = begCol + 1
    End If
    For j = begCol To .Cols - 1
     If rDok.EOF Then Exit For
     .col = j
     If VorDokuSp = 0 Then VorDokuSp = .col
     .Text = rDok!Art & " " & rDok!DokuDatum
     If rDok!Art = "ED" Then .CellBackColor = 16767449
     rDok.Move 1
    Next j
   End If
   Set rDok = Nothing
  .Row = .Row + 1
  rDPat.Move 1
 Loop
 .TextMatrix(0, GruSp) = "Lfdnr"
 .TextMatrix(0, PIDSp) = "Pat_id"
 .TextMatrix(0, NachNameSp) = "Nachname"
 .TextMatrix(0, 3) = "Vorname"
 .TextMatrix(0, 4) = "Kasse"
 .TextMatrix(0, 5) = "DMP"
 .TextMatrix(0, ICDSp) = "ICD"
 .TextMatrix(0, 7) = "lFall"
 .TextMatrix(0, VorDokuSp) = "lDoku"
 Call SizeColumns(MFG, Me)
 .col = NachNameSp
 .Row = gehezu
 End With
 Call mfg_entercell
 Me.Command1(0).Caption = "&Erstelle"
 Me.Command1(1).Caption = "In T&m anzeigen"
 Me.Command1(2).Caption = "&In Tm fertigst."
 Me.Command1(3).Caption = "&Suchen"
 Me.Command1(4).Caption = "&Weitersuchen"
 Me.Command1(5).Caption = "&Nächster Gleichartiger"
' Me.Command1(6).Caption = "&Typ 1 ED"
' Me.Command1(7).Caption = "T&yp 1 FD"
' Me.Command1(8).Caption = "Ty&p 2 ED"
' Me.Command1(9).Caption = "Typ 2 &FD"
 Me.Command1(6).Caption = "Erstelle &zu Pat_id" ' Dokubeliebig
 Me.Command1(7).Caption = "&Fertigst zu Pat_id"
 Me.Command1(8).Caption = "Anleitung"
 Me.Command1(9).Caption = "&HTML bewah"
 Me.Command1(10).Caption = "R&ückg"
 Me.Command1(11).Caption = "Auff&risch" ' Auffrisch
 Me.Command1(12).Caption = "&AlleFS" 'alleFS
 Me.Command1(13).Caption = "a&ktFS" ' aktFS
 For i = 14 To Me.Command1.Count - 1
  Me.Command1(i).Visible = False
 Next i
 Me.Label1.Caption = "'&DMP'=Reiternr.:"
 Me.Text1 = 12 ' DMP-Spalte, DMP-Tab
 On Error GoTo fehler
End Select ' art
End Select ' me.art

 Dim ctl As Control, lctl As Control
 For i = Me.Command1.LBound To Me.Command1.UBound + 2 ' 0 to 10
  If i <= Me.Command1.UBound Then
   Set ctl = Me.Command1(i)
  ElseIf i = Me.Command1.UBound + 1 Then
   Set ctl = Me.Label1
  Else
   Set ctl = Me.Text1
  End If
  If i <= Me.Command1.UBound Then
   ctl.Width = Me.TextWidth(ctl.Caption) + 200
  ElseIf i = Me.Command1.UBound + 1 Then
   ctl.Width = Me.TextWidth(ctl.Caption)
  Else
   ctl.Width = 300
  End If
  If i = 0 Then
   ctl.Left = 100
  Else
   ctl.Left = lctl.Left + lctl.Width + 50
  End If
  Set lctl = ctl
 Next i

Me.KeyPreview = True
Me.Visible = True
Me.MFG.SetFocus
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Form_LOAD (PatListe)/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Sub
End Sub ' Form_load

Private Sub Form_Resize()
 On Error Resume Next
 MFG.Height = Me.Height - MFG.Top - 400
 Call Me.TopAusricht
End Sub

Public Sub MFG_Click()
 Dim ctrl As Control, Fd0$, Fd1$, i%, rs As New ADODB.Recordset
 If cRow(MFGTyp) = cRowSel(MFGTyp) Then
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
 End If
 cCol(MFGTyp) = Me.MFG.col
 obtb = -1
 Select Case Art
  Case artDiag
   If MFG.col = GruSp Then
     obtb = 0
   End If
  Case artPat
 End Select
 If obtb = -1 Then
  cRow(MFGTyp) = Me.MFG.Row
  cRowSel(MFGTyp) = Me.MFG.RowSel
 Else
  If obtb = 0 Then
   Set ctrl = Me.Li1
   ctrl.Clear
   rs.Open "select * from diagnoseng1 order by rf", DBCn, adOpenStatic, adLockOptimistic
   Do While Not rs.EOF
    ctrl.AddItem (rs!gruppe)
    rs.Move 1
   Loop
   If AnfCode <> 0 Then
    For i = 0 To Li1.ListCount - 1
     If Li1.List(i) <> "" Then
      If Asc(LCase(Left(Li1.List(i), 1))) = AnfCode Then
       Cstumm = True
       Li1.ListIndex = i
       Cstumm = False
       Exit For
      End If
     End If
    Next i
    AnfCode = 0
   End If
  Else
  End If
  ctrl.Left = Me.MFG.CellLeft + Me.MFG.Left
  ctrl.Top = Me.MFG.CellTop + Me.MFG.Top
  Cstumm = True
  ctrl.Text = Me.MFG.Text
  Cstumm = False
  ctrl.Visible = True
  ctrl.Height = ctrl.ListCount * 255
  Call MFG_leavecell
  ctrl.SetFocus
  On Error Resume Next
  ctrl.SelStart = Len(ctrl.Text)
 End If
End Sub

Private Sub mfg_entercell()
 If noenter = 0 Then
  If fgespei(MFGTyp) = 0 Then
   altFarbe(MFGTyp) = Me.MFG.CellBackColor
   Me.MFG.CellBackColor = vbYellow
   fgespei(MFGTyp) = -1
  End If
 End If
End Sub ' mfg_entercell

Private Sub MFG_leavecell()
 If noenter = 0 Then
  If altFarbe(MFGTyp) <> 0 Then ' am Anfang
   Me.MFG.CellBackColor = IIf(Me.MFG.Row = 0, vbActiveBorder, altFarbe(MFGTyp))
  End If
'  If MFG.CellBackColor = 0 Then
  fgespei(MFGTyp) = 0
 End If
End Sub ' MFG_leavecell
Private Sub MFG_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub ' MFG_KeyDown
Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub
Private Sub Command1_KeyDown(index As Integer, keyCode As Integer, Shift As Integer)
 Call Key(keyCode, Shift, Me, Me.hlese.MyDB.Name)
End Sub

