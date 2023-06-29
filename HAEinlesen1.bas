Attribute VB_Name = "HAEinlesen"
Option Explicit
Const HADBName$ = "haerzte"
Dim KVDate As Date
Dim HACn As New ADODB.Connection, rs As New ADODB.Recordset, rAF&

Enum UAbTyp
 uabName = 0
 NLrt
' Fachrichtung
' Straße
' Ort
' Tel
' Fax
' Mail
' BSArt
' Sprechzeiten
' LANRBSNR
 zusatzbezeichnung
 weiterbildung
 genehmigung
 leistungsumfang
 vertragsangebot
 fremdsprache
 ÄrztederPraxis
End Enum

Type AbschTyp
 art As Integer
 Bis As Long
End Type

Private Type arzttyp
 idarzt As Long
 Nachname As String
 Vorname As String
 Namenszusatz As String
 titel_id As Long
' fachrichtung_id AS Long
 Lanr As Long
 nlart_id As Long
 AktZeit As Date
End Type

Private Type BSTyp
 idbs As Long
 Straße As New CString
 Hausnr As New CString
 plz As String
 Ort_id As Long
 BSNR As String
 obNBS As Integer
 obAng As Integer
 bsart_id As Long
 sprechzeiten_id As Long
 Rollst As Integer
 AktZeit As Date
 name As String
End Type
Const ArztSuchS$ = "KVB_Arztsuche*.pdf"

Public Function doalleKVDateien()
 Dim i&, erg&, db$
 Dim SL As New SortierListe, SD As SortierDatei, Fil$
 Call Lese.ProgStart
 '#Const obadobe = 1
 #If obadobe = 1 Then
 Fil = Dir(vVerz & ArztSuchS) ' "v:\KVB_Arztsuche*.pdf"
 Do While LenB(Fil) <> 0
  Set SD = New SortierDatei
  Set SD.File = FSO.GetFile(vVerz & Fil)
  SL.sCAdd SD
  Fil = Dir
 Loop
 erg = MsgBox("Sollen wirklich alle " & SL.COUNT & " Hausarztdaten aus '" & vVerz & "' neu eingelesen werden?", vbYesNo)
 If erg = vbYes Then
  db = HADBName & "_neu"
  Verbinde db
'  IF HALöschen THEN
   If True Then
   For i = 1 To SL.COUNT
    dodoHausärzteEinlesen SL.Item(i).File.path, False, db
   Next i
  End If
 End If
 #Else
 MsgBox "geht nur mit adobe"
 #End If
End Function ' alleKVDateien

Public Function lebetest()
 Dim rs As ADODB.Recordset
 Dim lebed As Date
 Open "p:\namausg.txt" For Output As #205
 Call Lese.ProgStart
 myFrag rs, "select gesname(pat_id) Nam, n.* from namenlb n", adOpenStatic
 If Not rs.BOF Then
  Do While Not rs.EOF
   rs.MoveNext
   lebed = lebe(rs!Pat_id)
   Dim lbeh As Date
   lbeh = rs!lbeh
   If rs!lbeh <> lebed Then
    Debug.Print rs!Pat_id, lbeh, lebe(rs!Pat_id)
    Print #205, rs!Pat_id & " " & lbeh & " " & lebe(rs!Pat_id)
   Else
'    Print #205, rs!Pat_id & " (stimmt)"
'    Debug.Print rs!Pat_id, "(stimmt)"
   End If
   DoEvents
  Loop
 End If
 Debug.Print "Fertig!"
 Print #205, "Fertig!"
 Close #205
End Function

Function TLösch(Tb$)
' HACn.Execute "TRUNCATE `" & Tb & "`", rAF
 myEFrag "DELETE FROM `" & Tb & "`", rAF, HACn
 If rAF <> 0 Then lies.Ausgeb rAF & " Datensätze aus `" & Tb & "` gelöscht", True
' ON Error Resume Next
 myEFrag "ALTER TABLE `" & Tb & "` auto_increment=1", rAF, HACn
End Function ' TLösch

Function HALöschen%()
  Dim erg&
  erg = MsgBox("Sollen wirklich alle Tabellen der Datenbank `" & CurDB(HACn) & "` der Verbindung:" & vbCrLf & HACn & vbCrLf & "gelöscht werden?", vbYesNo, "Sicherheitsrückfrage")
  If erg = vbNo Then
   HALöschen = 0
  Else
' mit ALTER TABLE geht irgendwie nicht 30.11.09
   TLösch "bs"
   TLösch "bsart"
   TLösch "arzt"
   TLösch "fachrichtung"
   TLösch "ort"
   TLösch "bsart"
   TLösch "fremdsprache"
   TLösch "genehmigung"
   TLösch "leistungsumfang"
   TLösch "nlart"
   TLösch "sprechzeiten"
   TLösch "titel"
   TLösch "vertragsangebot"
   TLösch "zusatzbezeichnung"
   TLösch "weiterbildung"
   TLösch "arzt_has_bs" ' "ALTER TABLE `arzt_has_bs` auto_increment=0"
   TLösch "arzt_has_fachrichtung"
   TLösch "arzt_has_fremdsprache"
   TLösch "arzt_has_genehmigung"
   TLösch "arzt_has_leistungsumfang"
   TLösch "arzt_has_vertragsangebot"
   TLösch "arzt_has_weiterbildung"
   TLösch "arzt_has_zusatzbezeichnung"
   TLösch "bs_has_genehmigung"
   TLösch "bs_has_leistungsumfang"
   TLösch "bs_has_vertragsangebot"
   TLösch "fax"
   TLösch "mail"
   TLösch "tel"
   HALöschen = True
  End If
End Function ' HALöschen

Public Function getHAPDF$(Optional obStumm%)
  On Error GoTo fehler
  With Lese.CommonDialogLese
   .DialogTitle = "PDF-Datei mit Hausarztdaten von KVB aussuchen"
   .initDir = vVerz
' "c:\v\KVB_Arztsuche_-20090919-221246.pdf"
   .DefaultExt = "pdf"
   .Filename = ArztSuchS
   Dim jFil$, jDat As Date, erg$
   erg = Dir(.initDir & .Filename)
   Do While LenB(erg) <> 0
    If erg <> "." And erg <> ".." Then
     If FileDateTime(.initDir & erg) > jDat Then
'      jDat = FileDateTime(.InitDir & erg)
      jDat = ZiehDat(.initDir & erg)
      jFil = erg
     End If
    End If
    erg = Dir
   Loop
   .Filename = .initDir & jFil
'   .Flags = .Flags AND NOT FileOpenConstants.cdlOFNFileMustExist
   If Not obStumm Then .ShowOpen
   getHAPDF = .Filename
  End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getHAPDF/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getHAPDF
Public Function ZiehDat(Datei$) As Date
      Dim p1&, pz&, bst$, zds$
      p1 = Len(Datei)
      pz = 0
      Do
       bst = Mid$(Datei, p1, 1)
       If bst = "-" Then pz = pz + 1
       If pz = 2 Or p1 = 1 Then Exit Do
       p1 = p1 - 1
      Loop
      zds = Mid$(Datei, p1 + 7, 2) & "." & Mid$(Datei, p1 + 5, 2) & "." & Mid$(Datei, p1 + 1, 4) & " " & Mid$(Datei, p1 + 10, 2) & ":" & Mid$(Datei, p1 + 12, 2) & ":" & Mid$(Datei, p1 + 14, 2)
      If IsDate(zds) Then
       ZiehDat = CDate(zds)
      Else
       ZiehDat = FileDateTime(Datei)
      End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZiehDat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZiehDat

Public Sub doHausärzteEinlesen()
 Dim Filename$
 Filename = getHAPDF(True)
 dodoHausärzteEinlesen Filename, True
End Sub ' doHausärzteEinlesen

Sub Verbinde(Optional db$)
 Dim i%, ErrNr&
 On Error GoTo fehler
 If Not HACn Is Nothing Then Set HACn = Nothing
 HACnS = DBCnS
 HACn.Open HACnS ' DBCn.ConnectionString
 If LenB(db) = 0 Then db = HADBName
 For i = 1 To 2
  If i = 1 Then On Error Resume Next
  myEFrag "USE `" & db & "`", , HACn, True, ErrNr
  If ErrNr <> 0 Then
#If Modul_MachDBhaerzte_bas_dabei Then
   doMach_haerzte db, DBCn, MachDatenbank.GetServr(DBCn) ' keine "`" hier
#End If
  End If
 Next i
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Verbinde/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Verbinde

Sub dodoHausärzteEinlesen(Filename$, obRückfrage%, Optional db$)
' Const HADBName$ = "hausaerzte"
'#Const obadobe = 1
#If obadobe = 1 Then
 Dim PDTextS As Acrobat.CAcroPDTextSELECT
 Dim Result&, PDDoc As Acrobat.CAcroPDDoc
 Dim PDPage As Acrobat.CAcroPDPage
 Dim PDHili As Acrobat.CAcroHiliteList
 Dim B_erg As Boolean, erg$
 Dim PDFStream$
' Strings Falls Text von Anfang ausgelesen werden soll (bzw bis zum Ende)
' Dim FromFirst$, ToLast$
 Dim Str As New CString, NTL&
 Dim i%, j&, jj&, splitt$(), s2$(), pid$
 Dim Uhrzeit$, maxP&
 
' FromFirst = Chr(169) & Chr(170) & Chr(172)
' ToLast = Chr(163) & Chr(165) & Chr(164)
' ON Error Resume Next
 On Error GoTo fehler
 Set PDDoc = CreateObject("AcroExch.pdDoc")
' ON Error GoTo fehler
 KVDate = ZiehDat(Filename)
 Result = PDDoc.Open(Filename)
 If Not Result Then
  MsgBox "Fehler beim Öffnen von: " & Filename, vbCritical, "Hausärzte einlesen"
  Exit Sub
 Else
  maxP = PDDoc.GetNumPages
  If obRückfrage Then
   erg = InputBox("Wie viele Seiten von '" & Filename & "' sollen eingelesen werden?", "Rückfrage", maxP)
   If IsNumeric(erg) Then
    maxP = CDbl(erg)
    If maxP = 0 Then Exit Sub
   Else
    Exit Sub
   End If
  End If
 End If

#Const hadbneu = 0
#If hadbneu Then
' IF LenB(DB) <> 0 THEN ' ist auch schon in verbinde
  Call doMach_haerzte(db, DBCn, DBVerb.Cpt)
' END IF
#End If
' Filename = Environ("userprofile") & "\Desktop\" & "KVB_Arztsuche_-20090905-045301.pdf"
' Nehme die Erste Seite - Index 0
 Do
  Set PDPage = PDDoc.AcquirePage(j)
  ' Erzeuge ein Highlight Objekt und weise ihm 2000 Elemente bei (keine Grenzprobleme)
  Set PDHili = CreateObject("AcroExch.HiliteList")
  B_erg = PDHili.Add(0, 32000)
  ' Erzeuge eine Textauswahl aus dem gesamten Text
  On Error Resume Next
  Set PDTextS = PDPage.CreatePageHilite(PDHili)
  If Err.Number <> 0 Then Exit Do
  On Error GoTo fehler
  ' Hole Anzahl der "Textblöcke"
  NTL = PDTextS.GetNumText
  ' Gebe den Text der Textauswahl zurück
  For i = 0 To NTL - 1
   Str.Append PDTextS.GetText(i)
  Next i
  DoEvents
  Str.Append vbCrLf
  j = j + 1
  lies.Ausgeb "Hausärzte einlesen: " & Filename & ": Lese S.: " & j, 0
' die ersten 30 Seiten einlesen
  If j > maxP Then Exit Do
 Loop
' in jeder Zeile von splitt steht eine Zeile aus der PDF-Datei
 SplitNeu Str.Value, vbCrLf, splitt, "(", ")"
 
 Verbinde db
 
 Dim Absch() As AbschTyp, AbZ&
 AbZ = 0
 ReDim Absch(0)
 Dim aktab&, LANRZeile&, aktName$ ' ob die letzte Zeile eine Facharztzeile war -> dann nicht zählen
 For aktab = LBound(splitt) To UBound(splitt)
  Select Case splitt(aktab)
   Case "Fachärztin für Psychosomatische Medizin und", "Facharzt für Psychosomatische Medizin und"
    splitt(aktab) = splitt(aktab) & " " & splitt(aktab + 1)
    splitt(aktab + 1) = vNS
  End Select
'  IF aktab = 46637 THEN Stop
'  IF aktab = 1199 THEN Stop
  Select Case splitt(aktab)
' eine der folgenden Ausdrücke steht hinter dem Namen, wodurch dieser als solches klassifizierbar ist
' 23.4.10: hier dürfen auch Facharztbezeichnungen aufgeführt werden, die sonst erst als übernächstes kommen,
   Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin", "Psychotherapeutisch tätige Ärztin", "Psychotherapeutisch tätiger Arzt", "Kinder- und Jugendlichen-Psychotherapeutin", "Kinder- und Jugendlichen-Psychotherapeut", "Fachärztin für Psychiatrie und Psychotherapie", "Facharzt für Psychiatrie und Psychotherapie", "Facharzt für Neurologie", "Facharzt für Nervenheilkunde", "Fachärztin für Neurologie", "Fachärztin für Nervenheilkunde", "Fachärztin für Psychosomatische Medizin und Psychotherapie", "Facharzt für Psychosomatische Medizin und Psychotherapie", "Schwerpunkt Kinder-Kardiologie", "Facharzt für Kinder- und Jugendmedizin (Schwerpunkt Neonatologie)", "Fachärztin für Neurochirurgie" ' , "Abrechnung Mammographie-Screening Abrechnung Mammographie-Screening"
   ' "Fachärztin für Neurochirurgie" = Ulrike Metz 23.4.10
    If LANRZeile <> 0 Then
     GoSub bisfestleg
     GoSub neuerAbschnitt
    End If
   Case vNS ' Nach Zusammenzug von "Fachärztin für Psychosomatische Medizin und Psychotherapie"
'    Stop
   Case Else
    If splitt(aktab) = "Ermächtigte Einrichtung" Or splitt(aktab) = "Ermächtigtes Krankenhaus" Then
      LANRZeile = aktab
      For jj = 0 To 50
       If splitt(aktab - jj) Like "##### *" Then
        Absch(AbZ).Bis = aktab - jj - 3
' Abteilung für Hämatologie und Abteilung für Hämatologie und, internistische
' Onkologie
        If Absch(AbZ).Bis > 0 Then
         If InStr(12, splitt(Absch(AbZ).Bis), Left$(splitt(Absch(AbZ).Bis), 12)) <> 0 Then
          Absch(AbZ).Bis = Absch(AbZ).Bis - 1
         End If
        End If
'        Open "c:\lanrzeilen.txt" For Append AS #317
'        Print #317, splitt(LANRZeile)
'        Close #317
        Exit For
       End If
      Next jj
      GoSub neuerAbschnitt
      Absch(AbZ).art = 1

'     ElseIf splitt(aktab) = "Ermächtigte Einrichtung" THEN
    ElseIf LANRZeile = 0 Then
'     IF Left$(splitt(aktab), 4) = "LANR" THEN
     If splitt(aktab) = "Zusatzbezeichnungen:" Or splitt(aktab) = "Leistungsumfang:" Or splitt(aktab) = "Genehmigungen:" Or splitt(aktab) = "Besondere Vertragsangebote:" Then
      LANRZeile = aktab
     ElseIf Left$(splitt(aktab), 4) = "LANR" Then
      LANRZeile = aktab
'      IF splitt(aktab + 1) = "758369301 698013700" THEN Stop
     End If
' wenn LANRZeile <> 0, darf normal kein Ort mehr kommen, weil der schon vor LANR kommt.
' Kommt doch noch später einer, so handelt es sich um einen neuen Arzt ohne Niederlassungsart -> Gerald Beier
    Else ' IF LANRZeile <> 0 THEN
     If aktName = "Dr.med. Christiane Beier" Then
      If IsNumeric(Left$(splitt(aktab), 1)) Then ' weil "like" so lang brauchen soll
       If splitt(aktab) Like "##### *" And Not splitt(aktab) Like "* * * *" And Not InStrB(splitt(aktab), ",") <> 0 And Not InStrB(splitt(aktab), ".") <> 0 And Not InStrB(splitt(aktab), " EBM") <> 0 And Not InStrB(splitt(aktab), " und ") <> 0 And Not InStrB(splitt(aktab), "Komplex") <> 0 Then ' PLZ mit Ort
'        MsgBox "vermutlich Fehlerhafte Verarbeitung der Zeile " & aktab & ": " & vbCrLf & splitt(aktab)
        GoSub bisfestleg
        GoSub neuerAbschnitt
       End If
      End If
     End If
    End If ' LANRZeile = True / else
  End Select
 Next aktab
' GoSub neuerAbschnitt
 Absch(AbZ).Bis = UBound(splitt) - 1
 Const AbschnDat$ = pVerz & "AbschnDat.txt"
 Open AbschnDat For Output As #340
 For aktab = 0 To AbZ
  Dim ausgj&
  If aktab = 0 Then ausgj = 0 Else ausgj = Absch(aktab - 1).Bis + 1
  If Left$(splitt(ausgj), 8) = "Hinweis:" Then ausgj = ausgj + 1
  Print #340, aktab & " " & Absch(aktab).Bis + 1 & " " & splitt(ausgj)
 Next
 Close #340
 zeigan AbschnDat
 
 For aktab = 0 To AbZ
'  IF aktab = 41 THEN Stop
  Call proTeilnehmer(splitt, Absch, aktab, HACn)
 Next aktab
 
 If AbZ > 1000 Then
  lies.Ausgeb "Korrigiere Geschlechter ...", 0
  myEFrag "UPDATE `arzt` b LEFT JOIN (SELECT DISTINCT vorname FROM `arzt` a LEFT JOIN `nlart` ON a.nlart_id = idnlart LEFT JOIN `arzt_has_fachrichtung` ahf ON a.idarzt = ahf.arzt_id LEFT JOIN `fachrichtung` f ON idfachrichtung = fachrichtung_id WHERE fachrichtung LIKE '%ärztin%' OR fachrichtung LIKE '%eutin%' OR niederlassungsart LIKE '%eutin%') a ON b.vorname = a.vorname SET b.obweibl=1 WHERE NOT ISNULL(a.vorname)", rAF, HACn
 End If
 lies.Ausgeb (AbZ - 1) & " Ärzte aus '" & Filename & "' eingelesen; " & rAF & " davon auf weiblich geändert", 1
' Dim Datum As Date
' SplitNeu splitt(0), " ", s2
' For j = 0 To UBound(s2)
'  IF InStrB(s2(j), "-") > 0 THEN
'   s2 = Split(s2(j), "-")
'   Datum = CDate(s2(0))
'   Exit For
'  END IF
' Next j
 Result = PDDoc.Close
 Set PDPage = Nothing
 Set PDHili = Nothing
 Set PDTextS = Nothing
 Set PDDoc = Nothing
 Exit Sub
 
neuerAbschnitt: ' Abz weiterzählen, absch(abz).bis festlegen, aktName festlegen
     If AbZ = UBound(Absch) Then
      ReDim Preserve Absch(UBound(Absch) + 500)
     End If
     AbZ = AbZ + 1
     Return
bisfestleg:
'     IF AbZ > 88 THEN Debug.Print "Abz:", AbZ: Stop
     If aktab = 1 Then
      Absch(AbZ).Bis = aktab - 2
     ElseIf InStrB(splitt(aktab - 1), " ") = 0 Then ' wenn der Name sich über zwei Zeilen zieht, dann ist in der zweiten Zeile evtl. kein Leerzeichen mehr; Angerpointner
      Absch(AbZ).Bis = aktab - 3
      splitt(aktab - 2) = splitt(aktab - 2) & " " & splitt(aktab - 1)
      splitt(aktab - 1) = vNS
'     ElseIf splitt(aktab) = "Abrechnung Mammographie-Screening Abrechnung Mammographie-Screening" THEN
'      Absch(AbZ).Bis = aktab - 1
     Else
      Absch(AbZ).Bis = aktab - 2
     End If
'     aktab = aktab + 1 ' sonst wird "Psychotherapie" nicht zusammengezogen, s. Becker-Jakubaß
     LANRZeile = 0
     aktName = splitt(aktab - 1)
     Return
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HausärzteEinlesen_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
#End If ' obAdobe
End Sub ' HausärzteEinlesen_Click

Sub proTeilnehmer(ByRef splitt$(), ByRef Absch() As AbschTyp, aktab&, ByRef HACn As ADODB.Connection)
  Dim Uab As UAbTyp, Arzt As arzttyp, BS As BSTyp, ort As New CString, SZ As New CString, s2$()
  Dim tel$(), telz&, fax$(), faxz&, mail() As New CString, mailz&
  Dim OrtZeile&, LANRZeile&, szzz&  ' SprechzeitenZeilenZahl
  Dim fren&() ' Fachrichtungen des aktuellen Teilnehmers
'  ReDim fren&(0)
  Dim j&, jj&, jjj&, szz&, buch$, pos&, p2&
  Dim idIns&, idIns1&, idIns2& ' idIns1 = bs, idIns2 = arzt
  On Error GoTo fehler
' alle Zeilen des aktuellen Teilnehmers
  Uab = uabName
  Dim ausgj&
  If aktab = 0 Then ausgj = 0 Else ausgj = Absch(aktab - 1).Bis + 1
  If Left$(splitt(ausgj), 8) = "Hinweis:" Then ausgj = ausgj + 1
  Debug.Print splitt(ausgj)
  
  For j = ausgj To Absch(aktab).Bis
   Select Case Uab
    Case uabName
     If Absch(aktab).art = 0 Then ' Wenn Arzt und nicht ermächtigte Einrichtung
         SplitNeu splitt(j), " ", s2
         If UBound(s2) > 1 Then
          If s2(UBound(s2) - 1) = "Al" Then
           s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
           ReDim Preserve s2(UBound(s2) - 1)
          End If
         End If
    '     IF s2(UBound(s2)) = "von" THEN
    '      Arzt.Namenszusatz = s2(UBound(s2))
    ''      s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
    '      ReDim Preserve s2(UBound(s2) - 1)
    '     END IF
    '     IF s2(UBound(s2)) = "Freiherr von" THEN
    '      Arzt.Namenszusatz = s2(UBound(s2))
    ''      s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
    '      ReDim Preserve s2(UBound(s2) - 1)
    '     END IF
          For jj = UBound(s2) To 1 Step -1
           Select Case s2(jj)
            Case "Freiherr", "Freifrau", "von", "van", "van den", "v.", "de", "den", "vom", "zur", "genannt", "jun.", "sen."
             Arzt.Namenszusatz = s2(jj) & IIf(LenB(Arzt.Namenszusatz) = 0, vNS, " ") & Arzt.Namenszusatz
             For jjj = jj + 1 To UBound(s2)
              s2(jjj - 1) = s2(jjj)
             Next jjj
             ReDim Preserve s2(UBound(s2) - 1)
           End Select
          Next jj
'         IF s2(UBound(s2)) = "32123," THEN Stop
         Arzt.Nachname = s2(UBound(s2))
         If UBound(s2) = 1 Then
          Arzt.Vorname = s2(0)
         Else
          Select Case s2(0)
           Case "Afssaneh", "Alex", "Ali", "Amir", "Andreas", "Angela", "Anna", "Anne", "Berthold", "Birgit", "Carl", "Caroline", "Christian", "Christoph", "Claus", "Detlef", "Dieter", "Edith", "Elisabeth", "Eugen", "Evelyn", "Frank", "Franz", "Frida", "Gabriele", "Gisela", "Gerald", "Hamdi", "Hans", "Harriet", "Harry", "Heide", "Heiner", "Heinz", "Helga", "Hermina", "Ilse", "Ingeborg", "Irmgard", "Jhonny", "Joachim", "Jörn", "Johannes", "Johannus", "John", "Jorge", "Jorj", "Kai", "Kambiz", "Karin", "Karl", "Kerstin", "Klaus", "Lara", "Manfred", "Mara", "Maria", "Martina", "MartinH.", "Mirjam", "Mitra", "Oliver", "Peter", "Petra", "Rebecca", "Robert", "Rolf", "Roser", "Ruth", "Shervin", "Sigrid", "Silke", "Stefan", "Tahereh", "Thomas", "Thorsten", "Trude", "Ursula", "Ute", "Veit", "Vladimiro", "Werner", "Wilhelm", "Wolf", "Wolfgang", "Z."
            s2(1) = s2(0) & " " & s2(1)
            s2(0) = vNS
           Case "Kinder-u.Jugendlichen-Psychotherapeutin"  ' Fehler in Datenbank: Läuft auch als Titel
            s2(0) = vNS
          End Select
          Arzt.Vorname = s2(1)
          For jj = 2 To UBound(s2) - 1
           Arzt.Vorname = Arzt.Vorname & " " & s2(jj)
          Next jj
           Do
            pos = InStr(Arzt.Vorname, ". ")
            If pos = 0 Then Exit Do
            buch = Mid$(Arzt.Vorname, pos - 1, 1)
            ' "Thomas O. Cook"
            If (buch >= "A" And buch <= "Z") Or buch = "Ä" Or buch = "Ö" Or buch = "Ü" Then Exit Do
            s2(0) = s2(0) & Left$(Arzt.Vorname, pos)
            Arzt.Vorname = Mid$(Arzt.Vorname, pos + 2)
           Loop
          pos = InStr(Arzt.Vorname, ").")
          If pos <> 0 Then
           s2(0) = s2(0) & Left$(Arzt.Vorname, pos)
           Arzt.Vorname = Trim$(Mid$(Arzt.Vorname, pos + 2))
          End If
          pos = InStr(Arzt.Vorname, ")")
          If pos <> 0 Then
           s2(0) = s2(0) & Left$(Arzt.Vorname, pos)
           Arzt.Vorname = Trim$(Mid$(Arzt.Vorname, pos + 1))
          End If
          pos = InStr(Arzt.Vorname, "/")
          If pos <> 0 Then
           p2 = InStr(pos, Arzt.Vorname, " ")
           If p2 <> 0 Then
            s2(0) = s2(0) & Left$(Arzt.Vorname, p2 - 1)
            Arzt.Vorname = Mid$(Arzt.Vorname, p2 + 1)
           End If
          End If
          pos = 0
          Dim Wort$
          Wort = "Ancona"
          If pos = 0 Then
           Wort = "Bologna"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Budapest"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Bukarest"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Debrecen"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Dipl.-Psychologe"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Dipl.Päd.Univ"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Diplom-Lehrer"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Diplom-Psychologe"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "en Médecine"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Florenz"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Fünfkirchen"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Guadalupe"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "III"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "in Psychologie"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Kaschau"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Klausenburg"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Madjid"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Medicina y Cirugia"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "medicinae universae"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Medizinisch"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "M.E.S."
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Neumarkt"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "of B. Adm. Dr.med."
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Olmütz"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Perugia"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          pos = InStr(Arzt.Vorname, Wort)
          If pos = 0 Then
           Wort = "Prag"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Pressburg"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Rom"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Siena"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Szeged"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Tel-Aviv"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Temeschburg"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos = 0 Then
           Wort = "Zagreb"
           pos = InStr(Arzt.Vorname, Wort)
          End If
          If pos <> 0 Then
           s2(0) = s2(0) & Left$(Arzt.Vorname, pos + Len(Wort))
           Arzt.Vorname = Trim$(Mid$(Arzt.Vorname, pos + Len(Wort) + 1))
          End If
          
          pos = 0
          pos = InStr(Arzt.Vorname, "Beier Perlasberger")
          If pos = 0 Then pos = InStr(Arzt.Vorname, "Theisen")
          If pos <> 0 Then
           Arzt.Nachname = Mid$(Arzt.Vorname, pos) & " " & Arzt.Nachname
           Arzt.Vorname = Trim$(Left$(Arzt.Vorname, pos - 1))
          End If
          Arzt.titel_id = indIns(HACn, HACnS, "titel", "titel", s2(0), "idtitel")
         End If ' rs.bof
         lies.Ausgeb "Bearbeite: " & Arzt.Nachname & ", " & Arzt.Vorname, 0
     Else ' absch(aktab).art=0 -> 1
      BS.name = splitt(j)
      For jj = 10 To Len(BS.name) ' doppelten Anfangsteil entfernen
       If Left$(BS.name, jj - 1) = Mid$(BS.name, jj + 1, jj - 1) Then
        BS.name = Mid$(BS.name, jj + 1)
        Exit For
       End If
      Next jj
      lies.Ausgeb "Bearbeite: " & BS.name, 0
     End If
' Ende Case uabName
    Case NLrt
' die Zeile, in der PLZ und Ort stehen, enthält als erstes eine Ziffer
     For jj = j To Absch(aktab).Bis  ' j+1 gestrichen wegen Gerald Beier
      If IsNumeric(Left$(splitt(jj), 5)) And Len(splitt(jj)) > 6 Then OrtZeile = jj: Exit For
     Next jj
     If OrtZeile >= Absch(aktab).Bis Then MsgBox "Ortszeile > " & Absch(aktab).Bis: Stop
'     IF OrtZeile = 1890 THEN Stop
' in der Zeile nach dem Namen könnte, muß aber nicht, eine Niederlassungsart stehen
     Select Case splitt(j)
      Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin"
       Arzt.nlart_id = indIns(HACn, HACnS, "nlart", "niederlassungsart", splitt(j), "idnlart")
       j = j + 1
     End Select
' über die Zeile getrennte Fachrichtungen zusammenführen
     For jj = j To OrtZeile - 2
      jjj = 0
      While (InStrB(splitt(jj), "(") <> 0 And InStrB(splitt(jj), ")") = 0) Or Right$(splitt(jj), 2) = "/-"
       jjj = jjj + 1
       If Right$(splitt(jj), 1) <> " " Then splitt(jj) = splitt(jj) & " "
       splitt(jj) = splitt(jj) & splitt(jj + jjj)
       splitt(jj + jjj) = vNS
      Wend
     Next jj
' Teilgebiete und Schwerpunkte gesondert aufnehmen
     Dim FrSep$(1), ergstr$(), ez&
     FrSep(0) = "Schwerpunkt"
     FrSep(1) = "Teilgebiet"
     ReDim fren(1)
     For jj = j To OrtZeile - 2
      If LenB(splitt(jj)) <> 0 Then
       ez = SplitNeuArr(splitt(jj), FrSep, ergstr, True)
       For jjj = 0 To ez - 1
        buch = Right$(ergstr(jjj), 1)
        If buch = "(" Or buch = ")" Or buch = "," Then
         ergstr(jjj) = Left$(ergstr(jjj), Len(ergstr(jjj)) - 1)
        End If
        Dim l1 As New CString
        l1 = ergstr(jjj)
        l1.LCase
        If l1.Instr("Ärztin") <> 0 Or l1.Instr("ärztin") <> 0 Or l1.Instr("arzt") <> 0 Or l1.Instr("olog") <> 0 Or l1.Instr("medizin") <> 0 Or l1.Instr("ologie") <> 0 Or l1.Instr("psych") <> 0 Or l1.Instr("schwerpunkt") <> 0 Or l1.Instr("teilgebiet") <> 0 Or l1.Left(12) = "schwerpunkt " Then ' Ä wird mit lcase scheinbar nicht zu ä
         If l1.Instr("mvz") = 0 And l1.Instr("versorgungszentrum") = 0 And l1.Instr("abteilung") = 0 And l1.Instr("klinik") = 0 And l1.Instr("ambulanz") = 0 And l1.Instr("gmbh") = 0 And l1.Instr("institut") <> 0 Then
          GoTo stimmt
         End If
        End If
        If BS.name <> vNS Then BS.name = BS.name & " " & splitt(jj) Else BS.name = splitt(jj)
        If l1.Value Like "* #*" Or l1.Value Like "#*" Or IsNumeric(l1.Value) Or Len(l1.Value) < 3 Or Left$(l1.Value, 4) = "str." Or l1.Value Like "*str.*" Or InStrB(l1.Value, ",") <> 0 Then
        Else
         Select Case ergstr(jjj)
'         Case "Berufsausübungsgem. / MVZ-München, 80799", "Einsteinstr. 130", "München, Heßstr. 22", "Ofenthaler Weg 20"
          Case "München", "Erl", "Ansbach", "Hasenkopf", "Nbg", "Windsheim", "Nürnberg", "Kollegen", "Marktheidenfeld-Michelrieth", "Partenkirchen/Auenst", "Rosenheimer Str.", "Roßhaupter-96", "Steenbeek/Groneberg", "Wasserburg", "Wessels-Str.", "Wichtermann", "ÜÖ BAG, Dres.med. Neher/Hilber & Kollegen", "Aurnhammer-Str.", "Bad Gögging", "Bahnhofspl.", "Gemünden-Langenprozelten", "Günzburger Str.41", "Hohenwarter Str", "Kalbskopf-Str", "Preuschwitz.Str.101"
            Debug.Print "proTeiln. ergstr(jjj): " & ergstr(jjj)
          Case Else
stimmt:
           fren(UBound(fren)) = indIns(HACn, HACnS, "fachrichtung", "Fachrichtung", ergstr(jjj), "idFachrichtung")
           ReDim Preserve fren(UBound(fren) + 1)
         End Select
        End If
       Next jjj
      End If
     Next jj
     
     jj = 1
     Do
      buch = Mid$(splitt(OrtZeile - 1), jj, 1)
      If IsNumeric(buch) Then
       BS.Hausnr = Mid$(splitt(OrtZeile - 1), jj)
       Exit Do
      End If
      jj = jj + 1
      If jj > Len(splitt(OrtZeile - 1)) Then Exit Do
     Loop
     BS.Straße = Trim$(Left$(splitt(OrtZeile - 1), jj - 1))
     pos = InStr(BS.Straße, "Perlasberger")
     If pos <> 0 Then
      BS.Straße = Mid$(BS.Straße, pos)
     End If
     SplitNeu splitt(OrtZeile), " ", s2
     If Not IsNumeric(s2(0)) Then ' Leerzeichen vergessen
      ReDim s2(1)
      For jjj = 1 To Len(splitt(OrtZeile))
       If IsNumeric(Mid$(splitt(OrtZeile), jjj, 1)) Then
        s2(0) = s2(0) & Mid$(splitt(OrtZeile), jjj, 1)
       Else
        Exit For
       End If
      Next jjj
      s2(1) = Mid$(splitt(OrtZeile), Len(s2(0)) + 1)
     End If
     BS.plz = s2(0)
     If UBound(s2) > 0 Then
      ort = s2(1)
      For jj = 2 To UBound(s2)
       ort.Append " "
       ort.Append s2(jj)
      Next jj
      BS.Ort_id = indIns(HACn, HACnS, "ort", "Ort", ort.Value, "idOrt")
     End If
'     SELECT CASE splitt(j)
'      Case "Facharzt", "Hausarzt" ' danach immer Fachrichtung
'      Case "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin"
'       ' manchmal allein, manchmal vor dem nächsten
'      Case "Kinder- und Jugendlichen-Psychotherapeutin", "Kinder- und Jugendlichen-Psychotherapeut"
'       ' manchmal allein stehend, manchmal nach dem letzten, danach immer Straße
'      Case "Fachärztin für Psychiatrie und Psychotherapie", "Facharzt für Psychiatrie und Psychotherapie"
'       ' manchmal nach Facharzt, manchmal allein, manchmal vor dem nächsten
'      Case "Psychotherapeutisch tätige Ärztin", "Psychotherapeutisch tätiger Arzt"
'       ' manchmal allein, manchmal nach dem letzten
'       Uab = Uab + 1
'      Case Else
'     END SELECT
     j = OrtZeile + 1
     telz = -1
     ReDim tel(0)
     Do While Left$(splitt(j), 5) = "Tel.:"
      telz = telz + 1
      If telz > UBound(tel) Then ReDim Preserve tel(telz)
      tel(telz) = Mid$(splitt(j), 7)
      j = j + 1
     Loop
     faxz = -1
     ReDim fax(0)
     Do While Left$(splitt(j), 5) = "Fax.:"
      faxz = faxz + 1
      If faxz > UBound(fax) Then ReDim Preserve fax(faxz)
      fax(faxz) = Mid$(splitt(j), 7)
      j = j + 1
     Loop
     For jj = j + 1 To Absch(aktab).Bis
      If splitt(jj) = "Sprechzeiten:" Then
       szz = jj
       Exit For
      End If
     Next jj
     If szz = 0 Then
' kam bisher nicht vor
      MsgBox "Keine Sprechzeiten bei Arzt " & Arzt.Nachname & ", " & Arzt.Vorname
      szz = Absch(aktab).Bis
     End If
     If splitt(j) = "Angestellt in MVZ und überörtliche" Then
      splitt(j + 1) = splitt(j) & " " & splitt(j + 1)
      splitt(j) = vNS
      j = j + 1
     End If
     mailz = -1
     ReDim mail(0)
     For jj = j To szz - 2
      If Left$(splitt(jj), 8) = "E-Mail.:" Then ' in einer Zeile stand: "E-Mail.:", dann nä Zeile
       mailz = mailz + 1
       If mailz > UBound(mail) Then ReDim Preserve mail(mailz)
       mail(mailz).Append Mid$(splitt(jj), 10)
      ElseIf mailz <> -1 Then
       mail(mailz).Append splitt(jj)
      End If
     Next jj
     Dim BSArt$
     BSArt = splitt(szz - 1)
'     IF BSArt = "Angestellt in MVZ" THEN Stop
     If Left$(BSArt, 14) = "Angestellt in " Then BS.obAng = 1: BSArt = Mid$(BSArt, 15)
     BS.bsart_id = indIns(HACn, HACnS, "bsart", "BSArt", BSArt, "idbsart")
     j = szz
     For jj = j + 1 To Absch(aktab).Bis
      If splitt(jj) = "LANR: BSNR:" Then LANRZeile = jj: BS.obNBS = False: Exit For
      If splitt(jj) = "LANR: NBSNR:" Then LANRZeile = jj: BS.obNBS = 1: Exit For
      If splitt(jj) = "BSNR:" Then
        LANRZeile = jj
        BS.obNBS = False
        Exit For
      End If
      If splitt(jj) = "Zusatzbezeichnungen:" Or splitt(jj) = "Genehmigungen:" Or splitt(jj) = "Leistungsumfang:" Or splitt(jj) = "Besondere Vertragsangebote:" Then
       LANRZeile = jj
       BS.obNBS = False
       Exit For
      End If
     Next jj
     If LANRZeile = 0 Then
      MsgBox "Fehler in proTeilnehmer: LANRZeile = 0"
      Stop
     End If
     szzz = (LANRZeile - j - 1) * 0.5
     SZ.Clear
     For jj = j + 1 To j + szzz + 1
      If jj + szzz >= LANRZeile Then Exit For
      SZ.Append splitt(jj)
      If jj + szzz >= LANRZeile Or szzz = 0 Then Exit For
      SZ.Append " "
      SZ.Append splitt(jj + szzz)
      If jj < j + 1 + szzz Then SZ.Append ", "
     Next jj
'     IF Left$(SZ.Value, 6) <> "Montag" AND Left$(SZ.Value, 8) <> "Dienstag" AND Left$(SZ.Value, 8) <> "Mittwoch" AND Left$(SZ.Value, 10) <> "Donnerstag" AND Left$(SZ.Value, 7) <> "Freitag" AND SZ.Value <> "keine Angaben" THEN Stop
     BS.sprechzeiten_id = indIns(HACn, HACnS, "sprechzeiten", "Sprechzeiten", SZ.Value, "idsprechzeiten")
'     IF UBound(s2) <> 1 THEN Stop
     If Absch(aktab).art = 0 Then
      
      SplitNeu splitt(LANRZeile + 1), " ", s2
      If IsNumeric(s2(0)) Then
       If IsNumeric(s2(0)) Then Arzt.Lanr = s2(0)
'      IF BS.obNBS = 1 THEN BS.NBSNR = s2(1): BS.BSNR = 0 ELSE BS.NBSNR = 0: BS.BSNR = s2(1)
       If UBound(s2) = 1 Then
        BS.BSNR = s2(1)
       Else ' Schricke hat 2 Lanr, diese erscheinen wie auch die BSNR in einer Extra-Zeile; 2. LANR ignorieren wir
        p2 = 1
        Do
         If Not IsNumeric(splitt(LANRZeile + p2)) Then Exit Do
         p2 = p2 + 1
        Loop
        BS.BSNR = splitt(LANRZeile + p2 - 1)
       End If
      Else
       BS.BSNR = 0
      End If
     Else
      BS.BSNR = splitt(LANRZeile + 1)
     End If
     Dim KennZeile&(zusatzbezeichnung To ÄrztederPraxis)
     Dim ZB() As New CString, Wb() As New CString, Geneh() As New CString, LeiU() As New CString, VertA() As New CString, fS() As New CString, ÄP() As New CString
     Dim ZBz&, WBz&, Genehz&, VertAz&, FSz&, ÄPz&, LeiUz&, überspringen&
     For jj = LANRZeile To Absch(aktab).Bis
'      IF LEFT(splitt(jj), 6) = "Email:" THEN KennZeile(Email) = jj
      Select Case splitt(jj)
       Case "Zusatzbezeichnungen:"
         KennZeile(zusatzbezeichnung) = jj: Uab = zusatzbezeichnung
       Case "Weiterbildungen:":            KennZeile(weiterbildung) = jj: Uab = weiterbildung
       Case "Genehmigungen:":              KennZeile(genehmigung) = jj: Uab = genehmigung
       Case "Leistungsumfang:":            KennZeile(leistungsumfang) = jj: Uab = leistungsumfang
       Case "Besondere Vertragsangebote:": KennZeile(vertragsangebot) = jj: Uab = vertragsangebot
       Case "Fremdsprachen:":              KennZeile(fremdsprache) = jj: Uab = fremdsprache
       Case "Ärzte der Praxis:":           KennZeile(ÄrztederPraxis) = jj: Uab = ÄrztederPraxis
       Case "Rollstuhlgerechte Praxis":    BS.Rollst = 1
       Case Else
        überspringen = 0
        Do While jj <= Absch(aktab).Bis And ((Left$(splitt(jj + 1 + überspringen), 4) = "Arzt" And KennZeile(genehmigung) <> 0 And KennZeile(vertragsangebot) = 0) Or _
        Right$(splitt(jj), 5) = "nicht" Or Right$(splitt(jj), 3) = "und" Or Right$(splitt(jj), 2) = "u." Or Right$(splitt(jj), 1) = ";" Or Left$(splitt(jj + 1 + überspringen), 4) = "bzw." Or Right$(splitt(jj), 2) = "f." Or Right$(splitt(jj), 3) = "mit" Or Left$(splitt(jj + 1 + überspringen), 1) = "(" Or Left$(splitt(jj + 1 + überspringen), 1) = "/" Or Right$(splitt(jj), 1) = "/" Or Right$(splitt(jj), 1) = "-" Or Right$(splitt(jj), 1) = "," Or Right$(splitt(jj), 5) = "sches" Or Left$(splitt(jj + 1 + überspringen), 3) = "SGB" Or Left$(splitt(jj + 1 + überspringen), 14) = "sonographisch)" Or Left$(splitt(jj + 1 + überspringen), 8) = "mammogr." Or Right$(splitt(jj), 5) = "spez." Or Left$(splitt(jj + 1 + überspringen), 3) = "IPM" Or Left$(splitt(jj + 1 + überspringen), 18) = "Kinder/Jugendliche" Or Left$(splitt(jj + 1 + überspringen), 12) = "Operationen)" Or Right$(splitt(jj), 6) = "in der" Or splitt(jj) = "DMP Brustkrebs teilnehmender Facharzt")
'         IF Left$(splitt(jj + 1 + überspringen), 5) = "Arzt" THEN Stop
         überspringen = überspringen + 1
'         IF Uab <> ÄrztederPraxis THEN Stop
         splitt(jj) = splitt(jj) & " " & splitt(jj + überspringen)
        Loop
        Select Case Uab
         Case zusatzbezeichnung
          ZBz = ZBz + 1
          If ZBz = 1 Then ReDim ZB(60) Else If ZBz > UBound(ZB) Then ReDim Preserve ZB(UBound(ZB) + 60)
          ZB(ZBz) = splitt(jj)
         Case weiterbildung
          WBz = WBz + 1
          If WBz = 1 Then ReDim Wb(60) Else If WBz > UBound(Wb) Then ReDim Preserve Wb(UBound(Wb) + 60)
          Wb(WBz) = splitt(jj)
         Case genehmigung
          Genehz = Genehz + 1
          If Genehz = 1 Then ReDim Geneh(60) Else If Genehz > UBound(Geneh) Then ReDim Preserve Geneh(UBound(Geneh) + 60)
          Geneh(Genehz) = splitt(jj)
         Case leistungsumfang
          LeiUz = LeiUz + 1
          If LeiUz = 1 Then ReDim LeiU(60) Else If LeiUz > UBound(LeiU) Then ReDim Preserve LeiU(UBound(LeiU) + 60)
          LeiU(LeiUz) = splitt(jj)
         Case vertragsangebot
          VertAz = VertAz + 1
          If VertAz = 1 Then ReDim VertA(60) Else If VertAz > UBound(VertA) Then ReDim Preserve VertA(UBound(VertA) + 60)
          VertA(VertAz) = splitt(jj)
         Case fremdsprache
' 23.1.10: beim Seitenumbruch kann auch nach der Fremdsprache noch ein Teil des Textes der genehmigten Ziffern kommen: Dr. Peter Bernius
          If splitt(jj) Like "*#*" Or splitt(jj) Like "* * *" Or InStrB(splitt(jj), "ologie") <> 0 Or splitt(jj) = "wurden" Then
           Uab = Uab + 1
          Else
           FSz = FSz + 1
           If FSz = 1 Then ReDim fS(60) Else If FSz > UBound(fS) Then ReDim Preserve fS(UBound(fS) + 60)
           fS(FSz) = splitt(jj)
          End If
         Case ÄrztederPraxis
          If splitt(jj) Like "*#*" Or splitt(jj) Like "* * *" Then
           Uab = Uab + 1
          Else
           ÄPz = ÄPz + 1
           If ÄPz = 1 Then ReDim ÄP(60) Else If ÄPz > UBound(ÄP) Then ReDim Preserve ÄP(UBound(ÄP) + 60)
           ÄP(ÄPz) = splitt(jj)
          End If
        End Select
      End Select
      jj = jj + überspringen
     Next jj
     If Absch(aktab).art = 0 Then ' Wenn keine ermächtigte Einrichtung
      If Arzt.titel_id = 0 Then Arzt.titel_id = indIns(HACn, HACnS, "titel", "titel", vNS, "idtitel")
      If Arzt.nlart_id = 0 Then Arzt.nlart_id = indIns(HACn, HACnS, "nlart", "niederlassungsart", vNS, "idnlart")
      Set rs = Nothing
      myFrag rs, "SELECT `idarzt` FROM `arzt` WHERE `LANR` = " & Arzt.Lanr, adOpenStatic, HACn, adLockReadOnly
      If rs.EOF Then
       Set rs = Nothing
       InsKorr HACn, HACnS, "INSERT INTO `arzt`(`Nachname`,`Vorname`,`titel_id`,`LANR`,`nlart_id`,`Namenszusatz`,`aktzeit`,`seit`) VALUES('" & doUmwfSQL(Arzt.Nachname, LVobMySQL) & "','" & doUmwfSQL(Arzt.Vorname, LVobMySQL) & "'," & Arzt.titel_id & "," & Arzt.Lanr & "," & Arzt.nlart_id & ",'" & Arzt.Namenszusatz & "'," & DatFor_k(KVDate) & "," & DatFor_k(KVDate) & ")", rAF
       Set rs = myEFrag("SELECT last_insert_id()", , HACn)
      Else
       myEFrag "UPDATE `arzt` SET `titel_id` = " & Arzt.titel_id & ", `nlart_id` = " & Arzt.nlart_id & ", `Namenszusatz` = '" & Arzt.Namenszusatz & "', `aktzeit` = " & DatFor_k(KVDate) & ", `Nachname` = '" & doUmwfSQL(Arzt.Nachname, LVobMySQL) & "', `Vorname` = '" & doUmwfSQL(Arzt.Vorname, LVobMySQL) & "' WHERE LANR = " & Arzt.Lanr, rAF, HACn
       If rAF > 1 Then
        MsgBox "raf = " & rAF & " beim Einfügen des Arztes: " & Arzt.Nachname & ", " & Arzt.Vorname & " (" & Arzt.Lanr & ")"
       End If
      End If
      idIns2 = rs.Fields(0)
'      HACn.BeginTrans
'      HACn.CommitTrans
'      HACn.Execute "begin"
      myEFrag "DELETE FROM `arzt_has_leistungsumfang` WHERE `arzt_id` = " & idIns2, rAF, HACn
'      Debug.Print "proTeiln.: " & rAF & " Sätze aus arzt_has_leistungsumfang zu arzt_id: " & idIns2 & " gelöscht"
'      HACn.CommitTrans
'      HACn.Execute "commit"
      For jj = 1 To LeiUz
'leineu:
       idIns = indIns(HACn, HACnS, "leistungsumfang", "leistungsumfang", LeiU(jj).Value, "idleistungsumfang")
'       ON Error Resume Next
       InsKorr HACn, HACnS, "INSERT INTO `arzt_has_leistungsumfang`(`arzt_id`,`leistungsumfang_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
'       IF rAF <> 1 THEN Stop
'       GoTo leineu
'       ON Error GoTo fehler
      Next jj
      myEFrag "DELETE FROM `arzt_has_zusatzbezeichnung` WHERE `arzt_id` = " & idIns2, rAF, HACn
      For jj = 1 To ZBz
       idIns = indIns(HACn, HACnS, "zusatzbezeichnung", "Zusatzbezeichnung", ZB(jj).Value, "idZusatzbezeichnung")
'       ON Error Resume Next
       InsKorr HACn, HACnS, "INSERT INTO `arzt_has_zusatzbezeichnung`(`arzt_id`,`Zusatzbezeichnung_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
'       ON Error GoTo fehler
      Next jj
      myEFrag "DELETE FROM `arzt_has_weiterbildung` WHERE `arzt_id` = " & idIns2, rAF, HACn
      For jj = 1 To WBz
       idIns = indIns(HACn, HACnS, "weiterbildung", "Weiterbildung", Wb(jj).Value, "idWeiterbildung")
       InsKorr HACn, HACnS, "INSERT INTO `arzt_has_weiterbildung`(`arzt_id`,`Weiterbildung_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
      Next jj
      myEFrag "DELETE FROM `arzt_has_genehmigung` WHERE `arzt_id` = " & idIns2, rAF, HACn
      Dim dop%
      Dim jjjj%
      For jj = 1 To Genehz
       dop = 0
       For jjjj = 1 To jj - 1
        If Geneh(jjjj) = Geneh(jj) Then
         dop = True
         Exit For
        End If
       Next jjjj
       If Not dop Then
        idIns = indIns(HACn, HACnS, "genehmigung", "genehmigung", Geneh(jj).Value, "idgenehmigung")
        InsKorr HACn, HACnS, "INSERT INTO `arzt_has_genehmigung`(`arzt_id`,`genehmigung_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
       End If
      Next jj
'erneut:
      myEFrag "DELETE FROM `arzt_has_vertragsangebot` WHERE `arzt_id` = " & idIns2, rAF, HACn
      For jj = 1 To VertAz
       dop = 0
'       IF jj = 20 THEN Stop
'       IF VertA(jj).Value = "Franz-Josef-Strauß-Allee 11" THEN Stop
' bei "Franz-Josef-Strauß-Allee 11" / "Franz-Josef-Strauss-Allee 11" versagt es
       For jjjj = 1 To jj - 1
        If VertA(jjjj) = VertA(jj) Then
         dop = True
         Exit For
        End If
       Next jjjj
       If Not dop Then
        idIns = indIns(HACn, HACnS, "vertragsangebot", "vertragsangebot", VertA(jj).Value, "idvertragsangebot")
'        GoTo erneut
'        ON Error Resume Next
        InsKorr HACn, HACnS, "INSERT INTO `arzt_has_vertragsangebot`(`arzt_id`,`vertragsangebot_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
'        ON Error GoTo fehler
       End If
      Next jj
      myEFrag "DELETE FROM `arzt_has_fremdsprache` WHERE `arzt_id` = " & idIns2, rAF, HACn
      For jj = 1 To FSz
       dop = 0
       For jjjj = 1 To jj - 1
        If fS(jjjj) = fS(jj) Then
         dop = True
         Exit For
        End If
       Next jjjj
       If Not dop Then
        If LenB(fS(jj).Value) <> 0 Then
         idIns = indIns(HACn, HACnS, "fremdsprache", "fremdsprache", fS(jj).Value, "idfremdsprache")
         InsKorr HACn, HACnS, "INSERT INTO `arzt_has_fremdsprache`(`arzt_id`,`fremdsprache_id`) VALUES(" & idIns2 & "," & idIns & ")", rAF
        End If
       End If
      Next jj
      myEFrag "DELETE FROM `arzt_has_fachrichtung` WHERE `arzt_id` = " & idIns2, rAF, HACn
      For jj = 1 To UBound(fren) - 1
       ' Fachrichtungen können offenbar versehentlich auch doppelt genannt werden (-> Eckhard Rudolf, hat 2 x Unfallchirurgie und 2 x Viszeralchirurgie)
       Set rs = Nothing
       myFrag rs, "SELECT arzt_id FROM `arzt_has_fachrichtung` WHERE arzt_id = " & idIns2 & " AND fachrichtung_id = " & fren(jj), adOpenStatic, HACn, adLockReadOnly
       If rs.BOF Then
        InsKorr HACn, HACnS, "INSERT INTO `arzt_has_fachrichtung`(`arzt_id`,`fachrichtung_id`) VALUES(" & idIns2 & "," & fren(jj) & ")", rAF
       End If
      Next jj
      myEFrag "COMMIT", , HACn
'      HACn.CommitTrans
     End If ' Absch(aktab).art = 0 THEN ' Wenn keine ermächtigte Einrichtung
' Ende case nlArt
   End Select
   Select Case Uab
    Case uabName, NLrt
     Uab = Uab + 1
   End Select
  Next j
  
'  HACn.Execute "INSERT INTO `arzt`(nachname,vorname,titel_id,fachrichtung_id,lanr,nlart_id) VALUES('" & Arzt.Nachname & "','" & Arzt.Vorname & "'," & Arzt.titel_id & "," & Arzt.fachrichtung_id & "," & Arzt.LANR & " ," & Arzt.nlart_id & ") ON duplicate key update nachname = '" & Arzt.Nachname & "', vorname = '" & Arzt.Vorname & "', titel_id = " & Arzt.titel_id & ", fachrichtung_id = " & Arzt.fachrichtung_id & ", nlart_id = " & Arzt.nlart_id, rAF
  Set rs = Nothing
  myFrag rs, "SELECT `idbs` FROM `bs` WHERE `bsnr` = " & BS.BSNR, adOpenStatic, HACn, adLockReadOnly
  If rs.BOF Then
   Set rs = Nothing
   InsKorr HACn, HACnS, "INSERT INTO `bs`(`name`,`straße`,`hausnr`,`plz`,`ort_id`,`bsnr`,`bsart_id`,`sprechzeiten_id`,`rollst`,`aktzeit`,`seit`) VALUES('" & doUmwfSQL(BS.name, LVobMySQL) & "','" & doUmwfSQL(BS.Straße.Value, LVobMySQL) & "','" & BS.Hausnr & "','" & BS.plz & "'," & BS.Ort_id & ",'" & BS.BSNR & "'," & BS.bsart_id & "," & BS.sprechzeiten_id & "," & BS.Rollst & "," & DatFor_k(KVDate) & "," & DatFor_k(KVDate) & ")", rAF
   If rAF = 0 Then MsgBox "Einrichtung: " & BS.name & ", " & BS.Straße & ", " & BS.Ort_id & " konnte nicht in bs eingefügt werden!"
   Set rs = myEFrag("SELECT last_insert_id()", , HACn)
  Else
   myEFrag "UPDATE `bs` SET `name` = '" & doUmwfSQL(BS.name, LVobMySQL) & "',`straße` = '" & doUmwfSQL(BS.Straße.Value, LVobMySQL) & "',`hausnr` = '" & BS.Hausnr & "',`plz` = '" & BS.plz & "',`ort_id` = " & BS.Ort_id & ",`bsart_id` = " & BS.bsart_id & ",`sprechzeiten_id` = " & BS.sprechzeiten_id & ",`rollst` = " & BS.Rollst & ",`aktzeit` = " & DatFor_k(KVDate) & " WHERE `bsnr` = '" & BS.BSNR & "'", rAF, HACn
   If rAF > 1 Then
    MsgBox "raf = " & rAF & " beim Einfügen der Betriebsstätte: " & BS.Straße & ", " & BS.plz & " (" & BS.BSNR & ")"
   End If
  End If
  idIns1 = rs.Fields(0)
  HACn.BeginTrans
  myEFrag "DELETE FROM `tel` WHERE `bs_id` = " & idIns1, rAF, HACn
  For jj = 0 To telz
   myFrag rs, "SELECT idTel FROM `tel` WHERE `Tel` = '" & tel(jj) & "' AND bs_id = " & idIns1, adOpenStatic, HACn, adLockReadOnly
'   Set rs = Nothing
'   rs.Open "SELECT idTel FROM `tel` WHERE `Tel` = '" & tel(jj) & "' AND bs_id = " & idIns1, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    InsKorr HACn, HACnS, "INSERT INTO `tel`(`Tel`,`bs_id`) VALUES('" & tel(jj) & "'," & idIns1 & ")", rAF
   End If
  Next jj
  myEFrag "DELETE FROM `fax` WHERE `bs_id` = " & idIns1, rAF, HACn
  For jj = 0 To faxz
   myFrag rs, "SELECT idFax FROM `fax` WHERE `Fax` = '" & fax(jj) & "' AND bs_id = " & idIns1, adOpenStatic, HACn, adLockReadOnly
'   Set rs = Nothing
'   rs.Open "SELECT idFax FROM `fax` WHERE `Fax` = '" & fax(jj) & "' AND bs_id = " & idIns1, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    InsKorr HACn, HACnS, "INSERT INTO `fax`(`Fax`,`Faxzahl`,`bs_id`) VALUES('" & fax(jj) & "','" & REPLACE$(fax(jj), "-", "") & "'," & idIns1 & ")", rAF
   End If
  Next jj
  myEFrag "DELETE FROM `mail` WHERE `bs_id` = " & idIns1, rAF, HACn
  For jj = 0 To mailz
   myFrag rs, "SELECT `idmail` FROM `mail` WHERE `Mail` = '" & mail(jj) & "' AND bs_id = " & idIns1, adOpenStatic, HACn, adLockReadOnly
'   Set rs = Nothing
'   rs.Open "SELECT `idmail` FROM `mail` WHERE `Mail` = '" & mail(jj) & "' AND bs_id = " & idIns1, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    InsKorr HACn, HACnS, "INSERT INTO `mail`(`mail`,`bs_id`) VALUES('" & mail(jj) & "'," & idIns1 & ")", rAF
   End If
  Next jj
  HACn.CommitTrans
  
  If Absch(aktab).art = 0 Then ' Wenn keine ermächtigte Einrichtung
   myEFrag "UPDATE `arzt_has_bs` SET aktzeit = " & DatFor_k(KVDate) & ", obneben = " & BS.obNBS & ", obang = " & BS.obAng & " WHERE `bs_id` = " & idIns1 & " AND `arzt_id` = " & idIns2, rAF, HACn
   If rAF <> 1 Then
    HACn.BeginTrans
 '   HACn.Execute "DELETE FROM `arzt_has_bs` WHERE `bs_id` = " & idIns1 & " AND `arzt_id` = " & idIns2, rAF
    myFrag rs, "SELECT `obneben` FROM `arzt_has_bs` WHERE `bs_id` = " & idIns1 & " AND `arzt_id` = " & idIns2, adOpenStatic, HACn, adLockReadOnly
'    Set rs = Nothing ' da ein Arzt versehentlich doppelt drin stehen kann (Georg Hochheuser)
'    rs.Open "SELECT `obneben` FROM `arzt_has_bs` WHERE `bs_id` = " & idIns1 & " AND `arzt_id` = " & idIns2, HACn, adOpenStatic, adLockReadOnly
    If rs.BOF Then
     InsKorr HACn, HACnS, "INSERT INTO `arzt_has_bs`(`bs_id`,`arzt_id`,`obneben`,`obang`,`aktzeit`,`seit`) VALUES(" & idIns1 & "," & idIns2 & "," & BS.obNBS & "," & BS.obAng & "," & DatFor_k(KVDate) & "," & DatFor_k(KVDate) & ")", rAF
    End If
    HACn.CommitTrans
   End If
  Else ' Absch(aktab).art = 0 THEN ' Wenn keine ermächtigte Einrichtung
'      HACn.Execute "begin"
      myEFrag "DELETE FROM `bs_has_leistungsumfang` WHERE `bs_id` = " & idIns1, rAF, HACn
      For jj = 1 To LeiUz
       idIns = indIns(HACn, HACnS, "leistungsumfang", "leistungsumfang", LeiU(jj).Value, "idleistungsumfang")
'       ON Error Resume Next
       InsKorr HACn, HACnS, "INSERT INTO `bs_has_leistungsumfang`(`bs_id`,`leistungsumfang_id`) VALUES(" & idIns1 & "," & idIns & ")", rAF
'       ON Error GoTo fehler
      Next jj
      myEFrag "DELETE FROM `bs_has_genehmigung` WHERE `bs_id` = " & idIns1, rAF, HACn
'      Dim dop%
'      Dim jjjj%
      For jj = 1 To Genehz
       dop = 0
       For jjjj = 1 To jj - 1
        If Geneh(jjjj) = Geneh(jj) Then
         dop = True
         Exit For
        End If
       Next jjjj
       If Not dop Then
        idIns = indIns(HACn, HACnS, "genehmigung", "genehmigung", Geneh(jj).Value, "idgenehmigung")
        InsKorr HACn, HACnS, "INSERT INTO `bs_has_genehmigung`(`bs_id`,`genehmigung_id`) VALUES(" & idIns1 & "," & idIns & ")", rAF
       End If
      Next jj
      myEFrag "DELETE FROM `bs_has_vertragsangebot` WHERE `bs_id` = " & idIns1, rAF, HACn
      For jj = 1 To VertAz
       dop = 0
'       IF jj = 20 THEN Stop
'       IF VertA(jj).Value = "Franz-Josef-Strauß-Allee 11" THEN Stop
' bei "Franz-Josef-Strauß-Allee 11" / "Franz-Josef-Strauss-Allee 11" versagt es
       For jjjj = 1 To jj - 1
        If VertA(jjjj) = VertA(jj) Then
         dop = True
         Exit For
        End If
       Next jjjj
       If Not dop Then
        idIns = indIns(HACn, HACnS, "vertragsangebot", "vertragsangebot", VertA(jj).Value, "idvertragsangebot")
'        GoTo erneut
'        ON Error Resume Next
        InsKorr HACn, HACnS, "INSERT INTO `bs_has_vertragsangebot`(`bs_id`,`vertragsangebot_id`) VALUES(" & idIns1 & "," & idIns & ")", rAF
'        ON Error GoTo fehler
       End If
      Next jj
      myEFrag "COMMIT", , HACn
  
  End If ' Absch(aktab).art = 0 THEN ' Wenn keine ermächtigte Einrichtung
' Plausibilitätskontrollen:
'SELECT * FROM mail WHERE not mail RLIKE '^[\-\_\.[:alnum:]]+@[\-\_\.[:alnum:]]+\.(org|de|com|ag|info|net|biz)+$';
'SELECT * FROM tel WHERE not tel RLIKE '^[-0123456789]+$';
'SELECT * FROM fax WHERE not fax RLIKE '^[-0123456789]+$';
 Exit Sub
 
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
'If InStrB(Err.Description, "Doppelter Eintrag '76-61'") <> 0 THEN Resume Next
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF, HACn
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in proTeilnehmer/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' proTeilnehmer

