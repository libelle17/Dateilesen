Attribute VB_Name = "HAEinlesen"
Option Explicit
Const DBName$ = "haerzte"
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
 Zusatzbezeichnung
 Weiterbildung
 Genehmigung
 Leistungsumfang
 Vertragsangebot
 Fremdsprache
 ÄrztederPraxis
End Enum

Private Type ArztTyp
 idarzt As Long
 Nachname As String
 Vorname As String
 Namenszusatz As String
 titel_id As Long
' fachrichtung_id As Long
 LANR As Long
 nlart_id As Long
 AktZeit As Date
End Type

Private Type BSTyp
 idbs As Long
 Straße As New CString
 Hausnr As New CString
 Plz As String
 Ort_id As Long
 BSNR As String
 obNBS As Integer
 obAng As Integer
 bsart_id As Long
 sprechzeiten_id As Long
 Rollst As Integer
 AktZeit As Date
End Type
Const ArztSuchS$ = "KVB_Arztsuche*.pdf"
Public Function doalleKVDateien()
 Dim i&, erg&
 Dim SL As New SortierListe, SD As SortierDatei, Fil$
 Call Lese.ProgStart
 Fil = Dir(vVerz & ArztSuchS) ' "v:\KVB_Arztsuche*.pdf"
 Do While LenB(Fil) <> 0
  Set SD = New SortierDatei
  Set SD.File = FSO.GetFile(Fil)
  SL.sCAdd SD
  Fil = Dir
 Loop
 erg = MsgBox("Sollen wirklich alle " & SL.Count & " Hausarztdaten aus '" & vVerz & "' neu eingelesen werden?", vbYesNo)
 If erg = vbYes Then
  If HALöschen Then
   For i = 1 To SL.Count
    dodoHausärzteEinlesen SL.Item(i).File.path, False
   Next i
  End If
 End If
End Function ' alleKVDateien
Function TLösch(Tb$)
' HACn.Execute "truncate `" & Tb & "`", rAF
 HACn.Execute "delete from `" & Tb & "`", rAF
 If rAF <> 0 Then lies.Ausgeb rAF & " Datensätze aus `" & Tb & "` gelöscht", True
' On Error Resume Next
 HACn.Execute "alter table `" & Tb & "` auto_increment=1", rAF
End Function

Function HALöschen%()
  Dim erg&
  Verbinde
  erg = MsgBox("Sollen wirklich alle Tabellen der Tabelle " & DBName & " der Verbindung:" & vbCrLf & HACn & vbCrLf & "gelöscht werden?", vbYesNo, "Sicherheitsrückfrage")
  If erg = vbNo Then
   HALöschen = 0
  Else
' mit alter table geht irgendwie nicht 30.11.09
   TLösch "bs"
   TLösch "bsart"
   TLösch "arzt"
   TLösch "fachrichtung"
   TLösch "ort"
   TLösch "bsart"
   TLösch "fremdsprache"
   TLösch "genehmigung"
   TLösch "nlart"
   TLösch "sprechzeiten"
   TLösch "titel"
   TLösch "vertragsangebot"
   TLösch "zusatzbezeichnung"
   TLösch "weiterbildung"
   TLösch "arzt_has_bs" ' "alter table `arzt_has_bs` auto_increment=0"
   TLösch "arzt_has_fachrichtung"
   TLösch "arzt_has_fremdsprache"
   TLösch "arzt_has_genehmigung"
   TLösch "arzt_has_vertragsangebot"
   TLösch "arzt_has_weiterbildung"
   TLösch "arzt_has_zusatzbezeichnung"
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
   .InitDir = vVerz
' "c:\v\KVB_Arztsuche_-20090919-221246.pdf"
   .DefaultExt = "pdf"
   .Filename = ArztSuchS
   Dim jFil$, jDat As Date, erg$
   erg = Dir(.InitDir & .Filename)
   Do While LenB(erg) <> 0
    If erg <> "." And erg <> ".." Then
     If FileDateTime(.InitDir & erg) > jDat Then
'      jDat = FileDateTime(.InitDir & erg)
      jDat = ZiehDat(.InitDir & erg)
      jFil = erg
     End If
    End If
    erg = Dir
   Loop
   .Filename = .InitDir & jFil
'   .Flags = .Flags And Not FileOpenConstants.cdlOFNFileMustExist
   If Not obStumm Then .ShowOpen
   getHAPDF = .Filename
  End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getHAPDF/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getHAPDF
Public Function ZiehDat(Datei$) As Date
      Dim p1&, pz&, bst$, zds$
      p1 = Len(Datei)
      pz = 0
      Do
       bst = Mid(Datei, p1, 1)
       If bst = "-" Then pz = pz + 1
       If pz = 2 Or p1 = 1 Then Exit Do
       p1 = p1 - 1
      Loop
      zds = Mid(Datei, p1 + 7, 2) & "." & Mid(Datei, p1 + 5, 2) & "." & Mid(Datei, p1 + 1, 4) & " " & Mid(Datei, p1 + 10, 2) & ":" & Mid(Datei, p1 + 12, 2) & ":" & Mid(Datei, p1 + 14, 2)
      If IsDate(zds) Then
       ZiehDat = CDate(zds)
      Else
       ZiehDat = FileDateTime(Datei)
      End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZiehDat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZiehDat

Public Sub doHausärzteEinlesen()
 Dim Filename$
 Filename = getHAPDF(True)
 dodoHausärzteEinlesen Filename, True
End Sub ' doHausärzteEinlesen

Sub Verbinde()
 Dim i%
 HACn.Open DBCn.ConnectionString
 For i = 1 To 2
  If i = 1 Then On Error Resume Next
  HACn.Execute "use `" & DBName & "`"
  If Err.Number <> 0 Then doMach_haerzte DBName, MachDatenbank.GetServer(DBCn) ' keine "`" hier
 Next i
End Sub ' Verbinde

Sub dodoHausärzteEinlesen(Filename$, obRückfrage%)
' Const DBName$ = "hausaerzte"
 Dim PDTextS As Acrobat.CAcroPDTextSelect
 Dim Result&, PDDoc As Acrobat.CAcroPDDoc
 Dim PDPage As Acrobat.CAcroPDPage
 Dim PDHili As Acrobat.CAcroHiliteList
 Dim B As Boolean, erg$
 Dim PDFStream$
' Strings Falls Text von Anfang ausgelesen werden soll (bzw bis zum Ende)
' Dim FromFirst$, ToLast$
 Dim Str As New CString, NTL&
 Dim i%, j&, jj&, splitt$(), s2$(), Pid$, Datum As Date
 Dim Uhrzeit$, maxP&
 
' FromFirst = Chr(169) & Chr(170) & Chr(172)
' ToLast = Chr(163) & Chr(165) & Chr(164)
' On Error Resume Next
 On Error GoTo fehler
 Set PDDoc = CreateObject("AcroExch.pdDoc")
' On Error GoTo fehler
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
 Call doMach_haerzte(DBName, DBVerb.Cpt)
#End If
' Filename = Environ("userprofile") & "\Desktop\" & "KVB_Arztsuche_-20090905-045301.pdf"
' Nehme die Erste Seite - Index 0
 Do
  Set PDPage = PDDoc.AcquirePage(j)
  ' Erzeuge ein Highlight Objekt und weise ihm 2000 Elemente bei (keine Grenzprobleme)
  Set PDHili = CreateObject("AcroExch.HiliteList")
  B = PDHili.Add(0, 32000)
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
  lies.Ausgeb "Hausärzte einlesen: Lese S.: " & j, 0
' die ersten 30 Seiten einlesen
  If j > maxP Then Exit Do
 Loop
' in jeder Zeile von splitt steht eine Zeile aus der PDF-Datei
 SplitNeu Str.Value, vbCrLf, splitt, "(", ")"
 
 Verbinde
 
 Dim AbschBis&(), AbZ&
 AbZ = 0
 ReDim AbschBis(0)
 Dim aktab&, LANRZeile&, aktName$ ' ob die letzte Zeile eine Facharztzeile war -> dann nicht zählen
 For aktab = LBound(splitt) To UBound(splitt)
  Select Case splitt(aktab)
   Case "Fachärztin für Psychosomatische Medizin und", "Facharzt für Psychosomatische Medizin und"
    splitt(aktab) = splitt(aktab) & " " & splitt(aktab + 1)
    splitt(aktab + 1) = vbNullString
  End Select
'  If aktab = 46637 Then Stop
  Select Case splitt(aktab)
' eine der folgenden Ausdrücke steht hinter dem Namen, wodurch dieser als solches klassifizierbar ist
   Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin", "Psychotherapeutisch tätige Ärztin", "Psychotherapeutisch tätiger Arzt", "Kinder- und Jugendlichen-Psychotherapeutin", "Kinder- und Jugendlichen-Psychotherapeut", "Fachärztin für Psychiatrie und Psychotherapie", "Facharzt für Psychiatrie und Psychotherapie", "Facharzt für Neurologie", "Facharzt für Nervenheilkunde", "Fachärztin für Neurologie", "Fachärztin für Nervenheilkunde", "Fachärztin für Psychosomatische Medizin und Psychotherapie", "Facharzt für Psychosomatische Medizin und Psychotherapie", "Schwerpunkt Kinder-Kardiologie", "Facharzt für Kinder- und Jugendmedizin (Schwerpunkt Neonatologie)"
    If LANRZeile <> 0 Then
neuerAbschnitt:
     If AbZ = UBound(AbschBis) Then
      ReDim Preserve AbschBis(UBound(AbschBis) + 500)
     End If
     AbZ = AbZ + 1
     If aktab = 1 Then
      AbschBis(AbZ) = aktab - 2
     ElseIf InStrB(splitt(aktab - 1), " ") = 0 Then ' wenn der Name sich über zwei Zeilen zieht, dann ist in der zweiten Zeile evtl. kein Leerzeichen mehr; Angerpointner
      AbschBis(AbZ) = aktab - 3
      splitt(aktab - 2) = splitt(aktab - 2) & " " & splitt(aktab - 1)
      splitt(aktab - 1) = vbNullString
     Else
      AbschBis(AbZ) = aktab - 2
     End If
'     aktab = aktab + 1 ' sonst wird "Psychotherapie" nicht zusammengezogen, s. Becker-Jakubaß
     LANRZeile = 0
    End If
    aktName = splitt(aktab - 1)
   Case "" ' Nach Zusammenzug von "Fachärztin für Psychosomatische Medizin und Psychotherapie"
'    Stop
   Case Else
    If LANRZeile = 0 Then
'     If Left$(splitt(aktab), 4) = "LANR" Then
     If Left$(splitt(aktab), 4) = "LANR" Then
      LANRZeile = aktab
     End If
' wenn LANRZeile <> 0, darf normal kein Ort mehr kommen, weil der schon vor LANR kommt.
' Kommt doch noch später einer, so handelt es sich um einen neuen Arzt ohne Niederlassungsart -> Gerald Beier
    ElseIf LANRZeile <> 0 Then
     If aktName = "Dr.med. Christiane Beier" Then
      If IsNumeric(Left(splitt(aktab), 1)) Then ' weil "like" so lang brauchen soll
       If splitt(aktab) Like "##### *" And Not splitt(aktab) Like "* * * *" And Not InStrB(splitt(aktab), ",") <> 0 And Not InStrB(splitt(aktab), ".") <> 0 And Not InStrB(splitt(aktab), " EBM") <> 0 And Not InStrB(splitt(aktab), " und ") <> 0 And Not InStrB(splitt(aktab), "Komplex") <> 0 Then ' PLZ mit Ort
'        MsgBox "vermutlich Fehlerhafte Verarbeitung der Zeile " & aktab & ": " & vbCrLf & splitt(aktab)
        GoTo neuerAbschnitt
       End If
      End If
     End If
    End If ' LANRZeile = True / else
  End Select
 Next aktab
 If AbZ = UBound(AbschBis) Then
  ReDim Preserve AbschBis(UBound(AbschBis) + 500)
 End If
 AbZ = AbZ + 1
 AbschBis(AbZ) = UBound(splitt) - 1
 For aktab = 2 To AbZ
  Call proTeilnehmer(splitt, AbschBis, aktab, HACn)
 Next aktab
 
 lies.Ausgeb "Korrigiere Geschlechter ...", 0
 HACn.Execute "update arzt b left join (select distinct vorname from arzt a left join nlart on a.nlart_id = idnlart left join arzt_has_fachrichtung ahf on a.idarzt = ahf.arzt_id left join fachrichtung f on idfachrichtung = fachrichtung_id where fachrichtung like '%ärztin%' or fachrichtung like '%eutin%' or niederlassungsart like '%eutin%') a on b.vorname = a.vorname set b.obweibl=1 where not isnull(a.vorname)", rAF
 lies.Ausgeb (AbZ - 1) & " Ärzte aus '" & Filename & "' eingelesen; " & rAF & " davon auf weiblich geändert", 1
 SplitNeu splitt(0), " ", s2
 For j = 0 To UBound(s2)
  If InStrB(s2(j), "-") > 0 Then
   s2 = Split(s2(j), "-")
   Datum = CDate(s2(0))
   Exit For
  End If
 Next j
 Result = PDDoc.Close
 Set PDPage = Nothing
 Set PDHili = Nothing
 Set PDTextS = Nothing
 Set PDDoc = Nothing
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HausärzteEinlesen_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' HausärzteEinlesen_Click

Sub proTeilnehmer(ByRef splitt$(), ByRef AbschBis&(), aktab&, ByRef HACn As ADODB.Connection)
  Dim Uab As UAbTyp, Arzt As ArztTyp, BS As BSTyp, Ort As New CString, SZ As New CString, s2$()
  Dim tel$(), telz&, fax$(), faxz&, mail() As New CString, mailz&
  Dim OrtZeile&, LANRZeile&, szzz&  ' SprechzeitenZeilenZahl
  Dim fren&() ' Fachrichtungen des aktuellen Teilnehmers
'  ReDim fren&(0)
  Dim j&, jj&, jjj&, szz&, buch$, pos&, p2&
  Dim idIns&, idIns2
  On Error GoTo fehler
' alle Zeilen des aktuellen Teilnehmers
  Uab = uabName
  For j = AbschBis(aktab - 1) + 1 To AbschBis(aktab)
'   If Left$(splitt(j), 12) = "Gerald Beier" Then Stop
   Select Case Uab
    Case uabName
'     If InStrB(splitt(j), "Massoudy") <> 0 Or InStrB(splitt(j), "Al-Iassin") <> 0 Or InStrB(splitt(j), "Guarch") <> 0 Then Stop
     SplitNeu splitt(j), " ", s2
'     If InStr(splitt(j), "Ackern") <> 0 Then Stop
     If UBound(s2) > 1 Then
      If s2(UBound(s2) - 1) = "Al" Then
       s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
       ReDim Preserve s2(UBound(s2) - 1)
      End If
     End If
'     If s2(UBound(s2)) = "von" Then
'      Arzt.Namenszusatz = s2(UBound(s2))
''      s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
'      ReDim Preserve s2(UBound(s2) - 1)
'     End If
'     If s2(UBound(s2)) = "Freiherr von" Then
'      Arzt.Namenszusatz = s2(UBound(s2))
''      s2(UBound(s2) - 1) = s2(UBound(s2) - 1) & " " & s2(UBound(s2))
'      ReDim Preserve s2(UBound(s2) - 1)
'     End If
      For jj = UBound(s2) To 1 Step -1
       Select Case s2(jj)
        Case "Freiherr", "Freifrau", "von", "van", "van den", "v.", "de", "den", "vom", "zur", "genannt", "jun.", "sen."
         Arzt.Namenszusatz = s2(jj) & IIf(Arzt.Namenszusatz = "", "", " ") & Arzt.Namenszusatz
         For jjj = jj + 1 To UBound(s2)
          s2(jjj - 1) = s2(jjj)
         Next jjj
         ReDim Preserve s2(UBound(s2) - 1)
       End Select
      Next jj
     If s2(UBound(s2)) = "32123," Then Stop
     Arzt.Nachname = s2(UBound(s2))
     If UBound(s2) = 1 Then
      Arzt.Vorname = s2(0)
     Else
      Select Case s2(0)
       Case "Afssaneh", "Alex", "Ali", "Amir", "Andreas", "Angela", "Anna", "Anne", "Berthold", "Birgit", "Carl", "Caroline", "Christian", "Christoph", "Claus", "Detlef", "Dieter", "Edith", "Elisabeth", "Eugen", "Evelyn", "Frank", "Franz", "Frida", "Gabriele", "Gisela", "Gerald", "Hamdi", "Hans", "Harriet", "Harry", "Heide", "Heiner", "Heinz", "Helga", "Hermina", "Ilse", "Ingeborg", "Irmgard", "Jhonny", "Joachim", "Jörn", "Johannes", "Johannus", "John", "Jorge", "Jorj", "Kai", "Kambiz", "Karin", "Karl", "Kerstin", "Klaus", "Lara", "Manfred", "Mara", "Maria", "Martina", "MartinH.", "Mirjam", "Mitra", "Oliver", "Peter", "Petra", "Rebecca", "Robert", "Rolf", "Roser", "Ruth", "Shervin", "Sigrid", "Silke", "Stefan", "Tahereh", "Thomas", "Thorsten", "Trude", "Ursula", "Ute", "Veit", "Vladimiro", "Werner", "Wilhelm", "Wolf", "Wolfgang", "Z."
        s2(1) = s2(0) & " " & s2(1)
        s2(0) = vbNullString
       Case "Kinder-u.Jugendlichen-Psychotherapeutin"  ' Fehler in Datenbank: Läuft auch als Titel
        s2(0) = vbNullString
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
       Arzt.Nachname = Mid(Arzt.Vorname, pos) & " " & Arzt.Nachname
       Arzt.Vorname = Trim(Left$(Arzt.Vorname, pos - 1))
      End If
      Arzt.titel_id = indIns(HACn, "titel", "titel", s2(0), "idtitel")
     End If ' rs.bof
     lies.Ausgeb "Bearbeite: " & Arzt.Nachname & ", " & Arzt.Vorname, 0
    Case NLrt
' die Zeile, in der PLZ und Ort stehen, enthält als erstes eine Ziffer
     For jj = j To AbschBis(aktab) ' j+1 gestrichen wegen Gerald Beier
      If IsNumeric(Left$(splitt(jj), 5)) And Len(splitt(jj)) > 6 Then OrtZeile = jj: Exit For
     Next jj
     If OrtZeile >= AbschBis(aktab) Then Stop
'     If OrtZeile = 1890 Then Stop
' in der Zeile nach dem Namen könnte, muß aber nicht, eine Niederlassungsart stehen
     Select Case splitt(j)
      Case "Facharzt", "Hausarzt", "Fach- und Hausarzt", "Psychologischer Psychotherapeut", "Psychologische Psychotherapeutin"
       Arzt.nlart_id = indIns(HACn, "nlart", "niederlassungsart", splitt(j), "idnlart")
       j = j + 1
     End Select
' über die Zeile getrennte Fachrichtungen zusammenführen
     For jj = j To OrtZeile - 2
      jjj = 0
      While (InStrB(splitt(jj), "(") <> 0 And InStrB(splitt(jj), ")") = 0) Or Right$(splitt(jj), 2) = "/-"
       jjj = jjj + 1
       If Right$(splitt(jj), 1) <> " " Then splitt(jj) = splitt(jj) & " "
       splitt(jj) = splitt(jj) & splitt(jj + jjj)
       splitt(jj + jjj) = vbNullString
      Wend
     Next jj
' Teilgebiete und Schwerpunkte gesondert aufnehmen
     Dim FrSep$(1), ergstr$(), ez&
     FrSep(0) = "Schwerpunkt"
     FrSep(1) = "Teilgebiet"
'     If Arzt.Nachname = "Abelein" And Arzt.Vorname = "Richard" Then Stop
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
        If l1.Instr("ärztin") <> 0 Or l1.Instr("arzt") <> 0 Or l1.Instr("olog") <> 0 Or l1.Instr("medizin") <> 0 Or l1.Instr("ologie") <> 0 Or l1.Instr("psych") <> 0 Or l1.Instr("schwerpunkt") <> 0 Or l1.Instr("teigebiet") <> 0 Then
         If l1.Instr("MVZ") = 0 And l1.Instr("Versorgungszentrum") = 0 Then
          GoTo stimmt
         End If
        End If
        If l1.Value Like "* #*" Or l1.Value Like "#*" Or IsNumeric(l1.Value) Or Len(l1.Value) < 3 Or Left(l1.Value, 4) = "str." Or l1.Value Like "*str.*" Or InStrB(l1.Value, ",") <> 0 Then
        Else
         Select Case ergstr(jjj)
'         Case "Berufsausübungsgem. / MVZ-München, 80799", "Einsteinstr. 130", "München, Heßstr. 22", "Ofenthaler Weg 20"
          Case "München", "Erl", "Ansbach", "Hasenkopf", "Nbg", "Windsheim", "Nürnberg", "Kollegen", "Marktheidenfeld-Michelrieth", "Partenkirchen/Auenst", "Rosenheimer Str.", "Roßhaupter-96", "Steenbeek/Groneberg", "Wasserburg", "Wessels-Str.", "Wichtermann", "ÜÖ BAG, Dres.med. Neher/Hilber & Kollegen", "Aurnhammer-Str.", "Bad Gögging", "Bahnhofspl.", "Gemünden-Langenprozelten", "Günzburger Str.41", "Hohenwarter Str", "Kalbskopf-Str", "Preuschwitz.Str.101"
            Debug.Print ergstr(jjj)
          Case Else
stimmt:
           fren(UBound(fren)) = indIns(HACn, "fachrichtung", "Fachrichtung", ergstr(jjj), "idFachrichtung")
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
      BS.Straße = Mid(BS.Straße, pos)
     End If
     SplitNeu splitt(OrtZeile), " ", s2
     If Not IsNumeric(s2(0)) Then ' Leerzeichen vergessen
      ReDim s2(1)
      For jjj = 1 To Len(splitt(OrtZeile))
       If IsNumeric(Mid(splitt(OrtZeile), jjj, 1)) Then
        s2(0) = s2(0) & Mid(splitt(OrtZeile), jjj, 1)
       Else
        Exit For
       End If
      Next jjj
      s2(1) = Mid(splitt(OrtZeile), Len(s2(0)) + 1)
     End If
     BS.Plz = s2(0)
     If UBound(s2) > 0 Then
      Ort = s2(1)
      For jj = 2 To UBound(s2)
       Ort.Append " "
       Ort.Append s2(jj)
      Next jj
      BS.Ort_id = indIns(HACn, "ort", "Ort", Ort.Value, "idOrt")
     End If
'     Select Case splitt(j)
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
'     End Select
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
     For jj = j + 1 To AbschBis(aktab)
      If splitt(jj) = "Sprechzeiten:" Then
       szz = jj
       Exit For
      End If
     Next jj
     If szz = 0 Then
' kam bisher nicht vor
      MsgBox "Keine Sprechzeiten bei Arzt " & Arzt.Nachname & ", " & Arzt.Vorname
      szz = AbschBis(aktab)
     End If
     If splitt(j) = "Angestellt in MVZ und überörtliche" Then
      splitt(j + 1) = splitt(j) & " " & splitt(j + 1)
      splitt(j) = vbNullString
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
'     If BSArt = "Angestellt in MVZ" Then Stop
     If Left$(BSArt, 14) = "Angestellt in " Then BS.obAng = 1: BSArt = Mid(BSArt, 15)
     BS.bsart_id = indIns(HACn, "bsart", "BSArt", BSArt, "idbsart")
     j = szz
     For jj = j + 1 To AbschBis(aktab)
      If splitt(jj) = "LANR: BSNR:" Then LANRZeile = jj: BS.obNBS = False: Exit For
      If splitt(jj) = "LANR: NBSNR:" Then LANRZeile = jj: BS.obNBS = 1: Exit For
     Next jj
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
'     If Left$(SZ.Value, 6) <> "Montag" And Left$(SZ.Value, 8) <> "Dienstag" And Left$(SZ.Value, 8) <> "Mittwoch" And Left$(SZ.Value, 10) <> "Donnerstag" And Left$(SZ.Value, 7) <> "Freitag" And SZ.Value <> "keine Angaben" Then Stop
     BS.sprechzeiten_id = indIns(HACn, "sprechzeiten", "Sprechzeiten", SZ.Value, "idsprechzeiten")
     SplitNeu splitt(LANRZeile + 1), " ", s2
'     If UBound(s2) <> 1 Then Stop
     Arzt.LANR = s2(0)
'     If BS.obNBS = 1 Then BS.NBSNR = s2(1): BS.BSNR = 0 Else BS.NBSNR = 0: BS.BSNR = s2(1)
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
     Dim KennZeile&(Zusatzbezeichnung To ÄrztederPraxis)
     Dim ZB() As New CString, Wb() As New CString, Geneh() As New CString, VertA() As New CString, fS() As New CString, ÄP() As New CString
     Dim ZBz&, WBz&, Genehz&, VertAz&, FSz&, ÄPz&, überspringen&
     For jj = LANRZeile To AbschBis(aktab)
'      If Left(splitt(jj), 6) = "Email:" Then KennZeile(Email) = jj
      Select Case splitt(jj)
       Case "Zusatzbezeichnungen:"
         KennZeile(Zusatzbezeichnung) = jj: Uab = Zusatzbezeichnung
       Case "Weiterbildungen:":            KennZeile(Weiterbildung) = jj: Uab = Weiterbildung
       Case "Genehmigungen:":              KennZeile(Genehmigung) = jj: Uab = Genehmigung
       Case "Leistungsumfang:":            KennZeile(Leistungsumfang) = jj: Uab = Leistungsumfang
       Case "Besondere Vertragsangebote:": KennZeile(Vertragsangebot) = jj: Uab = Vertragsangebot
       Case "Fremdsprachen:":              KennZeile(Fremdsprache) = jj: Uab = Fremdsprache
       Case "Ärzte der Praxis:":           KennZeile(ÄrztederPraxis) = jj: Uab = ÄrztederPraxis
       Case "Rollstuhlgerechte Praxis":    BS.Rollst = 1
       Case Else
        überspringen = 0
        Do While jj <= AbschBis(aktab) And ((Left$(splitt(jj + 1 + überspringen), 4) = "Arzt" And KennZeile(Genehmigung) <> 0 And KennZeile(Vertragsangebot) = 0) Or _
        Right$(splitt(jj), 5) = "nicht" Or Right$(splitt(jj), 3) = "und" Or Right$(splitt(jj), 2) = "u." Or Right$(splitt(jj), 1) = ";" Or Left$(splitt(jj + 1 + überspringen), 4) = "bzw." Or Right$(splitt(jj), 2) = "f." Or Right$(splitt(jj), 3) = "mit" Or Left$(splitt(jj + 1 + überspringen), 1) = "(" Or Left$(splitt(jj + 1 + überspringen), 1) = "/" Or Right$(splitt(jj), 1) = "/" Or Right$(splitt(jj), 1) = "-" Or Right$(splitt(jj), 1) = "," Or Right$(splitt(jj), 5) = "sches" Or Left$(splitt(jj + 1 + überspringen), 3) = "SGB" Or Left$(splitt(jj + 1 + überspringen), 14) = "sonographisch)" Or Left$(splitt(jj + 1 + überspringen), 8) = "mammogr." Or Right$(splitt(jj), 5) = "spez." Or Left$(splitt(jj + 1 + überspringen), 3) = "IPM" Or Left$(splitt(jj + 1 + überspringen), 18) = "Kinder/Jugendliche" Or Left$(splitt(jj + 1 + überspringen), 12) = "Operationen)" Or Right$(splitt(jj), 6) = "in der" Or splitt(jj) = "DMP Brustkrebs teilnehmender Facharzt")
'         If Left$(splitt(jj + 1 + überspringen), 5) = "Arzt" Then Stop
         überspringen = überspringen + 1
'         If Uab <> ÄrztederPraxis Then Stop
         splitt(jj) = splitt(jj) & " " & splitt(jj + überspringen)
        Loop
        Select Case Uab
         Case Zusatzbezeichnung
          ZBz = ZBz + 1
          If ZBz = 1 Then ReDim ZB(60) Else If ZBz > UBound(ZB) Then ReDim Preserve ZB(UBound(ZB) + 60)
          ZB(ZBz) = splitt(jj)
         Case Weiterbildung
          WBz = WBz + 1
          If WBz = 1 Then ReDim Wb(60) Else If WBz > UBound(Wb) Then ReDim Preserve Wb(UBound(Wb) + 60)
          Wb(WBz) = splitt(jj)
         Case Genehmigung
          Genehz = Genehz + 1
          If Genehz = 1 Then ReDim Geneh(60) Else If Genehz > UBound(Geneh) Then ReDim Preserve Geneh(UBound(Geneh) + 60)
          Geneh(Genehz) = splitt(jj)
         Case Vertragsangebot
          VertAz = VertAz + 1
          If VertAz = 1 Then ReDim VertA(60) Else If VertAz > UBound(VertA) Then ReDim Preserve VertA(UBound(VertA) + 60)
          VertA(VertAz) = splitt(jj)
         Case Fremdsprache
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
      If Arzt.titel_id = 0 Then Arzt.titel_id = indIns(HACn, "titel", "titel", "", "idtitel")
      If Arzt.nlart_id = 0 Then Arzt.nlart_id = indIns(HACn, "nlart", "niederlassungsart", "", "idnlart")
      Set rs = Nothing
      rs.Open "select `idarzt` from `arzt` where `LANR` = " & Arzt.LANR, HACn, adOpenStatic, adLockReadOnly
      If rs.EOF Then
       Set rs = Nothing
       HACn.Execute "insert into `arzt`(`Nachname`,`Vorname`,`titel_id`,`LANR`,`nlart_id`,`Namenszusatz`,`aktzeit`) values('" & doUmwfSQL(Arzt.Nachname, LVobMySQL) & "','" & doUmwfSQL(Arzt.Vorname, LVobMySQL) & "'," & Arzt.titel_id & "," & Arzt.LANR & "," & Arzt.nlart_id & ",'" & Arzt.Namenszusatz & "'," & DatForm(KVDate) & ")"
       Set rs = HACn.Execute("select last_insert_id()")
      Else
       HACn.Execute "update `arzt` set `titel_id` = " & Arzt.titel_id & ", `nlart_id` = " & Arzt.nlart_id & ", `Namenszusatz` = '" & Arzt.Namenszusatz & "', `aktzeit` = " & DatForm(KVDate) & ", `Nachname` = '" & doUmwfSQL(Arzt.Nachname, LVobMySQL) & "', `Vorname` = '" & doUmwfSQL(Arzt.Vorname, LVobMySQL) & "' where LANR = " & Arzt.LANR, rAF
       If rAF > 1 Then
        MsgBox "raf = " & rAF & " beim Einfügen des Arztes: " & Arzt.Nachname & ", " & Arzt.Vorname & " (" & Arzt.LANR & ")"
       End If
      End If
      idIns2 = rs.Fields(0)
'      HACn.BeginTrans
'      HACn.CommitTrans
      HACn.Execute "begin"
      HACn.Execute "delete from `arzt_has_zusatzbezeichnung` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To ZBz
       idIns = indIns(HACn, "zusatzbezeichnung", "Zusatzbezeichnung", ZB(jj).Value, "idZusatzbezeichnung")
       HACn.Execute "insert into `arzt_has_zusatzbezeichnung`(`arzt_id`,`Zusatzbezeichnung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      HACn.Execute "delete from `arzt_has_weiterbildung` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To WBz
       idIns = indIns(HACn, "weiterbildung", "Weiterbildung", Wb(jj).Value, "idWeiterbildung")
       HACn.Execute "insert into `arzt_has_weiterbildung`(`arzt_id`,`Weiterbildung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      HACn.Execute "delete from `arzt_has_genehmigung` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To Genehz
       idIns = indIns(HACn, "genehmigung", "genehmigung", Geneh(jj).Value, "idgenehmigung")
       HACn.Execute "insert into `arzt_has_genehmigung`(`arzt_id`,`genehmigung_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      HACn.Execute "delete from `arzt_has_vertragsangebot` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To VertAz
       idIns = indIns(HACn, "vertragsangebot", "vertragsangebot", VertA(jj).Value, "idvertragsangebot")
       HACn.Execute "insert into `arzt_has_vertragsangebot`(`arzt_id`,`vertragsangebot_id`) values(" & idIns2 & "," & idIns & ")"
      Next jj
      HACn.Execute "delete from `arzt_has_fremdsprache` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To FSz
       If LenB(fS(jj).Value) <> 0 Then
        idIns = indIns(HACn, "fremdsprache", "fremdsprache", fS(jj).Value, "idfremdsprache")
        HACn.Execute "insert into `arzt_has_fremdsprache`(`arzt_id`,`fremdsprache_id`) values(" & idIns2 & "," & idIns & ")"
       End If
      Next jj
      HACn.Execute "delete from `arzt_has_fachrichtung` where `arzt_id` = " & idIns2, rAF
      For jj = 1 To UBound(fren) - 1
       ' Fachrichtungen können offenbar versehentlich auch doppelt genannt werden (-> Eckhard Rudolf, hat 2 x Unfallchirurgie und 2 x Viszeralchirurgie)
       Set rs = Nothing
       rs.Open "select arzt_id from `arzt_has_fachrichtung` where arzt_id = " & idIns2 & " and fachrichtung_id = " & fren(jj), HACn, adOpenStatic, adLockReadOnly
       If rs.BOF Then
        HACn.Execute "insert into `arzt_has_fachrichtung`(`arzt_id`,`fachrichtung_id`) values(" & idIns2 & "," & fren(jj) & ")"
       End If
      Next jj
      HACn.Execute "commit"
'      HACn.CommitTrans
   End Select
   Select Case Uab
    Case uabName, NLrt
     Uab = Uab + 1
   End Select
  Next j
'  HACn.Execute "insert into `arzt`(nachname,vorname,titel_id,fachrichtung_id,lanr,nlart_id) values('" & Arzt.Nachname & "','" & Arzt.Vorname & "'," & Arzt.titel_id & "," & Arzt.fachrichtung_id & "," & Arzt.LANR & " ," & Arzt.nlart_id & ") on duplicate key update nachname = '" & Arzt.Nachname & "', vorname = '" & Arzt.Vorname & "', titel_id = " & Arzt.titel_id & ", fachrichtung_id = " & Arzt.fachrichtung_id & ", nlart_id = " & Arzt.nlart_id, rAF
  Set rs = Nothing
  rs.Open "select `idbs` from `bs` where `bsnr` = " & BS.BSNR, HACn, adOpenStatic, adLockReadOnly
  If rs.BOF Then
   Set rs = Nothing
   HACn.Execute "insert into `bs`(`straße`,`hausnr`,`plz`,`ort_id`,`bsnr`,`bsart_id`,`sprechzeiten_id`,`rollst`,`aktzeit`) values('" & doUmwfSQL(BS.Straße.Value, LVobMySQL) & "','" & BS.Hausnr & "','" & BS.Plz & "'," & BS.Ort_id & ",'" & BS.BSNR & "'," & BS.bsart_id & "," & BS.sprechzeiten_id & "," & BS.Rollst & "," & DatForm(KVDate) & ")"
   Set rs = HACn.Execute("select last_insert_id()")
  Else
   HACn.Execute "update `bs` set `straße` = '" & doUmwfSQL(BS.Straße.Value, LVobMySQL) & "', `hausnr` = '" & BS.Hausnr & "', `plz` = '" & BS.Plz & "', `ort_id` = " & BS.Ort_id & ", `bsart_id` = " & BS.bsart_id & ", `sprechzeiten_id` = " & BS.sprechzeiten_id & ", `rollst` = " & BS.Rollst & ", `aktzeit` = " & DatForm(KVDate) & " where `bsnr` = '" & BS.BSNR & "'", rAF
   If rAF > 1 Then
    MsgBox "raf = " & rAF & " beim Einfügen der Betriebsstätte: " & BS.Straße & ", " & BS.Plz & " (" & BS.BSNR & ")"
   End If
  End If
  idIns = rs.Fields(0)
  HACn.BeginTrans
  HACn.Execute "delete from `tel` where `bs_id` = " & idIns, rAF
  For jj = 0 To telz
   Set rs = Nothing
   rs.Open "select idTel from `tel` where `Tel` = '" & tel(jj) & "' and bs_id = " & idIns, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    HACn.Execute "insert into `tel`(`Tel`,`bs_id`) values('" & tel(jj) & "'," & idIns & ")", rAF
   End If
  Next jj
  HACn.Execute "delete from `fax` where `bs_id` = " & idIns, rAF
  For jj = 0 To faxz
   Set rs = Nothing
   rs.Open "select idFax from `fax` where `Fax` = '" & fax(jj) & "' and bs_id = " & idIns, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    HACn.Execute "insert into `fax`(`Fax`,`Faxzahl`,`bs_id`) values('" & fax(jj) & "','" & Replace$(fax(jj), "-", "") & "'," & idIns & ")", rAF
   End If
  Next jj
  HACn.Execute "delete from `mail` where `bs_id` = " & idIns, rAF
  For jj = 0 To mailz
   Set rs = Nothing
   rs.Open "select `idmail` from `mail` where `Mail` = '" & mail(jj) & "' and bs_id = " & idIns, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    HACn.Execute "insert into `mail`(`mail`,`bs_id`) values('" & mail(jj) & "'," & idIns & ")", rAF
   End If
  Next jj
  HACn.CommitTrans
  
  HACn.Execute "update `arzt_has_bs` set aktzeit = " & DatForm(KVDate) & ", obneben = " & BS.obNBS & ", obang = " & BS.obAng & " where `bs_id` = " & idIns & " and `arzt_id` = " & idIns2, rAF
  If rAF <> 1 Then
   HACn.BeginTrans
 '  HACn.Execute "delete from `arzt_has_bs` where `bs_id` = " & idIns & " and `arzt_id` = " & idIns2, rAF
   Set rs = Nothing ' da ein Arzt versehentlich doppelt drin stehen kann (Georg Hochheuser)
   rs.Open "select `obneben` from `arzt_has_bs` where `bs_id` = " & idIns & " and `arzt_id` = " & idIns2, HACn, adOpenStatic, adLockReadOnly
   If rs.BOF Then
    HACn.Execute "insert into `arzt_has_bs`(`bs_id`,`arzt_id`,`obneben`,`obang`,`aktzeit`) values(" & idIns & "," & idIns2 & "," & BS.obNBS & "," & BS.obAng & "," & DatForm(KVDate) & ")", rAF
   End If
   HACn.CommitTrans
  End If
' Plausibilitätskontrollen:
'select * from mail where not mail rlike '^[\-\_\.[:alnum:]]+@[\-\_\.[:alnum:]]+\.(org|de|com|ag|info|net|biz)+$';
'select * from tel where not tel rlike '^[-0123456789]+$';
'select * from fax where not fax rlike '^[-0123456789]+$';
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
'If InStrB(Err.Description, "Doppelter Eintrag '76-61'") <> 0 Then Resume Next
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 HACn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in proTeilnehmer/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' proTeilnehmer

Function indIns&(cn As ADODB.Connection, Tb$, fld$, Wert$, idFld$)
 Dim rs As New ADODB.Recordset, rAF&
 Dim varlen&, pos&
 Dim STyp$
 On Error GoTo fehler
 rs.Open "select * from `" & Tb & "` where `" & fld & "` = '" & doUmwfSQL(Wert, LVobMySQL) & "'", cn, adOpenStatic, adLockReadOnly
 If rs.BOF Then
  Set rs = Nothing
  rs.Open "SELECT data_type dt, character_maximum_length cml FROM information_schema.`COLUMNS` C where table_schema = '" & DBName & "' and table_name = '" & Tb & "' and column_name = '" & fld & "'", cn
'  rs.Open "show full columns from `" & Tb & "` where field = '" & fld & "'", cn
  If Not rs.BOF Then
   STyp = rs!DT
'   If InStrB(STyp, "varchar(") <> 0 Then
   If STyp = "varchar" Then
'    pos = InStr(STyp, "(")
'    varlen = Mid(STyp, pos + 1, InStr(STyp, ")") - pos - 1)
    varlen = rs!cml
    If varlen < Len(Wert) Then
     cn.Execute "Alter Table `" & Tb & "` modify `" & fld & "` varchar(" & Len(Wert) & ")"
    End If
   End If
  End If
  cn.Execute "insert into `" & Tb & "`(`" & fld & "`) values('" & doUmwfSQL(Wert, LVobMySQL) & "')", rAF
  Set rs = cn.Execute("select last_insert_id()")
  indIns = rs.Fields(0)
 Else
  indIns = rs.Fields(idFld)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in indIns/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' indIns

