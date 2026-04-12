Attribute VB_Name = "AnAnpassen"
Option Explicit
Public AbN$() ' AbschnittName
Public AbI$() ' AbschnittInhalt
Public AbIDate As Date
Public G1$(4), G2$(30), G3$(5), G4$(17)

' in alleSpeichern
Function AnTrennZeichen()
 Dim H0 As New CString, h1 As New CString, i&, midakt&, midnext&, j&
 G1(0) = "Wichtige bisherige Krankheiten und Operationen:"
 G1(1) = "Größe:"
 G1(2) = "Früher geraucht"
 G1(3) = "Rauchen Sie noch?"
 G1(4) = "Wie viel Alkohol trinken "
 
 G2(0) = "Teilnahme am DMP:"
 G2(1) = "Diabetes Typ "
 G2(2) = "Grund für Vorstellung "
 G2(3) = "Familienanamnese:"
 G2(4) = "Körpergröße"
 G2(5) = "Diabetes-Medikament 1:"
 G2(6) = "Broteinheiten ca. früh"
 G2(7) = "Essenszeiten ca. früh"
 G2(8) = "Spritz-Eß-Abstand bei"
 G2(9) = "Spritzort früh"
 G2(10) = "Letzte Diabetesschulung"
 G2(11) = "Letztes HbA1c"
 G2(12) = "gemessen am"
 G2(13) = "BZ-Messung selbst?"
 G2(14) = "Durchschnittliche Zahl der BZ-Messungen"
 G2(15) = "Wie hoch ist der BZ meist vor dem Essen:"
 G2(16) = "Wie oft in den letzten 12 Monaten waren Ketoazidosen"
 G2(17) = "Schwangerschaft?"
 G2(18) = "Bluthochdruck?"
 G2(19) = "Wann war die letzte Augenspiegelung?"
 G2(20) = "Ist ein diabetischer Nierenschaden bekannt?"
 G2(21) = "Herzkrankheit?"
 G2(22) = "Hirndurchblutungsstörung?"
 G2(23) = "Beindurchblutungsstörung?"
 G2(24) = "Gefühlsstörung / Ameisenlaufen an den Füßen?"
 G2(25) = "Druckstellen/Schwielen?"
 G2(26) = "Gehen Sie regelmäßig zur Fußpflege?"
 G2(27) = "Haben Sie diabetesgerechte Einlagen/Schuhe?"
 G2(28) = "Neue Fußkomplikationen in den letzten 12 Monaten?"
 G2(29) = "Entleerungsstörung Magen?"
 G2(30) = "Sonstige Folgeerkrankungen:"
 
 G3(0) = "WS:"
 G3(1) = "NNH:"
 G3(2) = "Beinödeme/Venenkrankheiten:"
 G3(3) = "Lymphknoten:"
 G3(4) = "KR:"
 G3(5) = "Neuro sonst:"
 
 G4(0) = "Fußbefund (Inspektion):"
 G4(1) = "Hyperkeratosen u.ä.:"
 G4(2) = "Ulcera:"
 G4(3) = "grobe Kraft Zehenheber:"
 G4(4) = "grobe Kraft Zehenbeuger:"
 G4(5) = "grobe Kraft Knie:"
 G4(6) = "ASR:"
 G4(7) = "PSR:"
 G4(8) = "Oberflächensensibilität:"
 G4(9) = "Monofilamenttest Fußsohle:"
 G4(10) = "Kalt-Warm:"
 G4(11) = "Vibration Innenknöchel:"
 G4(12) = "Vibration Großzehenballen:"
 G4(13) = "Puls in der Leiste:"
 G4(14) = "Puls in der Kniekehle:"
 G4(15) = "Puls der re A.tib.post.:"
 G4(16) = "Puls der re A.dors.ped.:"
 G4(17) = "aktuellen Blutdruck und"
 
 For i = 1 To UBound(rEi)
  Select Case rEi(i).art
   Case "anal", "andm", "andm2", "usal", "usd", "usdm", "usdm1", "usdm2"
    H0.Clear
    h1.Clear
    H0 = rEi(i).Inhalt
    midakt = 1
  End Select
  Select Case rEi(i).art
   Case "anal"
    For j = 0 To UBound(G1)
     midnext = H0.Instr(G1(j))
     If midnext = 0 Or midnext <= midakt Then
      Exit For
     End If
     h1.AppVar Array(H0.cMid(midakt, midnext - midakt), "; ")
     midakt = midnext
    Next j
   Case "andm"
    For j = 0 To UBound(G2)
     midnext = H0.Instr(G2(j))
     If midnext = 0 Or midnext <= midakt Then
      Exit For
     End If
     h1.AppVar Array(H0.cMid(midakt, midnext - midakt), "; ")
     midakt = midnext
    Next j
   Case "usal"
    For j = 0 To UBound(G3)
     midnext = H0.Instr(G3(j))
     If midnext = 0 Or midnext <= midakt Then
      Exit For
     End If
     h1.AppVar Array(H0.cMid(midakt, midnext - midakt), "; ")
     midakt = midnext
    Next j
   Case "usd", "usdm", "usdm1", "usdm2"
    For j = 0 To UBound(G4)
     midnext = H0.Instr(G4(j))
     If midnext = 0 Or midnext <= midakt Then
      Exit For
     End If
     h1.AppVar Array(H0.cMid(midakt, midnext - midakt), "; ")
     midakt = midnext
    Next j
  End Select
  Select Case rEi(i).art
   Case "anal", "andm", "usal", "usd", "usdm", "usdm1", "usdm2"
    If left$(rEi(i).art, 3) <> "usd" Or h1 = "" Then h1.Append H0.cMid(midakt) ' OR h1="" 10.11.19 wg. PID 52832
    rEi(i).Inhalt = h1
  End Select
 Next i
End Function ' AnTrennZeichen

' in alleSpeichern
Function dmpErg()
 Dim trz
 On Error GoTo fehler
 trz = Array("DMP halbj?", _
             ", HbA1c-Ziel:", _
             ", Vertretung:", _
             ", Arztwechsel:", _
             ", Z.d.schw.Hypo.i.Q.:", _
             ", Z.d.KhsEw.w.D.m.i.Q.:", _
             ", D.m.-Schul.empf.:", _
             ", D.m.-Schul.wahrg.:", _
             ", Hypert-Schul.empf.:", _
             ", Hypert-Schul.wahrg:", _
             ", KK-Info Tabak empf.:", _
             ", KK-Info Ernäh. empf:", _
             ", KK-Info körp.Train.empf.:", _
             ", Üw Fußeinricht.:", _
             ", Einw Diab\'klin.:", _
             ", Mitarbeiter:", "")
 Dim Wert$()
 Dim p1&, pe1&, p2&, p2t&, j&, k&, m&, rFai&, WertLen&, npe1k& ' Index auf trz zur Bestimmung der nächsten Trennzeichenlänge für pe1
 For j = 1 To UBound(rEi)
  If rEi(j).art = "dmperg" Then
   Call Kusd(trz, Wert, j)
   
   rFai = 0
   For m = 1 To UBound(rFa)
    If rFa(m).BhFB > rEi(j).Zeitpunkt Then
     rFai = m - 1
     Exit For
    End If
   Next
   If rFai = 0 Then rFai = UBound(rFa)
   
'   Dim C1 AS Class1
'   Dim tliapp, tliinfo
'   SET tliapp = CreateObject("TLI.TLIApplication")
'   IF Not tliapp Is Nothing THEN
'    SET tliinfo = tliapp.TypeInfoFromRecordVariant(C1)
'    IF Not tliinfo Is Nothing THEN
'      Stop ' Look at tliinfo AND analyze the properties
'    END IF
'   END IF
   
'    Dim ri AS RecordInfo
'    SET ri = TLI.TypeInfoFromRecordVariant(rFa)
'    Dim member AS MemberInfo
'    For Each member In ri.Members
'     Dim memberVal AS Variant
'     memberVal = TLI.RecordField(someUDT, member.name)
'     Debug.Print member.name & " : " & memberVal
'    Next member
   
   If (Wert(0) <> "u" Or rFa(rFai).dmphalbj = "") And InStr(Wert(0), "~") = 0 Then rFa(rFai).dmphalbj = Wert(0)
   If (Wert(1) <> "u" Or rFa(rFai).dmpHbA1cZiel = "") And InStr(Wert(1), "~") = 0 Then rFa(rFai).dmpHbA1cZiel = Wert(1)
   If (Wert(2) <> "u" Or rFa(rFai).dmpVertret = "") And InStr(Wert(2), "~") = 0 Then rFa(rFai).dmpVertret = Wert(2)
   If (Wert(3) <> "u" Or rFa(rFai).dmpArztw = "") And InStr(Wert(3), "~") = 0 Then rFa(rFai).dmpArztw = Wert(3)
   If (Wert(4) <> "u" Or rFa(rFai).dmpHypos = "") And InStr(Wert(4), "~") = 0 Then rFa(rFai).dmpHypos = Wert(4)
   If (Wert(5) <> "u" Or rFa(rFai).dmpKhsA = "") And InStr(Wert(5), "~") = 0 Then rFa(rFai).dmpKhsA = Wert(5)
   If (Wert(6) <> "u" Or rFa(rFai).dmpDMSchulEmpf = "") And InStr(Wert(6), "~") = 0 Then rFa(rFai).dmpDMSchulEmpf = Wert(6)
   If (Wert(7) <> "u" Or rFa(rFai).dmpDMSchulWahrg = "") And InStr(Wert(7), "~") = 0 Then rFa(rFai).dmpDMSchulWahrg = Wert(7)
   If (Wert(8) <> "u" Or rFa(rFai).dmpHypertSchulEmpf = "") And InStr(Wert(8), "~") = 0 Then rFa(rFai).dmpHypertSchulEmpf = Wert(8)
   If (Wert(9) <> "u" Or rFa(rFai).dmpHypertSchulWahrg = "") And InStr(Wert(9), "~") = 0 Then rFa(rFai).dmpHypertSchulWahrg = Wert(9)
   If (Wert(10) <> "u" Or rFa(rFai).dmpKKTabakEmpf = "") And InStr(Wert(10), "~") = 0 Then rFa(rFai).dmpKKTabakEmpf = Wert(10)
   If (Wert(11) <> "u" Or rFa(rFai).dmpKKErnEmpf = "") And InStr(Wert(11), "~") = 0 Then rFa(rFai).dmpKKErnEmpf = Wert(11)
   If (Wert(12) <> "u" Or rFa(rFai).dmpKKkTrainEmpf = "") And InStr(Wert(12), "~") = 0 Then rFa(rFai).dmpKKkTrainEmpf = Wert(12)
   If (Wert(13) <> "u" Or rFa(rFai).dmpUewFuss = "") And InStr(Wert(13), "~") = 0 Then rFa(rFai).dmpUewFuss = Wert(13)
   If (Wert(14) <> "u" Or rFa(rFai).dmpEinwDM = "") And InStr(Wert(14), "~") = 0 Then rFa(rFai).dmpEinwDM = Wert(14)
   If (Wert(15) <> "u" Or rFa(rFai).dmpMA = "") And InStr(Wert(15), "~") = 0 Then rFa(rFai).dmpMA = Wert(15)
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpErg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dmperg

' in usdm0, usdm1, usdm2, usfuss, usulcus, usvkgd, usvkgd2 sowie dmperg
' fuer: usdm, fuss, ulcus, vkgd, faelle
Function Kusd(ByRef trz, ByRef Wert$(), j&)
 Dim Offs%(), fehler%(), trzlen%(), k%, lOffs%, letztk%, tzz%
 Dim p0ende%
 On Error GoTo fehler
 tzz = UBound(trz)
 If trz(tzz) = "" Then tzz = tzz - 1 ' Trennzeichenzahl
 ReDim Wert(tzz)
 ReDim Offs(tzz)
 ReDim trzlen(tzz)
 ' für jedes Trennzeichen ...
 For k = 0 To tzz
  trzlen(k) = Len(trz(k)) ' ... Länge bestimmen
  ' ... und Offset des Vorkommens in rei(j).Inhalt bestimmen:
  letztk = k - 1 ' letztk = Index des Trennzeichens, hinter dem gesucht werden soll
  Do
   Do While letztk >= 0 ' wenn das letzte Trennzeichen nicht gefunden wurde, dann ein voriges nehmen
    lOffs = Offs(letztk)
    If lOffs > 0 Then Exit Do
    letztk = letztk - 1
   Loop
   If lOffs = 0 Then lOffs = 1 ' wenn wir noch am Anfang stehen oder kein voriges gefunden wurde
   p0ende = lOffs
   If letztk >= 0 Then p0ende = p0ende + trzlen(letztk)
   lOffs = InStr(p0ende, rEi(j).Inhalt, trz(k))
   letztk = letztk - 1
  Loop Until lOffs > 0 Or letztk < 0
  Offs(k) = lOffs
 Next k
 
 Dim p0%, d0% ' dann ab da nochmal anfangen, danach zu suchen, über Distanz d0
   Dim jj&, kk& ' vorne(kk) oder hinten(jj) verstümmelte Trennzeichen versuchen zu erkennen
   Dim akttrz$
 Dim reihepasst%, fehlerz%, vorfz%, aktuelles%, LK%
 Do While True
  fehlerz = 0
  ReDim fehler(tzz) ' ein Offset ist 0 oder kleiner als der vorige
  LK = 0
  lOffs = Offs(LK) + trzlen(LK)
  For k = 1 To tzz
   If Offs(k) = 0 Then
    fehler(k) = True
    fehlerz = fehlerz + 1
   Else
    If Offs(k) <= lOffs Then
     aktuelles = True
     If k < tzz Then
      If Offs(k + 1) < lOffs And Offs(k + 1) > Offs(k) + trzlen(k) Then
' wenn das vorige die Ordnung stört, dann jenes statt dessen als fehlerhaft markieren
       aktuelles = False
       fehler(LK) = True
       Offs(LK) = 0
       LK = k
       lOffs = Offs(LK) + trzlen(LK)
      End If
     End If
     If aktuelles Then
      fehler(k) = True
      Offs(k) = 0
     End If
     fehlerz = fehlerz + 1
    Else
     LK = k
     lOffs = Offs(LK) + trzlen(LK)
    End If
   End If
  Next k
  If fehlerz = 0 Or (vorfz <> 0 And fehlerz >= vorfz) Then Exit Do Else vorfz = fehlerz
  reihepasst = True
  lOffs = 0
  For k = 0 To tzz
   If fehler(k) Then
    Dim kvor%, knach%
    p0 = 0: d0 = 0
    kvor = k: knach = k
    Do While p0 >= d0
     kvor = kvor - 1: If kvor < 0 Then kvor = 0: p0 = 1 Else p0 = Offs(kvor) + trzlen(kvor)
     If p0 < d0 Then Exit Do
     knach = knach + 1: If knach > tzz Then knach = tzz: d0 = Len(rEi(j).Inhalt) Else d0 = Offs(knach)
    Loop
    reihepasst = False
    Offs(k) = 0
'    IF j = 128 AND k = 5 THEN Stop ' Pat. 23
    For kk = 0 To 2 ' vorne(kk) oder hinten(jj) verstümmelte Trennzeichen versuchen zu erkennen
     akttrz = Mid$(trz(k), kk + 1)
     For jj = Len(akttrz) To 4 Step -1
      akttrz = left$(akttrz, Len(akttrz) - 1)
      Offs(k) = InStr(p0, rEi(j).Inhalt, akttrz)
      If Offs(k) >= d0 Then Offs(k) = 0
      If Offs(k) <> 0 Then Exit For
     Next jj
     If Offs(k) <> 0 Then
      trzlen(k) = Len(akttrz)
      Exit For
     End If
    Next kk
   End If
  Next k
  If reihepasst Then Exit Do
 Loop
 d0 = Len(rEi(j).Inhalt) + 1
 For k = tzz To 0 Step -1
  If Offs(k) <> 0 Then
   p0ende = Offs(k) + trzlen(k)
   If d0 > p0ende Then
    Wert(k) = Trim$(Mid$(rEi(j).Inhalt, p0ende, d0 - p0ende))
    Do While True
     Select Case left$(Wert(k), 1)
      Case ":", ",", ";": Wert(k) = Trim$(Mid$(Wert(k), 2))
      Case Else: Exit Do
     End Select
    Loop
    Do While True
     Select Case Right$(Wert(k), 1)
      Case ":", ",", ";": Wert(k) = Trim$(left$(Wert(k), Len(Wert(k)) - 1)) ' "." nicht wg. "n.u."
      Case Else: Exit Do
     End Select
    Loop
   End If
   d0 = Offs(k)
  End If
 Next k
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KOffs/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Kusd

#If False Then
' in usdm0, usdm1, usdm2, usfuss, usulcus, usvkgd, usvkgd2 sowie dmperg
Function KernUsd(ByRef trz, ByRef Wert$(), j&)
 Dim p1&, pe1&, p2&, p2t&, p3&, k&, m&, rFai&, WertLen&, lWertLen&, npe1k& ' Index auf trz zur Bestimmung der nächsten Trennzeichenlänge für pe1
 Dim lpe1& ' letztes pe1
 Dim Offs%(), lenge%(), trzlen%()
 
 On Error GoTo fehler
   ReDim Wert(UBound(trz) - 1)
   ReDim Offs(UBound(Wert))
   ReDim lenge(UBound(Wert))
   ReDim trzlen(UBound(Wert))
   Dim naefehlt%
   For k = 0 To UBound(Wert)
    If p2 = p3 And p2 <> 0 Then naefehlt = True Else naefehlt = 0 ' wenn eins ganz fehlt, z.B. ", Puls" in vkgd
    trzlen(k) = Len(trz(k))
    If k = 0 Then p1 = InStr(rEi(j).Inhalt, trz(k)): npe1k = k Else p1 = p2
    lpe1 = pe1
    pe1 = p1 + Len(trz(npe1k))
    npe1k = k + 1
'    IF k = 4 THEN Stop
    If k = UBound(Wert) Then
     p2 = Len(rEi(j).Inhalt) + 1
    Else
     p2 = InStr(pe1, rEi(j).Inhalt, trz(npe1k))
     If p2 = 0 And k < UBound(Wert) - 1 Then p2 = InStr(pe1, rEi(j).Inhalt, trz(k + 2))
    End If
    ' wenn ein Trennzeichen wie ", li:" aktuell verstümmelt, aber später nochmal unverstümmelt vorkommt, z.B. pat_id 52040
    ' geht dort noch nicht, wo trz(npe1k+1) auch verstümmelt :-)
    If k < UBound(Wert) - 1 Then p3 = InStr(pe1, rEi(j).Inhalt, trz(npe1k + 1)) Else p3 = Len(rEi(j).Inhalt) + 1
    If p2 = 0 Or (p3 <> 0 And p2 > p3) Then
     If lpe1 = 0 Then lpe1 = 1
     If k = UBound(Wert) Then
      p2 = InStr(lpe1, rEi(j).Inhalt, trz(npe1k))
     Else
      p2 = Len(rEi(j).Inhalt) + 1 ' +1=31.8.17, Pat. 2848
     End If
     If p2 = 0 Then
      Dim jj&, kk& ' vorne(kk) oder hinten(jj) verstümmelte Trennzeichen versuchen zu erkennen
      Dim akttrz$
      For kk = 0 To 2
       akttrz = Mid$(trz(npe1k), kk + 1)
       For jj = Len(akttrz) To 4 Step -1
        akttrz = left$(akttrz, Len(akttrz) - 1)
        p2 = InStr(lpe1, rEi(j).Inhalt, akttrz)
        If p2 <> 0 Then
         Exit For
        End If
       Next jj
       If p2 <> 0 Then
        trzlen(npe1k) = Len(akttrz)
        lpe1 = pe1 ' 7.9.15
        Exit For
       End If
      Next kk
     Else ' 7.9.15 pe1 wird (hoffentlich immer) weiter unten korrigiert
'      Stop
     End If
     pe1 = lpe1 ' 6.9.15
    End If
    If p2 = 0 And k < UBound(Wert) Then
'     IF k > 0 THEN Wert(k) = Wert(k - 1)
     npe1k = npe1k + 1
     p2 = InStr(pe1, rEi(j).Inhalt, trz(npe1k)) ' Korrektur für übernächstes Feld, falls z.B. Doppelpunkt gelöscht
    ElseIf k <= UBound(Wert) And k <= UBound(trz) - 2 Then
     If p2 > pe1 + 40 Then ' wenn mehr als 40 Buchstaben, Überprüfung, für aktuelles Feld, ob z.B. Doppelpunkt gelöscht
      p2t = InStr(pe1, rEi(j).Inhalt, trz(k + 2))
      If p2t > 0 And p2t < p2 Then
       npe1k = npe1k + 1
       p2 = p2t ' 7.9.15
       If k > 0 Then If pe1 < Offs(k - 1) + lenge(k - 1) + trzlen(k) Then pe1 = Offs(k - 1) + lenge(k - 1) + trzlen(k)
      End If
     End If
    End If
    lWertLen = WertLen
    If p2 >= pe1 Then
     WertLen = p2 - pe1
     If Not naefehlt Then
      Wert(k) = Trim$(Mid$(rEi(j).Inhalt, pe1, WertLen))
      Offs(k) = pe1
     Else
      If rEi(j).art <> "vkgd" And rEi(j).art <> "vkg" Then
       MsgBox "Unerwarteter Zustand in KernUsd: rEi(" & j & ").art <> 'vkgd'/'vkg' (" & rEi(j).vkgd & ")"
      End If
      p2 = p1 ' wenn eins ausgelassen war, dann das darauf folgende noch berücksichtigen
     End If
    Else
     WertLen = Len(rEi(j).Inhalt)
     Wert(k) = Trim$(Mid$(rEi(j).Inhalt, pe1, WertLen))
     Offs(k) = pe1
    End If
    lenge(k) = WertLen
    If npe1k < UBound(trz) Then
     If LenB(trz(npe1k + 1)) <> 0 Then
      Dim ünäpos&
      ünäpos = InStr(Wert(k), trz(npe1k + 1))
      If ünäpos > 0 Then Wert(k) = left$(Wert(k), ünäpos - 1): lenge(k) = ünäpos - 1
     End If
    End If
    If npe1k = k + 2 Then
     k = k + 1
     If k <= UBound(Wert) Then Wert(k) = Wert(k - 1): Offs(k) = Offs(k - 1): lenge(k) = lenge(k - 1)
    End If
   Next k
   ' dass kein Feld über das nächste hinausreicht ...
   For k = UBound(Wert) To 1 Step -1
    If Offs(k - 1) = Offs(k) Then
     If lenge(k - 1) > lenge(k) Then
      Wert(k - 1) = Wert(k)
     End If
    Else
     If Offs(k - 1) + lenge(k - 1) + trzlen(k) > Offs(k) And Offs(k) - trzlen(k) > 0 Then
      Wert(k - 1) = left$(Wert(k - 1), Offs(k) - trzlen(k))
     End If
    End If
   Next k
   For k = UBound(Wert) To 0 Step -1
    If Right$(Wert(k), 1) = ";" Then Wert(k) = Trim$(left$(Wert(k), Len(Wert(k)) - 1))
   Next k
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kernusb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' kernusd
#End If 'false

Function rUsRest(j&)
   rUs(UBound(rUs)).Pat_id = rEi(j).Pat_id
   rUs(UBound(rUs)).FID = rEi(j).FID
   rUs(UBound(rUs)).Zeitpunkt = rEi(j).Zeitpunkt
   rUs(UBound(rUs)).art = rEi(j).art
   rUs(UBound(rUs)).absPos = rEi(j).absPos
   rUs(UBound(rUs)).QS = rEi(j).QS
   rUs(UBound(rUs)).QT = rEi(j).QT
   rUs(UBound(rUs)).StByte = rEi(j).StByte
   rUs(UBound(rUs)).aktZeit = rEi(j).aktZeit
End Function ' rUsRest

Function rFuRest(j&)
   rFu(UBound(rFu)).Pat_id = rEi(j).Pat_id
   rFu(UBound(rFu)).FID = rEi(j).FID
   rFu(UBound(rFu)).Zeitpunkt = rEi(j).Zeitpunkt
   rFu(UBound(rFu)).art = rEi(j).art
   rFu(UBound(rFu)).absPos = rEi(j).absPos
   rFu(UBound(rFu)).QS = rEi(j).QS
   rFu(UBound(rFu)).QT = rEi(j).QT
   rFu(UBound(rFu)).StByte = rEi(j).StByte
   rFu(UBound(rFu)).aktZeit = rEi(j).aktZeit
End Function ' rUsRest

Function rVkRest(j&)
   rVk(UBound(rVk)).Pat_id = rEi(j).Pat_id
   rVk(UBound(rVk)).FID = rEi(j).FID
   rVk(UBound(rVk)).Zeitpunkt = rEi(j).Zeitpunkt
'   rvk(UBound(rvk)).Art = rEi(j).Art
   rVk(UBound(rVk)).absPos = rEi(j).absPos
'   rvk(UBound(rvk)).QS = rEi(j).QS
'   rvk(UBound(rvk)).QT = rEi(j).QT
   rVk(UBound(rVk)).StByte = rEi(j).StByte
   rVk(UBound(rVk)).aktZeit = rEi(j).aktZeit
End Function ' rUsRest

Function rUlRest(j&)
   rUl(UBound(rUl)).Pat_id = rEi(j).Pat_id
   rUl(UBound(rUl)).FID = rEi(j).FID
   rUl(UBound(rUl)).Zeitpunkt = rEi(j).Zeitpunkt
'   rUl(UBound(rUl)).Art = rEi(j).Art
   rUl(UBound(rUl)).absPos = rEi(j).absPos
'   rUl(UBound(rUl)).QS = rEi(j).QS
'   rUl(UBound(rUl)).QT = rEi(j).QT
   rUl(UBound(rUl)).StByte = rEi(j).StByte
   rUl(UBound(rUl)).aktZeit = rEi(j).aktZeit
End Function ' rUsRest

' in alleSpeichern
Function usdm0()
 Dim trz, j&
 Dim trzn
 On Error GoTo fehler
 trz = Array("Liphypertrophien: re:", ", li:", _
             "Fußbefund (Inspektion): re:", "li:", _
             "Hyperkeratosen u.ä.: re:", ", li:", _
             "Ulcera: re:", ", li", _
             "grobe Kraft Zehenheber: re:", "li:", _
             "grobe Kraft Zehenbeuger: re:", ", li:", _
             "grobe Kraft Knie: re:", ", li:", _
             "ASR: re:", ", li:", _
             "PSR: re:", ", li:", _
             "Oberflächensensibilität: re:", ", li:", _
             "Monofilamenttest Fußsohle: re:", ", li:", _
             "Kalt-Warm: re:", ", li:", _
             "Vibration Innenknöchel: re:", ", li:", _
             "Vibration Großzehenballen: re:", ", li:", _
             "Puls in der Leiste: re:", ", li:", _
             "Puls in der Kniekehle: re:", ", li:", _
             "Puls der re A.tib.post.: re:", ", li:", _
             "Puls der re A.dors.ped.: re:", ", li:", _
             "aktuellen Blutdruck und ggf. Puls bitte extra eingeben; Mitarbeiter:", "")
 trzn = Array("Spritzstellen (Bauch/OS):", ", Ort:", _
              "; Kraft Zehenheber: re:", "; li:", _
              ";  Kraft Zehenbeuger: re:", "; li:", _
              "; Kraft Knie: re:", "; li:", _
              ";  ASR (Ferse): re:", "; li:", _
              "; PSR (Knie): re:", "; li:", _
              "; Monofilament Fußsohle(/5): re:", "; li:", _
              "; Kalt-Warm(/5): re:", "; li:", _
              "; Vibr. Innenknö. (/8); re:", "; li:", _
              ";  Vibr. Großzehenballen (/8): re:", "; li:", _
              "; Puls Leiste: re:", "; li:", _
              "; Puls Kniekehle: re:", "; li:", _
              "; Puls Atp.(Innenkn): re:", "; li:", _
              "; Puls Adp.(Fußrü): re:", "; li:", _
              "; Mitarbeiter:", "")
             
 For j = 1 To UBound(rEi)
  If LCase$(rEi(j).art) = "usdm" And rEi(j).Zeitpunkt < #3/18/2025# Then
   If rUs(UBound(rUs)).absPos <> 0 Or UBound(rUs) = 0 Then ReDim Preserve rUs(UBound(rUs) + 1)
   Dim Wert$()
'   IF j = 198 THEN Stop ' Pid 1115
   Call Kusd(IIf(rEi(j).absPos = -1, trzn, trz), Wert, j)
   Call rUsRest(j)
   If UBound(Wert) = 28 Then
    If (Wert(0) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(0), "~") = 0 Or _
      (Wert(1) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(1), "~") = 0 _
      Then rUs(UBound(rUs)).Spritzst = Wert(0) & "|" & Wert(1)
    If (Wert(2) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_re = "") And InStr(Wert(2), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_re = Wert(2)
    If (Wert(3) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_li = "") And InStr(Wert(3), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_li = Wert(3)
    If (Wert(4) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_re = "") And InStr(Wert(4), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_re = Wert(4)
    If (Wert(5) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_li = "") And InStr(Wert(5), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_li = Wert(5)
    If (Wert(6) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_re = "") And InStr(Wert(6), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_re = Wert(6)
    If (Wert(7) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_li = "") And InStr(Wert(7), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_li = Wert(7)
    If (Wert(8) <> "u" Or rUs(UBound(rUs)).ASR_re = "") And InStr(Wert(8), "~") = 0 Then rUs(UBound(rUs)).ASR_re = Wert(8)
    If (Wert(9) <> "u" Or rUs(UBound(rUs)).ASR_li = "") And InStr(Wert(9), "~") = 0 Then rUs(UBound(rUs)).ASR_li = Wert(9)
    If (Wert(10) <> "u" Or rUs(UBound(rUs)).PSR_re = "") And InStr(Wert(10), "~") = 0 Then rUs(UBound(rUs)).PSR_re = Wert(10)
    If (Wert(11) <> "u" Or rUs(UBound(rUs)).PSR_li = "") And InStr(Wert(11), "~") = 0 Then rUs(UBound(rUs)).PSR_li = Wert(11)
    If (Wert(12) <> "u" Or rUs(UBound(rUs)).MF_re = "") And InStr(Wert(12), "~") = 0 Then rUs(UBound(rUs)).MF_re = Wert(12) & IIf(IsNumeric(Wert(12)), "/5", "")
    If (Wert(13) <> "u" Or rUs(UBound(rUs)).MF_li = "") And InStr(Wert(13), "~") = 0 Then rUs(UBound(rUs)).MF_li = Wert(13) & IIf(IsNumeric(Wert(13)), "/5", "")
    If (Wert(14) <> "u" Or rUs(UBound(rUs)).KW_re = "") And InStr(Wert(14), "~") = 0 Then rUs(UBound(rUs)).KW_re = Wert(14) & IIf(IsNumeric(Wert(14)), "/5", "")
    If (Wert(15) <> "u" Or rUs(UBound(rUs)).KW_li = "") And InStr(Wert(15), "~") = 0 Then rUs(UBound(rUs)).KW_li = Wert(15) & IIf(IsNumeric(Wert(15)), "/5", "")
    If (Wert(16) <> "u" Or rUs(UBound(rUs)).Vibr_IK_re = "") And InStr(Wert(16), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_re = Wert(16) & IIf(IsNumeric(Wert(16)), "/8", "")
    If (Wert(17) <> "u" Or rUs(UBound(rUs)).Vibr_IK_li = "") And InStr(Wert(17), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_li = Wert(17) & IIf(IsNumeric(Wert(17)), "/8", "")
    If (Wert(18) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_re = "") And InStr(Wert(18), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_re = Wert(18) & IIf(IsNumeric(Wert(18)), "/8", "")
    If (Wert(19) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_li = "") And InStr(Wert(19), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_li = Wert(19) & IIf(IsNumeric(Wert(19)), "/8", "")
    If (Wert(20) <> "u" Or rUs(UBound(rUs)).PulsL_re = "") And InStr(Wert(20), "~") = 0 Then rUs(UBound(rUs)).PulsL_re = Wert(20)
    If (Wert(21) <> "u" Or rUs(UBound(rUs)).PulsL_li = "") And InStr(Wert(21), "~") = 0 Then rUs(UBound(rUs)).PulsL_li = Wert(21)
    If (Wert(22) <> "u" Or rUs(UBound(rUs)).PulsKK_re = "") And InStr(Wert(22), "~") = 0 Then rUs(UBound(rUs)).PulsKK_re = Wert(22)
    If (Wert(23) <> "u" Or rUs(UBound(rUs)).PulsKK_li = "") And InStr(Wert(23), "~") = 0 Then rUs(UBound(rUs)).PulsKK_li = Wert(23)
    If (Wert(24) <> "u" Or rUs(UBound(rUs)).PulsAtp_re = "") And InStr(Wert(24), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_re = Wert(24)
    If (Wert(25) <> "u" Or rUs(UBound(rUs)).PulsAtp_li = "") And InStr(Wert(25), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_li = Wert(25)
    If (Wert(26) <> "u" Or rUs(UBound(rUs)).PulsAdp_re = "") And InStr(Wert(26), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_re = Wert(26)
    If (Wert(27) <> "u" Or rUs(UBound(rUs)).PulsAdp_li = "") And InStr(Wert(27), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_li = Wert(27)
    If (Wert(28) <> "u" Or rUs(UBound(rUs)).Mitarbeiter = "") And InStr(Wert(28), "~") = 0 Then rUs(UBound(rUs)).Mitarbeiter = Wert(28)
   Else
    If (Wert(0) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(0), "~") = 0 Or _
      (Wert(1) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(1), "~") = 0 _
      Then rUs(UBound(rUs)).Spritzst = Wert(0) & "|" & Wert(1)
    If (Wert(2) <> "u" Or rUs(UBound(rUs)).Fußbef_re = "") And InStr(Wert(2), "~") = 0 Then rUs(UBound(rUs)).Fußbef_re = Wert(2)
    If (Wert(3) <> "u" Or rUs(UBound(rUs)).Fußbef_li = "") And InStr(Wert(3), "~") = 0 Then rUs(UBound(rUs)).Fußbef_li = Wert(3)
    If (Wert(4) <> "u" Or rUs(UBound(rUs)).Hyperk_re = "") And InStr(Wert(4), "~") = 0 Then rUs(UBound(rUs)).Hyperk_re = Wert(4)
    If (Wert(5) <> "u" Or rUs(UBound(rUs)).Hyperk_li = "") And InStr(Wert(5), "~") = 0 Then rUs(UBound(rUs)).Hyperk_li = Wert(5)
    If (Wert(6) <> "u" Or rUs(UBound(rUs)).Ulcera_re = "") And InStr(Wert(6), "~") = 0 Then rUs(UBound(rUs)).Ulcera_re = Wert(6)
    If (Wert(7) <> "u" Or rUs(UBound(rUs)).Ulcera_li = "") And InStr(Wert(7), "~") = 0 Then rUs(UBound(rUs)).Ulcera_li = Wert(7)
    If (Wert(8) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_re = "") And InStr(Wert(8), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_re = Wert(8)
    If (Wert(9) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_li = "") And InStr(Wert(9), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_li = Wert(9)
    If (Wert(10) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_re = "") And InStr(Wert(10), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_re = Wert(10)
    If (Wert(11) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_li = "") And InStr(Wert(11), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_li = Wert(11)
    If (Wert(12) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_re = "") And InStr(Wert(12), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_re = Wert(12)
    If (Wert(13) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_li = "") And InStr(Wert(13), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_li = Wert(13)
    If (Wert(14) <> "u" Or rUs(UBound(rUs)).ASR_re = "") And InStr(Wert(14), "~") = 0 Then rUs(UBound(rUs)).ASR_re = Wert(14)
    If (Wert(15) <> "u" Or rUs(UBound(rUs)).ASR_li = "") And InStr(Wert(15), "~") = 0 Then rUs(UBound(rUs)).ASR_li = Wert(15)
    If (Wert(16) <> "u" Or rUs(UBound(rUs)).PSR_re = "") And InStr(Wert(16), "~") = 0 Then rUs(UBound(rUs)).PSR_re = Wert(16)
    If (Wert(17) <> "u" Or rUs(UBound(rUs)).PSR_li = "") And InStr(Wert(17), "~") = 0 Then rUs(UBound(rUs)).PSR_li = Wert(17)
    If (Wert(18) <> "u" Or rUs(UBound(rUs)).Oberfl_re = "") And InStr(Wert(18), "~") = 0 Then rUs(UBound(rUs)).Oberfl_re = Wert(18)
    If (Wert(19) <> "u" Or rUs(UBound(rUs)).Oberfl_li = "") And InStr(Wert(19), "~") = 0 Then rUs(UBound(rUs)).Oberfl_li = Wert(19)
    If (Wert(20) <> "u" Or rUs(UBound(rUs)).MF_re = "") And InStr(Wert(20), "~") = 0 Then rUs(UBound(rUs)).MF_re = Wert(20) & IIf(IsNumeric(Wert(20)), "/5", "")
    If (Wert(21) <> "u" Or rUs(UBound(rUs)).MF_li = "") And InStr(Wert(21), "~") = 0 Then rUs(UBound(rUs)).MF_li = Wert(21) & IIf(IsNumeric(Wert(21)), "/5", "")
    If (Wert(22) <> "u" Or rUs(UBound(rUs)).KW_re = "") And InStr(Wert(22), "~") = 0 Then rUs(UBound(rUs)).KW_re = Wert(22) & IIf(IsNumeric(Wert(22)), "/5", "")
    If (Wert(23) <> "u" Or rUs(UBound(rUs)).KW_li = "") And InStr(Wert(23), "~") = 0 Then rUs(UBound(rUs)).KW_li = Wert(23) & IIf(IsNumeric(Wert(23)), "/5", "")
    If (Wert(24) <> "u" Or rUs(UBound(rUs)).Vibr_IK_re = "") And InStr(Wert(24), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_re = Wert(24) & IIf(IsNumeric(Wert(24)), "/8", "")
    If (Wert(25) <> "u" Or rUs(UBound(rUs)).Vibr_IK_li = "") And InStr(Wert(25), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_li = Wert(25) & IIf(IsNumeric(Wert(25)), "/8", "")
    If (Wert(26) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_re = "") And InStr(Wert(26), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_re = Wert(26) & IIf(IsNumeric(Wert(26)), "/8", "")
    If (Wert(27) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_li = "") And InStr(Wert(27), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_li = Wert(27) & IIf(IsNumeric(Wert(27)), "/8", "")
    If (Wert(28) <> "u" Or rUs(UBound(rUs)).PulsL_re = "") And InStr(Wert(28), "~") = 0 Then rUs(UBound(rUs)).PulsL_re = Wert(28)
    If (Wert(29) <> "u" Or rUs(UBound(rUs)).PulsL_li = "") And InStr(Wert(29), "~") = 0 Then rUs(UBound(rUs)).PulsL_li = Wert(29)
    If (Wert(30) <> "u" Or rUs(UBound(rUs)).PulsKK_re = "") And InStr(Wert(30), "~") = 0 Then rUs(UBound(rUs)).PulsKK_re = Wert(30)
    If (Wert(31) <> "u" Or rUs(UBound(rUs)).PulsKK_li = "") And InStr(Wert(31), "~") = 0 Then rUs(UBound(rUs)).PulsKK_li = Wert(31)
    If (Wert(32) <> "u" Or rUs(UBound(rUs)).PulsAtp_re = "") And InStr(Wert(32), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_re = Wert(32)
    If (Wert(33) <> "u" Or rUs(UBound(rUs)).PulsAtp_li = "") And InStr(Wert(33), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_li = Wert(33)
    If (Wert(34) <> "u" Or rUs(UBound(rUs)).PulsAdp_re = "") And InStr(Wert(34), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_re = Wert(34)
    If (Wert(35) <> "u" Or rUs(UBound(rUs)).PulsAdp_li = "") And InStr(Wert(35), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_li = Wert(35)
    If (Wert(36) <> "u" Or rUs(UBound(rUs)).Mitarbeiter = "") And InStr(Wert(36), "~") = 0 Then rUs(UBound(rUs)).Mitarbeiter = Wert(36)
    If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
   End If ' UBound(Wert) >= 36 Then
  End If ' LCase$(rEi(j).art) = "usdm" Then
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdm0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usdm0

Function usdm1()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Spritzstellen (Bauch/OS):", _
             ", Fußbefund (Inspektion): re:", ", li:", _
             ", Hyperkeratosen u.ä.: re:", ", li:", _
             ", Ulcera: re:", ", li:", _
             ", Kraft Zehenheber: re:", ", li:", _
             ", Kraft Zehenbeuger: re:", ", li:", _
             ", Kraft Knie: re:", ", li:", _
             ", ASR (Ferse): re:", ", li:", _
             ", PSR (Knie): re:", ", li:", _
             ", Monofilament Fußsohle(/5): re:", ", li:", _
             ", Kalt-Warm(/5): re:", ", li:", _
             ", Vibr. Innenknö. (/8): re:", ", li:", _
             ", Vibr. Großzehenballen (/8): re:", ", li:", _
             ", Puls Leiste: re:", ", li:", _
             ", Puls Kniekehle: re:", ", li:", _
             ", Puls Atp.(Innenkn): re:", ", li:", _
             ", Puls Adp.(Fußrü): re:", ", li:", _
             ", Mitarbeiter:", "")
 For j = 1 To UBound(rEi)
  If rEi(j).art = "usdm1" Then
   If rUs(UBound(rUs)).absPos <> 0 Or UBound(rUs) = 0 Then ReDim Preserve rUs(UBound(rUs) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rUsRest(j)
   If (Wert(0) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(0), "~") = 0 Then rUs(UBound(rUs)).Spritzst = Wert(0)
   If (Wert(1) <> "u" Or rUs(UBound(rUs)).Fußbef_re = "") And InStr(Wert(1), "~") = 0 Then rUs(UBound(rUs)).Fußbef_re = Wert(1)
   If (Wert(2) <> "u" Or rUs(UBound(rUs)).Fußbef_li = "") And InStr(Wert(2), "~") = 0 Then rUs(UBound(rUs)).Fußbef_li = Wert(2)
   If (Wert(3) <> "u" Or rUs(UBound(rUs)).Hyperk_re = "") And InStr(Wert(3), "~") = 0 Then rUs(UBound(rUs)).Hyperk_re = Wert(3)
   If (Wert(4) <> "u" Or rUs(UBound(rUs)).Hyperk_li = "") And InStr(Wert(4), "~") = 0 Then rUs(UBound(rUs)).Hyperk_li = Wert(4)
   If (Wert(5) <> "u" Or rUs(UBound(rUs)).Ulcera_re = "") And InStr(Wert(5), "~") = 0 Then rUs(UBound(rUs)).Ulcera_re = Wert(5)
   If (Wert(6) <> "u" Or rUs(UBound(rUs)).Ulcera_li = "") And InStr(Wert(6), "~") = 0 Then rUs(UBound(rUs)).Ulcera_li = Wert(6)
   If (Wert(7) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_re = "") And InStr(Wert(7), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_re = Wert(7)
   If (Wert(8) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_li = "") And InStr(Wert(8), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_li = Wert(8)
   If (Wert(9) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_re = "") And InStr(Wert(9), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_re = Wert(9)
   If (Wert(10) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_li = "") And InStr(Wert(10), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_li = Wert(10)
   If (Wert(11) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_re = "") And InStr(Wert(11), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_re = Wert(11)
   If (Wert(12) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_li = "") And InStr(Wert(12), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_li = Wert(12)
   If (Wert(13) <> "u" Or rUs(UBound(rUs)).ASR_re = "") And InStr(Wert(13), "~") = 0 Then rUs(UBound(rUs)).ASR_re = Wert(13)
   If (Wert(14) <> "u" Or rUs(UBound(rUs)).ASR_li = "") And InStr(Wert(14), "~") = 0 Then rUs(UBound(rUs)).ASR_li = Wert(14)
   If (Wert(15) <> "u" Or rUs(UBound(rUs)).PSR_re = "") And InStr(Wert(15), "~") = 0 Then rUs(UBound(rUs)).PSR_re = Wert(15)
   If (Wert(16) <> "u" Or rUs(UBound(rUs)).PSR_li = "") And InStr(Wert(16), "~") = 0 Then rUs(UBound(rUs)).PSR_li = Wert(16)
   If (Wert(17) <> "u" Or rUs(UBound(rUs)).MF_re = "") And InStr(Wert(17), "~") = 0 Then rUs(UBound(rUs)).MF_re = Wert(17) & IIf(IsNumeric(Wert(17)), "/5", "")
   If (Wert(18) <> "u" Or rUs(UBound(rUs)).MF_li = "") And InStr(Wert(18), "~") = 0 Then rUs(UBound(rUs)).MF_li = Wert(18) & IIf(IsNumeric(Wert(18)), "/5", "")
   If (Wert(19) <> "u" Or rUs(UBound(rUs)).KW_re = "") And InStr(Wert(19), "~") = 0 Then rUs(UBound(rUs)).KW_re = Wert(19) & IIf(IsNumeric(Wert(19)), "/5", "")
   If (Wert(20) <> "u" Or rUs(UBound(rUs)).KW_li = "") And InStr(Wert(20), "~") = 0 Then rUs(UBound(rUs)).KW_li = Wert(20) & IIf(IsNumeric(Wert(20)), "/5", "")
   If (Wert(21) <> "u" Or rUs(UBound(rUs)).Vibr_IK_re = "") And InStr(Wert(21), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_re = Wert(21) & IIf(IsNumeric(Wert(21)), "/8", "")
   If (Wert(22) <> "u" Or rUs(UBound(rUs)).Vibr_IK_li = "") And InStr(Wert(22), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_li = Wert(22) & IIf(IsNumeric(Wert(22)), "/8", "")
   If (Wert(23) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_re = "") And InStr(Wert(23), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_re = Wert(23) & IIf(IsNumeric(Wert(23)), "/8", "")
   If (Wert(24) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_li = "") And InStr(Wert(24), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_li = Wert(24) & IIf(IsNumeric(Wert(24)), "/8", "")
   If (Wert(25) <> "u" Or rUs(UBound(rUs)).PulsL_re = "") And InStr(Wert(25), "~") = 0 Then rUs(UBound(rUs)).PulsL_re = Wert(25)
   If (Wert(26) <> "u" Or rUs(UBound(rUs)).PulsL_li = "") And InStr(Wert(26), "~") = 0 Then rUs(UBound(rUs)).PulsL_li = Wert(26)
   If (Wert(27) <> "u" Or rUs(UBound(rUs)).PulsKK_re = "") And InStr(Wert(27), "~") = 0 Then rUs(UBound(rUs)).PulsKK_re = Wert(27)
   If (Wert(28) <> "u" Or rUs(UBound(rUs)).PulsKK_li = "") And InStr(Wert(28), "~") = 0 Then rUs(UBound(rUs)).PulsKK_li = Wert(28)
   If (Wert(29) <> "u" Or rUs(UBound(rUs)).PulsAtp_re = "") And InStr(Wert(29), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_re = Wert(29)
   If (Wert(30) <> "u" Or rUs(UBound(rUs)).PulsAtp_li = "") And InStr(Wert(30), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_li = Wert(30)
   If (Wert(31) <> "u" Or rUs(UBound(rUs)).PulsAdp_re = "") And InStr(Wert(31), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_re = Wert(31)
   If (Wert(32) <> "u" Or rUs(UBound(rUs)).PulsAdp_li = "") And InStr(Wert(32), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_li = Wert(32)
   If (Wert(33) <> "u" Or rUs(UBound(rUs)).Mitarbeiter = "") And InStr(Wert(33), "~") = 0 Then rUs(UBound(rUs)).Mitarbeiter = Wert(33)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdm1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usdm1

Function usdm2()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Spritzstellen (Bauch/OS):", _
             "; Kraft Zehenheber: re:", "; li:", _
             ";  Kraft Zehenbeuger: re:", "; li:", _
             "; Kraft Knie: re:", "; li:", _
             ";  ASR (Ferse): re:", "; li:", _
             "; PSR (Knie): re:", "; li:", _
             "; Monofilament Fußsohle(/5): re:", "; li:", _
             "; Kalt-Warm(/5): re:", "; li:", _
             "; Vibr. Innenknö. (/8); re:", "; li:", _
             ";  Vibr. Großzehenballen (/8): re:", "; li:", _
             "; Puls Leiste: re:", "; li:", _
             "; Puls Kniekehle: re:", "; li:", _
             "; Puls Atp.(Innenkn): re:", "; li:", _
             "; Puls Adp.(Fußrü): re:", "; li:", _
             "; Mitarbeiter:", "")
 For j = 1 To UBound(rEi)
  If rEi(j).art = "usdm2" Or (rEi(j).art = "usd" And rEi(j).Zeitpunkt < #3/18/2025#) Then
   If rUs(UBound(rUs)).absPos <> 0 Or UBound(rUs) = 0 Then ReDim Preserve rUs(UBound(rUs) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rUsRest(j)
   If (Wert(0) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(0), "~") = 0 Then rUs(UBound(rUs)).Spritzst = Wert(0)
   If (Wert(1) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_re = "") And InStr(Wert(1), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_re = Wert(1)
   If (Wert(2) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_li = "") And InStr(Wert(2), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_li = Wert(2)
   If (Wert(3) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_re = "") And InStr(Wert(3), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_re = Wert(3)
   If (Wert(4) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_li = "") And InStr(Wert(4), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_li = Wert(4)
   If (Wert(5) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_re = "") And InStr(Wert(5), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_re = Wert(5)
   If (Wert(6) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_li = "") And InStr(Wert(6), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_li = Wert(6)
   If (Wert(7) <> "u" Or rUs(UBound(rUs)).ASR_re = "") And InStr(Wert(7), "~") = 0 Then rUs(UBound(rUs)).ASR_re = Wert(7)
   If (Wert(8) <> "u" Or rUs(UBound(rUs)).ASR_li = "") And InStr(Wert(8), "~") = 0 Then rUs(UBound(rUs)).ASR_li = Wert(8)
   If (Wert(9) <> "u" Or rUs(UBound(rUs)).PSR_re = "") And InStr(Wert(9), "~") = 0 Then rUs(UBound(rUs)).PSR_re = Wert(9)
   If (Wert(10) <> "u" Or rUs(UBound(rUs)).PSR_li = "") And InStr(Wert(10), "~") = 0 Then rUs(UBound(rUs)).PSR_li = Wert(10)
   If (Wert(11) <> "u" Or rUs(UBound(rUs)).MF_re = "") And InStr(Wert(11), "~") = 0 Then rUs(UBound(rUs)).MF_re = Wert(11) & IIf(IsNumeric(Wert(11)), "/5", "")
   If (Wert(12) <> "u" Or rUs(UBound(rUs)).MF_li = "") And InStr(Wert(12), "~") = 0 Then rUs(UBound(rUs)).MF_li = Wert(12) & IIf(IsNumeric(Wert(12)), "/5", "")
   If (Wert(13) <> "u" Or rUs(UBound(rUs)).KW_re = "") And InStr(Wert(13), "~") = 0 Then rUs(UBound(rUs)).KW_re = Wert(13) & IIf(IsNumeric(Wert(13)), "/5", "")
   If (Wert(14) <> "u" Or rUs(UBound(rUs)).KW_li = "") And InStr(Wert(14), "~") = 0 Then rUs(UBound(rUs)).KW_li = Wert(14) & IIf(IsNumeric(Wert(14)), "/5", "")
   If (Wert(15) <> "u" Or rUs(UBound(rUs)).Vibr_IK_re = "") And InStr(Wert(15), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_re = Wert(15) & IIf(IsNumeric(Wert(15)), "/8", "")
   If (Wert(16) <> "u" Or rUs(UBound(rUs)).Vibr_IK_li = "") And InStr(Wert(16), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_li = Wert(16) & IIf(IsNumeric(Wert(16)), "/8", "")
   If (Wert(17) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_re = "") And InStr(Wert(17), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_re = Wert(17) & IIf(IsNumeric(Wert(17)), "/8", "")
   If (Wert(18) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_li = "") And InStr(Wert(18), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_li = Wert(18) & IIf(IsNumeric(Wert(18)), "/8", "")
   If (Wert(19) <> "u" Or rUs(UBound(rUs)).PulsL_re = "") And InStr(Wert(19), "~") = 0 Then rUs(UBound(rUs)).PulsL_re = Wert(19)
   If (Wert(20) <> "u" Or rUs(UBound(rUs)).PulsL_li = "") And InStr(Wert(20), "~") = 0 Then rUs(UBound(rUs)).PulsL_li = Wert(20)
   If (Wert(21) <> "u" Or rUs(UBound(rUs)).PulsKK_re = "") And InStr(Wert(21), "~") = 0 Then rUs(UBound(rUs)).PulsKK_re = Wert(21)
   If (Wert(22) <> "u" Or rUs(UBound(rUs)).PulsKK_li = "") And InStr(Wert(22), "~") = 0 Then rUs(UBound(rUs)).PulsKK_li = Wert(22)
   If (Wert(23) <> "u" Or rUs(UBound(rUs)).PulsAtp_re = "") And InStr(Wert(23), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_re = Wert(23)
   If (Wert(24) <> "u" Or rUs(UBound(rUs)).PulsAtp_li = "") And InStr(Wert(24), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_li = Wert(24)
   If (Wert(25) <> "u" Or rUs(UBound(rUs)).PulsAdp_re = "") And InStr(Wert(25), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_re = Wert(25)
   If (Wert(26) <> "u" Or rUs(UBound(rUs)).PulsAdp_li = "") And InStr(Wert(26), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_li = Wert(26)
   If (Wert(27) <> "u" Or rUs(UBound(rUs)).Mitarbeiter = "") And InStr(Wert(27), "~") = 0 Then rUs(UBound(rUs)).Mitarbeiter = Wert(27)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdm2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usdm2

Function usd()
 Dim trz, j&
 On Error GoTo fehler
 
 trz = Array("Spritzstellen (Bauch/OS):", _
             ";  Kraft Zehenheber: re:", ",  li:", _
             ";  Kraft Zehenbeuger: re:", ";  li:", _
             ";  Kraft Knie: re:", ";  li:", _
             ";   ASR (Ferse): re:", "; li:", _
             ";  PSR (Knie): re:", "; li:", _
             ";  Monofilament Fußsohle(/5): re:", "; li:", _
             ";  Kalt-Warm(/5): re:", "; li:", _
             ";  Vibr. Innenknö. (/8); re:", "; li:", _
             ";  Vibr. Großzehenballen (/8): re:", "; li:", _
             "  Puls Leiste: re:", "; li:", _
             "; Puls Kniekehle: re:", "; li:", _
             "; Puls Atp.(Innenkn): re:", "; li:", _
             "; Puls Adp.(Fußrü): re:", "; li:", _
             "; Mitarbeiter:", "")
 For j = 1 To UBound(rEi)
  If (rEi(j).art = "usd" Or rEi(j).art = "usdm") And rEi(j).Zeitpunkt >= #3/18/2025# Then
   If rUs(UBound(rUs)).absPos <> 0 Or UBound(rUs) = 0 Then ReDim Preserve rUs(UBound(rUs) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rUsRest(j)
   If (Wert(0) <> "u" Or rUs(UBound(rUs)).Spritzst = "") And InStr(Wert(0), "~") = 0 Then rUs(UBound(rUs)).Spritzst = Wert(0)
   If (Wert(1) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_re = "") And InStr(Wert(1), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_re = Wert(1)
   If (Wert(2) <> "u" Or rUs(UBound(rUs)).Kraft_Zh_li = "") And InStr(Wert(2), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zh_li = Wert(2)
   If (Wert(3) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_re = "") And InStr(Wert(3), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_re = Wert(3)
   If (Wert(4) <> "u" Or rUs(UBound(rUs)).Kraft_Zb_li = "") And InStr(Wert(4), "~") = 0 Then rUs(UBound(rUs)).Kraft_Zb_li = Wert(4)
   If (Wert(5) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_re = "") And InStr(Wert(5), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_re = Wert(5)
   If (Wert(6) <> "u" Or rUs(UBound(rUs)).Kraft_Knie_li = "") And InStr(Wert(6), "~") = 0 Then rUs(UBound(rUs)).Kraft_Knie_li = Wert(6)
   If (Wert(7) <> "u" Or rUs(UBound(rUs)).ASR_re = "") And InStr(Wert(7), "~") = 0 Then rUs(UBound(rUs)).ASR_re = Wert(7)
   If (Wert(8) <> "u" Or rUs(UBound(rUs)).ASR_li = "") And InStr(Wert(8), "~") = 0 Then rUs(UBound(rUs)).ASR_li = Wert(8)
   If (Wert(9) <> "u" Or rUs(UBound(rUs)).PSR_re = "") And InStr(Wert(9), "~") = 0 Then rUs(UBound(rUs)).PSR_re = Wert(9)
   If (Wert(10) <> "u" Or rUs(UBound(rUs)).PSR_li = "") And InStr(Wert(10), "~") = 0 Then rUs(UBound(rUs)).PSR_li = Wert(10)
   If (Wert(11) <> "u" Or rUs(UBound(rUs)).MF_re = "") And InStr(Wert(11), "~") = 0 Then rUs(UBound(rUs)).MF_re = Wert(11) & IIf(IsNumeric(Wert(11)), "/5", "")
   If (Wert(12) <> "u" Or rUs(UBound(rUs)).MF_li = "") And InStr(Wert(12), "~") = 0 Then rUs(UBound(rUs)).MF_li = Wert(12) & IIf(IsNumeric(Wert(12)), "/5", "")
   If (Wert(13) <> "u" Or rUs(UBound(rUs)).KW_re = "") And InStr(Wert(13), "~") = 0 Then rUs(UBound(rUs)).KW_re = Wert(13) & IIf(IsNumeric(Wert(13)), "/5", "")
   If (Wert(14) <> "u" Or rUs(UBound(rUs)).KW_li = "") And InStr(Wert(14), "~") = 0 Then rUs(UBound(rUs)).KW_li = Wert(14) & IIf(IsNumeric(Wert(14)), "/5", "")
   If (Wert(15) <> "u" Or rUs(UBound(rUs)).Vibr_IK_re = "") And InStr(Wert(15), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_re = Wert(15) & IIf(IsNumeric(Wert(15)), "/8", "")
   If (Wert(16) <> "u" Or rUs(UBound(rUs)).Vibr_IK_li = "") And InStr(Wert(16), "~") = 0 Then rUs(UBound(rUs)).Vibr_IK_li = Wert(16) & IIf(IsNumeric(Wert(16)), "/8", "")
   If (Wert(17) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_re = "") And InStr(Wert(17), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_re = Wert(17) & IIf(IsNumeric(Wert(17)), "/8", "")
   If (Wert(18) <> "u" Or rUs(UBound(rUs)).Vibr_GZ_li = "") And InStr(Wert(18), "~") = 0 Then rUs(UBound(rUs)).Vibr_GZ_li = Wert(18) & IIf(IsNumeric(Wert(18)), "/8", "")
   If (Wert(19) <> "u" Or rUs(UBound(rUs)).PulsL_re = "") And InStr(Wert(19), "~") = 0 Then rUs(UBound(rUs)).PulsL_re = Wert(19)
   If (Wert(20) <> "u" Or rUs(UBound(rUs)).PulsL_li = "") And InStr(Wert(20), "~") = 0 Then rUs(UBound(rUs)).PulsL_li = Wert(20)
   If (Wert(21) <> "u" Or rUs(UBound(rUs)).PulsKK_re = "") And InStr(Wert(21), "~") = 0 Then rUs(UBound(rUs)).PulsKK_re = Wert(21)
   If (Wert(22) <> "u" Or rUs(UBound(rUs)).PulsKK_li = "") And InStr(Wert(22), "~") = 0 Then rUs(UBound(rUs)).PulsKK_li = Wert(22)
   If (Wert(23) <> "u" Or rUs(UBound(rUs)).PulsAtp_re = "") And InStr(Wert(23), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_re = Wert(23)
   If (Wert(24) <> "u" Or rUs(UBound(rUs)).PulsAtp_li = "") And InStr(Wert(24), "~") = 0 Then rUs(UBound(rUs)).PulsAtp_li = Wert(24)
   If (Wert(25) <> "u" Or rUs(UBound(rUs)).PulsAdp_re = "") And InStr(Wert(25), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_re = Wert(25)
   If (Wert(26) <> "u" Or rUs(UBound(rUs)).PulsAdp_li = "") And InStr(Wert(26), "~") = 0 Then rUs(UBound(rUs)).PulsAdp_li = Wert(26)
   If (Wert(27) <> "u" Or rUs(UBound(rUs)).Mitarbeiter = "") And InStr(Wert(27), "~") = 0 Then rUs(UBound(rUs)).Mitarbeiter = Wert(27)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdm2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usdm2


Function USfuss()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Fußdeformität:", _
             "Hyperkeratose mit Einblutung:", "Weiteres:", _
             "Z.n. Ulcus:", "Z.n. Amputation:", _
             "Füße genau angeschaut:", "Ulcus:", _
             "Wundinfektion:", "nä US:", _
             "Mitarbeiter:", "")
 For j = 1 To UBound(rEi)
  If rEi(j).art = "fuß" And rEi(j).Zeitpunkt > #7/17/2017 8:00:00 AM# Then
   If rFu(UBound(rFu)).absPos <> 0 Or UBound(rFu) = 0 Then ReDim Preserve rFu(UBound(rFu) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rFuRest(j)
   If (Wert(0) <> "u" Or rFu(UBound(rFu)).Fußdeform = "") And InStr(Wert(0), "~") = 0 Then rFu(UBound(rFu)).Fußdeform = Wert(0)
   If (Wert(1) <> "u" Or rFu(UBound(rFu)).Hyper_mEin = "") And InStr(Wert(1), "~") = 0 Then rFu(UBound(rFu)).Hyper_mEin = Wert(1)
   If (Wert(2) <> "u" Or rFu(UBound(rFu)).Weiteres = "") And InStr(Wert(2), "~") = 0 Then rFu(UBound(rFu)).Weiteres = Wert(2)
   If (Wert(3) <> "u" Or rFu(UBound(rFu)).Zn_Ulcus = "") And InStr(Wert(3), "~") = 0 Then rFu(UBound(rFu)).Zn_Ulcus = Wert(3)
   If (Wert(4) <> "u" Or rFu(UBound(rFu)).Zn_Amput = "") And InStr(Wert(4), "~") = 0 Then rFu(UBound(rFu)).Zn_Amput = Wert(4)
   If (Wert(5) <> "u" Or rFu(UBound(rFu)).Fuß_ang = "") And InStr(Wert(5), "~") = 0 Then rFu(UBound(rFu)).Fuß_ang = Wert(5)
   If (Wert(6) <> "u" Or rFu(UBound(rFu)).Ulcera = "") And InStr(Wert(6), "~") = 0 Then rFu(UBound(rFu)).Ulcera = Wert(6)
   If (Wert(7) <> "u" Or rFu(UBound(rFu)).Wundinfektion = "") And InStr(Wert(7), "~") = 0 Then rFu(UBound(rFu)).Wundinfektion = Wert(7)
   If (Wert(8) <> "u" Or rFu(UBound(rFu)).nae_US = "") And InStr(Wert(8), "~") = 0 Then rFu(UBound(rFu)).nae_US = Wert(8)
   If (Wert(9) <> "u" Or rFu(UBound(rFu)).Mitarbeiter = "") And InStr(Wert(9), "~") = 0 Then rFu(UBound(rFu)).Mitarbeiter = Wert(9)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
   Dim k&
   For k = UBound(trz) - 1 To 1 Step -1
    If Wert(k) = Wert(k - 1) Then Wert(k) = ""
   Next k
  End If
 Next j
 
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usfuss/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usfuss

Function usVKGD()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Wohlempfinden", _
             ", Sättigung", ", Zielwerterreichung", ", Ketonkörper", _
             ", aktueller Gynäkologenbefund", ", Gewichtsentwicklung der Schwangeren", _
             ", HbA1c (monatlich)+ ggf. TSH (dreimonatlich) abgenommen", ", Blutdruck", ", Puls", _
             "; (", "")
'             Debug.Print UBound(trz)
 For j = 1 To UBound(rEi)
  If rEi(j).art = "vkgd" Or rEi(j).art = "vkg" Then
   If rVk(UBound(rVk)).absPos <> 0 Or UBound(rVk) = 0 Then ReDim Preserve rVk(UBound(rVk) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rVkRest(j)
   If (Wert(0) <> "u" Or rVk(UBound(rVk)).Wohlempfinden = "") And InStr(Wert(0), "~") = 0 Then rVk(UBound(rVk)).Wohlempfinden = Wert(0)
   If (Wert(1) <> "u" Or rVk(UBound(rVk)).Saettigung = "") And InStr(Wert(1), "~") = 0 Then rVk(UBound(rVk)).Saettigung = Wert(1)
   If (Wert(2) <> "u" Or rVk(UBound(rVk)).Zielwerterreichung = "") And InStr(Wert(2), "~") = 0 Then rVk(UBound(rVk)).Zielwerterreichung = Wert(2)
   If (Wert(3) <> "u" Or rVk(UBound(rVk)).Ketonkörper = "") And InStr(Wert(3), "~") = 0 Then rVk(UBound(rVk)).Ketonkörper = Wert(3)
   If (Wert(4) <> "u" Or rVk(UBound(rVk)).Gynaekologenbefund = "") And InStr(Wert(4), "~") = 0 Then rVk(UBound(rVk)).Gynaekologenbefund = Wert(4)
   If (Wert(5) <> "u" Or rVk(UBound(rVk)).Gewichtsentwicklung = "") And InStr(Wert(5), "~") = 0 Then rVk(UBound(rVk)).Gewichtsentwicklung = Wert(5)
   If (Wert(6) <> "u" Or rVk(UBound(rVk)).HbA1c = "") And InStr(Wert(6), "~") = 0 Then rVk(UBound(rVk)).HbA1c = Wert(6)
   If (Wert(7) <> "u" Or rVk(UBound(rVk)).Blutdruck = "") And InStr(Wert(7), "~") = 0 Then rVk(UBound(rVk)).Blutdruck = Wert(7)
   If (Wert(8) <> "u" Or rVk(UBound(rVk)).Puls = "") And InStr(Wert(8), "~") = 0 Then rVk(UBound(rVk)).Puls = Wert(8)
   If (Wert(9) <> "u" Or rVk(UBound(rVk)).Mitarbeiter = "") And InStr(Wert(9), "~") = 0 Then rVk(UBound(rVk)).Mitarbeiter = Wert(9)
   If Right$(rVk(UBound(rVk)).Mitarbeiter, 1) = ")" Then rVk(UBound(rVk)).Mitarbeiter = left$(rVk(UBound(rVk)).Mitarbeiter, Len(rVk(UBound(rVk)).Mitarbeiter) - 1)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
   Dim k&
   For k = UBound(trz) - 1 To 1 Step -1
    If Wert(k) = Wert(k - 1) Then Wert(k) = ""
   Next k
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usulcus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usVKGD()

Function usVKGD2()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Wohlempfinden", _
             ", Sättigung", ", Zielwerterreichung", _
             ", aktueller Gynäkologenbefund", ", Gewichtsentwicklung der Schwangeren", _
             ", Bewegung: Art:", ", Minuten pro Woche:", ", Blutdruck", ", Puls", _
             "; (", "")
'             Debug.Print UBound(trz)
 For j = 1 To UBound(rEi)
  If rEi(j).art = "vkgd2" Then
   If rVk(UBound(rVk)).absPos <> 0 Or UBound(rVk) = 0 Then ReDim Preserve rVk(UBound(rVk) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rVkRest(j)
   If (Wert(0) <> "u" Or rVk(UBound(rVk)).Wohlempfinden = "") And InStr(Wert(0), "~") = 0 Then rVk(UBound(rVk)).Wohlempfinden = Wert(0)
   If (Wert(1) <> "u" Or rVk(UBound(rVk)).Saettigung = "") And InStr(Wert(1), "~") = 0 Then rVk(UBound(rVk)).Saettigung = Wert(1)
   If (Wert(2) <> "u" Or rVk(UBound(rVk)).Zielwerterreichung = "") And InStr(Wert(2), "~") = 0 Then rVk(UBound(rVk)).Zielwerterreichung = Wert(2)
   If (Wert(3) <> "u" Or rVk(UBound(rVk)).Gynaekologenbefund = "") And InStr(Wert(3), "~") = 0 Then rVk(UBound(rVk)).Gynaekologenbefund = Wert(3)
   If (Wert(4) <> "u" Or rVk(UBound(rVk)).Gewichtsentwicklung = "") And InStr(Wert(4), "~") = 0 Then rVk(UBound(rVk)).Gewichtsentwicklung = Wert(4)
   If (Wert(5) <> "u" Or rVk(UBound(rVk)).Bewegung = "") And InStr(Wert(5), "~") = 0 Then rVk(UBound(rVk)).Bewegung = Wert(5)
   If (Wert(6) <> "u" Or rVk(UBound(rVk)).Minuten = "") And InStr(Wert(6), "~") = 0 Then rVk(UBound(rVk)).Minuten = Wert(6)
   If (Wert(7) <> "u" Or rVk(UBound(rVk)).Blutdruck = "") And InStr(Wert(7), "~") = 0 Then rVk(UBound(rVk)).Blutdruck = Wert(7)
   If (Wert(8) <> "u" Or rVk(UBound(rVk)).Puls = "") And InStr(Wert(8), "~") = 0 Then rVk(UBound(rVk)).Puls = Wert(8)
   If (Wert(9) <> "u" Or rVk(UBound(rVk)).Mitarbeiter = "") And InStr(Wert(9), "~") = 0 Then rVk(UBound(rVk)).Mitarbeiter = Wert(9)
   If Right$(rVk(UBound(rVk)).Mitarbeiter, 1) = ")" Then rVk(UBound(rVk)).Mitarbeiter = left$(rVk(UBound(rVk)).Mitarbeiter, Len(rVk(UBound(rVk)).Mitarbeiter) - 1)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
   Dim k&
   For k = UBound(trz) - 1 To 1 Step -1
    If Wert(k) = Wert(k - 1) Then Wert(k) = ""
   Next k
  End If
 Next j
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usulcus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usVKGD2()

Function USUlcus()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Lokalisation:", _
             "Seite:", "Größe:", _
             "Beläge:", "Exsudat:", _
             "Geruch aus 3 cm Entfernung:", "Wundrand:", "Wundumgebung:", "Temperatur:", "Fotodoku", _
             "Wundversorgung:", _
             "Mitarbeiter:", "")
'             Debug.Print UBound(trz)
 For j = 1 To UBound(rEi)
  If rEi(j).art = "ulcus" Then
   If rUl(UBound(rUl)).absPos <> 0 Or UBound(rUl) = 0 Then ReDim Preserve rUl(UBound(rUl) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rUlRest(j)
   If (Wert(0) <> "u" Or rUl(UBound(rUl)).Lokalisation = "") And InStr(Wert(0), "~") = 0 Then rUl(UBound(rUl)).Lokalisation = Wert(0)
   If (Wert(1) <> "u" Or rUl(UBound(rUl)).Seite = "") And InStr(Wert(1), "~") = 0 Then rUl(UBound(rUl)).Seite = Wert(1)
   If (Wert(2) <> "u" Or rUl(UBound(rUl)).Größe = "") And InStr(Wert(2), "~") = 0 Then rUl(UBound(rUl)).Größe = Wert(2)
   If (Wert(3) <> "u" Or rUl(UBound(rUl)).Beläge = "") And InStr(Wert(3), "~") = 0 Then rUl(UBound(rUl)).Beläge = Wert(3)
   If (Wert(4) <> "u" Or rUl(UBound(rUl)).Exsudat = "") And InStr(Wert(4), "~") = 0 Then rUl(UBound(rUl)).Exsudat = Wert(4)
   If (Wert(5) <> "u" Or rUl(UBound(rUl)).Geruch = "") And InStr(Wert(5), "~") = 0 Then rUl(UBound(rUl)).Geruch = Wert(5)
   If (Wert(6) <> "u" Or rUl(UBound(rUl)).Wundrand = "") And InStr(Wert(6), "~") = 0 Then rUl(UBound(rUl)).Wundrand = Wert(6)
   If (Wert(7) <> "u" Or rUl(UBound(rUl)).Wundumgebung = "") And InStr(Wert(7), "~") = 0 Then rUl(UBound(rUl)).Wundumgebung = Wert(7)
   If (Wert(8) <> "u" Or rUl(UBound(rUl)).Temperatur = "") And InStr(Wert(8), "~") = 0 Then rUl(UBound(rUl)).Temperatur = Wert(8)
   If (Wert(9) <> "u" Or rUl(UBound(rUl)).Fotodoku = "") And InStr(Wert(9), "~") = 0 Then rUl(UBound(rUl)).Fotodoku = Wert(9)
   If (Wert(10) <> "u" Or rUl(UBound(rUl)).Wundversorgung = "") And InStr(Wert(10), "~") = 0 Then rUl(UBound(rUl)).Wundversorgung = Wert(10)
   If (Wert(11) <> "u" Or rUl(UBound(rUl)).Mitarbeiter = "") And InStr(Wert(11), "~") = 0 Then rUl(UBound(rUl)).Mitarbeiter = Wert(11)
   If rEi(j).Zeitpunkt > #3/18/2025# Then rUs(UBound(rUs)).Mitarbeiter = rEi(j).Ersteller
   Dim k&
   For k = UBound(trz) - 1 To 1 Step -1
    If Wert(k) = Wert(k - 1) Then Wert(k) = ""
   Next k
  End If
 Next j
 
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usulcus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usulcus

' 16.7.23: noch nicht aktiviert
Function USUlcusa()
 Dim trz, j&
 On Error GoTo fehler
 trz = Array("Lokalisation:", _
             "Seite:", "Größe:", _
             "Beläge:", "Exsudat:", _
             "Geruch aus 3 cm Entfernung:", "Wundrand:", "Wundgrund:", "Wundumgebung:", "Temperatur:", "Phase:", "Fotodoku", _
             "Wundversorgung:", _
             "Mitarbeiter:", "")
'             Debug.Print UBound(trz)
 For j = 1 To UBound(rEi)
  If rEi(j).art = "ulcus" Then
   If rUl(UBound(rUl)).absPos <> 0 Or UBound(rUl) = 0 Then ReDim Preserve rUl(UBound(rUl) + 1)
   Dim Wert$()
   Call Kusd(trz, Wert, j)
   Call rUlRest(j)
   If (Wert(0) <> "u" Or rUl(UBound(rUl)).Lokalisation = "") And InStr(Wert(0), "~") = 0 Then rUl(UBound(rUl)).Lokalisation = Wert(0)
   If (Wert(1) <> "u" Or rUl(UBound(rUl)).Seite = "") And InStr(Wert(1), "~") = 0 Then rUl(UBound(rUl)).Seite = Wert(1)
   If (Wert(2) <> "u" Or rUl(UBound(rUl)).Größe = "") And InStr(Wert(2), "~") = 0 Then rUl(UBound(rUl)).Größe = Wert(2)
   If (Wert(3) <> "u" Or rUl(UBound(rUl)).Beläge = "") And InStr(Wert(3), "~") = 0 Then rUl(UBound(rUl)).Beläge = Wert(3)
   If (Wert(4) <> "u" Or rUl(UBound(rUl)).Exsudat = "") And InStr(Wert(4), "~") = 0 Then rUl(UBound(rUl)).Exsudat = Wert(4)
   If (Wert(5) <> "u" Or rUl(UBound(rUl)).Geruch = "") And InStr(Wert(5), "~") = 0 Then rUl(UBound(rUl)).Geruch = Wert(5)
   If (Wert(6) <> "u" Or rUl(UBound(rUl)).Wundrand = "") And InStr(Wert(6), "~") = 0 Then rUl(UBound(rUl)).Wundrand = Wert(6)
'   If (Wert(7) <> "u" Or rUl(UBound(rUl)).Wundgrund = "") And InStr(Wert(7), "~") = 0 Then rUl(UBound(rUl)).Wundgrund = Wert(7)
   If (Wert(8) <> "u" Or rUl(UBound(rUl)).Wundumgebung = "") And InStr(Wert(8), "~") = 0 Then rUl(UBound(rUl)).Wundumgebung = Wert(8)
   If (Wert(9) <> "u" Or rUl(UBound(rUl)).Temperatur = "") And InStr(Wert(9), "~") = 0 Then rUl(UBound(rUl)).Temperatur = Wert(9)
'   If (Wert(10) <> "u" Or rUl(UBound(rUl)).Phase = "") And InStr(Wert(10), "~") = 0 Then rUl(UBound(rUl)).Phase = Wert(10)
   If (Wert(11) <> "u" Or rUl(UBound(rUl)).Fotodoku = "") And InStr(Wert(11), "~") = 0 Then rUl(UBound(rUl)).Fotodoku = Wert(11)
   If (Wert(12) <> "u" Or rUl(UBound(rUl)).Wundversorgung = "") And InStr(Wert(12), "~") = 0 Then rUl(UBound(rUl)).Wundversorgung = Wert(12)
   If (Wert(13) <> "u" Or rUl(UBound(rUl)).Mitarbeiter = "") And InStr(Wert(13), "~") = 0 Then rUl(UBound(rUl)).Mitarbeiter = Wert(13)
   Dim k&
   For k = UBound(trz) - 1 To 1 Step -1
    If Wert(k) = Wert(k - 1) Then Wert(k) = ""
   Next k
  End If
 Next j
 
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usulcusa/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usulcusa

Function AnAlle() ' -> im Menü als "Anamnesedatenblatt -> Anamnesen ergänzen"
 Dim i%, j&
 Dim zurZahl$, buch$, vglStr$
 On Error GoTo fehler
 Call anal
 Call AnDmFieseln
 Call AnDm2Fieseln
 Call usdmAlt
 Call usAl
'Jetzt noch Größe und Gewicht ggf. nachtragen
 For j = 1 To UBound(rEi)
  Select Case UCase$(rEi(j).art)
   Case "GROESSE", "GRÖSSE", "GRÖßE"
    If rsAnm.Fields(16) = 0 Or IsNull(rsAnm.Fields(16)) Then ' Größe 6.3.13
     zurZahl = vNS
     vglStr = "0123456789,."
     For i = 1 To Len(rEi(j).Inhalt)
      buch = Mid$(rEi(j).Inhalt, i, 1)
      If InStrB(vglStr, buch) <> 0 Then
       zurZahl = zurZahl + buch
       If InStrB(".,", buch) <> 0 Then
        vglStr = "0123456789"
       End If
      Else
       Exit For
      End If
     Next i
     If zurZahl <> vNS Then
      rsAnm.Fields(16) = CDbl(zurZahl)
'      IF rsAnm.fields(16) > 3 THEN rsAnm.fields(16) = rsAnm.fields(16) / 100
     End If
    End If ' rsAnm.fields(16) = 0
   Case "GEWI", "GEWICHT"
    If rsAnm!Gewicht = 0 Or IsNull(rsAnm!Gewicht) Then
      zurZahl = vNS
      vglStr = "0123456789,."
      For i = 1 To Len(rEi(j).Inhalt)
       buch = Mid$(rEi(j).Inhalt, i, 1)
       If InStrB(vglStr, buch) <> 0 Then
        zurZahl = zurZahl + buch
        If InStrB(".,", buch) <> 0 Then
         vglStr = "0123456789"
        End If
       Else
        Exit For
       End If
      Next i
      If zurZahl <> vNS Then
       rsAnm!Gewicht = CDbl(zurZahl)
       If rsAnm!Gewicht < 3 Then rsAnm!Gewicht = rsAnm!Gewicht * 100
      End If
    End If '  rsAnam!Gewicht = 0 THEN
  End Select
 Next j
 If rsAnm.Fields(16) < 3 Then rsAnm.Fields(16) = rsAnm.Fields(16) * 100 ' Größe 6.5.13
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnAlle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AnAlle()

' aufgerufen in anAlle
Function AnDmFieseln()
 Const imin% = -3
 Const imax% = 113
 Dim tr$(imin To imax)
 Dim Fd$(imin To imax)
 On Error GoTo fehler
 tr(-3) = "seit:"
 Fd(-3) = "Sexualstörung seit"
 tr(-2) = "Sexualstörung?"
 Fd(-2) = "Sexualstörung"
 tr(-1) = "Gelenke (Bewegungseinschränkungen ohne andere Erklärung)?"
 Fd(-1) = "Bewegungseinschränkungen"
 tr(0) = "Haut:"
 Fd(0) = "Folgeerkrankungen Haut"
 tr(1) = "Sonstige Folgeerkrankungen:"
 Fd(1) = vNS
 tr(2) = "Schwindel beim Aufstehen?"
 Fd(2) = "Schwindel Aufstehen"
 tr(3) = "Entleerungsstörung Harnblase?"
 Fd(3) = "Entleerungsstörungen Harnblase"
 tr(4) = "Entleerungsstörung Magen?"
 Fd(4) = "Entleerungsstörungen Magen"
 tr(5) = "Neue Fußkomplikationen in den letzten 12 Monaten?"
 Fd(5) = "Neue Fußkomplikationen"
 tr(6) = "Haben Sie diabetesgerechte Einlagen/Schuhe?"
 Fd(6) = "Einlagen"
 tr(7) = "Podologie?"
 Fd(7) = "Podologie"
 tr(8) = "Gehen Sie regelmäßig zur Fußpflege?"
 Fd(8) = "Fußpflege"
 tr(9) = "ggf. Beschreibung:"
 Fd(9) = "Verformungen Beschreibung"
 tr(10) = "Verformungen der Füße?"
 Fd(10) = "Verformungen"
 tr(11) = "Druckstellen/Schwielen?"
 Fd(11) = "Druckstellen"
 tr(12) = "nach oben reichend bis"
 Fd(12) = "Ameisen Ausmaß"
 tr(13) = "Gefühlsstörung / Ameisenlaufen an den Füßen?"
 Fd(13) = "Ameisenlaufen"
' fd(14,5) = "pAVK Beschreibung"
 tr(14) = "Amputation (wann, wo)?"
 Fd(14) = "Amputation"
 tr(15) = "Ulcus?"
 Fd(15) = "Geschwür"
 tr(16) = "Bypaß am Bein?"
 Fd(16) = "bypaß peripher"
 tr(17) = "Schaufensterkrankheit?"
 Fd(17) = "Schaufensterkrankheit"
 tr(18) = "Beindurchblutungsstörung?"
 Fd(18) = "Beindurchblutungsstörung"
 tr(19) = "Schlaganfall? (wann? Lähmung? wo?)"
 Fd(19) = "Schlaganfall"
 tr(19) = "(wann? Lähmung? wo?)"
 Fd(19) = "Schlaganfall"
 tr(21) = "Schlaganfall?"
 Fd(21) = "Schlaganfall"
 tr(22) = "Hirndurchblutungsstörung?"
 Fd(22) = "Hirndurchblutungsstörung"
' fd(21,5) = "Herzkrankheit Beschreibung"
 tr(23) = "Herzschwäche mit Luftnot bei Belastung oder Beinschwellungen?"
 Fd(23) = "Herzschwäche"
 tr(24) = "wann?"
 Fd(24) = "Bypass wann"
 tr(25) = "Bypaß am Herzen?"
 Fd(25) = "Bypass kardial"
 tr(26) = "PTCA oder Stent?"
 Fd(26) = "PTCA oder Stent"
 tr(27) = "wann?"
 Fd(27) = "Herzinfarkt wann"
 tr(28) = "Herzinfarkt?"
 Fd(28) = "Herzinfarkt"
 tr(29) = "Brustschmerz b. Belastung (A.p.):"
 Fd(29) = "Angina pectoris"
 tr(30) = "Herzkrankheit?"
 Fd(30) = "Herzkrankheit Beschreibung"
 tr(31) = "seit:"
 Fd(31) = "Dialyse seit"
' fd(30,6)="andere Nierenerkrankung"
 tr(32) = "Dialyse?"
 Fd(32) = "Dialyse"
 tr(33) = "erhöht?"
 Fd(33) = "erhöht?"
 tr(34) = "Albumin im Urin zuletzt untersucht:"
 Fd(34) = "Albumin zuletzt"
 tr(35) = "Ist ein diabetischer Nierenschaden bekannt?"
 Fd(35) = "Diabet Nierenschaden"
 tr(36) = "liegt eine nicht durch eine Brille behebbare Sehminderung vor?"
 Fd(36) = "Sehminderung unbehebbar"
 tr(37) = "wurde die Netzhaut schon gelasert?"
 Fd(37) = "Netzhaut gelasert"
 tr(38) = "Wie war der Befund?"
 Fd(38) = "Augensp Befund"
 tr(39) = "Wann war die letzte Augenspiegelung?"
 Fd(39) = "Augensp zuletzt"
 tr(40) = "wie hoch sind Ihre Blutdruckwerte meist?"
 Fd(40) = "Blutdruckwerte"
' fd(39,5)="BHD beh mit"
 tr(41) = "Selbstmessung?"
 Fd(41) = "BDselbst"
 tr(42) = "seit:"
 Fd(42) = "BHD seit"
 tr(43) = "Bluthochdruck?"
 Fd(43) = "Bluthochdruck"
 tr(44) = "aktuelle SSW:"
 Fd(44) = "Schwanger seit"
 tr(45) = "Schwangerschaft?"
 Fd(45) = "Schwanger"
 tr(46) = "wie oft im Monat ist der BZ > 300 mg/dl?"
 Fd(46) = "BZgr300 pM"
 tr(47) = "Wie oft in den letzten 12 Monaten waren Ketoazidosen mit Krankenhauseinweisung?"
 Fd(47) = "Keto pa"
 tr(48) = "Wie oft in den letzten 12 Monaten waren Ketoazidosen mit Krakenhauseinweisung?"
 Fd(48) = "Keto pa"
 tr(49) = "wie oft in den letzten 12 Monaten waren Sie deshalb bewußtlos?"
 Fd(49) = "Bewußtlos pa"
 tr(50) = "wie oft in den letzten 12 Monaten war deshalb fremde Hilfe nötig?"
 Fd(50) = "Fremde Hilfe pa"
 tr(51) = "werden sie rechtzeitg bemerkt?"
 Fd(51) = "UZ rechtzeitig"
 tr(52) = "wie of im Monat sind Unterzucker kleiner 50 mg/dl?"
 Fd(52) = "Unterzucker pM"
 tr(53) = "wann häufen sich Unterzucker?"
 Fd(53) = "UZ Tageszeit"
 tr(54) = "nach dem Essen:"
 Fd(54) = "BZWerte n d Essen"
 tr(55) = "Wie hoch ist der BZ meist vor dem Essen:"
 Fd(55) = "BZWerte v d Essen"
 tr(56) = "nachts:"
 Fd(56) = "BZMessungen p W nachts"
 tr(57) = "davon nach dem Essen:"
 Fd(57) = "BZMessungen pW ndE"
 tr(58) = "Durchschnittliche Zahl der BZ-Messungen pro Woche:"
 Fd(58) = "BZMessungen pW"
 tr(59) = "Aufschreiben:"
 Fd(59) = "Aufschreiben"
 tr(60) = "wie heißt das Gerät?"
 Fd(60) = "Gerät"
 tr(61) = "BZ-Messung selbst?"
 Fd(61) = "BZMessungen selbst"
 tr(62) = "vorherige Werte:"
 Fd(62) = "vorherige Werte"
 tr(63) = "gemessen am:"
 Fd(63) = "gemessen am"
 tr(64) = "Letztes HbA1c [%]:"
 Fd(64) = "letztes HbA1c"
 tr(65) = "letzte Diabetesschulung [Ort]:"
 Fd(65) = "Ort Schulung"
 tr(66) = "Letzte Diabetesschulung [Jahr]:"
 Fd(66) = "Jahr letzte Diabetesschulung"
 tr(67) = "spät"
 Fd(67) = "Spritzstelle nachts"
 tr(68) = "abends"
 Fd(68) = "Spritzstelle abends"
 tr(69) = "mittags"
 Fd(69) = "Spritzstelle mittags"
 tr(70) = "Spritzort früh"
 Fd(70) = "Spritzstelle früh"
 tr(71) = "abends"
 Fd(71) = "Spritz-Eß-Abstand abends"
 tr(72) = "mittags"
 Fd(72) = "Spritz-Eß-Abstand mittags"
 tr(73) = "Spritz-Eß-Abstand bei normalem BZ in Minuten ca.: früh"
 Fd(73) = "Spritz-Eß-Abstand früh"
 tr(74) = "abends"
 Fd(74) = "Essenszeit abends"
 tr(75) = "mittags"
 Fd(75) = "Essenszeit mittags"
 tr(76) = "Essenszeiten ca. früh"
 Fd(76) = "Essenszeit früh"
 tr(77) = "abends"
 Fd(77) = "Broteinheiten abends"
 tr(78) = "mittags"
 Fd(78) = "Broteinheiten mittags"
 tr(79) = "Broteinheiten ca. früh"
 Fd(79) = "Broteinheiten früh"
 tr(80) = "Marke:"
 Fd(80) = "Insulinpumpe Marke"
 tr(81) = "seit [Jahr]:"
 Fd(81) = "Insulinpumpe seit"
 tr(82) = "Insulinpumpe"
 Fd(82) = "Insulinpumpe"
 tr(83) = "spät"
 Fd(83) = "DiabetesMedikament 4 Menge"
 tr(84) = "abends"
 Fd(84) = "DiabetesMedikament 4 Menge"
 tr(85) = "mittags"
 Fd(85) = "DiabetesMedikament 4 Menge"
 tr(86) = "früh"
 Fd(86) = "DiabetesMedikament 4 Menge"
 tr(87) = "Diabetes-Medikament 4:"
 Fd(87) = "DiabetesMedikament 4"
 tr(88) = "spät"
 Fd(88) = "DiabetesMedikament 3 Menge"
 tr(89) = "abends"
 Fd(89) = "DiabetesMedikament 3 Menge"
 tr(90) = "mittags"
 Fd(90) = "DiabetesMedikament 3 Menge"
 tr(91) = "früh"
 Fd(91) = "DiabetesMedikament 3 Menge"
 tr(92) = "Diabetes-Medikament 3:"
 Fd(92) = "DiabetesMedikament 3"
 tr(93) = "spät"
 Fd(93) = "DiabetesMedikament 2 Menge"
 tr(94) = "abends"
 Fd(94) = "DiabetesMedikament 2 Menge"
 tr(95) = "mittags"
 Fd(95) = "DiabetesMedikament 2 Menge"
 tr(96) = "früh"
 Fd(96) = "DiabetesMedikament 2 Menge"
 tr(97) = "Diabetes-Medikament 2:"
 Fd(97) = "DiabetesMedikament 2"
 tr(98) = "spät"
 Fd(98) = "DiabetesMedikament 1 Menge"
 tr(99) = "abends"
 Fd(99) = "DiabetesMedikament 1 Menge"
 tr(100) = "mittags"
 Fd(100) = "DiabetesMedikament 1 Menge"
 tr(101) = "früh"
 Fd(101) = "DiabetesMedikament 1 Menge"
 tr(102) = "Diabetes-Medikament 1:"
 Fd(102) = "DiabetesMedikament 1"
 tr(103) = "Tendenz"
 Fd(103) = "Tendenz"
 tr(104) = "Gewicht [kg]"
 Fd(104) = "Gewicht"
 tr(105) = "Körpergröße [cm]"
 Fd(105) = "Größe"
 tr(106) = "Familienanamnese:"
 Fd(106) = "Familienanamnese"
 tr(107) = "Grund für Vorstellung"
 Fd(107) = "Grund für Vorstellung"
 tr(108) = "Insulin seit [Jahr]:"
 Fd(108) = "Insulin seit"
 tr(109) = "Tabletten seit [Jahr]:"
 Fd(109) = "Tabletten seit"
 tr(110) = "seit [Jahr]:"
 Fd(110) = "Diabetes seit"
 tr(111) = "Diabetes Typ"
 Fd(111) = "Diabetestyp"
 tr(112) = "Teilnahme am DMP:"
 Fd(112) = "DMP"
 tr(113) = "Hausarzt (insbes. bei Gemeinschaftspraxen und abweichendem ÜW-Schein):"
 Fd(113) = "Hausarzt"
 Call do_anImp(imin, imax, tr(), Fd(), "|andm|")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnDm/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AnDM

' aufgerufen in anAlle
Function AnDm2Fieseln()
 Const imin% = 0
 Const imax% = 74
 Dim tr$(imin To imax)
 Dim Fd$(imin To imax)
 On Error GoTo fehler
 
 tr(0) = "Parodontitis, ggf. mit Beschreibung?"
 Fd(0) = "Parodontitis"
 tr(1) = "seit:"
 Fd(1) = "Sexualstörung seit"
 tr(2) = "Sexualstörung?"
 Fd(2) = "Sexualstörung"
 tr(3) = "Schwindel beim Aufstehen?"
 Fd(3) = "Schwindel Aufstehen"
 tr(4) = "Entleerungsstörung Harnblase?"
 Fd(4) = "Entleerungsstörungen Harnblase"
 tr(5) = "Entleerungsstörung Magen?"
 Fd(5) = "Entleerungsstörungen Magen"
 tr(6) = "Neue Fußkomplikationen in den letzten 12 Monaten?"
 Fd(6) = "Neue Fußkomplikationen"
 tr(7) = "Haben Sie diabetesgerechte Einlagen/Schuhe?"
 Fd(7) = "Einlagen"
 tr(8) = "Podologie?"
 Fd(8) = "Podologie"
 tr(9) = "Gehen Sie regelmäßig zur Fußpflege?"
 Fd(9) = "Fußpflege"
 tr(10) = "ggf. Beschreibung:"
 Fd(10) = "Verformungen Beschreibung"
 tr(11) = "Verformungen der Füße?"
 Fd(11) = "Verformungen"
 tr(12) = "Druckstellen/Schwielen?"
 Fd(12) = "Druckstellen"
 tr(13) = "nach oben reichend bis"
 Fd(13) = "Ameisen Ausmaß"
 tr(14) = "Gefühlsstörung / Ameisenlaufen an den Füßen?"
 Fd(14) = "Ameisenlaufen"
 tr(15) = "Amputation (wann, wo)?"
 Fd(15) = "Amputation"
 tr(16) = "Ulcus?"
 Fd(16) = "Geschwür"
 tr(17) = "Bypaß am Bein?"
 Fd(17) = "bypaß peripher"
 tr(18) = "Schaufensterkrankheit?"
 Fd(18) = "Schaufensterkrankheit"
 tr(19) = "Beindurchblutungsstörung?"
 Fd(19) = "Beindurchblutungsstörung"
 tr(20) = "Schlaganfall/Hirndurchblutungsstörung? (wann? Lähmung? wo?)"
 Fd(20) = "Schlaganfall"
 tr(21) = "(wann? Lähmung? wo?)"
 Fd(21) = "Schlaganfall"
 tr(22) = "Schlaganfall?"
 Fd(22) = "Schlaganfall"
' fd(21,5) = "Herzkrankheit Beschreibung"
 tr(23) = "Herzschwäche mit Luftnot bei Belastung oder Beinschwellungen?"
 Fd(23) = "Herzschwäche"
 tr(24) = "wann?"
 Fd(24) = "Bypass wann"
 tr(25) = "Bypaß am Herzen?"
 Fd(25) = "Bypass kardial"
 tr(26) = "PTCA oder Stent?"
 Fd(26) = "PTCA oder Stent"
 tr(27) = "wann?"
 Fd(27) = "Herzinfarkt wann"
 tr(28) = "Herzinfarkt?"
 Fd(28) = "Herzinfarkt"
 tr(29) = "Brustschmerz b. Belastung (A.p.):"
 Fd(29) = "Angina pectoris"
 tr(30) = "Herzkrankheit?"
 Fd(30) = "Herzkrankheit"
 tr(31) = "seit:"
 Fd(31) = "Dialyse seit"
' fd(30,6)="andere Nierenerkrankung"
 tr(32) = "Dialyse?"
 Fd(32) = "Dialyse"
 tr(33) = "Kreatinin erhöht?"
 Fd(33) = "Kreatinin"
 tr(34) = "erhöht?"
 Fd(34) = "erhöht?"
 tr(35) = "Albumin im Urin zuletzt untersucht:"
 Fd(35) = "Albumin zuletzt"
 tr(36) = "Ist ein diabetischer Nierenschaden bekannt?"
 Fd(36) = "Diabet Nierenschaden"
 tr(37) = "liegt eine nicht durch eine Brille behebbare Sehminderung vor?"
 Fd(37) = "Sehminderung unbehebbar"
 tr(38) = "wurde die Netzhaut schon gelasert?"
 Fd(38) = "Netzhaut gelasert"
 tr(39) = "Wie war der Befund?"
 Fd(39) = "Augensp Befund"
 tr(40) = "Wann war die letzte Augenspiegelung?"
 Fd(40) = "Augensp zuletzt"
 tr(41) = "wie hoch sind Ihre Blutdruckwerte meist?"
 Fd(41) = "Blutdruckwerte"
' fd(39,5)="BHD beh mit"
 tr(42) = "Selbstmessung?"
 Fd(42) = "BDselbst"
 tr(43) = "seit:"
 Fd(43) = "BHD seit"
 tr(44) = "Bluthochdruck?"
 Fd(44) = "Bluthochdruck"
 tr(45) = "aktuelle SSW:"
 Fd(45) = "Schwanger seit"
 tr(46) = "Schwangerschaft?"
 Fd(46) = "Schwanger"
 tr(47) = "wie oft i.d. letzten 12 Monaten?"
 Fd(47) = "Keto pa"
 tr(48) = "Bisher Ketoazidosen mit Krankenhauseinweisung?"
 Fd(48) = "Keto"
 tr(49) = "wie oft in den letzten 12 Monaten?"
 Fd(49) = "Bewußtlos pa"
 tr(50) = "Bisher Unterzucker mit Bewußtlosigkeit oder Fremdhilfe?"
 Fd(50) = "Schwere Uzu"
 tr(51) = "Aufschreiben:"
 Fd(51) = "Aufschreiben"
 tr(52) = "wie heißt das Gerät?"
 Fd(52) = "Gerät"
 tr(53) = "BZ-Messung selbst?"
 Fd(53) = "BZMessungen selbst"
 tr(54) = "seit:"
 Fd(54) = "CGM seit"
 tr(55) = "Subcutane Zuckermessung?"
 Fd(55) = "subcutane Zuckermessung"
 tr(56) = "vorherige Werte:"
 Fd(56) = "vorherige Werte"
 tr(57) = "gemessen am:"
 Fd(57) = "gemessen am"
 tr(58) = "Letztes HbA1c [%]:"
 Fd(58) = "letztes HbA1c"
 tr(59) = "letzte Diabetesschulung [Ort]:"
 Fd(59) = "Ort Schulung"
 tr(60) = "Letzte Diabetesschulung [Jahr]:"
 Fd(60) = "Jahr letzte Diabetesschulung"
 tr(61) = "Ernährung:"
 Fd(61) = "Ernährung"
 tr(62) = "Verzögerungsinsulin"
 Fd(62) = "Spritzstelle nachts"
 tr(63) = "Spritzort Essensinsulin"
 Fd(63) = "Spritzstelle früh"
 tr(64) = "Marke:"
 Fd(64) = "Insulinpumpe Marke"
 tr(65) = "seit [Jahr]:"
 Fd(65) = "Insulinpumpe seit"
 tr(66) = "Insulinpumpe"
 Fd(66) = "Insulinpumpe"
 tr(67) = "Familienanamnese:"
 Fd(67) = "Familienanamnese"
 tr(68) = "Grund für Vorstellung"
 Fd(68) = "Grund für Vorstellung"
 tr(69) = "Insulin seit [Jahr]:"
 Fd(69) = "Insulin seit"
 tr(70) = "Tabletten seit [Jahr]:"
 Fd(70) = "Tabletten seit"
 tr(71) = "seit [Jahr]:"
 Fd(71) = "Diabetes seit"
 tr(72) = "Diabetes Typ"
 Fd(72) = "Diabetestyp"
 tr(73) = "Teilnahme am DMP:"
 Fd(73) = "DMP"
 tr(74) = "Hausarzt (insbes. bei Gemeinschaftspraxen und abweichendem ÜW-Schein):"
 Fd(74) = "Hausarzt"
 Call do_anImp(imin, imax, tr(), Fd(), "|andm2|")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnDm2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AnDM2

Function anal()
 Const imin% = -4
 Const imax% = 7
 Dim tr$(imin To imax)
 Dim Fd$(imin To imax)
 On Error GoTo fehler
 tr(-4) = "Mitarbeiter:"
 Fd(-4) = "Mitarbeiter"
 tr(-3) = "Wie viel Alkohol trinken Sie im Schnitt pro Woche?"
 Fd(-3) = "Alkohol"
 tr(-2) = "wie viel?"
 Fd(-2) = "TabakMenge"
 tr(-1) = "Rauchen Sie noch?"
 Fd(-1) = "TabakAkt"
 tr(0) = "bis wann:"
 Fd(0) = "TabakBis"
 tr(1) = "Früher geraucht?"
 Fd(1) = "TabakEx"
 tr(2) = "Rauchen:"
 Fd(2) = "Tabak"
 tr(3) = "Tendenz"
 Fd(3) = "Tendenz"
 tr(4) = "Gewicht:"
 Fd(4) = "Gewicht"
 tr(5) = "Größe:"
 Fd(5) = "Größe"
 tr(6) = "Wichtige bisherige Krankheiten und Operationen:"
 Fd(6) = "Weitere Anamnese"
 tr(7) = "Grund für Vorstellung in der Praxis:"
 Fd(7) = "Grund für Vorstellung"
 Call do_anImp(imin, imax, tr(), Fd(), "|anal|ana|")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnAl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'AnAl

'4.9.06: wird aufgerufen in: AnAl, UsAl, UsDm und AlDm
' arten: z.B. "|anal|ana|"
Function do_anImp(imin%, imax%, tr$(), Fd$(), arten$, Optional tfd1, Optional tfd2, Optional kDB%)  ' Testfeld, kein Datenbankeintrag
' Dim rsAnm AS DAO.Recordset, _
     rEi AS DAO.Recordset
 Dim neuinh$, fld$
 Dim i&, j%, k%, im&
 Dim s$, trennz$, tz2$
 Dim sp$(), spe$()
 'Dim endZ$(3), endZe ', vorwertnull%
 On Error GoTo fehler
 'endZ(0) = ","
 'endZ(1) = "."
 'endZ(2) = ";"
 'endZ(3) = "~"
 Const EndStr$ = ",.;~ "
' dtbInit
' SET rsAnm = TabÖff("Anamnesebogen", "Pat_id")
' SET rEi = TabÖff("eintraege", "Auswahl")
' rsAnm.Seek "=", Pat_id
' rEi.Seek "=", Pat_id, makro
 If Not kDB Then
  For i = 1 To UBound(rEi)
   If InStrB(arten, "|" & rEi(i).art & "|") Then
    im = i
    Exit For
   End If
  Next i
 End If
 
 If im <> 0 Or kDB <> 0 Then
  If Not kDB Then
   If Not IsMissing(tfd1) Then
'   Debug.Print "Test auf Vorhandensein:", rsAnm!Pat_ID, rsAnm!Nachname, rsAnm!Vorname
    If Not IsNull(rsAnm.Fields(tfd1)) And rsAnm.Fields(tfd1) <> vNS Then
    Exit Function
'   Else
    End If
   End If
   If Not IsMissing(tfd2) Then
    If Not IsNull(rsAnm.Fields(tfd2)) And rsAnm.Fields(tfd2) <> vNS Then
     Exit Function
'   Else
    End If
   End If
'  rsAnm.Edit
   If InStrB(arten, "|andm|") Then
    For k = 1 To 3
     rsAnm.Fields("DiabetesMedikament " & k & " Menge") = vNS
    Next k
   End If
  End If
  
  ReDim sp(1)
  If Not kDB Then
   sp(0) = rEi(im).Inhalt
  Else
   Dim rEin As New ADODB.Recordset
   myFrag rEin, "SELECT zeitpunkt,inhalt FROM `eintraege` WHERE art RLIKE '^(" & Mid$(arten, 2, Len(arten) - 2) & ")$' AND pat_id = " & rNa(0).Pat_id & " ORDER BY zeitpunkt DESC;"
   If Not rEin.BOF Then
    sp(0) = rEin!Inhalt
    AbIDate = rEin!Zeitpunkt
   End If
  End If
'  Dim AbN$() ' AbschnittName
'  Dim AbI$() ' AbschnittInhalt
  Dim ia% 'i für Abschnitte
  ReDim AbN(0)
  ReDim AbI(0)
  For i = imin To imax
   ReDim Preserve AbN(i - imin)
   ReDim Preserve AbI(i - imin)
   If sp(0) <> vNS Then
    trennz = tr(i)
    Do
     Dim ZifPos As Vorkommen
     trennz = REPLACE$(trennz, " ", "")
'     Debug.Print trennz
     ZifPos = doSp(sp(0), trennz)
     If ZifPos.Beg > 0 Then
      sp(1) = Mid$(sp(0), ZifPos.END + 1)
'      Debug.Print sp(1)
      sp(0) = left$(sp(0), ZifPos.Beg - 1)
#If False Then
     sp = Split(sp(0), trennz)
     If UBound(sp) > 0 Then
      For j = 0 To UBound(sp) - 2
       sp(0) = sp(0) + trennz + sp(j + 1)
      Next j
      If UBound(sp) > 1 Then sp(1) = sp(UBound(sp))
#End If
' wenn falsche Fragmente durch Löschung entstehen
      If InStrB(sp(1), ",") <> 0 And InStr(sp(1), ",") < Len(sp(1)) Then
       For k = 1 To 15
        If i + k > imax Then Exit For
        tz2 = tr(i + k)
        Do
         If InStrB(sp(1), tz2) <> 0 Then
          If InStrB(sp(0), tz2) <> 0 Then Exit Do ' wenn es in der Antwort steht und weiter vorne in der Frage
          sp(0) = sp(0) + trennz + sp(1)
          AbN(i - imin) = vNS ' richtige Trennung fehlte
          GoTo w2 ' falsches Fragment
         End If
         If InStrB(":) ", Right$(tz2, 1)) <> 0 And Right$(tz2, 1) <> vNS Then
          tz2 = left$(tz2, Len(tz2) - 1)
         Else
          Exit Do ' keine falschen Fragmente
         End If
        Loop
       Next
      End If ' instrb(sp(1), ",") <> 0 AND InStr(sp(1), ",") < Len(sp(1)) THEN
      Do
       If sp(1) = vNS Then Exit Do
       If InStrB(EndStr, Right$(sp(1), 1)) <> 0 Then
        sp(1) = left$(sp(1), Len(sp(1)) - 1)
       Else
        Exit Do
       End If
      Loop
      sp(1) = REPLACE$(Trim$(sp(1)), "~", "")
      Select Case sp(1)
       Case "ja": sp(1) = "j"
       Case "kein", "keine", "keinen", "nein", "normal", "nicht bekannt": sp(1) = "n"
       Case "unbekannt": sp(1) = "u"
      End Select
      AbN(i - imin) = trennz
      AbI(i - imin) = Trim$(sp(1))
      Exit Do
     End If ' IF UBound(sp) > 0 THEN
' hier kommt er nur an, wenn trennz nichts trennte
     If InStrB(":)? ", Right$(trennz, 1)) > 0 And Right$(trennz, 1) <> vNS Then
      trennz = left$(trennz, Len(trennz) - 1)
     Else
'      Debug.Print trennz + ":: ", "-----------------"
      AbN(i - imin) = vNS
      Exit Do
     End If
    Loop
   End If ' sp(0)<> ""
w2:
  Next i
' damit li nicht falsch zugeordnet wird, falls irgendwo li und re gelöscht worden sind
  For i = imin To imax - 2
   If AbN(i - imin) = "li" Or AbN(i - imin) = "li:" Then
    j = i - imin
    If AbI(i - imin) <> vNS And AbI(i - imin + 1) = vNS And _
      (AbN(i - imin + 2) = "li" Or AbN(i - imin + 2) = "li:") And AbI(i - imin + 2) = vNS Then
      AbI(i - imin + 2) = AbI(i - imin)
      AbI(i - imin) = vNS
    End If
   End If
  Next i
  
  If Not kDB Then
  
  Dim izuvor%
  For i = imin To imax
'   IF Pat_id = 828 AND makro = "usdm" AND i = -1 THEN
'   IF (makro = "usdm" AND i - imin = 105) OR (makro = "anal" AND i - imin = 4) THEN
   fld = Fd(i)
'   IF i = 82 THEN
   If AbN(i - imin) = vNS Then
    If fld <> vNS Then
     Dim jj%
     For jj = imin To imax
      If jj <> i Then
       If Fd(jj) = fld Then GoTo doppelt
      End If
     Next jj
     Select Case rsAnm.Fields(fld).name
      Case "Größe", "Gewicht", "Tendenz"
      Case Else
       If fld <> vNS Then
        Select Case rsAnm.Fields(fld).Type
         Case 11, 16, 17, 2, 18, 3, 19, 4, 5, 20, 21, 131, 139, 6, 14, 7, 64, 133, 134, 135
              rsAnm.Fields(fld) = 0
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
              rsAnm.Fields(fld) = vNS
        End Select
       End If
     End Select
doppelt:
    End If
   Else ' AbN(i - imin) = vns THEN
    If fld = vNS Then
     If Not i = -1 And InStrB(arten, "|usdm|") Then Err.Raise 999, , "Fehler bei AnAnpassen bei USDM mit Fld = """" und i <> -1"
     GoTo weiter
    End If
'    IF rsAnm.Fields(fld).Type = adUNSIGNEDTinyInt THEN
    AbI(i - imin) = gKw(AbI(i - imin), rsAnm.Fields(fld).Type)
'    Debug.Print AbN(i - imin) + ":: ", AbI(i - imin)
    If Not IsNull(fld) And Not fld = vNS Then
     Dim FproZielFeld% ' Felder pro ZielFeld, 1 = normal, 2 = mindestens 2
     FproZielFeld = 1
     If i > imin Then
      If fld = Fd(izuvor) And AbN(izuvor - imin) <> vNS Then FproZielFeld = 2
     End If
     If FproZielFeld = 2 Then
      If InStrB(fld, "Menge") <> 0 Then
       neuinh = IIf(LenB(AbI(i - imin)) = 0, "0", IIf(AbI(i - imin) = "-", "0", AbI(i - imin))) & " " & "-" & " " & rsAnm.Fields(Fd(izuvor))
       GoSub neuinh
      Else
       If tr(i - 1) = "li:" Or tr(i - 1) = "li" Then
'        rsAnm.Fields(fd(i-imin)) = abn(i-imin) & " " & abi(i-imin) + "/" + rsAnm.Fields(fd(i-1))
        neuinh = AbI(i - imin) + " | " + rsAnm.Fields(Fd(i - 1))
        GoSub neuinh
       Else
        If AbI(izuvor - imin) = vNS Then
         neuinh = AbN(i - imin) & " " & AbI(i - imin) ' falls vorwert leer => von vorne anfangen
        Else
         neuinh = AbI(i - imin) + " ," + AbN(izuvor - imin) & " " & LTrim$(rsAnm.Fields(Fd(izuvor)))
        End If
        GoSub neuinh
       End If
      End If
     Else ' => FproZielFeld = 1
      If fld <> vNS Then
       Select Case rsAnm.Fields(fld).name
       Case "Größe", "Gewicht", "Tendenz"
       Case Else
        Select Case rsAnm.Fields(fld).Type
         Case 11, 16, 17, 2, 18, 3, 19, 4, 5, 20, 21, 131, 139, 6, 14, 7, 64, 133, 134, 135
              rsAnm.Fields(fld) = 0
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
              rsAnm.Fields(fld) = vNS
        End Select
       End Select
      End If
      Select Case rsAnm.Fields(fld).Type
       Case adBoolean, adUnsignedTinyInt '1 ' Logisch
        If obNein(AbI(i - imin)) Then
         rsAnm.Fields(fld) = 0
        Else
         If lies.obMySQL Then
          rsAnm.Fields(fld) = 1
         Else
          rsAnm.Fields(fld) = -1
         End If
        End If
       Case Else
        If InStrB(fld, "Essenszeit") <> 0 Or InStrB(fld, "Broteinheiten") <> 0 Then
         spe = Split(AbI(i - imin), "/")
         If UBound(spe) = -1 Then
          neuinh = AbI(i - imin)
          GoSub neuinh
         Else
          If UBound(spe) = 0 Then
           spe = Split(spe(0), " ")
          End If
          neuinh = gKw(spe(0), rsAnm.Fields(fld).Type)
          GoSub neuinh
          If UBound(spe) > 0 Then
           neuinh = spe(1)
           On Error GoTo f1
           fld = nächstes(fld)
           GoSub neuinh
           On Error GoTo fehler
          End If
         End If
        Else ' instrb(fd(i-imin), "Essenszeit") <> 0 OR instrb(fd(i-imin), "Broteinheiten") <> 0 THEN
         Select Case fld
          Case "Tendenz"
           AbI(i - imin) = left$(AbI(i - imin), 1)
          Case "Diabetestyp"
           Select Case AbI(i - imin)
            Case "path.Glucosetoleranz"
             AbI(i - imin) = "p"
            Case "Gestations-"
             AbI(i - imin) = "g"
            Case Else
             If InStrB(AbI(i - imin), "ekund") <> 0 Or InStrB(AbI(i - imin), "ankreopri") <> 0 Then AbI(i - imin) = "s"
           End Select
         End Select ' fd(i-imin)
           If InStrB(fld, "Menge") <> 0 Then
            neuinh = IIf(AbI(i - imin) = "-", "0", AbI(i - imin)) 'replace$(AbI(i - imin), "-", "0")
            GoSub neuinh
           ElseIf InStrB(arten, "|anal|") And fld = "Grund für Vorstellung" Then
            If IsNull(rsAnm.Fields(fld)) Then
              neuinh = AbI(i - imin)
              GoSub neuinh
            Else
             If AbI(i - imin) <> vNS And InStrB(rsAnm.Fields(fld), AbI(i - imin)) = 0 Then
              If tr(i - 1) = "li:" Then
'              rsAnm.Fields(fd(i-imin)) = abn(i-imin) & " " & abi(i-imin) + "/" + rsAnm.Fields(fd(i - 1))
               neuinh = AbI(i - imin) + " | " + rsAnm.Fields(Fd(i - 1))
               GoSub neuinh
              Else
' bitte prüfen, ob die Reihenfolge stimmt, deshalb stop
               If rsAnm.Fields(fld).name <> "Grund für Vorstellung" Then
                MsgBox "Unerwarteter Zustand do_anImp: " & vbCrLf & "rsAnm.Fields(fld).name (<> ""Grund für Vorstellung""):" & rsAnm.Fields(fld).name
               End If
               neuinh = LTrim$(rsAnm.Fields(fld) & " " & AbI(i - imin))
               GoSub neuinh
              End If
             End If
            End If
           ElseIf InStrB(arten, "|anal|") = 0 Then
            Select Case rsAnm.Fields(fld).Type
             Case adBoolean, adUnsignedTinyInt, _
                  adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adBigInt, adUnsignedBigInt, _
                  adNumeric, adVarNumeric, adCurrency, adDecimal
'            Case 2,17,  16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
              neuinh = AbI(i - imin)
              rsAnm.Fields(fld) = MachNumerisch(neuinh)
             Case adUnsignedSmallInt, adUnsignedInt
'             Case 18, 19
              neuinh = AbI(i - imin)
              rsAnm.Fields(fld) = Abs(MachNumerisch(neuinh))
             Case adDate, adFileTime, adDBDate, adDBTime, adDBTimeStamp
'             Case 7, 64, 133, 134, 135
              If Not IsDate(AbI(i - imin)) Then
                rsAnm.Fields(fld) = CDate(0)
              Else
               rsAnm.Fields(fld) = CDate(datUmwandel(AbI(i - imin), "dd.mm.yy"))  'CDate(AbI(i - imin))
              End If
             Case adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adEmpty, _
              adIDispatch, adVariant, adIUnknown, adGUID, adBinary, adUserDefined, adPropVariant, _
              adVarBinary, adLongVarBinary, adError, adArray
'            Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205, 10, 8192
              neuinh = AbI(i - imin)
              rsAnm.Fields(fld).Value = left$(neuinh, rsAnm.Fields(fld).DefinedSize)
            End Select
           Else
            FproZielFeld = 2
            If (fld = "Größe" Or fld = "Gewicht" Or fld = "Tendenz") Then
             If rsAnm.Fields(fld) = vNS Or rsAnm.Fields(fld) = 0 Or IsNull(rsAnm.Fields(fld)) Then
              FproZielFeld = 1
             End If
            Else
             FproZielFeld = 1
            End If
            If FproZielFeld = 1 Then
              neuinh = AbI(i - imin)
              Select Case rsAnm.Fields(fld).Type
               Case 11, adUnsignedTinyInt, 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14, 7, 64, 133, 134, 135
                 neuinh = CDbl(IIf(LenB(neuinh) = 0, "0", neuinh))
                 GoSub neuinh
               Case adUnsignedTinyInt, 18, 19
                 neuinh = CDbl(IIf(LenB(neuinh) = 0, "0", neuinh))
                 rsAnm.Fields(fld) = Abs(neuinh)
               Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
                 neuinh = CStr(neuinh)
                 GoSub neuinh
              End Select
            End If
           End If
        End If ' instrb(fd(i-imin), "Essenszeit") <> 0 OR instrb(fd(i-imin), "Broteinheiten") <> 0 THEN
      End Select
     End If
    End If
weiter:
'    vorwertnull = 0
'    IF AbI(i - imin) = vns THEN vorwertnull = -1
   End If ' abn(i-imin)<> ""
   If fld <> "" Then
     Debug.Print fld, arten, rsAnm.Fields(fld)
   Else
     Debug.Print fld, arten
   End If
   izuvor = i
  Next i
  If InStrB(arten, "|andm|") Then
   For k = 4 To 2 Step -1
    If IsNull(rsAnm.Fields("DiabetesMedikament " + CStr(k) + " Menge")) Or (rsAnm.Fields("DiabetesMedikament " + CStr(k) + " Menge") = "0 - 0 - 0 - ") Then 'And rsAnm.Fields("DiabetesMedikament " + CStr(K)) = "") THEN
     neuinh = vNS
     fld = "DiabetesMedikament " + CStr(k) + " Menge"
     GoSub neuinh
    End If
   Next k
   For k = 4 To 2 Step -1
    If rsAnm.Fields("DiabetesMedikament " + CStr(k) + " Menge") <> vNS And rsAnm.Fields("DiabetesMedikament " + CStr(k - 1) + " Menge") = vNS Then
     neuinh = rsAnm.Fields("DiabetesMedikament " + CStr(k) + " Menge")
     fld = "DiabetesMedikament " + CStr(k - 1) + " Menge"
     GoSub neuinh
     neuinh = vNS
     fld = "DiabetesMedikament " + CStr(k) + " Menge"
     rsAnm.Fields(fld) = neuinh
    End If
   Next k
   If IsNull(rsAnm.Fields("Angina pectoris")) Then rsAnm.Fields("Angina pectoris") = vNS
   If IsNull(rsAnm.Fields("Herzinfarkt")) Then rsAnm.Fields("Herzinfarkt") = vNS
   If IsNull(rsAnm.Fields("PTCA oder Stent")) Then rsAnm.Fields("PTCA oder Stent") = vNS
   If IsNull(rsAnm.Fields("Herzschwäche")) Then rsAnm.Fields("Herzschwäche") = vNS
   If IsNull(rsAnm.Fields("Bypass kardial")) Then rsAnm.Fields("Bypass kardial") = 0
   If Not obNein(rsAnm.Fields("Angina pectoris")) Or Not obNein(rsAnm!Herzinfarkt) Or Not obNein(rsAnm.Fields("PTCA oder Stent")) Or Not obNein(rsAnm.Fields("Bypass kardial")) Or Not obNein(rsAnm!Herzschwäche) Then
    neuinh = "j"
    fld = "Herzkrankheit"
    rsAnm.Fields(fld) = neuinh
   ElseIf rsAnm.Fields("Angina pectoris") = vNS And rsAnm!Herzinfarkt = vNS And rsAnm.Fields("PTCA oder Stent") = vNS And rsAnm.Fields("Bypass kardial") = 0 And rsAnm!Herzschwäche = vNS Then
    neuinh = vNS
    fld = "Herzkrankheit"
    rsAnm.Fields(fld) = neuinh
   Else
    neuinh = "-"
    fld = "Herzkrankheit"
    rsAnm.Fields(fld) = neuinh
   End If
  End If ' makro = "andm"
' Nacharbeiten: Wenn Herzkrankheit Beschreibung positiv => dann auch Herzkrankheit
' Essenzeiten aufteilen
'  IF rsAnm!Pat_id = 0 THEN
'   rsAnm.CancelUpdate ' kommt auf mir noch unbekanntem Weg zustande
'  Else
'   rsAnm.Update
'  END IF
' Else
'  IF rsAnm.NoMatch THEN MsgBox "Pat. " + CStr(Pat_id) + " nicht in der Anamnesetabelle gefunden"
'  IF rEi.NoMatch THEN SysCmd 4, "Pat. " + CStr(Pat_id) + " nicht in Tabelle 'eintraege' gefunden" 'MsgBox "Pat. " + CStr(Pat_id) + " nicht in den eintraegen gefunden"
  
 End If ' Not rEi.NoMatch AND NOT rsAnm.NoMatch THEN
End If ' not kdm
Exit Function
neuinh:
If Right$(neuinh, 5) = "\r \r" Then
 neuinh = left$(neuinh, Len(neuinh) - 5)
ElseIf Right$(neuinh, 2) = "\r" Then
 neuinh = left$(neuinh, Len(neuinh) - 2)
End If
rsAnm.Fields(fld) = neuinh
Return
Dim Feld$
f1:
' Debug.Print fld, NeuInh
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147217887 Then ' Das Feld ist zu klein für die Datenmenge, die Sie hinzufügen wollten. Versuchen Sie, weniger Daten einzufügen. / Die Datenmenge ist zu groß
 Dim rsc As ADODB.Recordset
 Static Versuch%, MerkNeuInh$
 Dim AfN&
 AfN = 0
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "anamnesebogen", Empty))
 Dim ZCStr$, DBStrAlt$, SpName$
 ZCStr = rsAnm.source
 Versuch = Versuch + 1
 If MerkNeuInh <> neuinh Then Versuch = 0
 If Versuch > 5 Then
  If neuinh > vNS Then neuinh = left$(neuinh, Len(neuinh) - 1)
 End If
 If fld = vNS Then fld = Fd(i)
 Do
  If rsc!COLUMN_NAME = rsAnm.Fields(fld).name Then
   SpName = rsc!COLUMN_NAME
'   If SpName = "Diabetestyp" Then Stop
'   IF Not SpModAlt(NeuInh, "anamnesebogen", rsc) THEN ' dann Memo-Feld
'    Call myEFrag("UPDATE `anamnesebogen` SET " & "`" & SpName & "`" & " = """ & replace$(NeuInh, """", """""") & """ WHERE pat_id = " & rNa(0).Pat_id, AfN)
'   END IF
   If Not SpMod(Len(neuinh), "anamnesebogen", rsc, neuinh) Then ' dann Memo-Feld
    Call myEFrag("UPDATE `anamnesebogen` SET `" & SpName & "` = """ & REPLACE$(neuinh, """", """""") & """ WHERE pat_id = " & rNa(0).Pat_id, AfN)
   End If
   If rsAnm.State = 0 Then myFrag rsAnm, ZCStr, adOpenStatic, DBCn, adLockOptimistic
   MerkNeuInh = neuinh
   If AfN > 0 Then Resume Next Else Resume
  End If
  rsc.Move 1
 Loop
End If
If rsAnm.State = 0 Then
 myFrag rsAnm, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & rNa(0).Pat_id, adOpenStatic, DBCn, adLockOptimistic
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_anImp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'AnImp

' 4.9.06: wird aufgerufen in doTabakStAlt, do_AnImp und gkw, dodoPLZ
Function MachNumerisch#(ByVal ST$, Optional erstDatum%)
 Dim nachziffer%, pos%, stor$, runde%
 On Error GoTo fehler
 stor = ST
 If IsNumeric(ST) Then
'  MachNumerisch = ST
 ElseIf ST = vNS Then
'  MachNumerisch = "0"
 Else
  nachziffer = 0
  For pos = 1 To Len(ST)
   If InStrB("0123456789" + IIf(nachziffer, " ,.", ""), Mid$(ST, pos, 1)) = 0 And Mid$(ST, pos, 1) <> vNS Then
    ST = left$(ST, pos - 1) & " " & Mid$(ST, pos + 1)
   Else
    nachziffer = True
   End If
  Next
 End If
 ST = Trim$(ST)
 If ST = vNS Then
  ST = 0
 Else ' ST = vNS THEN
  pos = InStr(ST, "  ")
  If pos > 0 And IsNumeric(left$(ST, pos)) Then
   ST = left$(ST, pos)
  End If
  For runde = 1 To 2
   If (erstDatum And (runde = 1)) Or (Not erstDatum And (runde = 2)) Then
    If IsDate(ST) Then
     MachNumerisch = CDate(ST): Exit For
    ElseIf IsDate(left$(REPLACE$(ST, ",", "."), 10)) Then
     MachNumerisch = CDate(left$(REPLACE$(ST, ",", "."), 10)): Exit For
    ElseIf IsDate("1.1." & REPLACE(REPLACE$(ST, ",", "."), "/", ".")) Then
     MachNumerisch = CDate("1.1." & REPLACE(REPLACE$(ST, ",", "."), "/", ".")): Exit For
    ElseIf IsDate("1." & REPLACE(REPLACE$(ST, ",", "."), "/", ".")) Then
     MachNumerisch = CDate("1." & REPLACE(REPLACE$(ST, ",", "."), "/", ".")): Exit For
    End If
   Else ' (erstDatum AND (runde = 1)) OR (Not erstDatum AND (runde = 2)) THEN
    If IsNumeric(ST) Then
     Dim stk$
     If InStrB(ST, ".") <> 0 Then
      stk = REPLACE$(ST, ".", ",")
      If IsNumeric(stk) Then
       MachNumerisch = stk
      Else
       MachNumerisch = ST
      End If
     Else
      MachNumerisch = ST
     End If
     Exit For
    Else
     Dim stneu$, buchst$, Posi&
     stneu = vNS
     For Posi = 1 To Len(ST)
      buchst = Mid$(ST, Posi, 1)
      If IsNumeric(buchst) Or buchst = "." Or buchst = "," Then
       stneu = stneu + buchst
      Else
       Exit For
      End If
     Next Posi
     stneu = REPLACE$(stneu, ",,", ",")
     Posi = InStr(stneu, ",") ' 16.7.09
     If Posi > 0 Then
      Posi = InStr(Posi + 1, stneu, ",")
      If Posi <> 0 Then stneu = left$(stneu, Posi - 1)
     End If
     If IsNumeric(stneu) Then
      ST = stneu
     Else ' IsNumeric(stneu) THEN
      If IsDate(REPLACE$(stneu, ",", ".")) Then
       ST = CDate(REPLACE$(stneu, ",", "."))
      Else
 ' Achtung: abenteuerliche Typumwandlung:
       Select Case stor
        Case "Z.n. Aortenklappenersatz 1/2002": ST = "2002"
        Case "114,7 (vor der Schwangerschaft 116 kg": ST = "116"
        Case "76 kg (vor der Schwangerschaft 58 kg": ST = "58"
  '     Case "86 kg, Rauchen: nein"
        Case "54-55 kg": ST = "54"
        Case "73 vor gut einem Jahr bei 80 kg, sonst eher bei 73-75 kg, bei Beginn der Bauchspeicheldrüsenerkrankung 68 kg"
         ST = "73"
        Case "83 (hatte 95 kg bis vor 1 Monat"
         ST = "83"
        Case "87, im Dezember 93"
         ST = "87"
        Case Else
         Err.Raise 999, , "Unbehandelter Fall in machnumerisch: " & stor & " bei Pat.: " & rNa(0).Pat_id
         MsgBox stneu & " nicht numerisch zu bekommen (machmumerisch)"
         ST = "0"
       End Select
       Exit Function
      End If ' IsDate(replace$(stneu, ",", ".")) THEN
     End If ' IsNumeric(stneu) THEN / else
 '    Debug.Print ST
 '    Debug.Print ST
     If InStrB(ST, " ") <> 0 Then
      ST = left$(ST, InStr(ST, " ") - 1)
     End If
     If IsNumeric(ST) Then ' OR Not IsDate(replace$(ST, ",", ".")) THEN
      MachNumerisch = ST: Exit For
'     ELSE ' IsNumeric(ST) OR Not IsDate(replace$(ST, ",", ".")) THEN
'      MachNumerisch = replace$(ST, ",", ".")
     End If ' IsNumeric(ST) OR Not IsDate(replace$(ST, ",", ".")) THEN
    End If ' IsDate(ST) THEN elseif elseif else
   End If ' (erstDatum AND (runde = 1)) OR (Not erstDatum AND (runde = 2)) THEN
  Next runde
 End If ' ST = vNS THEN else
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachNumerisch/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MachNumerisch

Function usAl()
 Const imin% = 0
 Const imax% = 16
 Dim tr$(imin To imax)
 Dim Fd$(imin To imax)
 On Error GoTo fehler
 tr(0) = "Neuro sonst:"
 Fd(0) = "Neuro sonst"
 tr(1) = "LR:"
 Fd(1) = "Neuro sonst"
 tr(2) = "KR:"
 Fd(2) = "Neuro sonst"
 tr(3) = "Lymphknoten:"
 Fd(3) = "LK"
 tr(4) = "Beinödeme/Venenkrankheiten:"
 Fd(4) = "BeinödVen"
 tr(5) = "Venenkrankheiten:"
 Fd(5) = "BeinödVen"
 tr(6) = "Beinödeme"
 Fd(6) = "BeinödVen"
 tr(7) = "Mundhöhle:"
 Fd(7) = "Mundhöhle"
 tr(8) = "Zähne:"
 Fd(8) = "Zähne"
 tr(9) = "NNH:"
 Fd(9) = "NNH"
 tr(10) = "SD:"
 Fd(10) = "SD"
 tr(11) = "NL:"
 Fd(11) = "NL"
 tr(12) = "WS:"
 Fd(12) = "WS"
 tr(13) = "Abdomen:"
 Fd(13) = "Bauch"
 tr(14) = "Pulmo:"
 Fd(14) = "Lunge"
 tr(15) = "Halsschlagadern:"
 Fd(15) = "Carotiden"
 tr(16) = "Cor:"
 Fd(16) = "Herz"
 Call do_anImp(imin, imax, tr(), Fd(), "|usal|")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usAl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usAl


Function usdmAlt(Optional fürDMP%)
 Const imin% = -2
 Const imax% = 35
 Dim tr$(imin To imax)
 Dim Fd$(imin To imax)
 On Error GoTo fehler
 tr(-2) = "aktueller RR"
 tr(-1) = "aktuellen Blutdruck und ggf. Puls bitte extra eingeben"
 tr(0) = "li:"
 Fd(0) = "Puls Adp"
 tr(1) = "Puls der re A.dors.ped.: re:"
 Fd(1) = "Puls Adp"
 tr(2) = "li:"
 Fd(2) = "Puls Atp"
 tr(3) = "Puls der re A.tib.post.: re:"
 Fd(3) = "Puls Atp"
 tr(4) = "li:"
 Fd(4) = "Puls Kniekehle"
 tr(5) = "Puls in der Kniekehle: re:"
 Fd(5) = "Puls Kniekehle"
 tr(6) = "li:"
 Fd(6) = "Puls Leiste"
 tr(7) = "Puls in der Leiste: re:"
 Fd(7) = "Puls Leiste"
 tr(8) = "li:"
 Fd(8) = "Vibration Großzehe"
 tr(9) = "Vibration Großzehenballen: re:"
 Fd(9) = "Vibration Großzehe"
 tr(10) = "li:"
 Fd(10) = "Vibration IK"
 tr(11) = "Vibration Innenknöchel: re:"
 Fd(11) = "Vibration IK"
 tr(12) = "li:"
 Fd(12) = "Kalt-Warm"
 tr(13) = "Kalt-Warm: re:"
 Fd(13) = "Kalt-Warm"
 tr(14) = "li:"
 Fd(14) = "Monofilamenttest"
 tr(15) = "Monofilamenttest Fußsohle: re:"
 Fd(15) = "Monofilamenttest"
 tr(16) = "li:"
 Fd(16) = "Oberflächensensibilität"
 tr(17) = "Oberflächensensibilität: re:"
 Fd(17) = "Oberflächensensibilität"
 tr(18) = "li:"
 Fd(18) = "PSR"
 tr(19) = "PSR: re:"
 Fd(19) = "PSR"
 tr(20) = "li:"
 Fd(20) = "ASR"
 tr(21) = "ASR: re:"
 Fd(21) = "ASR"
 tr(22) = "li:"
 Fd(22) = "Kraft Knie"
 tr(23) = "grobe Kraft Knie: re:"
 Fd(23) = "Kraft Knie"
 tr(24) = "li:"
 Fd(24) = "Kraft Zehenbeuger"
 tr(25) = "grobe Kraft Zehenbeuger: re:"
 Fd(25) = "Kraft Zehenbeuger"
 tr(26) = "li:"
 Fd(26) = "Kraft Zehenheber"
 tr(27) = "grobe Kraft Zehenheber: re:"
 Fd(27) = "Kraft Zehenheber"
 tr(28) = "li"
 Fd(28) = "Ulcera"
 tr(29) = "Ulcera: re:"
 Fd(29) = "Ulcera"
 tr(30) = "li:"
 Fd(30) = "Hyperkeratosen"
 tr(31) = "Hyperkeratosen u.ä.: re:"
 Fd(31) = "Hyperkeratosen"
 tr(32) = "li:"
 Fd(32) = "Beinbefund"
 tr(33) = "Fußbefund (Inspektion): re:"
 Fd(33) = "Beinbefund"
 tr(34) = "li:"
 Fd(34) = "Liphypertrophien Abdomen"
 tr(35) = "Liphypertrophien: re:"
 Fd(35) = "Liphypertrophien Abdomen"
 Call do_anImp(imin, imax, tr(), Fd(), "|usd|usdm|usdm1|", "Kraft Zehenbeuger", "Vibration IK", fürDMP) ' wenn für DMP, dann kein Datenbankeintrag
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usDMAlt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' usDMAlt

Function nächstes(ByVal ez$)
 Dim Spl$(), vs$
 On Error GoTo fehler
 Spl = Split(ez)
 vs = Spl(0)
 ez = Spl(1)
 Select Case ez$
  Case "abends": nächstes = vs + IIf(vs = "Essenszeit", " spät", " nachts")
  Case "mittags": nächstes = vs + " nachmittags"
  Case "früh": nächstes = vs + IIf(vs = "Essenszeit", " vormittags", " ZM früh")
 End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in nächstes/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' nächstes(ez$)

Function obNein%(antw$)
 Dim a$
 On Error GoTo fehler
 a = Trim$(LCase$(antw))
 If InStrB(a, "nein,") <> 0 Then
  obNein = -1
 Else
  Select Case a
   Case vNS, "-", "n", "nie", "normal", "nein", "0", 0, "falsch", "false", "- | -", "|", "o.B | o.B.", "n | n", "keine | keine", "o.B", "o. B", "o.B.", "|-", "-|", "| -", "- |", "-|-", "| n", "|n", "n |", "n|", "entfällt | entfällt", "re: keine , li: keine", "keine|keine", "keine: | keine"
    obNein = -1
  End Select
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obNein/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obNein

Function obUnbek%(antw$)
 Dim a$
 On Error GoTo fehler
 a = Trim$(LCase$(antw))
 Select Case a
  Case vNS, "unbek"
   obUnbek = -1
 End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obUnbek/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obNein

Function gKw$(ST$, Optional Typ)
 Dim p1&, pe&, buch
' IF ST = "81kg" THEN Stop
 On Error GoTo fehler
 gKw = ST
 p1 = InStr(gKw, "~{")
 If p1 > 0 Then
  gKw = Trim$(left$(gKw, p1 - 1))
 Else
  Do
   p1 = InStr(gKw, "{")
   pe = InStr(p1 + 1, gKw, "}")
   If pe = 0 Then pe = Len(gKw)
   If p1 > 0 And pe > p1 Then
    gKw = left$(gKw, p1 - 1) + Mid$(gKw, pe + 1)
   End If
   If gKw = "{" Then gKw = vNS: p1 = 0
  Loop Until p1 = 0
 End If
 gKw = Trim$(gKw)
nochmal:
 If Not IsMissing(Typ) Then
  Select Case Typ
   Case adBoolean ' 6.4.07: Fehler bei Insulinpumpen mit logischem Feld
'   case 11
   Case adUnsignedTinyInt ' 9.9.08: jetzt hier Fehler bei Insulinpumpen
'   case 17
   Case adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adBigInt, adUnsignedBigInt, _
        adNumeric, adVarNumeric, adCurrency, adDecimal
'         Case 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
     gKw = MachNumerisch(gKw)
   Case adDate, adFileTime, adDBDate, adDBTime, adDBTimeStamp
'   Case 7, 64, 133, 134, 135
     gKw = MachNumerisch(gKw, True)
   Case adUnsignedSmallInt, adUnsignedInt
'   Case 18, 19
     gKw = Abs(MachNumerisch(gKw))
   Case adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adEmpty, _
        adIDispatch, adVariant, adIUnknown, adGUID, adBinary, adUserDefined, adPropVariant, _
        adVarBinary, adLongVarBinary, adError, adArray
'   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205, 10, 8192
  End Select
'  gKw = ST
'  GoTo nochmal
'  IF Typ < 8 THEN
'   Debug.Print gKw
'   gKw = MachNumerisch(gKw)
'  END IF
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in gKw/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' gKw

Function datUmwandel$(fldv$, Optional fmt$ = "mm\/yyyy") ' für gemessen am, 30.6.06
     datUmwandel = Format$(fldv, fmt)
     Dim gemaSpli$()
     gemaSpli = Split(fldv, "/")
     If UBound(gemaSpli) = 1 Then
      If IsNumeric(gemaSpli(0)) And IsNumeric(gemaSpli(1)) Then
       datUmwandel = Format$(DateSerial(gemaSpli(1), gemaSpli(0), 1), fmt)
      End If
     End If
End Function ' datUmwandel

