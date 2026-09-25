Attribute VB_Name = "DiagAnzeige"
Option Explicit
' ===========================================================================================
' DiagAnzeige.bas - neue Diagnosenanzeige (Stand 21.9.2026)
'
' Wird von DiagString (Importiere.bas) aufgerufen, wenn fuer den Patienten Zeilen mit MOStatus > 0
' in `diagnosen` stehen (= mit der neuen Uebertragung aus medoff geladen). Alle anderen Patienten
' behalten die alte Logik in DiagString unveraendert. Bei einem Fehler in diesem Modul faellt
' DiagString ebenfalls auf die alte Logik zurueck (ok = False).
'
' Ablauf (Referenzlogik in Referenz_Python\neuanzeige.py, gleiche Regeln):
'  1. Zeilen laden (SQL mit den Funktionen diag_info/diag_jahr aus der Datenbank quelle)
'  2. aktive Zeilen (MOStatus <> 4): identische Kopien zusammenfassen; aktive Zeilen ausser Z.n. und
'     Ausschluss wie bisher je ICD-3 + Gewissheit + erste 17 Zeichen zusammenfassen
'  3. aktive Z.n. gegen aktive gesichert je ICD-3: Paarregeln aus Tabelle diagpaare
'     (Code g z j b a oder Zahl, mit Bedingungen); Z.n. ohne gesicherten Partner: nach diagentscheid
'  4. abgeschlossene Zeilen (MOStatus 4): Kopien zusammenfassen; weglassen bei aktivem Gegenstueck
'     oder bei spezifischerer abgeschlossener Zeile; sonst Entscheidung aus Tabelle diagentscheid
'     (n weg, j einmal gesichert, jz einmal Z.n., je einzeln Z.n.)
'  5. Label (Z.n. / V.a. / Z.n. V.a. / Ausschluss) wird neu gesetzt, fuehrende Praefixe im Text entfallen
'  6. die Arrays Diag, ICD, DSic ... (Importiere.bas) werden gefuellt, MachDiagnosen erledigt wie
'     bisher die Umformungen (Diabetes, Fusssyndrom ...) und die Ausgabe
' Pflegetabellen in der Datenbank quelle: diagentscheid, diagpaare, diagnormal (siehe Offene_Punkte_Diagnosenanzeige.md)
' Hinweis: Die Umformungen in AnzeigeText sind eine Kopie der Umformungen in MachDiagnosen (Importiere.bas);
'          bei Aenderungen dort bitte hier nachziehen.
' ===========================================================================================

Private Type DZ
 Icd As String          ' ICD
 Text As String         ' Text ohne fuehrende Labels (gesichert, Z.n., V.a., ausgeschl.)
 Roh As String         ' Text wie gespeichert
 Sich As String         ' gespeichert: G V Z A oder Leerzeichen
 Seite As String        ' R L B oder Leerzeichen
 Attr As String         ' Erlaeuterung
 Status As Long         ' MOStatus (4 = abgeschlossen)
 obDauer As Integer
 Datum As Date
 Jahr As Long           ' Jahr im Text, sonst Eintragsjahr (0 = unbekannt)
 Dggel As Integer
 obKasse As Integer
 lKasse As Date
 KFdFA As String
 Aus As String
 iBm As String
 G1 As Long
 G2 As Long
 IcdK As String         ' ICD ohne abschliessenden Strich (J45.9- = J45.9), nur fuer den Abgleich
 Gew As String          ' Gewissheit: G, V oder A
 Verg As Boolean        ' selbst als Z.n. gekennzeichnet
 HatZ As Boolean        ' Text beginnt mit einem Z.n.-Praefix
 VZ As Boolean          ' Praefixfolge V.a. dann Z.n.
 B0 As String           ' normalisierter Text ohne Datum/Seite
 B1 As String           ' B0 mit Abkuerzungen/Synonymen (Tabelle diagnormal)
 Kern As String         ' B1 bis zum ersten Komma/Klammer
 W1 As String           ' Woerter von B1 (Satzzeichen durch Leerzeichen ersetzt)
 Rest As String         ' Woerter von B1 nach dem Kern
 Daten As String        ' Datumsangaben im Text
 Seiten As String       ' L, R oder LR (Text und Seitenfeld)
 ErlW As String         ' Woerter der Erlaeuterung
 RB As String           ' Text, den die Anzeige aus der ICD erzeugt (nur wenn er den Text ersetzt)
 TNorm As String        ' klein, ohne Praefix und Datum (fuer Textmuster)
 Aktiv As Boolean
 Weg As Boolean
 Gewaehlt As Boolean
 FlagB As Boolean
 FlagA As Boolean
 Gepaart As Boolean
 Label As String
End Type

Private Zl() As DZ
Private zN As Long
Private eN As Long, pN As Long
Private ePre() As String, eTxt() As String, eAlt() As Long, eDec() As String, ePrio() As Long
Private pPre() As String, pCond() As String, pCode() As String
Private eGeladen As Boolean
Private PatAlter As Long

Public DiagNeuAbschalten As Boolean ' True = immer die alte Logik (zum Vergleichen im Direktfenster)

' ------------------------------------------------------------------------------------------
Public Function DiagNeuAktiv(ByVal Pat_ID As String) As Boolean
 Dim rs As New ADODB.Recordset
 If DiagNeuAbschalten Then Exit Function
 On Error GoTo fehler
 myFrag rs, "SELECT 1 x FROM diagnosen WHERE pat_id = " & Pat_ID & " AND COALESCE(MOStatus,0) > 0 LIMIT 1"
 DiagNeuAktiv = Not rs.BOF
 Exit Function
fehler:
 DiagNeuAktiv = False
End Function

Public Sub DiagRegelnNeuLaden()
 eGeladen = False
End Sub

Public Function DiagStringNeu(ByVal Pat_ID As String, DiagTab() As CString, ByVal VorDat As Date, ByVal obBrief As Integer, ByVal dmseit As String, ByRef ok As Boolean) As String
 On Error GoTo fehler
 ok = False
 RegelnLaden
 PatAlterLaden Pat_ID
 LadeZeilen Pat_ID, VorDat, obBrief
 If zN = 0 Then
  ok = True
  Exit Function
 End If
 Auswahl dmseit
 AusgabeFuellen dmseit
 DiagStringNeu = MachDiagnosen(Pat_ID, DiagTab, dmseit)
 ok = True
 Exit Function
fehler:
 ok = False
 Debug.Print "DiagStringNeu (Pat " & Pat_ID & "): " & Err.Number & " " & Err.Description
End Function

' ------------------------------------------------------------------------------------------
' Hilfsfunktionen
Private Function U(ByVal s As String) As String
 ' ersetzt Platzhalter durch Umlaute, damit dieses Modul reines ASCII bleibt
 s = Replace(s, "{ae}", ChrW$(228))
 s = Replace(s, "{ue}", ChrW$(252))
 s = Replace(s, "{Ue}", ChrW$(220))
 s = Replace(s, "{ss}", ChrW$(223))
 U = s
End Function

Private Function Sp(ByVal s As String) As String
 s = Replace(s, vbTab, " ")
 Do While InStr(s, "  ") > 0
  s = Replace(s, "  ", " ")
 Loop
 Sp = Trim$(s)
End Function

Private Function Zahl(ByVal v As Variant) As Long
 ' tinyint(1) kommt je nach Treiber als Boolean: True = 1
 If IsNull(v) Then Exit Function
 If VarType(v) = vbBoolean Then
  Zahl = IIf(v, 1, 0)
 Else
  Zahl = CLng(Val("" & v))
 End If
End Function

Private Function FindeKey(c As Collection, ByVal key As String) As Long
 Dim v As Variant
 On Error Resume Next
 FindeKey = 0
 v = c(key)
 If Err.Number = 0 Then FindeKey = CLng(v)
 Err.Clear
End Function

Private Function Neuer(ByVal a As Long, ByVal b As Long) As Boolean
 ' a ist strikt neuer als b (Datum ohne Uhrzeit, dann laengerer Text)
 Dim da As Long, db As Long
 da = CLng(Int(Zl(a).Datum)): db = CLng(Int(Zl(b).Datum))
 If da <> db Then
  Neuer = (da > db)
 ElseIf Len(Zl(a).Text) <> Len(Zl(b).Text) Then
  Neuer = (Len(Zl(a).Text) > Len(Zl(b).Text))
 End If
End Function

Private Function NeuerSpaeter(ByVal a As Long, ByVal b As Long) As Boolean
 ' wie Neuer; bei Gleichstand gewinnt der spaetere Index
 Dim da As Long, db As Long
 da = CLng(Int(Zl(a).Datum)): db = CLng(Int(Zl(b).Datum))
 If da <> db Then
  NeuerSpaeter = (da > db)
 ElseIf Len(Zl(a).Text) <> Len(Zl(b).Text) Then
  NeuerSpaeter = (Len(Zl(a).Text) > Len(Zl(b).Text))
 Else
  NeuerSpaeter = (a > b)
 End If
End Function

Private Function TeilmengeW(ByVal a As String, ByVal b As String) As Boolean
 ' alle durch Leerzeichen getrennten Woerter von a kommen in b vor (leer = immer erfuellt)
 Dim v() As String, i As Long
 TeilmengeW = True
 a = Trim$(a)
 If Len(a) = 0 Then Exit Function
 b = " " & b & " "
 v = Split(a, " ")
 For i = 0 To UBound(v)
  If Len(v(i)) > 0 Then
   If InStr(1, b, " " & v(i) & " ", vbBinaryCompare) = 0 Then
    TeilmengeW = False
    Exit Function
   End If
  End If
 Next i
End Function

Private Function TeilmengeS(ByVal a As String, ByVal b As String) As Boolean
 ' jedes Zeichen (L, R) von a kommt in b vor
 Dim i As Long
 TeilmengeS = True
 For i = 1 To Len(a)
  If InStr(b, Mid$(a, i, 1)) = 0 Then
   TeilmengeS = False
   Exit Function
  End If
 Next i
End Function

Private Function SeitenGemeinsam(ByVal a As String, ByVal b As String) As Boolean
 Dim i As Long
 If Len(a) = 0 Or Len(b) = 0 Then Exit Function
 For i = 1 To Len(a)
  If InStr(b, Mid$(a, i, 1)) > 0 Then
   SeitenGemeinsam = True
   Exit Function
  End If
 Next i
End Function

Private Function PatLike(ByVal pat As String, ByVal s As String) As Boolean
 ' SQL-Muster mit % (ohne % = exakter Vergleich), Gross-/Kleinschreibung egal
 Dim p As String
 pat = LCase$(pat): s = LCase$(s)
 If InStr(pat, "%") = 0 Then
  PatLike = (s = pat)
 Else
  p = Replace(pat, "[", "[[]")
  p = Replace(p, "?", "[?]")
  p = Replace(p, "#", "[#]")
  p = Replace(p, "%", "*")
  PatLike = (s Like p)
 End If
End Function

Private Function IstBuchstabe(ByVal c As String) As Boolean
 If Len(c) = 0 Then Exit Function
 IstBuchstabe = (c Like "[A-Za-z]") Or (AscW(c) > 191)
End Function

Private Function WortEnde(ByVal u As String, ByVal pos As Long) As Boolean
 ' an Position pos (1-basiert) folgt kein Buchstabe mehr
 If pos > Len(u) Then
  WortEnde = True
 Else
  WortEnde = Not IstBuchstabe(Mid$(u, pos, 1))
 End If
End Function

Private Function Woerter(ByVal s As String) As String
 ' ersetzt alles ausser Buchstaben, Ziffern und Unterstrich durch Leerzeichen
 Dim i As Long, c As String, r As String
 For i = 1 To Len(s)
  c = Mid$(s, i, 1)
  If (c Like "[A-Za-z0-9_]") Or (AscW(c) > 191) Then
   r = r & c
  Else
   r = r & " "
  End If
 Next i
 Woerter = r
End Function

Private Function KlammerZn(ByVal u As String) As Long
 ' (Z.n.) [Z.n.] (wohl Z.n.) (Z.n.?) am Anfang: Laenge, sonst 0
 Dim c As String, p As Long, inner As String
 KlammerZn = 0
 c = Left$(u, 1)
 If c <> "(" And c <> "[" Then Exit Function
 p = InStr(2, u, IIf(c = "(", ")", "]"))
 If p = 0 Or p > 20 Then Exit Function
 inner = LCase$(Trim$(Mid$(u, 2, p - 2)))
 If Left$(inner, 5) = "wohl " Then inner = Trim$(Mid$(inner, 6))
 If Right$(inner, 1) = "?" Then inner = Trim$(Left$(inner, Len(inner) - 1))
 If inner = "z.n." Or inner = "z.n" Or inner = "z. n." Or inner = "zn" Or inner = "z n" Or inner = "zn." Then KlammerZn = p
End Function

Private Function PraefixAb(ByVal t As String, ByRef seq As String) As String
 ' entfernt fuehrende Labels; seq enthaelt deren Reihenfolge (Z, V, G, A)
 Dim u As String, n As Long, k As Long
 seq = ""
 Do
  t = LTrim$(t)
  u = LCase$(t)
  n = 0
  If Left$(u, 4) = "z.n." Then
   n = 4: seq = seq & "Z"
  ElseIf Left$(u, 5) = "z. n." Then
   n = 5: seq = seq & "Z"
  ElseIf Left$(u, 3) = "zn." Then
   n = 3: seq = seq & "Z"
  ElseIf Left$(u, 12) = "zustand nach" And WortEnde(u, 13) Then
   n = 12: seq = seq & "Z"
  ElseIf KlammerZn(u) > 0 Then
   n = KlammerZn(u): seq = seq & "Z"
  ElseIf Left$(u, 4) = "v.a." Then
   n = 4: seq = seq & "V"
  ElseIf Left$(u, 5) = "v. a." Then
   n = 5: seq = seq & "V"
  ElseIf Left$(u, 12) = "verdacht auf" And WortEnde(u, 13) Then
   n = 12: seq = seq & "V"
  ElseIf Left$(u, 8) = "verdacht" And WortEnde(u, 9) Then
   n = 8: seq = seq & "V"
  ElseIf Left$(u, 9) = "gesichert" And WortEnde(u, 10) Then
   n = 9: seq = seq & "G"
  ElseIf Left$(u, 14) = "ausgeschlossen" And WortEnde(u, 15) Then
   n = 14: seq = seq & "A"
  ElseIf Left$(u, 10) = "ausgeschl." Then
   n = 10: seq = seq & "A"
  ElseIf Left$(u, 9) = "ausgeschl" And WortEnde(u, 10) Then
   n = 9: seq = seq & "A"
  ElseIf Left$(u, 10) = "ausschluss" And WortEnde(u, 11) Then
   n = 10: seq = seq & "A"
  End If
  If n = 0 Then Exit Do
  t = Mid$(t, n + 1)
  t = LTrim$(t)
  If Len(t) > 0 Then
   k = InStr(":,-", Left$(t, 1))
   If k > 0 Then t = Mid$(t, 2)
  End If
 Loop
 PraefixAb = Trim$(t)
End Function

Private Function SkipICD(ByVal ic As String, ByVal obBrief As Integer) As Boolean
 If ic = "" Then SkipICD = True: Exit Function
 If ic Like "Z25*" Then SkipICD = True: Exit Function
 If obBrief <> 0 Then
  If ic Like "M20.*" Or ic Like "M21.*" Or ic Like "Q66.*" Or ic Like "B35.*" Or ic Like "K02.*" Then SkipICD = True: Exit Function
  If ic = "L84" Or ic = "R26.8" Or ic = "R29.6" Or ic = "R52.2" Or ic = "R68.8" Then SkipICD = True
 End If
End Function

Private Function AnzeigeText(ByVal ic As String, ByVal t As String, ByVal seite As String, ByVal attr As String, ByVal dmseit As String) As String
 ' Kopie der Umformungen aus MachDiagnosen (Importiere.bas)
 Dim m As String
 If Left$(ic, 2) = "E1" Then
  If InStr(t, "ankr") > 0 Or InStr(t, "ekund") > 0 Then
  Else
   Select Case Mid$(ic, 3, 1)
    Case "0": t = "Diabetes mellitus Typ 1"
    Case "1": t = "Diabetes mellitus Typ 2"
    Case "2": t = U("Diabetes mellitus in Verbindung mit Fehl- oder Mangelern{ae}hrung")
    Case "3": t = U("Diabetes mellitus (sekund{ae}r)")
    Case "4": t = "Diabetes mellitus"
   End Select
   If dmseit <> "" Then t = t & " seit " & dmseit
  End If
 End If
 If Left$(ic, 5) = "O24.4" Then t = "Gestationsdiabetes"
 If Left$(ic, 5) = "N08.3" And InStr(t, U("Glomerul{ae}re Krankheiten")) > 0 Then t = "Diabetische Nephropathie"
 If InStr(t, "Niereninsuff.") > 0 Then t = Replace(t, "Niereninsuff.", "Niereninsuffizienz")
 If Left$(ic, 5) = "G99.0" And InStr(t, "Autonome Neuropathie bei endokrinen") > 0 Then t = "Diabetische autonome Neuropathie"
 If (Left$(ic, 5) = "I79.9" Or Left$(ic, 5) = "I79.2") And InStr(t, "Periphere Angiopathie bei anderenorts") > 0 Then t = "Periphere Angiopathie"
 If Left$(ic, 5) = "M14.6" And InStr(t, "Neuropathische Arthropathie ") > 0 Then t = "Diabetische Osteoarthropathie"
 If Left$(ic, 5) = "M36.8" And InStr(t, "anderenorts") > 0 Then t = "Diabetische Bindegewebserkrankung"
 If Left$(ic, 3) = "E66" Then t = U("{Ue}bergewicht")
 m = Left$(ic, 5)
 If m = "K76.0" Or m = "K76.9" Or m = "K71.6" Or m = "K71.7" Or m = "K77.8" Then
  If InStr(t, "nbek") = 0 And InStr(t, "nklar") = 0 And InStr(attr, "nklar") = 0 Then t = "Hepatopathie"
 End If
 If Left$(ic, 3) = "L89" Then
  t = U("Diabetisches Fu{ss}syndrom ")
  If Mid$(ic, 5, 1) < "9" Then t = t & " im Stadium Wagner "
  Select Case Mid$(ic, 5, 1)
   Case "0", "1": t = t & "0"
   Case "2": t = t & "1"
   Case "3": t = t & "2"
   Case "4": t = t & "3"
  End Select
 End If
 Select Case seite
  Case "R": If Right$(t, 3) <> " re" Then t = t & " re"
  Case "L": If Right$(t, 3) <> " li" Then t = t & " li"
  Case "B": If InStr(t, "bds.") = 0 Then t = t & " bds."
 End Select
 If attr <> "" Then t = t & " (" & attr & ")"
 AnzeigeText = Replace(t, U(", nicht n{ae}her bezeichnet"), "")
End Function

Private Function Ersatz(ByVal t As String) As String
 ' wie in MachDiagnosen, wenn ein Label vorangestellt wird
 t = Replace(t, "Diab.", "diab.")
 t = Replace(t, "Diabeti", "diabeti")
 t = Replace(t, "Diabet.", "diabet.")
 t = Replace(t, "Auton", "auton")
 Ersatz = t
End Function

Private Function LabelFuer(ByVal gew As String, ByVal verg As Boolean, ByVal vz As Boolean) As String
 If gew = "A" Then
  LabelFuer = "Ausschluss "
 ElseIf vz Then
  LabelFuer = "V.a. Z.n. "
 ElseIf gew = "V" Then
  LabelFuer = IIf(verg, "Z.n. V.a. ", "V.a. ")
 Else
  LabelFuer = IIf(verg, "Z.n. ", "")
 End If
End Function

' ------------------------------------------------------------------------------------------
' Pflegetabellen laden
Private Sub RegelnLaden()
 Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
 If eGeladen Then Exit Sub
 eN = 0: pN = 0
 myFrag rs, "SELECT ICD_Praefix, COALESCE(Text_Muster,'') tm, COALESCE(Alter_Min,-1) am, Entscheidung, Prioritaet FROM diagentscheid ORDER BY ID"
 If Not rs.BOF Then
  Do While Not rs.EOF
   eN = eN + 1
   ReDim Preserve ePre(1 To eN): ReDim Preserve eTxt(1 To eN): ReDim Preserve eAlt(1 To eN): ReDim Preserve eDec(1 To eN): ReDim Preserve ePrio(1 To eN)
   ePre(eN) = rs!ICD_Praefix & ""
   eTxt(eN) = rs!tm & ""
   eAlt(eN) = Zahl(rs!am)
   eDec(eN) = rs!Entscheidung & ""
   ePrio(eN) = Zahl(rs!Prioritaet)
   rs.MoveNext
  Loop
 End If
 myFrag rs2, "SELECT ICD_Praefix, Bedingung, Code FROM diagpaare ORDER BY ICD_Praefix, Reihenfolge, ID"
 If Not rs2.BOF Then
  Do While Not rs2.EOF
   pN = pN + 1
   ReDim Preserve pPre(1 To pN): ReDim Preserve pCond(1 To pN): ReDim Preserve pCode(1 To pN)
   pPre(pN) = rs2!ICD_Praefix & ""
   pCond(pN) = rs2!Bedingung & ""
   pCode(pN) = rs2!Code & ""
   rs2.MoveNext
  Loop
 End If
 eGeladen = True
End Sub

Private Sub PatAlterLaden(ByVal Pat_ID As String)
 Dim rs As New ADODB.Recordset
 PatAlter = -1
 myFrag rs, "SELECT TIMESTAMPDIFF(YEAR, GebDat, CURDATE()) alt FROM namen WHERE pat_id = " & Pat_ID
 If Not rs.BOF Then
  If Not IsNull(rs!alt) Then PatAlter = Zahl(rs!alt)
 End If
End Sub

' ------------------------------------------------------------------------------------------
Private Sub LadeZeilen(ByVal Pat_ID As String, ByVal VorDat As Date, ByVal obBrief As Integer)
 Dim rs As New ADODB.Recordset, sql As String, info() As String, seq As String, ic As String, jj As Long, t1 As String
 sql = "SELECT d.DiagSicherheit, d.DiagText, d.DiagSeite, d.DiagAttr, d.ICD, d.obDauer, COALESCE(d.Dggel,0) Dggel, d.obKasse, d.lKasse, d.KFdFA, " & _
       "COALESCE(d.DiagDatum,0) DiagDatum, d.AusnBegr, d.intBemerk, COALESCE(d.MOStatus,0) MOStatus, COALESCE(g1.rf,0) rf, COALESCE(r.gi2,0) gi2, " & _
       "diag_info(d.DiagText, d.DiagSeite, d.DiagAttr) info, diag_jahr(d.DiagText) jahr " & vbCrLf & _
       "FROM diagnosen d LEFT JOIN diagreihe r ON d.icd = r.icd LEFT JOIN diagg1 g1 ON r.gi1 = g1.lfdnr " & vbCrLf & _
       "WHERE d.pat_id = " & Pat_ID & " AND (d.obDauer <> 0 OR d.DiagDatum > " & DatFor_k(VorDat) & ") " & vbCrLf & _
       "ORDER BY g1.rf, r.gi2 DESC, COALESCE(d.Dggel,0), d.ICD, d.ID1;"
 myFrag rs, sql
 zN = 0
 If rs.BOF Then Exit Sub
 Do While Not rs.EOF
  ic = rs!ICD & ""
  If Not SkipICD(ic, obBrief) Then
   zN = zN + 1
   ReDim Preserve Zl(1 To zN)
   With Zl(zN)
    .Icd = ic
    .IcdK = IIf(ic <> "-1" And Right$(ic, 1) = "-", Left$(ic, Len(ic) - 1), ic)
    .Roh = rs!DiagText & ""
    .Sich = rs!DiagSicherheit & ""
    .Seite = rs!DiagSeite & ""
    .Attr = rs!DiagAttr & ""
    .Status = Zahl(rs!MOStatus)
    .obDauer = IIf(Zahl(rs!obDauer) <> 0, -1, 0)
    If IsDate(rs!DiagDatum) Then .Datum = CDate(rs!DiagDatum) Else .Datum = 0
    If IsDate(rs!lKasse) Then .lKasse = CDate(rs!lKasse) Else .lKasse = 0
    .Dggel = Zahl(rs!Dggel)
    .obKasse = Zahl(rs!obKasse)
    .KFdFA = rs!KFdFA & ""
    .Aus = rs!AusnBegr & ""
    .iBm = rs!intBemerk & ""
    .G1 = Zahl(rs!rf)
    .G2 = Zahl(rs!gi2)
    info = Split(rs!info & "", Chr$(31))
    If UBound(info) < 8 Then ReDim Preserve info(0 To 8)
    .Daten = info(1): .B0 = info(2): .B1 = info(3): .Kern = info(4): .Rest = info(5): .Seiten = info(6): .ErlW = info(7): .TNorm = info(8)
    .W1 = Woerter(.B1)
    .Text = PraefixAb(.Roh, seq)
    jj = Zahl(rs!jahr)
    If jj = 0 And .Datum > 0 Then jj = Year(.Datum)
    .Jahr = jj
    .Gew = "G"
    If .Sich = "V" Or InStr(seq, "V") > 0 Then .Gew = "V"
    If .Sich = "A" Or InStr(seq, "A") > 0 Then .Gew = "A"
    .HatZ = (InStr(seq, "Z") > 0)
    .Verg = (.Sich = "Z" And .Status <> 4) Or .HatZ
    .VZ = (Left$(seq, 2) = "VZ")
    .Aktiv = (.Status <> 4)
    t1 = AnzeigeText(.Icd, .Text, " ", "", "")
    If Sp(t1) <> Sp(.Text) Then .RB = LCase$(Sp(t1))
   End With
  End If
  rs.MoveNext
 Loop
End Sub

' ------------------------------------------------------------------------------------------
' Abgleich: passt die (abgeschlossene / Z.n.) Zeile c zur (aktiven / gesicherten) Zeile a?
Private Function DatenGedeckt(ByVal c As Long, ByVal a As Long) As Boolean
 ' jede Datumsangabe im Text von c steht im Text von a oder ihre Zahlengruppen alle in der Erlaeuterung von a
 Dim v() As String, g() As String, i As Long, j As Long, ok As Boolean
 DatenGedeckt = True
 If Len(Trim$(Zl(c).Daten)) = 0 Then Exit Function
 v = Split(Trim$(Zl(c).Daten), " ")
 For i = 0 To UBound(v)
  If Len(v(i)) > 0 Then
   If InStr(1, " " & Zl(a).Daten & " ", " " & v(i) & " ", vbBinaryCompare) = 0 Then
    g = Split(Trim$(Woerter(v(i))), " ")
    ok = True
    For j = 0 To UBound(g)
     If Len(g(j)) > 0 Then
      If InStr(1, " " & Zl(a).ErlW & " ", " " & g(j) & " ", vbBinaryCompare) = 0 Then ok = False: Exit For
     End If
    Next j
    If Not ok Then DatenGedeckt = False: Exit Function
   End If
  End If
 Next i
End Function

Private Function Passt(ByVal c As Long, ByVal a As Long, Optional ByVal locker As Boolean = False) As Boolean
 ' locker: auch andere ICD derselben 3-Stelle, sofern die Anzeige den Text nicht aus der ICD erzeugt (RB leer)
 If Not (Zl(c).IcdK = Zl(a).IcdK Or Zl(c).Icd = "-1" Or (locker And Left$(Zl(c).Icd, 3) = Left$(Zl(a).Icd, 3) And Len(Zl(c).RB) = 0 And Len(Zl(a).RB) = 0)) Then Exit Function
 If Not DatenGedeckt(c, a) Then Exit Function
 If Not TeilmengeS(Zl(c).Seiten, Zl(a).Seiten) Then Exit Function
 If Not TeilmengeW(Zl(c).ErlW, Zl(a).ErlW) Then Exit Function
 If Zl(c).B0 = Zl(a).B0 Or Zl(c).B1 = Zl(a).B1 Then
  Passt = True
 ElseIf Len(Zl(c).RB) > 0 And Zl(c).RB = Zl(a).RB Then
  Passt = True
 ElseIf Len(Zl(c).Kern) > 0 And Zl(c).Kern = Zl(a).Kern And TeilmengeW(Zl(c).Rest, Zl(a).Rest & " " & Zl(a).ErlW) Then
  Passt = True
 ElseIf Len(Zl(c).W1) > 0 And TeilmengeW(Zl(c).W1, Zl(a).W1 & " " & Zl(a).ErlW) Then
  Passt = True
 End If
End Function

Private Function KeyKopie(ByVal i As Long, ByVal mitVerg As Boolean) As String
 Dim b As String
 b = Zl(i).RB
 If Len(b) = 0 Then b = Zl(i).B1
 KeyKopie = Zl(i).IcdK & "|" & b & "|" & Zl(i).Seiten & "|" & Zl(i).Daten & "|" & Zl(i).ErlW & "|" & Zl(i).Gew & IIf(mitVerg, "|" & IIf(Zl(i).Verg, "1", "0"), "")
End Function

Private Function Entscheid(ByVal ic As String, ByVal tnorm As String) As String
 ' Entscheidung aus diagentscheid: laengster ICD-Praefix, dann hoechste Prioritaet; Standard je
 Dim i As Long, bestKey As Long, k As Long
 Entscheid = "je"
 bestKey = -1
 For i = 1 To eN
  If Left$(ic, Len(ePre(i))) = ePre(i) Then
   If Len(eTxt(i)) = 0 Or PatLike(eTxt(i), tnorm) Then
    If eAlt(i) < 0 Or (PatAlter >= 0 And PatAlter > eAlt(i)) Then
     k = Len(ePre(i)) * 1000 + ePrio(i)
     If k > bestKey Then
      bestKey = k
      Entscheid = eDec(i)
     End If
    End If
   End If
  End If
 Next i
End Function

Private Function BedingungOK(ByVal c As String, ByVal iz As Long, ByVal ig As Long) As Boolean
 Dim zt As String, gt As String, zAussen As Boolean, zInnen As Boolean, gAussen As Boolean, gInnen As Boolean
 If c = "" Then
  BedingungOK = True
 ElseIf c = "GLEICH" Then
  BedingungOK = Passt(iz, ig) Or Passt(ig, iz)
 ElseIf c = "GJUENGER" Then
  BedingungOK = (Zl(ig).Jahr > Zl(iz).Jahr)
 ElseIf c = "SEITE" Or c = "SEITE_OHNE_AUSSENINNEN" Then
  If Not SeitenGemeinsam(Zl(iz).Seiten, Zl(ig).Seiten) Then Exit Function
  If c = "SEITE" Then
   BedingungOK = True
  Else
   zt = LCase$(Zl(iz).Text): gt = LCase$(Zl(ig).Text)
   zAussen = (InStr(zt, "au" & ChrW$(223) & "en") > 0 Or InStr(zt, "aussen") > 0)
   gAussen = (InStr(gt, "au" & ChrW$(223) & "en") > 0 Or InStr(gt, "aussen") > 0)
   zInnen = (InStr(zt, "innen") > 0): gInnen = (InStr(gt, "innen") > 0)
   BedingungOK = Not ((zAussen And gInnen) Or (gAussen And zInnen))
  End If
 ElseIf Left$(c, 6) = "GTEXT:" Then
  BedingungOK = PatLike(Mid$(c, 7), Zl(ig).TNorm)
 ElseIf Left$(c, 6) = "ZTEXT:" Then
  BedingungOK = PatLike(Mid$(c, 7), Zl(iz).TNorm)
 End If
End Function

Private Function PaarCode(ByVal iz As Long, ByVal ig As Long) As String
 ' Regeln aus diagpaare: laengster Praefix, der auf beide ICD passt; darin die erste Regel (nach Reihenfolge), deren Bedingung zutrifft
 Dim i As Long, ml As Long
 ml = -1
 For i = 1 To pN
  If Left$(Zl(iz).Icd, Len(pPre(i))) = pPre(i) And Left$(Zl(ig).Icd, Len(pPre(i))) = pPre(i) Then
   If Len(pPre(i)) > ml Then ml = Len(pPre(i))
  End If
 Next i
 PaarCode = "b"
 If ml < 0 Then Exit Function
 For i = 1 To pN
  If Len(pPre(i)) = ml Then
   If Left$(Zl(iz).Icd, ml) = pPre(i) And Left$(Zl(ig).Icd, ml) = pPre(i) Then
    If BedingungOK(pCond(i), iz, ig) Then
     PaarCode = pCode(i)
     Exit Function
    End If
   End If
  End If
 Next i
End Function

' ------------------------------------------------------------------------------------------
Private Sub Auswahl(ByVal dmseit As String)
 Dim i As Long, j As Long, g As Long, grpN As Long, col As Collection, bestOf() As Long, key As String
 Dim zi() As Long, zAnz As Long, gi() As Long, gAnz As Long, code As String, yz As Long, yg As Long
 Dim ci() As Long, cn As Long, ers() As Boolean, ok As Boolean, d As String, best As Long, first As Long, cand() As Boolean
 Dim hid() As Boolean, dd() As String, zc() As Long, zcn As Long, zers() As Boolean
 Dim deckung As String, hatG0 As Boolean

 ' --- 1. aktive Zeilen: identische Kopien zusammenfassen
 ReDim bestOf(1 To zN)
 Set col = New Collection
 grpN = 0
 For i = 1 To zN
  If Zl(i).Aktiv Then
   key = KeyKopie(i, True)
   g = FindeKey(col, key)
   If g = 0 Then
    grpN = grpN + 1: bestOf(grpN) = i: col.Add grpN, key
   ElseIf Neuer(i, bestOf(g)) Then
    Zl(bestOf(g)).Weg = True: bestOf(g) = i
   Else
    Zl(i).Weg = True
   End If
  End If
 Next i
 ' --- 1b. aktive Zeilen ausser Z.n. und Ausschluss: wie bisher je ICD-3, Gewissheit und ersten 17 Zeichen
 ReDim bestOf(1 To zN)
 Set col = New Collection
 grpN = 0
 For i = 1 To zN
  If Zl(i).Aktiv And Not Zl(i).Weg And Not Zl(i).Verg And Zl(i).Gew <> "A" Then
   key = Left$(Zl(i).Icd, 3) & "|" & Zl(i).Gew & "|" & LCase$(Left$(Zl(i).Roh, 17))
   g = FindeKey(col, key)
   If g = 0 Then
    grpN = grpN + 1: bestOf(grpN) = i: col.Add grpN, key
   ElseIf Neuer(i, bestOf(g)) Then
    Zl(bestOf(g)).Weg = True: bestOf(g) = i
   Else
    Zl(i).Weg = True
   End If
  End If
 Next i
 ' --- 1c. aktive Z.n.: weniger spezifische durch spezifischere aktive Z.n. derselben Diagnose verdraengen
 ReDim zc(1 To zN)
 zcn = 0
 For i = 1 To zN
  If Zl(i).Aktiv And Not Zl(i).Weg And Zl(i).Verg And Zl(i).Gew <> "A" Then zcn = zcn + 1: zc(zcn) = i
 Next i
 ReDim zers(1 To zN)
 For i = 1 To zcn
  For j = 1 To zcn
   If j <> i Then
    If Zl(zc(i)).Gew = Zl(zc(j)).Gew Then
     If Passt(zc(i), zc(j)) Then
      If Not Passt(zc(j), zc(i)) Then
       zers(zc(i)) = True
      ElseIf NeuerSpaeter(zc(j), zc(i)) Then
       zers(zc(i)) = True
      End If
     End If
    End If
   End If
  Next j
 Next i
 ' aktive gesicherte Zeilen: ebenso, auch ueber verschiedene ICD derselben 3-Stelle
 zcn = 0
 For i = 1 To zN
  If Zl(i).Aktiv And Not Zl(i).Weg And Not Zl(i).Verg And Zl(i).Gew = "G" Then zcn = zcn + 1: zc(zcn) = i
 Next i
 For i = 1 To zcn
  For j = 1 To zcn
   If j <> i Then
    If Passt(zc(i), zc(j), True) Then
     If Not Passt(zc(j), zc(i), True) Then
      zers(zc(i)) = True
     ElseIf NeuerSpaeter(zc(j), zc(i)) Then
      zers(zc(i)) = True
     End If
    End If
   End If
  Next j
 Next i
 For i = 1 To zN
  If zers(i) Then Zl(i).Weg = True
 Next i
 ' Label der aktiven Zeilen
 For i = 1 To zN
  If Zl(i).Aktiv And Not Zl(i).Weg Then
   Zl(i).Label = LabelFuer(Zl(i).Gew, Zl(i).Verg, Zl(i).VZ)
   Zl(i).Gewaehlt = True
  End If
 Next i

 ' --- 2. abgeschlossene Zeilen
 ReDim bestOf(1 To zN)
 Set col = New Collection
 grpN = 0
 For i = 1 To zN
  If Not Zl(i).Aktiv Then
   key = KeyKopie(i, False)
   g = FindeKey(col, key)
   If g = 0 Then
    grpN = grpN + 1: bestOf(grpN) = i: col.Add grpN, key
   ElseIf Neuer(i, bestOf(g)) Then
    Zl(bestOf(g)).Weg = True: bestOf(g) = i
   Else
    Zl(i).Weg = True
   End If
  End If
 Next i
 ' Kandidaten: ohne aktives Gegenstueck
 ReDim ci(1 To zN)
 cn = 0
 For i = 1 To zN
  If Not Zl(i).Aktiv And Not Zl(i).Weg Then
   ok = True
   For j = 1 To zN
    If Zl(j).Aktiv Then
     If Passt(i, j) Then ok = False: Exit For
    End If
   Next j
   If ok Then cn = cn + 1: ci(cn) = i
  End If
 Next i
 ' Kandidaten, die von einer spezifischeren abgeschlossenen Zeile verdraengt werden
 ReDim ers(1 To zN)
 For i = 1 To cn
  For j = 1 To cn
   If j <> i Then
    If Passt(ci(i), ci(j)) Then
     If Not Passt(ci(j), ci(i)) Then
      ers(ci(i)) = True
     ElseIf NeuerSpaeter(ci(j), ci(i)) Then
      ers(ci(i)) = True
     End If
    End If
   End If
  Next j
 Next i
 ' Entscheidung je Kandidat
 ReDim dd(1 To zN)
 For i = 1 To cn
  If Not ers(ci(i)) Then
   d = Entscheid(Zl(ci(i)).Icd, Zl(ci(i)).TNorm)
   dd(ci(i)) = d
   If d = "je" Then
    Zl(ci(i)).Label = LabelFuer(Zl(ci(i)).Gew, True, False)
    Zl(ci(i)).Gewaehlt = True
   End If
  End If
 Next i
 ' j / jz: einmal je ICD-3 und Gewissheit (die juengste, bei Gleichstand die fruehere)
 For i = 1 To cn
  If Not ers(ci(i)) And (dd(ci(i)) = "j" Or dd(ci(i)) = "jz") Then
   best = 0
   For j = 1 To cn
    If Not ers(ci(j)) And (dd(ci(j)) = "j" Or dd(ci(j)) = "jz") Then
     If Left$(Zl(ci(j)).Icd, 3) = Left$(Zl(ci(i)).Icd, 3) And Zl(ci(j)).Gew = Zl(ci(i)).Gew Then
      If best = 0 Then
       best = ci(j)
      ElseIf Neuer(ci(j), best) Then
       best = ci(j)
      End If
     End If
    End If
   Next j
   If best = ci(i) Then
    Zl(ci(i)).Label = LabelFuer(Zl(ci(i)).Gew, (dd(ci(i)) = "jz" Or Zl(ci(i)).HatZ), False)
    Zl(ci(i)).Gewaehlt = True
   End If
  End If
 Next i

 ' --- 3. Paarregeln (diagpaare) ueber alle angezeigten Z.n.- und gesichert-Zeilen (aktive und abgeschlossene)
 ReDim zi(1 To zN): ReDim gi(1 To zN)
 zAnz = 0: gAnz = 0
 For i = 1 To zN
  If Zl(i).Gewaehlt And Zl(i).Gew = "G" Then
   If Zl(i).Label = "Z.n. " Then
    zAnz = zAnz + 1: zi(zAnz) = i
   ElseIf Zl(i).Label = "" Then
    gAnz = gAnz + 1: gi(gAnz) = i
   End If
  End If
 Next i
 ReDim hid(1 To zN)
 For i = 1 To zAnz
  For j = 1 To gAnz
   If Left$(Zl(zi(i)).Icd, 3) = Left$(Zl(gi(j)).Icd, 3) And Zl(zi(i)).Icd <> "-1" Then
    Zl(zi(i)).Gepaart = True: Zl(gi(j)).Gepaart = True
    code = PaarCode(zi(i), gi(j))
    yz = Zl(zi(i)).Jahr: yg = Zl(gi(j)).Jahr
    If code = "g" Then
     hid(zi(i)) = True
    ElseIf code = "z" Then
     hid(gi(j)) = True
    ElseIf code = "j" Then
     If yg >= yz Then hid(zi(i)) = True Else hid(gi(j)) = True
    ElseIf code = "b" Then
     Zl(zi(i)).FlagB = True: Zl(gi(j)).FlagB = True
    ElseIf code = "a" Then
     Zl(zi(i)).FlagA = True: Zl(gi(j)).FlagA = True
    ElseIf IsNumeric(code) Then
     If yg - yz >= CLng(code) Then
      Zl(zi(i)).FlagA = True: Zl(gi(j)).FlagA = True
     ElseIf yg > yz Then
      hid(zi(i)) = True
     Else
      hid(gi(j)) = True
     End If
    End If
   End If
  Next j
 Next i
 For i = 1 To zN
  If hid(i) Then Zl(i).Gewaehlt = False
 Next i
 ' Code b: von den nur mit b verbundenen Zeilen bleibt je Art (Z.n. / gesichert) und ICD-3 die juengste
 For i = 1 To zAnz
  If Zl(zi(i)).FlagB And Not Zl(zi(i)).FlagA And Zl(zi(i)).Gewaehlt Then
   For j = 1 To zAnz
    If j <> i Then
     If Zl(zi(j)).FlagB And Not Zl(zi(j)).FlagA And Not hid(zi(j)) And Left$(Zl(zi(j)).Icd, 3) = Left$(Zl(zi(i)).Icd, 3) Then
      If Neuer(zi(j), zi(i)) Or (Not Neuer(zi(i), zi(j)) And j < i) Then Zl(zi(i)).Gewaehlt = False
     End If
    End If
   Next j
  End If
 Next i
 For i = 1 To gAnz
  If Zl(gi(i)).FlagB And Not Zl(gi(i)).FlagA And Zl(gi(i)).Gewaehlt Then
   For j = 1 To gAnz
    If j <> i Then
     If Zl(gi(j)).FlagB And Not Zl(gi(j)).FlagA And Not hid(gi(j)) And Left$(Zl(gi(j)).Icd, 3) = Left$(Zl(gi(i)).Icd, 3) Then
      If Neuer(gi(j), gi(i)) Or (Not Neuer(gi(i), gi(j)) And j < i) Then Zl(gi(i)).Gewaehlt = False
     End If
    End If
   Next j
  End If
 Next i
 ' aktive Z.n. ohne gesicherten Partner: Tabelle diagentscheid (je = alle einzeln, sonst einmal je ICD-3)
 ReDim cand(1 To zN)
 For i = 1 To zAnz
  If Zl(zi(i)).Aktiv And Not Zl(zi(i)).Gepaart And Zl(zi(i)).Gewaehlt Then cand(zi(i)) = True
 Next i
 For i = 1 To zAnz
  If cand(zi(i)) Then
   first = 0: best = 0
   For j = 1 To zAnz
    If cand(zi(j)) Then
     If Left$(Zl(zi(j)).Icd, 3) = Left$(Zl(zi(i)).Icd, 3) Then
      If first = 0 Then first = zi(j)
      If best = 0 Then
       best = zi(j)
      ElseIf Neuer(zi(j), best) Then
       best = zi(j)
      End If
     End If
    End If
   Next j
   d = Entscheid(Zl(first).Icd, Zl(first).TNorm)
   If d <> "je" And best <> zi(i) Then Zl(zi(i)).Gewaehlt = False
  End If
 Next i
 ' L89: aktive gesicherte Wagner-0-Zeile deckt Z.n.-Wagner-0-Zeilen ab, die keine andere Seite betreffen
 deckung = "": hatG0 = False
 For i = 1 To zN
  If Zl(i).Gewaehlt And Zl(i).Label = "" And Left$(Zl(i).Icd, 3) = "L89" And (Mid$(Zl(i).Icd, 5, 1) = "0" Or Mid$(Zl(i).Icd, 5, 1) = "1") Then
   hatG0 = True
   For j = 1 To Len(Zl(i).Seiten)
    If InStr(deckung, Mid$(Zl(i).Seiten, j, 1)) = 0 Then deckung = deckung & Mid$(Zl(i).Seiten, j, 1)
   Next j
  End If
 Next i
 If hatG0 Then
  For i = 1 To zN
   If Zl(i).Gewaehlt And Zl(i).Label = "Z.n. " And Left$(Zl(i).Icd, 3) = "L89" And (Mid$(Zl(i).Icd, 5, 1) = "0" Or Mid$(Zl(i).Icd, 5, 1) = "1") Then
    If TeilmengeS(Zl(i).Seiten, deckung) Then Zl(i).Gewaehlt = False
   End If
  Next i
 End If
End Sub

' ------------------------------------------------------------------------------------------
Private Sub AusgabeFuellen(ByVal dmseit As String)
 Dim i As Long, k As Long, n As Long, vis() As Long, col As Collection, key As String, zeile As String, ind As String, t As String
 ReDim vis(1 To zN)
 Set col = New Collection
 n = 0
 For i = 1 To zN
  If Zl(i).Gewaehlt Then
   t = AnzeigeText(Zl(i).Icd, Zl(i).Text, Zl(i).Seite, Zl(i).Attr, dmseit)
   If Len(Zl(i).Label) > 0 Then t = Ersatz(t)
   zeile = IIf(Zl(i).obDauer = 0, "(q) ", "") & Zl(i).Label & t
   ind = IIf(Zl(i).G1 = 1 And Zl(i).G2 > 50 And Zl(i).G2 < 97, " ", "")
   key = Left$(Zl(i).Icd, 3) & "|" & ind & zeile
   If FindeKey(col, key) = 0 Then
    col.Add i, key
    n = n + 1: vis(n) = i
   End If
  End If
 Next i
 DiagNrSetzen n ' dimensioniert in Importiere.bas alle Diagnosen-Arrays neu (Namen wie G1 sind auch in AnAnpassen.bas vergeben)
 For k = 1 To n
  i = vis(k)
  With Zl(i)
   If Len(.Label) > 0 Then
    Importiere.Diag(k - 1) = .Label & Ersatz(.Text)
   Else
    Importiere.Diag(k - 1) = .Text
   End If
   If Left$(.Label, 4) = "Z.n." Then
    Importiere.DSic(k - 1) = "Z"
   ElseIf Left$(.Label, 4) = "V.a." Then
    Importiere.DSic(k - 1) = "V"
   ElseIf Left$(.Label, 10) = "Ausschluss" Then
    Importiere.DSic(k - 1) = "A"
   Else
    Importiere.DSic(k - 1) = "G"
   End If
   Importiere.ICD(k - 1) = .Icd
   Importiere.DiagSe(k - 1) = .Seite
   Importiere.DiagAttr(k - 1) = .Attr
   Importiere.DiagAus(k - 1) = .Aus
   Importiere.DiagiBm(k - 1) = .iBm
   Importiere.diagdt(k - 1) = .Datum
   Importiere.obDauer(k - 1) = .obDauer
   Importiere.obKasse(k - 1) = .obKasse
   Importiere.lKasse(k - 1) = .lKasse
   Importiere.Dggel(k - 1) = .Dggel
   Importiere.KFdFA(k - 1) = .KFdFA
   Importiere.G1(k - 1) = CInt(.G1)
   Importiere.G2(k - 1) = CInt(.G2)
  End With
 Next k
End Sub

' ------------------------------------------------------------------------------------------
' Test im Direktfenster:  DiagTest "155"   (gibt den alten und den neuen Diagnosestring im Direktfenster aus)
Public Sub DiagTest(ByVal Pat_ID As String)
 Dim dt() As CString, s As String
 Call Lese.ProgStart
 DiagNeuAbschalten = True
 s = DiagString(Pat_ID, dt, , True, "")
 Debug.Print "--- ALT (alte Logik) ---"
 Debug.Print Replace(s, vbVerticalTab, vbCrLf)
 DiagNeuAbschalten = False
 s = DiagString(Pat_ID, dt, , True, "")
 Debug.Print "--- NEU (DiagAnzeige) ---"
 Debug.Print Replace(s, vbVerticalTab, vbCrLf)
End Sub
