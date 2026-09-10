Attribute VB_Name = "Laborregel"
Option Explicit

' Laborregel-Engine, Phase 1 (Kern): ermittelt anhand der Tabelle "laborregel"
' (Diagnosen, letzter eigener Laborwert, BMI, Alter, Blutdruck, Versicherung,
' DMP-Status, Medikation) die fuer einen Patienten aktuell indizierten
' Laborwerte fuer die naechste Blutabnahme. Siehe Plan "Laborregel-Engine -
' Phase 1" vom 10.9.2026. Noch OHNE: "jemals"-Werte, anderer-Laborparameter-
' Bedingungen, pmg-Override, Knopf/Unterfenster im Laufzettel (spaetere Phasen).

Public Type RRAvg
 Syst As Double
 Diast As Double
 Zahl As Long
End Type

' Vergleichsterm wie ">24" oder "<18" gegen einen numerischen Wert pruefen.
' Nutzt MachNumerisch (AnAnpassen.bas) zum Parsen der Zahl, damit Einheiten/Text
' im Vergleichsterm ignoriert werden (z.B. "<50ml/min"). Leerer Vergleichsterm
' gilt als erfuellt (keine Bedingung).
Function ErfuelltVgl(Wert As Double, Vgl As String) As Boolean
 Dim v$, op$, schwelle#
 v = Trim$(Vgl)
 If Len(v) = 0 Then ErfuelltVgl = True: Exit Function
 If left$(v, 2) = ">=" Or left$(v, 2) = "<=" Then
  op = left$(v, 2)
 ElseIf left$(v, 1) = ">" Or left$(v, 1) = "<" Or left$(v, 1) = "=" Then
  op = left$(v, 1)
 Else
  op = "="
 End If
 schwelle = MachNumerisch(v)
 Select Case op
  Case ">": ErfuelltVgl = (Wert > schwelle)
  Case "<": ErfuelltVgl = (Wert < schwelle)
  Case ">=": ErfuelltVgl = (Wert >= schwelle)
  Case "<=": ErfuelltVgl = (Wert <= schwelle)
  Case Else: ErfuelltVgl = (Wert = schwelle)
 End Select
End Function ' ErfuelltVgl

' Hat der Patient eine (Verdachts-)Diagnose, deren ICD-Code zu Regex passt?
' Dauerdiagnosen zaehlen immer, andere (inkl. Verdacht) nur, wenn ihr
' DiagDatum nicht aelter als MonateNichtDauer Monate ist. Leeres Regex = immer erfuellt.
Function HatICD(pid As Long, Regex As String, Optional MonateNichtDauer As Integer = 12) As Boolean
 If Len(Trim$(Regex)) = 0 Then HatICD = True: Exit Function
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT 1 FROM diagnosen WHERE Pat_id = " & pid & " AND ICD RLIKE '" & Regex & "'" & _
   " AND DiagSicherheit IN ('G',' ','V','Z')" & _
   " AND (obDauer<>0 OR DiagDatum >= DATE_SUB(NOW(), INTERVAL " & MonateNichtDauer & " MONTH)) LIMIT 1"
 HatICD = Not r.BOF
 Set r = Nothing
End Function ' HatICD

' Hat der Patient im aktuellsten Medikamentenplan ein Medikament/einen Wirkstoff,
' der zu Regex passt? Leeres Regex = immer erfuellt.
Function HatMedikament(pid As Long, Regex As String) As Boolean
 If Len(Trim$(Regex)) = 0 Then HatMedikament = True: Exit Function
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT 1 FROM medplan mp WHERE mp.Pat_id = " & pid & _
   " AND mp.Zeitpunkt = (SELECT MAX(Zeitpunkt) FROM medplan WHERE Pat_id = mp.Pat_id)" & _
   " AND (mp.Medikament RLIKE '" & Regex & "' OR mp.Wirkstoff RLIKE '" & Regex & "') LIMIT 1"
 HatMedikament = Not r.BOF
 Set r = Nothing
End Function ' HatMedikament

' Letzter eigener Wert eines Laborparameters: Regex gegen Abkü in labor1a/labor2a
' (z.B. "^(CHOL|XCHOL)$"), da eine Bestellbezeichnung (laborregel.Laborparameter)
' nicht zwingend mit der/den tatsächlich verwendeten Abkürzung(en) übereinstimmt.
' Null, wenn kein Wert vorhanden. Liefert optional zusätzlich den Zeitpunkt dieses
' Werts zurück (für IntervallMonate-Prüfung).
Function LetzterEigenerWert(pid As Long, Regex As String, Optional ByRef LetzterZeitpunkt As Variant) As Variant
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT Wert,Zeitpunkt FROM (" & _
   "SELECT Wert,Zeitpunkt FROM labor1a WHERE Pat_id=" & pid & " AND Abkü RLIKE '" & Regex & "'" & _
   " UNION ALL SELECT Wert,Zeitpunkt FROM labor2a WHERE Pat_id=" & pid & " AND Abkü RLIKE '" & Regex & "'" & _
   ") x ORDER BY Zeitpunkt DESC LIMIT 1"
 If Not r.BOF Then
  LetzterEigenerWert = MachNumerisch(REPLACE$(nz(r!Wert, ""), ",", "."))
  LetzterZeitpunkt = r!Zeitpunkt
 Else
  LetzterEigenerWert = Null
  LetzterZeitpunkt = Null
 End If
End Function ' LetzterEigenerWert

' Blutdruck-Durchschnitt (nach RRzahl gewichtet) der letzten "Tage" Tage.
' Extrahiert aus Laufzettelneu.bas ("TH:Blutdruck") / LaufzMo.bas (Duplikat),
' die dort jetzt diese Funktion aufrufen statt die Query zu duplizieren.
Function RRDurchschnitt(pid As Long, Optional Tage As Long = 183) As RRAvg
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT SUM(rrsyst*rrzahl)/SUM(rrzahl) Rs, SUM(rrdiast*rrzahl)/SUM(rrzahl) Rd, SUM(rrzahl) Rz FROM rr WHERE pat_id=" & pid & " AND zeitpunkt> SUBDATE(NOW()," & Tage & ")"
 If Not IsNull(r!Rs) Then RRDurchschnitt.Syst = r!Rs
 If Not IsNull(r!Rd) Then RRDurchschnitt.Diast = r!Rd
 If Not IsNull(r!Rz) Then RRDurchschnitt.Zahl = r!Rz
 Set r = Nothing
End Function ' RRDurchschnitt

' Versicherungsstatus ("privat"/"kasse") anhand des juengsten Falls, analog
' zur bestehenden Logik in Laufzettelneu.bas (rFlSchGr = 90/89 => privat).
Function ErmittleVersicherung(pid As Long) As String
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT schgr FROM faelle f WHERE f.Pat_id = " & pid & " ORDER BY bhfb DESC LIMIT 1"
 If Not r.BOF Then
  If r!schgr = 90 Or r!schgr = 89 Then
   ErmittleVersicherung = "privat"
  Else
   ErmittleVersicherung = "kasse"
  End If
 Else
  ErmittleVersicherung = "kasse" ' kein Fall gefunden: konservative Annahme
 End If
 Set r = Nothing
End Function ' ErmittleVersicherung

' DMP-Einschreibestatus: namen.dmpklass = "DMP hier" (3) oder "DMP HA" (2)
' zaehlt als eingeschrieben (DMPEnum siehe ZielDBFunktionen.bas).
Function ErmittleDMP(pid As Long) As Boolean
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT dmpklass FROM namen WHERE Pat_ID = " & pid & " LIMIT 1"
 If Not r.BOF Then
  ErmittleDMP = (r!dmpklass = DMPEnum.HA Or r!dmpklass = DMPEnum.hier)
 End If
 Set r = Nothing
End Function ' ErmittleDMP

' Alter in Jahren, ueber die MySQL-Funktion patAlter() (auch anderswo im
' Projekt so verwendet, z.B. ZielDBFunktionen.bas).
Function ErmittleAlter(pid As Long) As Double
 ErmittleAlter = nz(myEFrag("SELECT patAlter(" & pid & ") A")!A, 0)
End Function ' ErmittleAlter

' BMI ueber die bestehende Funktion DMPString/DMPClass (ZielDBFunktionen.bas),
' damit die Berechnung (Groesse/Gewicht-Ermittlung) nicht dupliziert wird.
Function ErmittleBMI(pid As Long) As Double
 Dim dmpLokal As DMPClass
 Call DMPString(pid, dmpLokal, True)
 ErmittleBMI = dmpLokal.bmi
End Function ' ErmittleBMI

' Haupteinstiegspunkt: liefert die Liste der fuer den Patienten aktuell
' indizierten Laborparameter (Werte aus laborregel.Laborparameter, je einmal).
' Pro Regelzeile werden nur befuellte Felder geprueft und UND-verknuepft;
' mehrere Zeilen fuer denselben Laborparameter sind ODER-verknuepft.
Function ErmittleIndizierteLaborwerte(pid As Long) As Collection
 Dim Erg As New Collection
 Dim r As New ADODB.Recordset
 myFrag r, "SELECT * FROM laborregel WHERE GueltigAb <= NOW() AND (GueltigBis IS NULL OR GueltigBis >= NOW())"

 Dim BMI#, Alter#, RR As RRAvg, Versicherung$, obDMP%
 BMI = ErmittleBMI(pid)
 Alter = ErmittleAlter(pid)
 RR = RRDurchschnitt(pid)
 Versicherung = ErmittleVersicherung(pid)
 obDMP = ErmittleDMP(pid)

 Do While Not r.EOF
  Dim obErfuellt%
  obErfuellt = True

  If obErfuellt Then If LenB(nz(r!ICDRegex, "")) <> 0 Then obErfuellt = HatICD(pid, r!ICDRegex)

  If obErfuellt Then
   If LenB(nz(r!EigenerWertVgl, "")) <> 0 Or Not IsNull(r!IntervallMonate) Then
    Dim ewert, ewertZp
    Dim abkRegex$
    abkRegex = nz(r!AbkueRegex, "")
    If LenB(abkRegex) = 0 Then abkRegex = r!Laborparameter
    ewert = LetzterEigenerWert(pid, abkRegex, ewertZp)
    If LenB(nz(r!EigenerWertVgl, "")) <> 0 Then
     If IsNull(ewert) Then obErfuellt = False Else obErfuellt = ErfuelltVgl(CDbl(ewert), r!EigenerWertVgl)
    End If
    If obErfuellt And Not IsNull(r!IntervallMonate) Then
     If r!IntervallMonate = 0 Then
      obErfuellt = IsNull(ewertZp) ' 0 = einmalig, nur wenn noch nie gemessen
     Else
      ' Tage statt DateAdd("m",...) verwenden, da DateAdd Bruchteile von Monaten nicht praezise verarbeitet;
      ' 1 Monat wird dabei als 30 Tage gerechnet (siehe Spaltenkommentar IntervallMonate)
      obErfuellt = IsNull(ewertZp) Or (ewertZp < Now() - CDbl(r!IntervallMonate) * 30)
     End If
    End If
   End If
  End If

  If obErfuellt Then If LenB(nz(r!BMIVgl, "")) <> 0 Then obErfuellt = ErfuelltVgl(BMI, r!BMIVgl)

  If obErfuellt Then If LenB(nz(r!AlterVgl, "")) <> 0 Then obErfuellt = ErfuelltVgl(Alter, r!AlterVgl)

  If obErfuellt Then If LenB(nz(r!RRVgl, "")) <> 0 Then obErfuellt = ErfuelltVgl(RR.Syst, r!RRVgl)

  If obErfuellt Then If LenB(nz(r!Versicherung, "")) <> 0 Then obErfuellt = (LCase$(r!Versicherung) = Versicherung)

  If obErfuellt Then If LenB(nz(r!DMP, "")) <> 0 Then obErfuellt = ((LCase$(r!DMP) = "ja") = (obDMP <> 0))

  If obErfuellt Then If LenB(nz(r!MedikamentRegex, "")) <> 0 Then obErfuellt = HatMedikament(pid, r!MedikamentRegex)

  If obErfuellt Then
   Dim schonDrin%, i%
   schonDrin = False
   For i = 1 To Erg.Count
    If Erg(i) = r!Laborparameter Then schonDrin = True: Exit For
   Next i
   If Not schonDrin Then Erg.Add r!Laborparameter
  End If

  r.MoveNext
 Loop
 Set r = Nothing
 Set ErmittleIndizierteLaborwerte = Erg
End Function ' ErmittleIndizierteLaborwerte
