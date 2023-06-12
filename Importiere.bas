Attribute VB_Name = "Importiere"
#Const CString = 0
Option Explicit
Public ql$
Public AktID& ' zum Merken des aktuellen Patienten beim automatischen Formularschluß
Public Const Anmnb$ = "AnBog" ' "Anamnesebogen"
Public Anmnbi&, ABPat_ID&
'Public Const ExtraPatienten$ = "'zutun','bestellungen','seminar','cgms gerät','fehler','teambesprechung','ideen','einladungen','neues'"
Public Const ExtraPatienten$ = "'^zutun|^bestellungen|^seminar|^cgms ger|^fehler|^teambesprechung|^ideen|^einladungen|^neues|^zukaufen'"
Public kvä As New ADODB.Recordset
Public obÜP% ' ob übertragungsprotokoll geführt wird
Dim rLPbm
Public altRecordSource$
'Const obMy% = True ' ob ADO-Quelle eine MySQL-Datenbank ist
Const obÜProt% = 0 ' ob Übertragungsprotokoll
Public obHausBesuch% ' ob über Hausbesuchsmodul gespeichert wurde (nur das letzte Jahr jedes Patienten)
Public pCol As Collection ' Kollektion zu aktualisierender Patienten
Dim numA&
'Dim rs As Adodb.Recordset
'Dim rEzä() As Adodb.Recordset
'Public rsAnam As Adodb.Recordset
Dim obvorgestellt As Boolean, messDatum As Date, messDatumD As Date
Dim DText_$, ICD_$, DSic_$, DSe_$, DAt_$, DAus_$, DiBm_$, obD_%, f6010_%, f6011_$
Dim Lanr&, BSNR&
Dim FormID&
Dim FormAbk$, FormBez$, FormVorl$, FormInh$, obFormDMP%, FormSp$
Dim BGFallNr$ ' für dolies, neuQuartal
Public FoIDv& ' Pseudo-Foid
#If thaalt Then
 Public rmpn() As medplan ' nach Datum umsortierter Plan ' ausrangiert 12.12.20
#End If
Dim DokPfad$, DokArt$, DokName$
'Dim lVorl As Date ' letzte Vorlage, 4109
#Const obPraxis = False
#If obPraxis Then
'Public Const uVerz$ = "u:\"
'Public Const pVerz$ = "p:\"
#Else
'Public Const uVerz$ = "c:\u\"
'Public Const pVerz$ = "c:\p\"
#End If
'Public Const hVerz$ = uVerz & "TMImport\" ' HochladeVerzeichnis
'Public Const aVerz$ = uVerz & "Anamnese\"
'Public Const QmdB$ = aVerz & "Quelle.mdb" ' uverz & "Anamnese\Quelle.mdb
'Public Const eVerz$ = pVerz & "Patientenübergreifendes"
'Public Const üVerz$ = uVerz & "TMExport\"
'Public Const LabTransPfad$ = "\\server\BioWinBACKUP"
Public Arra() ' Array für Telefonnummern
Public ArraInd%
Dim lfdfl&, AktZeit As Date, f8000$, f8100$
Public lAktZeit As Date ' bei Namen steht die letzte Datenänderung aus irgend einem zu dem Namen gehörigen Satz aus BDT-Datei drin, um nur aktuelle aktualisieren zu können
    
Public Diag$() '  = rdi(0).DiagText
Public ICD$() '  = rdi(0).ICD
Public DSic$() ' = rdi(0).DiagSicherheit
Public DiagSe$() ' = rdi(0).DiagSeite
Public DiagAttr$() ' =rdi(0).DiagAttr
Public DiagAus$() ' rdi(0).AusnBegr
Public DiagiBm$() ' = rdi(0).intBemerk
Public diagdat() As Date ' = rdi(0).DiagDatum
Public obDauer%() ' = rdi(0).obDauer
Public obKasse%() ' = rdi(0).obKasse
Public lKasse() As Date ' = rdi(0).lKasse
Public f6010%() ' =rdi(0).f6010
Public f6011$() ' =rdi(0).f6011
Dim gk() As Boolean ' abgehakt
Public G1%()       ' diagreihe.gi1
Public G2%()     ' diagreihe.gi2
Dim DiagNr%
'Dim nachname$, vorname$, GebDat As Date, Titel$, nvorsatz$, GesamtName$

Const DiagMaxZahl% = 40
' Variablen für Anamnesebogen
Dim DMSchulz%, DMSchL%, RRSchulz%, dmpdat As Date, Tkz%, VorStDat As Date, obMedNetz%
'Dim obMedPlan%, obMedPlanGelesen%
Dim NEin0% ' Nameneintr0 wurde aufgerufen
Dim arr$, gesRR$, lRR$  ' gesamter, Anfangs- und letzter Blutdruck
Dim obEinweisung%
'Dim EintrArt$, EintrInh$
Dim FStG$, Abkü$, AbküLabor$ ' , Wert$ ', Einheit$ ' , Kommentar$, Langtext$
Dim obLaborEintrag% ' nicht bei vorläufigen Befunden nach den endgültigen
Dim AlbErhöht$, AlbDat As Date
Dim lHbA1c$, lHbA1cDat As Date  ' für Anamnesebogen
Dim absPos&
Dim DStr$ ' Datumsudt.mwandel-String für Blutdruckangabe
Dim MPDatum As Date
'Public Med$(), Dos$()
#If altMed Then
Public MedNr%
#End If
Public MedZahl%, obHergang%, pMpnr& ' für Medikamentenplan: ..., patientenspez. mpnr
Dim obmed As Boolean, lDos%, MedArt$
Dim ZeileNr% ' wenn mehrere 6299 auf ein 6298 folgen, dann können alle eingetragen werden
Dim rAF& ' records affected
Public Const obForeign% = True
Public PcDokPfad$
Public Fil As File
Public medSL As SortierListe ' für Medklass
 ' ehemals statische Variablen einzelner Prozeduren, hier für erneuten Einlesvorgang rausgezogen
Public rsAnm As New ADODB.Recordset
Dim Quartal$ ' 4101
Dim lfdnr& ' für Tabelle Namen / in GesLies: Zahl der eingelesenen Patienten
Dim lKennung% ', lk%(15)
Dim jetztKopf%, rFm_Nr&, jetztCGM%
Dim lFormID& ' ID des letzten Formulars
Dim obZweiteRunde%
Dim lLang$, lLangVW&
Dim lKomm$, lKommVW&
Dim SpeicherZt As Date
Dim AktQ$ ' aktuelles Quartal
Dim obDStr% ' ob Diagnose in Diagnosestring einzutragen ist
Dim maxBhFB As Date, imaxBhFB% ' Fall mit dem maximalen Behandlungsfallbeginn
Dim mSL As SortierListe ' für medarten
'Public rAna() AS Ana
'Public rFm() AS `forminhfeld`
Public ls% 'Laborsatz
Public sFld As sFeld, sListFeld As New SortierListe
Public sFldI As sFeldInh, sListFldInh As New SortierListe
Public sLp As sLpar, sListLpar As New SortierListe
Public altAusgabe$ ' Ausgabefensterinhalt, der bleiben soll
Public AltBDT$, AltBDTDat As Date  ' letzter Einlesevorgang
Enum WieWeiter
 D1Weiter = -1
 beidemvgl = 0
 d2weiter = 1
 beideovgl = 2
End Enum ' WieWeiter
Enum FormUntTyp
 Fgz
 AnamnKörperMasse
 BefundBlutdruck
 Schwangerschaft
End Enum ' FormUntTyp
Dim Kassengeändert% ' Kassen wurden geändert, Kategorie muß bestimmt werden
Dim IKneu% ' IK wurde aus Feld 6299 neu gesetzt, war vorher "" => Kassenname auch aktualisieren
 
 Type iLanrTyp
  id As Long
  Lanr As Long
  krz As String
 End Type
 Public iLanr() As iLanrTyp
'Module General section
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long ' ob saArray initialisiert
' alternativ:
'     Dim ubiL&
'     ubiL = -1
'     ON Error Resume Next
'     ubiL = UBound(iLanr)
'     ON Error GoTo fehler
'     IF ubiL <> -1 THEN
Dim auti%() ' aut idem, hat bei Langrezepten umgekehrte Bedeutung wie bei Kassenrezepten und in der Datenbank
Dim anzl%() ' Anzahl bei Kassenrezepten
Dim mdnr% ' Medikament Nr im Rezept (0-2)
Public qbeg As Date

' in holiAzu
Public Function fuellilanr()
' LANR-Tabelle vorbereiten
  If SafeArrayGetDim(iLanr) = 0 Then ' wenn iLanr initialisiert
   Dim rila As New ADODB.Recordset
   myFrag rila, "SELECT `ID`,`LANR`,krz FROM `lanrpraxis`"
   If Not rila.BOF Then
    Do While Not rila.EOF
     If SafeArrayGetDim(iLanr) = 0 Then ReDim iLanr(0) Else ReDim Preserve iLanr(UBound(iLanr) + 1)
     iLanr(UBound(iLanr)).id = rila!id
     iLanr(UBound(iLanr)).Lanr = rila!Lanr
     iLanr(UBound(iLanr)).krz = rila!krz ' Kürzel
     rila.MoveNext
    Loop
   End If ' Not rila.BOF Then
   Set rila = Nothing
  End If ' SafeArrayGetDim(iLanr) = 0 THEN
End Function ' fuellilanr()

' in dolies
Public Function holiAzu&(ByVal Lanr&) ' hole interne Arztzuordnungsnummer
  Dim i&
  Call fuellilanr
  If SafeArrayGetDim(iLanr) <> 0 Then ' wenn iLanr initialisiert
   For i = 0 To UBound(iLanr)
    If iLanr(i).Lanr = Arra(1) Then
     holiAzu = iLanr(i).id
     Exit Function
    End If
   Next i
  End If
  myEFrag "INSERT INTO `lanrpraxis`(lanr,Nachname,Vorname,Titel,Anschrz,Strasse,Hausnummer,PLZ,Stadt,Telefon,Telefax,Abteilung,Name,KVNr,BSNR,Betriebsstaettename) VALUES(" & Lanr & ",'','','','','','','','','','','','','','','')", rAF
  If rAF = 1 Then
   holiAzu = myEFrag("SELECT last_insert_id() id")!id
' folgende Alternative brauchts ned, wenn DBCn nicht gleichzeitig von zweitem Thread bedient wird:
'   holiAzu = myEFrag("SELECT id FROM lanrpraxis WHERE lanr=" & LANR)!id
  Else
   holiAzu = -1
  End If
  If SafeArrayGetDim(iLanr) = 0 Then ReDim iLanr(0) Else ReDim Preserve iLanr(UBound(iLanr) + 1)
  iLanr(UBound(iLanr)).id = holiAzu
  iLanr(UBound(iLanr)).Lanr = Lanr
End Function ' holiAzu&(LANR&)

Function ArrayString(ParamArray tokens()) As String()
    Dim arr
    ReDim arr(UBound(tokens)) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        arr(i) = tokens(i)
    Next
    ArrayString = arr
End Function ' ArrayString

Function ASMod(ParamArray tokens()) As String()
    Dim arr
    ReDim arr(UBound(tokens)) As String
    Dim i&
    For i = 0 To UBound(tokens)
     arr(i) = " " & tokens(i) & "=\"""
    Next
    ASMod = arr
End Function ' ASMod

' Tabellen vorbereiten für BDT-Import
' in Geslies
Function doTabVorb(frm As Lese, obInhalt%, obmitFormularen%)
 Const MaxLauf& = 100
 Dim lauf&
 Dim i&, T2#
 On Error GoTo fehler
 Dim rs As New ADODB.Recordset, rAF&
' Dim cn As New ADODB.Connection,
' cn.Open CSStr, , , 0
' SET rs = myEFrag("SELECT * FROM quelle.`forminhaltfeld`", rAf) ' + adAsyncExecute,,cn)
' Call rs.Open("quelle.`forminhaltfeld`", cn, adOpenStatic, adLockReadOnly)
 Set rs = myEFrag("SELECT abkü, labor, langtext, einheit FROM `laborparameter`")
 Do While Not rs.EOF
  Set sLp = New sLpar
  sLp.Abkü = IIf(IsNull(rs.Fields(0)), vNS, rs.Fields(0)) ' Abkü 6.3.13
  sLp.Labor = IIf(IsNull(rs!Labor), vNS, rs!Labor)
  sLp.Einheit = IIf(IsNull(rs!Einheit), vNS, rs!Einheit)
  sLp.Langtext = IIf(IsNull(rs!Langtext), vNS, rs!Langtext)
  Call sListLpar.sCAdd(sLp, -1)
  rs.Move 1
 Loop
 rs.Close
 On Error GoTo F0
vorabfrag:
 lauf = lauf + 1
' Set rs = myEFrag("SELECT feld,feldvw FROM `forminhaltfeld`")
 myFrag rs, "SELECT feld,feldvw FROM `forminhaltfeld`"
 Do While Not rs.EOF
  On Error GoTo fehler
  Dim rsFeld
  rsFeld = rs!Feld
  Set sFld = New sFeld
  If IsNull(rsFeld) Then
   sFld.Feld = vNS
  Else
   sFld.Feld = rsFeld
  End If
  sFld.FeldVW = rs!FeldVW
' IF Not sListFeld.SuchItem(sFld) THEN
  Call sListFeld.sCAdd(sFld, -1)
' END IF
  rs.Move 1
 Loop
 rs.Close
 If obInhalt Then
 ' Call rs.Open("quelle.forminhaltfeldinh", cn, adOpenStatic, adLockReadOnly)
  Set rs = myEFrag("SELECT feldinh,feldinhvw FROM `forminhaltfeldinh`", rAF) ', adAsyncExecute)
  frm.Bytes = "(FormInhaltFeldInh)"
  Do While Not rs.EOF
   Set sFldI = New sFeldInh
   Dim rsFeldInh
   
   rsFeldInh = rs!FeldInh
   If IsNull(rsFeldInh) Then rsFeldInh = vNS
   sFldI.FeldInh = rsFeldInh
   sFldI.FeldInhVW = rs!FeldInhVW
 '  IF Not sListFldInh.SuchItem(sFldI) THEN
    Call sListFldInh.sCAdd(sFldI, -1)
 '  END IF
   rs.Move 1
   i = i + 1
   If i Mod 100 = 0 Then
    frm.Zeilen = i
    DoEvents
   End If
  Loop
  frm.Bytes = 0
  T2 = Now
 ' For i = 1 To sListFldInh.Count
 '  Debug.Print sListFldInh.Item(i).FeldInh
 ' Next i
 End If ' obInhalt
 
 Set rs = myEFrag("SELECT DISTINCT kennung FROM `unbekannte kennungen`", rAF)
 Do While Not rs.EOF
  ReDim Preserve rUn(UBound(rUn) + 1)
  rUn(UBound(rUn)).Kennung = rs!Kennung
  rs.Move 1
 Loop
 rUn1 = UBound(rUn)
 
 rFo(0).FormID = 0
 ReDim rFo(0)
 If obmitFormularen Then
 ' Call rAdo.Open("SELECT * FROM `formulare` ORDER BY formid", DBCn, adOpenDynamic, adLockReadOnly)
  Set rs = myEFrag("SELECT * FROM `formulare` ORDER BY formid")
 ' myFrag rs, "SELECT * FROM `formulare` ORDER BY formid" ', adOpenForwardOnly
  If Not rs.BOF Then
   rs.MoveFirst
   Do While Not rs.EOF
    ReDim Preserve rFo(UBound(rFo) + 1)
    rFo(UBound(rFo)).Form_Abk = IIf(IsNull(rs!Form_Abk), vNS, rs!Form_Abk)
    rFo(UBound(rFo)).FormBez = rs!FormBez
    rFo(UBound(rFo)).FormVorl = rs!FormVorl
    rFo(UBound(rFo)).FormID = rs!FormID
    rs.MoveNext
   Loop ' While Not rs.EOF
  End If ' Not rs.BOF Then
 End If ' obmitformularen
 rFo1 = UBound(rFo)
 
 rFi(0).Form_AbkVW = 0
 Set rs = myEFrag("SELECT Form_Abk,Form_AbkVW FROM forminhaltform_abk ORDER BY form_abkvw")
' myFrag rs, "SELECT * FROM forminhaltform_abk ORDER BY form_abkvw"
 If Not rs.BOF Then
  rs.MoveFirst
  Do While Not rs.EOF
   ReDim Preserve rFi(UBound(rFi) + 1)
   rFi(UBound(rFi)).Form_Abk = IIf(IsNull(rs!Form_Abk), vNS, rs!Form_Abk)
   rFi(UBound(rFi)).Form_AbkVW = rs!Form_AbkVW
   rs.MoveNext
  Loop
 End If
 rFi1 = UBound(rFi)
 Exit Function
F0:
 If Err.Number = 3704 And lauf < MaxLauf Then ' Der Vorgang ist für ein geschlossenes Objekt nicht zugelassen.
  Resume vorabfrag
 End If
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabVorb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doTabVorb

' Versuch mit Load Data Infile sehr negativ ausgefallen: 45" ggü. 9" Einlesezeit und falscher Zeichensatz,
' unabhängig vom Zeichensatz der der Datenbank (auch Dos usw. versucht)
' falls nicht über MySQL Query Browser eingelesen!
' Datenbank einl2 daher wieder gelöscht. Hier noch die Creates:
'
'`einl2`.CREATE DATABASE `einl2` /*!40100 DEFAULT CHARACTER SET cp850 */;
'
'CREATE TABLE  `einl2`.`latin2` (
'  `lenge` varchar(3) DEFAULT NULL,
'  `kennung` varchar(4) DEFAULT NULL,
'  `inhalt` varchar(150) DEFAULT NULL
') ENGINE=MyISAM DEFAULT CHARSET=latin2 ROW_FORMAT=DYNAMIC;
'
'CREATE TABLE  `einl2`.`zuleing` (
'  `dname` varchar(120) DEFAULT NULL,
'  `bytes` varchar(12) DEFAULT NULL,
'  `geändert` datetime DEFAULT NULL
') ENGINE=MyISAM DEFAULT CHARSET=cp850;
Function DTEst()
' Call DVgl(uverz & "TMExport\x0119153b.BDT", uverz & "TMExport\x0119153aBDT.txt")
 Call DVgl(uVerz & "TMExport\x0119153b.BDT", uVerz & "TMExport\x0119153va.BDT")
End Function ' DTEst()

Function DVgl(D1$, D2$)
 Dim T1$, T2$, k1$, K2$, i1$, i2$, Pat1$, Pat2$
 Const i2max& = &H7FFFFFFF
 Dim steuer As WieWeiter  ' -1 = D1 muss weiter gelesen werden,
 Dim neuSteuer As WieWeiter, obD1neu%, obD2neu%
'               0 = beide mit Vergleich,
'               2 = beide ohne Vergleich,
'               1 = D2 muss weiter gelesen werden
 On Error GoTo fehler
 Open D1 For Input As #317
 Open D2 For Input As #319
 Set pCol = New Collection
 steuer = beidemvgl
 Do
   obD1neu = 0
   obD2neu = 0
   If steuer <> d2weiter Then
    If EOF(317) Then
     If i1 <> i2max Then
      k1 = "3000"
      i1 = i2max
     Else
      Exit Do
     End If
    Else
     Line Input #317, T1
     k1 = Mid$(T1, 4, 4)
     i1 = Mid$(T1, 8)
    End If
    If k1 = "3000" Then
     If Pat1 <> i1 Then
      obD1neu = True
     End If
    End If
   End If
   If steuer <> D1Weiter Then
    If EOF(319) Then
     If i2 <> i2max Then
      K2 = "3000"
      i2 = i2max
     End If
    Else
     Line Input #319, T2
     K2 = Mid$(T2, 4, 4)
     i2 = Mid$(T2, 8)
    End If
    If K2 = "3000" Then
     If Pat2 <> i2 Then
      obD2neu = True
      DoEvents
     End If
    End If
   End If
   If Not obD1neu And Not obD2neu Then
    If k1 = K2 And i1 = i2 Then
    Else
     If steuer = beidemvgl Then
      If LenB(Pat1) <> 0 Then Call pCol.Add(Pat1)
      neuSteuer = beideovgl
     End If
    End If
   ElseIf obD1neu And obD2neu Then
    If i1 = i2 Then
     neuSteuer = beidemvgl
    ElseIf i1 < i2 Then
     Call pCol.Add(i1)
     neuSteuer = D1Weiter
    Else
     neuSteuer = d2weiter
    End If
   ElseIf obD1neu And Not obD2neu Then
    If steuer = beidemvgl Then
     Call pCol.Add(Pat1)
     neuSteuer = d2weiter
    ElseIf steuer = beideovgl Then
     neuSteuer = d2weiter
    ElseIf steuer = D1Weiter Then
     If i1 = Pat2 Then
      neuSteuer = beidemvgl
     ElseIf i1 < Pat2 Then
      Call pCol.Add(i1)
      neuSteuer = D1Weiter
     ElseIf i1 > Pat2 Then
      neuSteuer = d2weiter
     End If
    End If
   Else ' IF Not obD1neu AND obD2neu THEN
    If steuer = beidemvgl Then
     Call pCol.Add(Pat1)
     neuSteuer = D1Weiter
    ElseIf steuer = beideovgl Then
     neuSteuer = D1Weiter
    ElseIf steuer = d2weiter Then
     If Pat1 = i2 Then
      neuSteuer = beidemvgl
     ElseIf Pat1 < i2 Then
      Call pCol.Add(Pat1)
      neuSteuer = D1Weiter
     ElseIf i1 > Pat2 Then
      neuSteuer = d2weiter
     End If
    End If
   End If
   If obD1neu Then Pat1 = i1
   If obD2neu Then Pat2 = i2
   steuer = neuSteuer
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DVgl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DVgl

Public Function kompakt(frm As Lese)
  Dim AdoJRO As New JRO.JetEngine, neuDat$
  Dim obCnOffen%
  On Error GoTo fehler
  If frm.obAusgedehnt Then
   If Not DBCn Is Nothing Then If DBCn.State = 1 Then obCnOffen = -1
   If obCnOffen Then DBCn.Close
   SetDBCn Nothing, vNS
   Call AdoJRO.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & frm.dlg.MdB, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & frm.dlg.MdB & "comp")
   neuDat = Dir(frm.dlg.MdB & "comp")
   If LenB(neuDat) <> 0 Then
    Kill frm.dlg.MdB
    Name frm.dlg.MdB & "comp" As frm.dlg.MdB
    frm.Ausgeb "Fertig mit Komprimieren von " & frm.dlg.MdB, True
   End If
   If obCnOffen Then
    Call acon(quelleT)
   End If
   frm.obAusgedehnt = 0
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kompakt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' kompakt
'use quelle4;
'DROP TABLE eintrhist2;
'DROP TABLE eintrhist1;
'ALTER ignore TABLE `eintraege` drop COLUMN `ungültigab`;
'CREATE TABLE eintrhist1 LIKE `eintraege`;
'ALTER TABLE `eintrhist1` ADD COLUMN `ID` INTEGER(10) UNSIGNED NOT NULL DEFAULT NULL AUTO_INCREMENT AFTER `StByte`,
' ADD PRIMARY KEY (`ID`);
'CREATE TABLE eintrhist2 LIKE eintrhist1;
'ALTER TABLE `eintrhist1`
' DROP COLUMN `FID`,
' DROP COLUMN `absPos`,
' DROP COLUMN `AktZeit`,
' DROP COLUMN `StByte`;
'ALTER TABLE `eintrhist2`
' DROP COLUMN `Pat_ID`,
' DROP COLUMN `ZeitPunkt`,
' DROP COLUMN `Art`,
' DROP COLUMN `Inhalt`,
' DROP COLUMN `QS`,
' DROP COLUMN `QT`;
'ALTER TABLE `eintrhist2` MODIFY COLUMN `ID` INTEGER UNSIGNED NOT NULL DEFAULT 0,
' DROP PRIMARY KEY;
'ALTER TABLE `eintrhist2` ADD CONSTRAINT `FK_id` FOREIGN KEY `FK_id` (`ID`)
'    REFERENCES `eintrhist1` (`ID`)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT;
'DROP TABLE eintrhist;
'DROP VIEW eintrhist;
'CREATE VIEW eintrhist AS SELECT * FROM eintrhist1 INNER JOIN eintrhist2 USING (id);
Function eintrhist()
 Dim hist$, rq As New ADODB.Recordset, rz1 As New ADODB.Recordset, rz2 As New ADODB.Recordset, rAF&
 On Error GoTo fehler
#If True Then
Call myEFrag("INSERT INTO eintrhist1(pat_id,zeitpunkt,art,inhalt,qs,qt)" & vbCrLf & _
"SELECT pat_id,zeitpunkt,art,inhalt,qs,qt" & vbCrLf & _
"FROM eintraege e" & vbCrLf & _
"WHERE pat_id IN (SELECT pat_id FROM namen WHERE nachname RLIKE " & ExtraPatienten & " AND art<>'PKon')" & vbCrLf & _
"AND NOT EXISTS" & vbCrLf & _
"(SELECT 0 FROM eintrhist1 WHERE pat_id=e.Pat_ID AND zeitpunkt=e.zeitpunkt AND art=e.art AND inhalt=e.Inhalt AND qs=e.qs AND qt=e.qt)" _
, rAF, DBCn)
Call myEFrag("INSERT INTO eintrhist2(fid,abspos,aktzeit,stbyte,id)" & vbCrLf & _
"SELECT e.fid,e.abspos,e.aktzeit,e.stbyte,eh.id" & vbCrLf & _
"FROM eintraege e LEFT JOIN eintrhist1 eh USING (pat_id,zeitpunkt,art,inhalt,qs,qt)" & vbCrLf & _
"WHERE e.pat_id IN (SELECT pat_id FROM namen WHERE nachname RLIKE " & ExtraPatienten & " AND art<>'PKon')" & vbCrLf & _
"AND NOT EXISTS" & vbCrLf & _
"(SELECT 0 FROM eintrhist2 FORCE INDEX (FälleEinträge) WHERE fid=e.fid AND abspos=e.abspos AND aktzeit=e.aktzeit AND stbyte=e.stbyte AND id=eh.id)" _
, rAF, DBCn)
#Else
 Set rq = myEFrag("SELECT GROUP_CONCAT(pat_id) FROM `namen` WHERE nachname RLIKE " & ExtraPatienten, , , , , , 1000)
 While IsNull(rq.Fields(0))
  hist = ""
 Wend
 If Not IsNull(rq.Fields(0)) Then
  hist = rq.Fields(0)
 End If
 If LenB(hist) <> 0 Then
  Set rq = Nothing
  myFrag rq, "SELECT * FROM `eintraege` WHERE pat_id IN (" & Left$(hist, Len(hist) - 1) & ") AND art <> 'Pkon'", adOpenStatic
  Do While Not rq.EOF
   Do
    Set rz1 = Nothing
    myFrag rz1, "SELECT id FROM `eintrhist1` WHERE pat_id = " & rq!Pat_id & " AND zeitpunkt = " & DatFor_k(rq!Zeitpunkt) & " AND art = '" & rq!art & "' AND inhalt = '" & UmwfSQL(rq!Inhalt) & "' AND qs = '" & rq!QS & "' AND qt = '" & rq!QT & "'"
    If rz1.BOF Then
     InsKorr DBCn, DBCnS, "INSERT INTO `eintrhist1`(pat_id,zeitpunkt,art,inhalt,qs,qt) VALUES(" & rq!Pat_id & "," & DatFor_k(rq!Zeitpunkt) & ",'" & rq!art & "','" & UmwfSQL(rq!Inhalt) & "','" & rq!QS & "','" & rq!QT & "')", rAF
    Else
     Exit Do
    End If
   Loop
   Set rz2 = Nothing
   myFrag rz2, "SELECT id FROM `eintrhist2` WHERE fid = " & rq!FID & " AND abspos = " & rq!absPos & " AND aktzeit = " & DatFor_k(rq!AktZeit) & " AND stbyte = " & rq!StByte & " AND id = " & rz1!id
   If rz2.BOF Then
    myEFrag "INSERT INTO `eintrhist2`(fid,abspos,aktzeit,stbyte,id) VALUES(" & rq!FID & "," & rq!absPos & "," & DatFor_k(rq!AktZeit) & "," & rq!StByte & "," & rz1!id & ")", rAF
   End If
   rq.Move 1
  Loop ' While Not rq.EOF
 End If ' LenB(hist) <> 0 Then
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf, vbAbortRetryIgnore, "Aufgefangener Fehler in eintrhist/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' eintrhist

'In code
'If SafeArrayGetDim(myArray) <> 0 THEN
'    MsgBox "Array has been Initialized"
'END IF
'Public FUNCTION ilanrtest()
'  Dim rs As New ADODB.Recordset
'  Lese.ProgStart
'  myFrag rs, "SELECT * FROM `lanrpraxis`"
'  Do While Not rs.EOF
'   IF SafeArrayGetDim(iLanr) = 0 THEN ReDim iLanr(0) ELSE ReDim iLanr(UBound(iLanr) + 1)
'   rs.Move 1
'  Loop
'  rs.Close
'End Function

' wird nur in doEinles aufgerufen
Function GesLies(frm As Lese, BDTDatei$, BDTName$, EinlAb&, EinlBis&, obLaborDirekt%, obLDneu%, obLaborQuer%, obLQneu%, obEmails%, EmDatei$)
' Dim rAdo As New Adodb.Recordset
' Dim rs As New Adodb.Recordset
 
 Dim ltxtS As New CString 'letzter Text
 Dim tx1S As New CString, tx2S As New CString ' Eingelesener Text aus BDT (324), aus frm.dlg.LDatei
 Dim K1S As New CString, K2S As New CString ' Kürzel1 (BDT), Kürzel2
' Dim i1s As New CString, i2s As New CString ' Inhalt 1 (-> Patient 1), Inhalt 2 (->Patient 2)
 Dim i1s$, i2s$, GesBytes#
 Dim zwi$, znr&, obmitFormularen%
 Dim Pat1$, Pat2$
 Const i2max& = &H7FFFFFFF ' Maximale Long-Zahl
 Dim steuer As WieWeiter  ' Steuerung des Lesesvorgangs: -1 = D1 muss weiter gelesen werden,
 Dim neuSteuer As WieWeiter ' Steuerung des nächsten Lesevorgangs
 Dim obD1neu% ' Vergleich mit letzter Einlesung gewünscht und letzte Zeile in BDT-Datei enthält neuen Patienten
 Dim obD2neu% ' Vergleich mit letzter Einlesung gewünscht und letzte Zeile in alter Datei enthält neuen Patienten
 Dim Zeilen&, Bytes# ' für den Anzeigevorgang, inhaltlich nicht wichtig
 Dim Cpt$, i&
 Dim Tm1!, Tm2!, Dauer!, DauerZ&
 Dim rsinl As New ADODB.Recordset
 On Error GoTo fehler
 
 If IsNumeric(frm.GesBytes) Then GesBytes = frm.GesBytes
 Call controlsLauf(frm) ' Unterbrechbefehl aktivieren, Abbrech-Befehl deaktivieren
 Call StatischInit
 Call medartenhier(mSL)
 Dauer = 0
 DauerZ = 0
 BezFeh = pVerz & "BezFehler_" & DefDB(DBCn) & "_" & Format(Now(), "YYYYMMDD_hhmmss") & ".txt"
 If obÜProt Then Open üVerz & "ÜbertrProtMySQL.txt" For Output As #322 ' Nummer ab 256: gleichzeitig für andere Prozesse zugreifbar
 If frm.dlg.obBDT Then
  If frm.dlg.ÜberTabelle <> 0 Then
   frm.Ausgeb "Entferne Vortabelle `" & DefDB(DBCn) & "`.`inlalt` ...", False
   myEFrag "DROP TABLE IF EXISTS `inlalt`", rAF
   frm.Ausgeb "Kopiere Tabelle `" & DefDB(DBCn) & "`.`inlk` in Vortabelle `" & DefDB(DBCn) & "`.`inlalt` ...", False
   On Error Resume Next
   myEFrag "RENAME TABLE `inlk` to `inlalt`", rAF
   On Error GoTo fehler
   frm.Ausgeb "Erstelle Tabelle `" & DefDB(DBCn) & "`.`inlk` ...", False
   myEFrag "CREATE TABLE IF NOT EXISTS `inlk` (id int(10) UNSIGNED NOT NULL AUTO_INCREMENT, stbyte int(10), breite char(3) NOT NULL, kennung char(4) NOT NULL, inhalt varchar(10000) NOT NULL, PRIMARY KEY (id), KEY kennunginhalt (kennung,inhalt(6))) ENGINE=MyISAM DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci", rAF
   frm.Ausgeb "Leere Tabelle `" & DefDB(DBCn) & "`.`inlk` ...", False
   myEFrag "DELETE FROM inlk", rAF
   frm.Ausgeb "Stelle Zähler in Tabelle `" & DefDB(DBCn) & "`.`inlk` ...", False
   myEFrag "ALTER TABLE `inlk` auto_increment=0", rAF
   frm.Ausgeb "Lade Daten in Tabelle `" & DefDB(DBCn) & "`.`inlk` (mit: LOAD DATA INFILE), kann dauern ...", False
   myEFrag "LOAD DATA INFILE '/DATA/eigene Dateien/TMExport/" & BDTName & "' INTO TABLE `inlk` fields terminated by '' enclosed by '' lines terminated by '\r\n'  (breite,kennung,inhalt)", rAF
   If frm.dlg.NurInTabelle Then
    Exit Function
   End If
   frm.Ausgeb "Lese Tabelle `" & DefDB(DBCn) & "`.`inlk` ...", False
   myFrag rsinl, "SELECT * FROM `inlk`"
  Else ' frm.dlg.ÜberTabelle <> 0 Then
   Open BDTDatei For Input As #324
  End If ' frm.dlg.ÜberTabelle <> 0 Then else
 
  Call Tinit
'  Call Zinit
  #If False Then
   If ProgrammLauf(0, Cpt) Then ' wenn Programm schon (anderweitig) aufgerufen, dann nicht nochmal gleichzeitig aufrufen
    MsgBox "Programm schon auf " + Cpt + " aufgerufen, breche ab!"
    GoTo abbrech
   End If
  #End If
  Call eintrhist ' 12.10.08
  If AllePat Then
   Call AllesLösch(frm)
'  Call löschBezügeausLaborux
  End If
  obmitFormularen = FSO.GetFile(BDTDatei).size > 1000000
  Call doTabVorb(frm, obVorb, obmitFormularen)
 
  Set medSL = New SortierListe ' für Medklass
  Call Eintragszl(1)
#If ohnebezug Then
  Call ForeignNo0
  Call ForeignNo1
#End If
  T1a = Now
  Dim obEOF%, obAngeh%, obSchluss% ' ob die BDT-Datei zuende ist
  If frm.dlg.ÜberTabelle <> 0 Then
   obEOF = rsinl.EOF
  Else
   obEOF = EOF(324)
  End If
 
  Dim cKenn(1000000) As String * 4
  Dim cInha(1000000) As String '* 72 ' feste Breiten 12", variable 14" => aber feste führt zu unerwarteten Ergebnissen bei right$(... ; 100000: 11"; 300000: 12", 500000: 12",1000000: 12"
  Dim ind& ' Index auf cKenn und cInha
  Dim altpat$ ' letzter Patient
  Dim aktdat As Date, aktDatD As Date ' Aktualisierungzeitpunkt / -datum (BDT-Felder 5000 und 6200)
  Dim fallzahl& ', Zahl der verarbeiteten Fälle
  Dim obVorspannVorbei%
  Dim obAndPat%
  Dim obAuslassen% ' Variable nicht mit obVorspannvorbei identisch setzen, damit auch bei den ausgelassenen Patienten lfdnr hochgezählt werden kann
  Dim aktPatAusg As New CString ' aktuelle Patientendaten für Ausgabe
  Dim spos& ' für seek

  If frm.dlg.obVglMitLetzterEinlesung <> 0 And Not obHausBesuch Then
   Close #318
   If Not FSO.FileExists(frm.dlg.LDatei) Then
    If LenB(frm.dlg.LDatei) = 0 Then
     MsgBox ("Fehler: dlg.LDatei nicht angegeben!")
    Else
     MsgBox ("Fehler: dlg.LDatei: '" & frm.dlg.LDatei & "' nicht gefunden!")
    End If
    If frm.dlg.ÜberTabelle <> 0 Then
     Set rsinl = Nothing
    Else
     Close #324
     obEOF = True
    End If
    Exit Function
   End If
   Open frm.dlg.LDatei For Input As #318
   Set pCol = New Collection
   steuer = beidemvgl
  End If ' frm.dlg.obVglMitLetzterEinlesung <> 0 AND NOT obHausBesuch THEN
 
  If EinlAb > 0 Then obAuslassen = True
  Do ' Große Einleseschleife
   obD1neu = 0
   obD2neu = 0
' Es folgt die Zeitabschätzung für den Zuschauer
   Zeilen = Zeilen + 1
   Bytes = Bytes + tx1S.length + 2
   If Zeilen Mod 100 = 0 Then
    frm.Zeilen = Zeilen
    frm.Bytes = Format$(Bytes, "###,###,###,###,###,###,###,##0")
    frm.Sekunden = Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    If GesBytes <> 0 Then frm.Prozent = Format$(frm.Bytes / GesBytes * 100, "0.00")
    frm.Beginn = Format$(T1, "hh:mm:ss")
    If GesBytes <> 0 Then
     frm.GesDauer = Format$((T1a - T1 + GesBytes / frm.Bytes * (Now - T1a)), "hh:mm:ss")
     frm.EndZp = Format$(T1 + (T1a - T1 + GesBytes / frm.Bytes * (Now - T1a)), "hh:mm:ss")
    End If
    DoEvents
   End If ' Zeilen Mod 100 = 0 THEN
   
' Wenn die letzte Einlesung nicht verglichen wird oder nicht D2 weitergelesen werden muss
   If frm.dlg.obVglMitLetzterEinlesung = 0 Or obHausBesuch Or steuer <> d2weiter Then
    If obEOF Then ' wenn außerdem D1 schon zuende ist
     ' 9.11.21: falls erste Bedingung nicht gilt <> falls zweite Bedingung gilt!
     If Not (frm.dlg.obVglMitLetzterEinlesung = 0 Or obHausBesuch) And i2s <> CStr(i2max) Then ' und wenn, falls D2 nicht weitergelesen werden muss, I2 noch nicht auf i2max steht
      K2S = "3000"
      i2s = i2max ' dann i2 auf i2max stellen
     Else
      Exit Do ' sonst mit großer Einleseschleife aufhören
     End If
    Else ' bzw., falls also D1 noch nicht zuende,
     If frm.dlg.ÜberTabelle <> 0 Then
      rsinl.Move 1
      If Not rsinl.EOF Then tx1S = rsinl!breite & rsinl!Kennung & rsinl!Inhalt Else tx1S = vNS: obEOF = True
     Else
      znr = znr + 1
      On Error GoTo ef324
      Line Input #324, zwi ' dann D1 einlesen
      On Error GoTo fehler
      tx1S = zwi
     End If
'     IF tx1 = "01330002148" THEN      ' 2.11.07: ist der letzte Patient in alter Datei länger -> Endlosschleife mit steuer = d2weiter
     If steuer <> d2weiter Then ' und ggf. für den Vergleich zur Verfügung stellen
      K1S = tx1S.Mid(4, 4)
      i1s = tx1S.Mid(8)
     End If ' xxx
    End If
    If steuer <> d2weiter Then
     If K1S = "3000" Then
      If Pat1 <> i1s Then
       obD1neu = True ' wenn neuer Patient, dann kennzeichnen
       Pat1 = i1s
      End If
     End If ' K1S = "3000" Then
    End If ' steuer <> d2weiter Then
    
    If frm.dlg.ÜberTabelle <> 0 Then
     obEOF = rsinl.EOF
    Else
     If Not obEOF Then obEOF = EOF(324) ' wenn zuende, dann merken
    End If
    If InStrB(tx1S.Left(7), "#######") <> 0 Then 'IsNumeric(LEFT(tx1, 7)) THEN
     ltxtS.AppVar Array(" ", tx1S) ' ggf. über die Variable obAngeh kennzeichnen, dass die nächste Zeile noch angehängt werden muss
     obAngeh = -1
    ElseIf obAngeh Then
     obAngeh = 0
    End If
    
    Do ' kleine Schleife, Lesen einzelner Zeilen
'LiesEineZeile D1
     If obAngeh Then
      If obEOF Then
       obSchluss = True ' dann aufhören
      Else
       Exit Do          ' gleich aufhören
      End If
     Else ' obAngeh
      cKenn(ind) = ltxtS.Mid(4, 4) ' eingelesene Zeile verdauen
      cInha(ind) = ltxtS.Mid(8)
'     IF ZS THEN RInhalt = ZSU1(RInhalt)
      Select Case cKenn(ind)
       Case 3000
        If cInha(ind) <> altpat Then ' oder am Schluss!
         obAndPat = True ' eingelesene Zeile trug neuen Patienten
         If obVorspannVorbei > 0 Then lfdnr = lfdnr + 1 ' Zahl der Patienten steigt um 1
        End If
      End Select
      If obVorspannVorbei > 0 Then
       Select Case cKenn(ind)
 '     Case 3610, 3649, 5999, 4102, 4109, 4150, 5000, 6200 ' dauert 10" länger als die kurze Version!
           Case 3101: rNa(0).Nachname = ZSU1(doUmwfSQL(cInha(ind), lies.obMySQL))
           Case 3102: rNa(0).Vorname = ZSU1(doUmwfSQL(cInha(ind), lies.obMySQL))
           Case 3103:
            If Left$(cInha(ind), 4) = "0000" Then cInha(ind) = "0101" & Mid$(cInha(ind), 5) ' 19.7.21 Pat. 1083
            rNa(0).GebDat = BDTtoDate(ZSU1(doUmwfSQL(cInha(ind), lies.obMySQL)))
       End Select
'       IF frm.dlg.obVglMitLetzterEinlesung = 0 THEN ' probeweise auskommentiert 24.10.08
        Select Case cKenn(ind)
         Case 3649, 5000, 5999, 6200 ' 6200 = Tagesdatum (BDT-Beschreibung ZI 1994) ' SELECT idn, dt.kennung FROM (SELECT id-1 AS idn FROM `inlk` WHERE kennung = 6201) AS i LEFT JOIN `inlk` dt ON dt.id = i.idn GROUP BY dt.kennung; => 28.10.08: 3649: 18687 Stück, 5000: 50532 Stück, 5999: 1294 Stück, 6200: 527443 Stück
          aktdat = BDTtoDate(cInha(ind))
          aktDatD = aktdat
          If aktdat > lAktZeit Then
           lAktZeit = aktdat
          End If
 '     Case 4110, 6201
         Case 6201
          aktdat = aktDatD + BDTtoTime(cInha(ind))
          If aktdat > lAktZeit Then
           lAktZeit = aktdat
          End If
      ' 4110: Uhrzeit letzter Vorlage zu 4109
      ' 6201: zu 5000
        End Select
'       END IF
      End If ' obvorspannvorbei
     End If ' obAngeh -> not obAngeh
     If obAndPat Or obSchluss Then
      If obVorspannVorbei > 0 Then ' wenn neuer Patient und Vorspann schon vorbei
       If EinlBis <> 0 Then
        If altpat > EinlBis Then Exit Do ' kleine Schleife beenden
       End If
       If EinlAb = 0 Or altpat >= EinlAb Then
' wenn Pat von einlab bis einlbis
'      IF Not rs Is Nothing THEN IF rs.State = 1 THEN rs.Close
'      call foreignno
'      Call rs.Open("DELETE FROM `namen` WHERE pat_id = " & 0)
        Dim oblies%
        If obHausBesuch Then
         oblies = -1
        ElseIf frm.dlg.obVglMitLetzterEinlesung <> 0 Then ' AND NOT obHausBesuch Then
         oblies = 0
         If pCol.COUNT > 0 Then
          If pCol(pCol.COUNT) = cInha(0) Then
           oblies = -1
          ElseIf pCol.COUNT > 1 Then
           If pCol(pCol.COUNT - 1) = cInha(0) Then
            oblies = -1
           End If
          End If ' pCol(pCol.COUNT) = cInha(0) Then
         End If ' pCol.COUNT > 0 Then
        Else ' frm.dlg.obVglMitLetzterEinlesung <> 0 AND NOT obHausBesuch THEN
         Dim rsaktZeit As Date, rsAZS$
'         IF Not rs Is Nothing THEN IF rs.State = 1 THEN rs.Close
         Dim rsaz As New ADODB.Recordset
         myFrag rsaz, "SELECT aktzeit,nachname,vorname,gebdat,aufndat FROM `namen` WHERE pat_id = " & altpat
         oblies = 0
         If rsaz.BOF Then
          oblies = -1
          rsAZS = "          -        "
         Else
          If IsNull(rsaz!AktZeit) Or lAktZeit - rsaz!AktZeit > 0.0000115 Or (lAktZeit = 0 And rsaz!AktZeit = 0) Or FSO.GetFile(BDTDatei).size < 1000000 Then ' FileLen zu kurz 21.8.21
           oblies = -1
           rsaktZeit = IIf(IsNull(rsaz!AktZeit), 0, rsaz!AktZeit)
           rsAZS = Format$(rsaktZeit, "dd.mm.yyyy hh:mm:ss")
          End If
         End If ' rs.BOF Then else
         Set rsaz = Nothing
        End If ' frm.dlg.obVglMitLetzterEinlesung <> 0 AND NOT obhausbesuch THEN else
'       altAusgabe = frm.Ausgabe
'       aktPatAusg = lfdnr & ") " & Format$(VorStDat, "dd.mm.yy") & " - " & lAktZeit & " (eingetragen:" & rsAZS & ")" & ": (" & altpat & ") "
        aktPatAusg.Clear
        aktPatAusg.AppVar Array(lfdnr, ") ", Format$(VorStDat, "dd.mm.yy"), " - ", lAktZeit, " (eingetragen:", rsAZS, ")", ": ( ", altpat, " ) ")
        Lese.lzPID = CLng("0" & altpat)
        If oblies Then
         frm.Ausgeb aktPatAusg.Value & " verarbeiten ... ", False
         fallzahl = fallzahl + 1
         Tm1 = Timer
         DoEvents
         For i = 0 To ind - 1
'         Debug.Print cKenn(i), cInha(i)
          Call dolies(frm, cKenn(i), ZSU1(doUmwfSQL(cInha(i), lies.obMySQL, False)), obSchluss, znr, obmitFormularen) ' false = 13.8.15, um bei zusammenhängenden Zeilen das Leerzeichen zu erhalten
         Next i
#If thaalt Then
         Call rmeSort    ' ausrangiert 12.12.20
         Call THAfestleg ' ausrangiert 12.12.20
#End If
'        aktPatAusg = aktPatAusg & rNa(0).Nachname & " " & rNa(0).Vorname & ", " & rNa(0).GebDat & ";"
'         aktPatAusg.Clear
         aktPatAusg.AppVar Array(GesNameFn(rNa(0).NVorsatz, rNa(0).Nachname, rNa(0).Titel, rNa(0).Vorname, rNa(0).GebDat), ";")
         frm.Ausgeb aktPatAusg & " speichern ... ", False
         Call alleSpeichern(frm)
         frm.Ausgeb aktPatAusg & " fertig!", True
         Tm2 = Timer
         Dauer = Dauer + Tm2 - Tm1
         DauerZ = DauerZ + 1
         Lese.DurchschnDauer = Round(Dauer / DauerZ / 100 * 100, 2)
         Call Tinit
        Else ' oblies THEN
'         aktPatAusg = aktPatAusg & rNa(0).Nachname & " " & rNa(0).Vorname & ", " & rNa(0).GebDat & ";"
'         aktPatAusg.Clear
         aktPatAusg.AppVar Array(GesNameFn(rNa(0).NVorsatz, rNa(0).Nachname, rNa(0).Titel, rNa(0).Vorname, rNa(0).GebDat), ";")
         frm.Ausgeb aktPatAusg & " übersprungen", True
        End If ' oblies THEN
       Else ' EinlAb = 0 OR altPat >= EinlAb THEN
        frm.Ausgeb lfdnr & ") Pat_id: " & altpat & " nicht berücksichtigt!", True
       End If ' EinlAb = 0 OR altPat >= EinlAb THEN
      Else ' obVorspannVorbei
       obVorspannVorbei = obVorspannVorbei + 1 ' wenn neuer Patient und Vorspann bisher nicht vorbei, dann laß ihn vorbei sein
      End If ' obVorspannVorbei
       
      DoEvents
'     IF obVorspannVorbei > 0 THEN
      If Not obSchluss Then
       cInha(0) = cInha(ind)
       altpat = cInha(ind)
       cKenn(0) = cKenn(ind)
      End If
      lAktZeit = 0
      ind = 0
'     END IF
      obAndPat = 0
      If BrichAb Then
       Exit Do
      End If
      frm.Ausgeb "Lese Daten von Datei ...", False
     End If ' obAndPat OR obschluss
     ind = ind + 1
     
     If obSchluss Then
      Exit Do
     End If
     ltxtS = tx1S
     If Not obEOF Then Exit Do
     obSchluss = True ' dann aufhören
    Loop ' kleine Schleife, Lesen einzelner Zeilen
    If LenB(altpat) = 0 Then altpat = -1
    If EinlBis <> 0 Then
     If altpat > EinlBis Then Exit Do
    End If
    If frm.dlg.obVglMitLetzterEinlesung <> 0 And Not obHausBesuch Then If obSchluss Then Exit Do
    If BrichAb Then
     Exit Do
    End If
   End If ' frm.dlg.obVglMitLetzterEinlesung = 0 OR obHausBesuch OR Steuer <> D2Weiter THEN
   
   If frm.dlg.obVglMitLetzterEinlesung <> 0 And Not obHausBesuch Then ' wenn Vergleich mit letzter Einlesung
    If steuer <> D1Weiter Then ' wenn nicht D1 weiter
     If EOF(318) Then      ' Wenn Kopie schon zuende
      If i2s <> CStr(i2max) Then   ' und i2 noch nicht i2max 6.11.10
       K2S = "3000"         ' dann setzen
       i2s = i2max
      End If
     Else
      On Error GoTo ef318
      Line Input #318, zwi ' wenn noch nicht, dann Zeile einlesen
      On Error GoTo fehler
      tx2S = zwi
      K2S = tx2S.Mid(4, 4)
      i2s = tx2S.Mid(8)
     End If
     If K2S = "3000" Then
      If Pat2 <> i2s Then
       If EOF(318) Then
        obD2neu = True ' 2.11.07: Progendelosschleife bei letztem Pat. in alter Datei länger-> hier wohl stop, um Fehler zu korrigieren
       Else
        obD2neu = True       ' wenn neuer Patient auf D2
       End If
       DoEvents
      End If
     End If ' K2 = "3000" THEN
    End If ' Steuer <> D1Weiter THEN
    If Not obD1neu And Not obD2neu Then
     If K1S = K2S And i1s = i2s Then
     Else
      If steuer = beidemvgl Then
       If LenB(Pat1) <> 0 Then Call pCol.Add(Pat1)
       neuSteuer = beideovgl
      ElseIf steuer = d2weiter Then
       If EOF(318) Then
        neuSteuer = beidemvgl
       End If
      End If
     End If
    ElseIf obD1neu And obD2neu Then
     If i1s = i2s Then
      neuSteuer = beidemvgl
     ElseIf i1s < i2s Then
      Call pCol.Add(i1s)
      neuSteuer = D1Weiter
     Else
      neuSteuer = d2weiter
     End If
    ElseIf obD1neu And Not obD2neu Then
     If steuer = beidemvgl Then
      Call pCol.Add(Pat1)
      neuSteuer = d2weiter
     ElseIf steuer = beideovgl Then
      neuSteuer = d2weiter
     ElseIf steuer = D1Weiter Then
      If i1s = Pat2 Then
       neuSteuer = beidemvgl
      ElseIf i1s < Pat2 Then
       Call pCol.Add(i1s)
       neuSteuer = D1Weiter
      ElseIf i1s > Pat2 Then
       neuSteuer = d2weiter
      End If
     End If
    Else ' IF Not obD1neu AND obD2neu THEN
     If steuer = beidemvgl Then
      Call pCol.Add(Pat1)
      neuSteuer = D1Weiter
     ElseIf steuer = beideovgl Then
      neuSteuer = D1Weiter
     ElseIf steuer = d2weiter Then
      If Pat1 = i2s Then
       neuSteuer = beidemvgl
      ElseIf Pat1 < i2s Then
       Call pCol.Add(Pat1)
       neuSteuer = D1Weiter
      ElseIf i1s > Pat2 Then
       neuSteuer = d2weiter
      End If
     End If ' steuer = beidemvgl THEN elseif elseif
    End If 'Not obD1neu AND NOT obD2neu THEN
    If obD1neu Then Pat1 = i1s
    If obD2neu Then Pat2 = i2s
    steuer = neuSteuer
   End If ' frm.dlg.obVglMitLetzterEinlesung <> 0 AND NOT obHausBesuch THEN
  Loop 'While Not EOF(324) ' Große Einleseschleife
  frm.Ausgabe = altAusgabe
  T1a = Now
'   IF Not mitLab THEN MsgBox (t1a - t1) * 60 * 60 * 24 ' dauert hingegen nur 9" für die ganze Datei und hat den richtigen Zeichensatz, mit Berechnung von LAktZeit 32"
 
  Call Eintragszl(2, fallzahl, (Now - T1) * 60 * 60 * 24) ' Eintragsschleife verlassen
  If obSchluss Then
   Call Eintragszl(3) ' ... und zwar an der vorgesehenen Stelle
  End If
 
  Call medklass2
#If ohnebezug Then
  Call ForeignYes0
  Call ForeignYes1
#End If

  #If False Then
   Call ProgrammLauf(-1)
  #End If
'abbrech:
  If frm.dlg.ÜberTabelle <> 0 Then
   Set rsinl = Nothing
  Else
   Close #324 ' BDT-Datei
   obEOF = True
  End If
 End If ' frm.dlg.obBDT THEN
 
 Call Eintragszl(4) ' nach Medklass
 If Not BrichAb And (FSO.FileExists(frm.dlg.LaborPfadBeispiel) Or FSO.FileExists(frm.dlg.LaborPfadBeispiel)) Then
  If obLaborDirekt Then Call LaborDirektImport(frm, absPos, frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern, frm.dlg.LaborPfadBeispiel, obLDneu)
  Call Eintragszl(5)
  If obLaborQuer Then Call LaborErgPatId(frm, obLQneu)
  Call Eintragszl(6)
  If obEmails Then Call EmailsImport(EmDatei, frm)
  Call Eintragszl(7)
' Folgendes aus vermuteten Performancegründen hier statt beim einzeln Speichern
  If frm.dlg.obBDT Then
   frm.Ausgeb "Ergänze Namen im Anmnesebogen ...", False
   If LenB(sqliif) = 0 Then Zinit (LVobMySQL)
   myEFrag "UPDATE `anamnesebogen` a INNER JOIN `namen` n ON a.pat_id = n.pat_id  SET a.nachname = n.nachname, a.vorname = n.vorname, a.anrede = " & sqliif & "(n.geschlecht='m','Herr','Frau'), a.titel = n.titel, a.gebdat = n.gebdat WHERE (a.nachname = '' OR ISNULL(a.nachname))", rAF
   If frm.dlg.bereinigeFormInhFeld <> 0 Then
    ' 11.10.20: als contab-job ausgelagert
'    frm.Ausgeb "Lösche waise Sätze aus forminhfeld ...", False
'    IF lies.obMySQL THEN ' wieder eingesetzt am 7.11.09
'     myEFrag "DELETE FROM `forminhfeld` WHERE NOT EXISTS (SELECT foid FROM `forminhkopf` WHERE foid = `forminhfeld`.foid)", rAF
'    Else
'     IF lies.obMySQL THEN
'      myEFrag "DROP TABLE IF EXISTS `forminhfeld2`", rAF
'     Else
'      ON Error Resume Next
'      myEFrag "DROP TABLE `forminhfeld2`", rAF
'      ON Error GoTo fehler
'     END IF
'     Dim VorZahl&, GelZahl&
'     myEFrag "CREATE TABLE `forminhfeld2` LIKE `forminhfeld`", rAF
'     rs.Open "SELECT COUNT(0) ct FROM `forminhfeld`", DBCn
'     VorZahl = rs!ct
'     myEFrag "INSERT INTO `forminhfeld2` SELECT fif.* FROM `forminhkopf` fik LEFT JOIN `forminhfeld` fif ON fif.foid = fik.foid", rAF
'     GelZahl = VorZahl - rAF
'     myEFrag "SET foreign_key_checks=0"
'     myEFrag "DROP TABLE `forminhfeld`"
'     myEFrag "ALTER TABLE `forminhfeld2` rename `forminhfeld`"
'     myEFrag "SET foreign_key_checks=1"
'     SET rs = Nothing
'    END IF ' frm.dlg.bereinigeFormInhFeld <> 0 THEN
'    frm.Ausgeb rAF & " waise Sätze aus `forminhfeld` gelöscht", True
   End If
  End If ' frm.dlg.obBDT THEN
'  IF Not lies.obMySQL THEN
'   IF rs.State = 1 THEN rs.Close
'   SET rs = Nothing
'  Call kompakt(frm)
'  END IF
  If obLaborDirekt Then  ' 1.11.14, zur Verkürzung des Einzelimports
   Call Lese.FalschAbgehakteUngueltig_Click ' 11.7.09
  End If
  Call Eintragszl(8)
 End If ' Not BrichAb THEN
 
 If obÜProt Then Close #322
 frm.Ausgeb "Zeitdauer gesamt: " & Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0"), True
 frm.Ausgeb "Zeitdauer nach Haupteinlesen: " & Format$((Now - T1a) * 60 * 60 * 24, "###,###,###,###,###,###,##0"), True
 If GesBytes > 500000000 Then
  Call fzsfuell(frm, 999, -1)
  Call fzsfuell(frm, 999, 0)
'  MsgBox "Rufe jetzt Fallzahlstand auf"
  Call dofallzahlstand(frm, IIf(Hour(Now()) < 9, 1, 0))
 End If
 Call ControlsStop(frm)
 Exit Function
ef318:
 spos = Seek(318)
 Close #318
 Open frm.dlg.LDatei For Input As #318
 Seek #318, spos
 Resume
ef324:
 spos = Seek(324)
 Close #324
 Open frm.dlg.LDatei For Input As #324
 Seek #324, spos
 Resume
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then
 If DBCn.State = 0 Then
  DBCn.Open
  Resume
 End If
End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + frm.dlg.LDatei, vbAbortRetryIgnore, "Aufgefangener Fehler in GesLies/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GesLies

Public Function ControlsStop(frm As Lese)
 Dim i&
 For i = 0 To frm.Controls.COUNT - 1
  If frm.Controls(i).name = "Beenden" Or frm.Controls(i).name = "Zurücksetzen" Then
   frm.Controls(i).Enabled = True
  ElseIf frm.Controls(i).name = "Abbrechen" Then
   frm.Controls(i).Enabled = False
  End If
 Next i
End Function ' ControlsStop

Public Function controlsLauf(frm As Lese)
 Dim i&
 For i = 0 To frm.Controls.COUNT - 1
  If frm.Controls(i).name = "Beenden" Or frm.Controls(i).name = "Zurücksetzen" Then
   frm.Controls(i).Enabled = False
  ElseIf frm.Controls(i).name = "Abbrechen" Then
   frm.Controls(i).Enabled = True
  End If
 Next i
End Function 'ControlsLauf

Function GetMed(lang$, Einrück%) As CString ' mit Einrück=0 in doMedklassT, THAfstleg, dolies, MedArtenPruef, alleSpeichern, therauskunft, mit 3 auch in: allespeichern
'  medi.Med = Trim$(UCase$(LEFT(rMe(i).Medikament, IIf(InStrB(rMe(i).Medikament, " ") > 0, InStr(rMe(i).Medikament, " ") - 1, Len(rMe(i).Medikament)))))
'  medi.Med = Trim$(UCase$(LEFT(rRe(i).Medikament, IIf(InStrB(Mid$(rRe(i).Medikament, 4), " ") > 0, InStr(Mid$(rRe(i).Medikament, 4), " ") - 1 + 4, Len(rRe(i).Medikament)))))
  Dim links$, lzpos&
  Set GetMed = New CString
  GetMed = lang
'  IF GetMed.Instr("Humalog® ") <> 0 THEN Stop
'  IF GetMed = "Humalog       ca." THEN Stop
'  IF GetMed = "Berlinsulin Basal" THEN Stop
'  GetMed.UCase
  GetMed = UCase$(GetMed) ' .ucase lässt Umlaute aus, 26.3.22
  lzpos = GetMed.Instr(" ", 1 + Einrück)
  If lzpos > 0 Then
   links = Left$(GetMed, lzpos - 1)
   Select Case links
    Case "ACCU", "ACCU-CHEK", "HUMALOG", "LIPROLOG", "INSUMAN", "BERLINSULIN", "HUMINSULIN", "INSULIN" ' ,"INSULIN RATIO","INSULIN HUMINSULIN","INSULIN HM","INSULIN BRAUN RATIO","INSULIN BBM RATIO","INSULIN B.BRAUN","INSULIN B.","INSULIN B"
     lzpos = GetMed.Instr(" ", 1 + MAX(Len(links) + 1, Einrück))
     While GetMed.Mid(lzpos, 1) = " " And lzpos <= GetMed.length: lzpos = lzpos + 1: Wend
'     links = Left$(GetMed, lzpos - 1)
   End Select
   If links = "ACCU" Or InStr(links, "®") <> 0 Then
     lzpos = GetMed.Instr(" ", 1 + MAX(lzpos, Einrück))
     While GetMed.Mid(lzpos, 1) = " " And lzpos <= GetMed.length: lzpos = lzpos + 1: Wend
   End If
   If lzpos > 0 Then
    GetMed.Cut (lzpos - 1)
   End If
   GetMed = Trim$(GetMed)
  End If
End Function ' GetMed(Lang$, Einrück%) AS CString

Function doMedklassT() ' kommt vor in MedklassT_Click
 Dim rMe As New ADODB.Recordset
' Dim medSL As New SortierListe
 Dim medi As MediCl
 Dim mZ&, iru&
 Set medSL = New SortierListe
 Lese.Ausgeb "Beginne mit Einlesen für doMedklassT", True, True
 mZ = myEFrag("SELECT COUNT(0) FROM (SELECT 0 FROM medplan GROUP BY medanfang) i").Fields(0)
' rMe.Open "SELECT mp.Pat_id,mp.Medikament,Medanfang FROM `medplan` mp LEFT JOIN medarten ma ON mp.MedAnfang=ma.Medikament WHERE ISNULL(ma.medikament) GROUP BY medanfang ORDER BY medikament,pat_id", DBCn, adOpenDynamic, adLockReadOnly
 myFrag rMe, "SELECT mp.Pat_id,mp.Medikament,Medanfang FROM `medplan` mp LEFT JOIN medarten ma ON mp.MedAnfang=ma.Medikament WHERE ISNULL(ma.medikament) GROUP BY medanfang ORDER BY medikament,pat_id"
' rMe.Open "SELECT Pat_id,Medikament,Medanfang FROM `medplan` GROUP BY medanfang ORDER BY medikament,pat_id", DBCn, adOpenDynamic, adLockReadOnly
 Lese.Ausgeb "Beginne mit Do while in doMedklassT, Datensatzzahl: " & CStr(mZ), True, True
 Do While Not rMe.EOF
  Set medi = New MediCl
'  medi.Med = Trim$(UCase$(LEFT(rMe!Medikament, IIf(InStrB(rMe!Medikament, " ") > 0, InStr(rMe!Medikament, " ") - 1, Len(rMe!Medikament)))))
  medi.Med = GetMed(rMe!Medikament, 0)
  iru = iru + 1
  If iru Mod 1000 = 0 Then
   Lese.Ausgeb "Do while in doMedklassT:" & CStr(iru) & "/" & CStr(mZ), True, True
  End If
#Const medikorr = 1
#If medikorr Then
  If rMe!MedAnfang <> medi.Med Then
   myEFrag "UPDATE `medplan` SET medanfang = UPPER('" & medi.Med & "') WHERE medikament = '" & rMe!Medikament & "'", rAF
  End If
#End If
  medi.LT = rMe!Medikament
  medi.Pat_id = rMe!Pat_id
  Call medSL.sCAdd(medi, True)
  DoEvents
  rMe.Move 1
 Loop
 Lese.Ausgeb "Fertig mit Einlesen für doMedklassT", True, True
 Call medklass2
 Lese.Ausgeb "Fertig mit doMedklassT", True, True
End Function ' doMedklassT()

 Function medklass2() ' kommt vor in GesLies(), doMedklassT()
' Medklass, 2. Teil
 Dim Med$, i&
 Dim rs As ADODB.Recordset
 On Error Resume Next
 DBCn.BeginTrans: obTrans = 1
 On Error GoTo fehler
' Call ForeignNo
  For i = 1 To medSL.COUNT
  Med = medSL.Item(i).Med
  If LenB(Med) <> 0 And Med <> """" Then
   Set rs = myEFrag("SELECT 0 FROM `medarten` WHERE medikament = '" & REPLACE$(REPLACE$(Med, "'", "\\\'"), "\\\\'", "\\\'") & "'")
   If rs.BOF Then
    InsKorr DBCn, DBCnS, "INSERT INTO `medarten`(medikament,hinzugefügt,langname,pat_id) values ('" & Med & "'," & DatFor_k(Now) & ",'" & medSL.Item(i).LT & "'," & medSL.Item(i).Pat_id & ")", rAF
   Else
    Set rs = myEFrag("SELECT 0 FROM `medarten` WHERE medikament = '" & Med & "' AND (pat_id = 0 OR ISNULL(pat_id))")
    If Not rs.BOF Then
     Call myEFrag("UPDATE `medarten` SET pat_id = " & medSL.Item(i).Pat_id & " WHERE medikament = '" & Med & "'", rAF)
    End If
' , Langname = '" & medSL.Item(i).LT & "'
    Set rs = myEFrag("SELECT 0 FROM `medarten` WHERE medikament = '" & Med & "' AND (langname = '' OR ISNULL(langname))")
    If Not rs.BOF Then
     Call myEFrag("UPDATE `medarten` SET LangName = '" & medSL.Item(i).LT & "' WHERE medikament = '" & Med & "'", rAF)
    End If
    Call myEFrag("UPDATE `medarten` SET anzahl = " & medSL.Item(i).mvPatZahl & " WHERE medikament = '" & Med & "'", rAF)
   End If
  End If
 Next i
' Call ForeignYes
 On Error Resume Next
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medklass2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Medklass2

Function EintragZusatz()
 Dim rAF&, erg$
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, dae As Date, sql$
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 Call Lese.ProgStart
 myFrag rs, "SELECT datei,dateiaend,beginn FROM `eintragszahlen` WHERE NOT ISNULL(datei) AND datei<>'' ORDER BY beginn DESC"
 Do While Not rs.EOF
   On Error Resume Next
   erg = Dir(rs!Datei)
   On Error GoTo fehler
   If LenB(erg) <> 0 Then
    myFrag rs1, "SELECT datei,dateiaend, beginn FROM `eintragszahlen` WHERE datei = '" & IIf(LVobMySQL, REPLACE$(rs!Datei, "\", "\\"), rs!Datei) & "' ORDER BY beginn DESC"
    If Not rs1.BOF Then
     dae = FileDateTime(rs!Datei) '  FSO.GetFile(rs!Datei).DateLastModified
     If rs1!DateiAend <> dae Then
 '    sql = "UPDATE " & IIf(LVobMySQL, vns, " top 1 ") & "`" & DefDB(dbcn) & "`.`eintragszahlen` SET DateiAend = " & DatFor_k(dae) & " WHERE Datei = '" & umwfSQL(rs!Datei) & "'" & IIf(LVobMySQL, " LIMIT 1", "")
      sql = "UPDATE `eintragszahlen` SET DateiAend = " & DatFor_k(dae) & " WHERE beginn = " & DatFor_k(rs1!Beginn)
      Call myEFrag(sql, rAF)
      If rAF <> 1 Then
       Tüt 1200, 1000
       MsgBox "Fehler bei `eintragszahlen`.beginn = " & rs1!Beginn & ", dae = " & dae & "(rAF=" & rAF & " statt 1 bei update"
       Debug.Print rs1!Beginn, dae
       Exit Function
      End If ' rAF <> 1
     End If ' rs1!DateiAend
    End If ' Not rs1.BOF THEN
    Set rs1 = Nothing ' eins weiter oben geht nur bei Access
   End If ' LenB(erg) <> 0
  rs.Move 1
 Loop
 Screen.MousePointer = vbNormal
' MsgBox "Fertig mit Eintragszusatz"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EintragZusatz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EintragZusatz

Function testload()
 Dim rs As New ADODB.Recordset, rdn As New ADODB.Recordset, erg$
 Dim T1!, T2!
 Call Lese.ProgStart
 Call myFrag(rdn, "SELECT stbyte, datei FROM `namen` LEFT JOIN `eintragszahlen` USING (stbyte) GROUP BY stbyte", DBCn)
 Call getAktByte
 Do While Not rdn.EOF
  erg = Dir(rdn!Datei)
  If LenB(erg) <> 0 Then
'   myEFrag "LOAD DATA INFILE '/DATA/eigene Dateien/TMExport/" & erg & "' INTO TABLE `inlk` fields terminated by '' enclosed by '' lines terminated by '\r\n'  (breite,kennung,inhalt) SET StByte = " & AktByte & ", pat_id = IF(kennung=3000,inhalt,0)", rAF
   T1 = Timer
   myEFrag "LOAD DATA INFILE '/DATA/eigene Dateien/TMExport/" & erg & "' INTO TABLE `inlk` fields terminated by '' enclosed by '' lines terminated by '\r\n'  (breite,kennung,inhalt) SET StByte = " & AktByte, rAF
'  myEFrag "LOAD DATA INFILE '/DATA/eigene Dateien/TMExport/Temiz x0119153.BDT' INTO TABLE `inlk` fields terminated by '' enclosed by '' lines terminated by '\r\n'  (breite,kennung,inhalt) SET StByte = " & AktByte & ", pat_id = IF(kennung=3000,inhalt,0)", rAF
   T2 = Timer
   DoEvents
   Debug.Print (T2 - T1) & " 's' für Datei " & erg
   DoEvents
  End If
  rdn.Move 1
 Loop
End Function ' testload

Function getAktByte()
 Dim rs As New ADODB.Recordset
 On Error GoTo fehler
 Set rs = myEFrag("SELECT stbyte FROM `eintragszahlen` ORDER BY stbyte DESC")
 Call Lese.ProgStart
 If rs.BOF Then
  VorByte = 0
 Else
  VorByte = rs.Fields(0)
 End If
 AktByte = VorByte + 1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getAktByte/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getAktyte

Function EintragStart&(frm As Lese) ' wird in EinlesenClick aufgerufen
 On Error GoTo fehler
 Dim sql$, DateiAend#
 Dim rs As ADODB.Recordset
 rEzäBeg = Now
 If DBCn.State = 0 Then
  DBCnOpen
 End If
 Call getAktByte
 If FSO.FileExists(frm.QDatei) Then
  DateiAend = FSO.GetFile(frm.QDatei).DateLastModified
 End If
 sql = "INSERT INTO `eintragszahlen` (beginn,stbyte,datei,dateiaend,TabellenEntleeren,ZurücksetzenLAktDat,obVglMitLetzterEinlesung,Pat_IDVon,Pat_IDbis,VorladenFFI,Übertabelle,NurInTabelle,SammelInsert,bereinigeFormInhFeld, LaborDirektEinlesen,LaborDirektNeu,LaborQuerVerb,LaborQuerNeu,LaborPfadBeispiel,AlterTab,obmitEmails) " & _
       "values (" & DatFor_k(rEzäBeg) & "," & AktByte & ",'" & doUmwfSQL(frm.QDatei, lies.obMySQL) & "'," & DatFor_k(DateiAend) & "," & Abs(frm.dlg.TabellenEntleeren) & "," & Abs(frm.dlg.ZurücksetzenLAktDat) & "," & Abs(frm.dlg.obVglMitLetzterEinlesung) & ",'" & frm.dlg.Pat_IDVon & "','" & frm.dlg.Pat_IDBis & "'," & Abs(frm.dlg.VorladenFFI) & "," & Abs(frm.dlg.ÜberTabelle) & "," & Abs(frm.dlg.NurInTabelle) & "," & Abs(frm.dlg.SammelInsert) & "," & Abs(frm.dlg.bereinigeFormInhFeld) & "," & Abs(frm.dlg.LaborDirektEinlesen) & "," & Abs(frm.dlg.LaborDirektNeu) & "," & Abs(frm.dlg.LaborQuerVerb) & "," & Abs(frm.dlg.LaborQuerNeu) & ",'" & doUmwfSQL(frm.dlg.LaborPfadBeispiel, lies.obMySQL) & "'," & Abs(frm.dlg.AlterTab) & "," & Abs(frm.dlg.obmitEmails) & ")"
 InsKorr DBCn, DBCnS, sql, rAF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EintragStart/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EintragStart()

' aufgerufen öfter in geslies
Function Eintragszl(nr%, Optional fallzahl&, Optional Sekunden&)
 Dim sql$
 On Error GoTo fehler
 Lese.Ausgeb "Eintragszahlen eintragen ...", False
 If fallzahl = 0 And Sekunden = 0 Then
  sql = "UPDATE `eintragszahlen` SET zp" & nr & " = " + DatFor_k(Now) + " WHERE beginn = " + DatFor_k(rEzäBeg)
 Else
  sql = "UPDATE `eintragszahlen` SET zp" & nr & " = " & DatFor_k(Now) & ", FallZahl = " & fallzahl & ", Sekunden = " & Sekunden & " WHERE beginn = " + DatFor_k(rEzäBeg)
 End If
' Call Lese.ProgStart
' Call myEFrag(sql, numA)
 myEFrag sql, numA
 If numA <> 1 Then Err.Raise 999, , "Fehler in Eintragszl: numA <> 1, und zwar: " & numA
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
ElseIf InStrB(Err.Description, "gone away") <> 0 Then
 Call Lese.ProgStart
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Eintrag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Eintragszl

' aufgerufen in geslies
Function StatischInit()
 Quartal = vNS ' 4101
 lfdnr = 0
 lKennung = 0 ', lk%(15)
 jetztKopf = 0
 rFm_Nr = 0
 lFormID = 0  ' ID des letzten Formulars
 obZweiteRunde = 0
 lLang = "diesen schräcklichen Langtext gibt es garantiert nicht"
 lLangVW = 0
 lKomm = "Düsen Kommentar gübt es nücht" ' 21.8.21, sonst scheitert kommeinfüg, sodann wegen foreign-key-Fehler gar kein Eintrag in laborneu mehr!
 lKommVW = 0
 SpeicherZt = 0
 AktQ = ZQuart(Now - Verspätung) ' definiert in ZielDBFunktionen
End Function ' StatischInit

Public Function medartenhier(SL As SortierListe)
 Dim rs As New ADODB.Recordset, SM As SortierMedi, T1!, T2!
 On Error GoTo fehler
 T1 = Timer
 Set mSL = New SortierListe
 myFrag rs, "SELECT medikament,glib, metf, gluci, shglin, glit, dpp4, glp1, sglt2, sonstad, insart, ins, anal, puzu FROM `medarten` WHERE glib<>0 OR metf<>0 OR gluci<>0 OR shglin<>0 OR glit<>0 OR dpp4<>0 OR glp1<>0 OR sglt2<>0 OR sonstad<>0 OR insart<>0 OR ins<>0 OR anal<>0 OR puzu<>0"
 Do While Not rs.EOF
  Set SM = New SortierMedi
  Set SM.medi = rs!Medikament
  SM.anal = rs!anal
  SM.glib = rs!glib
  SM.glit = rs!glit
  SM.gluci = rs!gluci
  SM.InS = rs!InS
  If Not IsNull(rs!insart) Then If IsNumeric(rs!insart) Then SM.insart = rs!insart
  SM.metf = rs!metf
  SM.puzu = rs!puzu
  SM.shglin = rs!shglin
  SM.dpp4 = rs!dpp4
  SM.glp1 = rs!glp1
  SM.sglt2 = rs!sglt2
  SM.sonstad = rs!sonstad
  SL.sCAdd SM
  rs.MoveNext
 Loop
 T2 = Timer
 Debug.Print T2 - T1 & " Sekunden"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medartenhier/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' medartenhier

Function thtest()
 Lese.ProgStart
 Call rufThFestleg(10)
End Function ' thtest

' in TherapieartenFestlegen_Click(), thtest()
Function rufThFestleg(Pat_id&)
 syscmd acSysCmdSetStatus, "ermittle und speichere Therapiearten für Pat. " & Pat_id
#If Not thaalt Then
 'RunCommandLine ("ssh root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & CStr(Pat_id) & ")'")
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
#If mitfenster Then
 rufauf "ssh", "root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & CStr(Pat_id) & ")'", 2, "c:\windows\system32\openssh\", -1, 0
#Else
 Call TheraErmitt(Pat_id)
#End If
' myEFrag "CALL fuellThaP(" & CStr(Pat_id) & ")"
#Else ' ausrangiert 12.12.20
 Dim altfoid&
 Dim rsNa As New ADODB.Recordset
 Dim rsFm As New ADODB.Recordset
 Dim rsMe As New ADODB.Recordset
 ReDim rTh(0)
 ReDim rNa(0)
 ReDim rFm(0)
 ReDim rMe(0)
 rNa(0).Pat_id = Pat_id
 myFrag rsNa, "SELECT aufndat FROM `namen` WHERE pat_id = " & Pat_id
 If Not rsNa.BOF Then rNa(0).AufnDat = rsNa!AufnDat
 myFrag rsFm, "SELECT foid,feldinh,feld,zeitpunkt FROM `formular` WHERE pat_id = " & Pat_id & " AND feld IN ('medikament','txtmedkey') ORDER BY foid, feldnr"
 If Not rsFm.BOF Then
  Do While Not rsFm.EOF()
   ReDim Preserve rFm(UBound(rFm) + 1)
   rFm(UBound(rFm)).Foid = rsFm!Foid
   rFm(UBound(rFm)).FeldInh = rsFm!FeldInh
   rFm(UBound(rFm)).Feld = rsFm!Feld
   If rsFm!Foid <> altfoid Then
    If (0 / 1) + (Not Not rFr) = 0 Then ReDim rFr(1) Else ReDim Preserve rFr(UBound(rFr) + 1)
    rFr(UBound(rFr)).Zeitpunkt = rsFm!Zeitpunkt
    rFr(UBound(rFr)).Foid = rsFm!Foid
    altfoid = rsFm!Foid
   End If
   rsFm.MoveNext
  Loop
 End If
 myFrag rsMe, "SELECT MPNr, Zeitpunkt, AktZeit, Medikament, Mo, Mi, NM, Ab, ZN, StByte FROM `medplan` WHERE pat_id = " & Pat_id & " ORDER BY zeitpunkt, mpnr"
 If Not rsMe.BOF Then
  Do While Not rsMe.EOF()
   ReDim Preserve rMe(UBound(rMe) + 1)
   rMe(UBound(rMe)).MPNr = rsMe!MPNr
   rMe(UBound(rMe)).Zeitpunkt = rsMe!Zeitpunkt
   rMe(UBound(rMe)).Medikament = rsMe!Medikament
   rMe(UBound(rMe)).MedAnfang = GetMed(rsMe!Medikament, 0).Value ' 17.6.17
   rMe(UBound(rMe)).mo = rsMe!mo
   rMe(UBound(rMe)).mi = rsMe!mi
   rMe(UBound(rMe)).nm = rsMe!nm
   rMe(UBound(rMe)).ab = rsMe!ab
   rMe(UBound(rMe)).Zn = rsMe!Zn
   rMe(UBound(rMe)).AktZeit = rsMe!AktZeit
   AktByte = rsMe!StByte
   rsMe.MoveNext
  Loop
 End If
 If mSL Is Nothing Then Call medartenhier(mSL)
 Call THAfestleg
 Call therartenSpeichern(0, 0)
#End If
End Function ' rufThFestleg(pat_id&)

#If thaalt Then ' ausrangiert 12.12.20
' aufgerufen in GesLies, benötigt nur für ThaFestleg, befüllt rmpn mit einer nach Zeitpunkt sortierten Version von rMe
Function rmeSort()
 ReDim rmpn(0)
 Dim ldat As Date, adat As Date
 Dim i&
 On Error GoTo fehler
 Do
  adat = 0
  For i = 1 To UBound(rMe)
   If rMe(i).Zeitpunkt > ldat And (adat = 0 Or rMe(i).Zeitpunkt < adat) Then adat = rMe(i).Zeitpunkt
  Next i
  If adat = 0 Then Exit Do
  For i = 1 To UBound(rMe)
   If rMe(i).Zeitpunkt = adat Then
    ReDim Preserve rmpn(UBound(rmpn) + 1)
    rmpn(UBound(rmpn)).FID = rMe(i).FID
    rmpn(UBound(rmpn)).Pat_id = rMe(i).Pat_id
    rmpn(UBound(rmpn)).MPNr = rMe(i).MPNr
    rmpn(UBound(rmpn)).Zeitpunkt = rMe(i).Zeitpunkt
    rmpn(UBound(rmpn)).Datum = rMe(i).Datum
    rmpn(UBound(rmpn)).Medikament = rMe(i).Medikament
    rmpn(UBound(rmpn)).MedAnfang = rMe(i).MedAnfang
    rmpn(UBound(rmpn)).Wirkstoff = rMe(i).Wirkstoff
    rmpn(UBound(rmpn)).FeldNr = rMe(i).FeldNr
    rmpn(UBound(rmpn)).mo = rMe(i).mo
    rmpn(UBound(rmpn)).mi = rMe(i).mi
    rmpn(UBound(rmpn)).nm = rMe(i).nm
    rmpn(UBound(rmpn)).ab = rMe(i).ab
    rmpn(UBound(rmpn)).Zn = rMe(i).Zn
    rmpn(UBound(rmpn)).bBed = rMe(i).bBed
    rmpn(UBound(rmpn)).Bemerkung = rMe(i).Bemerkung
    rmpn(UBound(rmpn)).Grund = rMe(i).Grund
    rmpn(UBound(rmpn)).Stärke = rMe(i).Stärke
    rmpn(UBound(rmpn)).Einheit = rMe(i).Einheit
    rmpn(UBound(rmpn)).Form = rMe(i).Form
    rmpn(UBound(rmpn)).absPos = rMe(i).absPos
    rmpn(UBound(rmpn)).AktZeit = rMe(i).AktZeit
    rmpn(UBound(rmpn)).StByte = rMe(i).StByte
   End If
  Next i
  ldat = adat
 Loop
 ' nach Zeitpunkt sortieren sortierdatum
schluß:
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rmeSort/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rmeSort

' ausrangiert 12.12.20
' s. therart_erm
Function THAfestleg() ' in dolies und geslies / parallel zu Therauskunft
 Dim i&, Tha$, altTha$, insart%, altinsart%, mEl As New SortierMedi, obschluß%, Grund$
 Dim altMPNr%, altZp As Date
 Dim glp1z%, oadz%, moins1%, miins1%, nmins1%, abins1%, znins1%, insz%, puzuz%, RpCSII% ' für Therapiearten
 Dim moins3%, miins3%, nmins3%, abins3%, znins3%
 Dim sFK As SortierFormularKopf, fkSL As New SortierListe
 Dim SD As SortierDatum, dSL As New SortierListe
 Dim nachhak% ' ob Durchlauf nach dem Medikamentenplan, bzw. ob dann tatsächlich noch neues Pumpenzubehör gefunden
 On Error GoTo fehler
 
 If ((0 / 1) + Not Not rFr) <> 0 Then
  For i = 1 To UBound(rFr)
   Set sFK = New SortierFormularKopf
   sFK.Zeitpunkt = rFr(i).Zeitpunkt
   sFK.Foid = rFr(i).Foid
   fkSL.sCAdd sFK
  Next i
 End If
 If ((0 / 1) + Not Not rFm) <> 0 Then
 For i = 1 To UBound(rFm)
  Select Case LCase$(rFm(i).Feld)
   Case "medikament", "txtMedKey"
    Dim ifi$
    ifi = LCase$(rFm(i).FeldInh)
'    IF rFm(i).Feld <> "medikament" THEN
'     MsgBox "Fehler in THAfestleg: mFm(i).Feld <> ""medikament"" "
'     Stop
'    END IF
'    Debug.Print ifi
Debug.Print ifi
    If (InStrB(ifi, "reservoir") <> 0 Or InStrB(ifi, "rapid d link") <> 0 Or InStrB(ifi, "rap d li") <> 0 Or InStrB(ifi, "rapid-d li") <> 0 Or InStrB(ifi, "tenderl") <> 0 Or InStrB(ifi, "flexl") <> 0 Or InStrB(ifi, "check spirit") <> 0 Or InStrB(ifi, "insight") <> 0 Or InStrB(ifi, "chek spirit") <> 0 Or InStrB(ifi, "pumpenträg") <> 0 Or InStrB(ifi, "kunststoffampu") <> 0 Or InStrB(ifi, "spritzampull") <> 0 Or InStrB(ifi, "batteriefachdeckel") <> 0 Or InStrB(ifi, "h-tron") <> 0 Or InStrB(ifi, "d-tron") <> 0 Or InStrB(ifi, "paradigm") <> 0 Or InStrB(ifi, "csii") <> 0 Or InStrB(ifi, "linpumpe") <> 0 Or InStrB(ifi, "omnipod") <> 0 Or InStrB(ifi, "minimed") <> 0 Or InStrB(ifi, "640g") <> 0 Or InStrB(ifi, "carelink") <> 0 Or InStrB(ifi, "mio ") <> 0 Or InStrB(ifi, "quickset") <> 0 Or InStrB(ifi, "silhouette") <> 0 Or InStrB(ifi, "sure-t") <> 0 _
    Or InStrB(ifi, "sure t") <> 0 Or InStrB(ifi, "paradigm") <> 0 Or InStrB(ifi, " veo") <> 0 Or InStrB(ifi, "animas") <> 0 Or InStrB(ifi, "caridge") <> 0) And InStrB(FeldInh, "%menveo%") = 0 Then
     Set SD = New SortierDatum
     Set sFK = New SortierFormularKopf
     sFK.Foid = rFm(i).Foid
     Set sFK = fkSL.GetItem(sFK)
     If Not sFK Is Nothing Then
      SD.Datum = sFK.Zeitpunkt 'Int(sFK.Zeitpunkt)
      SD.Text = ifi
      If dSL.GetItem(SD) Is Nothing Then
       dSL.Add SD
'      Else
'       Stop
      End If
     End If
     RpCSII = RpCSII + 1
    End If
  End Select
 Next i
 End If
'      SET sD = New SortierDatum
'      sD.Datum = Int(NOW())
'      dSL.sCInit
'      dSL.Eingrenz sD
'      SET sD = New SortierDatum
'      sD.Datum = Int(CDate("1.5.2010"))
'      dSL.sCInit
'      dSL.Eingrenz sD
'      SET sD = New SortierDatum
'      sD.Datum = Int(CDate("1.11.2009"))
'      dSL.sCInit
'      dSL.Eingrenz sD
 Dim rsAna As New ADODB.Recordset
 myFrag rsAna, "SELECT `insulinpumpe` ip FROM `anamnesebogen` a WHERE pat_id = " & rNa(0).Pat_id
 If Not rsAna.EOF Then
  If Not IsNull(rsAna!IP) Then
   If rsAna!IP <> 0 Then
    Tha = "CSII"
    Grund = "Anamnese: Insulinpumpe"
    insart = 1
    If UBound(rTh) < 2 Or (Tha <> rTh(UBound(rTh)).therart And Int(rNa(0).AufnDat) <> Int(rTh(UBound(rTh)).Zp)) Then _
       ReDim Preserve rTh(UBound(rTh) + 1)
    rTh(UBound(rTh)).MPNr = 0
    rTh(UBound(rTh)).Pat_id = rNa(0).Pat_id
    rTh(UBound(rTh)).therart = Tha
    rTh(UBound(rTh)).absPos = absPos
    rTh(UBound(rTh)).Zp = rNa(0).AufnDat
    rTh(UBound(rTh)).Grund = Grund
    altTha = Tha
    altinsart = insart
   End If
  End If
 End If
 
 For i = 0 To UBound(rmpn)
'  IF i >= 1331 THEN
'  Debug.Print i
'  END IF
  If rmpn(i).MPNr <> altMPNr Then
'   IF rmpn(i).mpnr = 166 THEN Stop
thains:
   If nachhak <> 0 Then puzuz = 0
   If altMPNr <> 0 Then
    If Not dSL Is Nothing Then
     If dSL.COUNT <> 0 Then
      Set SD = New SortierDatum
      SD.Datum = altZp ' Int(altZp)
      dSL.scinit
      dSL.Eingrenz SD
      Dim DatLetzRez As Date
      DatLetzRez = altZp ' Int(altZp)
      Dim cAkt&
      cAkt = -1
      If dSL.sCa <= dSL.sCe Then
       If Int(dSL.Item(dSL.sCa).Datum) = Int(altZp) Then
        cAkt = dSL.sCa
       Else
'       ginge auch:
'       IF dSL.sCa + nachhak > 1 THEN cAkt = dSL.sCa - 1 + nachhak
        If nachhak Then
         If dSL.sCa > 0 Then cAkt = dSL.sCa ' DatLetzRez = dSL.Item(dSL.sCa).Datum
        Else
         If dSL.sCa > 1 Then cAkt = dSL.sCa - 1 ' DatLetzRez = dSL.Item(dSL.sCa - 1).Datum
        End If
       End If
      Else
       cAkt = dSL.sCe ' DatLetzRez = dSL.Item(dSL.sCe).Datum
       nachhak = 0
      End If
      ' dann noch nach vorangegangenen Rezepten schauen
      If cAkt <> -1 Then
       DatLetzRez = dSL.Item(cAkt).Datum
' wenn das letzte Pumpenrezept innerhalb des letzten Jahres war
       If altZp - DatLetzRez <= 365 Then
        puzuz = puzuz + 1
        Do While True
         If cAkt < 2 Then Exit Do
         If dSL.Item(cAkt - 1).Datum <= rTh(UBound(rTh)).Zp Then Exit Do
         cAkt = cAkt - 1
        Loop
        DatLetzRez = dSL.Item(cAkt).Datum
        Grund = dSL.Item(cAkt).Text
       End If
      End If
     End If ' dsL.Count<>0
    End If ' Not dSL Is Nothing THEN
    If puzuz <> 0 Then
     Tha = "CSII"
    ElseIf ((moins1 <> 0) + (miins1 <> 0) + (nmins1 <> 0) + (abins1 <> 0) + (znins1 <> 0) <> 0) Or _
    (insart = 4 And (moins1 + moins3 <> 0) + (miins1 + miins3 <> 0) + (nmins1 + nmins3 <> 0) + (abins1 + abins3 <> 0) + (znins1 + znins3 <> 0) <> 0) _
    Then
     If glp1z <> 0 Then
      Tha = "GLP1ICT"
     Else
      Tha = "ICT"
     End If
    ElseIf insz = 0 Then
     If glp1z <> 0 Then
      Tha = "GLP1"
     ElseIf oadz <> 0 Then
      Tha = "OAD"
     Else
      Tha = "Diät"
     End If
    Else
     If glp1z <> 0 Then
      Tha = "GLP1Ins"
     ElseIf oadz <> 0 Then
      Tha = "Komb"
     Else
      Tha = "CT"
     End If
    End If ' puzuz <> 0 THEN usw.
   End If ' altMPNr <> 0 THEN
   Dim dsangehängt As Boolean
   dsangehängt = False
'   IF Tha <> altTha AND NOT ((Tha = "OAD" OR Tha = "Diät") AND altTha = "CSII") THEN ' führte bei Z.n. Gestationsdiabetes mit CSII und jetzt OAD zum Fehler, Pat. 2457
'   IF (Tha <> altTha OR insart <> altinsart) AND NOT ((Tha = "Diät") AND altTha = "CSII") THEN '
' => 13.1.20 Pat. 62648: ICT Medplan 13 2019 würde von ICT Medplan 14 2020 überschrieben und gelöscht, da insart-Wechsel
   If (Tha <> altTha) And Not ((Tha = "Diät") And altTha = "CSII") Then '
' 8.1.16 diese Bedingung scheint aktuell ins Leere zu laufen, da zu deren Erfüllung oder Ablehnung aktuell kein Pumpenrezept
' herangezogen wird und ein leerer Medikamentenplan vor einem ICT-Plan die Verkennung desselben bewirkt (Pat. 51612)
'    IF altZp >= rTh(UBound(rTh)).Zp THEN ' nicht: Pseudo-ICT vor Pumpenrezept
' nach dem ersten Plan am gleichen Tag nur eine Therapieform
     If UBound(rTh) < 2 Or (Tha <> rTh(UBound(rTh)).therart And Int(altZp) <> Int(rTh(UBound(rTh)).Zp)) Then _
       ReDim Preserve rTh(UBound(rTh) + 1)
     rTh(UBound(rTh)).MPNr = altMPNr
     rTh(UBound(rTh)).Pat_id = rNa(0).Pat_id
     rTh(UBound(rTh)).therart = Tha
     rTh(UBound(rTh)).Grund = Grund
     rTh(UBound(rTh)).insart = insart
     rTh(UBound(rTh)).absPos = absPos
     rTh(UBound(rTh)).AktZeit = AktZeit
     rTh(UBound(rTh)).Zp = IIf(puzuz, DatLetzRez, altZp) ' nachhak <> 0 And
     dsangehängt = True
'    END IF
' das folgende bewirkt bei Pat. 51612 Auslassung aller Therapien nach ersten und vor der Pumpe
'#If False THEN
'' Bei Pumpenträgern könnte auch der Medikamentenplan ab dem Zp d Pumpenbehandlung nicht mehr geändert worden sein
'' oder rein theoretisch erstmals über ein halbes Jahr nach Pumpenrezepten eingetragen worden sein
'    IF Tha <> "CSII" THEN
'     IF Not dSL Is Nothing THEN
'      IF dSL.Count <> 0 THEN
'       IF puzuz = 0 THEN
'        IF dSL.sCa <= dSL.sCe THEN
'         Tha = "CSII"
'         IF altTha = Tha THEN
'' Pumpenrezept nach Pseudo-ICT
'          IF dsangehängt THEN
'           ReDim Preserve rTh(UBound(rTh) - 1)
'' falls Pseudo-ICT-Datensatz angehängt (Bedinung von oben abgeschrieben), dann wieder löschen
'          Else
'           rTh(UBound(rTh)).therart = Tha
'          END IF
'         Else
'' am gleichen Tag nur eine Therapieform
'          IF UBound(rTh) < 2 OR (Tha <> rTh(UBound(rTh)).therart AND Int(dSL.Item(dSL.sCa).Datum) <> Int(rTh(UBound(rTh)).Zp)) THEN _
'           ReDim Preserve rTh(UBound(rTh) + 1)
'          rTh(UBound(rTh)).mpnr = 0
'          rTh(UBound(rTh)).Pat_id = rNa(0).Pat_id
'          rTh(UBound(rTh)).therart = Tha
'          rTh(UBound(rTh)).absPos = absPos
'          rTh(UBound(rTh)).AktZeit = AktZeit
'          rTh(UBound(rTh)).Zp = dSL.Item(dSL.sCa).Datum
'         END IF
'        END IF ' dSL.sCa <= dSL.sCe THEN
'       END IF ' puzuz = 0 THEN
'      END IF ' dSL.Count <> 0 THEN
'     END IF ' Not dSL Is Nothing THEN
'    END IF ' Tha <> "CSII" THEN
'#END IF ' IF false
    altTha = Tha
    altinsart = insart
   End If ' Tha <> altTha THEN
   If obschluß Then
    GoTo schluß
   End If
   oadz = 0
   glp1z = 0
   moins1 = 0
   miins1 = 0
   nmins1 = 0
   abins1 = 0
   znins1 = 0
   moins3 = 0
   miins3 = 0
   nmins3 = 0
   abins3 = 0
   znins3 = 0
   insz = 0
   puzuz = 0
   insart = 0 ' 24.11.18
   Grund = ""
   altMPNr = rmpn(i).MPNr
   altZp = rmpn(i).Zeitpunkt ' Int(rmpn(i).Zeitpunkt)
   AktZeit = rmpn(i).AktZeit
  End If ' rmpn(i).mpnr <> altMPNr THEN
'    IF i >= 1083 THEN Stop
  Set mEl = Nothing
  mEl.medi = UCase$(GetMed(rmpn(i).Medikament, 0))
  Set mEl = mSL.GetItem(mEl)
  If Not mEl Is Nothing Then
   If mEl.glib <> 0 Or mEl.metf <> 0 Or mEl.gluci <> 0 Or mEl.shglin <> 0 Or mEl.glit <> 0 Or mEl.dpp4 <> 0 Or mEl.sglt2 <> 0 Or mEl.sonstad <> 0 Then
    oadz = oadz + 1
   ElseIf mEl.glp1 <> 0 Then
    glp1z = glp1z + 1
   ElseIf mEl.puzu <> 0 Then
    puzuz = puzuz + 1
    Grund = Grund & IIf(Grund = "", "", " /") & "Puzu: " & mEl.medi
   ElseIf mEl.insart = 1 Then
    If rmpn(i).mo <> vNS Then moins1 = moins1 + 1
    If rmpn(i).mi <> vNS Then miins1 = miins1 + 1
    If rmpn(i).nm <> vNS Then nmins1 = nmins1 + 1
    If rmpn(i).ab <> vNS Then abins1 = abins1 + 1
    If rmpn(i).Zn <> vNS Then znins1 = znins1 + 1 ' obict = (moins1<>0) + (miins1<>0)+(nmins1<>0)+(abins1<>0)+(znins1<>0)
    insz = insz + 1
   ElseIf mEl.InS <> 0 Or mEl.anal <> 0 Then
    If mEl.insart = 3 Then ' 19.12.18: Pat. 54212 soll als ICT eingestuft werden
     If rmpn(i).mo <> vNS Then moins3 = moins3 + 1
     If rmpn(i).mi <> vNS Then miins3 = miins3 + 1
     If rmpn(i).nm <> vNS Then nmins3 = nmins3 + 1
     If rmpn(i).ab <> vNS Then abins3 = abins3 + 1
     If rmpn(i).Zn <> vNS Then znins3 = znins3 + 1
    End If
    insz = insz + 1
   End If
   Dim bishinsart%
   bishinsart = insart
   ' Insart: '0=keines, 1=nur Mahlzeiten,2=nur Verzögerungs,3=nur Misch, 4=verschiedene
   Select Case mEl.insart ' wird fetgelegt in GetMed(rmpn(i).Medikament, 0)
    Case 0 ' keines
    Case 1 ' schnell
     If insart = 0 Then insart = 1 Else If insart = 2 Or insart = 3 Then insart = 4
    Case 2 ' langsam
     If insart = 0 Then insart = 2 Else If insart = 1 Or insart = 3 Then insart = 4
    Case 3 ' Misch
     If insart = 0 Then insart = 3 Else If insart = 1 Or insart = 2 Then insart = 4
   End Select
   If insart <> bishinsart Then
    Grund = Grund & " /" & rmpn(i).MedAnfang
   End If
  End If
 Next i
 'wenn das letzte Pumpenrezept nach dem letzten Medikamentenplan kam
 obschluß = True
 If UBound(rmpn) > 0 Then
  nachhak = 1
  GoTo thains
 End If
schluß:
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in THAfestleg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' THAfestleg
#End If ' thaalt

#If False Then
Public Function theraktakt() ' Therapiearten auf einmal aktualisieren 11.7.10, sollte nur nötig sein, wenn Algorithmus von thafestleg geändert wird
#If Not thaalt Then
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
#If mitfenster Then
 rufauf "ssh", "root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(0)'", 2, "c:\windows\system32\openssh\", -1, 0
#Else
 Call TheraErmitt(0)
#End If
' myEFrag "CALL fuellThaP(0)"
#Else
 Dim rsa As New Recordset, rs1 As New Recordset, rs2 As New Recordset, rAF&
 Lese.ProgStart
 myFrag rsa, "SELECT pat_id FROM `therarten` GROUP BY pat_id"
 Do While Not rsa.EOF
  Set rs1 = Nothing
  myFrag rs1, "SELECT therart FROM `therarten` WHERE pat_id = " & rsa!Pat_id & " ORDER BY zp DESC,mpnr DESC LIMIT 1"
  If Not rs1.BOF Then
   myEFrag "UPDATE `anamnesebogen` SET therakt = '" & rs1!therart & "' WHERE pat_id = " & rsa!Pat_id & " AND therakt <> '" & rs1!therart & "'", rAF
   Debug.Print rsa!Pat_id, rs1!therart, rAF
  End If
  rsa.MoveNext
 Loop
#End If
End Function ' theraktakt()
#End If

Public Function testzwi() ' 3750
 Dim ri$, s1$, s2$
 ri = "TM#<DesktopObject Id=""0-47330""><erstellungsZeitpunkt>25.07.2006 19:51:18</erstellungsZeitpunkt><executeOnLoad>0</executeOnLoad><hideTitel>0</hideTitel><iconPath>$\TurboMed\Symbole\DesktopObjekte\Achtung.ico</iconPath><noteBkColor>0</noteBkColor><noteFgColor>0</noteFgColor><positionBottom>109</positionBottom><positionLeft>118</positionLeft><positionRight>177</positionRight><positionTop>399</positionTop><showAsNote>1</showAsNote><syncInfoList>SyncInfoList</syncInfoList><titel>97274A</titel><toolTipText>Achtung</toolTipText><verankert>0</verankert></DesktopObject>"
 s1 = "DesktopObject Id="""
 s2 = """"
 s1 = "<erstellungsZeitpunkt>"
 s2 = "<"
 testzwi = ZwischenStr(ri, s1, s2)
End Function ' testzwi

Function neuQuartal(frm As Lese, rInhalt$)
    Dim i&
'   If rNa(0).Pat_id = 68076 Then Stop
    ReDim Preserve rFa(UBound(rFa) + 1)
    lfdfl = lfdfl + 1
    rFa(UBound(rFa)).s8000 = f8000
    f8000 = vNS
    rFa(UBound(rFa)).s8100 = f8100
    f8100 = vNS
    rFa(UBound(rFa)).Pat_id = rNa(0).Pat_id
    rFa(UBound(rFa)).Vorname = rNa(0).Vorname
    rFa(UBound(rFa)).Nachname = rNa(0).Nachname
    rFa(UBound(rFa)).AktZeit = AktZeit
    rFa(UBound(rFa)).lfdnr = lfdfl
    If LenB(BGFallNr) <> 0 Then rFa(UBound(rFa)).BGFallNr = BGFallNr: BGFallNr = vNS
    If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
    If UBound(rFa) = 1 Then
'     rsAdo.Open ("SELECT MAX(fid) AS fid FROM faelle")
     Set rsAdo = myEFrag("SELECT MAX(fid) fid FROM faelle")
     If IsNull(rsAdo!FID) Then
      If lies.obMySQL Then
       Call myEFrag("ALTER TABLE `faelle` AUTO_INCREMENT = 1")
      Else
       rsAdo.Close
       ZielDbS = frm.dlg.MdB
       Dim zwiCStr$
       zwiCStr = DBCnS ' DBCn.ConnectionString
       If DBCn.State = 1 Then DBCn.Close
       For i = 0 To 1000
       Next i
       Call BezLöschA
       For i = 0 To 1000
       Next i
'       Call CnOpen(False, zwiCStr, frm)
       Call acon(quelleT)
'       IF lies.obmysql THEN Call myEFrag("use " & myDB) ' nicht nötig, da hier kein lies.obmysql
       Call myEFrag("ALTER TABLE `faelle` ALTER COLUMN fid COUNTER(1,1)")
       Call BezHerstA
      End If
      rFa(UBound(rFa)).FID = 1
     Else
      rFa(UBound(rFa)).FID = IIf(IsNull(rsAdo!FID), 0, rsAdo!FID) + 1
     End If
    Else ' ISNULL(rsAdo!FID) THEN
     rFa(UBound(rFa)).FID = rFa(UBound(rFa) - 1).FID + 1
    End If ' ISNULL(rsAdo!FID) THEN
    rFa(UBound(rFa)).Quartal = rInhalt
    rFa(UBound(rFa)).QAnf = QAnf(rFa(UBound(rFa)).Quartal)
    rFa(UBound(rFa)).QEnd = QEnd(rFa(UBound(rFa)).Quartal)
    rFa(UBound(rFa)).lanrid = Lanr ' :    Lanr = 0   ' 21.3.21
End Function ' neuQuartal

Function dolies(frm As Lese, RKennung$, rInhalt$, obSchluss%, znr&, obmitFormularen%)
' IF rsAnam Is Nothing THEN
'  SET rsAnam = New ADODB.Recordset
'  rsAnam.Open "anamnesebogen", DBCn, adOpenDynamic, adLockOptimistic
' END IF
 Static FormUnt As FormUntTyp
 Dim FeldInhalt$, i&, j&, pos%
 Dim p1&, p2&
 Static aktDiNr%
 Static IKrz$
 Dim rFoNeu%, rFoAbkNeu% ' ob Formular/Abkürzungseintrag neu
 '  FID ggf zuweisen
 Dim obPat%
 Dim FNam() As String, FNam2() As String, ifn&
 Dim rsP As New ADODB.Recordset ' Pharmazentralnummer
 Dim ErrNr&, ErrDes$
 Static ausrrxml!
 On Error GoTo fehler
 FNam = ASMod("a", "w", "p", "m", "d", "v", "h", "i", "r", "s", "t", "du", "fd")
 FNam2 = ASMod("t")
 absPos = absPos + 1
 
  AktZeit = Now ' 27.12.06: wird für die Tabelle `namen` jetzt aus dem Feld 9103 gezogen
' Exit Function
 If RKennung Like "####" Then
'  Dim lPatNeu%
'  IF lk(13) = 3000 THEN
'  IF lKennung = 4101 THEN
  Select Case CLng(RKennung)
' Hier stand noch was
   Case 3000 ' Patientennummer
    If rInhalt = "" Then
     rInhalt = 0
    End If
    If rInhalt <> rNa(0).Pat_id Then ' da 3000 pro Pat. öfter vorkommt
'     lPatNeu = -1
#If thaalt Then
     If rNa(0).Pat_id <> 0 Then
      Call THAfestleg ' ausrangiert 12.12.20
     End If
#End If
     If rInhalt = "0" Then
      If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
'      myFrag rsAdo, "SELECT (MAX(pat_id) + 1) AS pid FROM `namen`"
      Set rsAdo = myEFrag("SELECT (MAX(pat_id) + 1) pid FROM `namen`")
      Err.Raise 999, , "Achtung: Patientennummer 0! Muß geändert werden, z.B. auf " + CStr(rsAdo!pid) + ": Bitte mit Turbomed abgleichen und auch dort ändern"
     End If
     myEFrag "UPDATE namen SET stbytea = " & AktByte & " WHERE pat_id = " & rInhalt ' 17.11.19 um Programmabbrüchen während Import nachspüren zu können
     rNa(0).StByteA = AktByte
     rNa(0).lfdnr = lfdnr
     rNa(0).Pat_id = rInhalt
     rNa(0).AktZeit = lAktZeit
     Call PatInit
     ausrrxml = 0
    End If ' RInhalt <> rna(0).Pat_ID
   Case 3004 ' f3004 unbekanntes Feld, 0 oder 2
    rNa(0).f3004 = rInhalt
   Case 3006 ' f3006 unbekanntes Feld, 5.1.0 oder 5.1.0
    rNa(0).f3006 = rInhalt
   Case 3100 ' Namensvorsatz
    rNa(0).NVorsatz = rInhalt
'    rAna(0).NVorsatz = rna(0).NVorsatz
 ' update macht dann den SQLBindParameter not used for all parameters mysq-fehler
 '  IF rAna.EditMode = 1 OR rAna.EditMode = adEditInProgress THEN
 '      rsAnam!NVorsatz = Rinhalt
 '   END IF
   Case 3101 ' Nachname
    If rInhalt Like "zzz*" Then
     rInhalt = Mid$(rInhalt, 4)
     Tkz = -1
    End If
    rNa(0).Nachname = rInhalt
'    rAna(0).Nachname = rna(0).Nachname
   Case 3102 ' Vorname
    rNa(0).Vorname = rInhalt
'    rAna(0).Vorname = rNa(0).Vorname
   Case 3103 ' Geburtsdatum
    rNa(0).GebDat = BDTtoDate(rInhalt)
'    rAna(0).GebDat = rNa(0).GebDat
   Case 3104 ' Titel
    rNa(0).Titel = rInhalt
'    rAna(0).Titel = rNa(0).Titel
   Case 3105, 3119 ' Versichertennummer
    rNa(0).Versichertennummer = rInhalt
   Case 3107
    rNa(0).Straße = rInhalt
   Case 3108 ' KVKStatus
    rNa(0).KVKStatus = rInhalt
   Case 3109 ' Hausnr
    rNa(0).Hausnr = rInhalt
   Case 3110 ' Geschlecht
    rNa(0).geschlecht = IIf(rInhalt = "1" Or rInhalt = "M", "m", IIf(rInhalt = "2" Or rInhalt = "W", "w", " "))
'    rAna(0).Anrede = IIf(rNa(0).Geschlecht = "m", "Herr", "Frau")
   Case 3112 ' Postleitzahl
    rNa(0).plz = rInhalt
   Case 3113 ' Ort
    rNa(0).ort = rInhalt
   Case 3114 ' Lkz
    rNa(0).Lkz = rInhalt
   Case 3115 ' Anschrzus
    rNa(0).Anschrzus = rInhalt
   Case 3120 ' NVors
    rNa(0).NVors = rInhalt
   Case 3121 ' Postfach-Postleitzahl
    rNa(0).PFPlz = rInhalt
   Case 3122 ' Postfach-Ort
    rNa(0).PFOrt = rInhalt
   Case 3123 ' Postfach-Nummer
    rNa(0).PFNr = rInhalt
   Case 3124 ' unbenk. Feld, bisher immer leer
    rNa(0).f3124 = rInhalt
   Case 3215 ' Anschriftenzusatz, aufgeteilt
    Call aufSplit(rInhalt)
    If ArraInd > 1 Then
     rNa(0).AnschrZus_2 = Arra(1)
    End If
   Case 3216 ' Postfach 2, aufgeteilt
    Call aufSplit(rInhalt)
    If ArraInd > 1 Then
     rNa(0).Postfach_2 = Arra(1)
    End If
   Case 3217 ' Länderkode, vermutlich Herkunftsland, aufgeteilt
    Call aufSplit(rInhalt)
    If ArraInd > 1 Then
     rNa(0).LK_2 = Arra(1)
    End If
   Case 3603 ' BG-Fall-Nummer
    BGFallNr = rInhalt
   Case 3610 ' Aufnahmedatum
    rNa(0).AufnDat = BDTtoDate(rInhalt)
    Call VorstellSetz(rNa(0).AufnDat)
   Case 3612 ' Beruf
    rNa(0).Beruf = rInhalt
   Case 3625 ' Arbeitgeber
    rNa(0).Arbeitgeber = rInhalt
   Case 3626, 3208 ' Telefonnummer von 3629 dupliziert
    If rInhalt <> vNS Then
     rNa(0).PrivatTel = rInhalt
    End If
   Case 3629 ' verschiedene Telefonnummern mit Formatierung
'    Debug.Print RInhalt
    Call aufSplit(rInhalt)
    FeldInhalt = vNS
    For i = 1 To ArraInd - 3
     FeldInhalt = FeldInhalt + Arra(i)
    Next
    Select Case Arra(ArraInd - 2)
     Case "1"
      Select Case Arra(ArraInd - 1)
       Case "Privat"
        If Arra(ArraInd) = "Tel" Then rNa(0).PrivatTel = FeldInhalt Else If Arra(ArraInd) = "Fax" Then rNa(0).PrivatFax = FeldInhalt Else If Arra(ArraInd) = "Mobil" Then rNa(0).PrivatMobil = FeldInhalt
       Case "Dienst"
        If Arra(ArraInd) = "Tel" Then rNa(0).DienstTel = FeldInhalt Else Err.Raise 999
       Case Else: Err.Raise 999
      End Select
     Case "2"
      Select Case Arra(ArraInd - 1)
       Case "Privat"
        If Arra(ArraInd) = "Tel" Then rNa(0).PrivatTel_2 = FeldInhalt Else Err.Raise 999
       Case Else: Err.Raise 999
      End Select
    End Select
   Case 3630 ' KV-Nummer des überweisenden / -> Hausarzt
'    IF ISNULL(RInhalt) THEN
'     IF obÜProt THEN Print #322, "Pat: " + CStr(rNa(0).Pat_ID) + ": KV-Nr des Überweisenden (Feld 3630) leer!"
'    Else
     If rNa(0).KVNr = "" Then
      rNa(0).KVNr = Left$(REPLACE$(rInhalt, "/", ""), 10)
     ElseIf rNa(0).KVNr2 = "" Then
      rNa(0).KVNr2 = Left$(REPLACE$(rInhalt, "/", ""), 10)
     ElseIf rNa(0).KVNr3 = "" Then
      rNa(0).KVNr3 = Left$(REPLACE$(rInhalt, "/", ""), 10)
     ElseIf rNa(0).KVNr4 = "" Then
      rNa(0).KVNr4 = Left$(REPLACE$(rInhalt, "/", ""), 10)
     End If
     Dim obrKV%, iirunde%, aktKVNr$
     For iirunde = 1 To 4
      obrKV = -1
      Select Case iirunde
       Case 1: aktKVNr = rNa(0).KVNr
       Case 2: aktKVNr = rNa(0).KVNr2
       Case 3: aktKVNr = rNa(0).KVNr3
       Case 4: aktKVNr = rNa(0).KVNr4
      End Select
      If LenB(aktKVNr) = 0 Then Exit For ' iirunde
      For j = 1 To UBound(rKv)
       If rKv(j).KVNr = aktKVNr And rKv(j).Pat_id = rNa(0).Pat_id Then
        obrKV = 0
        Exit For
       End If
      Next j
      If obrKV Then
       ReDim Preserve rKv(UBound(rKv) + 1)
       rKv(UBound(rKv)).absPos = absPos
       rKv(UBound(rKv)).KVNr = aktKVNr
       rKv(UBound(rKv)).Pat_id = rNa(0).Pat_id
       rKv(UBound(rKv)).AktZeit = AktZeit
      End If
     Next iirunde
'    END IF
   Case 3631 ' Weggeldzone
    Call aufSplit(rInhalt)
    If ArraInd > 1 Then
     rNa(0).Weggeldzone = Arra(1)
     rNa(0).WeggzZahl = Arra(2)
    End If
   Case 3632: rNa(UBound(rNa)).Verwandt = IIf(LenB(rNa(UBound(rNa)).Verwandt) = 0, vNS, rNa(UBound(rNa)).Verwandt & vbCrLf) & rInhalt
   Case 3633: rNa(UBound(rNa)).zubenach = IIf(LenB(rNa(UBound(rNa)).zubenach) = 0, vNS, rNa(UBound(rNa)).zubenach & vbCrLf) & rInhalt
   Case 3634: ' Notiz, Zeile für Zeile
    rNa(UBound(rNa)).Notiz = IIf(LenB(rNa(UBound(rNa)).Notiz) = 0, vNS, rNa(UBound(rNa)).Notiz & vbCrLf) & rInhalt
   Case 3654: rNa(UBound(rNa)).Cave = IIf(LenB(rNa(UBound(rNa)).Cave) = 0, vNS, rNa(UBound(rNa)).Cave & vbCrLf) & rInhalt
   Case 3635 ' interne Zuordnung Arzt bei GP / => LANR <- interne Zordnung
    Call aufSplit(rInhalt)
    If ArraInd > 0 Then
     Lanr = holiAzu(Arra(1))
    End If
   Case 3628 ' Sprache
    rNa(0).Sprache = rInhalt
   Case 3636 ' Betriebsstättennummer
    Call aufSplit(rInhalt)
    If ArraInd > 0 Then BSNR = Arra(1)
   Case 3649, 5999 ' Diagnosedatum
'    If RKennung = 3649 Then Stop
    messDatum = BDTtoDate(rInhalt)
    messDatumD = messDatum
   Case 3650, 6000 ' Dauer-, Einzeldiagnose
    obD_ = IIf(RKennung = 3650, 1, 0)
'    If rInhalt = "Harninkontinenz" Then Stop
    DText_ = rInhalt
   Case 3652, 6210 ' Medikament auf Rezept (neben 6210) 3652 = bis Mitte 2012, 6210 = danach
    If InStr(FormVorl, "Kassenrezept") <> 0 Or InStr(FormVorl, "Privatrezept") <> 0 Then ' Tauchte, wohl nach Teillöschung, auch einmal mit "Medikamentenplan" auf
     Call RezEintr(rInhalt, False, InStr(FormVorl, "2002") = 0)
    End If
   Case 3750 ' Desktopobjekt
    ReDim Preserve rDe(UBound(rDe) + 1)
    rDe(UBound(rDe)).Pat_id = rNa(0).Pat_id
    rDe(UBound(rDe)).AktZeit = AktZeit
    Dim ZStr$
    ZStr = ZwischenStr(rInhalt, "<erstellungsZeitpunkt>", "<")
    If IsDate(ZStr) Then rDe(UBound(rDe)).erstZP = CDate(ZStr)
    rDe(UBound(rDe)).IDS = ZwischenStr(rInhalt, "DesktopObject Id=""", """")
    If rDe(UBound(rDe)).IDS = vNS Then rDe(UBound(rDe)).IDS = ZwischenStr(rInhalt, "DesktopObject Id=\""", "\""")
    rDe(UBound(rDe)).exoL = ZwischenStr(rInhalt, "<executeonLoad>", "<")
    rDe(UBound(rDe)).iconPath = ZwischenStr(rInhalt, "<iconPath>", "<")
    rDe(UBound(rDe)).syncInfoList = ZwischenStr(rInhalt, "<syncInfoList>", "<")
    rDe(UBound(rDe)).Titel = ZwischenStr(rInhalt, "<titel>", "<")
    rDe(UBound(rDe)).toolTipText = ZwischenStr(rInhalt, "<toolTipText>", "<")
    ZStr = ZwischenStr(rInhalt, "<hideTitel>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).hideT = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<noteBkColor>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).noteBkColor = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<noteFgColor>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).noteFgColor = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<positionBottom>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).positionBottom = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<positionLeft>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).positionLeft = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<positionRight>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).positionRight = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<positionTop>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).positionTop = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<showAsNote>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).showAsNote = CDbl(ZStr)
    ZStr = ZwischenStr(rInhalt, "<verankert>", "<")
    If IsNumeric(ZStr) Then rDe(UBound(rDe)).verankert = CDbl(ZStr)
   Case 3800 ' Chroniker angekreuzt, '-' => '?'
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then
      rNa(0).obChk = Arra(1)
     End If
   Case 4101 ' Quartal
    If Not obHausBesuch Or QAnf(rInhalt) < QAnf(rFa(UBound(rFa)).Quartal) Then
     Call neuQuartal(frm, rInhalt)
    End If
 '   IF UBound(rfa) = 0 AND rfa(0).Pat_ID = 0 THEN ReDim Preserve rfa(UBound(rfa)) ELSE ReDim Preserve rfa(UBound(rfa) + 1)
   Case 4102 ' ausgestellt am <> letzter Vorgang (dieser nicht exportiert!)
    rFa(UBound(rFa)).lanrid = Lanr ' :    Lanr = 0 ' Kommentar 21.3.21
    rFa(UBound(rFa)).ausgst = BDTtoDate(rInhalt)
    Call VorstellSetz(BDTtoDate(rInhalt))
   Case 4104 ' VKNr
    rFa(UBound(rFa)).VKNr = rInhalt
   Case 4106 ' Kostenträgerabrechnungsbereich, immer 00
    rFa(UBound(rFa)).KtrAbrB = rInhalt
   Case 4107 ' Abrechnungsart, 1 oder 2
    rFa(UBound(rFa)).AbrAr = rInhalt
   Case 4108 ' ? -> allenfalls mit "103" befüllt
    rFa(UBound(rFa)).f4108 = rInhalt
   Case 4109 ' letzte Vorlage
    rFa(UBound(rFa)).lVorl = BDTtoDate(rInhalt)
    Call VorstellSetz(rFa(UBound(rFa)).lVorl)
   Case 4110 ' Uhrzeit zu letzter Vorlage
    rFa(UBound(rFa)).lVorl = rFa(UBound(rFa)).lVorl + BDTtoTime(rInhalt)
   Case 4111 ' IK
    rFa(UBound(rFa)).IK = rInhalt
   Case 4112 'KVKStatus
    rFa(UBound(rFa)).KVKs = rInhalt
   Case 4113 ' Status-Ergänzung / Ost-West-Status
    rFa(UBound(rFa)).KVKserg = rInhalt
   Case 4121 ' Gebührenordnung, 1 oder 2
    rFa(UBound(rFa)).GebOr = rInhalt
   Case 4122 ' Abrechnungsgebiet, 00
    rFa(UBound(rFa)).AbrGb = rInhalt
   Case 4123 ' Personenkreis/Untersuchungskategorie
    rFa(UBound(rFa)).PersKreis = rInhalt
   Case 4124 ' SKT-Zusatzangaben
    rFa(UBound(rFa)).SKtZusatz = rInhalt
   Case 4131 ' unbek. Feld, immer 0
    rFa(UBound(rFa)).f4131 = rInhalt
   Case 4132 ' unbek. Feld, 0-6
    rFa(UBound(rFa)).f4132 = rInhalt
   Case 4133 ' Versichtenschutzbeginn
    rFa(UBound(rFa)).VschBeg = RBDTtoDate(rInhalt)
   Case 4134 ' KKasse_2
    rFa(UBound(rFa)).KKasse_2 = rInhalt
   Case 4136 ' Faktor persönlich
    rFa(UBound(rFa)).FaktPers = CDbl(REPLACE(rInhalt, ".", ","))
   Case 4137 ' Faktor persönlich
    rFa(UBound(rFa)).FaktTechn = CDbl(REPLACE(rInhalt, ".", ","))
   Case 4138 ' Faktor Labor
    rFa(UBound(rFa)).FaktLabor = CDbl(REPLACE(rInhalt, ".", ","))
   Case 4144 ' Turbomed-Fallnummer
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then
      rFa(UBound(rFa)).TMFNr = Arra(1)
     End If
   Case 4150 ' Behandlungsfall: Beginn
     rFa(UBound(rFa)).BhFB = BDTtoDate(rInhalt)
     If rFa(UBound(rFa)).BhFB >= maxBhFB Then
      maxBhFB = rFa(UBound(rFa)).BhFB
      imaxBhFB = UBound(rFa)
     End If
   Case 4151 ' Behandlungsfall: Progende (Musterwoman) / Wohl Ende des Behandlungsfallbeginnquartals
     rFa(UBound(rFa)).BhFE1 = BDTtoDate(rInhalt)
   Case 4152 ' Behandlungsfall: Progende (Musterwoman),bei offenem Fall 00000000, sonst z.B. 30092006 für 3/06
     rFa(UBound(rFa)).BhFE2 = BDTtoDate(rInhalt)
   Case 4202 ' Unfall, Unfallfolgen nach 4152, bei Pat.2775,312,2470 "1"
     rFa(UBound(rFa)).f4202 = rInhalt
   Case 4205 ' Auftrag bei Überweisung
     rFa(UBound(rFa)).Auftrag = rFa(UBound(rFa)).Auftrag & Stutz(rInhalt)
   Case 4206 ' letzter Tag der Regel
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then
      rFa(UBound(rFa)).letzteRegel = BDTtoDate(Arra(1))
'      rFa(UBound(rFa)).letzteRegel = rInhalt
     End If
   Case 4207 ' Verdacht bei Überweisung
     rFa(UBound(rFa)).Verdacht = rFa(UBound(rFa)).Verdacht & Stutz(rInhalt)
   Case 4208 ' Befund bei Überweisung
     rFa(UBound(rFa)).Befund = rFa(UBound(rFa)).Befund & Stutz(rInhalt)
   Case 4209 ' Überweisungstext / Auftrag bei Privatpatienten
     rFa(UBound(rFa)).ÜwText = rInhalt
   Case 4266 ' unbekanntes Daumsfeld bisher nur bei Musterwoman
     rFa(UBound(rFa)).f4266 = BDTtoDate(rInhalt)
     ' Überwiesen von
   Case 4247, _
        4231, 4241, 4285 ' Überweiser, hat sich offenbar von 4231 auf 4241 geändert! (31.7.05)
   ' 10.1.10: Muß sich irgendwann auf 4285 geändert haben
     Call aufSplit(rInhalt)
     If ArraInd >= 5 Then
        rFa(UBound(rFa)).ÜWNNr = IIf(IsNull(Arra(ArraInd)), vNS, REPLACE$(Arra(ArraInd), "/", ""))
        rFa(UBound(rFa)).ÜWNaN = Arra(ArraInd - 1)
        rFa(UBound(rFa)).ÜWVsw = Arra(ArraInd - 2)
        rFa(UBound(rFa)).ÜWVor = Arra(ArraInd - 3)
        rFa(UBound(rFa)).ÜWTit = Arra(ArraInd - 4)
      Dim uers As New ADODB.Recordset, uerunde%, insrunde%, Tabl$, KVNr$, kvnrrs As ADODB.Recordset
      Dim nnafeld$
      For insrunde = 1 To 2
       If insrunde = 1 Then Tabl = "liuez": nnafeld = "name" Else Tabl = "ueberwvon": nnafeld = "nachname"
       KVNr = REPLACE$(Arra(5), "/", "")
       If LenB(KVNr) = 0 Then
        KVNr = myFrag(kvnrrs, "SELECT MAX(kvnr)+1 from " & Tabl & " WHERE kvnr<700000000").Fields(0)
       End If
       For uerunde = 1 To 2
resume_4247:
        On Error GoTo fehler_4247
        Set uers = Nothing
        myFrag uers, "SELECT id FROM " & Tabl & " WHERE kvnr = '" & KVNr & "' AND titel = '" & Arra(1) & "' AND vorname = '" & Arra(2) & "' AND zusatz = '" & Arra(3) & "' AND " & nnafeld & " = '" & Arra(4) & "'"
        If uers.EOF Then
         InsKorr DBCn, DBCnS, "INSERT INTO " & Tabl & "(kvnr,titel,vorname,zusatz," & nnafeld & ") VALUES('" & KVNr & "','" & Arra(1) & "','" & Arra(2) & "','" & Arra(3) & "','" & Arra(4) & "')", rAF
        Else ' uers.EOF
         Exit For
        End If ' uers.EOF
        On Error GoTo fehler
       Next uerunde
      Next insrunde
      If Not uers.EOF Then
       rFa(UBound(rFa)).üwvid = uers!id
      End If
      Set uers = Nothing
     Else
      MsgBox "Stop in dolies: Überwiesen von mit falscher Zahl an Unterteilungen (" & ArraInd & " statt 5)"
      Stop
     End If
   Case 4290 ' Auftrag (Privatpatienten)
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).ÜwText = Arra(1)
   Case 4210 ' Ankreuzfeld LSR
     rFa(UBound(rFa)).f4210 = rInhalt
   Case 4211 ' Ankreuzfeld HAH
     rFa(UBound(rFa)).AkfHAH = rInhalt
   Case 4213 '  Ankreuzfeld AK
     rFa(UBound(rFa)).AkfAK = rInhalt
   Case 4218 ' Überwiesen von
     Call aufSplit(rInhalt)
     If ArraInd > 2 Then
      rFa(UBound(rFa)).ÜbWVLANR = REPLACE$(Arra(1), "/", vNS)
      rFa(UBound(rFa)).ÜbWVBSNR = REPLACE$(Arra(2), "/", vNS)
      rFa(UBound(rFa)).ÜbWVKVNR = REPLACE$(Arra(3), "/", vNS)
     ElseIf ArraInd > 1 Then
      rFa(UBound(rFa)).ÜbWVKVNR = REPLACE$(Arra(2), "/", vNS)
     ElseIf ArraInd > 0 Then
      rFa(UBound(rFa)).ÜbWVKVNR = REPLACE$(Arra(1), "/", vNS)
     End If
' umgestellt und geändert 25.10.14
     If rFa(UBound(rFa)).ÜbWVKVNR <> vNS Then
      rFa(UBound(rFa)).Übwr = rFa(UBound(rFa)).ÜbWVKVNR  ' Left$(rFa(UBound(rFa)).ÜbWVKVNR & "00", 9)
     ElseIf rFa(UBound(rFa)).ÜbWVBSNR <> vNS Then
      rFa(UBound(rFa)).Übwr = rFa(UBound(rFa)).ÜbWVBSNR ' Left$(rFa(UBound(rFa)).ÜbWVBSNR & "00", 9)
     End If
'     IF rInhalt <> vNS THEN rFa(UBound(rFa)).Übw = LEFT(replace$(rInhalt, "/", vNS), 7)
   Case 4219 ' Anderer Überweiser
     rFa(UBound(rFa)).AndÜw = REPLACE$(rInhalt, "/", "")
     If rFa(UBound(rFa)).Übwr = vNS And rInhalt <> vNS Then rFa(UBound(rFa)).Übwr = rInhalt ' Left$(rFa(UBound(rFa)).AndÜw & "00", 9) ' geändert 25.10.14
'     IF rInhalt <> vNS THEN rFa(UBound(rFa)).Übw = LEFT(replace$(rInhalt, "/", vNS), 7)
   Case 4220 ' Überweisungsziel
     rFa(UBound(rFa)).ÜWZiel = rInhalt
   Case 4242 ' LANR des Überweisers
     rFa(UBound(rFa)).ÜbwLANR = rInhalt
   Case 4233
     rFa(UBound(rFa)).statNuller = rInhalt
   Case 4236 ' bei stationär: klasse
     rFa(UBound(rFa)).statKlasse = rInhalt
   Case 4237 ' unbek
     rFa(UBound(rFa)).f4237 = rInhalt
   Case 4238 ' bei stationär: Behandlungstage
     rFa(UBound(rFa)).statBehTage = rInhalt
   Case 4239 ' Scheingruppe, 24, 41, 90
     rFa(UBound(rFa)).SchGr = rInhalt
     ' zuvor: testaltQuartal
     Select Case rFa(UBound(rFa)).SchGr
      Case "00", "20", "21", "23", "24"  ' 23 = kurativ Konsiliaruntersuchung 21 = Auftragsleistung
       If rFa(UBound(rFa)).Quartal <> ZQuart(rFa(UBound(rFa)).BhFB) Then
        rFa(UBound(rFa)).altQuart = rFa(UBound(rFa)).Quartal
        rFa(UBound(rFa)).Quartal = ZQuart(rFa(UBound(rFa)).BhFB)
       End If
      Case "90", "92", "41", "43" ' 43 = Notfall 41 = ÄND, 92 = nur Musterwoman
      Case Else
'    Debug.Print aFäl!Pat_id, aFäl!Nachname, aFäl!Vorname, aFäl!SchGr, aFäl!FID, aFäl!Quartal, aFäl!BhFB
    End Select
   Case 4243 ' Weiterbehandelnder Arzt in Worten
     rFa(UBound(rFa)).Weiterbeh = rInhalt
   Case 4401 ' 4401, Praxisgebühr -Array
     Call aufSplit(rInhalt)
     rFa(UBound(rFa)).PGeb = Arra(1)
   Case 4402 ' PGebErg  '4402, Array ?
     Call aufSplit(rInhalt)
     rFa(UBound(rFa)).PGebErg = Arra(1)
   Case 4403 ' 4403, Mahnfrist bis
     Call aufSplit(rInhalt)
     rFa(UBound(rFa)).Mahnfrist = Arra(1)
   Case 4505 ' Unfallort
     rFa(UBound(rFa)).Unfallort = rInhalt
   Case 4506 ' Beschäftigt als
     rFa(UBound(rFa)).BeschAls = rInhalt
   Case 4507 ' Beschäftigt seit
     rFa(UBound(rFa)).BeschSeit = BDTtoDate(rInhalt)
   Case 4509 ' Unfallbetrieb, eigentlich aufgeteilt
     rFa(UBound(rFa)).Unfallbetrieb = rInhalt
   Case 4570 ' f4570, unbekanntes Feld
     rFa(UBound(rFa)).f4570 = rInhalt
   Case 4580 ' Abrechnungsart / Privatabrechnungskatalog
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).GOÄKatNr = Arra(1)
     If ArraInd > 1 Then rFa(UBound(rFa)).GOÄKatName = Arra(2)
   Case 4585 ' abrechnender Arzt
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).abrArzt = Arra(1)
   Case 4586 ' private Versicherung
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).privVers = Arra(1)
   Case 4590 ' Auftrag bei Privatpatienten
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then
      rFa(UBound(rFa)).ÜwText = Arra(1)
     End If
   Case 4602 ' Rechnungsanschrift (Musterfrau)
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).AdNam = Arra(1)
     If ArraInd > 1 Then If Arra(1) <> "" Then rFa(UBound(rFa)).AdNam = rFa(UBound(rFa)).AdNam + Arra(2)
     If ArraInd > 2 Then If Arra(1) <> "" Then rFa(UBound(rFa)).AdNam = rFa(UBound(rFa)).AdNam + Arra(3)
     If ArraInd > 3 Then If Arra(1) <> "" Then rFa(UBound(rFa)).AdNam = rFa(UBound(rFa)).AdNam + Arra(4)
     If ArraInd > 4 Then rFa(UBound(rFa)).AdStr = Arra(5)
     If ArraInd > 5 Then rFa(UBound(rFa)).AdPlz = Arra(6)
     If ArraInd > 6 Then rFa(UBound(rFa)).AdOrt = Arra(7)
   Case 4603 ' Überweiser BG, eigentlich aufteilt
     rFa(UBound(rFa)).ÜwBG = rInhalt
   Case 4604  'Behandlungsfall: Progende (Musterwoman)
     rFa(UBound(rFa)).BhFE = BDTtoDate(rInhalt)
   Case 5000, 6200 ' 6200 = Tagesdatum (BDT-Beschreibung ZI 1994)
    messDatum = BDTtoDate(rInhalt)
    messDatumD = messDatum
    If obHausBesuch Then
     If UBound(rFa) = 0 Then
      Call neuQuartal(frm, QuartalStr(messDatum))
     End If
     While messDatum >= QEnd(rFa(UBound(rFa)).Quartal) + 1
      Call neuQuartal(frm, QuartalStr(QEnd(rFa(UBound(rFa)).Quartal) + 1))
     Wend
    End If ' obHausBesuch Then
    If messDatum > rNa(0).lAktTM Then rNa(0).lAktTM = messDatum
' Hier fängt auf jeden Fall ein neues Formular an
    FormID = 0
    FormAbk = vNS
    DokPfad = vNS
    DokArt = vNS
    DokName = vNS
   Case 5001 ' Leistungsziffer
    Call LeistEintr0(rInhalt)
    Call SchulungszifferZuord(rInhalt, DMSchulz, RRSchulz)
    If Not obvorgestellt Then
     Call VorstellSetz(messDatum)
    End If
   Case 9999 ' Chargennummer
    If lKennung = 5001 Then
     rLe(UBound(rLe)).Charge = rInhalt
    End If
   Case 5101 ' letzter Vorgang (abgerechnet?)
     Call aufSplit(rInhalt)
     If ArraInd > 0 Then rLe(UBound(rLe)).letzVorg = BDTtoDate(Arra(1))
   Case 5098 ' 0000000000
    rLe(UBound(rLe)).lanrid = Lanr ' : Lanr = 0 ' Kommentar 21.3.21
    rLe(UBound(rLe)).f5098 = rInhalt
   Case 5099 ' LANR
    rLe(UBound(rLe)).Lanr = rInhalt
   Case 5002 ' Art der Untersuchung
    rLe(UBound(rLe)).f5002 = rInhalt
   Case 5005 ' Anzahl
    rLe(UBound(rLe)).f5005 = rInhalt
   Case 5006 ' um Uhrzeit
    rLe(UBound(rLe)).f5006 = rInhalt
   Case 5009 ' freier Begründungstext
    rLe(UBound(rLe)).f5009 = rInhalt
   Case 5010 ' Medikament
    rLe(UBound(rLe)).Med = rInhalt
   Case 5011 ' Sachkostenbezeichnung
    rLe(UBound(rLe)).Sachkbez = rInhalt
   Case 5012 ' Sachkosten in Cent
    rLe(UBound(rLe)).Sachkct = MachNumerisch(rInhalt)
   Case 5015 ' Organ
    rLe(UBound(rLe)).f5015 = rInhalt
   Case 5016 ' Name des Arztes (Briefempfänger)
    rLe(UBound(rLe)).f5016 = rInhalt
   Case 5018 ' Zone bei Besuchen
    rLe(UBound(rLe)).Zone = rInhalt
   Case 5021 ' Datum letzte Krebsvorsorge
    rLe(UBound(rLe)).f5021 = rInhalt
   Case 5026 ' Entlassungsdatum
    rLe(UBound(rLe)).f5026 = rInhalt
   Case 5062 ' Faktor
    rLe(UBound(rLe)).Faktor = rInhalt
   Case 6001, 3673 ' ICD, "ICD-Code Dauerdiagnose
    ICD_ = rInhalt
   Case 6003, 3674 ' Diagnosensicherheit, "Diagnosesicherheit Dauerdiagnose"
    DSic_ = rInhalt
   Case 6004, 3675 ' Diagnosenseite
    DSe_ = rInhalt
   Case 6006, 3676 ' Diagnosenattribut = optionale Erläuterung
    DAt_ = rInhalt
   Case 6008, 3677 ' Ausnahmebegründung
    DAus_ = rInhalt
   Case 6009 ' interne Bemerkung
    DiBm_ = rInhalt
' Quartalsdiagnose: 6010 Falsch, 6011 ?
' anamnest. Dauerdiagnose: kommt nur einmal vor, 6010 Falsch, 6011 nein
' anamnest. für diesen Schein relevant: 2x, 6010 einmal falsch, einmal wahr, 6011 nein
' relevant: 2 x , 6010 einmal falsch (=ohne Kodierrichtlinien relevanter Eintrag), einmal wahr(=mit KR relevant), 6011 ja
' relevant und dd-Eintrag gelöscht: 1 x 6010 Wahr
' relevant und dann bdd-Eintrag gelöscht: 1 x 6010 Falsch, 6011 ja
' gelöscht: 6330bddg, 6331gesichert Allergie, 6330ddg, 6331gesichert Allergie
   Case 6010 ' Diagnose gelöscht (Feld dgg)
'    If DText_ = "Harninkontinenz" Then Stop
    Call aufSplit(rInhalt)
    If ArraInd > 0 Then If Arra(1) = "Falsch" Or Arra(1) = "False" Then f6010_ = 0 Else f6010_ = 1
   Case 6011 ' TM#?
    Call aufSplit(rInhalt)
    If ArraInd > 0 Then f6011_ = Left$(Arra(1), 1)
       ' 27.3.07: von "dd" auf "d" abgeänderte Diagnosen werden in der BDT-Datei aufgrund eines Turbomed-Fehlers doppelt ausgegeben -> dann "d" nehmen
    Dim inr%
    obDStr = False
    If UBound(rDi) > 0 Then
'     If DText_ = "Harninkontinenz" Then Stop
     For inr = 1 To UBound(rDi)
'      If rDi(inr).DiagText = DText_ Then Stop
' wenn die gleichartige Diagnose schon mal gefunden wurde (üblicherweise der bdd-Eintrag zu dd oder umgekehrt),
' dann keinen neuen Eintrag erzeugen.
' f6010-Feld falsch = dd-Eintrag existiert (wird nur bei ausgeschalteten Kodierrichtlinien in Abrechnungsdatei geschrieben)
' f6010-Feld wahr = nur bdd-Eintrag existiert (wird nur bei eingeschalteten Kodierrichtlinien in Abrechnungsdatei geschrieben)
      If rDi(inr).DiagText = DText_ _
         And rDi(inr).ICD = ICD_ _
         And rDi(inr).DiagSicherheit = DSic_ _
         And rDi(inr).DiagSeite = DSe_ _
         And rDi(inr).DiagAttr = DAt_ _
         And rDi(inr).AusnBegr = DAus_ _
         And rDi(inr).intBemerk = DiBm_ _
         And (rDi(inr).obDauer = obD_ And obD_ <> 0) Then
       If f6010_ <> 0 Then ' Einzeldiagnosen immer übermittelt
        rDi(inr).obKasse = 1
        rDi(inr).lKasse = MAX(rDi(inr).lKasse, messDatum)
       Else
        rDi(inr).f6010 = 0
       End If ' f6010_ <> 0 Or obD_ = 0 Then ' Einzeldiagnosen immer übermittelt
       Dim jnr%
       For jnr = 1 To UBound(Diag) ' Schattenarrays für DiagString
        If Diag(jnr) = DText_ _
           And ICD(jnr) = ICD_ _
           And DSic(jnr) = DSic_ _
           And DiagSe(jnr) = DSe_ _
           And DiagAttr(jnr) = DAt_ _
           And DiagAus(jnr) = DAus_ _
           And DiagiBm(jnr) = DiBm_ _
           And obDauer(jnr) = obD_ Then
         If f6010_ <> 0 Or obD_ = 0 Then
          obKasse(jnr) = 1
          lKasse(jnr) = MAX(lKasse(jnr), messDatum)
         Else
          f6010(jnr) = 0
         End If ' f6010_ <> 0 Or obD_ = 0 Then
         Exit For
        End If ' Diag(jnr) = DText_
       Next jnr
       GoTo difertig
      End If ' rDi(inr).DiagText = DText_
     Next inr
    End If ' UBound(rDi) > 0 Then
    If (obD_ <> 0 Or rFa(UBound(rFa)).Quartal = AktQ) And DiagNr <= UBound(diagdat) Then ' formal müßte hier evtl altQuart rein, da dessen Belegung auf diese Schleife vorverlegt wurde
     obDStr = True
    End If
    ReDim Preserve rDi(UBound(rDi) + 1)
    aktDiNr = UBound(rDi)
    rDi(aktDiNr).DiagText = DText_
    rDi(aktDiNr).ICD = ICD_
    rDi(aktDiNr).DiagSicherheit = DSic_
    rDi(aktDiNr).DiagSeite = DSe_
    rDi(aktDiNr).DiagAttr = DAt_
    rDi(aktDiNr).AusnBegr = DAus_
    rDi(aktDiNr).intBemerk = DiBm_
    rDi(aktDiNr).DiagDatum = messDatum
    rDi(aktDiNr).obDauer = obD_
    rDi(aktDiNr).f6010 = f6010_
    If f6010_ <> 0 Then
     rDi(aktDiNr).obKasse = 1 ' Einzeldiagnosen werden immer übermittelt
     rDi(aktDiNr).lKasse = MAX(rDi(aktDiNr).lKasse, messDatum)
    End If
    rDi(aktDiNr).f6011 = f6011_
    rDi(aktDiNr).FID = rFa(UBound(rFa)).FID
    rDi(aktDiNr).Pat_id = rNa(0).Pat_id
'    rDi(aktDiNr).GesName = GN
    rDi(aktDiNr).absPos = absPos
    rDi(aktDiNr).AktZeit = AktZeit
    If obDStr Then
     Diag(DiagNr) = DText_
     ICD(DiagNr) = ICD_
     For j = Len(DText_) To 1 Step -1
      If Mid$(Diag(DiagNr), j, 1) = "(" Then
       Dim IcdRoh$
'     IF Pat_id = 33 THEN Halt
       If Len(Diag(DiagNr)) >= j Then
        IcdRoh = vNS
       Else
        If IsNumeric(Mid$(IcdRoh, 2, 1)) Then
         ICD(DiagNr) = IcdRoh
         Diag(DiagNr) = Left$(Diag(DiagNr), j - 1)
        End If
       End If
       Exit For
      End If ' Mid$(Diag(DiagNr), j, 1) = "(" Then
     Next j
     DSic(DiagNr) = DSic_
     DiagSe(DiagNr) = DSe_
     DiagAttr(DiagNr) = DAt_
     DiagAus(DiagNr) = DAus_
     DiagiBm(DiagNr) = DiBm_
     diagdat(DiagNr) = messDatum
     obDauer(DiagNr) = obD_
     If f6010_ <> 0 Or obD_ = 0 Then
'      If obD_ Then
       obKasse(DiagNr) = 1 ' Einzeldiagnosen werden immer übermittelt
       lKasse(DiagNr) = MAX(lKasse(DiagNr), messDatum)
'      End If
     End If
     f6010(DiagNr) = f6010_
     f6011(DiagNr) = f6011_
     DiagNr = DiagNr + 1
    End If ' obDStr
difertig:
    DText_ = ""
    ICD_ = ""
    DSic_ = ""
    DSe_ = ""
    DAt_ = ""
    DAus_ = ""
    DiBm_ = ""
    obD_ = 0
    f6010_ = 0
    f6011_ = ""
   Case 6201 ' Uhrzeit
    messDatum = messDatumD + BDTtoTime(rInhalt)
    If messDatum > rNa(0).lAktTM Then rNa(0).lAktTM = messDatum
   Case 6203 ' Formularunterteilung s. 6340
'     511533, '6203', 1634, 2, 'TM#?##'
'         27, '6203', 1634, 2, 'TM#Anamnese##Koerpermasse'
'      13122, '6203', 1634, 2, 'TM#Befund##Blutdruck'
    Select Case rInhalt
     Case "TM#?##":                    FormUnt = Fgz
     Case "TM#Anamnese##Koerpermasse": FormUnt = AnamnKörperMasse
     Case "TM#Befund##Blutdruck":      FormUnt = BefundBlutdruck
    End Select
    If Left$(rInhalt, 10) = "TM#Befund#" And InStr(rInhalt, "#Schwangerschaft") <> 0 Then FormUnt = Schwangerschaft
   Case 6340 ' internes XML-Formular
    Dim Wort$, Diast$, irunde%, nachstr$, umzuw$
    If InStrB(rInhalt, "KarteiSchwangerschaftsEintrag") <> 0 Then ' 9.4.22: es können auch Blutdruck und Sws in einer Zeile stehen, Pat. 64124
      ReDim Preserve rSw(UBound(rSw) + 1)
      rSw(UBound(rSw)).Pat_id = rNa(0).Pat_id
      rSw(UBound(rSw)).Zeitpunkt = messDatum
      rSw(UBound(rSw)).absPos = absPos
      rSw(UBound(rSw)).AktZeit = AktZeit
      rSw(UBound(rSw)).FID = rFa(UBound(rFa)).FID
      For irunde = 1 To 9
       Select Case irunde
        Case 1: nachstr = "voraussichtlicher ET:"
        Case 2: nachstr = "EndeDatum>"
        Case 3: nachstr = "EffektiveLetzteRegel>"
        Case 4: nachstr = "EndeArt>"
        Case 5: nachstr = "ErfassteLetzteRegel>"
        Case 6: nachstr = "KorrigierterGeburtstermin>"
        Case 7: nachstr = "MutterschutzBeginn>"
        Case 8: nachstr = "anzeigeText>LR:"
        Case 9: nachstr = "~ Datum:"
       End Select
       p1 = InStr(rInhalt, nachstr)
       If (p1 > 0) Then
        p2 = InStr(p1, rInhalt, "<")
'        IF irunde >= 1 THEN
         umzuw = Mid$(rInhalt, p1 + Len(nachstr), p2 - p1 - Len(nachstr))
         If umzuw <> "" Then
          Select Case irunde
           Case 1:
            rSw(UBound(rSw)).vorET = CDate(umzuw)
            rFa(UBound(rFa)).vorET = rSw(UBound(rSw)).vorET
           Case 2:
            rSw(UBound(rSw)).ED = BDTtoDate(umzuw)
            rFa(UBound(rFa)).vorET = rSw(UBound(rSw)).ED
           Case 3:
            rSw(UBound(rSw)).efLR = BDTtoDate(umzuw)
           Case 4:
            rSw(UBound(rSw)).EndeArt = umzuw
           Case 5:
            rSw(UBound(rSw)).erLR = BDTtoDate(umzuw)
           Case 6:
            rSw(UBound(rSw)).kGT = BDTtoDate(umzuw)
           Case 7:
            rSw(UBound(rSw)).MB = BDTtoDate(umzuw)
           Case 8:
            If InStrB(umzuw, ";") <> 0 Then umzuw = Trim$(Left$(umzuw, InStr(umzuw, ";") - 1))
            rSw(UBound(rSw)).lR = IIf(umzuw = "", 0, umzuw)
           Case 9:
            rSw(UBound(rSw)).ET = CDate(umzuw)
            rSw(UBound(rSw)).vorET = rSw(UBound(rSw)).ET ' von mir dazu erfunden
            rFa(UBound(rFa)).vorET = rSw(UBound(rSw)).ET
          End Select
'          Exit For
         End If
'        Else
'         umzuw = Mid$(rInhalt, p1 + Len(nachstr) + 1, p2 - p1 - Len(nachstr) - 1)
'         IF umzuw <> "" THEN
'          rSw(UBound(rSw)).vorET = CDate(umzuw)
'          rFa(UBound(rFa)).vorET = rSw(UBound(rSw)).vorET
''          Exit For
'         END IF
'        END IF
'       Else
'        p1 = InStr(rInhalt, "Entbindung ~ SSW:")
'        IF (p1 > 0) THEN
'         Const nachstrd$ = "~ Datum:"
'         p2 = InStr(p1, rInhalt, nachstrd)
'         umzuw = Mid$(rInhalt, p2 + Len(nachstrd) + 1, 10)
'         IF umzuw <> "" THEN
'          rSw(UBound(rSw)).ET = CDate(umzuw)
'          rFa(UBound(rFa)).vorET = rSw(UBound(rSw)).ET
''          Exit For
'         END IF
'        END IF
       End If
      Next irunde
    End If ' InStrB(rInhalt, "KarteiSchwangerschaftsEintrag") <> 0 THEN
    Select Case FormUnt
'     Case Schwangerschaft
     Case AnamnKörperMasse ' dann einträge machen
      If InStrB(rInhalt, "KarteiKoerpermassEintrag") <> 0 Then ' Stand 8.12.10 Tautologie
     'TM#<KarteiKoerpermassEintrag Id="0-2014164"><anzeigeText>BMI: 29,4 (Übergewicht)\nGröße: 173,0 cm\nGewicht: 88,000 kg</anzeigeText><attribute><GewichtInG>88000</GewichtInG><GroesseInMM>1730</GroesseIn
       For irunde = 1 To 4
        Select Case irunde
         Case 1: Wort = "GewichtInG"
         Case 2: Wort = "GroesseInMM"
         Case 3: Wort = "BauchumfangInMM"
         Case 4: Wort = "HueftumfangInMM"
        End Select
        p1 = InStr(rInhalt, "<" & Wort & ">")
        p2 = InStr(rInhalt, "</" & Wort & ">")
        If p1 <> 0 And p2 <> 0 Then
         p1 = p1 + 2 + Len(Wort)
         Diast = Mid$(rInhalt, p1, p2 - p1)
         If LenB(Diast) <> 0 Then
          GoSub rEiVorb
          Select Case irunde
           Case 1: rEi(UBound(rEi)).art = "gewicht": rEi(UBound(rEi)).Inhalt = Left$(Diast, MAX(Len(Diast) - 3, 0)) & " kg" ' g -> kg
           Case 2: rEi(UBound(rEi)).art = "groesse": rEi(UBound(rEi)).Inhalt = Left$(Diast, Len(Diast) - 1) & " cm" ' mm -> cm
           Case 3: rEi(UBound(rEi)).art = "taille":  rEi(UBound(rEi)).Inhalt = Left$(Diast, Len(Diast) - 1) & " cm" ' mm -> cm
           Case 4: rEi(UBound(rEi)).art = "hüfte":   rEi(UBound(rEi)).Inhalt = Left$(Diast, Len(Diast) - 1) & " cm" ' mm -> cm
          End Select
         End If
        End If ' p1 <> 0 AND p2 <> 0 THEN
       Next irunde
       If False Then
rEiVorb:
        ReDim Preserve rEi(UBound(rEi) + 1)
'        rEi(UBound(rEi)).Art = rInhalt
        rEi(UBound(rEi)).absPos = absPos
        rEi(UBound(rEi)).AktZeit = AktZeit
        rEi(UBound(rEi)).FID = rFa(UBound(rFa)).FID
        rEi(UBound(rEi)).Pat_id = rNa(0).Pat_id
        rEi(UBound(rEi)).Zeitpunkt = messDatum
        rEi(UBound(rEi)).QS = ZQSort(rEi(UBound(rEi)).Zeitpunkt)
        rEi(UBound(rEi)).QT = ZQuart(rEi(UBound(rEi)).Zeitpunkt)
        Return
       End If
      End If 'InStrB(rInhalt, "KarteiKoerpermassEintrag
     Case BefundBlutdruck
' 11.3.15: liefert, z.B. bei Pat. 60147, falsche Ergebnisse, weil RR z.B. perpetuiert, deshalb wegen Tautotologie auskommentiert
' 25.9.16: doch wieder aktiviert, da sonst Information "Mittelwert der letzten ..." verloren geht
'#If False THEN
      If InStrB(rInhalt, "KarteiZusatzdatenEintrag") <> 0 Then ' Stand 8.12.10 Tautologie
      Call aufSplit(rInhalt, "</KarteiZusatzdatenEintrag>")
      Dim jsp%
      For jsp = 0 To ArraInd
          If Arra(jsp) = "" Then Exit For
'          IF InStrB(Arra(jsp), "0-23177064") <> 0 THEN Stop
'          IF InStrB(Arra(jsp), "0-23177069") <> 0 THEN Stop
           ReDim Preserve rRr(UBound(rRr) + 1)
           ausrrxml = True
           rRr(UBound(rRr)).Pat_id = rNa(0).Pat_id
           rRr(UBound(rRr)).Zeitpunkt = messDatum
    '       rRr(UBound(rRr)).RR = RR
           rRr(UBound(rRr)).absPos = absPos
           rRr(UBound(rRr)).AktZeit = AktZeit
           rRr(UBound(rRr)).FID = rFa(UBound(rFa)).FID
    '       Call RREintr
           For irunde = 1 To 5
            Select Case irunde
             Case 1: Wort = "SystolischerWert"
             Case 2: Wort = "DiastolischerWert"
             Case 3: Wort = "Puls"
             Case 4: Wort = "Quelle"
             Case 5: Wort = "Kommentar"
            End Select
            p1 = InStr(Arra(jsp), "<" & Wort & ">")
            p2 = InStr(Arra(jsp), "</" & Wort & ">")
            If p1 <> 0 And p2 <> 0 Then
             p1 = p1 + 2 + Len(Wort)
             Diast = Mid$(Arra(jsp), p1, p2 - p1)
             Select Case irunde
              Case 1: rRr(UBound(rRr)).RR = Diast
              Case 2: rRr(UBound(rRr)).RR = rRr(UBound(rRr)).RR & "/" & Diast
              Case 3: rRr(UBound(rRr)).Puls = IIf(LenB(Diast) = 0, "0", Diast)
              Case 4: rRr(UBound(rRr)).Quelle = Diast
              Case 5: rRr(UBound(rRr)).Bemerkung = Diast
             End Select
            End If ' p1 <> 0 AND p2 <> 0 THEN
           Next irunde
           Dim jrue%
           For jrue = UBound(rRr) - 1 To UBound(rRr) - 4 Step -1
            If jrue < 0 Then Exit For
            If rRr(jrue).RR = rRr(UBound(rRr)).RR And rRr(jrue).Bemerkung = rRr(UBound(rRr)).Bemerkung Then
             ReDim Preserve rRr(UBound(rRr) - 1)
            End If
           Next jrue
       Next jsp
      End If ' InStrB(rInhalt, "KarteiZusatzdatenEintrag") <> 0
'#END IF
    End Select ' FormUnt
   Case 6218 ' Rezepteintrag, z.B. HeilHilfsmittel, LangRezeptEintrag, Sprechstundenbedarf-Rezept
    Call RezEintr(rInhalt, True)
   Case 6220 ' Symptome, bisher nur bei Musterfrau
   Case 6230 ' Blutdruck
    If Left$(rInhalt, 3) = "RR " Then ' 8.12.10: dann redundant ' 11.3.15 umgruppiert
     rInhalt = Mid$(rInhalt, 3)
    End If
    rInhalt = LTrim$(rInhalt)
    Dim nichteintrag%
    If ausrrxml And messDatum = rRr(UBound(rRr)).Zeitpunkt Then
      nichteintrag = 1
    ' 25.9.16: in KarteiZusatzdatenEintrag steht der falsche Blutdruck
      Dim posp%
      posp = InStr(rInhalt, "P")
      If posp > 0 Then
       If posp < Len(rInhalt) Then rRr(UBound(rRr)).Puls = MachNumerisch(Mid$(rInhalt, posp + 1))
       If posp > 1 Then rRr(UBound(rRr)).RR = Left$(rInhalt, posp - 1)
      Else
       rRr(UBound(rRr)).RR = rInhalt
'       rRr(UBound(rRr)).Bemerkung = "" ' 21.9.20; stimmt doch nicht
      End If
    End If
    If nichteintrag = 0 Then Call RREintr(rInhalt)
    DStr = " (" + Format$(messDatum, "dd/mm/yy") + ")"
    If arr = vNS Then arr = rInhalt & DStr
    lRR = rInhalt & DStr
    gesRR = gesRR & IIf(LenB(gesRR) = 0, vNS, ", ") & rInhalt & DStr
    ausrrxml = 0
  Case 6260 ' Therapie, bei uns: Eintrag Tamara Hartmann
   Case 6260 ' Therapie, bei uns Eintrag Tamara Hartmann
    If lKennung = 6280 Then
     rEi(UBound(rEi)).Inhalt = Left$(rEi(UBound(rEi)).Inhalt, Len(rEi(UBound(rEi)).Inhalt) - IIf(Right$(rEi(UBound(rEi)).Inhalt, 1) = "^", 1, 0)) + rInhalt
    Else
     ReDim Preserve rEi(UBound(rEi) + 1)
     rEi(UBound(rEi)).art = "th"
     rEi(UBound(rEi)).absPos = absPos
     rEi(UBound(rEi)).AktZeit = AktZeit
     rEi(UBound(rEi)).FID = rFa(UBound(rFa)).FID
     rEi(UBound(rEi)).Pat_id = rNa(0).Pat_id
     rEi(UBound(rEi)).Zeitpunkt = messDatum
     rEi(UBound(rEi)).QS = ZQSort(rEi(UBound(rEi)).Zeitpunkt)
     rEi(UBound(rEi)).QT = ZQuart(rEi(UBound(rEi)).Zeitpunkt)
     rEi(UBound(rEi)).Inhalt = rInhalt
    End If
   Case 6280 ' Laboranforderungstext
    If lKennung = 6280 Then
     rLb(UBound(rLb)).AnfText = Left$(rLb(UBound(rLb)).AnfText, Len(rLb(UBound(rLb)).AnfText) - IIf(Right$(rLb(UBound(rLb)).AnfText, 1) = "^", 1, 0)) + rInhalt
    Else
     ReDim Preserve rLb(UBound(rLb) + 1)
     rLb(UBound(rLb)).Pat_id = rNa(0).Pat_id
     rLb(UBound(rLb)).Zeitpunkt = messDatum
     rLb(UBound(rLb)).AnfText = rInhalt
     rLb(UBound(rLb)).FID = rFa(UBound(rFa)).FID
     rLb(UBound(rLb)).absPos = absPos
     rLb(UBound(rLb)).AktZeit = AktZeit
    End If
   Case 6285 ' AU-Dauer
    ReDim Preserve rAu(UBound(rAu) + 1)
    rAu(UBound(rAu)).Pat_id = rNa(0).Pat_id
    rAu(UBound(rAu)).Zeitpunkt = messDatum
    rAu(UBound(rAu)).Beginn = Left$(rInhalt, 8)
    rAu(UBound(rAu)).Ende = Right$(rInhalt, 8)
    rAu(UBound(rAu)).FID = rFa(UBound(rFa)).FID
    rAu(UBound(rAu)).AktZeit = AktZeit
   Case 6286 ' AU-Begründung
    If rAu(UBound(rAu)).ICDs = "" Then
     rAu(UBound(rAu)).ICDs = rInhalt
    Else
     rAu(UBound(rAu)).ICDs = rAu(UBound(rAu)).ICDs & " " & rInhalt
    End If
   Case 6291 ' Krankenhauseinweisung
    ReDim Preserve rKh(UBound(rKh) + 1)
    rKh(UBound(rKh)).Pat_id = rNa(0).Pat_id
    rKh(UBound(rKh)).Zeitpunkt = messDatum
    rKh(UBound(rKh)).Ziel = rInhalt
    rKh(UBound(rKh)).absPos = absPos
    rKh(UBound(rKh)).FID = rFa(UBound(rFa)).FID
    rKh(UBound(rKh)).AktZeit = AktZeit
   Case 6295 ' Formularabkürzung
    FormAbk = rInhalt
'    If FormAbk Like "uew" Then Stop
    Select Case rInhalt
     Case "DMPDTYP2", "DMPDTYP1", "DMPKHK", "DMPAB", "DMPCOPD", "DMPB2", "DMPK2", _
          "eDMPDM2", "eDMPDM1", "eDMPKHK", "eDMPAB", "eDMPCOPD", "eDMPB2", "eDMPK2"
      obFormDMP = True
      ReDim Preserve rDm(UBound(rDm) + 1)
      rDm(UBound(rDm)).Abk = rInhalt
      rDm(UBound(rDm)).Pat_id = rNa(0).Pat_id
      rDm(UBound(rDm)).KarteiDatum = messDatum
      rDm(UBound(rDm)).AktZeit = AktZeit
      rDm(UBound(rDm)).lanrid = Lanr ' : Lanr = 0 ' Kommentar 21.3.21
'      rDm(UBound(rDm)).NachName = rNa(0).NachName
'      rDm(UBound(rDm)).VorName = rNa(0).VorName
'      rDm(UBound(rDm)).GebDat = rNa(0).GebDat
     Case "TUG", "ADL" ' Timed Up AND Go - Test, Activity of Daily Life
        ReDim Preserve rEi(UBound(rEi) + 1)
        rEi(UBound(rEi)).art = rInhalt
        rEi(UBound(rEi)).absPos = absPos
        rEi(UBound(rEi)).AktZeit = AktZeit
        rEi(UBound(rEi)).FID = rFa(UBound(rFa)).FID
        rEi(UBound(rEi)).Pat_id = rNa(0).Pat_id
        rEi(UBound(rEi)).Zeitpunkt = messDatum
        rEi(UBound(rEi)).QS = ZQSort(rEi(UBound(rEi)).Zeitpunkt)
        rEi(UBound(rEi)).QT = ZQuart(rEi(UBound(rEi)).Zeitpunkt)
     Case Else
      obFormDMP = False
    End Select ' case rInhalt
   Case 6296 ' Formular-Bezeichnung
    FormBez = rInhalt
'    If FormBez = "Überweisung gezielt an: Neurologie" Then Stop
'    IF FormBez = "Medikamentenplan" THEN
'     12.6.17: beide Variablen werden nicht verwendet
'     IF obMedPlanGelesen THEN
'      obMedPlan = False
'     Else
'      obMedPlan = True
'      obMedPlanGelesen = True
'     END IF
'   ElseIf ...
'
' Falsche Formbez ermitteln:
' SELECT (select COUNT(0) from forminhkopf where form_id=f.formid) zahl, (select MAX(pat_id) from forminhkopf where form_id=f.formid) pat_id, (select MAX(zeitpunkt) from forminhkopf where form_id=f.formid) zp, f.* FROM formulare f WHERE formbez='Privatrezept';
'    IF FormBez = "Lang-Rezept" OR FormBez = "Rezept" OR FormBez = "Privatrezept" THEN
' Rezept: formvorl RLIKE 'kassenrezept' AND formvorl NOT RLIKE 'sprechstundenbedarf'
'         formvorl RLIKE 'rezeptlang'
'         formvorl RLIKE 'privatrezept'
    If InStrB(FormBez, "ezept") <> 0 Then
     ReDim auti(3) ' Pat. 11081
     ReDim anzl(3)
     mdnr = 0
    ElseIf Left$(rInhalt, 7) = "CGM BMP" Or Left$(rInhalt, 30) = "Elektronischer Medikationsplan" Then
     FormBez = "cgm_bmp"
    ElseIf obFormDMP Then
     If InStrB(rInhalt, "(unvollst") > 0 Then
      rDm(UBound(rDm)).obvoll = 0
     ElseIf InStrB(rInhalt, "(vollst") > 0 Then
      If VarType(rDm(UBound(rDm)).obvoll) = 17 Then rDm(UBound(rDm)).obvoll = 1 Else rDm(UBound(rDm)).obvoll = -1
     End If
     If InStrB(rInhalt, "Erst") > 0 Then
      rDm(UBound(rDm)).art = "ED"
     ElseIf InStrB(rInhalt, "Folge") > 0 Or InStrB(rInhalt, "Verlaufs") > 0 Then
      rDm(UBound(rDm)).art = "FD"
     End If
     If InStrB(rInhalt, "(ok") <> 0 Then
      rDm(UBound(rDm)).Ok = True
     End If
     If InStrB(rInhalt, "ausgedruckt)") <> 0 Then
      rDm(UBound(rDm)).ausgedruckt = True
     End If
     Dim exppos%, expdat$
     exppos = InStr(rInhalt, "exportiert am")
     If exppos > 0 Then
      expdat = Mid$(rInhalt, exppos + 14, 10)
      rDm(UBound(rDm)).exportiert = CDate(expdat)
     End If
    End If
   Case 6297 ' Formular-Vorlage (Windows-Pfad)
    FormVorl = rInhalt 'replace$(RInhalt, "€", " ")
' Kommt irgendwie nicht mehr vor 13.5.23
    If FormBez = "Medikamentenplan" Or FormBez = "cgm_bmp" Then
' "Medikamentenplan"_1
     MedZahl = -1
     obHergang = 0
'     IF pmpnr = 0 THEN ' wenn zwei Einlesungen gleichzeitig stattfinden ...
'      IF Not rsAdo Is Nothing THEN IF rsAdo.State = 1 THEN rsAdo.Close
''      myFrag rsAdo, "SELECT (MAX(mpnr) + 1) AS mpnr FROM `medplan`"
'      SET rsAdo = myEFrag("SELECT (MAX(mpnr) + 1) AS mpnr FROM `medplan`")
'      mpnr = IIf(ISNULL(rsAdo!mpnr), 1, rsAdo!mpnr)
'     Else
      pMpnr = pMpnr + 1
'     END IF
    Else ' FormBez = "Medikamentenplan"
     rFoNeu = -1
     For i = 1 To UBound(rFo)
      If rFo(i).Form_Abk = FormAbk And rFo(i).FormBez = FormBez Then
       If doUmwfSQL(rFo(i).FormVorl, lies.obMySQL) = FormVorl Or rFo(i).FormVorl = FormVorl Then
        rFoNeu = 0
        lFormID = rFo(i).FormID
        Exit For
       Else
'        err.raise 999,,"Fehler bei dolies"
       End If
      End If
     Next i
     If rFoNeu And Not obmitFormularen Then
      Dim rsf As ADODB.Recordset
      Call myFrag(rsf, "SELECT * FROM formulare WHERE Form_Abk='" & FormAbk & "' AND FormBez='" & FormBez & "' AND FormVorl='" & FormVorl & "' ORDER BY FormID", adOpenStatic)
      If Not rsf.BOF Then
       rFoNeu = 0
       ReDim Preserve rFo(UBound(rFo) + 1)
       rFo(UBound(rFo)).absPos = rsf!absPos
       rFo(UBound(rFo)).AktZeit = rsf!AktZeit
       rFo(UBound(rFo)).Form_Abk = IIf(IsNull(rsf!Form_Abk), vNS, rsf!Form_Abk)
       rFo(UBound(rFo)).FormBez = rsf!FormBez
       rFo(UBound(rFo)).FormID = rsf!FormID
       lFormID = rsf!FormID
       rFo(UBound(rFo)).FormVorl = rsf!FormVorl
       rFo(UBound(rFo)).StByte = rsf!StByte
      End If
     End If
     If rFoNeu Then
      ReDim Preserve rFo(UBound(rFo) + 1)
      rFo(UBound(rFo)).Form_Abk = FormAbk
      rFo(UBound(rFo)).absPos = absPos
      rFo(UBound(rFo)).AktZeit = AktZeit
      rFo(UBound(rFo)).FormBez = FormBez
      rFo(UBound(rFo)).FormID = rFo(UBound(rFo) - 1).FormID + 1 ' muss noch in `forminhkopf` angpasst werden, wird deshalb dort negativ gespeichert
      lFormID = rFo(UBound(rFo)).FormID
      rFo(UBound(rFo)).FormVorl = FormVorl
      rFo(UBound(rFo)).StByte = AktByte
     End If ' rFoNeu Then
     rFoAbkNeu = -1
     For i = 1 To UBound(rFi)
      If UCase$(rFi(i).Form_Abk) = UCase$(FormAbk) Then
       rFoAbkNeu = 0
       Exit For
      End If
     Next i
     If rFoAbkNeu Then
      ReDim Preserve rFi(UBound(rFi) + 1)
      rFi(UBound(rFi)).Form_Abk = FormAbk
     End If
     rFm_Nr = 0
     jetztKopf = True
    End If ' FormBez = "Medikamentenplan"
   Case 6298 ' Formularfeld und ggf. Spalte
    FormSp = rInhalt ' nur für DMP-Bogen und Kasse benötigt, auch TUG und ADL
    Call aufSplit(rInhalt)
#If altMed Then
    If FormBez = "Medikamentenplan" Then
' "Medikamentenplan"_2
'     MedArt = Arra(0)
     MedNr = -1
'     ReDim Med$(0)
'     ReDim Dos$(5, 0)
     Select Case Arra(0)
      Case "medikament"
       MedNr = Arra(1)
       While CLng(Arra(1)) > MedZahl
        Call MPerg
       Wend
      Case "Hergang" ' 21.1.12
        If MedZahl = -1 Then
         Call MPerg
        End If
      Case Else
'       Debug.Print Arra(0)
'       Stop
     End Select ' arra(0)
#Else
    If FormBez = "Medikamentenplan" Then
' "Medikamentenplan"_2 neu
#End If
    ElseIf FormBez = "cgm_bmp" Then
     If FormSp = "MedPlan" Then jetztCGM = True
    Else
     If jetztKopf Then
      ReDim Preserve rFr(UBound(rFr) + 1)
      rFr(UBound(rFr)).absPos = absPos
      rFr(UBound(rFr)).AktZeit = AktZeit
      rFr(UBound(rFr)).FID = rFa(UBound(rFa)).FID
      rFr(UBound(rFr)).Form_ID = lFormID '-lFormID ' negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFo angepaßt werden muss
      rFr(UBound(rFr)).Pat_id = rNa(0).Pat_id
      rFr(UBound(rFr)).StByte = AktByte
      rFr(UBound(rFr)).Zeitpunkt = messDatum
      rFr(UBound(rFr)).Satzart = f8000
      rFr(UBound(rFr)).lanrid = Lanr ' :    Lanr = 0 ' Kommentar 21.3.21
      f8000 = vNS
      rFr(UBound(rFr)).Satzlänge = f8100
      f8100 = vNS
      If FoIDv = 0 Then
       If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
       Set rsAdo = myEFrag("SELECT (MAX(foid) + 1) AS mfoid FROM `forminhkopf`")
       FoIDv = IIf(IsNull(rsAdo!mfoid), 0, rsAdo!mfoid)
      End If
      rFr(UBound(rFr)).Foid = FoIDv ' Pseudo-Foid
      FoIDv = FoIDv + 1
      rFm_Nr = rFm_Nr + 1
      jetztKopf = False
     End If ' jetztKopf
    End If ' Formbez
    ZeileNr = 0
   Case 6299 ' Formularfeldinhalt auch Überweisungstext
    FormInh = rInhalt
'    IF InStr(FormInh, "SVA") THEN Stop
    Select Case FormBez
     Case "Lang-Rezept"
      If FormInh = "X" And Arra(0) = "Check2" Then
        Select Case Arra(1)
         Case "4"
          auti(0) = 1
         Case "5"
          auti(1) = 1
         Case "6"
          auti(2) = 1
        End Select
       End If
     Case "Rezept"
      Select Case Arra(0)
       Case "IV_MED_AUTIDEM"
        auti(Arra(1) - 1) = FormInh
       Case "IV_MED_ANZAHL"
        anzl(Arra(1) - 1) = FormInh
      End Select
    End Select
    Select Case FormSp
     Case "Kasse"
       If IKneu And Not (rFa(UBound(rFa)).SchGr <> 90 And InStrB(LCase$(rInhalt), "privat") <> 0) Then
        rFa(UBound(rFa)).Kasse = FormInh ' Kasse aus Formularen gewinnen
       End If
       If rFa(UBound(rFa)).Kasse = "" Then
        If IKrz <> "" Then
         If InStr(IKrz, rFa(UBound(rFa)).IK) Then
          rFa(UBound(rFa)).Kasse = FormInh
         End If
        End If
       End If
       IKrz = ""
     Case "Ik"
      IKrz = rInhalt
      If rFa(UBound(rFa)).IK = "" Then
       IKneu = True
       rFa(UBound(rFa)).IK = Mid$(FormInh, 3)
      Else
       IKneu = 0
      End If
     Case "Status"
      Select Case FormAbk ' Status aus Formularen ziehen
       Case "rp", "uew", "lar", "spbrp", "uewL", "AnLU", "DMPDTYP2", "DMPDTYP1", "eDMPDM2", "eDMPDM1"
        If Not Len(FormInh) < Len(rFa(UBound(rFa)).Status) Then
         rFa(UBound(rFa)).Status = FormInh
        End If
      End Select
    End Select
    Select Case FormAbk
     Case "TUG" ' Time up AND go
       If FormSp = "sek" Then rEi(UBound(rEi)).Inhalt = FormInh & " " & FormSp
     Case "ADL" ' Barthel-Index
       If FormSp = "txtdyn#37" Then rEi(UBound(rEi)).Inhalt = "Gesamtpunktzahl " & FormInh
    End Select
    If obFormDMP Then
     If FormInh Like "?" Then
     ElseIf FormSp = "ZusatzDaten" Then
      rDm(UBound(rDm)).Zusatzdaten = FormInh
     ElseIf FormSp = "Nachname" Then
      rDm(UBound(rDm)).Nachname = FormInh
     ElseIf FormSp = "Vorname" Then
      rDm(UBound(rDm)).Vorname = FormInh
     ElseIf FormInh Like "*.*.*" Then
' Eigentlich heißen die relevanten Felder "Datum" bzw. "Erstellungsdatum". Die Namen können aber bei Download von falschem Computer aus vertauscht sein:
' -> allgemeinere Methode muß gewählt werden
' das erste Datenfeld mit Datumsformatierung entspricht bei richtigem Formular dem Feld "Datum"
      If rDm(UBound(rDm)).DokuDatum = 0 Then
       If Len(FormInh) > 1 Then
        If IsDate(FormInh) Then
         rDm(UBound(rDm)).DokuDatum = FormInh
        End If
       End If
      ElseIf rDm(UBound(rDm)).GebDat = 0 Then
       If Len(FormInh) > 1 Then
        If IsDate(FormInh) Then
         rDm(UBound(rDm)).GebDat = FormInh
        End If
       End If
      End If
     End If
    End If ' obFormDMP
    If jetztCGM Then
     jetztCGM = 0
'     Call aufSplit(rInhalt, "<M ")
' Korrektur wegen mehrseitigen Plänen 8.1.18
     Dim Pos1&, pos2&
     pos2 = 1: ArraInd = 0
     Do
      Pos1 = InStr(pos2, rInhalt, "<M ") + 3
      If Pos1 = 3 Then Exit Do
      pos2 = InStr(Pos1, rInhalt, "/>")
      If pos2 Then
       ArraInd = ArraInd + 1
       ReDim Preserve Arra(ArraInd)
       Arra(ArraInd) = Mid$(rInhalt, Pos1, pos2 - Pos1)
       Debug.Print Pos1, pos2, Arra(ArraInd)
      Else
       pos2 = Pos1
      End If
     Loop
'     Dim FNam() AS String, ifn&
'     FNam = ASMod("a", "w", "p", "m", "d", "v", "h", "i", "r", "s", "t", "du", "fd")
     MedZahl = ArraInd
     Dim mnr%
     If ArraInd = 0 Then ' 4.4.20: leerer Medplan
       Call MPerg
     End If
     For mnr = 1 To ArraInd
      Arra(mnr) = " " & Arra(mnr)
      Call MPerg
      For ifn = 0 To UBound(FNam)
       Dim ps1%, ps2%, gef$
       ps1 = InStr(Arra(mnr), FNam(ifn))
       If (ps1 <> 0) Then
        ps1 = ps1 + Len(FNam(ifn)) ' 5 oder 6
        ps2 = InStr(ps1, Arra(mnr), "\""")
        gef = Mid$(Arra(mnr), ps1, ps2 - ps1)
        If gef <> "" Then
         Select Case ifn
          Case 3 ' "m"
           rMe(UBound(rMe)).mo = gef
          Case 4 ' "d"
           rMe(UBound(rMe)).mi = gef
          Case 5 ' "v"
           rMe(UBound(rMe)).ab = gef
          Case 6 ' "h"
           rMe(UBound(rMe)).Zn = gef
          Case 7 ' "i"
           rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung + " ") + gef
          Case 8 ' "r"
           rMe(UBound(rMe)).Grund = gef
          Case 9 ' "s"
           rMe(UBound(rMe)).Stärke = gef
'          Case 10 ' "t"
'           rMe(ubound(rme)).Bemerkung = IIf(rMe(ubound(rme)).Bemerkung = "", "", rMe(ubound(rme)).Bemerkung + " ") + gef
          Case 11 ' "du"
           rMe(UBound(rMe)).Einheit = gef
          Case 12 ' "fd"
           rMe(UBound(rMe)).Form = gef
          Case 6 ' "h"
           rMe(UBound(rMe)).Zn = gef
          Case 0 ' "a"
           rMe(UBound(rMe)).Medikament = gef
           rMe(UBound(rMe)).MedAnfang = GetMed(gef, 0).Value
          Case 2 ' "p"
           rMe(UBound(rMe)).PZN = gef
           If rMe(UBound(rMe)).Medikament = "" Then
            Dim rpzn As New ADODB.Recordset
            Set rpzn = Nothing
            myFrag rpzn, "SELECT Medikament,MedAnfang,Wirkstoff FROM quelle.medplan WHERE pzn='" & gef & "' AND COALESCE(medikament,'')<>'' LIMIT 1"
            If Not rpzn.BOF Then
             rMe(UBound(rMe)).Medikament = rpzn!Medikament
             rMe(UBound(rMe)).MedAnfang = rpzn!MedAnfang
             rMe(UBound(rMe)).Wirkstoff = rpzn!Wirkstoff
            Else ' Not rpzn.BOF THEN
             While Len(gef) < 8
              gef = "0" & gef
             Wend
             Set rsP = Nothing
             Dim medrunde%
             For medrunde = 0 To 1
              If medrunde = 0 Then On Error Resume Next Else On Error GoTo fehler
              myFrag rsP, "SELECT pck_rezeptinfo FROM `glalle` WHERE pck_pzn='" & gef & "'", , , , , , True, ErrNr, ErrDes
              If ErrNr = 0 Then Exit For
              Call DBCnOpen
             Next medrunde
             If Not rsP.BOF Then
              rMe(UBound(rMe)).Medikament = rsP.Fields(0)
              rMe(UBound(rMe)).MedAnfang = GetMed(rsP.Fields(0), 0).Value
'             rMe(UBound(rMe)).wirkstoff = rsP.Fields(1)
             End If ' not rsp.bof
            End If ' not rpzn.bof
           End If ' rMe(UBound(rMe)).Medikament = "" THEN
         End Select
        End If ' gef <> "" THEN
'        Debug.Print gef, FNam(ifn)
       End If ' (ps1 <> 0) THEN
      Next
     Next
' komplette Freitextzeilen
     Call aufSplit(rInhalt, "<X ")
     MedZahl = ArraInd
     For mnr = 1 To ArraInd
      Arra(mnr) = " " & Arra(mnr)
      For ifn = 0 To UBound(FNam2)
       ps1 = InStr(Arra(mnr), FNam2(ifn))
       If (ps1 <> 0) Then
        ps1 = ps1 + Len(FNam2(ifn)) ' 5 oder 6
        ps2 = InStr(ps1, Arra(mnr), "\""")
        gef = Mid$(Arra(mnr), ps1, ps2 - ps1)
        If gef <> "" Then
         Select Case ifn
          Case 0 ' "t"
'           rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung + " ") + gef
' 22.4.19: Bemerkungen sollen extra stehen => geändert in:
            Call MPerg
            rMe(UBound(rMe)).Bemerkung = gef
         End Select
        End If
'        Debug.Print gef, FNam(ifn)
       End If ' (ps1 <> 0) THEN
      Next ' ifn = 0 To UBound(FNam2)
     Next ' mnr = 1 To ArraInd
    ElseIf FormBez = "Medikamentenplan" Then
    Else
     ReDim Preserve rFm(UBound(rFm) + 1)
     rFm(UBound(rFm)).FeldInh = FormInh
     rFm(UBound(rFm)).Feld = Arra(0)
     If ArraInd > 0 Then
'      rFm(UBound(rFm)).FeldNr = Arra(1)
      If IsNumeric(Arra(1)) Then rFm(UBound(rFm)).FeldNr = Arra(1) ' 7.10.08, Pat. 1524
     Else
      rFm(UBound(rFm)).FeldNr = ZeileNr ' damit mehrere 6299 auf ein 6298 kommen können
     End If
     rFm(UBound(rFm)).nr = rFm_Nr
     rFm_Nr = rFm_Nr + 1
     rFm(UBound(rFm)).Foid = FoIDv - 1 '-FoID 'negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFr angepaßt werden muss
    End If
'    Call FormInhEintrag(rtr.AbsolutePosition)
' $\TurboMed\Formulare\Patientenmenue\DMP Erst-Dokumentation Diabetes mellitus Typ 2.tmf
' $\TurboMed\Formulare\Patientenmenue\eDMP Datenerfassung.tmf
    If FormBez Like "Erst-Dokumentation Diabetes mellitus Typ 2*" Then
' $\TurboMed\Formulare\Patientenmenue\DMP Folge-Dokumentation Diabetes mellitus Typ 2.tmf
    ElseIf FormBez Like "Folge-Dokumentation Diabetes mellitus Typ 2*" Then
    Else
     Select Case FormBez
      Case "cgm_bmp"
        If FormSp = "CreationDate" Then
         MPDatum = FormInh
        End If
#If altMed Then
' "Medikamentenplan"_3
      Case "Medikamentenplan"
       If Arra(0) = "Datum" Then
        MPDatum = FormInh
        For i = 0 To MedZahl
         rMe(UBound(rMe) - i).Datum = MPDatum
        Next i
       ElseIf MedNr <> -1 Then
        ' Arra(0)="medikament"
        rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).Medikament = FormInh
        rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).MedAnfang = GetMed(FormInh, 0).Value
       ElseIf MedZahl <> -1 Or Arra(0) = "Check1" Or Arra(0) = "ab" Then
        If ArraInd > 0 Then
         While CLng(Arra(1)) > MedZahl
          Call MPerg
         Wend
        End If
        FormInh = REPLACE$(REPLACE$(REPLACE$(FormInh, "«", "½"), "ó", "¾"), "¬", "¼")
        Select Case Arra(0)
         Case "Check1"
          If FormInh = "X" Then rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).bBed = 1
         Case "Hergang"
          rMe(UBound(rMe) - MedZahl).Bemerkung = LTrim$(rMe(UBound(rMe) - MedZahl).Bemerkung + vbCrLf) + FormInh
         Case "mo"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).mo = FormInh
         Case "mi"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).mi = FormInh
         Case "nm"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).nm = FormInh
         Case "ab"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).ab = FormInh
         Case "zn"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).Zn = FormInh
         Case "Datum"
          MPDatum = FormInh
         Case "Nachname", "Vorname", "Name", "Stempel", "Datenfeld1602", "Datenfeld40€1€0", "Datenfeld40_1_0", "geb", "geb€0", "geb_0", "geb€0€0", "geb_0_0", "geb€0€0€0", "geb_0_0_0", "geb€0€0€0€0", "geb_0_0_0_0", "geb€0€0€0€0€0", "geb_0_0_0_0_0", "geb€0€0€1", "geb_0_0_1", "geb€0€1", "geb_0_1"
         Case Else
          Err.Raise 999, , "Fehler in dolies: Arra(0) im Medplan hat den unvorhergesehenen Wert: " & Arra(0)
        End Select
       ElseIf Arra(0) = "Hergang" Then
'        rMe(UBound(rMe)).Bemerkung = LTrim$(rMe(UBound(rMe)).Bemerkung + vbCrLf) + FormInh
       End If ' IF MedZahl <> -1 OR Arra(0) = "Check1" OR Arra(0) = "ab" THEN
#Else
' "Medikamentenplan"_3 neu
      Case "Medikamentenplan"
       Select Case Arra(0)
        Case "Nachname", "Vorname", "Name", "Stempel", "Datenfeld1602", "Datenfeld40€1€0", "Datenfeld40_1_0", "geb", "geb€0", "geb_0", "geb€0€0", "geb_0_0", "geb€0€0€0", "geb_0_0_0", "geb€0€0€0€0", "geb_0_0_0_0", "geb€0€0€0€0€0", "geb_0_0_0_0_0", "geb€0€0€1", "geb_0_0_1", "geb€0€1", "geb_0_1"
        Case "Datum"
         MPDatum = FormInh
         For i = 0 To MedZahl
          rMe(UBound(rMe) - i).Datum = MPDatum
         Next i
        Case "Hergang"
         If obHergang = 0 Then
          Call MPerg
          obHergang = 1
         End If
         rMe(UBound(rMe)).Bemerkung = LTrim$(rMe(UBound(rMe)).Bemerkung + vbCrLf) + FormInh
        Case "Check1", "ab", "medikament", "mi", "mo", "nm", "ab", "zn"
         If ArraInd > 0 Then
          While CLng(Arra(1)) > MedZahl - obHergang
           Call MPerg
          Wend
         End If
         FormInh = REPLACE$(REPLACE$(REPLACE$(FormInh, "«", "½"), "ó", "¾"), "¬", "¼")
         Dim mZeile%
         mZeile = UBound(rMe) - MedZahl + CLng(Arra(1))
         Select Case Arra(0)
          Case "Check1"
           If FormInh = "X" Then rMe(mZeile).bBed = 1
          Case "mo"
           rMe(mZeile).mo = FormInh
          Case "mi"
           rMe(mZeile).mi = FormInh
          Case "nm"
           rMe(mZeile).nm = FormInh
          Case "ab"
           rMe(mZeile).ab = FormInh
          Case "zn"
           rMe(mZeile).Zn = FormInh
          Case "medikament"
           rMe(mZeile).Medikament = FormInh
           rMe(mZeile).MedAnfang = GetMed(FormInh, 0).Value
          Case Else
           Err.Raise 999, , "Fehler in dolies: Arra(0) im Medplan hat den unvorhergesehenen Wert: " & Arra(0)
         End Select
        Case Else
       End Select
#End If
      Case "Privatrezept"
      Case "Rezept"
      Case "Überweisung Labor"
      Case "GDT"
      Case "Lang-Rezept"
      Case "Sprechstundenbedarf Medikamente"
      Case "Sprechstundenbedarf BTM"
      Case "Sprechstundenbedarf Heil- und Hilfsmittel"
      Case "Überweisung gezielt an: Augenheilkunde"
      Case "Überweisung gezielt an: Psychologie"
      Case "Überweisung gezielt an: Dermatologie"
      Case "Überweisung gezielt an: Psychotherapie"
      Case "Überweisung gezielt an: Gynäkologie"
      Case "Überweisung gezielt an: ->Psychotherapie"
      Case "Überweisung gezielt an: Neurologie"
      Case "Heilmittelverordnung Physikalische Therapie"
      Case "Überweisung gezielt an: Orthopädie"
      
      Case Else
       If obÜProt Then Print #322, "FormBez: " + FormBez + " in do_DatenEintragen bisher nicht berücksichtigt!"
'     Debug.Print "Formbezeichnung: " + FormBez
     End Select
    End If ' FormBez Like "Erst-Dokumentation Diabetes mellitus Typ 2*" Then else else else
    Select Case FormAbk
     Case "DMPDTYP2", "DMPDTYP1", "eDMPDM2", "eDMPDM1"
      If dmpdat = 0 Or dmpdat < messDatum Then dmpdat = messDatum
    End Select
    ZeileNr = ZeileNr + 1
   Case 6324 ' Brief mit Pfad
    Call aufSplit(rInhalt)
    ReDim Preserve rBr(UBound(rBr) + 1)
    rBr(UBound(rBr)).absPos = absPos
    rBr(UBound(rBr)).AktZeit = AktZeit
    If ArraInd > -1 Then rBr(UBound(rBr)).Pfad = Arra(0)
    If ArraInd > 0 Then rBr(UBound(rBr)).art = Arra(1)
    If ArraInd > 1 Then rBr(UBound(rBr)).name = Arra(2)
    If ArraInd > 2 Then rBr(UBound(rBr)).Typ = Arra(3)
    rBr(UBound(rBr)).FID = rFa(UBound(rFa)).FID
    rBr(UBound(rBr)).Pat_id = rNa(0).Pat_id
    rBr(UBound(rBr)).Zeitpunkt = messDatum
    rBr(UBound(rBr)).QT = ZQuart(rBr(UBound(rBr)).Zeitpunkt)
    rBr(UBound(rBr)).QS = ZQSort(rBr(UBound(rBr)).Zeitpunkt)
    
    Dim Datei$
    Datei = rBr(UBound(rBr)).Pfad
'   IF instrb(Datei, "\\\\") > 0 OR instrb(Datei, ":\\") > 0 THEN
    Datei = REPLACE$(Datei, "\\", "\")
    Datei = REPLACE$(Datei, "\\", "\")
'   END IF
'   Datei = replace$(Datei, "$\\TurboMed\\Dokumente", PcDokPfad)
'    Datei = replace$(lcase(Datei), "$\turbomed\dokumente\video\$\turbomed\dokumente", PcDokPfad)
'    Datei = replace$(lcase(Datei), "$\turbomed\dokumente", PcDokPfad)
'    Datei = replace$(lcase(Datei), "\linserv\daten\turbomed\dokumente", PcDokPfad)
    Do
     pos = InStr(3, LCase$(Datei), "\turbomed\dokumente")
     If pos = 0 Then pos = InStr(1, Datei, "\turbomed\dokumente", vbTextCompare)
     If pos = 0 Then Exit Do
     If pos = 2 Then Exit Do
     Datei = "$" & Mid$(Datei, pos)
    Loop
    If LenB(PcDokPfad) = 0 Then getDokPfad
    If FSO.FileExists(REPLACE$(LCase$(Datei), "$\turbomed\dokumente", PcDokPfad)) Then
     Set Fil = FSO.GetFile(REPLACE$(LCase$(Datei), "$\turbomed\dokumente", PcDokPfad))
     If lies.obMySQL Then
      rBr(UBound(rBr)).Pfad = REPLACE$(Datei, "\", "\\") ' fil.Path
     Else
      rBr(UBound(rBr)).Pfad = Datei ' fil.Path
     End If
     If Not Fil Is Nothing Then
      rBr(UBound(rBr)).DokGroe = Fil.size
      rBr(UBound(rBr)).DokAenD = Fil.DateLastModified
     End If
    Else
' Entwicklung: hier kann das Fehlt-Kennzeichen gesetzt werden und evtl.
'       rDo(UBound(rDo)).DokPfad = Datei 'stehen
    End If
    rBr(UBound(rBr)).Quelldatum = doQuelldatum(rBr(UBound(rBr)).name, rBr(UBound(rBr)).DokAenD) ' nach außen verschoben 22.4.23
   Case 6325 ' Dokumentpfad
    If Not obZweiteRunde Then
     ReDim Preserve rDo(UBound(rDo) + 1)
    End If
    rDo(UBound(rDo)).DokPfad = rDo(UBound(rDo)).DokPfad + Stutz(rInhalt)
    If Right$(rInhalt, 1) = "^" Then
     obZweiteRunde = -1
    Else
     obZweiteRunde = 0
     If LenB(PcDokPfad) = 0 Then getDokPfad
     Datei = rDo(UBound(rDo)).DokPfad
'    IF instrb(Datei, "\\\\") > 0 OR instrb(Datei, ":\\") > 0 THEN
     Datei = REPLACE$(Datei, "\\", "\")
     Datei = REPLACE$(Datei, "\\", "\")
'    END IF
'    Datei = replace$(Datei, "$\\TurboMed\\Dokumente", PcDokPfad)
'     Datei = replace$(lcase(Datei), "$\turbomed\dokumente\video\$\turbomed\dokumente", PcDokPfad)
'     Datei = replace$(lcase(Datei), "$\turbomed\dokumente", PcDokPfad)
'     Datei = replace$(lcase(Datei), "\linserv\daten\turbomed\dokumente", PcDokPfad)
     Do
      pos = InStr(3, LCase$(Datei), "\turbomed\dokumente")
      If pos = 0 Then pos = InStr(1, Datei, "\turbomed\dokumente", vbTextCompare)
      If pos = 0 Then Exit Do
      If pos = 2 Then Exit Do
      Datei = "$" & Mid$(Datei, pos)
     Loop
'     Datei = replace$(rDo(UBound(rDo)).DokPfad, "$\\TurboMed\\Dokumente", PcDokPfad)
     If FSO.FileExists(REPLACE$(LCase$(Datei), "$\turbomed\dokumente", PcDokPfad)) Then
      Set Fil = FSO.GetFile(REPLACE$(LCase$(Datei), "$\turbomed\dokumente", PcDokPfad))
      If lies.obMySQL Then
       rDo(UBound(rDo)).DokPfad = REPLACE$(Datei, "\", "\\") ' fil.Path
      Else
       rDo(UBound(rDo)).DokPfad = Datei ' fil.Path
      End If
      If Not Fil Is Nothing Then
       rDo(UBound(rDo)).DokGroe = Fil.size
       rDo(UBound(rDo)).DokAenD = Fil.DateLastModified
      End If
     Else
' Entwicklung: hier kann das Fehlt-Kennzeichen gesetzt werden und evtl.
'       rDo(UBound(rDo)).DokPfad = Datei 'stehen
     End If
    End If
   Case 6326 ' Dokumentart (Windows-Dateityp)
    rDo(UBound(rDo)).DokArt = rInhalt
   Case 6327, 6300 ' Dokumentname / 6300 = ohne Ortsangabe (11.12.06: 6300 kommt nicht (mehr) vor)
    rDo(UBound(rDo)).DokName = rDo(UBound(rDo)).DokName + Stutz(rInhalt)
    rDo(UBound(rDo)).absPos = absPos
    rDo(UBound(rDo)).AktZeit = AktZeit
    rDo(UBound(rDo)).FID = rFa(UBound(rFa)).FID
    rDo(UBound(rDo)).Pat_id = rNa(0).Pat_id
    rDo(UBound(rDo)).Zeitpunkt = messDatum
    rDo(UBound(rDo)).QS = ZQSort(rDo(UBound(rDo)).Zeitpunkt)
    rDo(UBound(rDo)).QT = ZQuart(rDo(UBound(rDo)).Zeitpunkt)
    rDo(UBound(rDo)).Quelldatum = doQuelldatum(rDo(UBound(rDo)).DokName, messDatum)
   Case 6330 ' eintraege: Art
    If rInhalt = "DiagTxt" Then
     obEinweisung = True
    Else
     obEinweisung = False
     ReDim Preserve rEi(UBound(rEi) + 1)
     rEi(UBound(rEi)).art = rInhalt
     rEi(UBound(rEi)).absPos = absPos
     rEi(UBound(rEi)).AktZeit = AktZeit
     rEi(UBound(rEi)).FID = rFa(UBound(rFa)).FID
     rEi(UBound(rEi)).Pat_id = rNa(0).Pat_id
     rEi(UBound(rEi)).Zeitpunkt = messDatum
     rEi(UBound(rEi)).QS = ZQSort(rEi(UBound(rEi)).Zeitpunkt)
     rEi(UBound(rEi)).QT = ZQuart(rEi(UBound(rEi)).Zeitpunkt)
     If rInhalt Like "*. sitzung*" Then DMSchL = DMSchL + 1
    End If
   Case 6331 ' eintraege: Inhalt
'    EintrInh = EintrInh + Stutz(RInhalt)
    If obEinweisung Then
     rKh(UBound(rKh)).Diagnose = rKh(UBound(rKh)).Diagnose + Stutz(rInhalt)
    Else '
    '     rEi(UBound(rEi)).Inhalt = LEFT(rEi(UBound(rEi)).Inhalt, Len(rEi(UBound(rEi)).Inhalt) - IIf(Right$(rEi(UBound(rEi)).Inhalt, 1) = "^", 1, 0)) + rInhalt
     If (rEi(UBound(rEi)).art = "Hüfte" And Left$(rInhalt, 12) = "Hüftumfang: ") Or _
        (rEi(UBound(rEi)).art = "Größe" And Left$(rInhalt, 7) = "Größe: ") Or _
        (rEi(UBound(rEi)).art = "Gewicht" And Left$(rInhalt, 9) = "Gewicht: ") Or _
        (rEi(UBound(rEi)).art = "taille" And Left$(rInhalt, 13) = "Bauchumfang: ") _
     Then
      ReDim Preserve rEi(UBound(rEi) - 1) ' beim internen Formular die nochmal einzelnen streichen
     Else
      If LCase$(rEi(UBound(rEi)).art) = "dak" Then ' 4.12.19: DAK-Datum dokumentieren
       If rNa(0).dakab = 0 Then
        If InStr(rInhalt, ",") > 30 Then
         If IsDate(REPLACE$(Mid$(rInhalt, 30, InStr(rInhalt, ",") - 30), ",", ".")) Then
          rNa(0).dakab = CDate(REPLACE$(Mid$(rInhalt, 30, InStr(rInhalt, ",") - 30), ",", "."))
         End If
        End If ' If InStr(rInhalt, ",") > 30 Then
       End If ' rNa(0).dakab = 0 Then
      End If ' LCase$(rEi(UBound(rEi)).art) = "dak" Then ' 4.12.19: DAK-Datum dokumentieren
'      IF InStr(rEi(UBound(rEi)).Inhalt, "KK-Info Tabak empf") <> 0 THEN Stop
      rEi(UBound(rEi)).Inhalt = rEi(UBound(rEi)).Inhalt + Stutz(rInhalt)
      Call PraxisHbA1c(rEi(UBound(rEi)).art, rEi(UBound(rEi)).Inhalt)
      If rEi(UBound(rEi)).art = "schul" And InStrB(rEi(UBound(rEi)).Inhalt, "schulungsverein") > 0 Then
       obMedNetz = True
      End If
      If rInhalt = "bfd" And lKennung = 6330 And InStrB(rInhalt, "icral") > 0 Then
       AlbErhöht = rInhalt
       AlbDat = messDatum
      End If
     End If
    End If
  Case 8000 ' Satzart (Turbomed) / Satzstatus (Internet-PDF-Dokument)
' 0020 Datenträger-Header, 0021 Datenträger-Abschluß, 0022 Datenpaket-Header, 0023 Datenpaket-Abschluß,
' 0010 Praxisdaten, 0101 Ärztliche Behandlung, 0102 Überweisungsfall, 0103 Belegeärztliche Behandlung,
' 0104 Notfalldienst/Vertretung/Notfall, 0190 Privatabrechnung, 0191 BG-Abrechnung, 0199 Unstrukturierte Fälle,
' 6100 Patientenstamm (Kennung > 3600), 6200 Behandlungsdaten (mit Tagesdatum 6200)
' BDT-Beschreibung ZI 1994: Abfolge der Satzarten < 0099 festgelegt, darüber beliebig
    f8000 = rInhalt
  Case 8100 ' Satzlänge (?)
    f8100 = rInhalt
  Case 8401 ' Labor: Fertigstellungsgrad
    FStG = rInhalt
  Case 8410 ' Labor: Abkürzung
    If InStrB(rInhalt, "<>") <> 0 Then Call aufSplit(rInhalt, "<>") Else Call aufSplit(rInhalt)
    If ArraInd > 1 Then
     Abkü = Arra(1)
     AbküLabor = Arra(2)
    Else
     Abkü = rInhalt 'replace$(RInhalt, "€", "_") ' ZS1 wäre hier nötig
     AbküLabor = vNS
    End If
    Call LaborEintr0
  Case 8411 ' Labor: Langtext
    If obLaborEintrag And rInhalt <> "" Then
     rLa(ls).Langtext = REPLACE$(rInhalt, "³", "ü") ' scheint die einzig angezeigte Ersetzung aus Zeichensatz() zu sein
     rLa(ls).LangtextVW = LTEinfüg&(rLa(ls).Langtext)
    End If
'    rLa(UBound(rLa)).Langtext = ZeichenSatz(RInhalt)
  Case 8420 ' Labor: Wert
    If obLaborEintrag Then rLa(ls).Wert = rInhalt
'    rLa(UBound(rLa)).Wert = RInhalt
'    IF rLa(UBound(rLa)).Abkü = "HBA1C" THEN
    If (Abkü = "HBA1" Or Abkü = "HBA1C" Or Abkü = "HbA1c Eigenlabor" Or Abkü = "HBA1C_" Or Abkü = "HBAC") And IIf(LenB(rInhalt) = 0, 0, rInhalt) < 22 Then ' Grenze für Unterscheidung %, mmol/l
     If messDatum > lHbA1cDat Then
      lHbA1c = rInhalt
      lHbA1cDat = messDatum
     End If
    End If
  Case 8421
   If obLaborEintrag Then rLa(ls).Einheit = REPLACE$(rInhalt, "³", "ü")  ' scheint die einzig angezeigte Ersetzung aus Zeichensatz() zu sein
'   rLa(UBound(rLa)).Einheit = RInhalt
  Case 8480
'   rLa(UBound(rLa)).Kommentar = rLa(UBound(rLa)).Kommentar + Stutz(RInhalt)
   If obLaborEintrag Then
    rLa(ls).Kommentar = rLa(ls).Kommentar + Stutz(rInhalt)
    rLa(ls).KommentarVW = KomEinfüg&(rLa(ls).Kommentar)
' jetzt kommen noch die nicht-numerischen Werte
    If rLa(ls).Wert = "" Then
     If InStrB(rInhalt, "<") > 0 Then
      rLa(ls).Wert = Mid$(rInhalt, InStr(rInhalt, "<"))
      If InStr(4, rLa(ls).Wert, " ") > 0 Then
       rLa(ls).Wert = Left$(rLa(ls).Wert, InStr(4, rLa(ls).Wert, " ") - 1)
      End If
     ElseIf InStrB(rInhalt, ">") > 0 Then
      rLa(ls).Wert = Mid$(rInhalt, InStr(rInhalt, ">"))
      Dim LeerPos% ' Korrektur 19.6.15
      LeerPos = InStr(4, rLa(ls).Wert, " ")
      If LeerPos > 0 Then
       rLa(ls).Wert = Left$(rLa(ls).Wert, LeerPos - 1)
      End If
     ElseIf InStrB(rInhalt, ":") > 0 Then ' 1:-Größen ohne < und >
      p1 = InStr(rInhalt, ":")
      For p2 = p1 To 1 Step -1
       If Mid$(rInhalt, p2, 1) = " " Then Exit For
      Next p2
      rLa(ls).Wert = Mid$(rInhalt, p2 + 1)
      If InStrB(rLa(ls).Wert, " ") > 0 Then
       rLa(ls).Wert = Left$(rLa(ls).Wert, InStr(rLa(ls).Wert, " ") - 1) ' Erklärungen weg
      End If
     End If
    End If ' rla(ls).wert=""
   End If
'    Call Laboreintr1
  Case 9601 ' Zeitraum der Speicherung
   SpeicherZt = BDTtoDate(Right$(rInhalt, 8))
  Case 9602 ' Beginn der Übertragung
   SpeicherZt = SpeicherZt + BDTtoTime(Left$(rInhalt, 6))
   Call myEFrag("UPDATE `eintragszahlen` SET speicherzt = " & DatFor_k(SpeicherZt) & " WHERE beginn = " & DatFor_k(rEzäBeg) & vNS, numA)
'  Case 4211 ' Ankreuzfeld HAH
'  Case 5010 ' Medikament
'  10.1.10: folgende 5 Zeilen in unbekannten Kennungen gerade entdeckt, noch nicht implementiert
  Case 8460 ' Normalwert-Text
  Case 8461 ' Normalwert untere Grenze
  Case 8462 ' Normalwert obere Grenze
'  Case 8470 ' Anmerkung
'  case 8490 ' 'Abschluss-Zeile', - #Test-Ident -- nur bei header 6200 -- -; bei Pat_id 11: "Auftragsnummer DOB350B"
' steht noch aus:
' case 3750 ' Desktop-Objekt
' case 4266 ' Kurabbruch
'  case 8409 ' Karteiartspalte für Blutdruck- und BMI-Formular
  Case 9100 ' Arztnummer des Absenders
  Case 9103 ' Erstellungsdatum
'     NAktZeit = BDTtoDate(RInhalt)
  Case 9105 ' Ordnungsnummer Datenträger (Header) des DP
  Case 9106 ' verwendeter Zeichensatz (7-bit oder 8-bit)
  Case 9210 ' Version ADT-Satzbeschreibung
  Case 9213 ' Version BDT-Satzbeschreibung
  Case 9600 ' Archivierungsart (BDT-Beschreibung ZI 1994: Zeitraum, aus dem Behandlungsdaten übermittelt werden
  
  Case 9202 ' Gesamtlänge des Datenpaketes
  Case 9203 ' Anzahl der Datenträger des Datenpaketes
  Case 4261 ' Rezeptart
  Case 4262 '
  Case 4263 ' Kurdauer
  Case 4264 ' Anreisetag
  Case 4265 ' Abreisetag
  Case 4271 ' Kurverlängerung
  
  Case 101 ' KBV -Prüfnummer            : A001011
  Case 102 ' Software -Verantwortlicher : Turbomed EDV Gmbh
  Case 103 ' Software                   : Turbomed@Windows
  Case 104 ' Hardware                   : IBM PC/AT
  Case 201 ' Arztnummer                 : 6419153
  Case 202 ' Praxistyp                  : 1 (Einzelpraxis)
  Case 203 ' Arztname                   : Gerald Schade
  Case 204 ' Arztgruppe                 : Innere Medizin
  Case 205 ' Straße der Praxis          : Münchner Straße 61b
  Case 207 ' Arzt mit Leistungskennzeichen <- KBV-Feldtabelle, nicht Turbomed
  Case 215 ' PLZ der Praxisadresse      : 85221
  Case 216 ' Ort der Praxisadresse      : Dachau
  Case 208 ' Telefonnummer der Praxis   : 08131 / 616 380
  Case 209 ' Telefaxnummer der Praxis   : 08131 / 616 381
' case 211 ' Arzt (Labor)
' case 212 ' LANR (Labor)
'  Case 101, 102, 103, 104 ' Installation
'  Case 201, 202, 203, 204, 205, 215, 216, 208, 209 ' KV-Nummer, meine Adresse
  Case Else ' Unbekannte Kennung
   ReDim Preserve rUn(UBound(rUn) + 1)
   rUn(UBound(rUn)).absPos = absPos
   rUn(UBound(rUn)).Kennung = RKennung
   rUn(UBound(rUn)).Pat_id = rNa(0).Pat_id
   rUn(UBound(rUn)).Inhalt = rInhalt
 End Select
 lKennung = RKennung
'  For i = UBound(lk) To 1 Step -1
'   lk(i) = lk(i - 1)
'   lk(0) = Rkennung
'  Next i
 End If ' Rkennung LIKE "####" THEN
 Exit Function
fehler_4247:
 If Err.Number = -2147467259 Or Err.Number = 3704 Then
  Call DBCnOpen
  Resume resume_4247
 End If
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away
 Call DBCnOpen
 Resume
ElseIf Err.Number = 3704 Then ' Der Vorgang ist für ein geschlossenes Objekt nicht zugelassen.
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dolies/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dolies

Function SchulungszifferZuord(rInhalt, ByRef DMSchulz%, ByRef RRSchulz%, Optional Zp As Date, Optional ByRef zpd1$, Optional ByRef zpdL$, Optional ByRef zpr1$, Optional ByRef zprL$)
    If zpd1 = "" Then zpd1 = "0"
    If zpdL = "" Then zpdL = "0"
    If zpr1 = "" Then zpr1 = "0"
    If zprL = "" Then zprL = "0"
    Select Case UCase$(rInhalt)
    Case "9261B", "9262B", "9263B", "9264B", "9268", "9270", "9271", _
         "97261B", "97262B", "97263B", "97264B", "97268", "97270", "97271", _
         "97269B", "97269", "97274", "97278", "97279", _
         "92292A", "92292B", "92290A", "92291A", _
         "92282", "92281", "92278", "92277", "97280", "92261A", "92262A", "92263A", "92269A", "92264A", "97273B", "92271A"
         
      DMSchulz = DMSchulz + 1
     If CDate(zpd1) > Zp Then zpd1 = Format$(Zp, "d.m.yy")
     If CDate(zpdL) < Zp Then zpdL = Format$(Zp, "d.m.yy")
     Case "9265B", "9266B", _
          "97265B", "97266B", _
          "97268B", _
          "92292C", "92292D", "92292E", _
          "92265A", "92266A", "92268A"
     RRSchulz = RRSchulz + 1
     If CDate(zpr1) > Zp Then zpr1 = Format$(Zp, "d.m.yy")
     If CDate(zprL) < Zp Then zprL = Format$(Zp, "d.m.yy")
    End Select
End Function ' SchulungszifferZuord(rInhalt, DMSchulz%, RRSchulz%)
' aufgerufen in dolies(): eine Medikamentenzeile hinzufügen
Function MPerg()
       ReDim Preserve rMe(UBound(rMe) + 1)
       MedZahl = MedZahl + 1
       rMe(UBound(rMe)).FeldNr = MedZahl
       rMe(UBound(rMe)).absPos = absPos
       rMe(UBound(rMe)).StByte = AktByte
       rMe(UBound(rMe)).AktZeit = AktZeit
       rMe(UBound(rMe)).Pat_id = rNa(0).Pat_id
       rMe(UBound(rMe)).FID = rFa(UBound(rFa)).FID
       rMe(UBound(rMe)).Zeitpunkt = messDatum
       rMe(UBound(rMe)).Datum = MPDatum
       rMe(UBound(rMe)).MPNr = pMpnr
       If obHergang <> 0 Then
        rMe(UBound(rMe)).Bemerkung = rMe(UBound(rMe) - 1).Bemerkung
        rMe(UBound(rMe) - 1).Bemerkung = ""
       End If
End Function ' MPerg

Function QAnf(Quartal) As Date ' Wird zZt nur für falleintragen gebraucht, nicht in Abfragen, da die zugehörige Abfrage zu langsam ist
Dim q As String * 1, DatStr$
If Quartal = "" Then
 QAnf = CDate(0)
Else
 q = Left$(Quartal, 1)
 Select Case q
  Case 1: DatStr = "1.1."
  Case 2: DatStr = "1.4."
  Case 3: DatStr = "1.7."
  Case 4: DatStr = "1.10."
 End Select
 QAnf = CDate(DatStr & Right$(Quartal, 4))
End If
End Function ' QAnf(quartal) As Date ' Wird zZt nicht gebraucht, da die zugehörige Abfrage zu langsam ist

Function QEnd(Quartal) As Date ' Wird zZt nur für falleintragen gebraucht, nicht in Abfragen, da die zugehörige Abfrage zu langsam ist
Dim q As String * 1, DatStr$
If Quartal = vNS Then
 QEnd = CDate(0)
Else
 q = Left$(Quartal, 1)
 Select Case q
  Case 1: DatStr = "31.3."
  Case 2: DatStr = "30.6."
  Case 3: DatStr = "30.9."
  Case 4: DatStr = "31.12."
 End Select
 QEnd = CDate(DatStr & Right$(Quartal, 4))
End If
End Function ' qend(quartal) As Date ' Wird zZt nicht gebraucht, da die zugehörige Abfrage zu langsam ist

Function QuartalStr$(Datum As Date, Optional obinv%)
 Dim q As String * 1
 Select Case Month(Datum)
  Case 1, 2, 3: q = 1
  Case 4, 5, 6: q = 2
  Case 7, 8, 9: q = 3
  Case 10, 11, 12: q = 4
 End Select
 If obinv Then
  QuartalStr = Year(Datum) & q
 Else
  QuartalStr = q & Year(Datum)
 End If
End Function ' Quartal

Function LeistEintr0(lG)
 On Error GoTo fehler
 If Not IsNull(lG) Then
  ReDim Preserve rLe(UBound(rLe) + 1)
  rLe(UBound(rLe)).Pat_id = rNa(0).Pat_id
  rLe(UBound(rLe)).Zeitpunkt = messDatum
  rLe(UBound(rLe)).Leistung = lG
  rLe(UBound(rLe)).QS = ZQSort(messDatum)
  rLe(UBound(rLe)).QT = ZQuart(messDatum)
  rLe(UBound(rLe)).absPos = absPos
  rLe(UBound(rLe)).AktZeit = AktZeit
  rLe(UBound(rLe)).FID = rFa(UBound(rFa)).FID
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistEintr0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistEintr(Lg)

Function test7()
 Dim rs As ADODB.Recordset, DT$
 Set rs = myEFrag("SELECT Inhalt FROM `eintraege` WHERE pat_id = 1786 AND art=""andm""")
 Debug.Print rs!Inhalt
 DT = Mid$(rs!Inhalt, InStr(rs!Inhalt, "Diabetes Typ") + 13, InStr(rs!Inhalt, "seit") - InStr(rs!Inhalt, "Diabetes Typ") - 13)
 Debug.Print DT
End Function ' test7

Function fDTyp$()
' DTypeintrag
 Dim DT$, i& ', erg$ ', rAn As New ADODB.Recordset
 On Error GoTo fehler
 For i = 1 To UBound(rEi)
  If rEi(i).art = "andm" Or rEi(i).art = "andm2" Then
   Dim pos&
   pos = InStr(rEi(i).Inhalt, "Diabetes Typ")
   If pos > 0 Then
   On Error Resume Next
   DT = Mid$(rEi(i).Inhalt, InStr(rEi(i).Inhalt, "Diabetes Typ") + 13, InStr(rEi(i).Inhalt, "seit") - InStr(rEi(i).Inhalt, "Diabetes Typ") - 13)
   On Error GoTo fehler
   If InStrB(DT, "{") > 0 Then
    DT = Left$(DT, InStr(DT, "{") - 1)
   End If
   DT = Trim$(DT)
'   IF Not rAn Is Nothing THEN IF rAn.State = 1 THEN rAn.Close
'   rAn.Open "SELECT * FROM `anamnesebogen` WHERE pat_id = " & rEi(i).Pat_id, DBCn, adOpenDynamic, adLockOptimistic
'   IF Not rAn.BOF THEN
'   IF ISNULL(rAn!Diabetestyp) OR rAn!Diabetestyp = "" THEN
'   IF rAna(0).Diabetestyp = "" THEN
     If InStrB(DT, "estat") > 0 Then fDTyp = "g"
     If InStrB(DT, "ekund") > 0 Or InStrB(DT, "pankreopri") > 0 Then fDTyp = "s"
     If InStrB(DT, "path") > 0 Then fDTyp = "p"
     If InStrB("12", Left$(DT, 1)) > 0 And Left$(DT, 1) <> "" Then fDTyp = DT
     If InStrB(DT, "?") > 0 Then fDTyp = fDTyp + "?"
'     erg = LEFT(erg, rAn.Fields("Diabetestyp").DefinedSize)
'     Call myEFrag("UPDATE `anamnesebogen` SET diabetestyp = '" & erg & "' WHERE pat_id = " & rEi(i).Pat_id, rAf)
'     rAna(0).Diabetestyp = erg
'   END IF
'   END IF
   End If
  End If
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fDTyp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fDtyp

Function fDMSchL&()
 Dim i&
 fDMSchL = 0
 For i = 1 To UBound(rEi)
  If rEi(i).Inhalt Like "*. sitzung*" Then
   fDMSchL = fDMSchL + 1
  End If
 Next i
End Function ' fDMSchl

Function fFanfFuell()
'FAnfFuell
 On Error GoTo fehler
 Dim lpid&, aktpid&, Bm$, lAusgst As Date, i&, j&
 lAusgst = 0
 For i = 1 To UBound(rFa)
  rFa(i).Fanf = 0
  If rFa(i).SchGr = "90" Or rFa(i).SchGr = "41" Then
   rFa(i).Fanf = rFa(i).BhFB
  ElseIf rFa(i).ausgst <> 0 Then
   rFa(i).Fanf = rFa(i).ausgst
  End If
  If rFa(i).Fanf = 0 Then
   If rFa(i).lVorl <> 0 And DateValue(rFa(i).lVorl) <= lAusgst Then
    rFa(i).Fanf = rFa(i).lVorl
   End If
  End If
  If rFa(i).Fanf = 0 Then
   Dim Eintr1 As Date
   Eintr1 = 0
' sql = "SELECT * FROM `eintraege` WHERE pat_id = " + CStr(rFa(i).Pat_ID) + " AND zeitpunkt >= DATE('" + format$(rFa(i).BhFB, "DD.MM.YY") + "') "
   For j = 1 To UBound(rEi)
    If rEi(j).Zeitpunkt > rFa(i).BhFB And (rFa(i).BhFE1 = 0 Or rEi(j).Zeitpunkt <= rFa(i).BhFE1) Then
     If rEi(j).Zeitpunkt < Eintr1 Or Eintr1 = 0 Then
      Eintr1 = rEi(j).Zeitpunkt
     End If
    End If
   Next j
   If Eintr1 <> 0 Then
    If Eintr1 > rFa(i).lVorl And rFa(i).lVorl <> 0 Then
     rFa(i).Fanf = rFa(i).lVorl
    Else
     rFa(i).Fanf = Eintr1
    End If
   Else
    If rFa(i).lVorl <> 0 Then
     rFa(i).Fanf = rFa(i).lVorl
    ElseIf rFa(i).Pat_id = 2 And rFa(i).Nachname Like "Muster*" Then
     rFa(i).Fanf = #7/1/2004#
    Else
     If obÜProt Then Print #322, rFa(i).Pat_id & ": Hatnäckiger Fall bei der Fallanfangsfestlegung"
     rFa(i).Fanf = rFa(i).BhFB
    End If
   End If
  End If
  lAusgst = rFa(i).ausgst
 Next i
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fAnfFuell/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fFAnfFuell

Function trs()
 Lese.ProgStart
 rsAnamOpen
End Function

Function rsAnamOpen()
 Dim i%, sql$
 On Error Resume Next
 If Not rsAnm Is Nothing Then If rsAnm.State = 1 Then If rsAnm.EditMode <> 0 Then rsAnm.CancelUpdate: rsAnm.Close
 On Error GoTo fehler
 If Not rsAnm Is Nothing Then If rsAnm.State = 1 Then rsAnm.Close
 rsAnm.CursorLocation = adUseServer
 For i = 1 To 2
nochmal:
  If Not rsAnm Is Nothing Then If rsAnm.State = 1 Then rsAnm.Close
'  Call rsAnm.Open("SELECT -obmednetz AS j_obmednetz, -tkz AS j_tkz, a.* FROM `anamnesebogen` a WHERE pat_id = " & rNa(0).Pat_id, DBCn, adOpenDynamic, adLockOptimistic)
'  Call rsAnm.Open("SELECT COALESCE(`diabetes seit`,'') `diabetes seit`, a.* FROM `anamnesebogen` a WHERE pat_id = " & rNa(0).Pat_id, DBCn, adOpenDynamic, adLockOptimistic)
  myFrag rsAnm, "SELECT COALESCE(`diabetes seit`,'') `diabetes seit`, a.* FROM `anamnesebogen` a WHERE pat_id = " & rNa(0).Pat_id, adOpenDynamic, DBCn, adLockOptimistic
  If rsAnm.BOF Then
   Dim primnr&
   If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
   myFrag rsAdo, "SELECT MAX(prim) mprim FROM `anamnesebogen`"
'   SET rsAdo = myEFrag("SELECT MAX(prim) AS mprim FROM `anamnesebogen`")
   If IsNull(rsAdo!mprim) Then primnr = 1 Else primnr = rsAdo!mprim + 1
   sql = "INSERT INTO `anamnesebogen`(pat_id,prim) VALUES(" & rNa(0).Pat_id & "," & primnr & ")"
   InsKorr DBCn, DBCnS, sql, rAF
   If obFor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
  Else
   Exit For
  End If
 Next i
 Exit Function
fehler:
 If Err.Number = 2006 Or Err.Number = 3704 Then ' Server has gone away o.ä.
  Set rsAnm = Nothing
  GoTo nochmal
 ElseIf Err.Number = -2147467259 Then  ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
  ErrDescription = Err.Description
  Call doBezFeh(sql, lies.dlg.BeziehungsfehlerSpeichern, ErrDescription)
  Resume
 End If
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rsAnamOpen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rsAnamOpen

' ehemals in alleSpeichern, jetzt nirgends mehr
Function kvnrpruef()
 Dim rs As New ADODB.Recordset
 On Error GoTo fehler
 myFrag rs, "SELECT 0 FROM `hausaerzte` WHERE kvnr = '" & rNa(0).KVNr & "'"
 If rs.BOF Then
 ' nicht: schwerpunkt,zusatzbezeichnung,bemerkung,beme,sprechstunden,von_bis,internetadressen,`Behandlung in Fremdsprachen`,`Rollstuhlgerechte Praxis`,verkehrsmittel,linie,`Haltestelle Parkplätze`,Wegbeschreibung,`Entfernung zur Praxis`,zahl,nichtmehr,gelöscht
  Set rs = Nothing
  myFrag rs, "SELECT 0 FROM `aktlue` WHERE kvnro = '" & rNa(0).KVNr & "'"
  If rs.BOF Then
   myEFrag "INSERT INTO `hausaerzte`(kvnr) VALUES('" & rNa(0).KVNr & "')", rAF
  Else
   On Error Resume Next
   myEFrag "INSERT INTO `hausaerzte`(überschrift,name,vorname,nachname,anschrift,kvnr,telefon,telefax,e_mail,zulassungsgebiet,arzttyp,`Gemeinschaftspraxis mit`,titel,geschlecht,`straße`,plz,ort,dmpt2,dmpt1) SELECT überschrift, CONCAT_WS(' ',anrede,titelt,vorname,name),vorname,name,CONCAT_WS(', ',strasse,CONCAT_WS(' ',plz,ort)),kvnr,telefon,fax,email,fachgruppe,arzttyp,gemmit,titel,geschlecht,strasse,plz,ort,dmpt2,dmpt1  FROM `aktlue` WHERE kvnro = '" & rNa(0).KVNr & "'", rAF
   On Error GoTo fehler
  End If
  Debug.Print rAF
 End If
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrpruef/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' kvnrpruef
Function FIDsetz()
 Dim i&
 For i = 1 To UBound(rEi)
  If rEi(i).FID = 0 Then
   If UBound(rFa) > 0 Then
    rEi(i).FID = rFa(1).FID
   End If
  Else
   Exit For
  End If
 Next i
 For i = 1 To UBound(rFr)
  If rFr(i).FID = 0 Then
   If UBound(rFa) > 0 Then
    rFr(i).FID = rFa(1).FID
   End If
  Else
   Exit For
  End If
 Next i
 For i = 1 To UBound(rLe)
  If rLe(i).FID = 0 Then
   If UBound(rFa) > 0 Then
    rLe(i).FID = rFa(1).FID
   End If
  Else
   Exit For
  End If
 Next i
 For i = 1 To UBound(rRe)
  If rRe(i).FID = 0 Then
   If UBound(rFa) > 0 Then
    rRe(i).FID = rFa(1).FID
   End If
  Else
   Exit For
  End If
 Next i
End Function ' FIDsetz

' in alleSpeichern
Function MedArtenPruef()
 Dim rs As New ADODB.Recordset, i&
 For i = 1 To UBound(rMe)
  myFrag rs, "SELECT 0 FROM `medarten` WHERE medikament = '" & GetMed(rMe(i).Medikament, 0) & "'"
  If rs.BOF Then
   InsKorr DBCn, DBCnS, "INSERT INTO `medarten`(langname,medikament,hinzugefügt,pat_id) VALUES('" & rMe(i).Medikament & "','" & UCase$(GetMed(rMe(i).Medikament, 0)) & "'," & DatFor_k(Now) & "," & rMe(i).Pat_id & ")", rAF
  Else
   myFrag rs, "UPDATE medarten SET pat_id= " & rMe(i).Pat_id & " WHERE medikament='" & GetMed(rMe(i).Medikament, 0) & "' AND pat_id=0"
  End If
  Set rs = Nothing
 Next i
End Function ' MedArtenPruef

Function aktqanf(Optional diff%) As Date
 Dim jetzt#, mon$
 jetzt = Now() - diff
 Select Case Month(jetzt)
  Case Is > 9: mon = "10"
  Case Is > 6: mon = "7"
  Case Is > 3: mon = "4"
  Case Else: mon = "1"
 End Select
 aktqanf = CDate("01." & mon & "." & Year(jetzt))
End Function ' aktqanf() As Date

' Aufruf nur in GesLies(
Function alleSpeichern(frm As Lese)
' rsAnam!Vorgestellt = MYDAT(Vorgestellt)
 Dim Cpt$, i&, j&
 Dim DMSchL&
 Dim rsc As ADODB.Recordset
 On Error GoTo fehler
 frm.SBez.BackColor = &HFF&
 DoEvents
 Call fFanfFuell
 
 On Error Resume Next
 Call DBCn.BeginTrans: obTrans = 1
 On Error GoTo fehler
 
' Medklass, redesigned, erster Teil
' Call ForeignNo
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
#If ohnebezug Then ' auskommentiert 22.10.22, da es den Patienten 0 gibt
 Call ForeignNo0
 Call ForeignNo1
#End If
 myEFrag "UPDATE `medarten` SET pat_id = 0 WHERE pat_id = " & rNa(0).Pat_id, rAF
#If ohnebezug Then ' auskommentiert 22.10.22
 Call ForeignYes0
 Call ForeignYes1
#End If
 Dim medi As MediCl
 For i = 1 To UBound(rMe)
  Set medi = New MediCl
  medi.Med = GetMed(rMe(i).Medikament, 0)
  medi.LT = rMe(i).Medikament
  medi.Pat_id = rNa(0).Pat_id
  Call medSL.sCAdd(medi, True)
 Next i
 For i = 1 To UBound(rRe)
  Set medi = New MediCl
  medi.Med = GetMed(rRe(i).Medikament, 3)
  medi.LT = rRe(i).Medikament
  medi.Pat_id = rNa(0).Pat_id
  Call medSL.sCAdd(medi, True)
 Next i

' Call myEFrag("SET autocommit = 0", , adAsyncExecute)
' Call myEFrag("start transaction") '("begin")
#If False Then
 If ProgrammLauf(-1, Cpt, -1) Then GoTo Ende
#End If
anamneseanfang:
 On Error Resume Next
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
 On Error GoTo fehler
nachFehler:
 Call rsAnamOpen
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "anamnesebogen", Empty))
 Dim Inhalt$, Länge%
 Do While Not rsc.EOF
  Select Case rsc!COLUMN_NAME
   Case "Nachname": Inhalt = rNa(0).Nachname: Länge = Len(rNa(0).Nachname)
   Case "Vorname":  Inhalt = rNa(0).Vorname:  Länge = Len(rNa(0).Vorname)
   Case "NVorsatz": Inhalt = rNa(0).NVorsatz: Länge = Len(rNa(0).NVorsatz)
   Case "Titel":    Inhalt = rNa(0).Titel:    Länge = Len(rNa(0).Titel)
  End Select
  Select Case rsc!COLUMN_NAME
   Case "Nachname", "Vorname", "NVorsatz", "Titel"
    Dim Cn$
    Cn = rsc!COLUMN_NAME
    If rsAnm.Fields(Cn).DefinedSize < Länge Then
     Call SpMod(Länge, "anamnesebogen", rsc, Inhalt)
     GoTo anamneseanfang
    End If
  End Select
  rsc.MoveNext
 Loop
 On Error Resume Next
 DBCn.BeginTrans: obTrans = 1
 On Error GoTo fehler
 If rsAnm!Nachname <> rNa(0).Nachname Or IsNull(rsAnm!Nachname) Then rsAnm!Nachname = rNa(0).Nachname
 If rsAnm!Vorname <> rNa(0).Vorname Or IsNull(rsAnm!Vorname) Then rsAnm!Vorname = rNa(0).Vorname
 If rsAnm!GebDat <> rNa(0).GebDat Or IsNull(rsAnm!GebDat) Then rsAnm!GebDat = rNa(0).GebDat
 If rsAnm!NVorsatz <> rNa(0).NVorsatz Or IsNull(rsAnm!NVorsatz) Then rsAnm!NVorsatz = rNa(0).NVorsatz
 If rsAnm!Titel <> rNa(0).Titel Or IsNull(rsAnm!Titel) Then rsAnm!Titel = rNa(0).Titel
 If rsAnm!Vorgestellt <> VorStDat Or IsNull(rsAnm!Vorgestellt) And VorStDat <> 0 Then rsAnm!Vorgestellt = VorStDat

 
' 1. Angabe bei der Anamnese
' IF rsAnm!Diabetestyp = vNS OR IsNull(rsAnm!Diabetestyp) THEN ' Kommentar 12.2.22
  rsAnm!Diabetestyp = Left$(fDTyp, rsAnm!Diabetestyp.DefinedSize)
' END IF
' 2. Angabe bei den Diagnosen
 Dim aktDCyp$
 aktDCyp = getDTyp
 If IsNull(rsAnm!Diabetestyp) Or ((aktDCyp = "1" Or aktDCyp = "2" Or aktDCyp = "s" Or aktDCyp = "g" Or aktDCyp = "p" Or aktDCyp = "" Or aktDCyp = "?") And aktDCyp <> rsAnm!Diabetestyp) Then
   rsAnm!Diabetestyp = aktDCyp
 End If
' Hier sollte noch die Therapieart mit TheArtSp festgelegt werden
 If rsAnm!DMSchulz <> DMSchulz Or IsNull(rsAnm!DMSchulz) Then rsAnm!DMSchulz = DMSchulz
 If rsAnm!RRSchulz <> RRSchulz Or IsNull(rsAnm!RRSchulz) Then rsAnm!RRSchulz = RRSchulz
 DMSchL = fDMSchL()
 If rsAnm!DMSchL <> DMSchL Or IsNull(rsAnm!DMSchL) Then rsAnm!DMSchL = DMSchL
 If rsAnm!DMPhier <> dmpdat Or IsNull(rsAnm!DMPhier) Then rsAnm!DMPhier = dmpdat
 Dim anrede$
 anrede = IIf(rNa(0).geschlecht = "m", "Herr", "Frau")
 If rsAnm!anrede <> anrede Or IsNull(rsAnm!anrede) Then rsAnm!anrede = anrede
' IF (rsAnm!j_obMedNetz = 0 OR ISNULL(rsAnm!j_obMedNetz)) AND obMedNetz THEN rsAnm!obMedNetz = obMedNetz
 If (rsAnm!obMedNetz = 0 Or IsNull(rsAnm!obMedNetz)) And obMedNetz Then rsAnm!obMedNetz = Abs(obMedNetz)
' IF Tkz AND (Not rsAnm!j_tkz OR ISNULL(rsAnm!j_tkz)) THEN rsAnm!j_tkz = True ELSE IF ISNULL(rsAnm!j_tkz) THEN rsAnm!j_tkz = 0
 If Tkz And (Not rsAnm!Tkz Or IsNull(rsAnm!Tkz)) Then rsAnm!Tkz = 1 Else If IsNull(rsAnm!Tkz) Then rsAnm!Tkz = 0
' Versicherungsart / Scheingruppe
 If rsAnm!Versicherungsart <> rFa(UBound(rFa)).SchGr Or IsNull(rsAnm!Versicherungsart) Then rsAnm!Versicherungsart = rFa(UBound(rFa)).SchGr
 Dim DiagStr$
 DiagStr = MacheDiagnosen(IIf(IsNull(rsAnm![diabetes seit]), "", rsAnm![diabetes seit]))
 If rsAnm!Diagnosen <> DiagStr Then
  On Error Resume Next
  For i = 1 To 2
   rsAnm!Diagnosen = DiagStr
   If Err.Number > 0 Then
    DiagStr = Left$(DiagStr, Len(DiagStr) - 1)
   Else
    Exit For
   End If
  Next
  On Error GoTo fehler
 End If
 On Error GoTo fehler
 'If rsAnm.EditMode = 1 THEN rsAnm.Update
 Call dmpErg
 Call usdm0
 Call usdm1
 Call usdm2
 Call USfuss
 Call USUlcus
 Call usVKGD
 Call usVKGD2
 Call AnAlle
 Call AnTrennZeichen
 
 Select Case rsAnm.EditMode
  Case adEditAdd, adEditInProgress
   On Error Resume Next
   rsAnm.Update
   If Err.Number <> 0 Then
'    SET rsAnm = Nothing
    If Not lies.obMySQL Then rsAnm.CancelUpdate
'   GoTo anamneseanfang
    Call anaUpd(rsAnm, rsAnm)
   End If
'   GoTo anamneseanfang
   On Error GoTo fehler
 End Select
 
' eintragen
 frm.SBez = "."
 DoEvents
 
 For i = 1 To UBound(rLa)
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  Dim sL1 As sLpar, sL2 As sLpar
  Set sL1 = New sLpar
  sL1.Abkü = rLa(i).Abkü
'  sL1.Labor = rLa(i).Labor
  sL1.Einheit = rLa(i).Einheit
  sL1.Langtext = rLa(i).Langtext
  Set sL2 = sListLpar.GetItem(sL1)
  If sL2 Is Nothing Then
   InsKorr DBCn, DBCnS, "INSERT INTO `laborparameter` (`abkü`,`einheit`,`langtext`,`aktzeit`) VALUES('" & rLa(i).Abkü & "','" & rLa(i).Einheit & "','" & rLa(i).Langtext & "'," & DatFor_k(Now()) & ")", rAF
'   SET rsAdo = myEFrag("SELECT abkü,eiheit,langtext FROM `laborparameter` WHERE abkü = '" & rLa(i).Abkü & "' AND einheit = '" & rLa(i).Einheit & "')")
   Call sListLpar.sCAdd(sL1)
  Else
  End If
 Next i
 
 For i = 1 To UBound(rFm)
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  If obVorb Then
   Dim fi1 As sFeldInh, fi2 As sFeldInh
   Set fi1 = New sFeldInh
   fi1.FeldInh = rFm(i).FeldInh
   Set fi2 = sListFldInh.GetItem(fi1)
   If fi2 Is Nothing Then
    myEFrag "INSERT INTO `forminhaltfeldinh` (feldinh,stbyte) VALUES('" & rFm(i).FeldInh & "'," & AktByte & ")", rAF
'nochmal:
    Set rsAdo = myEFrag("SELECT feldinhvw,feldinh FROM `forminhaltfeldinh` WHERE feldinh = '" & rFm(i).FeldInh & "'")
    If rsAdo.EOF Then
     Set rsAdo = myEFrag("SELECT feldinhvw,feldinh FROM `forminhaltfeldinh` ORDER BY feldinhvw DESC")
    End If
    fi1.FeldInhVW = rsAdo!FeldInhVW
    Call sListFldInh.sCAdd(fi1)
    rFm(i).FeldInhVW = fi1.FeldInhVW
   Else
    rFm(i).FeldInhVW = fi2.FeldInhVW
   End If
  Else
   If lies.obMySQL Then
    For j = 1 To 2
     Set rsAdo = myEFrag("SELECT feldinhvw,feldinh FROM `forminhaltfeldinh` WHERE feldinh = '" & rFm(i).FeldInh & "'")
     If rsAdo.BOF Then
      Set rsAdo = myEFrag("INSERT INTO `forminhaltfeldinh` (feldinh,stbyte) VALUES('" & rFm(i).FeldInh & "'," & AktByte & ")", rAF)
     Else
      Exit For
     End If
    Next j
   Else
    Dim sqlakt$
    sqlakt = "SELECT feldinhvw,feldinh,stbyte FROM `forminhaltfeldinh` WHERE feldinh = '" & rFm(i).FeldInh & "'"
    If rsAdo.State = 1 Then Set rsAdo = New ADODB.Recordset
'    rsAdo.Open sqlakt, DBCn, adOpenDynamic, adLockOptimistic
    myFrag rsAdo, sqlakt
    If rsAdo.BOF Then
     rsAdo.AddNew
     rsAdo!FeldInh = rFm(i).FeldInh
     rsAdo!StByte = AktByte
     rsAdo.Update
    End If
   End If
   rFm(i).FeldInhVW = rsAdo!FeldInhVW
  End If
    
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  If obVorb Or 1 = 1 Then
   Dim f1 As sFeld, f2 As sFeld
   Set f1 = New sFeld
   f1.Feld = rFm(i).Feld
   Set f2 = sListFeld.GetItem(f1)
   If f2 Is Nothing Then
    Set rsAdo = myEFrag("INSERT INTO `forminhaltfeld` (Feld,stbyte) VALUES('" & rFm(i).Feld & "'," & AktByte & ")")
    Set rsAdo = myEFrag("SELECT Feldvw,Feld FROM `forminhaltfeld` WHERE Feld = '" & rFm(i).Feld & "'")
    f1.FeldVW = rsAdo!FeldVW
    Call sListFeld.sCAdd(f1)
    rFm(i).FeldVW = f1.FeldVW
   Else
    rFm(i).FeldVW = f2.FeldVW
   End If
  Else
   For j = 1 To 2
    Set rsAdo = myEFrag("SELECT feldvw,feld FROM `forminhaltfeld` WHERE feld = '" & rFm(i).Feld & "'")
    If rsAdo.BOF Then
     Set rsAdo = myEFrag("INSERT INTO `forminhaltfeld` (feld,stbyte) VALUES('" & rFm(i).Feld & "'," & AktByte & ")")
    Else
     Exit For
    End If
   Next j
   rFm(i).FeldVW = rsAdo!FeldVW
  End If
 Next i
 
 frm.SBez = "S"
 DoEvents
 'rFm.FormID muss nach dem Speichern noch auf rFr.Form_id übertragen werden
 'rFo.FoID muss noch auf rFm.FoID übertragen werden
 ' rFm formulare -> rfmo
 ' rfo forminhkopf -> rfoo
 ' rfr forminhfeld -> rFm
'Public rFo() AS formulare
'Public rFr() AS forminhkopf
'Public rFm() AS forminhfeld

' Anamnesebogenfelder einstellen (-> in Access: FUNCTION FehlendeAnamneseBögen())
 For i = 1 To UBound(rDo)
  If InStrB(rDo(i).DokName, "namnese") > 0 Then
   If InStrB(rDo(i).DokName, "allg") > 0 Then
    rNa(0).AnAllgda = 1
   ElseIf InStrB(rDo(i).DokName, "2") > 0 Then
    rNa(0).An2da = 1
   ElseIf InStrB(rDo(i).DokName, "1") > 0 Then
    rNa(0).An1da = 1
   End If
  ElseIf InStrB(rDo(i).DokName, "heckliste") > 0 Then
   rNa(0).Checkda = 1
  End If
 Next i
' obhierdmp
' rNa(0).dmpklass = obhierdmpfn(rNa(0).Notiz, , , , rNa(0).dmpbeg, rNa(0).dmpkhkklass, rNa(0).dmpkhkbeg, rNa(0).dmpcopdklass, rNa(0).dmpcopdbeg, rNa(0).dmpcopdklass)
 obhierdmpfn rNa(0).Notiz, rNa(0).dmpklass, rNa(0).dmpbeg, rNa(0).dmpkhkklass, rNa(0).dmpkhkbeg, rNa(0).dmpcopdklass, rNa(0).dmpcopdbeg, rNa(0).dmpabklass, rNa(0).dmpabbeg, rNa(0).HzV, rNa(0).HzVbeg, rNa(0).DS, rNa(0).DSbeg
 Dim Infos12$()
'  Dim rKv1() AS kvnrue
 Call getHausarzt1(Infos12, rFa, rKv, True)
 For i = 0 To UBound(Infos12, 2)
  If InStrB(Infos12(12, i), " ") Then Infos12(12, i) = REPLACE$(Infos12(12, i), " ", "") ' 19.12.14, '64 16653'
 Next i
 rNa(0).getHA0 = IIf(Infos12(12, 0) = vNS, 0, Infos12(12, 0))
 rNa(0).fnHA0 = IIf(Infos12(10, 0) = vNS, 0, Infos12(10, 0))
 If UBound(Infos12, 2) > 0 Then
  rNa(0).getHA1 = IIf(Infos12(12, 1) = vNS, 0, Infos12(12, 1))
  rNa(0).fnHA1 = IIf(Infos12(10, 1) = vNS, 0, Infos12(10, 1))
  If UBound(Infos12, 2) > 1 Then
   rNa(0).getHA2 = IIf(Infos12(12, 2) = vNS, 0, Infos12(12, 2))
   rNa(0).fnHA2 = IIf(Infos12(10, 2) = vNS, 0, Infos12(10, 2))
  End If
 End If ' UBound(Infos12, 2) > 0
 For i = 0 To UBound(Infos12, 2)
'  IF i < 3 THEN
'   myEFrag "UPDATE `namen` SET getha" & CStr(i) & " = " & IIf(Infos12(12, i) = vNS, 0, Infos12(12, i)) & " WHERE pat_id = " & rNa(0).Pat_id, rAF
'  END IF
  If Infos12(12, i) <> vNS Then
'   myEFrag "SELECT kvnr FROM `hareal` WHERE kvnr = " & Infos12(12, i), rAF
   Dim rKVNr As New ADODB.Recordset, obFehlt%
   myFrag rKVNr, "SELECT 0 FROM `hareal` WHERE kvnr = " & Infos12(12, i)
   obFehlt = rKVNr.EOF
   Set rKVNr = Nothing
   If obFehlt Then
    InsKorr DBCn, DBCnS, "INSERT INTO `hareal`(Anrede,Adressat,Straße,PLZOrt,Fax,Überschrift,dmp2,dmp1,Niederlassungsgebiet,Vorname,InnereAllg,kvnr,Tel,Nachname) VALUES(" & IIf(Infos12(0, i) = "Herr", 1, 0) & ",'" & Infos12(1, i) & "','" & Infos12(2, i) & "','" & Infos12(3, i) & "','" & Infos12(4, i) & "','" & Infos12(5, i) & "'," & IIf(Infos12(6, i) = vNS, 0, 1) & "," & IIf(Infos12(7, i) = vNS, 0, 1) & ",'" & Infos12(8, i) & "','" & Infos12(9, i) & "'," & IIf(Infos12(11, i) = "", 0, 1) & "," & IIf(Infos12(12, i) = vNS, 0, Infos12(12, i)) & ",'" & Infos12(13, i) & "','" & Infos12(14, i) & "')", rAF
   Else
    myEFrag "UPDATE `hareal` SET Anrede=" & IIf(Infos12(0, i) = "Herr", 1, 0) & ",Adressat='" & Infos12(1, i) & "',Straße='" & Infos12(2, i) & "',PLZOrt='" & Infos12(3, i) & "',Fax='" & Infos12(4, i) & "',Überschrift='" & Infos12(5, i) & "',dmp2=" & IIf(Infos12(6, i) = vNS, 0, 1) & ",dmp1=" & IIf(Infos12(7, i) = vNS, 0, 1) & ",Niederlassungsgebiet='" & Infos12(8, i) & "',Vorname='" & Infos12(9, i) & "',InnereAllg=" & IIf(Infos12(11, i) = vNS, 0, 1) & ",Tel='" & Infos12(13, i) & "',Nachname='" & Infos12(14, i) & "' WHERE kvnr = " & Infos12(12, i), rAF
   End If
  End If
 Next i
 
 Call kassenspeichern(frm, CStr(rNa(0).Pat_id))
' Call kvnrpruef
 Call FIDsetz
 Call MedArtenPruef
#If altMed Then
' Folgendes 21.1.12
' "Medikamentenplan"_4
 For i = UBound(rMe) To 2 Step -1
  If rMe(i).Medikament = vNS And rMe(i).ab = vNS And rMe(i).mi = vNS And rMe(i).mo = vNS And rMe(i).nm = vNS And rMe(i).Zn = vNS And rMe(i).Bemerkung = vNS Then
   ReDim Preserve rMe(i - 1)
  Else
   Exit For
  End If
 Next i
#End If
 Call rrParsen
 If obHausBesuch Then
  tuLaden
  qbeg = aktqanf()
  ' 13.6.22: Nach Hausbesuchsmodulimport war die DAK-Einverständniserklärung nicht mehr erkennbar:
  If roNa(1).dakab <> 0 And rNa(0).dakab = 0 Then
   rNa(0).dakab = roNa(1).dakab
  End If
  Call faelleEinf
  Call auEinf
  Call briefeEinf
  Call diagnosenEinf
  Call dokumenteEinf
  Call eintraegeEinf
  Call forminhkopfEinf
  Call kheinweisEinf
  Call lbanforderungenEinf
  Call laborneuEinf
  Call leistungenEinf
  Call medplanEinf
  Call rezepteintraegeEinf
  Call rrEinf
  Call dmpreiheEinf
  Call desktopEinf
  Call usdmEinf
  Call fussEinf
  Call ulcusEinf
  Call vkgdEinf
  ' SELECT i.pat_id,gfid,f.* FROM (SELECT COUNT(0) zl, group_concat(fid) gfid,f.* FROM faelle f GROUP BY pat_id, bhfe1,fanf,lvorl,vknr,tmfnr,STATUS,privvers) i LEFT JOIN faelle f ON f.pat_id=i.pat_id AND INSTR(gfid,f.fid)<>0 WHERE zl>1; => bringt ein Minimum an Datensätzen
 End If ' obHausBesuch
 Call tuSpeichern(frm, frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern)
 ' korrigiertes Aufnahmedatum
 myEFrag "UPDATE namen n LEFT JOIN (SELECT pat_id, MIN(bhfb) bhfb, MIN(fanf) fanf FROM faelle f GROUP BY pat_id) f ON n.pat_id=f.pat_id SET kAufDat=date(IF(fanf>bhfb,fanf,bhfb)) WHERE f.pat_id=" & CStr(rNa(0).Pat_id), rAF
 If rsAnm Is Nothing Or rsAnm.State = 0 Then Call rsAnamOpen
 Call rrParseSpeichern
' Call TherapieArtEinzelnFestlegen(rNa(0).Pat_id, rsAnm)
#If thaalt Then ' GegenStück nach anamnesebogen, sonst Sperre
 If UBound(rTh) > 0 Then ' 11.7.10
  rsAnm!TherAkt = rTh(UBound(rTh)).therart
 End If
#End If
' Call ForeignYes
' Ende:
' Call myEFrag("commit")
 On Error Resume Next
 If obTrans <> 0 Then Call DBCn.CommitTrans: obTrans = 0
#If Not thaalt Then
'Shell "ssh root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & CStr(rNa(0).Pat_id) & ")'", vbMaximizedFocus
'rufauf "c:\windows\system32\openssh\ssh.exe", "root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & CStr(rNa(0).Pat_id) & ")'", 1, "c:\windows\system32\openssh\", -1
'Debug.Print rNa(0).Pat_id
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
Dim fengroe%
#If mitfenster Then
rufauf "ssh", "root@linux1 mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & CStr(rNa(0).Pat_id) & ")'", 2, "c:\windows\system32\openssh\", -1, 0
#Else
 Call TheraErmitt(rNa(0).Pat_id)
#End If
' myEFrag "CALL fuellThaP(" & CStr(rNa(0).Pat_id) & ")"
#End If
 If Kassengeändert Then
  myEFrag "UPDATE `kassenliste` k LEFT JOIN " & vbCrLf & _
  "(SELECT kateg,go,ik,vknr FROM `kassenliste`" & vbCrLf & _
  " WHERE kateg <> '' GROUP BY ik,vknr) k2 ON k.vknr = k2.vknr AND k.ik=k2.ik " & vbCrLf & _
  "SET k.kateg=k2.kateg, k.go=k2.go, geaen=" & Format(Now(), "yyyymmddHHMMSS") & vbCrLf & _
  " WHERE k.kateg = '' AND k2.kateg IS NOT NULL", rAF
  Ausgeb rAF & " Kassenkategorien anhand der VK-Nummern eingefügt"
'  Call Lese.KassenkategorienBestimmen_Click
  Kassengeändert = 0
 End If
' Call myEFrag("SET autocommit = 1")
 frm.SBez.BackColor = &HE0E0E0 ' hellgrau, vbgräulich&
 DoEvents
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away
 Dim altobTrans%
 altobTrans = obTrans
 If obTrans Then
'  If DBCn.State <> 0 Then DBCn.Close
  Set DBCn = Nothing ' geändert 21.10.22
 End If
 Call DBCnOpen
 If altobTrans <> 0 Then obTrans = 0: Resume nachFehler Else Resume
End If ' Err.Number = -2147467259
If Err.Number = -2147217833 Then
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborparameter", Empty))
 Do While Not rsc.EOF
  Debug.Print rsc!COLUMN_NAME
  If rsc!COLUMN_NAME = "Abkü" Then If SpMod(Len(rLa(i).Abkü), "laborparameter", rsc, rLa(i).Abkü) Then Exit Do
  If rsc!COLUMN_NAME = "Einheit" Then If SpMod(Len(rLa(i).Einheit), "laborparameter", rsc, rLa(i).Einheit) Then Exit Do
  If rsc!COLUMN_NAME = "Langtext" Then If SpMod(Len(rLa(i).Langtext), "laborparameter", rsc, rLa(i).Langtext) Then Exit Do
  rsc.Move 1
 Loop
 DBCn.BeginTrans: obTrans = 1
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' alleSpeichern

Function testmedarten()
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, T1!, T2!
 Dim SL As New SortierListe, SM As SortierMedi
 Lese.ProgStart
 medartenhier SL
 myFrag rs, "SELECT Medikament FROM `medarten`"
 T1 = Timer
 Do While Not rs.EOF
  Set SM = New SortierMedi
  SM.medi = rs!Medikament
  Set SM = SL.GetItem(SM)
'  IF Not SM Is Nothing THEN
'   Stop
'  END IF
'  SET rs1 = Nothing
'  myFrag rs1, "SELECT * FROM `medarten` WHERE medikament = '" & rs!Medikament & "'"
  rs.Move 1
 Loop
 T2 = Timer
 Debug.Print T2 - T1 & " Sekunden" ' bei 4000 Einträgen 0,48" vs. 20,6" mit medartenhier vs SELECT
End Function ' testmedarten

'Function testa()
'Dim rAf&, ohd%, od%
'Lese.ProgStart
'Dim rs As New ADODB.Recordset
'myFrag rs, "SELECT pat_id,notiz FROM quelle.`namen`"
'Do While Not rs.EOF
' ohd = obhierdmp(rs!Notiz, , , , od)
' myEFrag "UPDATE quelle.`namen` SET obhierdmp = " & ohd & ", obdmp = " & od & " WHERE pat_id = " & rs!Pat_id, rAf
' Debug.Print rAf
' rs.Move 1d
'Loop
'End FUNCTION ' testa

Function kassenspeichern(frm As Lese, pid$)
 Dim i%, rs As New ADODB.Recordset, kat$, keinetrans%, UKAS$, dokat%
 keinetrans = True
 Dim j%
 On Error GoTo fehler
 For i = 1 To UBound(rFa)
  dokat = 0
  Set rs = Nothing
  kat = vNS
  If rFa(i).Kasse = "" Then
   For j = 1 To UBound(rFa)
    If rFa(j).VKNr = rFa(i).VKNr And rFa(j).IK = rFa(i).IK And rFa(j).Kasse <> "" Then
     rFa(i).Kasse = rFa(j).Kasse
     Exit For
    End If
   Next
  End If
  myFrag rs, "SELECT kurzname,id,name,kateg FROM `kassenliste` WHERE vknr = '" & rFa(i).VKNr & "' AND ik = '" & rFa(i).IK & "'"  ' & IIf(rFa(i).Kasse = "", IIf(rFa(i).KKasse_2 = "", "", " AND name = '" & rFa(i).KKasse_2 & "'"), " AND name = '" & rFa(i).Kasse & "'")
  If rs.EOF Then ' 14.11.21
   Set rs = Nothing
   myFrag rs, "SELECT kurzname,id,name,kateg FROM `kassenliste` WHERE vknr = '" & rFa(i).VKNr & "'" ' & IIf(rFa(i).Kasse = "", IIf(rFa(i).KKasse_2 = "", "", " AND name = '" & rFa(i).KKasse_2 & "'"), " AND name = '" & rFa(i).Kasse & "'")
  End If
'  IF rFa(i).KID = 0 THEN Stop
  If rs.EOF Then
   dokat = True
  ElseIf (rs!Kateg = "" And (rs!name <> "" Or rs!kurzname <> "" Or rFa(i).Kasse <> "" Or rFa(i).KKasse_2 <> "")) Then
   dokat = True
   rFa(i).KID = rs!id
  End If
  If dokat Then
   UKAS = UCase$(rFa(i).Kasse)
   If UKAS = "" Then UKAS = UCase$(rFa(i).KKasse_2)
   If UKAS = "" Then If Not rs.EOF Then UKAS = UCase$(rs!name) ' dann nicht eof ' 14.11.21
   If UKAS = "" Then If Not rs.EOF Then UKAS = UCase$(rs!kurzname)
   If InStrB(UKAS, "Sozialhilfe") <> 0 Or InStrB(UKAS, "Sozialamt") <> 0 Or InStrB(UKAS, "SVA") <> 0 Or InStrB(UKAS, "SHV") <> 0 Then
    kat = "SHV"
   ElseIf InStrB(UKAS, "AOK") <> 0 Then ' instrb(ukas,"AOK")<>0
    kat = "AOK"
   ElseIf InStrB(UKAS, "BKK") <> 0 Or InStrB(UKAS, "BETRIEBSK") <> 0 Then
    kat = "BKK"
   ElseIf InStrB(UKAS, "IKK") <> 0 Then
    kat = "IKK"
   ElseIf InStrB(UKAS, "LKK") <> 0 Then
    kat = "LKK"
   ElseIf InStrB(UKAS, "BKN") <> 0 Or InStrB(UKAS, "KNAPP") <> 0 Then
    kat = "BKN"
   ElseIf InStrB(UKAS, "BARMKER") <> 0 Or InStrB(UKAS, "DAK") <> 0 Or InStrB(UKAS, "TKK") <> 0 Or InStrB(UKAS, "KKH") <> 0 Or InStrB(UKAS, "HEK") <> 0 Or InStrB(UKAS, "HMK") <> 0 Or InStrB(UKAS, "HKK") <> 0 Or InStrB(UKAS, "GEK") <> 0 Or InStrB(UKAS, "HZK") <> 0 Or InStrB(UKAS, "KEH") <> 0 Or InStrB(UKAS, "HAMBURG-MÜNCHENER") <> 0 Or InStrB(UKAS, "HANDELSKRANKENKASSE") <> 0 Or InStrB(UKAS, "TECHNIKER") <> 0 Or InStrB(UKAS, "HANSEATISCHE") <> 0 Or InStrB(UKAS, "AUS ") <> 0 Or InStrB(UKAS, "EK") <> 0 Then
    kat = "EK"
   End If
  End If
  If rs.EOF Then
'   Stop ' aus 6299 übernehmen
   On Error Resume Next
   If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
   If keinetrans = 0 Then keinetrans = Err.Number
   On Error GoTo fehler
   Dim GebOr$
   Select Case rFa(i).GebOr
    Case ""
     GebOr = vNS
    Case "1"
     GebOr = "1"
    Case "2"
     GebOr = "2"
   End Select
   InsKorr DBCn, DBCnS, "INSERT INTO `kassenliste`(vknr,ik,name,kurzname,go,kateg,eingef,pid) VALUES('" & rFa(i).VKNr & "','" & rFa(i).IK & "','" & rFa(i).Kasse & "','" & rFa(i).KKasse_2 & "','" & GebOr & "','" & kat & "'," & Format(Now(), "yyyymmddHHMMSS") & "," & pid & ")", rAF
   rFa(i).KID = myEFrag("SELECT last_insert_id()").Fields(0)
   Ausgeb rAF & " Kassen hinzugefügt (VK: " & rFa(i).VKNr & ", IK: " & rFa(i).IK & " => kkasse_2: " & rFa(i).KKasse_2 & "/kasse: " & rFa(i).Kasse & ")"
   Kassengeändert = True
   If keinetrans = 0 Then
    DBCn.BeginTrans: obTrans = 1
   End If
  ElseIf (rs!name = "" And rFa(i).Kasse <> "") Or (rs!kurzname = "" And rFa(i).KKasse_2 <> "") Or (rs!Kateg = "" And kat <> "") Then
'   Stop
'   On Error Resume Next
   If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
   If keinetrans = 0 Then keinetrans = Err.Number
   On Error GoTo fehler
   rAF = 0
   sql = "UPDATE `kassenliste` SET name = '" & rFa(i).Kasse & "', kurzname = '" & rFa(i).KKasse_2 & "', kateg = '" & kat & "',geaen=" & Format(Now(), "yyyymmddHHMMSS") & " WHERE vknr = '" & rFa(i).VKNr & "' AND ik = '" & rFa(i).IK & "'"
   Debug.Print sql
   Call myEFrag(sql, rAF)
   Ausgeb rAF & " Kassen mit Namen/Kategorie versehen (" & rFa(i).VKNr & ", IK: " & rFa(i).IK & " => kkasse_2: " & rFa(i).KKasse_2 & " /kasse: " & rFa(i).Kasse & " /Kateg: " & kat & ")"
   Kassengeändert = True
   If keinetrans = 0 Then
    DBCn.BeginTrans: obTrans = 1
   End If
  End If
 Next i
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kassenSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 '   rFaA.Open "SELECT * FROM `faelle` LEFT JOIN (SELECT DISTINCT vknr, kateg, name FROM `kassenliste`) AS kl ON faelle.vknr = kl.vknr WHERE pat_id = " & Me.PatID & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockReadOnly
End Function ' kassenspeichern

' vergleicht die hiesingen Funktionen mit denen in Mysql
Function testrr()
 Dim i&
 Dim rsr As New ADODB.Recordset
 Call Lese.ProgStart
 myFrag rsr, "SELECT rrsyst(rr) rs, rrdiast(rr) rd, rrzahl(bemerkung) rz, rr, bemerkung, pat_id, zeitpunkt FROM rr ORDER BY pat_id"
 If Not rsr.BOF Then
  Do While Not rsr.EOF
   Debug.Print rsr!Pat_id
   If rsr!rs <> holRRsyst(rsr!RR) Then Stop
   If rsr!rD <> holRRdiast(rsr!RR) Then Stop
   If rsr!rz <> holRRzahl(rsr!Bemerkung) Then Stop
   rsr.MoveNext
  Loop
  Debug.Print "Erfolg"
 End If
End Function ' testrr

Function rrParsen()
 Dim i%
 On Error GoTo fehler
 For i = 1 To UBound(rRr)
  Call dodoRRParse(rRr(i).RR, rRr(i).RRsyst, rRr(i).RRdiast)
  rRr(i).RRzahl = holRRzahl(rRr(i).Bemerkung)
  rRr(i).RRsyst = holRRsyst(rRr(i).RR)
  rRr(i).RRdiast = holRRdiast(rRr(i).RR)
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrParseSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rrParseSpeichern

Function holRRdiast%(RR$)
 Dim pos%, leng
 Dim ch As String * 1
 Dim rueck$
 pos = InStr(RR, "/")
 leng = Len(RR)
 If pos = 0 Then
  holRRdiast = "0"
  Exit Function
 End If
 Do While True
  pos = pos + 1
  If pos > leng Then
   Exit Do
  End If
  ch = Mid$(RR, pos, 1)
  If InStrB("0123456789", ch) <> 0 Then
   rueck = rueck & ch
  ElseIf (ch <> " " Or rueck <> "") Then
   Exit Do
  End If
 Loop
 If rueck = "" Then rueck = 0
 If CLng(rueck) > 199 Then rueck = Left$(rueck, 2)
 holRRdiast = rueck
End Function 'holRRdiast(RR$, RRdiast$)

Function holRRsyst%(RR$)
 Dim pos%, ppu%, schonkomma%, RRg$
 Dim ch As String * 1
 Dim rueck$
 RRg = UCase$(RR)
 pos = InStr(RRg, "/")
 If pos = 0 Then
  pos = InStr(RRg, " MM HG")
  If pos = 0 Then pos = InStr(RRg, " MMHG")
  ppu = InStr(RRg, "LETZTEN")
  If ppu > 0 And InStrB(RRg, "MESSUNGEN") <> 0 Then pos = ppu
 Else
  ppu = InStr(RR, "Puls") ' hier case sensitive,deshalb RR statt RRg, in mysql: ... binary ...
  If ppu > 0 And ppu < pos Then pos = ppu
 End If
 If pos = 0 Then pos = Len(RRg) + 1
 Do While True
  pos = pos - 1
  If pos <= 0 Then
   If rueck = "" Then rueck = "0"
   Exit Do
  End If
  ch = Mid$(RRg, pos, 1)
  If InStrB("0123456789,", ch) <> 0 And (schonkomma = 0 Or ch <> ",") Then
   If ch <> "," Or rueck <> "" Then rueck = ch & rueck
   If ch = "," Then schonkomma = 1
  ElseIf (ch <> " " And rueck <> "") Then
   Exit Do
  ElseIf ch = " " Then
   If rueck <> "" Then
    If CLng(rueck) > 50 Then
     Exit Do
    End If
   End If
  End If
 Loop
 pos = InStr(rueck, ",")
 If pos = 1 Then
  rueck = Mid$(rueck, 2)
 ElseIf pos > 1 Then
  rueck = Left$(rueck, pos - 1)
 End If
 If rueck < 30 Then rueck = 0
 If CLng(rueck) > 50000 Then
  If CLng(rueck) > 1000000 Then
   pos = Len(rueck)
   If Mid$(rueck, pos - 3, 1) = "7" Then
    rueck = Left$(rueck, 3)
   End If
  Else
   pos = Len(rueck)
   If Mid$(rueck, pos - 2, 1) = "7" Then
    rueck = Left$(rueck, pos - 3)
   End If
  End If
 End If
 If CLng(rueck) > 300 Then
  rueck = Left$(rueck, 3)
 End If
 holRRsyst = rueck
End Function ' holRRsyst

' s. SQL-Funktion holrrzahlahl
Function holRRzahl%(Bemk$)
 Const muster$ = "0123456789"
 Dim p0%, pos%, leng%, ppu%, Bemgr$, ch As String * 1
 Bemgr = UCase$(Bemk)
 pos = InStr(Bemgr, "MITTELWERT")
 If pos = 0 Then
   pos = InStr(Bemgr, "DURCHSCHNITT")
   If pos <> 0 Then pos = pos + 12
 Else
  pos = pos + 10
 End If
 If pos <> 0 Then
  p0 = pos
  leng = Len(Bemgr)
  ppu = InStr(Bemgr, "PULS")
  If ppu <> 0 And ppu > p0 Then leng = ppu
  Do While pos <= leng
   ch = Mid$(Bemgr, pos, 1)
   pos = pos + 1
   If ch = "," Then Exit Do
   If InStrB(muster, ch) <> 0 Then
    holRRzahl = holRRzahl * 10 + CInt(ch)
   Else
    If holRRzahl > 0 Then
     If ch = "-" Then
      If holRRzahl = 24 Then holRRzahl = 1
      Exit Do
     ElseIf ch = " " Then
      If InStrB(Mid$(Bemgr, pos), "TAGE") = 1 Then holRRzahl = 1
      Exit Do
     ElseIf ch = "/" Or ch = "." Then
      holRRzahl = 0
      Exit Do
     End If
    End If
   End If
   If holRRzahl > 999 Then
    holRRzahl = 1
    Exit Do
   End If
  Loop
 End If
 If holRRzahl = 0 Then holRRzahl = 1
End Function ' holholrrzahlahl(Bemk$, holrrzahl%)

' aufgerufen in alleSpeichern
Function rrParseSpeichern()
 Dim i%
 On Error GoTo fehler
 For i = 1 To UBound(rRr)
  Call do_RRParse(rRr(i).RR, rRr(i).Pat_id, rRr(i).Zeitpunkt, "Tabelle RR")
 Next i
 Exit Function
 If Not IsNull(rsAnm!Blutdruckwerte) Then
  Call do_RRParse(rsAnm!Blutdruckwerte, rNa(0).Pat_id, VorStDat, "An'bg.BW")
 End If
 If Not IsNull(rsAnm!RR) Then
  Call do_RRParse(rsAnm!RR, rNa(0).Pat_id, VorStDat, "An'bg.RR")
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrParseSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rrParseSpeichern

' in alleSpeichern
Function getDTyp$()
 Dim i%
 getDTyp = "?"
 For i = 1 To UBound(rDi)
  If (rDi(i).DiagSicherheit = "G" Or rDi(i).DiagSicherheit = "" Or rDi(i).DiagSicherheit = " ") Then '  And rDi(i).f6010 = 0
   If rDi(i).ICD = "O24.4" Then
    getDTyp = "g"
    Exit For
   ElseIf rDi(i).ICD = "R73.0" Then
    getDTyp = "p"
   ElseIf rDi(i).ICD Like "E1*" Then
    If (rDi(i).DiagText Like "*pankr*" Or rDi(i).DiagText Like "*pancr*" Or rDi(i).DiagText Like "*sekund*") And (rDi(i).DiagSicherheit = "G" Or rDi(i).DiagSicherheit = " ") Then ' And rDi(i).f6010 = 0
     getDTyp = "s"
     Exit For
    Else
     If rDi(i).ICD Like "E11*" Then
      getDTyp = "2"
      Exit For
     ElseIf rDi(i).ICD Like "E10*" Then
      getDTyp = "1"
      Exit For
     End If ' rDi(i).ICD LIKE "E11*" THEN else
    End If ' (rDi(i).DiagText LIKE "*pankr*" OR rDi(i).DiagText LIKE "*pancr*" OR rDi(i).DiagText LIKE "*sekund*") AND (rDi(i).DiagSicherheit = "G" OR rDi(i).DiagSicherheit = " ") AND rDi(i).f6010 = 0 THEN else
   End If ' rDi(i).ICD = "O24.4" THEN else
  End If ' (rDi(i).DiagSicherheit = "G" OR rDi(i).DiagSicherheit = "" OR rDi(i).DiagSicherheit = " ") AND rDi(i).f6010 = 0 THEN
 Next i ' i = 1 To UBound(rDi)
End Function ' getDTyp$()

' nur in alleSpeichern
Function MacheDiagnosen$(dmseit$) ' für AnaEintragen, MachSammelTab und DiagString
 Dim j&, k&, runde%
 On Error GoTo fehler
   
 Dim aktQB As Date, aktQE As Date
 For j = 1 To UBound(rFa)
  If rFa(j).BhFB < Now And rFa(j).BhFE1 > Now Or rFa(j).BhFE1 = 0 Then
   aktQB = rFa(j).BhFB
   aktQE = rFa(j).BhFE1 + 1
   If aktQE = 0 Then aktQE = Now
   Exit For
  End If
 Next j
 If aktQB = 0 Then
  For j = 1 To UBound(rFa)
   If rFa(j).BhFB > aktQB Then
    aktQB = rFa(j).BhFB
    aktQE = rFa(j).BhFE1 + 1
    If aktQE = 0 Then aktQE = Now
   End If
  Next j
 End If
    
' Dim Diag$()
' Dim ICD$()
' Dim DSic$()
' Dim f6010%()
' Dim DiagSe$()
' Dim DiagAttr$()
' Dim diagdat() As Date
' Dim obDauer() AS Boolean
' Dim gk() AS Boolean ' abgehakt
 DiagNr = 0
 ReDim Diag(0)
 ReDim ICD(0)
 ReDim DSic(0)
 ReDim DiagSe(0)
 ReDim DiagAttr(0)
 ReDim DiagAus(0)
 ReDim DiagiBm(0)
 ReDim diagdat(0)
 ReDim obDauer(0)
 ReDim obKasse(0)
 ReDim lKasse(0)
 ReDim f6010(0)
 ReDim f6011(0)
 ReDim gk(0)
 ReDim G1(0)
 ReDim G2(0)
 
 For runde = 1 To UBound(rDi)
  If rDi(runde).obDauer <> 0 Or (rDi(runde).DiagDatum >= aktQB And rDi(runde).DiagDatum < aktQE) Then
   ReDim Preserve Diag(DiagNr)
   ReDim Preserve ICD(DiagNr)
   ReDim Preserve DSic(DiagNr)
   ReDim Preserve DiagSe(DiagNr)
   ReDim Preserve DiagAttr(DiagNr)
   ReDim Preserve DiagAus(DiagNr)
   ReDim Preserve DiagiBm(DiagNr)
   ReDim Preserve diagdat(DiagNr)
   ReDim Preserve obDauer(DiagNr)
   ReDim Preserve obKasse(DiagNr)
   ReDim Preserve lKasse(DiagNr)
   ReDim Preserve f6010(DiagNr)
   ReDim Preserve f6011(DiagNr)
   ReDim Preserve gk(DiagNr)
   ReDim Preserve G1(DiagNr)
   ReDim Preserve G2(DiagNr)
   Diag(DiagNr) = rDi(runde).DiagText
   ICD(DiagNr) = rDi(runde).ICD
   DSic(DiagNr) = rDi(runde).DiagSicherheit
   DiagSe(DiagNr) = rDi(runde).DiagSeite
   DiagAttr(DiagNr) = rDi(runde).DiagAttr
   DiagAus(DiagNr) = rDi(runde).AusnBegr
   DiagiBm(DiagNr) = rDi(runde).intBemerk
   diagdat(DiagNr) = rDi(runde).DiagDatum
   obDauer(DiagNr) = rDi(runde).obDauer
   obKasse(DiagNr) = rDi(runde).obKasse
   lKasse(DiagNr) = rDi(runde).lKasse
   f6010(DiagNr) = rDi(runde).f6010
   f6011(DiagNr) = rDi(runde).f6011
   gk(DiagNr) = False
   DiagNr = DiagNr + 1
  End If
 Next runde
 Dim DTab() As CString ' Fake
 MacheDiagnosen = MachDiagnosen(Str$(rNa(0).Pat_id), DTab, dmseit, ohneTab:=True, Sort:=True)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MacheDiagnosen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MacheDiagnosen$()

Function testvor()
 Dim te As Date
 te = BDTtoDate("00000000")
 Call VorstellSetz(te)
End Function ' testvor

Function VorstellSetz(DaT) ' Datum setzen, an dem sich der Patient zum ersten Mal vorgestellt hat, 4102, 4109, 3610
 On Error GoTo fehler
 If Not obvorgestellt Then
  VorStDat = DaT
  If VorStDat <> 0 Then obvorgestellt = True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VorstellSetz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VorstellSetz(Dat$)

Function aufSplit%(ByVal q$, Optional Tz$)
Dim ind%, TzL%, pos%(), pakt%, i%
ql = q
On Error GoTo fehler
If Tz = vNS Then Tz = "#": TzL = 1 Else TzL = Len(Tz)
ind = 0: pakt = 1
ReDim pos(ind)
pos(ind) = 1
If Not IsNull(q) Then
 Do
  ind = ind + 1
  ReDim Preserve pos(ind)
  pos(ind) = InStr(pakt, q, Tz)
  If pos(ind) <> 0 Then
   pos(ind) = pos(ind)
   pakt = pos(ind) + 1
  Else
   pos(ind) = Len(q) + TzL
   Exit Do
  End If
 Loop
End If
If ind > 0 Then
 ReDim Arra(ind - 1)
 Arra(0) = Mid$(q, pos(i), pos(i + 1) - pos(i))
 For i = 1 To ind - 1
  Arra(i) = Mid$(q, pos(i) + TzL, pos(i + 1) - pos(i) - TzL)
 Next i
End If
aufSplit = ind - 1
ArraInd = aufSplit
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in aufSplit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit
'Function aufSplitalt%(ByVal q$, Optional Tz$)
'Dim ind%, TzL%
'On Error GoTo fehler
'If Tz = vNS THEN Tz = "#": TzL = 1 ELSE TzL = Len(Tz)
'ind = 0
'ReDim Arra(ind)
'If NOT ISNULL(q) THEN
' While InStrB(q, Tz) > 0
'  Arra(ind) = LEFT(q, InStr(q, Tz) - TzL)
'  ind = ind + 1
'  ReDim Preserve Arra(ind)
'  q = Right$(q, Len(q) - InStr(q, Tz) + 1 - TzL)
' Wend
' Arra(ind) = q
' ArraInd = ind
'END IF
'aufSplitalt = ind
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in aufSplit/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' aufSplit

'Function ZQuartSort$(Datum As Date) ' Für Abfragen mit Fallzuordnung
'Dim j AS String * 4, q AS String * 1
'On Error GoTo fehler
'j = YEAR(Datum)
'SELECT CASE Datum
' Case Is < CDate("1.4." + j): q = "1"
' Case Is < CDate("1.7." + j): q = "2"
' Case Is < CDate("1.10." + j): q = "3"
' Case Else: q = "4"
'End SELECT
'ZQuartSort = j + q
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuartSort/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT ' ZQuartSort
'
'End FUNCTION ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

'Function ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
'Dim j AS String * 4, Q AS String * 1
'On Error GoTo fehler
'j = YEAR(Datum)
'SELECT CASE Datum
' Case Is < CDate("1.4." + j): Q = "1"
' Case Is < CDate("1.7." + j): Q = "2"
' Case Is < CDate("1.10." + j): Q = "3"
' Case Else: Q = "4"
'End SELECT
'ZQuart = Q + j
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

Function RREintr(RR)
 On Error GoTo fehler
 ReDim Preserve rRr(UBound(rRr) + 1)
 rRr(UBound(rRr)).Pat_id = rNa(0).Pat_id
 rRr(UBound(rRr)).Zeitpunkt = messDatum
 rRr(UBound(rRr)).RR = RR
 rRr(UBound(rRr)).absPos = absPos
 rRr(UBound(rRr)).AktZeit = AktZeit
 rRr(UBound(rRr)).FID = rFa(UBound(rFa)).FID
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RREintr/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RREintr

' aufgerufen in dolies() (2x)
Function RezEintr(rez$, obLangrz%, Optional mitAutidem = True)
 On Error GoTo fehler
 Dim RRz As New ADODB.Recordset
 If Not IsNull(rez) Then
  Call aufSplit(rez)
  ReDim Preserve rRe(UBound(rRe) + 1)
  rRe(UBound(rRe)).Pat_id = rNa(0).Pat_id
  rRe(UBound(rRe)).Zeitpunkt = messDatum
  rRe(UBound(rRe)).QS = ZQSort(messDatum)
  rRe(UBound(rRe)).QT = ZQuart(messDatum)
  If obLangrz Then ' Langrezept
   If ArraInd > 0 Then rRe(UBound(rRe)).Rezkllang = Arra(1)
   rRe(UBound(rRe)).RKlnm = Left$(Arra(1), 2)
   rRe(UBound(rRe)).Rezeptklasse = Arra(2)
   rRe(UBound(rRe)).PZN = Arra(3)
   If ArraInd > 3 Then rRe(UBound(rRe)).Medikament = Arra(4)
   If ArraInd > 4 Then rRe(UBound(rRe)).Rezklkurz = Arra(5)
   If SafeArrayGetDim(auti) = 0 Then ' Privatrezept
    rRe(UBound(rRe)).auti = 1
   Else
    rRe(UBound(rRe)).auti = IIf(auti(0) = 1, 0, 1)
   End If
   rRe(UBound(rRe)).anzl = 1
  Else ' Rezept
   rRe(UBound(rRe)).Medikament = IIf(ArraInd > 0, Arra(1), vNS)
   rRe(UBound(rRe)).PZN = IIf(ArraInd > 1, Arra(2), vNS)
   If ArraInd > 3 Then rRe(UBound(rRe)).Rezklkurz = Arra(4)
   If SafeArrayGetDim(auti) = 0 Then ' Privatrezept
    rRe(UBound(rRe)).auti = 1
   Else
    If UBound(auti) < mdnr Then ReDim Preserve auti(mdnr)
    rRe(UBound(rRe)).auti = IIf(mitAutidem, auti(mdnr), 2)
   End If
   If SafeArrayGetDim(auti) = 0 Then ' Privatrezept
    rRe(UBound(rRe)).auti = 1
    rRe(UBound(rRe)).anzl = 1
   Else
    If UBound(auti) < mdnr Then ReDim Preserve auti(mdnr)
    If UBound(anzl) < mdnr Then ReDim Preserve anzl(mdnr)
    rRe(UBound(rRe)).auti = IIf(mitAutidem, auti(mdnr), 2)
    rRe(UBound(rRe)).anzl = MAX(anzl(mdnr), 1)
   End If
   mdnr = mdnr + 1
  End If
  rRe(UBound(rRe)).FID = rFa(UBound(rFa)).FID
  rRe(UBound(rRe)).AktZeit = AktZeit
  rRe(UBound(rRe)).Rezept = Arra(0)
  rRe(UBound(rRe)).absPos = absPos
  rRe(UBound(rRe)).lanrid = Lanr ' : Lanr = 0 ' Kommentar 21.3.21
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RezEintr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'RezEintr

' in dodoPLZ, do_Form_Current_AnBog, tubriefStandalone
Function DiagString$(Pat_id$, DiagTab() As CString, Optional VorDat As Date, Optional obBrief%, Optional dmseit$) ' für dynDiag, tubriefStandalone und dodoPLZ
 Dim runde%, rdDi As New ADODB.Recordset, sql$
 On Error GoTo fehler
 sql = "SELECT DiagSicherheit, DiagText, DiagSeite, DiagAttr, d.ICD, obdauer, COALESCE(d.f6010,0) f6010, obDauer<>0 j_obdauer,obKasse,lKasse,f6011, COALESCE(diagdatum,0) DiagDatum, AusnBegr, intBemerk, g1.rf,r.gi2 " & vbCrLf & _
       "FROM diagview d " & vbCrLf & _
       "LEFT JOIN `diagreihe` r ON d.icd = r.icd " & vbCrLf & _
       "LEFT JOIN `diagg1` g1 ON r.gi1 = g1.lfdnr " & vbCrLf & _
       "WHERE pat_id = " & Pat_id & " AND (d.obdauer <> 0 OR d.diagdatum >" & DatFor_k(VorDat) & ") " & vbCrLf & _
       "GROUP BY LEFT(d.icd,3), DiagSicherheit, d.f6010, LEFT(diagtext,17) " & vbCrLf & _
       "ORDER BY g1.rf,gi2 DESC,COALESCE(d.f6010,0),icd;"

' rdDi.Open sql, DBCn, adOpenDynamic, adLockReadOnly
' SET rdDi = myEFrag(sql)
' rdDi.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rdDi, sql
 If rdDi.BOF Then Exit Function
 runde = 0
 ReDim Diag(0)
 ReDim ICD(0)
 ReDim DSic(0)
 ReDim DiagSe(0)
 ReDim DiagAttr(0)
 ReDim DiagAus(0)
 ReDim DiagiBm(0)
 ReDim diagdat(0)
 ReDim obDauer(0)
 ReDim obKasse(0)
 ReDim lKasse(0)
 ReDim f6010(0)
 ReDim f6011(0)
 ReDim gk(0)
 ReDim G1(0)
 ReDim G2(0)
 Do While Not rdDi.EOF
  If rdDi!ICD Like "Z25*" Then ' Impfung
  ElseIf obBrief And (rdDi!ICD Like "M20.*" Or rdDi!ICD Like "M21.*" Or rdDi!ICD Like "Q66.*") Then ' erworbene Zehen / Extremitätendeformitäten, angeb. Fußdeformitäten,
  ElseIf obBrief And (rdDi!ICD Like "L84") Then ' Hornhautschwiele
  ElseIf obBrief And (rdDi!ICD Like "B35.*") Then ' Tinea
  ElseIf obBrief And (rdDi!ICD Like "K02.*") Then ' kariöse Zähne
  ElseIf obBrief And (rdDi!ICD = "R26.8") Then ' Multifaktorielle Motilitätsstörung
  ElseIf obBrief And (rdDi!ICD = "R29.6") Then ' Sturzneigung beim älteren Menschen
  ElseIf obBrief And (rdDi!ICD = "R52.2") Then ' Chronischer Schmerzpatient
  ElseIf obBrief And (rdDi!ICD = "R68.8") Then ' verminderte körperliche Aktivität
  Else
   On Error Resume Next
   Diag(runde) = rdDi!DiagText
   On Error GoTo fehler
   ICD(runde) = IIf(IsNull(rdDi!ICD), vNS, rdDi!ICD)
   DSic(runde) = IIf(IsNull(rdDi!DiagSicherheit), vNS, rdDi!DiagSicherheit)
   DiagSe(runde) = IIf(IsNull(rdDi!DiagSeite), vNS, rdDi!DiagSeite)
   DiagAttr(runde) = IIf(IsNull(rdDi!DiagAttr), vNS, rdDi!DiagAttr)
   DiagAus(runde) = IIf(IsNull(rdDi!AusnBegr), vNS, rdDi!AusnBegr)
   DiagiBm(runde) = IIf(IsNull(rdDi!intBemerk), vNS, rdDi!intBemerk)
   diagdat(runde) = rdDi!DiagDatum
   obDauer(runde) = IIf(IsNull(rdDi!j_obDauer), 0, -Abs(rdDi!j_obDauer))
   obKasse(runde) = IIf(IsNull(rdDi!obKasse), 0, rdDi!obKasse)
   lKasse(runde) = IIf(IsNull(rdDi!lKasse), 0, rdDi!lKasse)
   f6010(runde) = IIf(IsNull(rdDi!f6010), 0, rdDi!f6010)
   f6011(runde) = IIf(IsNull(rdDi!f6011), vNS, rdDi!f6011)
   G1(runde) = IIf(IsNull(rdDi!rf), 0, rdDi!rf)
   G2(runde) = IIf(IsNull(rdDi!gi2), 0, rdDi!gi2)
   gk(runde) = False
   runde = runde + 1
   ReDim Preserve Diag(runde)
   ReDim Preserve ICD(runde)
   ReDim Preserve DSic(runde)
   ReDim Preserve DiagSe(runde)
   ReDim Preserve DiagAttr(runde)
   ReDim Preserve DiagAus(runde)
   ReDim Preserve DiagiBm(runde)
   ReDim Preserve diagdat(runde)
   ReDim Preserve obDauer(runde)
   ReDim Preserve obKasse(runde)
   ReDim Preserve lKasse(runde)
   ReDim Preserve f6010(runde)
   ReDim Preserve f6011(runde)
   ReDim Preserve gk(runde)
   ReDim Preserve G1(runde)
   ReDim Preserve G2(runde)
  End If
  rdDi.MoveNext
 Loop
#Const debu = 0 ' noch in Formular
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", in DiagString)"
#End If
 DiagNr = runde
 DiagString = MachDiagnosen(Pat_id, DiagTab, dmseit)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", in DiagString 2)"
#End If
' DiagString = MachDiagnosen(Pat_id, DiagTab, dmseit)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiagString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagString$(Pat_id&)

' in DiagString
Function MachDiagnosen(Pat_id$, DiagTab() As CString, Optional dmseit$, Optional ohneTab%, Optional Sort%) As CString ' für DiagString
    Dim j&, k&, runde%, obkNeph%
    Dim aktD As CString
'    Dim did As New ADODB.Recordset
    On Error GoTo fehler
    obkNeph = obKeineNephropathie(Pat_id)
    If Not ohneTab Then ReDim DiagTab(0)
    For j = 0 To DiagNr - 1
     If Not IsNull(ICD(j)) And ICD(j) <> "" Then
     If Left$(ICD(j), 2) = "E1" Then
      If InStrB(Diag(j), "ankr") > 0 Or InStrB(Diag(j), "ekund") > 0 Then
      Else
       Select Case Mid$(ICD(j), 3, 1)
        Case "0"
         Diag(j) = "Diabetes mellitus Typ 1"
        Case "1"
         Diag(j) = "Diabetes mellitus Typ 2"
        Case "2"
         Diag(j) = "Diabetes mellitus in Verbindung mit Fehl- oder Mangelernährung"
        Case "3"
         Diag(j) = "Diabetes mellitus (sekundär)"
        Case "4"
         Diag(j) = "Diabetes mellitus"
       End Select
'       myFrag did, "SELECT COALESCE(min(`diabetes seit`),'') dd FROM anamnesebogen WHERE pat_id=" & Pat_id
'       IF did!dd <> "" THEN Diag(j) = Diag(j) & " seit " & did!dd
        If dmseit <> "" And Not InStrB(Diag(j), "seit") Then Diag(j) = Diag(j) & " seit " & dmseit
'       SET did = Nothing
      End If
     End If
     If Left$(ICD(j), 5) = "O24.4" Then
      Diag(j) = "Gestationsdiabetes"
     End If
     If Left$(ICD(j), 5) = "N08.3" Then
      If InStrB(Diag(j), "Glomeruläre Krankheiten") > 0 Then
       Diag(j) = "Diabetische Nephropathie"
      End If
     End If
     If InStrB(Diag(j), "Niereninsuff.") > 0 Then Diag(j) = REPLACE$(Diag(j), "Niereninsuff.", "Niereninsuffizienz")
     If Left$(ICD(j), 5) = "G99.0" Then
      If InStrB(Diag(j), "Autonome Neuropathie bei endokrinen") > 0 Then
       Diag(j) = "Diabetische autonome Neuropathie"
      End If
     End If
     If Left$(ICD(j), 5) = "I79.9" Or Left$(ICD(j), 5) = "I79.2" Then
      If InStrB(Diag(j), "Periphere Angiopathie bei anderenorts") > 0 Then
       Diag(j) = "Periphere Angiopathie"
      End If
     End If
' M14.2 = Diabetische Arthropathie, schon richtig
     If Left$(ICD(j), 5) = "M14.6" Then
      If InStrB(Diag(j), "Neuropathische Arthropathie ") > 0 Then
       Diag(j) = "Diabetische Osteoarthropathie" ' "Charcot-Fuß"
      End If
     End If
     If Left$(ICD(j), 5) = "M36.8" Then
      If InStrB(Diag(j), "anderenorts") > 0 Then
       Diag(j) = "Diabetische Bindegewebserkrankung"
      End If
     End If
     If Left$(ICD(j), 3) = "E66" Then
      Diag(j) = "Übergewicht"
     End If
     If InStrB("K76.0 K76.9 K71.6 K71.7 K77.8", Left$(ICD(j), 5)) > 0 _
      And InStrB(Diag(j), "nbek") = 0 And InStrB(Diag(j), "nklar") = 0 And InStrB(DiagAttr(j), "nklar") = 0 Then
'     IF LEFT(ICD(j), 5) = "K77.8" AND _
       instrb(Diag(j), "Leberkrankheiten bei sonstigen anderenorts klassifizierten Krankheiten") > 0 THEN
       Diag(j) = "Hepatopathie"
     End If
     If Left$(ICD(j), 3) = "L89" Then
      Diag(j) = "Diabetisches Fußsyndrom "
      If Mid$(ICD(j), 5, 1) < "9" Then _
       Diag(j) = Diag(j) + " im Stadium Wagner "
      Select Case Mid$(ICD(j), 5, 1)
       Case "1": Diag(j) = Diag(j) + "0"
       Case "2": Diag(j) = Diag(j) + "1"
       Case "3": Diag(j) = Diag(j) + "2"
       Case "4": Diag(j) = Diag(j) + "3"
      End Select
     End If
'
'     IF LEFT(ICD(j), 5) = "H28.0" THEN Diag(j) = " Diabetische Katarakt"
'     IF LEFT(ICD(j), 5) = "H36.0" THEN Diag(j) = " Diabetische Retinopathie"
'     IF LEFT(ICD(j), 5) = "G63.2" THEN Diag(j) = " Diabetische Polyneuropathie"
     Select Case DiagSe(j)
      Case "R"
       If Right$(Diag(j), 3) <> " re" Then Diag(j) = Diag(j) + " re"
      Case "L"
       If Right$(Diag(j), 3) <> " li" Then Diag(j) = Diag(j) + " li"
      Case "B"
       If InStrB(Diag(j), "bds.") = 0 Then Diag(j) = Diag(j) + " bds."
     End Select
     If DiagAttr(j) <> "" Then Diag(j) = Diag(j) + " (" + DiagAttr(j) + ")"
      Diag(j) = REPLACE$(Diag(j), ", nicht näher bezeichnet", "")
     End If
    Next j
' Doppelte entfernen
    For j = 0 To DiagNr - 1
     For k = j + 1 To DiagNr - 1
      If Diag(j) = Diag(k) And ICD(j) = ICD(k) And f6010(j) = f6010(k) Then
       If (obDauer(k) <> 0 And obDauer(j) = 0) Then
        If (DSic(j) = "g" Or DSic(j) = " ") And (DSic(k) = "V" Or DSic(k) = "Z") Then
         Diag(k) = vNS
         ICD(k) = vNS
         obDauer(j) = True
        Else
         Diag(j) = vNS
         ICD(j) = vNS
        End If
       Else
        If (DSic(k) = "g" Or DSic(k) = " ") And (DSic(j) = "V" Or DSic(j) = "Z") Then
         Diag(j) = vNS
         ICD(j) = vNS
         If obDauer(j) And Not obDauer(k) Then obDauer(k) = True
        Else
         Diag(k) = vNS
         ICD(k) = vNS
         If obDauer(k) And Not obDauer(j) Then obDauer(j) = True
        End If
       End If
      End If
     Next k
    Next j
'    IF obAnBog THEN    ' 6.10.13 auskommentiert, da sonst im Arztbrief V.a. und Z.n. nicht erscheint (tubriefStandalone)
     For j = 0 To DiagNr - 1
      Select Case DSic(j)
       Case "V"
        If Left$(Diag(j), 4) <> "V.a." Then Diag(j) = "V.a. " + REPLACE$(REPLACE$(REPLACE$(REPLACE$(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
       Case "Z"
        If Left$(Diag(j), 4) <> "Z.n." Then Diag(j) = "Z.n. " + REPLACE$(REPLACE$(REPLACE$(REPLACE$(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
       Case "A"
        If Left$(Diag(j), 8) <> "Ausschlu" Then Diag(j) = "Ausschluss " + REPLACE$(REPLACE$(REPLACE$(REPLACE$(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
      End Select
      If obDauer(j) = 0 Then
       Diag(j) = "(q) " & Diag(j)
      End If
     Next
'    END IF
'   MachDiagnosen = vNS
   Set MachDiagnosen = New CString
   Dim rua%, rue%
   Dim ohneNotwend%
   
   If Sort Then
    ohneNotwend = True
    rua = 0
    rue = 17
   Else
    rua = 1
    rue = 1
   End If
   For runde = rua To rue
    For j = 0 To DiagNr - 1
     If Not IsNull(ICD(j)) And ICD(j) <> "" Then ' And f6010(j) = 0
      If (Not Sort And (Diag(j) <> "" Or ICD(j) <> "")) Or _
      (Sort And (Diag(j) <> "" Or ICD(j) <> "") And _
        (Not ohneNotwend Or InStrB(Diag(j), "otwend") = 0) And _
        (Not obkNeph Or InStrB(ICD(j), "N08.3") = 0)) Then
       If (Not Sort Or ((runde = 0 And ((Left$(ICD(j), 2) = "E1" And Mid$(ICD(j), 3, 1) >= "0" And Mid$(ICD(j), 3, 1) <= "4" And obDauer(j) <> 0) Or InStrB("O24.4 R73.0", Left$(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And runde = 1 And (InStrB("N08.3", Left$(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And runde = 2 And (InStrB("N17 N18 N19", Left$(ICD(j), 3)) > 0)) Or _
        (Not gk(j) And runde = 3 And (InStrB("G59 G62", Left$(ICD(j), 3)) <> 0 Or _
             InStrB("H28.0 H36.0 G63.2 G99.0 I79.2 K77.8 K76.0 K76.9 K71.7 K71.6", Left$(ICD(j), 5)) <> 0 Or _
             InStrB("L89. I70.", Left$(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And runde = 4 And (Left$(ICD(j), 3) = "I10" Or Left$(ICD(j), 3) = "I15")) Or _
        (Not gk(j) And runde = 5 And (Left$(ICD(j), 2) = "I1" Or ICD(j) = "H35.0")) Or _
        (Not gk(j) And runde = 6 And _
             (InStrB("E65 E66 E67 E68", Left$(ICD(j), 3)) > 0 Or InStrB("E78. F17.", Left$(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And runde = 7 And _
             (InStrB("I25. I20. I21. I50.", Left$(ICD(j), 4)) > 0 Or InStrB("Z95.1 Z92.1", Left$(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And runde = 8 And (InStrB("R60 R01", Left$(ICD(j), 3)) > 0 Or InStrB("I83.1", Left$(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And runde = 9 And (InStrB("I63 I64", Left$(ICD(j), 3)) > 0)) Or _
        (Not gk(j) And runde = 10 And (InStrB("I7", Left$(ICD(j), 2)) > 0 Or InStrB("Z95.88", Left$(ICD(j), 6)) > 0)) Or _
        (Not gk(j) And runde = 11 And (InStrB("M36.8 T87.4", Left$(ICD(j), 5)) <> 0 Or _
             InStrB("M14. T89. T79.", Left$(ICD(j), 4)) <> 0)) Or _
        (Not gk(j) And runde = 12 And obDauer(j) <> 0 And (DSic(j) = "G" Or DSic(j) = " ")) Or _
        (Not gk(j) And runde = 13 And obDauer(j) <> 0 And DSic(j) = "V") Or _
        (Not gk(j) And runde = 14 And obDauer(j) = 0 And (DSic(j) = "G" Or DSic(j) = " ")) Or _
        (Not gk(j) And runde = 15 And obDauer(j) = 0 And DSic(j) = "V") Or _
        (Not gk(j) And runde = 16 And DSic(j) = "Z") Or _
        (Not gk(j) And runde = 17 And DSic(j) = "A"))) Then
'       MachDiagnosen = MachDiagnosen + IIf(runde > 0 AND runde < 4, " ", vNS) + Diag(j) + vbTab + "[" + ICD(j) + "]" + vbVerticalTab
        Set aktD = New CString
        If Sort Then
         aktD.AppVar Array(IIf(runde > 0 And runde < 4, " ", vNS), Diag(j), vbTab, "[", ICD(j), "]")
        Else
         aktD.AppVar Array(IIf(G1(j) = 1 And G2(j) > 50 And G2(j) < 97, " ", ""), Diag(j), vbTab, "[", ICD(j), "]")
'         aktD.AppVar Array(IIf(InStrB("N08.3 H28.0 H36.0 G63.2 G99.0 I79.2 K77.8 K76.0 K76.9 K71.7 K71.6", Left$(ICD(j), 5)) > 0 OR InStrB("N17. N18. N19. L89. I70.", Left$(ICD(j), 4)) > 0, " ", vNS), Diag(j), vbTab, "[", ICD(j), "]")
        End If
        MachDiagnosen.AppVar Array(aktD, vbVerticalTab)
        If Not ohneTab Then
         Set DiagTab(UBound(DiagTab)) = New CString
         DiagTab(UBound(DiagTab)).Append aktD
         ReDim Preserve DiagTab(UBound(DiagTab) + 1)
        End If
        gk(j) = True ' abgehakt
       End If ' (Runde = 0 AND (LEFT(icd(j), 2) = "E1") OR instrb("O24.4 R73.0", LEFT(icd(j), 5) > 0)) Or
      End If ' Diag(j) <> "" OR icd(j) <> "" THEN
     End If ' NOT ISNULL(ICD(j)) AND ICD(j) <> "" THEN
    Next ' j = 0 To DiagNr - 1
   Next ' Runde = 0 To 17
   If Not ohneTab Then If UBound(DiagTab) > 0 Then ReDim Preserve DiagTab(UBound(DiagTab) - 1)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachDiagnosen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machDiagnosen$()

Function obKeineNephropathie%(Pat_id$, Optional obMakroAlb%)
 Dim lAlbS$, lKreS$, sql$
' Dim rsAdo As New ADODB.Recordset
 Dim Labs As labtyp
 On Error GoTo fehler
 obMakroAlb = 0
 On Error Resume Next
'nochmal:
' sql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich AS NB " & _
'  "FROM `laborxus` u LEFT JOIN laborxwert w ON u.RefNr= w.RefNr " & _
'  "WHERE pat_id = " + CStr(Pat_id) + " AND NOT EXISTS (SELECT * FROM laborneu WHERE pat_id = u.pat_id AND abkü = w.Abkü AND wert = w.wert AND zeitpunkt > u.Eingang -3 AND zeitpunkt < u.Eingang+6)"
' sql = "SELECT * FROM labor1 WHERE pat_id = " & Pat_id & " UNION SELECT * FROM labor2 WHERE pat_id = " & Pat_id
  
 'lAlbS = Dtb.OpenRecordset("SELECT iif(ISNULL(wert),iif(ISNULL(kommentar),"""",kommentar),wert) AS erg FROM (" + sql + ") AS sql1 WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') ORDER BY zeitpunkt DESC")!erg
' myFrag rsAdo, "SELECT iif(ISNULL(wert),iif(ISNULL(kommentar),"""",kommentar),wert) AS erg FROM (" + sql + ") AS sql1 WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') ORDER BY zeitpunkt DESC"
' SET rsAdo = myEFrag("SELECT IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) AS erg FROM (" & sql & ") AS sql1 WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND (abkü <> 'ALBU' OR wert LIKE '%<%') ORDER BY zeitpunkt DESC")
' SET rsAdo = LabEPat(AlbCre, Pat_id)
' IF Not rsAdo.EOF THEN lAlbS = rsAdo!wert

 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(LA_AlbCre, CLng(Pat_id)) ' erster Aufruf: 0,34s
 If Labs.Abkü <> "" Then lAlbS = Labs.WertSg
 
 On Error GoTo fehler
 If lAlbS = "" Then GoTo schluss
 If InStrB(lAlbS, "<") = 0 Then
  lAlbS = REPLACE$(lAlbS, ".", ",")
  If Not IsNumeric(lAlbS) Then GoTo schluss
  If CDbl(lAlbS) >= 200 Then obMakroAlb = True
  If CDbl(lAlbS) >= 10 Then GoTo schluss ' setzen wir mal 10 als Grenze an
 End If
 On Error Resume Next
' rsAdo.Close
' myFrag rsAdo, "SELECT iif(ISNULL(wert),iif(ISNULL(kommentar),"""",kommentar),wert) AS erg FROM (" + sql + ") AS sql1 WHERE abkü = ""KREA"" ORDER BY zeitpunkt DESC"
' SET rsAdo = myEFrag("SELECT iif(ISNULL(wert),iif(ISNULL(kommentar),"""",kommentar),wert) AS erg FROM (" + sql + ") AS sql1 WHERE abkü IN (""CREAT"",""KRE02"", ""KREA"", ""KREA02"", ""KRES"") ORDER BY zeitpunkt DESC")
' SET rsAdo = LabEPat(Krea, Pat_id)
' IF Not rsAdo.EOF THEN lKreS = rsAdo!wert
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(LA_Krea, CLng(Pat_id))
 If Labs.Abkü <> "" Then lKreS = Labs.WertSg
' lKreS = Dtb.OpenRecordset("SELECT iif(ISNULL(wert),iif(ISNULL(kommentar),"""",kommentar),wert) AS erg FROM (" + sql + ") AS sql1 WHERE abkü IN (""CREAT"",""KRE02"", ""KREA"", ""KREA02"", ""KRES"") ORDER BY zeitpunkt DESC")!erg
 On Error GoTo fehler
 If IsNumeric(lKreS) Then
  If lKreS <> 0 Then
  If CDbl(REPLACE$(lKreS, ".", ",")) >= 1.3 Then
   GoTo schluss
  End If
  obKeineNephropathie = -1
  End If
 End If
schluss:
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obKeineNephropathie/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obKeineNephropathie%(Pat_id&, Optional obMakroAlb%)

Public Function GesNameFn(NVorsatz$, Nachname$, Titel, Vorname$, Optional GebDat As Date) As CString  ' s.a. Gesnam(
 On Error GoTo fehler
 Set GesNameFn = New CString
 GesNameFn.AppVar Array(NVorsatz, IIf(LenB(NVorsatz) = 0, vNS, " "), Nachname, ", ")
 If IsNull(Titel) Or LenB(Titel) = 0 Then
 Else
  GesNameFn.AppVar Array(Titel, ", ")
 End If
 GesNameFn.Append Vorname
 If GebDat <> 0 Then
  GesNameFn.AppVar Array(", *", Format$(GebDat, "dd/mm/yy"))
 End If
' GesName = NVorsatz + IIf(LenB(NVorsatz) = 0, vNS, " ") + Nachname & ", " & IIf(ISNULL(Titel) OR LenB(Titel) = 0, vNS, Titel + ", ") + Vorname
' IF GebDat > 0 THEN GesName = GesName + ", *" + Format$(GebDat, "dd/mm/yy")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesName/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GesName

Function PatInit()
' Pat_id = 0
 On Error GoTo fehler
 NEin0 = 0
' lfdPatNr = 0
   lRR = vNS
   gesRR = vNS ' RR aus Turbomed
'   Versicherung = vns
   AlbErhöht = vNS
   AlbDat = 0
   lHbA1c = vNS
   lHbA1cDat = 0
   arr = vNS ' Aktueller Blutdruck
'   obMedPlan = False
'   obMedPlanGelesen = False
   obvorgestellt = False
' Diagnosedaten
 DiagNr = 0
' Anamnesebogen
 DMSchulz = 0
 RRSchulz = 0
 DMSchL = 0
 Tkz = 0
 VorStDat = 0
 obMedNetz = 0
 
 RRSchulz = 0
 dmpdat = 0 'Null
' FID = 0
 ReDim Diag(DiagMaxZahl)
 ReDim ICD(DiagMaxZahl)
 ReDim DSic(DiagMaxZahl)
 ReDim DiagSe(DiagMaxZahl)
 ReDim DiagAttr(DiagMaxZahl)
 ReDim DiagAus(DiagMaxZahl)
 ReDim DiagiBm(DiagMaxZahl)
 ReDim diagdat(DiagMaxZahl)
 ReDim obDauer(DiagMaxZahl)
 ReDim obKasse(DiagMaxZahl)
 ReDim lKasse(DiagMaxZahl)
 ReDim f6010(DiagMaxZahl)
 ReDim f6011(DiagMaxZahl)
 ReDim gk(DiagMaxZahl)
 ReDim G1(DiagMaxZahl)
 ReDim G2(DiagMaxZahl)
 maxBhFB = 0
 imaxBhFB = 0
 pMpnr = 0 ' 10.7.10: ab jetzt Mpnr patientenspezifisch
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PatInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PatInit()
Function Stutz$(q$)
 On Error GoTo fehler
    If Right$(q, 1) = "^" Then
     Stutz = Left$(q, Len(q) - 1)
    Else
     Stutz = q
    End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Stutz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Stutz

Function PraxisHbA1c(EintrArt$, EintrInh$)
 Dim Ep%, ZwS$, erg$
 If UCase$(EintrArt) = "HBA1C" Then
  FStG = "E"
  Abkü = "HBA1C"
  AbküLabor = "Eigenlabor"
  Call LaborEintr0
  If obLaborEintrag Then
   rLa(ls).Einheit = "%"
  
   Ep = InStr(EintrInh, ",")
   If Ep > 1 Then
    ZwS = Left$(EintrInh, Ep - 1)
    erg = vNS
    If IsNumeric(Right$(ZwS, 2)) Then
     erg = Right$(ZwS, 2)
    ElseIf IsNumeric(Right$(ZwS, 1)) Then
     erg = Right$(ZwS, 1)
    End If
    If erg <> "" Then
     ZwS = Mid$(EintrInh, Ep + 1, 1)
     If IsNumeric(ZwS) Then
      erg = erg & "," & ZwS
     End If
    End If
   Else
    Ep = InStr(EintrInh, "%")
    If Ep > 0 Then
     ZwS = Trim$(Left$(EintrInh, Ep - 1))
     If IsNumeric(Right$(ZwS, 2)) Then
      erg = Right$(ZwS, 2)
     ElseIf IsNumeric(Right$(ZwS, 1)) Then
      erg = Right$(ZwS, 1)
     End If
    End If
   End If
   If IsNumeric(erg) Then
    rLa(ls).Wert = erg
    rLa(ls).LangtextVW = LTEinfüg("HbA1c")
    rLa(ls).Einheit = "%"
    rLa(ls).KommentarVW = KomEinfüg("Praxis-HbA1c")
'    rLa(ls).Kommentar = "Praxis-HbA1c"
   End If
  End If
 End If
End Function ' PraxisHbA1c
Function LTEinfüg&(Langtext$)
  Dim i&
  On Error GoTo fehler
  If Langtext = lLang Then
   LTEinfüg = lLangVW
  Else
   For i = 1 To 2
    If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
'   myFrag rsAdo, "SELECT langtext,langtextvw FROM `laborlangtext` WHERE langtext = '" & Langtext & "'"
    Set rsAdo = myEFrag("SELECT langtext,langtextvw FROM `laborlangtext` WHERE langtext = '" & Langtext & "'")
    If rsAdo.EOF Then
     InsKorr DBCn, DBCnS, "INSERT INTO `laborlangtext`(`langtext`) values ('" & Langtext & "')", rAF
    Else
     Exit For
    End If
   Next i
   LTEinfüg = rsAdo!LangtextVW
   lLangVW = LTEinfüg
   lLang = Langtext
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LTEinfüg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LTEinfüg&(Langtext$)

' verwendet in dolies, laboreintr0 und PraxisHbA1c
Function KomEinfüg&(Kommentar$)
  Dim i&
  On Error GoTo fehler
  If Kommentar = lKomm Then
   KomEinfüg = lKommVW
  Else
   For i = 1 To 2
    Set rsAdo = Nothing
'   myFrag rsAdo, "SELECT Kommentar,Kommentarvw FROM laborkommentar WHERE Kommentar = '" & Kommentar & "'"
    If lies.obMySQL = 0 And Len(Kommentar) > 255 Then
    ' nicht nachvollziehbare Besonderheit bei:
' 8430 Serum Erregerspezifische IgG-Ak nachweisbar. Der Befund ist serologisch mit einer aktuellen oder länger zurückliegenden Infektion odereiner Immunisierung vereinbar. Zur Abklärung der aktuellen Infektion (falls klinisch der Verdacht besteht), empfiehlt sich die Bestim- mung
'   hier geht auch bei Access nur "%", nicht "*"!
     Set rsAdo = myEFrag("SELECT Kommentar,Kommentarvw FROM laborkommentar WHERE Kommentar LIKE """ & Kommentar & "%" & """", rAF)
    Else
     Set rsAdo = myEFrag("SELECT Kommentar,Kommentarvw FROM laborkommentar WHERE Kommentar = '" & Kommentar & "'", rAF)
    End If
    If rsAdo.BOF Then
     InsKorr DBCn, DBCnS, "INSERT INTO `laborkommentar`(Kommentar) values ('" & Kommentar & "')", rAF
    Else
     Exit For
    End If
   Next i
   KomEinfüg = rsAdo!KommentarVW
   lKomm = Kommentar
   lKommVW = KomEinfüg
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LTEinfüg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function LaborEintr0()
 On Error GoTo fehler
 Dim obneu%, obüber%, i%
' Abkü = IIf(ISNULL(Abkü), vns, Abkü) 'If ISNULL(rLab!Abkü) THEN rLab!Abkü = vns
' Einheit = ZeichenSatz(IIf(ISNULL(Einheit), vns, Einheit)) ' IF ISNULL(rLab!Einheit) THEN rLab!Einheit = vns
 obneu = -1
 For i = 1 To UBound(rLa)
  If rLa(i).Zeitpunkt = messDatum And rLa(i).Abkü = Abkü Then
   If rLa(i).FertigStGrad = "T" And FStG = "E" Then
    obüber = -1: obneu = 0: ls = i
   ElseIf FStG = "T" Then
    obneu = 0
   End If
   Exit For
  End If
 Next i
 obLaborEintrag = (obneu Or obüber)
 If obLaborEintrag Then
  If obneu Then ReDim Preserve rLa(UBound(rLa) + 1): ls = UBound(rLa)
  rLa(ls).Pat_id = rNa(0).Pat_id
  rLa(ls).Zeitpunkt = messDatum
  rLa(ls).FertigStGrad = FStG
'  rLa(ls).Labor = AbküLabor
  rLa(ls).Abkü = Abkü
  rLa(ls).Pat_id = rNa(0).Pat_id
  rLa(ls).AktZeit = AktZeit
  rLa(ls).FID = rFa(UBound(rFa)).FID
  rLa(ls).absPos = absPos
  rLa(ls).Langtext = vNS
  rLa(ls).LangtextVW = LTEinfüg&(rLa(ls).Langtext)
  rLa(ls).Kommentar = vNS
  rLa(ls).KommentarVW = KomEinfüg&(rLa(ls).Kommentar)
  rLa(ls).Einheit = "kA"
  rLa(ls).Wert = vNS
 End If
 FStG = vNS: Abkü = vNS
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborEintr0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborEintr0()

'Function Laboreintr1()
' ON Error GoTo fehler
' IF rLa(ls).Kommentar <> "" AND NOT ISNULL(rLa(ls).Kommentar) THEN rLa(ls).KommentarVW = KomEinfüg&(rLa(ls).Kommentar)
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborEintr1/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' LaborEintr1()

'#If False THEN
'Function AlleLaborParameter()
' Dim rLab AS DAO.Recordset
' Dim rLaP AS DAO.Recordset, rLN AS DAO.Recordset, fld AS DAO.Field
  'On Error GoTo fehler
 'Call dtbInit
 'Set rLab = Dtb.OpenRecordset("LaborUnion")
 '#If False THEN
'  SET rLaP = Dtb.OpenRecordset("LaborParameter")
  'Set rLN = Dtb.OpenRecordset("LaborParameterNeu")
  'Do While Not rLaP.EOF
'   IF Not (rLaP!gruppe = 0 AND rLaP!reihe = 0 AND ISNULL(rLaP!unm) AND ISNULL(rLaP!onm) AND ISNULL(rLaP!unw) AND ISNULL(rLaP!onw)) THEN
   '' rLaP.Delete
'    rLN.AddNew
    'For Each fld In rLaP.Fields
'     rLN.Fields(fld.Name) = fld.Value
    'Next fld
    'rLN.Update
   'END IF
   'rLaP.Move 1
  'Loop
 '#END IF
' rLab.MoveFirst
' Do While Not rLab.EOF
  'Call LaborParameter(rLab)
  'rLab.MoveNext
' Loop
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AlleLaborParameter/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' AlleLaborParameter

'Function LaborParameter(rLab AS DAO.Recordset)
''Dim rLP AS DAO.Recordset, rlaL AS DAO.Recordset
'Dim Abkü$, Einheit$, Langtext$
'Dim obltvw%
'On Error Resume Next
'obltvw = IIf(Err.Number = 0, -1, 0)
'On Error GoTo fehler
'Abkü = IIf(ISNULL(rLab!Abkü), vns, trim$(rLab!Abkü)) 'If ISNULL(rLab!Abkü) THEN rLab!Abkü = vns
'Einheit = Zeichensatz(IIf(ISNULL(rLab!Einheit), vns, trim$(rLab!Einheit))) ' IF ISNULL(rLab!Einheit) THEN rLab!Einheit = vns
'If Einheit = "" THEN Einheit = "kA" ' 2.10.: ist so in Turbomed, => auch für die Direktübernahme so hindrehen
'Langtext = vns
'If obltvw THEN
' SET rlaL = TabÖff("LaborLangtext", "LangtextVW")
' IF Not rlaL Is Nothing AND NOT ISNULL(rLab!LangtextVW) THEN
  'rlaL.Seek "=", rLab!LangtextVW
  'If Not rlaL.NoMatch THEN
'   Langtext = rlaL!Langtext
  'END IF
 'END IF
'Else
' Langtext = IIf(ISNULL(rLab!Langtext), vns, rLab!Langtext)
'END IF
'Set rLP = TabÖff("LaborParameter", "Abkü")
'If Abkü <> "" THEN
' rLP.Seek "=", Abkü, Einheit
' IF rLP.BOF THEN
  'rLP.AddNew
 'ElseIf rLP.NoMatch THEN
'  rLP.AddNew
' END IF
' IF rLP.EditMode = 2 THEN
  'If obÜProt THEN Print #322, "Neuer Laborparameter: " + Abkü + " `" + Einheit + "`"
  'rLP!Abkü = Abkü
  'rLP!Einheit = Einheit
 'ElseIf (rLP!Langtext <> Langtext) THEN 'Or (rLP!Kommentar <> rLab!Kommentar AND NOT ISNULL(rLab!Kommentar)) THEN
'  rLP.Edit
' Else
  'rLPbm = rLP.Bookmark
  'GoTo fertig:
 'END IF
 'If rLP.EditMode = 1 OR rLP.EditMode = 2 THEN
'  IF NOT ISNULL(Langtext) AND Langtext <> "" THEN
   'rLP!Langtext = Langtext
  'END IF
''  rLP!Kommentar = rLab!Kommentar
  'rLP!AktZeit = AktZeit
  'rLP.Update
  'rLPbm = rLP.LastModified
 'END IF
'END IF
' GoTo fertig
'test:
' MsgBox "Hier Fehler bei Einheit"
' Resume
'fertig:
' rLP.Close
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborParameter/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' LaborParameter(rLab AS dao.recordset)
'#END IF

'auskommentiert 10.10.22
'Function formInhMach()
' MsgBox "Stop durch Aufruf von forminhmach()"
' Stop
' Dim nr&
' Dim altpat_id&, altZeitpunkt#, altform_id&
' Dim myA As New Adodb.Connection, fn As Adodb.Recordset, fk As New Adodb.Recordset, ff As New Adodb.Recordset
' myA.Open CStrAcc & QmdB '  aVerz & "quelle.mdb"
' Call myEFrag("DELETE FROM `forminhfeld`",,myA)
' Call myEFrag("DELETE FROM `forminhkopf`",,myA)
'' myEfag("ALTER TABLE `forminhkopf` alter column foid counter (1,1)",,myA)
' fk.Open "FormInhKopf", myA, adOpenDynamic, adLockOptimistic
' ff.Open "FormInhFeld", myA, adOpenDynamic, adLockOptimistic
' Set fn = myEFrag("SELECT * FROM forminhaltn ORDER BY pat_id, Form_ID, FeldVW, ZeitPunkt, FeldNr",,myA)
' Do While Not fn.EOF
'  If fn!Pat_id <> altpat_id Or fn!Zeitpunkt <> altZeitpunkt Or altform_id <> fn!Form_ID Then
'   fk.AddNew
'   fk!FID = fn!FID
'   fk!Pat_id = fn!Pat_id
'   fk!Form_ID = fn!Form_ID
'   fk!Form_AbkVW = fn!Form_AbkVW
'   fk!Zeitpunkt = fn!Zeitpunkt
'   fk!absPos = fn!absPos
'   fk!AktZeit = fn!AktZeit
'   fk!StByte = fn!StByte
'   fk.Update
'   altpat_id = fn!Pat_id
'   altZeitpunkt = fn!Zeitpunkt
'   altform_id = fn!Form_ID
'   nr = 1
'  Else
'   nr = nr + 1
'  End If
'  ff.AddNew
'  ff!FeldNr = fn!FeldNr
'  ff!FeldVW = fn!FeldVW
'  ff!FeldInhVW = fn!FeldInhVW
'  ff!Foid = fk!Foid
'  ff!nr = nr
'  ff.Update
'  fn.MoveNext
' Loop
'' call foreignno
'' Call myEFrag("INSERT INTO forminh SELECT * FROM forminhv;")
'' call foreignyes
'End Function ' formInhMach

Static Function REGEXP_INSTR%(tStr$, muster$)
 Dim rEx As New RegExp
 Dim regMC
 rEx.Pattern = muster
 rEx.IgnoreCase = True
 If rEx.test(tStr) Then
  Set regMC = rEx.Execute(tStr)
  REGEXP_INSTR = regMC(0).FirstIndex + 1 ' wenn das Muster gleich am Anfang kommt, geht's sonst nicht
 Else
  REGEXP_INSTR = 0
 End If ' rEx.test(tstr) Then
End Function ' REGEXP_INSTR

Static Function REGEXP_SUBSTR$(tStr$, muster$)
 Dim rEx As New RegExp
 Dim regMC
 rEx.Pattern = muster
 rEx.IgnoreCase = True
 If rEx.test(tStr) Then
  Set regMC = rEx.Execute(tStr)
  REGEXP_SUBSTR = regMC(0).Value
 Else
  REGEXP_SUBSTR = ""
 End If ' rEx.test(tstr) Then
End Function ' REGEXP_SUBSTR

Function std1(s$) As Date
' edat=STR_TO_DATE(REPLACE(umwname,':','.'),'%d.%m.%Y %H%i%s')
 Dim ns$, DS$, p1%
 ns = REPLACE$(s, ":", ".")
 p1 = InStr(ns, " ")
 DS = Left$(ns, p1) & Mid$(ns, p1 + 1, 2) & ":" & Mid$(ns, p1 + 3, 2) & IIf(Len(Mid$(ns, p1)) > 5, ":", "") & Mid$(ns, p1 + 5, 2)
 On Error Resume Next
 std1 = CDate(DS)
End Function ' std1(s$) As Date

Function std2(s$) As Date
 Dim DS$
' edat=STR_TO_DATE(umwname,'%Y%m%d_%H%i%s')
 DS = Mid$(s, 7, 2) & "." & Mid$(s, 5, 2) & "." & Left$(s, 4) & " " & Mid$(s, 10, 2) & ":" & Mid$(s, 12, 2) & IIf(Len(s) > 13, ":", "") & Mid(s, 14, 2)
 On Error Resume Next
 std2 = CDate(DS)
End Function ' std1(s$) As Date

Function std3(s$) As Date
 Dim ns$, DS$, p1%
 Dim p2%, p3%
' edat=STR_TO_DATE(REPLACE(REPLACE(umwname,':','.'),'-','.'),'%d.%m.%Y %H.%i.%s')
 ns = REPLACE$(REPLACE$(s, ":", "."), "-", ".")
 p1 = InStr(ns, " ")
 If p1 = 0 Then
  DS = ns
 Else
  p2 = InStr(p1, ns, ".")
  DS = Left$(ns, p1) & Mid$(ns, p1 + 1, IIf(p2, MIN(2, p2 - p1 - 1), 2))
  If p2 Then
   p3 = InStr(p2 + 1, ns, ".")
   DS = DS & ":" & Mid$(ns, p2 + 1, IIf(p3, MIN(2, p3 - p2 - 1), 2))
   If p3 Then
    DS = DS & ":" & Mid$(ns, p3 + 1, 2)
   End If
  End If
'  DS = Left$(ns, p1) & Mid$(ns, p1 + 1, MIN(2, p2 - p1 - 1)) & IIf(p2, ":" & Mid$(ns, p2 + 1, MIN(2, p3 - p2 - 1)) & IIf(p3, ":" & Mid$(ns, p3 + 1), ""), "")
'  DS = Left$(ns, p1)
 End If
 On Error Resume Next
 std3 = CDate(DS)
End Function ' std1(s$) As Date

Function std4(s$, dokd As Date) As Date
 Dim ns$, DS$
' edat=STR_TO_DATE(CONCAT(REPLACE(umwname,':','.'),YEAR(dokd)),'%d.%m.%Y');
' IF edat>dokd THEN SET edat=edat-INTERVAL 1 YEAR; END IF;
 ns = REPLACE$(s, ":", ".") & Year(dokd)
 On Error Resume Next
 std4 = CDate(ns)
 If std4 > dokd Then std4 = DateAdd("yyyy", -1, std4)
End Function ' std1(s$) As Date


' notwendig für diese Funktion: Verweis auf Microsoft VBScript Regular Expressions: c:\windows\syswow64\vbscript.dll\3
' aufgerufen in doLies (2x)
' zu koordinieren mit `quelldat` in doViewserstellen( in imporiere
Function doQuelldatum(ByVal name$, Optional dokd As Date) As Date
 Dim rEx As New RegExp
 Dim regMC
 Dim posf%, runde%, p1%, art%
 Const Flb$ = "Fremdlabor"
 Const m1$ = "(31\.((0|)[13578]|1[02])|30\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\.((0|)[0-9]|1[0-2]))\.(19|20|)([0-9]{2}) ([0-9]{6}|[0-9]{4})"
 Const m2$ = "(19|20)[0-9]{2}[01][0-9][0-3][0-9]_[0-2][0-9][0-6][0-9]{3}"
 Const m3$ = "(31\.((0|)[13578]|1[02])|30\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\.((0|)[0-9]|1[0-2]))\.(19|20|)([0-9]{2})( ([0-9]|[0-2][0-9])[.:][0-9]{2}([.:][0-9]{2}|)|)"
 Const m4$ = "(31\.((0|)[13578]|1[02])|30\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\.((0|)[0-9]|1[0-2]))\."
 Dim flname$, testname$, umwname$
 Dim eDat As Date
 posf = InStr(1, name, Flb, vbTextCompare)
 If posf > 0 Then
  flname = Mid$(name, InStrRev(name, Flb) + Len(Flb)) ' SUBSTRING_INDEX(name,Flb,-1)
  p1 = REGEXP_INSTR(flname, "BZ|RR|Medikamentenplan")
  If p1 <> 0 Then flname = Left(flname, p1 - 1)
 Else
  flname = name
 End If ' posf>0
 Do ' na
  art = 0
  testname = flname
  Do ' r1
   p1 = REGEXP_INSTR(testname, m1)
   If p1 <> 0 Then
    art = 1: umwname = REGEXP_SUBSTR(Mid$(testname, p1), m1)
   Else
    p1 = REGEXP_INSTR(testname, m2)
    If p1 <> 0 Then
     art = 2: umwname = REGEXP_SUBSTR(Mid$(testname, p1), m2)
    Else
     p1 = REGEXP_INSTR(testname, m3)
     If p1 <> 0 Then
      art = 3: umwname = REGEXP_SUBSTR(Mid$(testname, p1), m3)
     Else
      p1 = REGEXP_INSTR(testname, m4)
      If p1 <> 0 Then
       art = 4: umwname = REGEXP_SUBSTR(Mid$(testname, p1), m4)
      End If
     End If
    End If
   End If
   If p1 = 0 Then Exit Do
   flname = Mid$(testname, p1)
   testname = Mid$(flname, Len(umwname) + 1)
   If REGEXP_INSTR(testname, "[0-9]") = 0 Then Exit Do
  Loop ' r1
  Select Case art
   Case 1: eDat = std1(umwname)
   Case 2: eDat = std2(umwname)
   Case 3: eDat = std3(umwname)
   Case 4: eDat = std4(umwname, dokd)
   Case Else: eDat = 0
  End Select
  If runde = 1 Or posf = 0 Or eDat <> 0 Then Exit Do
  flname = name
  runde = runde + 1
 Loop ' na
 If eDat = 0 Or eDat > Now() Or eDat < #1/1/1985# Then eDat = CDate("30.12.1899")
 doQuelldatum = eDat
End Function ' doQuelldatum

#If zutesten Then
Public Function testqd()
 Dim rs As New ADODB.Recordset, Zahl&
 Lese.ProgStart
 rs.Open "SELECT quelldat(b.name,b.DokAenD) myd, b.* FROM briefe b", DBCn, adOpenStatic, adLockReadOnly
 DoEvents
 Do While Not rs.EOF
  Zahl = Zahl + 1
  If doQuelldatum(rs!name, rs!DokAenD) <> rs!myd Then
   Debug.Print Zahl, doQuelldatum(rs!name, rs!DokAenD), rs!myd
'  Else
'   Debug.Print Zahl, rs!myd, "stimmt!"
  End If
  rs.Move 1
 Loop
 Debug.Print "Fertig"
End Function ' testqd()
#End If

#If False Then
Function doQuelldatumAlt(ByVal DokName$, Optional dokdat As Date) As Date
 Dim Spli$(), i&, aru%, buch$, doumw$, auffuell%, auffStand%, testname$, p1&, p2&
 Dim doumwD As Date
 Dim gesdat$
 On Error GoTo fehler
' DokName = LCase$(DokName)
 p1 = InStr(1, DokName, "fremdlabor", vbTextCompare)
 If p1 <> 0 Then
  p2 = InStr(1, Mid$(DokName, p1), "bz", vbTextCompare)
  If p2 = 0 Then p2 = InStr(1, Mid$(DokName, p1), "rr", vbTextCompare)
  If p2 = 0 Then p2 = InStr(1, Mid$(DokName, p1), "medikamentenplan", vbTextCompare)
  If p2 <> 0 Then testname = Left$(Mid(DokName, p1), p2 - 1)
 End If
 If testname = "" Then testname = DokName
 For aru = 1 To IIf(p2 <> 0, 2, 1)
  If aru = 2 Then testname = DokName
  If dokdat = 0 Then dokdat = Now()
  Spli = Split(REPLACE$(REPLACE$(REPLACE$(REPLACE$(testname, ".jpg", ""), ".tif", ""), ".pdf", ""), ",", " "))
  For i = UBound(Spli) To 0 Step -1
'   gesdat = Spli(i - 1)
   If i > 0 Then
    If InStr(Spli(i - 1), "-") Then Spli(i - 1) = Mid(Spli(i - 1), InStr(Spli(i - 1), "-") + 1)
    Spli(i - 1) = REPLACE$(Spli(i - 1), "-", ".") ' 25.1.-6-2.23
    If i > 0 Then If IsDate(Spli(i - 1)) And InStrB(Spli(i - 1), ".") <> 0 Then If CDate(Spli(i - 1)) <= Now() Then i = i - 1 ' 2.4.23
   End If
   doumw = Spli(i)
   If IsDate(doumw) And InStrB(doumw, ".") <> 0 Then ' korrigiert 11.7.09
    If InStr(Mid$(doumw, InStr(doumw, ".") + 1), ".") = 0 Then
     gesdat = doumw & "." & Year(dokdat)
    ElseIf Right$(doumw, 1) = "." Then
     gesdat = doumw & Year(dokdat)
    Else
     gesdat = doumw
    End If
    If IsDate(gesdat) Then doumwD = CDate(gesdat) Else doumwD = CDate(doumw)
    If doumwD > dokdat Then doumwD = DateAdd("y", -1, doumwD)
    If doumwD <= Now Then ' 2.4.24
'    rdo.Edit
     buch = ""
     If UBound(Spli) > i Then
      If IsNumeric(Spli(i + 1)) And Len(Spli(i + 1)) = 6 Then
       buch = doumw & " " & Format(Spli(i + 1), "00:00:00")
      ElseIf Spli(i + 1) Like "##[.:]##[.:]##" Then
       buch = doumw & " " & REPLACE$(Spli(i + 1), ".", ":")
      End If
     End If
     If IsDate(buch) Then doQuelldatumAlt = CDate(buch) Else doQuelldatumAlt = doumwD
 '    rdo.Update
     Exit For
    End If
   ElseIf doumw Like "*########_######.png" Then
    doumw = Left$(Right$(doumw, 19), 15)
    doQuelldatumAlt = Mid$(doumw, 7, 2) & "." & Mid$(doumw, 5, 2) & "." & Left$(doumw, 4) & " " & Mid$(doumw, 10, 2) & ":" & Mid(doumw, 12, 2) & ":" & Mid(doumw, 14, 2) ' Format(doumw, "YYYYMMDD_HHMMSS")
   End If ' IsDate(Spli(i)) And InStrB(Spli(i), ".") <> 0
  Next i
  If doQuelldatumAlt <> 0 Then Exit For
 Next aru
 If doQuelldatumAlt = 0 Then doQuelldatumAlt = GetDatumAusString(testname)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doQuelldatum/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doQuelldatum
#End If

#If False Then
Function doFLQuelldatum(DokName$, Optional dokdat As Date) As Date
 Dim posfl%, p0%, ndok$
 Dim Spli$(), i&, buch$, cont As String * 10, auffuell%, auffStand%
 Const Fl$ = "Fremdlabor"
 posfl = InStr(1, DokName, Fl, vbTextCompare)
 If posfl Then
  p0 = posfl + Len(Fl)
  ndok = REPLACE$(Mid$(DokName, p0), ",", " ")
  Spli = Split(ndok)
  For i = 0 To UBound(Spli)
   If IsDate(Spli(i)) And InStrB(Spli(i), ".") <> 0 Then ' korrigiert 11.7.09
'    rdo.Edit
    If UBound(Spli) > i Then
     If IsNumeric(Spli(i + 1)) And Len(Spli(i + 1)) = 6 Then
      buch = Spli(i) & " " & Format(Spli(i + 1), "00:00:00")
      If IsDate(buch) Then
       doFLQuelldatum = CDate(buch)
      Else
       doFLQuelldatum = CDate(Spli(i))
      End If
     Else
      doFLQuelldatum = CDate(Spli(i))
     End If
    Else
     doFLQuelldatum = CDate(Spli(i))
    End If
'    rdo.Update
    Exit For
   End If
  Next i
  If doFLQuelldatum = 0 Then doFLQuelldatum = GetDatumAusString(DokName)
  ReDim Spli(0)
 Else
  doFLQuelldatum = doQuelldatum(DokName, dokdat)
 End If
End Function ' doFLQuelldatum
#End If

Public Function GetDatumAusString(Stri$) As Date
 Dim PPos%() ' Punkt-Positionen
 Dim LLen%(), RLen%() ' Ziffern-Zahl links von jeder Punktposition
 Dim obLP%, obRP% ' ob noch Ziffernzählung nach linker / rechter Punktposition aktiv
 Dim i%, DatStri$
 On Error GoTo fehler
  ReDim PPos(0)
  ReDim LLen(0)
  ReDim RLen(0)
  obLP = 0
  obRP = 0
   For i = Len(Stri) To 1 Step -1
    If Mid$(Stri, i, 1) = "." Then
     PPos(UBound(PPos)) = i
     ReDim Preserve PPos(UBound(PPos) + 1)
     ReDim Preserve LLen(UBound(LLen) + 1)
     ReDim Preserve RLen(UBound(RLen) + 1)
     obLP = True
    ElseIf InStrB("0123456789", Mid$(Stri, i, 1)) <> 0 Then
     If obLP Then LLen(UBound(PPos) - 1) = LLen(UBound(PPos) - 1) + 1
     RLen(UBound(PPos)) = RLen(UBound(PPos)) + 1
    Else
     obLP = 0
     RLen(UBound(PPos)) = 0
    End If
   Next i
  ' Folge zur Definition: 1) 2-4 Ziffern, ".", 1-2 Ziffern, ".", 1-2 Ziffern
   For i = 0 To UBound(LLen) - 1
    If (RLen(i) = 2 Or RLen(i) = 4) And (LLen(i) = 2 Or LLen(i) = 1) And (LLen(i + 1) = 2 Or LLen(i + 1) = 1) And PPos(i) - PPos(i + 1) = LLen(i) + 1 Then
     DatStri = Mid$(Stri, PPos(i + 1) - LLen(i + 1), RLen(i) + LLen(i) + LLen(i + 1) + 2)
     If IsDate(DatStri) Then
      GetDatumAusString = CDate(DatStri)
      If GetDatumAusString > Now() - 50 * 365 And GetDatumAusString <= Now Then
       Exit For
      End If
      GetDatumAusString = 0
     End If
    End If
   Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetDatumAusString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetDatumAusString$

' 28.10.18: nirgends aufgerufen
Function LöschDateiEintrag(DatID&) ' 3.2.07: Erstellt, auch schon verwendet
 Dim rAF&
 Dim rs As ADODB.Recordset
 Set rs = myEFrag("UPDATE `laborneu` SET refnr = null WHERE refnr IN (SELECT refnr FROM `laborxus` LEFT JOIN laborxeingel ON laborxus.datid = laborxeingel.datid WHERE laborxus.datid= " & DatID & ")", rAF)
 Debug.Print rAF
 Set rs = myEFrag("DELETE FROM `laborxleist` WHERE refnr IN (SELECT refnr FROM `laborxus` LEFT JOIN laborxeingel ON laborxus.datid = laborxeingel.datid WHERE laborxus.datid= " & DatID & ")", rAF)
 Debug.Print rAF
 Call ForeignNo0
 Call ForeignNo1
 Set rs = myEFrag("DELETE FROM `laborxwert` WHERE refnr IN (SELECT refnr FROM `laborxus` LEFT JOIN laborxeingel ON laborxus.datid = laborxeingel.datid WHERE laborxus.datid= " & DatID & ")", rAF)
' Call ForeignYes
 Debug.Print rAF
 Set rs = myEFrag("DELETE FROM `laborxbakt` WHERE refnr IN (SELECT refnr FROM `laborxus` LEFT JOIN `laborxeingel` ON `laborxus`.datid = `laborxeingel`.datid WHERE `laborxus`.datid= " & DatID & ")", rAF)
 Debug.Print rAF
' Call ForeignNo
 Set rs = myEFrag("DELETE FROM `laborxsaetze` WHERE satzid IN (SELECT satzid FROM `laborxus` LEFT JOIN `laborxeingel` ON `laborxus`.datid = `laborxeingel`.datid WHERE `laborxus`.datid= " & DatID & ")", rAF)
 Call ForeignYes0
 Call ForeignYes1
 Debug.Print rAF
 Set rs = myEFrag("DELETE FROM `laborxus` WHERE datid = " & DatID, rAF)
 Debug.Print rAF
 Set rs = myEFrag("DELETE FROM `laborxeingel` WHERE datid = " & DatID, rAF)
 Debug.Print rAF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LöschDateiEintrag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LöschDateiEintrag

'#If False THEN
'Function LöschRefNr()
  'On Error GoTo fehler
  'Call dtbInit
'#If True THEN
' myEFrag ("UPDATE `laborneu` SET refnr = null",,Dtb)
'#Else
  'Dim rs AS DAO.Recordset
  'Set rs = TabÖff("laborneu")
  'Do While Not rs.EOF
'   IF NOT ISNULL(rs!RefNr) THEN
    'rs.Edit
    'rs!RefNr = Null
    'rs.Update
   'END IF
'   rs.MoveNext
  'Loop
'#END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LöschRefNr/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' LöschRefNr
'#END IF
' 28.10.18: nirgends aufgerufen
Function testLöschAb(ab&, frm As Lese) ' s.a. LöschDateiEintrag und LabLösch
' Call frm.ConstrFestleg(0, frm)
 Call acon(quelleT)
' DBcnOpen ConStr
 Call myEFrag("DELETE FROM `laborxleist` WHERE refnr IN (SELECT refnr FROM `laborxus` WHERE datid>=" & ab & ")")
 Call myEFrag("DELETE FROM `laborxbakt` WHERE refnr IN (SELECT refnr FROM `laborxus` WHERE datid>=" & ab & ")")
 Call ForeignNo0
 Call ForeignNo1
 Call myEFrag("DELETE FROM `laborxwert` WHERE refnr IN (SELECT refnr FROM `laborxus` WHERE datid>=" & ab & ")")
 Call ForeignYes0
 Call ForeignYes1
 Call myEFrag("DELETE FROM `laborxus` WHERE datid >= " & ab)
 Call myEFrag("DELETE FROM `laborxeingel` WHERE datid >= " & ab)
End Function ' testLöschab

Function EmailsImport(EmDatei$, frm As Lese)
 Dim con As New ADODB.Connection  ' Connection
 Dim rNa As New ADODB.Recordset
 Dim rEx As New ADODB.Recordset
 Dim rX As New ADOX.Catalog
' Const EmDatei$ = pverz & "Patientenübergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
 On Error GoTo fehler
 If LenB(DBCn) = 0 Or DBCn = "" Then
  Call acon(quelleT) 'Call frm.ConstrFestleg(0, frm)
 End If
 If EmDatei <> "" Then
 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=2"";Data Source=" & EmDatei & ";" ' TABLE=Adressen$"
 Dim runde%, i%, pFeld$, eFeld$, obAnfang%, pNr&, pRoh$, email$, ka%, ke%, rAF&
  rX.ActiveConnection = con
  rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", con ' Hier Excel, nicht lies.obmysql = 0!
  Do While Not rEx.EOF
'  Debug.Print runde
   If obAnfang Then
    pRoh = rEx.Fields(pFeld)
    ka = InStr(pRoh, "(")
    ke = InStr(pRoh, ")")
    If ka > 0 And ke > 0 Then
     pNr = CLng(Mid$(pRoh, ka + 1, ke - ka - 1))
     Call myEFrag("UPDATE `namen` SET email = '" & rEx.Fields(eFeld) & "' WHERE pat_id = " & pNr, rAF)
' wenn Name noch nicht da
'     IF rAf <> 1 THEN Err.Raise 999, , "Fehler beim Einfügen der Email-Adresse '" & rEx.Fields(eFeld) & "' zum Patienten Nr. " & pNr
    End If
   ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) Then
    For i = 0 To rEx.Fields.COUNT - 1
     If rEx.Fields(i) = "Patient" Then
      pFeld = rEx.Fields(i).name
     ElseIf rEx.Fields(i) = "Email" Then
      eFeld = rEx.Fields(i).name
     End If
    Next i
    obAnfang = True
   End If
  runde = runde + 1
  rEx.MoveNext
  Loop
 End If
 On Error Resume Next
 rEx.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EmailsImport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Emails-Import

Function AnPack()
 Const TName$ = "anamnesebogen"
 Dim maxFuell&, DBlen&, SpName$, ZLen$
 Dim rsc As New ADODB.Recordset, rslen As New ADODB.Recordset
 On Error GoTo fehler
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, TName, Empty))
 Do
  SpName = rsc!COLUMN_NAME
  Debug.Print SpName, rsc!data_type, DBlen
  Select Case rsc!data_type
   Case 130
    DBlen = rsc!character_maximum_length
    If DBlen > 0 Then
     If Not rslen Is Nothing Then If rslen.State = 1 Then rslen.Close
     myFrag rslen, "SELECT MAX(" & sqllen & "(`" & rsc!COLUMN_NAME & "`)) AS mlen FROM " & TName & ";"
     ZLen = rslen!mlen
     If ZLen < DBlen Then
      Debug.Print SpName & ":" & DBlen & " -> " & ZLen
      rslen.Close
      On Error Resume Next
      Call myEFrag("ALTER TABLE `" & TName & "` " & sqlALTER & " Column " & " `" & rsc!COLUMN_NAME & "` " & sqlText & "(" & ZLen & ") " & " DEFAULT " & IIf(IsNull(rsc!Default), "NULL", rsc!Default) & IIf(IsNull(rsc!collation), vNS, " COLLATE " & rsc!collation) & " COMMENT '" & rsc!Comment & "'")
      Debug.Print "Erfolg: " & Err.Number & " " & Err.Description
      On Error GoTo fehler
     End If
    End If
  End Select
  rsc.Move 1
  If rsc.EOF Then Exit Do
 Loop
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnPack/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' AnPack

'Function SpModAlt%(SpVal$, TName, rs0 As Adodb.Recordset) ' Spalte modifizieren
'Dim ZCStr$, keinetrans%, SpName$, maxL%
'Dim rsc As New ADODB.Recordset
'On Error GoTo fehler
'SpName = rs0!COLUMN_NAME
'If ISNULL(rs0!character_maximum_length) THEN
' maxL = 32000
'Else
' maxL = IIf(rs0!character_maximum_length > 255, 5, rs0!character_maximum_length)
'END IF
''If rs0!data_type = 129 THEN obMemo = True
'If Len(SpVal) > maxL AND maxL > 0 THEN ' longtext
' IF obMitAlterTab = 0 THEN
'  Lese.Ausgeb "Feldinhaltverkürzung Tabelle '" & TName & "' Feld '" & SpName & "' (Länge:" & Len(SpVal) & "->" & maxL & ");" & SpVal & " -> " & LEFT(SpVal, maxL), True
'  SpVal = LEFT(SpVal, maxL)
' Else
'  ZCStr = DBCnS ' DBCn.ConnectionString
'  ON Error Resume Next
'  IF obTrans <> 0 THEN DBCn.CommitTrans: obTrans = 0
'  keinetrans = Err.Number
'  ON Error GoTo fehler
'  DBCn.Close
''  Call CnOpen(False, ZCStr)
'  Call acon(quelleT)
'  IF lies.obMySQL THEN Call myEFrag("USE `" & lies.MyDB & "`")
'  Call ForeignNo0
'  Call ForeignNo1
'  IF DBCn.State = 0 THEN
'   DBCnOpen
'  END IF
'  IF LenB(sqlText) = 0 THEN Zinit (lies.obMySQL)
'  IF Len(SpVal) > maxL THEN
'   IF Not lies.obMySQL AND Len(SpVal) > 255 THEN
'    Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & "MEMO")
'   Else
'    ON Error Resume Next
'nochmal:
'    Err.Clear
'    IF lies.obMySQL THEN
'     SET rsc = Nothing
'     myFrag rsc, "SHOW FULL COLUMNS FROM `" & TName & "` WHERE field = '" & SpName & "'"
'     Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & sqlText & "(" & Len(SpVal) & ") " & IIf(ISNULL(rsc!collation), vNS, " COLLATE " & rsc!collation) & " default " & IIf(ISNULL(rsc!Default), " null", rsc!Default) & " comment '" & rsc!Comment & "'", rAF)
'    Else
'     Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & sqlText & "(" & Len(SpVal) & ")", rAF)
'    END IF
'    Dim FNr&
'    FNr = Err.Number
'    Dim FText$
'    FText = Err.Description & " " & vbCrLf & Err.LastDllError
'    ON Error GoTo fehler
'   END IF
'   Dim ausgTxt$
'   IF FNr = 0 THEN
'    ausgTxt = "Feldvergrößerung Tabelle '" & TName & "' Feld '" & SpName & "';" & maxL & " -> " & Len(SpVal) & "; wg. Inhalt: '" & SpVal & "'"
'   Else
'    ausgTxt = FText & vbCrLf & "Feldvergrößerung gescheitert Tabelle '" & TName & "' Feld '" & SpName & "';" & maxL & " -> " & Len(SpVal) & "; wg. Inhalt: '" & SpVal & "', verkürze auf: '" & LEFT(SpVal, maxL) & "'"
'    SpVal = LEFT(SpVal, maxL)
'   END IF
'   Lese.Ausgeb ausgTxt, True
'   Open Lese.snst.DebugDatei For Append AS #399
'   Print #399, ausgTxt
'   Close #399
'   SpModAlt = True
'  END IF
'  DoEvents
'  Call ForeignYes0
'  Call ForeignYes1
'  IF keinetrans = 0 THEN
'   IF DBCn.State = 0 THEN
'    DBCnOpen
'   END IF
'   DBCn.BeginTrans
'   obTrans = 1
'  END IF
' END IF
'Else
'' MsgBox "Erst mal schauen, ob er hierherkommt"
'' Err.Raise 999, , "Erst mal schauen, ob er hierherkommt"
'' IF obMitAlterTab = 0 THEN
''  Lese.Ausgeb "Feldudt.mwandlung Tabelle '" & Tname & "' Feld '" & SpName & "' (Text mit Länge:" & Len(SpVal) & "-> Memo);" ,true
'' ElseIf Not lies.obmysql THEN ' Vermutlich tritt Phänomen nur bei Access auf
''  ZCStr = DBCn.ConnectionString
''  ON Error Resume Next
''  DBCn.CommitTrans
''  keineTrans = Err.Number
''  ON Error GoTo fehler
''  DBCn.Close
''  DBcnOpen ZCStr
''  IF lies.obmysql THEN Call myEFrag("use " & myDB)
''  call foreignno
''  Call myEFrag("ALTER TABLE " & "`" & Tname & "`" & " " & IIf(lies.obmysql, "MODIFY", "ALTER") & " Column " & "`" & SpName & "`" & " " & "MEMO")
''   ausgTxt = "Feldudt.mwandlung Tabelle '" & Tname & "' Feld '" & SpName & "';" & maxL & " -> Memo; wg. Inhalt: " & SpVal
''   Lese.Ausgeb ausgTxt ,true
''   Open Lese.snst.DebugDatei For Append AS #399
''   Print #399, ausgTxt
''   Close #399
''   SpMod = True
'' END IF
'END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SpModAlt/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' SpModAlt

Function LaborParameter(rlAb As ADODB.Recordset)
Dim raLP As New ADODB.Recordset, rlaL As New ADODB.Recordset
Dim Abkü$, Einheit$, Langtext$
Dim LTName$
Dim obltvw%
On Error Resume Next
LTName = rlAb.Fields("langtextvw").name
obltvw = IIf(Err.Number = 0, -1, 0)
On Error GoTo fehler
Abkü = IIf(IsNull(rlAb!Abkü), vNS, Trim$(rlAb!Abkü)) 'If ISNULL(rLab!Abkü) THEN rLab!Abkü = vns
Einheit = Zeichensatz(IIf(IsNull(rlAb!Einheit), vNS, Trim$(rlAb!Einheit))) ' IF ISNULL(rLab!Einheit) THEN rLab!Einheit = vns
If Einheit = "" Then Einheit = "kA" ' 2.10.: ist so in Turbomed, => auch für die Direktübernahme so hindrehen
Langtext = vNS
If obltvw Then
' SET rlaL = TabÖff("LaborLangtext", "LangtextVW")
 Set rlaL = Nothing
' rlaL.Open "laborlangtext", DBCn, adOpenDynamic, adLockReadOnly
 myFrag rlaL, "SELECT 0 FROM laborlangtext LIMIT 1"
 If Not rlaL Is Nothing And Not IsNull(rlAb!LangtextVW) Then
'  rlaL.Seek "=", rlAb!LangtextVW
  Set rlaL = Nothing
'  rlaL.Open "SELECT Langtext FROM `laborlangtext` WHERE langtextvw = " & rlAb!LangtextVW, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rlaL, "SELECT Langtext FROM `laborlangtext` WHERE langtextvw = " & rlAb!LangtextVW
  If Not rlaL.EOF Then
   Langtext = rlaL!Langtext
  End If
 End If
Else
 Langtext = IIf(IsNull(rlAb!Langtext), vNS, rlAb!Langtext)
End If
'Set raLP = TabÖff("LaborParameter", "Abkü")
Set raLP = Nothing
' Call raLP.Open("laborparameter", DBCn, adOpenDynamic, adLockReadOnly)
If Abkü <> "" Then
' rLP.Seek "=", Abkü, Einheit
 myFrag raLP, "SELECT abkü,einheit,langtext,aktzeit FROM laborparameter WHERE abkü = '" & Abkü & "' LIMIT 1"
 If raLP.BOF Then
  raLP.AddNew
 End If
 If raLP.EditMode = 2 Then
  If obÜP Then Print #322, "Neuer Laborparameter: " + Abkü + " `" + Einheit + "`"
  raLP!Abkü = Abkü
  raLP!Einheit = Einheit
 ElseIf (raLP!Langtext <> Langtext) Then 'Or (rLP!Kommentar <> rLab!Kommentar AND NOT ISNULL(rLab!Kommentar)) THEN
'  raLP.Edit
 Else
'  rLPbm = rLP.Bookmark
  GoTo Fertig:
 End If
 If raLP.EditMode = 1 Or raLP.EditMode = 2 Then
  If Not IsNull(Langtext) And Langtext <> "" Then
   raLP!Langtext = Langtext
  End If
'  rLP!Kommentar = rLab!Kommentar
  raLP!AktZeit = Now
  raLP.Update
'  rLPbm = rLP.LastModified
 End If
End If
 GoTo Fertig
test:
 MsgBox "Hier Fehler bei Einheit"
 Resume
Fertig:
 raLP.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborParameter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborParameter(rLab AS dao.recordset)

#If False Then
Function nulltest()
 Lese.ProgStart
 Dim rs As ADODB.Recordset
 myFrag rs, "SELECT NULL erg"
 Debug.Print "nix " & IIf(IsNull(rs!erg), "null", rs!erg)
End Function
#End If
