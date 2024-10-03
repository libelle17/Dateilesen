Attribute VB_Name = "ImportiereLabor"
Option Explicit
Const HADBName$ = "haerzte"
Public BezFeh$ ' Dateinamen für Beziehungsfehler beim Import
Public hVerz$, aVerz$, QmdB$, pügVerz$, üVerz$, LabTransPfad$
Public VorByte&, AktByte&, rEzäBeg As Date
Public rUn1&, rFo1& ' , rFi1& ' erster neuer Datensatz in formulare / forminhaltform_abk
Public obMitAlterTab% ' ALTER TABLE bei zu kurzen Feldern ' in MySQL
Dim rAf& ' records affected
 Type rlwType2 ' Ergänzung für laborpn (Tabelle für externe Labordateien)
  e As laborxwert
  Eingang As Date
  geschlecht As String
  NormU As String
  NormO As String
  Normbereich As String
  uid As Long
 End Type ' rlwType2

' 4.9.06: Wird in NachArbeit aufgerufen
Public Function LaborDirektImport(frm As Lese, absPos&, SammelInsert%, BezfSpei%, LPBeisp$, Optional oballe% = 0)
 Dim abPos&, i&, j&, ii&, erg$
 Const debugohneSpeichern% = False
' Const neuVerz$ = pverz & "Biowin\Backup\neu\"
 Const ZS As Boolean = True
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Dim repl$
 Dim fI As Files, fld As Folder, f1 As File, neuPfad$, keimz%
 Dim RefNr&
 Dim obB% ' obBakt: -1 = ja
 Dim txt$, SatzlängenArt%, obSonderSatz%
 Dim DebugZähler&
 Dim aktDateiNr&, DateiZahl&, obDateiNeu%
 Dim merkDatID&
 Dim rLw2() As rlwType2
 T1a = Now
 frm.ZeilenBez = "Bytes:"
 frm.Ausgeb "Lese Labor-DFÜ-Dateien ein ...", False
 On Error GoTo fehler
 If oballe Then 'Reihenfolge wichtig:
  Call LabLösch(frm)
  Call LöschRefNr
 End If
 If FSO Is Nothing Then Set FSO = New FileSystemObject
' Call VerzPrüf(neuVerz)
' IF 1 = 1 THEN
 Dim SL As New SortierListe, SD As SortierDatei
' IF Not FSO.FolderExists(LPBeisp) THEN Call LaborPfadDialog
 If Right$(LPBeisp, 1) <> "\" Then LPBeisp = LPBeisp & "\"
 If FSO.FolderExists(LPBeisp) Then
  erg = Dir(LPBeisp, vbDirectory)
 Else
  MsgBox LPBeisp & " nicht gefunden, höre auf."
  ProgEnde
 End If
 If LenB(erg) = 0 Then LPBeisp = Lese.LaborPfadDialog(LPBeisp)
' IF Not FSO.FolderExists(LPBeisp) THEN Exit Function
 erg = Dir(LPBeisp, vbDirectory)
 If LenB(erg) = 0 Then Exit Function
 Set fld = FSO.GetFolder(LPBeisp)
 Set fI = fld.Files
 frm.BytesBez = "Dateien:"
 frm.GesBytes = fI.COUNT
 frm.Ausgeb "Untersuche Dateien in Verzeichnis " & fld & " ...", True
 Call ForeignNo0
 Call ForeignNo1
 frm.Ausgeb vbCrLf, False
 For Each f1 In fI
  If f1.name Like "1b*.ld*" Or f1.name Like "X*.LD*" Or f1.name Like "Labor*.dat" Then
    aktDateiNr = aktDateiNr + 1
    frm.Bytes = aktDateiNr
'    repl = f1.Name
'    neuPfad = f1.Path
    obDateiNeu = 0
    For i = 1 To 2
'     Set rs = myEFrag("SELECT -fertig AS j_fertig, l.* FROM laborxeingel l WHERE name = '" & f1.name & "' AND pfad = '" & doUmwfSQL(f1.path, Lese.obMySQL) & "'")
     myFrag rs, "SELECT -fertig j_fertig, l.* FROM laborxeingel l WHERE name = '" & f1.name & "' AND pfad = '" & doUmwfSQL(f1.path, Lese.obMySQL) & "'"
     If rs.BOF Then
      obDateiNeu = True
     ElseIf rs!j_fertig = 0 Then
      obDateiNeu = 1  ' = schon da, aber nicht fertig
      merkDatID = rs!DatID
     Else
      If debugohneSpeichern Then
       merkDatID = rs!DatID ' andernfalls wird es später eingestellt
      End If
      Exit For
     End If
    Next i
    If debugohneSpeichern Then
     If obDateiNeu = 0 Then obDateiNeu = -1
    End If
    frm.Sekunden = Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    If obDateiNeu <> 0 Then
      frm.Ausgeb ".", False
      If aktDateiNr Mod 100 = 0 Then
       frm.Ausgeb vbCrLf, True
      End If
      DoEvents
      Set SD = New SortierDatei
      If merkDatID <> 0 Then ' bei Debuggen oder bei nicht fertiger Datei
       SD.DatID = merkDatID
       merkDatID = 0
      End If
      Set SD.File = f1
'      SD.DatID = rs!DatID ' für Debugging
      Call SL.sCAdd(SD)
    End If
  End If
  DoEvents
  If BrichAb Then
   Exit Function
  End If
 Next f1
 DateiZahl = SL.COUNT
 aktDateiNr = 0
 frm.BytesBez = "Dateien:"
 frm.GesBytes = DateiZahl
' frm.Ausgabe = altAusgabe
 frm.Bytes = 0
#If LaborDirektEinlesen Then ' Projekt -> Eigenschaften -> Erstellen Konstanten für bedingte Kompilierung
Const LDEProt$ = pVerz & "Patientenübergreifendes\Labordirekteingelesen.prot"
On Error Resume Next
Open LDEProt$ For Append As #251
For ii = 1 To SL.COUNT
 Print #251, Now() & ", zubearbeiten: " & SL.Item(ii).File.path
Next ii
Close #251
On Error GoTo fehler
#End If
 For ii = 1 To SL.COUNT
    aktDateiNr = aktDateiNr + 1
    frm.Ausgeb "Lese LaborDatei " + CStr(aktDateiNr) + "/" + CStr(DateiZahl) + " (" + SL.Item(ii).File.path + ") ein ...", True
    If SL.Item(ii).DatID = 0 Then ' wenn kein Debuggenohnespeichern und kein unfertiger Datensatz gefunden wurde
      Dim isql$
      isql = "INSERT INTO `laborxeingel`(name,pfad,zp) VALUES ('" & SL.Item(ii).File.name & "','" & doUmwfSQL(SL.Item(ii).File.path, Lese.obMySQL) & "'," & DatFor_k(Now) & ")"
      InsKorr DBCn, isql, rAf
      isql = "SELECT DatID FROM `laborxeingel` WHERE name = '" & SL.Item(ii).File.name & "' AND pfad = '" & doUmwfSQL(SL.Item(ii).File.path, Lese.obMySQL) & "'" ' *
      Set rs = Nothing
      myFrag rs, isql
      If Not rs.EOF Then
       SL.Item(ii).DatID = rs!DatID
      End If ' Not rs.EOF Then
    End If ' SL.Item(ii).DatID = 0 Then
'   IF f1.Name LIKE "1b*.ld*" OR f1.Name LIKE "Labor*.dat" THEN
'   IF f1.Name LIKE "Labor*.dat" THEN
'    repl = replace$(f1.Name, ".", "")
'    Dim DatID&
'    DatID = rs!DatID
'     Debug.Print SL.Item(ii).File.Path, Len(SL.Item(ii).File.Path)
     Open SL.Item(ii).File.path For Input As #1
     
  Dim obeof%, obAngeh%, obSchluss%, ltxt$
'  Dim rLsAkt As New ADODB.Recordset, rLuAkt As New ADODB.Recordset
  obSchluss = 0
  obeof = EOF(1)
  Do
   If obeof Then Exit Do
   Line Input #1, txt
   If frm.Zeilen = "" Then frm.Zeilen = "0"
   frm.Zeilen = frm.Zeilen + 1
   frm.Zeilen = Format$(frm.Bytes + Len(txt) + 2, "###,###,###,###,###,###,###,##0")
   If frm.Zeilen Mod 10 = 0 Then
    frm.Sekunden = Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    DoEvents
   End If
   obeof = EOF(1)
   If Not Left$(txt, 7) Like "#######" Then 'IsNumeric(LEFT(txt, 7)) THEN
    ltxt = ltxt & " " & txt
    obAngeh = -1
   ElseIf obAngeh Then
    obAngeh = 0
   End If
  
   Do
    If obAngeh Then
     If obeof Then
      obSchluss = True
     Else
      Exit Do
     End If
    End If
    
    If ZS Then ltxt = ZSU1(ltxt)
    If ltxt <> "" Then
      Dim lenge%, Kennung%, Inhalt$
      absPos = absPos + 1
      lenge = Left$(ltxt, 3)
      Kennung = Mid$(ltxt, 4, 4)
      Inhalt = REPLACE$(doUmwfSQL(Mid$(ltxt, 8), Lese.obMySQL), "³", "ü") ' Zeichensatz( ist meist falsch, nur das ü stimmt in Karlsfeld nicht
      DebugZähler = DebugZähler + 1
        Select Case Kennung
         Case 8000
          Select Case Left$(Inhalt, 4)
           Case 8220 'L-Datenpaket-Header (Turbomed)
            ReDim rLs(0)
            ReDim rLu(0)
            ReDim rLo(0)
            ReDim rLw(0)
            ReDim rLw2(0)
            ReDim rLL(0)
            rLs(0).Satzart = Inhalt
            RefNr = 0
            SatzlängenArt = 0
            obSonderSatz = -1
           Case 8221 ' Schluss ist erst bei 9202
            SatzlängenArt = -1
            obSonderSatz = -1
           Case Else ' 8201, 8202,8203
'            IF rLuAkt.State = 1 THEN rLuAkt.Update
            ReDim Preserve rLu(UBound(rLu) + 1)
            RefNr = RefNr - 1 ' negative Pseudo-Refnummern
            rLu(UBound(rLu)).DatID = SL.Item(ii).DatID
            rLs(0).DatID = SL.Item(ii).DatID
            rLu(UBound(rLu)).SatzID = rLs(0).SatzID
            rLu(UBound(rLu)).Satzart = Inhalt
            rLu(UBound(rLu)).RefNr = RefNr
            obSonderSatz = 0
          End Select
         Case 8410
          ReDim Preserve rLw(UBound(rLw) + 1)
          rLw(UBound(rLw)).RefNr = rLu(UBound(rLu)).RefNr
          ReDim Preserve rLw2(UBound(rLw2) + 1)
          rLw2(UBound(rLw2)).e.RefNr = rLu(UBound(rLu)).RefNr
          rLw2(UBound(rLw2)).Eingang = rLu(UBound(rLu)).Eingang
          rLw2(UBound(rLw2)).geschlecht = rLu(UBound(rLu)).geschlecht
          rLw2(UBound(rLw2)).uid = UBound(rLu)
          obB = 0
          rLw(UBound(rLw)).Abkü = Inhalt 'replace$(Inhalt, "€", " ")
          rLw2(UBound(rLw2)).e.Abkü = Inhalt 'replace$(Inhalt, "€", " ")
         Case 8434
          ReDim Preserve rLo(UBound(rLo) + 1)
          rLo(UBound(rLo)).RefNr = rLu(UBound(rLu)).RefNr
          obB = -1
          rLo(UBound(rLo)).Verf = Inhalt
         Case 8100: If obSonderSatz Then If Not SatzlängenArt Then rLs(0).Satzlänge = Inhalt Else rLs(0).SatzlängeSchluss = Inhalt Else rLu(UBound(rLu)).Satzlänge = Inhalt
         Case 8310: rLu(UBound(rLu)).Auftragsnummer = Inhalt
         Case 8311: rLu(UBound(rLu)).Auftragsschlüssel = Inhalt
         Case 8301: rLu(UBound(rLu)).Eingang = BDTtoDate(Inhalt)
         Case 8302: rLu(UBound(rLu)).Berichtsdatum = Inhalt
         Case 8303: rLu(UBound(rLu)).Berichtsdatum = rLu(UBound(rLu)).Berichtsdatum & " " & Inhalt
         Case 8609: rLu(UBound(rLu)).Abrechnungstyp = Inhalt
         Case 8401: rLu(UBound(rLu)).BefArt = Inhalt ' Fertigstellungsgrad
         Case 8615: rLu(UBound(rLu)).Auftraggeber = Inhalt ' LANR
         Case 8403: rLu(UBound(rLu)).GebüOrd = Inhalt
         Case 8405: rLu(UBound(rLu)).Patienteninformation = Inhalt
         Case 8407: rLu(UBound(rLu)).geschlecht = Inhalt ' 1=Mann, 2=Frau, 3=unbek, 4=Knabe, 5=Mädchen, 0=Name fehlt, 9=beide
                    If Not IsNumeric(rLu(UBound(rLu)).geschlecht) Then rLu(UBound(rLu)).geschlecht = 3
         Case 3101:
          rLu(UBound(rLu)).Nachname = Inhalt
          If rLu(UBound(rLu)).Nachname Like "zzz*" Then rLu(UBound(rLu)).Nachname = Mid$(rLu(UBound(rLu)).Nachname, 4)
         Case 3102: rLu(UBound(rLu)).Vorname = Inhalt
         Case 3110: rLu(UBound(rLu)).geschlecht = IIf(Inhalt = "W", 2, IIf(Inhalt = "M", 1, IIf(Inhalt = "X", 3, Inhalt)))
         Case 3103
          Dim glZ%, glZLocker%, glZoL%
'          Dim Bm$, bmLocker$, BmoL$
          Dim erwZ% 'Zahl der erwogenen Pat_ids
          Dim rLuGebDat As Date
          Dim merkPat_ID&
          erwZ = 0
          rLu(UBound(rLu)).GebDat = Inhalt
          If Left$(rLu(UBound(rLu)).GebDat, 4) = "0000" Then rLu(UBound(rLu)).GebDat = "0101" + Right$(rLu(UBound(rLu)).GebDat, 4)
          rLuGebDat = BDTtoDate(rLu(UBound(rLu)).GebDat) ' CDate(format$(rlu(ubound(rlu)).GebDat, "##/##/####"))
          If Not (IsNull(rLu(UBound(rLu)).Nachname) Or rLu(UBound(rLu)).Nachname = "") Or Not (IsNull(rLu(UBound(rLu)).Vorname) Or rLu(UBound(rLu)).Vorname = "") Then
           Dim runde%
           Dim Nachname$
           For runde = 1 To 2
'           SELECT CASE runde
'            Case 1: rNam.Seek "=", rlu(ubound(rlu)).Nachname, rlu(ubound(rlu)).Vorname, rlu(ubound(rlu)).GebDat
'            Case 2: rNam.Seek "=", rlu(ubound(rlu)).Nachname, rlu(ubound(rlu)).Vorname
'           END SELECT
            If rs.State = 1 Then rs.Close
            Select Case runde
             Case 1:
              myFrag rs, "SELECT * FROM `namen` WHERE nachname = '" & rLu(UBound(rLu)).Nachname & "' AND vorname = '" & rLu(UBound(rLu)).Vorname & "' AND gebdat = " & DatFor_k(rLuGebDat)
              If rs.BOF Then rs.Close: myFrag rs, "SELECT * FROM `namen` WHERE nachname = 'zzz" & rLu(UBound(rLu)).Nachname & "' AND vorname = '" & rLu(UBound(rLu)).Vorname & "' AND gebdat = " & DatFor_k(rLuGebDat)
             Case 2: myFrag rs, "SELECT * FROM `namen` WHERE nachname = '" & rLu(UBound(rLu)).Nachname & "' AND vorname = '" & rLu(UBound(rLu)).Vorname & "'"
              If rs.BOF Then rs.Close: myFrag rs, "SELECT * FROM `namen` WHERE nachname = 'zzz" & rLu(UBound(rLu)).Nachname & "' AND vorname = '" & rLu(UBound(rLu)).Vorname & "'"
            End Select
            If Not rs.BOF Then
             Do
              If runde = 1 Or rs!GebDat <> rLuGebDat Then  ' damit in der zweiten Runde nicht dieselben nochmal kommen / Null dürfte nicht vorkommen
               erwZ = erwZ + 1
               Select Case runde
                Case 1
                 rLu(UBound(rLu)).Pat_idErwVNG = rLu(UBound(rLu)).Pat_idErwVNG & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwVNG) Or LenB(rLu(UBound(rLu)).Pat_idErwVNG) = 0, vNS, "/") & CStr(rs!Pat_id)
                Case 2
                 rLu(UBound(rLu)).Pat_idErwVN = rLu(UBound(rLu)).Pat_idErwVN & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwVN) Or LenB(rLu(UBound(rLu)).Pat_idErwVN) = 0, vNS, "/") & CStr(rs!Pat_id)
               End Select
               merkPat_ID = rs!Pat_id
              End If
              rs.Move 1
              If rs.EOF Then Exit Do
              Nachname = rs!Nachname
              If Nachname Like "zzz*" Then Nachname = Mid$(Nachname, 4)
              If Nachname <> rLu(UBound(rLu)).Nachname Or rs!Vorname <> rLu(UBound(rLu)).Vorname Or (runde = 1 And rs!GebDat <> rLuGebDat) Then Exit Do
             Loop
            End If ' not nomatch
           Next runde
          End If ' ISNULL...
          If erwZ = 1 Then
           rLu(UBound(rLu)).Pat_id = merkPat_ID
           rLu(UBound(rLu)).Pat_idUrsp = "E"
          Else
           glZ = 0 ' Zahl der Patienten mit passendem Geburtstag, zu denen ein zeitlich passendes Labor in den anderen Labortabellen gefunden wurde
           glZLocker = 0 ' Zahl der Patienten mit passendem Geburtstag, die zum Zeitpunkt der Laboruntersuchung als schon bei mir behandelt gekennzeichnet waren
           glZoL = 0 ' Zahl der Patienten mit passendem Geburtstag
'           Bm = vns
'           bmLocker = vns
'           BmoL = vns
           If rs.State = 1 Then rs.Close
'           rs.Open "SELECT * FROM `anamnesebogen` WHERE gebdat = " & DatFor_k(rLuGebDat), DBCn, adOpenDynamic, adLockReadOnly
           myFrag rs, "SELECT * FROM `anamnesebogen` WHERE gebdat = " & DatFor_k(rLuGebDat)
           If Not rs.BOF Then
            Do
             glZoL = glZoL + 1
             erwZ = erwZ + 1
             rLu(UBound(rLu)).Pat_idErwG = rLu(UBound(rLu)).Pat_idErwG & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwG) Or LenB(rLu(UBound(rLu)).Pat_idErwG) = 0, vNS, "/") & CStr(rs!Pat_id)
'             BmoL = rsAnam.Bookmark
             If rs!Vorgestellt <= rLu(UBound(rLu)).Eingang + 5 Then ' mit etwas Reserve für Wochenende usw.
'              bmLocker = rsAnam.Bookmark
              glZLocker = glZLocker + 1
              If glZLocker = 1 Then
                merkPat_ID = rs!Pat_id
              End If
              rLu(UBound(rLu)).Pat_idErwGB = rLu(UBound(rLu)).Pat_idErwGB & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwGB) Or LenB(rLu(UBound(rLu)).Pat_idErwGB) = 0, vNS, "/") & rs!Pat_id
              If rs1.State = 1 Then rs1.Close
              myFrag rs1, "SELECT * FROM `laborneu` WHERE pat_id = " & rs!Pat_id & " AND zeitpunkt >= " & DatFor_k(rLu(UBound(rLu)).Eingang - 5) & " AND zeitpunkt <= " & DatFor_k(rLu(UBound(rLu)).Eingang + 15)
              If Not rs1.BOF Then
 '              Bm = rsAnam.Bookmark
               glZ = glZ + 1
               If glZ = 1 Then
                merkPat_ID = rs!Pat_id
               End If
               rLu(UBound(rLu)).Pat_idErwGL = rLu(UBound(rLu)).Pat_idErwGL & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwGL) Or LenB(rLu(UBound(rLu)).Pat_idErwGL) = 0, vNS, "/") & rs!Pat_id
              End If
             End If
             rs.Move 1
             If rs.EOF Then Exit Do
            Loop
            If glZLocker = 1 Or glZ = 1 Then
             rLu(UBound(rLu)).Pat_id = merkPat_ID
             rLu(UBound(rLu)).Pat_idUrsp = "E"
            End If
           End If
          End If ' erwZ = 1 nach Suche in Namen
         Case 3104: rLu(UBound(rLu)).Titel = Inhalt
         Case 3100: rLu(UBound(rLu)).NVorsatz = Inhalt
         Case 8411:
            rLw(UBound(rLw)).Langname = Inhalt
            rLw2(UBound(rLw2)).e.Langname = Inhalt
         Case 5001 ' Leistungsziffer
          ReDim Preserve rLL(UBound(rLL) + 1)
          rLL(UBound(rLL)).RefNr = rLu(UBound(rLu)).RefNr
          If obB Then rLL(UBound(rLL)).Verf = rLo(UBound(rLo)).Verf Else rLL(UBound(rLL)).Abkü = rLw(UBound(rLw)).Abkü
          rLL(UBound(rLL)).EBM = Inhalt
         Case 5005
          rLL(UBound(rLL)).Anzahl = Inhalt
         Case 8406
        ' zwingende Abfolge mit 5001 wurde 2/05 geprüft
          rLL(UBound(rLL)).goä = Inhalt
         Case 8614: rLL(UBound(rLL)).abrd = Inhalt
         Case 8428: rLo(UBound(rLo)).KuQu = Inhalt
         Case 8430: If obB Then rLo(UBound(rLo)).Quelle = Inhalt Else rLw(UBound(rLw)).Quelle = Inhalt: rLw2(UBound(rLw2)).e.Quelle = Inhalt
         Case 8431: If obB Then rLo(UBound(rLo)).Quelle = Inhalt Else rLw(UBound(rLw)).Quelle = Inhalt: rLw2(UBound(rLw2)).e.Quelle = Inhalt
         ' 8432 Abnahmedatum:
         Case 8432: If obB Then rLo(UBound(rLo)).AbnDat = BDTtoDate(Inhalt) Else rLw(UBound(rLw)).AbnDat = BDTtoDate(Inhalt): rLw2(UBound(rLw2)).e.AbnDat = BDTtoDate(Inhalt)
         Case 8433: rLo(UBound(rLo)).AbnDat = rLo(UBound(rLo)).AbnDat + BDTtoTime(Inhalt) ' Abnahmezeit
         Case 8418: rLw(UBound(rLw)).Teststatus = Inhalt: rLw2(UBound(rLw2)).e.Teststatus = Inhalt
         Case 8420: rLw(UBound(rLw)).Wert = Inhalt:       rLw2(UBound(rLw2)).e.Wert = Inhalt
         Case 8421: rLw(UBound(rLw)).Einheit = Inhalt:    rLw2(UBound(rLw2)).e.Einheit = Inhalt
         Case 8422: rLw(UBound(rLw)).Grenzwerti = Inhalt: rLw2(UBound(rLw2)).e.Grenzwerti = Inhalt
         Case 8460:
              Dim tt1$, tt2$, buch$, pos&, p2&, tt1n As New CString, tt2n As New CString, i1&, i2&
              Const notnu$ = "0123456789,."
              rLw2(UBound(rLw2)).Normbereich = Inhalt
erneut:
              tt1n.Clear
              tt2n.Clear
              pos = InStr(Inhalt, "-")
              If pos <> 0 Then
               tt1 = Trim$(Left$(Inhalt, pos - 1))
               For i1 = 1 To Len(tt1)
                buch = Mid$(tt1, i1, 1)
                If buch = " " Then tt1n.Clear
                If InStrB(notnu, buch) <> 0 Then tt1n.Append buch
               Next i1
               tt2 = Trim$(Mid$(Inhalt, pos + 1))
               For i2 = 1 To Len(tt2)
                buch = Mid$(tt2, i2, 1)
                If buch = " " Then Exit For
                If InStrB(notnu, buch) <> 0 Then tt2n.Append buch
               Next i2
              Else
               pos = InStr(Inhalt, "bis")
               If pos <> 0 Then
                tt2 = Trim$(Mid$(Inhalt, pos + 4))
                For i2 = 1 To Len(tt2)
                 buch = Mid$(tt2, i2, 1)
                 If buch = " " Then Exit For
                 If InStrB(notnu, buch) <> 0 Then tt2n.Append buch
                Next i2
               Else
                pos = InStr(Inhalt, "ab ")
                If pos <> 0 Then
                 tt1 = Trim$(Mid$(Inhalt, pos + 4))
                 For i1 = 1 To Len(tt1)
                  buch = Mid$(tt1, i1, 1)
                  If buch = " " Then Exit For
                  If InStrB(notnu, buch) <> 0 Then tt1n.Append buch
                 Next i1
                Else
                 p2 = InStr(Inhalt, ">")
                 pos = InStr(Inhalt, "<")
                 If p2 <> 0 Or (p2 = 0 And pos <> 0 And InStrB(Inhalt, "chwere") <> 0) Then ' Schwere Pankreasinsuffizienz:        < 100 µg/g Stuhl
                  If p2 <> 0 And pos <> 0 And p2 > pos Then
                   tt2 = Trim$(Mid$(Inhalt, pos + 1, p2 - pos - 1))
                   For i2 = 1 To Len(tt2)
                    buch = Mid$(tt2, i2, 1)
                    If buch = "=" Then Exit For
                    If InStrB(notnu, buch) <> 0 Then tt2n.Append buch
                   Next i2
                  ElseIf pos = 0 Then
                   tt1 = Trim$(Mid$(Inhalt, p2 + 1))
                   For i1 = 1 To Len(tt1)
                    buch = Mid$(tt1, i1, 1)
                    If buch = " " Then Exit For
                    If InStrB(notnu, buch) <> 0 Then tt1n.Append buch
                   Next
                  End If
                 ElseIf pos <> 0 Then
                  tt2 = Trim$(Mid$(Inhalt, pos + 1))
                  For i2 = 1 To Len(tt2)
                   buch = Mid$(tt2, i2, 1)
                   If buch = " " Then Exit For
                   If InStrB(notnu, buch) <> 0 Then tt2n.Append buch
                  Next i2
                 Else
                  If InStrB(Inhalt, "oxisch") <> 0 Or InStrB(Inhalt, "renztiter") <> 0 Then
                   tt2 = Inhalt
                   For i2 = 1 To Len(tt2)
                    buch = Mid$(tt2, i2, 1)
                    If tt2n.Value <> vNS And buch = " " Then Exit For
                    If InStrB(notnu, buch) <> 0 Then tt2n.Append buch
                   Next i2
                  End If
                 End If
                End If
               End If
              End If
'              GoTo erneut
              rLw2(UBound(rLw2)).NormU = tt1n
              rLw2(UBound(rLw2)).NormO = tt2n
         Case 8461:
                 rLw2(UBound(rLw2)).NormU = Inhalt
         Case 8462:
                 rLw2(UBound(rLw2)).NormO = Inhalt
         Case 8470
          If obB Then
           rLo(UBound(rLo)).Erklärung = IIf(IsNull(rLo(UBound(rLo)).Erklärung), vNS, rLo(UBound(rLo)).Erklärung & vbCrLf) & Inhalt
          Else
           rLw(UBound(rLw)).Erklärung = IIf(IsNull(rLw(UBound(rLw)).Erklärung), vNS, rLw(UBound(rLw)).Erklärung) & Inhalt
           rLw2(UBound(rLw2)).e.Erklärung = IIf(IsNull(rLw2(UBound(rLw2)).e.Erklärung), vNS, rLw2(UBound(rLw2)).e.Erklärung) & Inhalt
          End If
         Case 8480
           If obB Then
            rLo(UBound(rLo)).Kommentar = IIf(IsNull(rLo(UBound(rLo)).Kommentar), vNS, rLo(UBound(rLo)).Kommentar & vbCrLf) + Inhalt
            If keimz Then
             rLo(UBound(rLo)).Keimzahl = Inhalt
             keimz = 0
            End If
            If InStrB(Inhalt, "Keimzahl") > 0 Then keimz = -1
           Else
            rLw(UBound(rLw)).Kommentar = IIf(IsNull(rLw(UBound(rLw)).Kommentar), vNS, rLw(UBound(rLw)).Kommentar) & Inhalt
            rLw2(UBound(rLw)).e.Kommentar = IIf(IsNull(rLw2(UBound(rLw2)).e.Kommentar), vNS, rLw2(UBound(rLw2)).e.Kommentar) & Inhalt
           End If
         Case 8490:
'          rLU!auftrhinw = inhalt ' weiß noch nicht, auf welcher Ebene relevant
           rLw(UBound(rLw)).AuftrHinw = Inhalt
           rLw2(UBound(rLw2)).e.AuftrHinw = Inhalt
         Case 8615: ' Lebenslange Arztnummer
         Case 9212: rLs(0).VersionSatzb = Inhalt
         Case 201: rLs(0).Arztnr = Inhalt
         Case 203: rLs(0).Arztname = Inhalt
         Case 205: rLs(0).StraßePraxis = Inhalt
         Case 211: rLs(0).Arzt = Inhalt
         Case 212: rLs(0).Lanr = Inhalt
         Case 215: rLs(0).PLZPraxis = Inhalt
         Case 216: rLs(0).OrtPraxis = Inhalt
         Case 8300:
          Dim Spli$(), SpliZ&
          SpliZ = SplitNeu(Inhalt, ";", Spli)
          If SpliZ > 1 Then
           rLs(0).Labor = Spli(0)
           rLs(0).StraßeLabor = Spli(1)
           rLs(0).PLZLabor = Left$(Spli(2), 5)
           rLs(0).OrtLabor = Mid$(Spli(2), 6)
          Else
           rLs(0).Labor = Spli(0)
          End If
         Case 8320: rLs(0).Labor = Inhalt
         Case 8321: rLs(0).StraßeLabor = Inhalt
         Case 8322: rLs(0).PLZLabor = Inhalt
         Case 8323: rLs(0).OrtLabor = Inhalt
         Case 101: rLs(0).KBVPrüfnr = Inhalt
         Case 9106: rLs(0).Zeichensatz = Inhalt
         Case 8312: rLs(0).Kundenarztnr = Inhalt
         Case 9103: rLs(0).Erstellungsdatum = Inhalt
         Case 9202: rLs(0).Gesamtlänge = Inhalt
          If Not debugohneSpeichern Then
           Call laborxsaetzeSpeichern(SammelInsert, BezfSpei)
          End If
          If rs.State = 1 Then rs.Close
          Dim neuSatzID&
          If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
          myFrag rs, "SELECT MAX(satzid) msatzid FROM `laborxsaetze`"
          If Not rs.BOF Then neuSatzID = IIf(IsNull(rs!msatzid), 0, rs!msatzid)
          For i = 1 To UBound(rLu)
           rLu(i).SatzID = neuSatzID
           If Not debugohneSpeichern Then
            Call laborxusSpeichern(SammelInsert, BezfSpei, i)
           End If
'           IF rs1.State = 1 THEN rs1.Close
'           myFrag rs1, "SELECT MAX(refnr) AS mrefnr FROM `laborxus`;"
           Set rs1 = Nothing
           Set rs1 = myEFrag("SELECT last_insert_id()")
           Dim rs1refnr&
'           IF ISNULL(rs1!mrefnr) THEN rs1refnr = 0 ELSE rs1refnr = rs1!mrefnr
           rs1refnr = rs1.Fields(0)
           If rs1refnr = 0 Then MsgBox "Fehler in LaborDirektImport: last_insert_id()=0"
           For j = 1 To UBound(rLo)
            If rLo(j).RefNr = rLu(i).RefNr Then
             rLo(j).RefNr = rs1refnr
            End If
           Next j
           For j = 1 To UBound(rLw)
            If rLw(j).RefNr = rLu(i).RefNr Then
             rLw(j).RefNr = rs1refnr
             rLw2(j).e.RefNr = rs1refnr
            End If
           Next j
           For j = 1 To UBound(rLL)
            If rLL(j).RefNr = rLu(i).RefNr Then
             rLL(j).RefNr = rs1refnr
            End If
           Next j
          Next i
           If Not debugohneSpeichern Then
            Call laborxbaktSpeichern(SammelInsert, BezfSpei)
            For i = 1 To UBound(rLw2)
             Set rs = Nothing
             If LenB(rLw2(i).e.Einheit) = 0 Then rLw2(i).e.Einheit = "kA"
             myFrag rs, "SELECT 0 FROM `laborparameter` WHERE abkü = '" & rLw2(i).e.Abkü & "' AND einheit = '" & rLw2(i).e.Einheit & "'"
             If rs.BOF Then
              InsKorr DBCn, "INSERT INTO `laborparameter`(`abkü`,`einheit`,`langtext`,`aktzeit`) VALUES('" & rLw2(i).e.Abkü & "','" & rLw2(i).e.Einheit & "','" & rLw2(i).e.Langname & "'," & DatFor_k(Now()) & ")", rAf
             End If
             
             Dim laborid&
             laborid = indIns(DBCn, "laborxplab", "labor", rLs(0).Labor, "id")
             Set rs = Nothing
             myFrag rs, "SELECT `id` FROM `laborxpneu` WHERE `lid` = '" & laborid & "' AND `abkü` = '" & rLw2(i).e.Abkü & "' AND einheit = '" & rLw2(i).e.Einheit & "'"
             If rs.BOF Then
              InsKorr DBCn, "INSERT INTO `laborxpneu`(`lid`,`abkü`,`einheit`,`langtext`) VALUES('" & laborid & "','" & rLw2(i).e.Abkü & "','" & rLw2(i).e.Einheit & "','" & rLw2(i).e.Langname & "')", rAf
              Set rs = myEFrag("SELECT last_insert_id()")
             End If
             If Not rs.BOF Then
              If rs.Fields(0) = 0 Then MsgBox "Fehler in LaborDirektImport (2): last_insert_id()=0"
              Dim xpid&, Vorgehen% ' 0 = nehmen, 1 = hinzufügen, 2 = ändern, 3 = anderes Geschlecht suchen
              xpid = rs.Fields(0)
              If LenB(rLw2(i).geschlecht) = 0 Or Not IsNumeric(rLw2(i).geschlecht) Then rLw2(i).geschlecht = 3 ' unbekannt
              Set rs = Nothing
              myFrag rs, "SELECT `id`,`Geschlecht`,`Eingang` FROM `laborxpnb` WHERE `pid` = '" & xpid & "' AND geschlecht IN (9," & rLw2(i).geschlecht & ")  AND REPLACE(`NB`,' ','') = '" & REPLACE$(rLw2(i).Normbereich, " ", "") & "' AND uNg = '" & rLw2(i).NormU & "' AND ong = '" & rLw2(i).NormO & "'"
              If Not rs.BOF Then
               Vorgehen = 0
              Else
               Vorgehen = 3
              End If
              If Vorgehen = 3 Then
               Set rs = Nothing
               myFrag rs, "SELECT `id`,`Geschlecht`,`Eingang` FROM `laborxpnb` WHERE `pid` = '" & xpid & "' AND NOT geschlecht IN (9," & rLw2(i).geschlecht & ")  AND REPLACE(`NB`,' ','') = '" & REPLACE(rLw2(i).Normbereich, " ", "") & "' AND uNg = '" & rLw2(i).NormU & "' AND ong = '" & rLw2(i).NormO & "'"
               If Not rs.BOF Then
                Vorgehen = 2
               Else
                Vorgehen = 1
               End If
              End If
              Select Case Vorgehen
               Case 0
                If rs!Eingang > rLw2(i).Eingang Then
                 myEFrag "UPDATE `laborxpnb` SET eingang = " & DatFor_k(rLw2(i).Eingang) & ",uid = " & rLw2(i).e.RefNr & " WHERE id = " & rs!id, rAf
                 If rAf <> 1 Then
                  Debug.Print "Fehler in LaborDirektImport bei update eingang"
                 End If
                End If
               Case 1
                Set rs = Nothing
                InsKorr DBCn, "INSERT INTO `laborxpnb`(`pid`,`Geschlecht`,`Eingang`,`uNg`,`oNg`,`NB`,`uid`,`zahl`) VALUES(" & xpid & "," & rLw2(i).geschlecht & "," & DatFor_k(rLw2(i).Eingang) & ",'" & rLw2(i).NormU & "','" & rLw2(i).NormO & "','" & rLw2(i).Normbereich & "'," & rLw2(i).e.RefNr & ",1)", rAf
                Set rs = myEFrag("SELECT last_insert_id()")
                If Not rs.BOF Then If rs.Fields(0) = 0 Then MsgBox "Fehler in LaborDirektImport (3): last_insert_id()=0"
               Case 2
                myEFrag "UPDATE `laborxpnb` SET geschlecht = 9 WHERE id = " & rs!id, rAf
                If rAf <> 1 Then
                 Debug.Print "Fehler in LaborDirektImport bei update geschlecht"
                End If
              End Select
              rLw(i).nbid = rs.Fields(0)
              rLw2(i).e.nbid = rs.Fields(0)
              
              Select Case Vorgehen
               Case 0, 2
                myEFrag "UPDATE `laborxpnb` SET zahl = zahl+1 WHERE id = " & rs!id, rAf
              End Select
             End If ' Not rs.BOF Then
            Next i
'            Debug.Print SL.Item(ii).File.name
            Call laborxwertSpeichern(SammelInsert, BezfSpei)
            Call laborxleistSpeichern(SammelInsert, BezfSpei)
            Call myEFrag("UPDATE `laborxeingel` SET fertig = 1 WHERE datid = " & SL.Item(ii).DatID) ' Fertigkennzeichen setzen
'            Call DBCn.CommitTrans
'            Call DBCn.BeginTrans
           End If
         Case Else
          frm.Ausgeb "Unbekannte Kennung: " & Kennung & ", Inhalt: " & Inhalt, True
        End Select
'      Loop
     If obSchluss Or BrichAb Then
      Exit Do
     End If
    End If ' ltxt <> "" Then
    ltxt = txt
    If Not obeof Then Exit Do
    obSchluss = True
   Loop
   If BrichAb Then
    Exit Do
   End If
  Loop
  Close #1
'    END IF 'Not (oballe OR obDateiNeu) THEN ' oballe OR not FSO.fileexists(neupfad)
  frm.Bytes = frm.Bytes + 1
  If frm.Bytes Mod 100 = 0 Then
   frm.Prozent = Format$(frm.Bytes / frm.GesBytes * 100, "0.00")
   frm.Sekunden = Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
   frm.Beginn = Format$(T1, "hh:mm:ss")
   frm.GesDauer = Format$((T1a - T1 + frm.GesBytes / frm.Bytes * (Now - T1a)), "hh:mm:ss")
   frm.EndZp = Format$(T1 + (T1a - T1 + frm.GesBytes / frm.Bytes * (Now - T1a)), "hh:mm:ss")
  End If
  DoEvents
  If BrichAb Then
   frm.Ausgeb "Einlesen der Labor-DFÜ-Dateien abgebrochen ", True
   Exit Function
  End If
#If LaborDirektEinlesen Then ' Projekt -> Eigenschaften -> Erstellen Konstanten für bedingte Kompilierung
  On Error Resume Next
  Open LDEProt$ For Append As #251
  Print #251, Now() & ", fertig mit: " & SL.Item(ii).File.path
  Close #251
  On Error GoTo fehler
#End If
 Next
 frm.ZeilenBez = "Zeilen:"
 frm.Ausgeb "Fertig mit Einlesen der Labor-DFÜ-Dateien", True
 Call ForeignYes0
 Call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborDirektImport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborDirektImport

Function LaborErgPatId(lies As Lese, Optional mitLöschen%, Optional nurRef&)
' Pat_IDs ergänzen
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Const debugbit% = 0 ' liefert Analyse der Vorgänge in uverz & "anamnese\debug.txt", spart sich dafür Datenbankänderungen
 Dim altRefnr&, i&, sqlct$
 Dim j&
 Dim sPI As SortierPat_ID
 Dim SL As New SortierListe
' Dim rLX AS DAO.Recordset, rLN AS DAO.Recordset, rs AS DAO.Recordset
 Dim rLX As New ADODB.Recordset, rLN As New ADODB.Recordset
 Dim rLNZeitpunkt#, ZdüP%, LWerte$, GesAbk$, rAf&, laborneuAfn&, sql$, sql1$
 Dim obNeuerPat% ' Datensatz in rLX eintragen und mit SL neu anfangen
 Dim aktDS&, VergleichTot%
 
' nurRef = 3799
 On Error GoTo fehler
 T1a = Now
 lies.BytesBez = "Zeilen"
 lies.ZeilenBez = vNS
 
 GesAbk = "("
 lies.Ausgeb "Vergleiche die Laborübertragungsdateien mit dem BDT-Export ...", False
 lies.BytesBez = "Sätze:"
 lies.Bytes = 0
' IF LenB(DBCn) = 0 OR DBCn = "" THEN Call ConstrFestleg(0, lies.dlg)
  '"ZdüP, LWerte "
 ' mitLöschen = -1
  If mitLöschen Then
   If Not debugbit Then
    Call löschBezügeausLaborxus(nurRef)
   End If
  End If
nochmal:
  If LenB(sqliif) = 0 Then Zinit (lies.obMySQL)
  sql = "SELECT DISTINCT Auftragsnummer, Eingang, laborxus.Pat_id, Nachname, Vorname, GebDat, Geschlecht, GebüOrd, " & _
        "Erstellungsdatum, Wert, xw.Abkü, laborxus.RefNr, Pat_idUrsp, Pat_idlaborneu " & _
        "FROM (`laborxsaetze` INNER JOIN (`laborxeingel` INNER JOIN `laborxus` ON `laborxeingel`.DatID = `laborxus`.DatID) " & _
        "ON `laborxsaetze`.SatzID = `laborxus`.SatzID) INNER JOIN (SELECT refnr, abkü, " & sqliif & "(ISNULL(wert) OR wert = """",kommentar, wert) AS wert FROM `laborxwert` UNION SELECT refnr, verf AS abkü, keimzahl AS wert FROM `laborxbakt`) AS xw ON `laborxus`.RefNr = xw.RefNr " & _
        "WHERE NOT ISNULL(xw.abkü) AND ((ISNULL(verglichen) OR verglichen < " & DatFor_k(#1/1/2000#) & "))" ' NOT ISNULL(xw.Wert) AND xw.wert <> """" ' AND NOT ISNULL(xw.Wert) AND xw.wert <> """""
'        IF lies.obmysql THEN sql = replace$(sql, "iif", "if")
' 4.2.07: Bei Pat_id 107, laborident EIWE_T ist in der BDT-Datei kein Wert, obwohl in der übergebenen Datei "negativ" steht
'  GoTo nochmal
' einheit aus Ergebnisliste gestrichen
' die leeren werden oft gar nicht in der BDT-Datei mitgeliefert, was zu hohen Raten unangemessener Aberkennungen führt
'  IF NOT ISNULL(nurRef) THEN
   If nurRef > 0 Then
    sql = sql + " AND laborxus.Refnr IN (" & nurRef & ") "
   End If
'  END IF
'  sql = sql + " ORDER BY eingang, laborxus.refnr"
   sql = sql + " ORDER BY laborxus.refnr"
   
  sqlct = "SELECT COUNT(0) ct FROM (" & sql & ") AS Zl"
'  GoTo nochmal:
  Set rs = myEFrag(LCase$(sqlct))
  lies.GesBytes = rs!ct
   If mitLöschen Then
'   IF rs!ct > 30000 THEN
    If Not debugbit Then
     Call LöschRefNr(nurRef)
    End If
'   END IF
  End If

 If debugbit Then Open lies.snst.DebugDatei For Output As #300
' SET rLX = myEFrag(lcase(sql))
 obNeuerPat = -1 ' SL neu befüllen
' rLX.Open LCase$(sql), DBCn, adOpenDynamic, adLockOptimistic
 myFrag rLX, LCase$(sql)
 Do While Not rLX.EOF
  If debugbit Then
   Print #300, vbCrLf & "Überprüft wird der rLX-Datensatz:"
   For i = 0 To rLX.Fields.COUNT - 1
    Print #300, Left$(i & ". " & Space$(3), 3) & Left$(rLX.Fields(i).name & Space$(20), 20) & ": " & rLX.Fields(i).Value
   Next i
  End If
  Dim PrüfWert$
  PrüfWert = Trim$(REPLACE$(REPLACE$(REPLACE$(rLX!Wert, "<", ""), ":", ""), "-", ""))
  If debugbit Then Print #300, "Prüfwert: " & PrüfWert & " -> Prüfung erfolgt" & IIf(IsNumeric(PrüfWert), vNS, " nicht")
  If IsNumeric(PrüfWert) Then
   aktDS = aktDS + 1
'  SysCmd 4, "Vergleiche Labore: " + CStr(aktDS) + "/" + CStr(DSZahl) ' rs!ct von oben
' Falls in den entsprechenden Feldern von rLX schon was steht => löschen
'  rLN.Open "laborneu", DBCn
   ZdüP = ZdüP + 1 ' Zahl der überprüften Parameter
   LWerte = LWerte + rLX!Abkü + ":" + rLX!Wert + ", "
   If Not VergleichTot Then ' solange der Vergleich noch tot ist ...
'   Print #300, vbCrLf & "Refnr: " & rLX!RefNr & " (Pat_idlaborneu: " & rLX!Pat_idlaborneu & ": LWerte: " & LWerte
    If rLN.State = 1 Then rLN.Close
    
    Dim debugzl&
    sql = "SELECT * FROM (SELECT DISTINCT pat_id, abkü, wert, zeitpunkt, ABS(" & sqltodays & "(zeitpunkt)-" & sqltodays & "(" & DatFor_k(rLX!Eingang) & ")) AS abstand FROM `laborneu` WHERE abkü = '" & rLX!Abkü & "' AND wert = '" & rLX!Wert & "') AS innen ORDER BY abstand"
'    rLN.Open sql, DBCn, adOpenDynamic, adLockOptimistic
    myFrag rLN, sql
    If debugbit Then
     debugzl = 0
     Print #300, "Dazu wurden in `laborneu` gefunden:"
     Do
      If rLN.BOF Or rLN.EOF Then Exit Do
       debugzl = debugzl + 1
       Print #300, "  " & debugzl & ": rLN!Pat_id: " & rLN!Pat_id, rLN!Abkü, rLN!Wert, rLN!Zeitpunkt, rLN!abstand, IIf(rLN!abstand < 32, vNS, "Nicht berücksichtigt: Abstand zu groß")
      rLN.Move 1
     Loop
     If Not rLN.BOF Then
      rLN.MoveFirst
     End If
    End If
    If Not rLN.BOF Then
     GesAbk = GesAbk & "'" & rLN!Abkü & "',"
     If obNeuerPat Then
      Do
       If Abs(rLN!Zeitpunkt - rLX!Eingang) < 32 Then ' 1 Monat Grenze
' SL = Liste der Patienten, die für die aktuelle Laborzusammenstellung in Frage kommen
        Set sPI = New SortierPat_ID
        sPI.Pat_id = rLN!Pat_id
        sPI.Zeitpunkt = rLN!Zeitpunkt
        Call SL.Add(sPI)
       End If
       rLN.MoveNext
       If rLN.EOF Then Exit Do
'      IF rLN!Abkü <> rLX!Abkü OR rLN!Wert <> rLX!Wert THEN Exit Do
      Loop
     Else ' obNeuerPat
' diejenigen der bisher in Frage kommenden Patienten, die auch für den aktuellen Laborwert noch in Frage kommen, erneut kennzeichnen
      Do
'      For Each sPI In SL
       For i = 1 To SL.COUNT
        Set sPI = SL.Item(i)
        If sPI.Pat_id = rLN!Pat_id And sPI.Zeitpunkt = rLN!Zeitpunkt Then
         sPI.Knz = -1
         Exit For
        End If
       Next i
       rLN.MoveNext
       If rLN.EOF Then Exit Do
       If UCase$(rLN!Abkü) <> UCase$(rLX!Abkü) Or rLN!Wert <> rLX!Wert Then
        Exit Do
       End If
      Loop
' die restlichen löschen und die Kennzeichen zurücksetzen
      For i = SL.COUNT To 1 Step -1
       Set sPI = SL.Item(i)
       If sPI.Knz = 0 Then
        Call SL.Remove(i)
       Else
        sPI.Knz = 0
       End If
      Next i
      If SL.COUNT = 0 Then
       VergleichTot = -1
      End If ' SL.COUNT
     End If ' obNeuerPat
    Else ' ' Not rLN.NoMatch THEN else
    End If ' Not rLN.NoMatch THEN
    
    If debugbit Then
     Print #300, vbCrLf & " in der Liste befinden sich momentan (SL.Count): " & SL.COUNT & IIf(SL.COUNT > 0, ", und zwar:", "!!!!")
     For i = 1 To SL.COUNT
      Print #300, "   Pat_id: " & SL.Item(i).Pat_id
     Next i
     Print #300, "Gesamt-Abkürzungs-String: " & GesAbk
     Print #300, ""
    End If ' debugbit Then
   End If ' VergleichTot
   altRefnr = rLX!RefNr
  End If
  rLX.Move 1
'  IF rLX.EOF THEN obNeuerPat = -1 ELSE IF rLX!Auftragsnummer <> altAuftragsNummer THEN obNeuerPat = -1 ELSE obNeuerPat = 0
  If rLX.EOF Then obNeuerPat = -1 Else If rLX!RefNr <> altRefnr Then obNeuerPat = -1 Else obNeuerPat = 0
  If obNeuerPat Then
    If debugbit Then
     Print #300, vbCrLf & " Neuer Patient" & vbCrLf
    End If
' dann fängt später der nächste Auftrag an, der aktuelle wird eingetragen
    Dim nPat_idlaborneu$, nPat_ID&, nPat_idUrsp$, SLPat_id&
    VergleichTot = 0
    rLX.Move -1
    
    SLPat_id = -1
    If SL.COUNT = 0 Then
    ElseIf SL.COUNT = 1 Then
     SLPat_id = SL.Item(1).Pat_id
     rLNZeitpunkt = SL.Item(1).Zeitpunkt
     nPat_idlaborneu = SLPat_id
    ElseIf rLX!Pat_id <> 0 And Not IsNull(rLX!Pat_id) Then ' wenn die Laborwertekombination für mehrere Patienten stimmen würde, aber der richtige schon aus Namen/Geburtsdatum hervorgeht => diesen davon nehmen
     For i = 1 To SL.COUNT
      If SL.Item(i).Pat_id = rLX!Pat_id Then
       SLPat_id = SL.Item(i).Pat_id
       rLNZeitpunkt = SL.Item(i).Zeitpunkt
       nPat_idlaborneu = SLPat_id
       Exit For
      End If
     Next
    Else ' wenn also unbekannt ist, welcher Patient der richtige ist
     nPat_idlaborneu = SL.Item(1).Pat_id
     For i = 2 To SL.COUNT
      nPat_idlaborneu = nPat_idlaborneu & "/" & SL.Item(i).Pat_id
     Next
    End If
    
    If SLPat_id <> -1 Then ' SL.Count = 1
     nPat_ID = IIf(IsNull(rLX!Pat_id), 0, rLX!Pat_id)
     nPat_idUrsp = IIf(IsNull(rLX!Pat_idUrsp), vNS, rLX!Pat_idUrsp)
'    IF SL.Count = 1 THEN
     If nPat_ID = 0 Then
      nPat_ID = SLPat_id
      nPat_idUrsp = "L"
     ElseIf nPat_ID = SLPat_id Then
      nPat_idUrsp = "B" ' "E"rsteinlesung, "B"eide, "L"abordirektdateien
     Else
      nPat_idUrsp = "W" ' "W"idersprüchliche Patientenangaben!
     End If
     
     GesAbk = Left$(GesAbk, Len(GesAbk) - 1) + ")"
     If Not debugbit Then
      Call myEFrag("UPDATE `laborneu` SET refnr = " & rLX!RefNr & " WHERE pat_id = " & SLPat_id & " AND zeitpunkt = " & DatFor_k(rLNZeitpunkt) & " AND abkü IN " & GesAbk, laborneuAfn)
      If laborneuAfn = 0 Then Err.Raise 9999, , "laborneuAFN = 0"
     End If
'     END IF ' SL.count = 1
    
     If Not debugbit Then
      Call myEFrag("UPDATE `laborxus` SET pat_idUrsp= '" & nPat_idUrsp & "',Pat_id = " & nPat_ID & ",Zeitpunktlaborneu = " & DatFor_k(rLNZeitpunkt) & ",LWerte = '" & LWerte & "',ZdüP = " & ZdüP & ", Pat_idlaborneu = '" & nPat_idlaborneu & "', AfN = " & laborneuAfn & " WHERE refnr = " & rLX!RefNr, rAf)
     End If
    End If ' SL.count = 1
    If SLPat_id <> -1 Or (Now - BDTtoDate(rLX!Erstellungsdatum)) > 7 Then
     If Not debugbit Then
      Call myEFrag("UPDATE `laborxus` SET verglichen = " & DatFor_k(Now) & ",LWerte = '" & LWerte & "',ZdüP = " & ZdüP & ", ZdiP = " & SL.COUNT & " WHERE refnr = " & rLX!RefNr, rAf)
     End If
    End If
    rLX.Move 1 ' nach vorher -1
'   END IF '    IF SL.Count <= 1 OR ISNULL(rLX!Pat_ID) OR rLX!Pat_ID = 0 OR Not obEnthalten THEN
     For i = SL.COUNT To 1 Step -1
      Call SL.Remove(i)
     Next i
     ZdüP = 0
     GesAbk = "("
     LWerte = vNS
  End If ' obNeuerPat
  lies.Bytes = lies.Bytes + 1
  If lies.Bytes Mod 100 = 0 Then
   If lies.GesBytes <> 0 Then lies.Prozent = Format$(lies.Bytes / lies.GesBytes * 100, "0.00")
   lies.Sekunden = Format$((Now - T1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
   lies.Beginn = Format$(T1, "hh:mm:ss")
   DoEvents
  End If
  
  If rLX.EOF Then Exit Do
  If BrichAb Then
   lies.Ausgeb "LaborErgPatID (Vergleich der Labor-DFÜ-Datein mit dem BDT-Export) bei RefNr " & rLX!RefNr & " abgebrochen", True
   Exit Function
  End If
  DoEvents
 Loop ' while not rlx.eof
 lies.Ausgeb "Fertig mit LaborErgPatId (Vergleich der Labor-DFÜ-Datein mit dem BDT-Export)", True
 If debugbit Then
  Close #300
 End If
 lies.BytesBez = "Bytes:"
 lies.ZeilenBez = "Zeilen:"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborErgPatID/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborErgPatID

Function LöschRefNr(Optional RefNr&)
 Dim rAf&
 Dim rs As ADODB.Recordset
 If RefNr = 0 Then
  Set rs = myEFrag("UPDATE `laborneu` SET refnr = null WHERE NOT ISNULL(refnr)", rAf)
 Else
  Set rs = myEFrag("UPDATE `laborneu` SET refnr = null WHERE refnr = " & RefNr, rAf)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LöschRefNr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LöschRefNr

Function RBDTtoDate(DaT) As Date ' für umgekehrtes BDT-Format
 On Error GoTo fehler
 If DaT = Space$(8) Or DaT Like "00*" Then
 Else
  RBDTtoDate = CDate(Format$(Left$(DaT, 8), "####\.##\.##"))
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RBDTtoDate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rbdttodate

Function BDTtoDateTime(DaT) As Date
 On Error GoTo fehler
 If DaT = Space$(8) Or DaT Like "00*" Then
 Else
  On Error Resume Next
  BDTtoDateTime = CDate(Format$(Left$(DaT, 14), "####\.##\.## ##\:##\:##"))
  If BDTtoDateTime = 0 Then
   BDTtoDateTime = CDate(Format$(Left$(DaT, 14), "##\.##\.#### ##\:##\:##"))
  End If
  On Error GoTo fehler
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BDTtoDateTime/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' BDTtoDateTime

Function BDTtoDate(DaT) As Date ' für BDT-Format
 On Error GoTo fehler
' If Len(DaT) > 8 Then Stop
 If DaT = Space$(8) Or DaT Like "00*" Then
 Else
  On Error Resume Next
  BDTtoDate = CDate(Format$(Left$(DaT, 8), "##\.##\.####"))
  If BDTtoDate = 0 Then
   BDTtoDate = CDate(Format$(Left$(DaT, 8), "####\.##\.##"))
  End If
  On Error GoTo fehler
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BDTtoDate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' bdttodate

Function BDTtoTime(DaT) As Date ' für BDT-Format
 On Error GoTo fehler
 If DaT Like "????" Then
  BDTtoTime = CDate(Format$(DaT, "00:00"))
 ElseIf DaT Like "??????" Then
  BDTtoTime = CDate(Format$(DaT, "00:00:00"))
 ElseIf DaT = "*" Then
  BDTtoTime = "00:00"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BDTtoTime/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' BDTtoTime

Function indIns&(Cn As ADODB.Connection, Tb$, fld$, Wert$, idFld$)
 Dim rs As New ADODB.Recordset, rAf&
 Dim varlen&, pos&
 Dim STyp$
 On Error GoTo fehler
 Dim sql$
 sql = "SELECT * FROM `" & Tb & "` WHERE `" & fld & "` = '" & doUmwfSQL(Wert, LVobMySQL) & "'"
 myFrag rs, sql, adOpenStatic, Cn, adLockReadOnly
 If rs.BOF Then
  Set rs = Nothing
  myFrag rs, "SELECT data_type dt, character_maximum_length cml FROM information_schema.`COLUMNS` C WHERE table_schema = '" & HADBName & "' AND table_name = '" & Tb & "' AND column_name = '" & fld & "'", , Cn
'  rs.Open "SHOW FULL COLUMNS FROM `" & Tb & "` WHERE field = '" & fld & "'", cn
  If Not rs.BOF Then
   STyp = rs!DT
'   IF InStrB(STyp, "varchar(") <> 0 THEN
   If STyp = "varchar" Then
'    pos = InStr(STyp, "(")
'    varlen = Mid$(STyp, pos + 1, InStr(STyp, ")") - pos - 1)
    varlen = rs!cml
    If varlen < Len(Wert) Then
     If Len(Wert) > 767 Then
' hier müßte noch der Index verkürzt werden
      myEFrag "ALTER TABLE `" & Tb & "` MODIFY `" & fld & "` varchar(" & Len(Wert) & ")", , Cn
     Else
      myEFrag "ALTER TABLE `" & Tb & "` MODIFY `" & fld & "` varchar(" & Len(Wert) & ")", , Cn
     End If
    End If ' STyp = "varchar"
   End If ' NOT rs.BOF
  End If ' rs.BOF
  Set rs = Nothing
  InsKorr Cn, "INSERT INTO `" & Tb & "`(`" & fld & "`) VALUES('" & doUmwfSQL(Wert, LVobMySQL) & "')", rAf
  Set rs = myEFrag("SELECT last_insert_id()", , Cn)
  If Not rs.BOF Then If rs.Fields(0) = 0 Then MsgBox "Fehler in indIns: last_insert_id()=0"
  indIns = rs.Fields(0)
 Else
  indIns = rs.Fields(idFld)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in indIns/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' indIns

Function löschBezügeausLaborxus(Optional RefNr&)
   Call myEFrag("UPDATE `laborxus` SET " & _
                     "Zeitpunktlaborneu = " & DatFor_k(CDate(0)) & "," & _
                     "verglichen = null," & _
                     "LWerte = """"," & _
                     "ZdüP = 0," & _
                     "AfN = 0," & _
                     "Pat_idlaborneu = """" " & _
                     IIf(RefNr <> 0, "WHERE refnr = " & RefNr, vNS), rAf)
   Call myEFrag("UPDATE `laborxus` SET " & _
                     "pat_idursp = ""E"" " & _
                     IIf(RefNr <> 0, "WHERE refnr = " & RefNr & " and", "where") & " pat_idursp IN (""B"",""W"")", rAf)
   Call myEFrag("UPDATE `laborxus` SET " & _
                     "pat_idursp = """"," & _
                     "pat_id = 0 " & _
                     IIf(RefNr <> 0, "WHERE refnr = " & RefNr & " and", "where") & " pat_idursp = ""L""", rAf)
End Function ' löschBezügeausLaborxus

Function SpMod%(SpValLen%, TName, rs0 As ADODB.Recordset, Optional SpVal$) ' Spalte modifizieren
'Dim ZCStr$
Dim keinetrans%, SpName$, maxL%
Dim rsc As New ADODB.Recordset
Dim ausgTxt$, FNr&, FText$
On Error GoTo fehler
SpName = rs0!COLUMN_NAME
If IsNull(rs0!character_maximum_length) Then
 maxL = 32000
Else
' maxL = IIf(rs0!character_maximum_length > 255, 5, rs0!character_maximum_length)
 If rs0!character_maximum_length > 32000 Then
  maxL = 32000
 Else
  maxL = rs0!character_maximum_length
 End If
End If
'If rs0!data_type = 129 THEN obMemo = True
If SpValLen > maxL And maxL > 0 Then ' longtext
 If obMitAlterTab = 0 Then
'  Lese.Ausgeb "Feldinhaltverkürzung Tabelle '" & TName & "' Feld '" & SpName & "' (Länge:" & SpValLen & "->" & maxL & ");", True
  Lese.Ausgeb "Feldvergrößerung wäre nötig gewesen für Tabelle '" & TName & "' Feld '" & SpName & "';" & maxL & " -> " & SpValLen & "; verkürze Inhalt: " & Left$(SpVal, maxL) & IIf(SpVal <> vNS, "|", vNS) & Right$(SpVal, SpValLen - maxL), True
  SpMod = maxL
'  SpVal = LEFT(SpVal, maxL) ' Verkürzung
 Else
'  ZCStr = DBCnS ' DBCn.ConnectionString
  On Error Resume Next
  If obTrans <> 0 Then ComTrans DBCn: obTrans = 0 ' DBCn.CommitTrans: obTrans = 0 ' geändert 16.6.24
  keinetrans = Err.Number
  On Error GoTo fehler
  Call DBCnOpen
'  DBCn.Close
''  Call CnOpen(False, ZCStr)
'  'Call acon(quelleT)
'  DBCn.Open
  If Lese.obMySQL Then Call myEFrag("USE `" & Lese.MyDB & "`")
  Call ForeignNo0
  Call ForeignNo1
  If DBCn.State = 0 Then
   DBCnOpen
  End If
  If LenB(sqlText) = 0 Then Zinit (Lese.obMySQL)
  If SpValLen > maxL Then
   If Not Lese.obMySQL And SpValLen > 255 Then
    Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & "MEMO", rAf)
    SpMod = -1
   Else
    On Error Resume Next
nochmal:
    Err.Clear
    If Lese.obMySQL Then
     Set rsc = Nothing
     myFrag rsc, "SHOW FULL COLUMNS FROM `" & TName & "` WHERE field = '" & SpName & "'"
     Call myEFrag("SET SESSION innodb_strict_mode=off") ' wenn Zeilengröße schon > 8125 osä
     Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & sqlText & "(" & SpValLen & ") " & IIf(IsNull(rsc!collation), vNS, " COLLATE " & rsc!collation) & " default " & IIf(IsNull(rsc!Default), " null", "'" & rsc!Default & "'") & " comment '" & rsc!Comment & "'", rAf, , True, FNr, FText)
    Else ' Lese.obMySQL Then
     Call myEFrag("ALTER TABLE `" & TName & "`" & sqlALTER & " COLUMN `" & SpName & "` " & sqlText & "(" & SpValLen & ")", rAf, , True, FNr, FText)
    End If ' Lese.obMySQL Then else
'    FNr = Err.Number
'    FText = Err.Description & " " & vbCrLf & Err.LastDllError
    On Error GoTo fehler
   End If
   If FNr = 0 Then
    ausgTxt = "Feldvergrößerung Tabelle '" & TName & "' Feld '" & SpName & "';" & maxL & " -> " & SpValLen & ";" & IIf(SpVal <> vNS, " wg. Inhalt: " & Left$(SpVal, maxL) & "|" & Right$(SpVal, SpValLen - maxL) & ";" & rAf & " Datensätze geändert", vNS)
    SpMod = -1
   Else
    ausgTxt = FText & vbCrLf & "Feldvergrößerung gescheitert Tabelle '" & TName & "' Feld '" & SpName & "';" & maxL & " -> " & SpValLen & "; verkürze Inhalt: " & IIf(SpVal <> vNS, Left$(SpVal, maxL) & "|" & Right$(SpVal, SpValLen - maxL), vNS)
'    SpVal = LEFT(SpVal, maxL) ' Verkürzung
    SpMod = maxL
   End If
   Lese.Ausgeb ausgTxt, True
   On Error Resume Next
   Open Lese.snst.DebugDatei For Append As #399
   Dim akterr&
   akterr = Err.Number
   On Error GoTo fehler
   If akterr <> 0 Then
    Open REPLACE$(Lese.snst.DebugDatei, "u:", LiServer & "daten\eigene Dateien") For Append As #399
   End If
   Print #399, ausgTxt
   Close #399
  End If
  DoEvents
  Call ForeignYes0
  Call ForeignYes1
  If keinetrans = 0 Then
   If DBCn.State = 0 Then
    DBCnOpen
   End If
'   DBCn.BeginTrans: obTrans = 1
   BegTrans
   keinetrans = 0
  End If
 End If
Else
' MsgBox "Erst mal schauen, ob er hierherkommt"
' Err.Raise 999, , "Erst mal schauen, ob er hierherkommt"
' IF obMitAlterTab = 0 THEN
'  Lese.Ausgeb "Feldudt.mwandlung Tabelle '" & Tname & "' Feld '" & SpName & "' (Text mit Länge:" & Len(SpVal) & "-> Memo);" ,true
' ElseIf Not lese.obmysql THEN ' Vermutlich tritt Phänomen nur bei Access auf
'  ZCStr = DBCn.ConnectionString
'  ON Error Resume Next
'  DBCn.CommitTrans
'  keineTrans = Err.Number
'  ON Error GoTo fehler
'  DBCn.Close
'  DBcnOpen ZCStr
'  IF lese.obmysql THEN Call myEFrag("use " & myDB)
'  call foreignno
'  Call myEFrag("ALTER TABLE " & "`" & Tname & "`" & " " & IIf(lese.obmysql, "MODIFY", "ALTER") & " Column " & "`" & SpName & "`" & " " & "MEMO")
'   ausgTxt = "Feldudt.mwandlung Tabelle '" & Tname & "' Feld '" & SpName & "';" & maxL & " -> Memo; wg. Inhalt: " & SpVal
'   Lese.Ausgeb ausgTxt ,true
'   Open Lese.snst.DebugDatei For Append AS #399
'   Print #399, ausgTxt
'   Close #399
'   SpMod = True
' END IF
End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SpMod/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SpMod

