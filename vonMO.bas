Attribute VB_Name = "vonMo"
Option Explicit
Option Compare Text

Const MoSer$ = "wser"
Public Const MOCStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSer & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
' Public MOCon As New ADODB.Connection
  
Type memoType ' zur Übertragung der FMemo-Felder aus Medical Office
' patnr As Long
' FSur As Long
 znr As Integer
 mx As Integer
 ebn As Integer
 enr As String
 Text As String
End Type ' memoType

Type ebType
 nr As Integer
 Wert As Integer
End Type ' ebtype

Declare Sub CopyMemoryPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination&, ByVal Sourc&, ByVal length&)

Public Sub ebTypeDelete(ByRef sArray() As ebType, ByVal nDelPos&, ByVal bRedimSize%)
  Dim nPtr&, nSize&
  Static loezahl%
  ' Größe des Arrays bestimmen, falls nicht angegeben
  nSize = UBound(sArray)

  If nDelPos <= nSize Then loezahl = loezahl + 1
  If nDelPos < nSize Then
   ' Element aus Array löschen und alle nachfolgende Elemente nach vorne schieben
   nPtr = VarPtr(ByVal sArray(nDelPos))
   CopyMemoryPtr VarPtr(sArray(nDelPos)), VarPtr(sArray(nDelPos + 1)), VarPtr(sArray(nSize)) - VarPtr(sArray(nDelPos))
   CopyMemoryPtr VarPtr(sArray(nSize)), VarPtr(nPtr), Len(nPtr)
  End If ' nDelPos < nSize Then
 
  ' Array ggf. autom. um 1 Element verkleinern
  If bRedimSize And loezahl <> 0 Then
   nSize = nSize - loezahl
   If nSize < 0 Then nSize = 0
   ReDim Preserve sArray(nSize)
   loezahl = 0
  End If ' bRedimSize Then
End Sub ' ebTypeDelete(ByRef sArray() As ebtype, ByVal nDelPos&, ByVal bRedimSize%)

Public Sub memoTypeDelete(ByRef sArray() As memoType, ByVal nDelPos&, ByVal bRedimSize%)
  Dim nPtr&, nSize&
  Static loezahl%
  ' Größe des Arrays bestimmen, falls nicht angegeben
  nSize = UBound(sArray)

  If nDelPos <= nSize Then loezahl = loezahl + 1
  If nDelPos < nSize Then
   ' Element aus Array löschen und alle nachfolgende Elemente nach vorne schieben
   nPtr = VarPtr(ByVal sArray(nDelPos))
   CopyMemoryPtr VarPtr(sArray(nDelPos)), VarPtr(sArray(nDelPos + 1)), VarPtr(sArray(nSize)) - VarPtr(sArray(nDelPos))
   CopyMemoryPtr VarPtr(sArray(nSize)), VarPtr(nPtr), Len(nPtr)
  End If ' nDelPos < nSize Then
 
  ' Array ggf. autom. um 1 Element verkleinern
  If bRedimSize And loezahl <> 0 Then
   nSize = nSize - loezahl
   If nSize < 0 Then nSize = 0
   ReDim Preserve sArray(nSize)
   loezahl = 0
  End If ' bRedimSize Then
End Sub ' memoTypeDelete(ByRef sArray() As ebtype, ByVal nDelPos&, ByVal bRedimSize%)

' Omit plngLeft & plngRight; they are used internally during recursion
Public Function ebQuickSort(ByRef pvarArray() As ebType, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long) As ebType()
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As ebType
    Dim varSwap As ebType
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst).nr < varMid.nr And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid.nr < pvarArray(lngLast).nr And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then ebQuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then ebQuickSort pvarArray, lngFirst, plngRight
    ebQuickSort = pvarArray
End Function ' ebQuickSort


Public Function ParseMemo(FMemo$, MStr() As memoType) ' pNr&, FSur&,
 Dim gl&, pos&, MAX&, aktmax&, altmax&, tlen&, ie&, altie&, mznr%, i&
 Dim obDruck%, txt$, ebS$
 Dim eb() As ebType
 If FMemo <> "" Then
  pos = 1
  ie = 0
 ' Open "v:\Parsememo31.txt" For Output As #255
  Do
   obDruck = 0
   tlen = Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1))
   If pos = 1 Then gl = tlen: MAX = gl
   altmax = MAX
   aktmax = tlen + pos + 1
   If aktmax >= 0 And aktmax <= gl Then
    If ((Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Or pos > MAX) Then
     altie = ie
     If aktmax <= MAX Then ie = ie + 1
     If pos > MAX Then
      mznr = -1
      ie = 0
      If SafeArrayGetDim(MStr) <> 0 Then
       For i = 0 To UBound(MStr)
        If MStr(i).mx >= aktmax And MStr(i).znr > mznr Then mznr = MStr(i).znr ' MStr(i).patnr = pNr And
       Next i
       For i = 0 To UBound(MStr)
        If MStr(i).mx >= aktmax And MStr(i).znr = mznr And MStr(i).ebn > ie Then ie = MStr(i).ebn ' MStr(i).patnr = pNr And
       Next i
      End If ' SafeArrayGetDim(MStr) <> 0 Then
      ie = ie + 1
     End If ' pos > MAX
     If ie <= altie Then
      If ie < altie Then
       For i = UBound(eb) To 0 Step -1
        If eb(i).nr > ie Then Call ebTypeDelete(eb, i, 0)
       Next i
       Call ebTypeDelete(eb, UBound(eb) + 1, -1)
      End If ' ie < AltID Then
      For i = 0 To UBound(eb)
       If eb(i).nr = ie Then eb(i).Wert = eb(i).Wert + 1
      Next i
     Else ' ie < AltID Then
      If SafeArrayGetDim(eb) = 0 Then ReDim eb(0) Else ReDim Preserve eb(UBound(eb) + 1)
      eb(UBound(eb)).nr = ie
      eb(UBound(eb)).Wert = 1
     End If ' ie <= altie Then
     MAX = aktmax
     obDruck = 1
    End If ' ((Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Or pos > MAX) Then
   End If ' aktmax >= 0 And aktmax <= gl Then
   If obDruck = 1 Or pos = 1 Then ' or pos=1 or true
    txt = Mid(FMemo, pos + 2, tlen)
    If txt <> String$(tlen, Chr$(0)) Then
     ebS = ""
     Dim eb2() As ebType
     If SafeArrayGetDim(eb) <> 0 Then
      eb2 = ebQuickSort(eb)
      For i = 0 To UBound(eb)
       ebS = ebS & eb(i).Wert & "."
      Next i
      If Right$(ebS, 1) = "." Then ebS = Left$(ebS, Len(ebS) - 1)
     End If ' SafeArrayGetDim(eb) <> 0 Then
     If SafeArrayGetDim(MStr) = 0 Then ReDim MStr(0) Else ReDim Preserve MStr(UBound(MStr) + 1)
'     MStr(UBound(MStr)).patnr = pNr
'     MStr(UBound(MStr)).FSur = FSur
     MStr(UBound(MStr)).znr = pos
     MStr(UBound(MStr)).mx = MAX
     MStr(UBound(MStr)).ebn = ie
     MStr(UBound(MStr)).enr = ebS
     If Asc(Right$(txt, 1)) = 0 Then txt = Left$(txt, Len(txt) - 1)
     MStr(UBound(MStr)).Text = txt
'     Print #255, pos & "|" & MAX & "|" & ie & "|" & ebS & "|" & Mid$(FMemo, pos, 1) & "|" & Asc(Mid$(FMemo, pos, 1)) & "|" & Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1)) & "| aktMax: " & aktmax & "| altMax: " & altmax & "| Laenge: " & Len(txt) & "| EndByte: " & Asc(Mid$(FMemo, aktmax, 1)) & "|" & txt & vbCrLf
    Else ' txt <> String$(Chr$(0), tlen) Then
     pos = pos + tlen
    End If ' txt <> String$(Chr$(0), tlen) Then
   End If ' obdruck = 1 Then ' or pos=1 or true
   pos = pos + 1
   If pos >= gl Then Exit Do
  Loop
'  Close #255
  For i = UBound(MStr) - 1 To 0 Step -1
   If InStr(MStr(i + 1).enr, MStr(i).enr & ".") = 1 And MStr(i + 1).enr <> MStr(i).enr Then ' MStr(i).patnr = MStr(i + 1).patnr And MStr(i).FSur = MStr(i + 1).FSur And
    Call memoTypeDelete(MStr, i, 0)
   End If
   Call memoTypeDelete(MStr, UBound(MStr) + 1, True)
  Next i
 End If ' FMEMO<>""
End Function ' ParseMemo


Public Function doPatvonMO(pNr&, pid&)
 Dim rsNa As New ADODB.Recordset, rsFa As New ADODB.Recordset
 Dim MStr() As memoType
' ReDim MStr(0)
 Call LöschePat(pid)
 Tinit
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
 rsNa.Open "SELECT * from patstamm WHERE FSurogat=" & pNr, MOCon, adOpenStatic, adLockReadOnly
 rNa(0).Pat_id = pNr
 rNa(0).Nachname = rsNa!FNachname
 rNa(0).NVors = rsNa!FNamensvorsatz
 rNa(0).NVorsatz = rsNa!FNamenszusatz
 rNa(0).Titel = rsNa!FTitel
 rNa(0).Vorname = rsNa!FVorname
 rNa(0).Vorname = rsNa!FVorname
 Select Case rsNa!FGeschlecht:  Case "2": rNa(0).geschlecht = "w":  Case "1": rNa(0).geschlecht = "m":  Case Else: rNa(0).geschlecht = rsNa!FGeschlecht: End Select
 rNa(0).GebDat = DateSerial(Mid$(rsNa!fgeburtsdatum, 1, 4), Mid$(rsNa!fgeburtsdatum, 5, 2), Mid$(rsNa!fgeburtsdatum, 7, 2))
 rNa(0).Versichertennummer = rsNa!FAktversichertennr
 rNa(0).AufnDat = CDate("1.1.1890") + rsNa!FErstkontakt
 rNa(0).Straße = rsNa!FStrasse
 rsFa.Open "SELECT COALESCE(CONVERT(fmemo USING latin1),'') fmemo FROM patfall WHERE fpatnr=" & pNr, MOCon, adOpenStatic, adLockReadOnly
 If Not rsFa.BOF Then
  Do While Not rsFa.EOF
   Call ParseMemo(rsFa!FMemo, MStr()) ' rsFa!fpatnr, rsFa!fsurogat,
   rsFa.MoveNext
  Loop
 End If
 Set rsFa = Nothing
 Set rsFa = Nothing ' wirkt witzigerweise erst beim zweiten Mal (?!)
 rsFa.Open "SELECT COALESCE(CONVERT(fmemo USING latin1),'') fmemo FROM patstamm WHERE fsurogat=" & pNr, MOCon, adOpenStatic, adLockReadOnly
 If Not rsFa.BOF Then
   Call ParseMemo(rsFa!FMemo, MStr()) ' rsFa!fpatnr, rsFa!fsurogat,
   rsFa.MoveNext
 End If
End Function ' doPatvonMO(pNr&, pid&)

Public Function suchfi(pNr&, fI$)
 Dim altt$, gefu%, i&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
 rst.Open "SELECT c.TABLE_NAME tn -- , a.column_name cn" & vbCrLf & _
          "FROM information_schema.columns c" & vbCrLf & _
          "-- LEFT JOIN information_schema.columns a ON c.TABLE_CATALOG=a.TABLE_CATALOG AND c.TABLE_SCHEMA=a.TABLE_SCHEMA AND c.table_name=a.TABLE_NAME" & vbCrLf & _
          "WHERE c.table_catalog='def' AND c.table_schema='medoff' AND (c.COLUMN_NAME='fpatnr' OR (c.TABLE_NAME='patstamm' AND c.COLUMN_NAME='fsurogat')) GROUP BY c.table_name ORDER BY c.table_name", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
  If rst!Tn <> altt Then
   Set rsu = Nothing
   rsu.Open "SELECT * FROM `" & rst!Tn & "` WHERE `" & IIf(rst!Tn = "patstamm", "fsurogat", "fpatnr") & "`=" & pNr, MOCon, adOpenStatic, adLockReadOnly
   gefu = 0
   Do While Not rsu.EOF
    For i = 0 To rsu.Fields.COUNT - 1
     If Not (IsNull(rsu.Fields(i))) Then
      If CStr(rsu.Fields(i)) = fI Then
       Debug.Print rst!Tn, rsu.Fields(i).name, rsu.Fields(i)
       gefu = True
      End If
     End If
    Next i
    If gefu Then GoTo weiter
   rsu.MoveNext
  Loop
weiter:
  End If
  rst.MoveNext
 Loop ' While Not rst.EOF
 Debug.Print "Fertig"
End Function ' suchfi(pNr&, fI$)

Public Function suchal(fI$, obrlike%)
 Dim altt$, j&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
#If True Then
 rst.Open "SELECT TABLE_NAME tn, GROUP_CONCAT(CONCAT('CAST(',column_name,' AS CHAR)" & IIf(obrlike, " RLIKE ", "=") & "''" & fI & "''') SEPARATOR ' OR ') cn FROM information_schema.columns c WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_NAME NOT RLIKE 'fsurogat' group by table_name" _
 , MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
'  Debug.Print rst!Tn, rst!Cn
  DoEvents
  Set rsu = Nothing
  rsu.Open "SELECT * from `" & rst!Tn & "` WHERE " & rst!Cn, MOCon, adOpenStatic, adLockReadOnly
  If Not rsu.EOF Then
   Debug.Print rst!Tn & ": " & rsu.Fields(0).name & ": " & rsu.Fields(0)
  Else
'   Debug.Print "nichts gefunden in: " & rst!Tn
  End If
  rst.MoveNext
 Loop
#Else
 rst.Open "SELECT TABLE_NAME tn FROM information_schema.tables t WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_NAME  NOT RLIKE 'fsurogat'" & vbCrLf & _
 "and not exists (select 0 from information_schema.columns where  table_catalog='def' AND table_schema='medoff' AND TABLE_NAME =t.table_name and column_Name = 'fpatnr')" & vbCrLf & _
 "and table_name not in ('datafile')" _
 , MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
  If rst!Tn <> altt Then
   Set rsu = Nothing
   rsu.Open "SELECT * from `" & rst!Tn & "`", MOCon, adOpenStatic, adLockReadOnly
   j = 0
   Do While Not rsu.EOF
    j = j + 1
    Debug.Print rst!Tn, j
    DoEvents
'    If j = 8374 Then Stop
    For i = 0 To rsu.Fields.COUNT - 1
     If Not (IsNull(rsu.Fields(i))) Then
      If CStr(rsu.Fields(i)) = fI Then
       Debug.Print rsu.Fields(i).name
      End If
     End If
    Next i
   rsu.MoveNext
  Loop
  End If
  rst.MoveNext
 Loop ' While Not rst.EOF
#End If
 Debug.Print "Fertig"
End Function


#If False Then
' hier in SQL:
CREATE DEFINER=`medoff`@`%` PROCEDURE `procmpatstamm`(
    IN `PNr` INT
)
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'analysiert FMemo aus Patstamm'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(1024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmpatstamm (
    patnr INT(11) UNSIGNED NOT NULL,
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(patnr,znr),
    Key zuloe(patnr, enr)
);
DELETE FROM tmpmpatstamm WHERE pnr=0 OR patnr=pnr;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT fsurogat fpatnr, fmemo FROM patstamm WHERE (pnr=0 OR fsurogat=pnr) AND fmemo IS NOT NULL) DO
 SET pos=1;
 SET ie=0;
 TRUNCATE tmpeb;
labt:  Loop
  SET obdruck=0;
  SET tlen=CONV(HEX(CONCAT(MID(r.fmemo,pos+1,1),MID(r.fmemo,pos,1))),16,10);
  if pos=1 then SET gl=tlen; SET MAX=gl; END if;
  SET altmax=MAX;
  SET aktMAX=tlen+pos+1; -- aktuell angegebene Länge aus dem und dem nä Byte
  IF aktmax BETWEEN 0 AND gl AND -- wenn die angegebene Länge vertretbar
    (( CONV(HEX(MID(r.fmemo,aktMax,1)),16,10)=0 AND -- und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen
     aktmax<=MAX) OR pos>MAX) then -- wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde
    SET altie=ie; -- letzte Ebene
    if aktmax<=MAX then SET ie=ie+1; END if; -- im ersten Fall wird die Ebene erhöht
    if pos>MAX then -- im zweiten Fall wird zurückgegriffen
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmpatstamm WHERE patnr=r.fpatnr AND mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmpatstamm WHERE patnr=r.fpatnr AND mx>=aktMAX);
     END if;
    if ie<=altie then -- wenn die Ebene nicht erhöht worden ist
     if ie<altie then -- wenn sie vielmehr reduziert wurde
      DELETE FROM tmpeb WHERE nr>ie; -- dann werden die höheren Einträge wieder gelöscht
     END if;
     UPDATE tmpeb SET wert=wert+1 WHERE nr=ie; -- dann wird die Zählung der akt. Ebene erhöht
    Else
     INSERT INTO tmpeb(nr,wert) VALUES(ie,1); -- sonst muss ein neuer Eintrag für die hohe Ebene erstellt werden
    END if;
    SET MAX=aktmax; -- neue vorbestehende Länge
    SET obdruck=1; -- und drucken
  END if;
  IF obdruck=1
--     OR pos=1 -- zum Debuggen der ersten Gesamt-Zeile
--   OR TRUE -- zum Debuggen aller Zeilen
   then
    SET txt=MID(r.fmemo,pos+2,tlen);
    If txt <> Repeat(Chr(0), tlen) Then
     INSERT INTO tmpmpatstamm(patnr,znr,mx,ebn,enr,TEXT) VALUES(r.fpatnr,pos,max,ie,
       (SELECT GROUP_CONCAT(wert ORDER BY nr SEPARATOR '.') FROM tmpeb),
--    CONCAT(MID(r.fmemo,pos,1),'|',HEX(MID(r.fmemo,pos,1)),'|',CONV(HEX(MID(r.fmemo,pos,1)),16,10),'|',CONV(HEX(CONCAT(MID(r.fmemo,pos+1,1),MID(r.fmemo,pos,1))),16,10),'|','aktMax:',aktmax, '|','Max:',altMAX,'|','Laenge: ',LENGTH(txt),'|','Endbyte:',COALESCE(CONV(HEX(MID(r.fmemo,aktMax,1)),16,10),'-'),
--     CONCAT(IF(obdruck<>1,'- ',''),'"',
      txt
--    ,'"'))
       );
     Else
      SET pos=pos+tlen; -- 0-Strings nicht auffieseln
     END if;
  END IF; -- obdruck=1
  SET pos=pos+1;
  if pos>=gl then LEAVE labt; END if;
 END loop labt;
END FOR laba;
 -- folende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmpatstamm WHERE (pnr=0 OR patnr=pnr) AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpmpatstamm i WHERE i.patnr=tmpmpatstamm.patnr AND INSTR(i.enr,CONCAT(tmpmpatstamm.enr,'.'))=1 AND i.enr<>tmpmpatstamm.enr);
-- COMMIT;
-- LEAVE tp;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmpatfall`(
    IN `pnr` INT
)
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FMemo aus patfall in die Tabelle tmpfmemo auslesen'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(1024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmpatfall (
    patnr INT(11) UNSIGNED NOT null,
    fsur INT(11) UNSIGNED NOT NULL,
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(patnr,fsur,znr),
    Key zuloe(patnr, fsur, enr)
);
DELETE FROM tmpmpatfall WHERE pnr=0 OR patnr=pnr;
CREATE TEMPORARY TABLE if NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT fsurogat fsur, fpatnr, fmemo FROM patfall WHERE (pnr=0 OR fpatnr=pnr) AND fmemo IS NOT NULL) DO
 SET pos=1;
 SET ie=0;
 TRUNCATE tmpeb;
labt:  Loop
  SET obdruck=0;
  SET tlen=CONV(HEX(CONCAT(MID(r.fmemo,pos+1,1),MID(r.fmemo,pos,1))),16,10);
  if pos=1 then SET gl=tlen; SET MAX=gl; END if;
  SET altmax=MAX;
  SET aktMAX=tlen+pos+1; -- aktuell angegebene Länge aus dem und dem nä Byte
  IF aktmax BETWEEN 0 AND gl AND -- wenn die angegebene Länge vertretbar
    (( CONV(HEX(MID(r.fmemo,aktMax,1)),16,10)=0 AND -- und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen
     aktmax<=MAX) OR pos>MAX) then -- wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde
    SET altie=ie; -- letzte Ebene
    if aktmax<=MAX then SET ie=ie+1; END if; -- im ersten Fall wird die Ebene erhöht
    if pos>MAX then -- im zweiten Fall wird zurückgegriffen
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmpatfall WHERE patnr=r.fpatnr and fsur=r.fsur AND mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmpatfall WHERE patnr=r.fpatnr and fsur=r.fsur AND mx>=aktMAX);
     END if;
    if ie<=altie then -- wenn die Ebene nicht erhöht worden ist
     if ie<altie then -- wenn sie vielmehr reduziert wurde
      DELETE FROM tmpeb WHERE nr>ie; -- dann werden die höheren Einträge wieder gelöscht
     END if;
     UPDATE tmpeb SET wert=wert+1 WHERE nr=ie; -- dann wird die Zählung der akt. Ebene erhöht
    Else
     INSERT INTO tmpeb(nr,wert) VALUES(ie,1); -- sonst muss ein neuer Eintrag für die hohe Ebene erstellt werden
    END if;
    SET MAX=aktmax; -- neue vorbestehende Länge
    SET obdruck=1; -- und drucken
  END if;
  IF obdruck=1
--   OR pos=1 -- zum Debuggen der ersten Gesamt-Zeile
--   OR TRUE -- zum Debuggen aller Zeilen
   then
    SET txt=MID(r.fmemo,pos+2,tlen);
    If txt <> Repeat(Chr(0), tlen) Then
     INSERT INTO tmpmpatfall(patnr,fsur,znr,mx,ebn,enr,TEXT) VALUES(r.fpatnr,r.fsur,pos,max,ie,
       (SELECT GROUP_CONCAT(wert ORDER BY nr SEPARATOR '.') FROM tmpeb),
--    CONCAT(MID(r.fmemo,pos,1),'|',HEX(MID(r.fmemo,pos,1)),'|',CONV(HEX(MID(r.fmemo,pos,1)),16,10),'|',CONV(HEX(CONCAT(MID(r.fmemo,pos+1,1),MID(r.fmemo,pos,1))),16,10),'|','aktMax:',aktmax, '|','Max:',altMAX,'|','Laenge: ',LENGTH(txt),'|','Endbyte:',COALESCE(CONV(HEX(MID(r.fmemo,aktMax,1)),16,10),'-'),
--     CONCAT(IF(obdruck<>1,'- ',''),'"',
      txt
--    ,'"'))
       );
     ELSE SET pos=pos+tlen; -- 0-Strings nicht auffieseln
     END if;
  END IF; -- obdruck=1
  SET pos=pos+1;
  if pos>=gl then LEAVE labt; END if;
 END loop labt;
END FOR laba;
 -- folende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmpatfall WHERE (pnr=0 OR patnr=pnr) AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpmpatfall i WHERE i.patnr=tmpmpatfall.patnr AND i.fsur=tmpmpatfall.fsur AND INSTR(i.enr,CONCAT(tmpmpatfall.enr,'.'))=1 AND i.enr<>tmpmpatfall.enr);
-- folgende Variante bringt, aus der Prozedur aufgerufen, mariadb zum Absturz:
-- DELETE a FROM tmpmpatfall a LEFT JOIN (SELECT LEAD(enr) OVER(PARTITION BY patnr,fsur ORDER BY znr) nenr, patnr,fsur,znr FROM tmpmpatfall) i USING (patnr,fsur,znr) WHERE INSTR(i.nenr,a.enr)=1 AND i.nenr<>a.enr;
-- COMMIT;
-- LEAVE tp;
End

#End If
