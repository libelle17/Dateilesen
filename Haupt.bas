Attribute VB_Name = "Haupt"
Option Explicit
#Const ohneDAO = True
Public Const vNS$ = vbNullString
Dim DateiNachzuholen$, DebugDatei$, QuelleMDB$
'Const DateiNachzuholen$ = "c:\u\anamnese\Nachholen.txt"
'Const DebugDatei$ = "c:\u\anamnese\debug.txt"
'Const QuelleMDB$ = "c:\u\anamnese\quelle.mdb"
'Public rs As New ADODB.Recordset
Public FSO As New FileSystemObject
Public imAufbauLese As Boolean
Public imAufbauDialog As Boolean
Public imAufbauSonstige As Boolean
Public obStart% ' ob Startvorgang (da in ConstrFestleg Datei÷ffnen-Dialog f¸r die MDB-Datei zeigen und dbcn nicht verbinden)
Public RegStelle$
Public Const RegWurzel$ = "Software\GSProducts\"
Public FPos& ' Fehlerposition
'Public dlg AS Dialog
Const TbZ1% = 26 ' 27 (therarten)
Const TbZ2% = TbZ1 + 8
'Public Lese1
Public AllePat% ' alle Patienten sind in der Datei
Public obVorber% ' ob Formularvorbereitung
Dim tbn$(TbZ2), Tbk$(TbZ2), noDup%(TbZ2)
Public ZielDbS$
Public BrichAb% ' bricht gleich ab
Public ProgL‰uft%
Public KRein% ' Kodierrichtlinien eingeschaltet
Public T1 As Date, T1a As Date
Dim Ausschl(7) As String * 1
Type Vorkommen
 Beg As Integer
 END As Integer
End Type
Public Enum KeyModeConstants
  kmName
  kmPath
End Enum
Dim ForeignKAus0%, ForeignKAus1%, rAf&
Public KVƒDatei1$, BriefZiel$, AutoBriefZiel$, AutoBriefProtok$
'Public Const KVƒDatei1$ = AnamneseVerZeichnis1 + "KV-ƒrzte neu.mdb"
'Public Const BriefZiel$ = pVerz
'Public Const AutoBriefZiel$ = BriefZiel & "unkorrigiert\"
'Public Const AutoBriefProtok$ = BriefZiel & "zufaxen.txt"
Public lies As Lese
Public oboffenlassen%
#If False Then
Dim oPatient As New Risk
#End If
Public aktOSV As WindowsVersion
Public Pausenl‰nge%
Public userprof$

Type rstype
 Quartal As String
 tage As String
 kassenpat As String
End Type

Public Function lQAnfuEnd$(frist$, Optional nurDatum%)
' letztes Quartal Anfang und Ende, f¸r Format datetime (mit Zeit, z.B. ' 2012-07-01 AND 2012-10-01): bei Format date falsches Ergebnis!
' nurdatum = einen Tag abziehen, da Vergleichsparameter im Format '2012-10-21' statt '2012-10-21 17:21'

'lQAnfuEnd = " CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01') AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')" & IIf(nurDatum <> 0, "-INTERVAL 1 day", "")
 lQAnfuEnd = tuAbstand(frist, 0)
End Function ' lqanfuend(frist$)

' anstat: qbeg(SUBDATE(qanf(),274)) AND qend()
Public Function Khtsfl(Optional frist$)
 Khtsfl = tuAbstand(frist, 274)
End Function ' Khtsfl(Optional frist$)

' qanf() AND qende(adddate(qend(),274)
Public Function KhtsflZuk(Optional frist$)
 KhtsflZuk = tuAbstand(frist, 0, 274)
End Function ' KhtsflZuk(Optional frist$)

' aufgerufen in: lQAnfuEnd, Khtsfl, KhtsflZuk
Public Function tuAbstand(frist$, Abst%, Optional Zuk% = 0)
 Dim qa As Date, qe As Date
 If frist = "" Then
  If Versp‰tung = "" Then
   Call holFrist
  End If
  frist = Versp‰tung
 End If
 qa = QAnf(QuartalStr(Now() - frist - Abst))
 qe = QEnd(QuartalStr(Now() - frist + Zuk))
 tuAbstand = Year(qa) & Right$("0" & Month(qa), 2) & Right$("0" & Day(qa), 2) & " AND " & Year(qe) & Right$("0" & Month(qe), 2) & Right$("0" & Day(qe), 2) & "235959"
End Function ' tuAbstand

' aufgerufen in ZeigSQL (vielfach)
Public Function qtAnf$(frist$)
 Dim qa As Date
 Dim Abst%
 If frist = "" Then
  If Versp‰tung = "" Then
   Call holFrist
  End If
  frist = Versp‰tung
 End If
 qa = QAnf(QuartalStr(Now() - frist - Abst))
 qtAnf = Year(qa) & Right$("0" & Month(qa), 2) & Right$("0" & Day(qa), 2)
End Function ' qtAnf$(frist$)

Public Function qtEnd$(frist$)
 Dim qe As Date
 Dim Abst%
 If frist = "" Then
  If Versp‰tung = "" Then
   Call holFrist
  End If
  frist = Versp‰tung
 End If
 qe = QEnd(QuartalStr(Now() - frist - Abst))
 qtEnd = Year(qe) & Right$("0" & Month(qe), 2) & Right$("0" & Day(qe), 2) & "235959"
End Function ' qtEnd$(frist$)

' 12.10.19: steht auch im Postbeamtenvertrag https://www.kvb.de/praxis/alternative-versorgungsformen/dmp/rechtsgrundlagen/: "KVB-RQ-Diabetesvereinbarung-Postbeamtenkrankenkasse.pdf"
Public Function motsql() ' 13.4.21 ktag liefert falsche Ergebnisse, wenn auch inhalt gew‰hlt
 motsql = "SELECT f.Pat_ID,gesname(f.pat_id)PName, PatAlter(f.pat_id) PAlter, ICD, DATE(e.zeitpunkt) Zeitpunkt, e.art AS Art, e.inhalt Inhalt FROM `aktfv` f " & vbCrLf & _
       "LEFT JOIN `kassenliste` kl ON f.vknr = kl.vknr AND f.ik=kl.ik " & vbCrLf & _
       "LEFT JOIN `diagnosen` d ON f.pat_id = d.pat_id AND icd REGEXP '^E1[0-4]' AND diagsicherheit IN ('G',' ') " & vbCrLf & _
       "LEFT JOIN `leistungen` l ON f.pat_id = l.pat_id AND l.leistung IN (SELECT leistung FROM genehmigungen g WHERE obschulung<>0 AND wert >19) " & vbCrLf & _
                        "AND qbeg(l.zeitpunkt) + INTERVAL 350 day > qanf() " & vbCrLf & _
       "LEFT JOIN `eintraege` e ON f.pat_id = e.pat_id AND (e.art IN ('andm','andm2'" & artSpezBerat & ")) " & vbCrLf & _
                        "and e.zeitpunkt BETWEEN " & lQAnfuEnd$(Versp‰tung) & " " & vbCrLf & _
       "LEFT JOIN leistungen l0 ON l0.pat_id = f.pat_id AND l0.leistung IN ('92278','92282') AND l0.zeitpunkt>SUBDATE(qende(e.zeitpunkt),365) " & vbCrLf & _
       "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
       "WHERE NOT ISNULL(id1) AND ISNULL(l.leistung) AND ISNULL(l0.leistung) AND kl.kateg NOT IN ('LKK','SHV') AND (dmpklass<>2 OR kl.kateg='PBe') AND dmpklass IN (0,2,3) " & vbCrLf & _
       "GROUP BY f.pat_id, art, e.zeitpunkt " & vbCrLf & _
       "ORDER BY MID(icd,2,1), pat_id, e.zeitpunkt " ' STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL " & FristS & " DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d') AND SUBDATE(adddate(STR_TO_DATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & FristS & " DAY)),'/',((month(SUBDATE(NOW(),INTERVAL 20 DAY))-1) div 3)*3+1,'/1'),'%Y/%m/%d'),INTERVAL 3 MONTH),INTERVAL 1 day
' AND COALESCE(f6010,0)=0
End Function ' motsql()

'Public FUNCTION DefDB$(DBCn AS ADODB.Connection)
' DefDB = DBCn.DefaultDatabase
' ON Error Resume Next
' IF LenB(DefDB) = 0 THEN DefDB = DBCn.Properties("Data Source Name").Value
' IF LenB(DefDB) = 0 THEN DefDB = "quelle"
'End FUNCTION ' DefDB


'Public cn As New ADODB.Connection
'Type TabInd
' Ctl AS String
' TI As Integer
'End Type
Function machAbK¸$()
 Dim i%, j%, k%
 tbn(0) = "Namen"
 tbn(1) = "Faelle"
 tbn(2) = "AU"
 tbn(3) = "Briefe"
 tbn(4) = "Diagnosen"
 tbn(5) = "Dokumente"
 tbn(6) = "eintraege"
 tbn(7) = "FormInhaltForm_Abk"
 tbn(8) = "Formulare"
 tbn(9) = "FormInhKopf" 'altNeu"
 tbn(10) = "FormInhFeld"
 tbn(11) = "KHEinweis"
 tbn(12) = "LBAnforderungen"
 tbn(13) = "laborneu"
 tbn(14) = "Leistungen"
 tbn(15) = "MedPlan"
 tbn(16) = "Rezepteintraege"
 tbn(17) = "RR"
 tbn(18) = "KVNrUe"
 tbn(19) = "unbekannte kennungen"
 tbn(20) = "dmpreihe"
 tbn(21) = "desktop"
 tbn(22) = "usdm"
 tbn(23) = "fuss"
 tbn(24) = "ulcus"
 tbn(25) = "vkgd"
 tbn(26) = "sws"
' tbn(26) = "therarten"
 tbn(TbZ1 + 1) = "laborxsaetze"
 tbn(TbZ1 + 2) = "laborxeingel"
 tbn(TbZ1 + 3) = "laborxus"
 tbn(TbZ1 + 4) = "laborxbakt"
 tbn(TbZ1 + 5) = "laborxwert"
 tbn(TbZ1 + 6) = "laborxleist"
 tbn(TbZ1 + 7) = "liuez" ' zum Vorladen, noch nicht implementiert 1.11.14
 tbn(TbZ1 + 8) = "anamnesebogen"
 
 noDup(TbZ1 + 5) = True
 
' Tbn(17) = "FormInhaltForm_abk"
' Tbn(18) = "FormInhaltFeldinh"
' Tbn(19) = "FormInhaltFeld"
 For i = 0 To TbZ2
  Select Case tbn(i)
   Case "FormInhaltForm_Abk": Tbk(i) = "Fi"
'   Case "Formulare": Tbk(i) = "Fo"
   Case "laborxsaetze": Tbk(i) = "Ls"
'   Case "laborxbakt":   Tbk(i) = "Lo"
   Case "laborxleist":  Tbk(i) = "LL"
   Case "laborxwert":   Tbk(i) = "Lw"
   Case "laborxus":     Tbk(i) = "Lu"
   Case "laborxeingel": Tbk(i) = "Lg"
'   Case "laborxpneu": Tbk(i) = "Lg"
'   Case "laborxplab": Tbk(i) = "Lg"
'   Case "laborxpnb": Tbk(i) = "Lg"
   Case "anamnesebogen": Tbk(i) = "Ana"
   Case Else
    For k = 2 To Len(tbn(i))
     Tbk(i) = UCase$(Left$(tbn(i), 1)) + LCase$(Mid$(tbn(i), k, 1))
     For j = 0 To i - 1
      If Tbk(j) = Tbk(i) Then GoTo hammaschon
     Next j
     Exit For
hammaschon:
    Next k
  End Select
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in Zeichensatz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machAbK¸$()

' ADO-Datentypen
'# DataTypeEnum
'   adByRef = empty
'   adVector = empty          wird mit logischem OR-Argument mit anderem Typ verbunden
'   adEmpty = 0
'   adSMALLINT = 2            2 Bytes , Ganzzahl mit Vorzeichen
'   adInteger = 3             4 Bytes, mit Vorzeichen
'   adSingle = 4              4 , Gleitpunktwert
'   adDouble = 5              8 , Gleitpunktwert
'   adCurrency = 6
'   adDate = 7               als Double gespeichert, Tage seit dem 30.12.1899
'   adBSTR = 8              null-terminierten Character-String (Unicode)
'   adIDispatch = 9         Zeiger auf IDispatch-Schnittstelle in OLE-Objekt
'   adError = 10              32 Bit
'   adBoolean = 11           Boolscher Wert
'   adVariant = 12           Variant-Wert f¸r Automatisierung
'   adIUnknown = 13          Zeiger auf IUnknown-Schnittstelle in Ole-Objekt
'   adDecimal = 14
'   adTinyInt = 16           1 , Ganzzahl mit Vorzeichen
'   adUNSIGNEDTinyInt = 17   1 Byte
'   adUNSIGNEDSMALLINT = 18  2 Bytes
'   adUNSIGNEDInt = 19       4 Bytes
'   adBigInt = 20            8 Byte Ganzzahl mit Vorzeichen
'   adUNSIGNEDBigInt = 21    8 Bytes
'   adFileTime = 64
'   adGUID = 72 '# e.g. {E5D50A9B-33D2-11D3-AAB3-00104BA31425}
'   adBinary = 128           Bin‰rwert
'   adChar = 129           String
'   adWChar = 130         mit NullZeichen endende Zeichenfolge
'   adNumeric = 131        exakter numerischer Wert mit fester Genauigkeit und Skalierung
'   adUserDefined = 132      benutzerdefinierte Variable
'   adDBDate = 133           jjjjmmtt
'   adDBTime = 134           hhmmss
'   adDBTimeStamp = 135      jjjjmmtthhmmss plus Bruch in Milliardstel
'    adChapter    = 136       4 Byte
'    adPropVariant      = 138
'    adVarNumeric = 139
'   adVarChar = 200        ' String-Wert (nur Parameter-Objekt)
'   adLongVarChar = 201    ' langer String-Wert (nur Parameter-Objekt)
'   adVarWChar = 202       ' mit 0-Zeichen endende Unicode-Zeichenfolge
'   adLongVarWChar = 203   ' langer, mit 0-Zeichen endender Zeichenfolgenwert (nur Parameter-Objekt)
'   adVarBinary = 204      ' Bin‰rwert (nur Parameter-Objekt)
'   adLongVarBinary = 205 ' langer Bin‰rwert (nur Parameter-Objekt)
'   adArray = 8192        ' mit logischem OR verbinden

'SQL-JET
'BINARY:  1 Byte pro Zeichen
'BIT:     1 Byte
'BYTE:    1 Byte
'COUNTER: 4 Bytes
'CURRENCY:8 Bytes
'DATETIME:8 Bytes
'GUID     128 Bit
'SINGLE   4 Bytes
'DOUBLE   8 Bytes
'SHORT    2 Bytes
'LONG     4 Bytes
'LONGTEXT 1 Byte pro Zeichen
'LONGBINARY nach Bedarf, f¸r OLE-Objekte
'TEXT     1 Byte pro Zeichen

'VB-Datentypen:
'Byte     1 Byte
'Boolean  2 Bytes
'Integer  2 Bytes
'Long     4 Bytes
'Single   4 Bytes
'Double   8 Bytes
'Currency 8 Bytes
'Decimal  14 Bytes
'Date     8 Bytes
'Object   4 Bytes
'String (variabel)     10 Bytes + Zeichenfolgenl‰nge
'String (fest)         Zeichenfolgenl‰nge
'Variant (mit Zahlen)  16 Bytes
'Variant (mit Zeichen) 22 Bytes
'Datenfelder: 20 Bytes + 4 Bytes /Datenfelddef

'MySQL:
'BIT       1 Bit
'TINYINT   1 Byte
'BOOL      = Tinyint(1)
'SMALLINT  2 Bytes
'MEDIUMINT 3 Bytes
'INT       4 Bytes
'BIGINT    8 Bytes
'FLOAT     8 Bytes (0-24)
'DOUBLE    16 Bytes (25-53)
'DECIMAL   10 Bytes, falls nicht anders angegeben
'DATE,TIME,DATETIME,TIMESTAMP,YEAR
'BLOB usw.

Function doMacheTypen(Tabelle$) ' aufgerufen aus MacheTypen
 Dim Tabl$, art$, rsAdSc As New ADODB.Recordset
 Dim TDesc As New ADODB.Recordset
 Dim lenge&, obZusatzFeld%
 On Error GoTo fehler
 syscmd 4, "doMacheTypen (" & Tabelle & ") ..."
 Tabl = LCase$(Tabelle)
 Print #257, ""
 Print #257, "Public type " + REPLACE$(Tabl, " ", "_")
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do While Not rsAdSc.EOF
  DoEvents
'  Debug.Print rsAdSc.Fields(2), colw(rsadsc!column_name), rsAdSc!column_propid, rsAdSc!column_flags, rsAdSc!data_type, rsAdSc!type_guid, rsAdSc!character_maximum_length, rsAdSc!numeric_scale, rsAdSc!numeric_precision, rsAdSc!Description
  Dim colN$
  colN = colW(rsAdSc!COLUMN_NAME)
  art = " " + colN + " AS "
  Select Case rsAdSc!data_type
' numerisch:
   Case "tinyint", 16, 17:           art = art + "byte"
'   Case 11:              Art = Art + "boolean" kˆnnte in String "Wahr" o.‰. umgewandelt werden
   Case "bit", "smallint", 2, 11, 18:        art = art + "integer" ' z.T.: "boolean"
   Case "int", 3, 19:            art = art + "long" '20+21 -> double
   Case "float", 4:                art = art + "single"
   Case "decimal", "double", 5, 20, 21, 131, 139: art = art + "double"
   Case 6:                art = art + "currency"
   Case 14:               art = art + "decimal"
   Case 10: art = art + "variant" '(error)
'   Case 9, 72:            art = art + "object" lieber nicht
' Datum:
   Case "datetime", "date", "timestamp", "time", 7, 64, 133, 134, 135: art = art + "date"
' String
   Case "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205: art = art + "variant"
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203: art = art + "string"
   Case Else: art = art + "variant"
  End Select
  art = art & " '" & colW(rsAdSc!COLUMN_NAME) & " " & rsAdSc!data_type
  If Forms(0).obMySQL Then
   Set TDesc = Nothing
   myFrag TDesc, "SHOW FULL COLUMNS FROM `" & LCase$(Tabelle) & "` WHERE field = '" & rsAdSc!COLUMN_NAME & "'"
   If Not TDesc.EOF Then
    art = art & " '" & TDesc!Comment
   End If
  Else
   art = art & " '" & rsAdSc!Description
  End If
  Print #257, art
  If colN Like "Feld*VW" Or colN = "LangtextVW" Or colN = "KommentarVW" Then
   Print #257, " " + Left$(colN, Len(colN) - 2) + " AS string"
  End If
  rsAdSc.MoveNext
 Loop
#If False Then
' IF lies.obmysql THEN
  Set rsAdo = myEFrag("SHOW columns FROM `" & Tabl & "`")
' END IF
 Do While Not rsAdo.EOF
  art = " " + rsAdo.Fields(0) + " AS "
  If rsAdo.Fields(1) Like "varchar*" Then
   lenge = Len(rsAdo.Fields(1))
   art = art + "string" '  * " + LEFT(mid$(rsAdo.Fields(1), 9), lenge - 9)
  ElseIf rsAdo.Fields(1) Like "char*" Then
   lenge = Len(rsAdo.Fields(1))
   art = art + "string" ' * " + LEFT(mid$(rsAdo.Fields(1), 6), lenge - 6)
  ElseIf rsAdo.Fields(1) Like "int(10)*" Or rsAdo.Fields(1) Like "int(4)*" Or rsAdo.Fields(1) Like "integer" Or rsAdo.Fields(1) Like "int(11)*" Then
   art = art + "long"
  ElseIf rsAdo.Fields(1) Like "date*" Then
    art = art + "date"
  ElseIf rsAdo.Fields(1) = "longtext" Then
    art = art + "string"
  ElseIf rsAdo.Fields(1) = "tinyint(1)" Then
    art = art + "boolean"
  ElseIf rsAdo.Fields(1) = "SMALLINT(6)" Or rsAdo.Fields(1) = "tinyint(4)" Then
   art = art + "integer"
  ElseIf rsAdo.Fields(1) = "bit(8)" Or rsAdo.Fields(1) = "int(1) UNSIGNED" Then
   art = art + "byte"
  Else
    Err.Raise 999, , "Fehler in doMacheTypen: rsAdo.Fields(1): " & rsAdo.Fields(1) & "(Datentyp nicht vorgesehen)"
  End If
  Print #257, art
  If rsAdo.Fields(0) Like "Feld*VW" Then
   Print #257, " " + Left$(rsAdo.Fields(0), Len(rsAdo.Fields(0)) - 2) + " AS string"
  End If
  rsAdo.Move 1
 Loop
#End If
 Print #257, "end type"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doMacheTypen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMacheTypen

Function obAutoIncr%(Tabl$, rs As ADODB.Recordset)
 Dim test%, Spalt$
 On Error GoTo fehler
 If lies.obMySQL Then
#If True Then
 Dim rsDir As New ADODB.Recordset, fI%, pi%
 rsDir.CursorLocation = adUseServer
' For fi = 0 To rsDir.Fields.Count - 1
'  For pi = 0 To rsDir.Fields(fi).Properties.Count - 1
'   Debug.Print rsDir.Fields(fi).Properties(pi).Name, rsDir.Fields(fi).Properties(pi).Value
   Dim colname$
   colname = rs!COLUMN_NAME
   myFrag rsDir, "SELECT `" & colname & "` FROM `" & Tabl & "` LIMIT 1"
   If Not rsDir.BOF Then
    If rsDir.Fields(colname).Properties("ISAUTOINCREMENT") Then
     obAutoIncr = True
    End If
   End If
'  Next pi
' Next fi
#Else
' 29.7.2010: geht nicht
  If rs!data_type = 3 And rs!column_flags = 120 And rs!column_hasdefault = True And rs!column_default = 0 Then
   If rs!is_nullable And rs!COLUMN_NAME <> "dmpklass" And rs!COLUMN_NAME <> "dmpkhkklass" And rs!COLUMN_NAME <> "dmpcopdklass" Then
    Stop
   Else
    obAutoIncr = True
   End If
  End If
#End If
 Else
  Static catx As New ADOX.Catalog
  If catx.ActiveConnection Is Nothing Then
   catx.ActiveConnection = DBCnS ' DBCn.ConnectionString
  End If
  On Error Resume Next
  Spalt = rs!COLUMN_NAME
  test = catx.Tables(Tabl).Columns(Spalt).Properties.COUNT
  If Err.Number <> 0 Then
   Exit Function
  End If
  On Error GoTo fehler
  If catx.Tables(Tabl).Columns(Spalt).Properties.COUNT > 0 Then
   If catx.Tables(Tabl).Columns(Spalt).Properties("AutoIncrement") Then
    obAutoIncr = True
   End If
  End If
 End If ' lies.obmysql
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in obAutoIncr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obAutoIncr

Function colW$(col$)
 colW = REPLACE$(REPLACE$(REPLACE$(col, " ", "_"), "?", ""), "-", "_")
End Function ' colW$(col$)

Function doMachSQL0$(TName$, NobAI%) ' NobAI = es existiert eine Zeile ohne Autoincrement
 On Error GoTo fehler
 Dim Tabl$, art$, VStr$, lenge&, rsAdSc As New ADODB.Recordset, SpZ%
 Dim trennz$
 Dim obHierKAI%, colN$ ' ob hier kein AutoIncrement
 NobAI = 0
 Tabl = LCase$(TName)
' IF Tabl = "forminhaltform_abk" THEN
' IF DBCn.State = 0 THEN Call DBCnOpen(CSStr)
' Call myEFrag("use quelle1;")
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do
  If rsAdSc.BOF Then Exit Do
  If rsAdSc.EOF Then
   trennz = ")"
  Else
   obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
   If obHierKAI Then trennz = ","
  End If
  If rsAdSc.EOF Or obHierKAI Then
   If LenB(doMachSQL0) = 0 Then
     doMachSQL0 = " ("
   Else
     doMachSQL0 = doMachSQL0 & trennz
   End If
   If Not rsAdSc.EOF Then If SpZ Mod 10 = 3 Then doMachSQL0 = doMachSQL0 + """ & _" + vbCrLf + "     """
   SpZ = SpZ + 1
   If rsAdSc.EOF Then
     Exit Do
   Else
    doMachSQL0 = doMachSQL0 & colW(rsAdSc!COLUMN_NAME)
    NobAI = True
   End If ' rsAdSc.EOF Then else
  End If ' rsAdSc.EOF Or obHierKAI Then
  rsAdSc.Move 1
 Loop
 
#If False Then
 Set rsAdo = myEFrag("SHOW columns FROM `" & Tabl & "`")
 Do While Not rsAdo.EOF
  doMachSQL0 = doMachSQL0 + rsAdo.Fields(0)
  doMachSQL0 = doMachSQL0 & ","
  rsAdo.Move 1
 Loop
#End If
' doMachSQL0 = LEFT(doMachSQL0, Len(doMachSQL0) - 1) + ")"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doMachSQL0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMachSQL0(ti%, sql$)

#If False Then
Function doMachSQL1$(TI%, ptxt$)
 Dim Tabl$, VStr$, umbrZahl%, obBegonnen%
 Dim trennz$, tza$, tze1$, tze2$
 Dim obHierKAI%, rDT$ ' ob hier kein AutoIncrement / Column-Name / Recordset-Datatype
 Dim rsAdSc As New ADODB.Recordset
 On Error GoTo fehler
 umbrZahl = 1
 Tabl = LCase$(tbn(TI))
 
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do
  If rsAdSc.BOF Then Exit Do
  If rsAdSc.EOF Then
   trennz = ")"""
  Else
   obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
   If obHierKAI Then trennz = ","
  End If
  If rsAdSc.EOF Or obHierKAI Then
   If Not obBegonnen Then
    obBegonnen = True
   Else
    ptxt = ptxt & tze2 & trennz
   End If
   If rsAdSc.EOF Then
    Exit Do
   Else
    VStr = "r" & Tbk(TI) & "(i)." & colW(rsAdSc!COLUMN_NAME)
    Select Case rsAdSc!data_type
     Case "int", "tinyint", "smallint", "bigint", "decimal", "double", "float", 2, 3, 4, 5, 16, 17, 18, 19, 20, 21, 131, 139, 6, 14: tza = """ & ": tze1 = " & ": tze2 = """" ' Zahlen
     Case "bit", 11, 128:                            tza = """ & cstr(-(": tze1 = "<>0)) & ": tze2 = """" 'tza = """ & cstr(cint(": tze1 = ")) & ": tze2 = """" ' VStr = "cStr(cint(" & VStr & "))" & trennz ' int, z.T.: "boolean" / Zuordnung von 128 nicht ganz sicher
     Case "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205: tza = "'"" & ": tze1 = " & ": tze2 = """'" 'VStr = """'"" + " & VStr & " + ""'" & trennz & """" ' Variant
     Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203:    tza = "'"" & ": tze1 = " & ": tze2 = """'" 'VStr = """'"" + " & VStr & " + ""'" & trennz & """"
     Case "datetime", "date", "timestamp", "time", 7, 64, 133, 134, 135:               tza = """ & DatFor_k(": tze1 = ") & ": tze2 = """" 'VStr = " DatFor_k(" & VStr & ")" & trennz & """"
     Case Else: Err.Raise 999, , "Fehler in doMachSQL1: rsAdSc!data_type: " & rsAdSc!data_type & " (Datentyp nicht vorgesehen)"
    End Select
    Select Case rsAdSc!data_type
     Case "decimal", "double", "float": VStr = REPLACE$(VStr, ",", ".")
    End Select
    ptxt = ptxt & tza
    If Len(ptxt) / umbrZahl > 150 Then
     umbrZahl = umbrZahl + 1
     If umbrZahl Mod 15 = 1 Then
      ptxt = ptxt & """" & vbCrLf & "  sql = sql & "
     Else
      ptxt = ptxt & " _" & vbCrLf & "   "
     End If
    End If
    If Tabl = "namen" And InStrB(LCase$(VStr), "aktzeit") <> 0 Then
     ptxt = ptxt & " 0 " & tze1
    Else
     ptxt = ptxt & VStr & tze1
    End If
   End If
  End If
  rsAdSc.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doMachSQL1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMachSQL1(ti%, sql$)
'Function tabK¸$(Tl$)
' tabK¸ = "r" + ucase$(LEFT(Tl, 1)) + lcase(mid$(Tl, 2, 1))
'End Function
#End If

Function doMachSQL2$(TI%, ptxt$)
 Dim Tabl$, VStr$, umbrZahl%, obBegonnen%
 Dim trennz$, tza$, tze1$, tze2$
 Dim obHierKAI%, rDT$ ' ob hier kein AutoIncrement / Column-Name / Recordset-Datatype
 Dim rsAdSc As New ADODB.Recordset
 On Error GoTo fehler
 umbrZahl = 1
 Tabl = LCase$(tbn(TI))
 
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do
  If rsAdSc.BOF Then Exit Do
  If rsAdSc.EOF Then
   trennz = ")"""
  Else
   obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
   If obHierKAI Then trennz = ","
  End If
  If rsAdSc.EOF Or obHierKAI Then
   If Not obBegonnen Then
    obBegonnen = True
   Else
    ptxt = ptxt & tze2 & trennz
   End If
   If rsAdSc.EOF Then
    Exit Do
   Else

'Datentypen Recordset
'Const adEmpty = 0
'Const adSMALLINT = 2
'Const adInteger = 3
'Const adSingle = 4
'Const adDouble = 5
'Const adCurrency = 6
'Const adDate = 7
'Const adBSTR = 8
'Const adIDispatch = 9
'Const adError = 10
'Const adBoolean = 11
'Const adVariant = 12
'Const adIUnknown = 13
'Const adDecimal = 14
'Const adTinyInt = 16
'Const adUNSIGNEDTinyInt = 17
'Const adUNSIGNEDSMALLINT = 18
'Const adUNSIGNEDInt = 19
'Const adBigInt = 20
'Const adUNSIGNEDBigInt = 21
'Const adGUID = 72
'Const adBinary = 128
'Const adChar = 129
'Const adWChar = 130
'Const adNumeric = 131
'Const adUserDefined = 132
'Const adDBDate = 133
'Const adDBTime = 134
'Const adDBTimeStamp = 135
'Const adVarChar = 200
'Const adLongVarChar = 201
'Const adVarWChar = 202
'Const adLongVarWChar = 203
'Const adVarBinary = 204
'Const adLongVarBinary = 205
'    IF rsAdSc!column_name = "FaktPers" THEN Stop
    VStr = "r" & Tbk(TI) & "(i)." & colW(rsAdSc!COLUMN_NAME)
    Select Case rsAdSc!data_type
     Case "int", "tinyint", "smallint", "bigint", "decimal", "double", "float", 2, 3, 4, 5, 16, 17, 18, 19, 20, 21, 131, 139, 6, 14: tza = """ , ": tze1 = ", ": tze2 = """"  ' Zahlen
     Case "bit", 11, 128:                           tza = """ , cstr(-(": tze1 = "<>0)) , ": tze2 = """" ' VStr = "cStr(cint(" , VStr , "))" , trennz ' int, z.T.: "boolean" / Zuordnung von 128 nicht ganz sicher
     Case "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205: tza = "'"" , ": tze1 = " , ": tze2 = """'" 'VStr = """'"" + " , VStr , " + ""'" , trennz , """" ' Variant
     Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203:    tza = "'"" , ": tze1 = ", ": tze2 = """'" 'VStr = """'"" + " , VStr , " + ""'" , trennz , """"
     Case "datetime", "date", "timestamp", "time", 7, 64, 133, 134, 135:            tza = """ , DatFor_k(": tze1 = "), ": tze2 = """" 'VStr = " DatFor_k(" , VStr , ")" , trennz , """"
     Case Else: Err.Raise 999, , "Fehler in doMachSQL2: rsAdSc!data_type: ", rsAdSc!data_type, " (Datentyp nicht vorgesehen)"
    End Select
    '7.4.15: Aufteilung geht nicht
    If Len(ptxt) / umbrZahl > 200 Then
     umbrZahl = umbrZahl + 1
     If umbrZahl Mod 15 = 1 Then
'  sql = "  csql.AppVar Array(""("
      ptxt = ptxt & """)" & vbCrLf & "  csql.AppVar Array(" & Mid$(tza, 5)
     Else
      ptxt = ptxt & tza
      ptxt = ptxt & " _" & vbCrLf & "   "
     End If
    Else
     ptxt = ptxt & tza
    End If
    If Tabl = "namen" And InStrB(LCase$(VStr), "aktzeit") <> 0 Then
     ptxt = ptxt & " 0 " & tze1
    Else
     Select Case rsAdSc!data_type
      Case "double", "float", 4, 5 ' float, double
       ptxt = ptxt & "replace$(" & VStr & ","","",""."")" & tze1
      Case Else
       ptxt = ptxt & VStr & tze1
     End Select
    End If
   End If
  End If
  rsAdSc.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doMachSQL2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMachSQL2(ti%, sql$)
'Function tabK¸$(Tl$)
' tabK¸ = "r" + ucase$(LEFT(Tl, 1)) + lcase(mid$(Tl, 2, 1))
'End Function

Function doMachSQL3$(TI%, ptxt$)
 Dim Tabl$, VStr$, umbrZahl%, obBegonnen%
 Dim trennz$, tza$, tze1$, tze2$
 Dim obHierKAI%, rDT$ ' ob hier kein AutoIncrement / Column-Name / Recordset-Datatype
 Dim rsAdSc As New ADODB.Recordset
 On Error GoTo fehler
 umbrZahl = 1
 Tabl = LCase$(tbn(TI))
 
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, Tabl, Empty))
 Do
  If rsAdSc.BOF Then Exit Do
  If rsAdSc.EOF Then
   trennz = """" ' ")"""
  Else
'   obHierKAI = Not obAutoIncr(Tabl, rsAdSc)
'   IF obHierKAI THEN trennz = ","
   trennz = " AND "
  End If
  If True Then ' rsAdSc.EOF OR obHierKAI THEN
   If Not obBegonnen Then
    obBegonnen = True
   Else
    ptxt = ptxt & tze2 & trennz
   End If
   If rsAdSc.EOF Then
    Exit Do
   Else
    VStr = "r" & Tbk(TI) & "(i)." & colW(rsAdSc!COLUMN_NAME)
    Select Case rsAdSc!data_type
     Case "int", "tinyint", "smallint", "bigint", "decimal", "double", "float", 2, 3, 4, 5, 16, 17, 18, 19, 20, 21, 131, 139, 6, 14: tza = """ & ": tze1 = " & ": tze2 = """" ' Zahlen
     Case "bit", 11, 128:                            tza = """ & cstr(-(": tze1 = "<>0)) & ": tze2 = """" ' VStr = "cStr(cint(" , VStr , "))" , trennz ' int, z.T.: "boolean" / Zuordnung von 128 nicht ganz sicher
     Case "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205: tza = "'"" & ": tze1 = " & ": tze2 = """'" 'VStr = """'"" + " , VStr , " + ""'" , trennz , """" ' Variant
     Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203:    tza = "'"" & ": tze1 = " & ": tze2 = """'" 'VStr = """'"" + " , VStr , " + ""'" , trennz , """"
     Case "datetime", "date", "timestamp", "time", 7, 64, 133, 134, 135:               tza = """ & DatFor_k(": tze1 = ") & ": tze2 = """" 'VStr = " DatFor_k(" , VStr , ")" , trennz , """"
     Case Else: Err.Raise 999, , "Fehler in doMachSQL3: rsAdSc!data_type: ", rsAdSc!data_type, " (Datentyp nicht vorgesehen)"
    End Select
    Select Case rsAdSc!data_type
     Case "decimal", "double", "float": VStr = REPLACE$(VStr, ",", ".")
    End Select
    ptxt = ptxt & "`" & rsAdSc!COLUMN_NAME & "` = " & tza
    If Len(ptxt) / umbrZahl > 150 Then
     umbrZahl = umbrZahl + 1
     If umbrZahl Mod 15 = 1 Then
      ptxt = ptxt & """" & vbCrLf & "  sql = sql & "
     Else
      ptxt = ptxt & " _" & vbCrLf & "   "
     End If
    End If
    If Tabl = "namen" And InStrB(LCase$(VStr), "aktzeit") <> 0 Then
     ptxt = ptxt & " 0 " & tze1
    Else
     Select Case rsAdSc!data_type
      Case "double", "float", 4, 5 ' float, double
       ptxt = ptxt & "replace$(" & VStr & ","","",""."")" & tze1
      Case Else
       ptxt = ptxt & VStr & tze1
     End Select
    End If
   End If
  End If
  rsAdSc.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doMachSQL3/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doMachSQL3(ti%, sql$)
'Function tabK¸$(Tl$)
' tabK¸ = "r" + ucase$(LEFT(Tl, 1)) + lcase(mid$(Tl, 2, 1))
'End Function

Function ZPFeld$(i%)
 Select Case i
  Case 0: ZPFeld = "kAufDat" ' Namen, na
  Case 1: ZPFeld = "fanf" ' Faelle, fa
  Case 4: ZPFeld = "DiagDatum" ' Diagnosen, di
  Case 20: ZPFeld = "Dokudatum" ' dmpreihe, dm
  Case 21: ZPFeld = "erstZP" ' desktop, de
  Case Else: ZPFeld = "ZeitPunkt"
 End Select
End Function ' ZPFeld

Function MacheTypen(frm As Lese)
 Const SammelIns% = 0
 Dim rsAdSc As New ADODB.Recordset
 Dim i%, sql$, pText$, aktTbn$, aktTbk$, raText$, ii&, fat% ' fallabh‰ngige Tabelle
 Dim typFile$
 On Error GoTo fehler
 syscmd 4, "Mache Typen (1) ..."
 typFile = App.path + "\" + "typen.bas"
 On Error Resume Next
 Const ZdAltenDateien& = 8
 For i = ZdAltenDateien To 0 Step -1
  Name REPLACE$(typFile, ".bas", i & ".bas") As REPLACE$(typFile, ".bas", (i + 1) & ".bas")
 Next i
 Name typFile As REPLACE$(typFile, ".bas", "0.bas")
 syscmd 4, "Mache Typen (2) ..."
 On Error GoTo fehler
 Open typFile For Output As #257
 Call machAbK¸
' DBCnOpen CSStr
' Call myEFrag("use " & myDB)
 Print #257, "Option Explicit"
 Print #257, "Public obForK%"
 Print #257, "Dim sql$, T1!, T2!, maxL%"
 
' Print #257, ""
 syscmd 4, "Mache Typen (3) ..."
 For i = 0 To TbZ2
  Call doMacheTypen(tbn(i))
  DoEvents ' 2.4.13 sonst Absturz
 Next i
 Print #257, ""
 Print #257, "Type fzu"
 Print #257, " falt As Long"
 Print #257, " fneu As Long"
 Print #257, "End Type ' fzu"

 syscmd 4, "Mache Typen (3a) ..."
 Print #257, ""
 For i = 0 To TbZ2
  aktTbn = LCase$(REPLACE$(tbn(i), " ", "_"))
  Print #257, "Public r" & Tbk(i) & "() AS " & aktTbn & IIf(aktTbn = "formulare", "' kommt vor in: formulareSpeichern, doTabVorb, dolies", "")
 Next i
 Print #257, ""
 For i = 0 To TbZ2
  aktTbn = LCase$(REPLACE$(tbn(i), " ", "_"))
  Print #257, "Public ro" & Tbk(i) & "() AS " & aktTbn
 Next i
 syscmd 4, "Mache Typen (4) ..."
 Print #257, ""
 Print #257, "' in Geslies(2x)"
 Print #257, "Public FUNCTION Tinit()"
 Print #257, " static wdh%"
 Print #257, " ReDim rAna(0)"
' Print #257, " ReDim rFm(0)"
 For i = 0 To TbZ1
  Select Case Tbk(i)
   Case "Fo", "Fi"
   Print #257, " IF wdh = 0 THEN ReDim r" + Tbk(i) + "(0)"
   Case Else
   Print #257, " ReDim r" + Tbk(i) + "(0)"
  End Select
 Next i
 Print #257, " wdh = -1"
 Print #257, "End FUNCTION ' Tinit"
 Print #257, ""
 Print #257, "Public FUNCTION LabInit()"
 Print #257, " static wdh%"
' Print #257, " ReDim rFm(0)"
 syscmd 4, "Mache Typen (5) ..."
 For i = TbZ1 + 1 To TbZ2 - 1
  Select Case Tbk(i)
   Case "Fo", "Fi"
   Print #257, " IF wdh = 0 THEN ReDim r" + Tbk(i) + "(0)"
   Case Else
   Print #257, " ReDim r" + Tbk(i) + "(0)"
  End Select
 Next i
 syscmd 4, "Mache Typen (6) ..."
 Print #257, " wdh = -1"
 Print #257, "End FUNCTION ' Tinit"
 Print #257, ""
 Print #257, "' in AllesLˆsch, LabLˆsch"
 Print #257, "Public FUNCTION doEntleer(frm AS lese, Tbl$)"
 Print #257, " Dim rs As ADODB.Recordset"
 Print #257, "' SET rs = myEFrag(""SELECT COUNT(0) ct FROM `"" & Tbl & ""`"")"
 Print #257, " myFrag rs, ""SELECT COUNT(0) ct FROM `"" & Tbl & ""`"""
 Print #257, " If Not rs.BOF then"
 Print #257, "  frm.Ausgeb ""Lˆsche: `"" & Tbl & ""` ("" & rs!ct & "" Datens‰tze)"", True"
 Print #257, "  sql = sqlDeletefrom & ""`"" & Tbl & ""`"""
 Print #257, " End If ' Not rs.BOF then"
 Print #257, " Call myEFrag(sql) ' ,,adAsyncExecute"
 Print #257, " DoEvents"
 Print #257, "End FUNCTION ' doEntleer"
 Print #257, ""
 Print #257, "' in Pat_loeschen_Click, doPatvonMO"
 Print #257, "Public Sub LˆschePat(PID&, True)"
 Print #257, " Dim Tb, tbn, rAf&, ergeb$"
 Dim tbnS$
 tbnS = " tbn = Array("
 For i = 0 To TbZ1
  Select Case LCase$(tbn(i))
   Case "forminhaltform_abk", "formulare", "forminhfeld", "unbekannte kennungen"
   Case Else
    tbnS = tbnS & """" & LCase$(tbn(i)) & """"
    If i <> TbZ1 Then tbnS = tbnS & ", "
  End Select
 Next i
 tbnS = tbnS & ")"
 Print #257, " ON Error GoTo fehler"
 Print #257, tbnS
 Print #257, " For Each Tb In tbn"
 Print #257, "  myEFrag ""DELETE FROM `"" & Tb & ""` WHERE PAT_ID = "" & PID, rAf"
 Print #257, "  ergeb = ergeb & vbCrLf & rAf & "" S‰tze aus `"" & Tb & ""` gelˆscht."""
 Print #257, " Next"
 Print #257, " MsgBox ergeb"
 Print #257, " Debug.Print ergeb"
 Print #257, " Exit Function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, " #If VBA6 THEN"
 Print #257, "  AnwPfad = currentDB.Name"
 Print #257, " #Else"
 Print #257, "  AnwPfad = App.Path"
 Print #257, " #END IF"
 Print #257, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in LˆschePat/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " END SELECT"
 Print #257, "End FUNCTION ' LˆschePat"
 Print #257, ""
 Print #257, "Public FUNCTION AllesLˆsch(frm AS lese)"
 Print #257, " Dim ct&, rs As New ADODB.recordset"
 Print #257, " ON Error GoTo fehler"
 Print #257, " call ForeignNo0"
 Print #257, " call ForeignNo1"
 For i = TbZ1 To 0 Step -1
  Print #257, " call doEntleer(frm, """ & LCase$(tbn(i)) & """)"
 Next i
 Print #257, " call ForeignYes0"
 Print #257, " call ForeignYes1"
 Print #257, " Exit Function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, " #If VBA6 THEN"
 Print #257, " AnwPfad = currentDB.Name"
 Print #257, " #Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, " #END IF"
 Print #257, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in allesLˆsch/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " END SELECT"
 Print #257, "End FUNCTION ' AllesLˆsch"
 Print #257, ""
 Print #257, "Public FUNCTION LabLˆsch(frm AS lese)"
 Print #257, " Dim ct&, rs As New ADODB.recordset"
 Print #257, " ON Error GoTo fehler"
 Print #257, " call ForeignNo0"
 Print #257, " call ForeignNo1"
 For i = TbZ2 - 1 To TbZ1 + 1 Step -1
  Print #257, " call doEntleer(frm, """ & LCase$(tbn(i)) & """)"
 Next i
 Print #257, " call doentleer(frm, ""laborxplab"")"
 Print #257, " call doentleer(frm, ""laborxpneu"")"
 Print #257, " call doentleer(frm, ""laborxpnb"")"
 Print #257, " call ForeignYes0"
 Print #257, " call ForeignYes1"
 Print #257, " Exit Function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, " #If VBA6 THEN"
 Print #257, " AnwPfad = currentDB.Name"
 Print #257, " #Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, " #END IF"
 Print #257, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in LabLˆsch/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " END SELECT"
 Print #257, "End FUNCTION ' LabLˆsch"
 Print #257, ""
 Print #257, "Function doBezFeh(csqlVal$, obSpei%, ErrDes$)"
 Print #257, " Call ForeignNo0"
 Print #257, " Call ForeignNo1"
 Print #257, " obforK = True"
 Print #257, " IF obSpei <> 0 THEN"
 Print #257, "  Open BezFeh For Append AS #299"
 Print #257, "  Print #299, vbCrLf & vbCrLf & Now() & "": "" & csqlVal"
 Print #257, "  Print #299, vbCrLf & ErrDes"
 Print #257, "  Close #299"
 Print #257, " END IF"
 Print #257, "End FUNCTION 'doBezFeh"
 Print #257, ""
 Print #257, "' aufgerufen in alleSpeichern"
 Print #257, "Function fidSetz()"
 Print #257, " Dim i&, j&"
 syscmd 4, "Mache Typen (7) ..."
 For i = 0 To TbZ2 - 1
  Select Case Tbk(i)
   Case "Ei", "Fr", "Le", "Re", "Di", "Br", "Fu", "Me"
   Print #257, " For i = 1 To UBound(r" & Tbk(i) & ")" & " ' " & tbn(i)
'   Print #257, "  If r" & Tbk(i) & "(i).FID = 0 Then"
'   Print #257, "   If UBound(rFa) > 0 Then"
'   Print #257, "    r" & Tbk(i) & "(i).FID = rFa(1).FID"
'   Print #257, "   End If"
'   Print #257, "  Else ' r" & Tbk(i) & "(i).FID = 0 Then"
   Print #257, "  For j = 1 To UBound(rFa)"
   Print #257, "   If r" & Tbk(i) & "(i).FID = rFa(j).FID Then Goto " & Tbk(i) & "weiter"
   Print #257, "  Next j"
   Print #257, "  For j = 1 To UBound(rFa)"
   Print #257, "   If int(r" & Tbk(i) & "(i)." & IIf(Tbk(i) = "Di", "DiagDatum", "Zeitpunkt") & ") >= int(rFa(j).BhFB) AND int(r" & Tbk(i) & "(i)." & IIf(Tbk(i) = "Di", "diagDatum", "Zeitpunkt") & ") <= int(rFa(j).BhFE1) Then"
   Print #257, "    r" & Tbk(i) & "(i).FID = rFa(j).FID"
   Print #257, "    Exit For"
   Print #257, "   End If ' If int(r" & Tbk(i) & "(i)." & IIf(Tbk(i) = "Di", "DiagDatum", "Zeitpunkt") & ")) >= int(rFa(j).BhFB) AND int(r" & Tbk(i) & "(i)." & IIf(Tbk(i) = "Di", "diagDatum", "Zeitpunkt") & ") <= int(rFa(j).BhFE1) Then"
   Print #257, "  Next j"
   Print #257, Tbk(i) & "weiter:"
'   Print #257, "  End If ' r" & Tbk(i) & "(i).FID = 0"
   Print #257, " Next i"
  End Select
 Next i
 Print #257, "End Function ' FIDsetz"
 syscmd 4, "Mache Typen (8) ..."
 
#If False Then
 Print #257, "Public FUNCTION rAna_Hin(rs As Adodb.Recordset, rAna AS " & LCase$(tbn(TbZ2)) & ")" ' anamnesebogen; auf die Ersetzung von "rAna" durch tbk(tbz2) wurde aus mehreren Gr¸nden verzichtet
 Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(TbZ2)), Empty))
 Do While Not rsAdSc.EOF
  Print #257, " rsAnam." & colW(rsAdSc!COLUMN_NAME) & " = rs.fields(""" & rsAdSc!COLUMN_NAME & """)"
  rsAdSc.Move 1
 Loop
 Print #257, "End FUNCTION ' rAna_Hin"
#End If
 For i = 0 To TbZ2 - 1
  aktTbn = LCase$(REPLACE$(tbn(i), " ", "_"))
  Select Case aktTbn
   Case "forminhaltform_abk", "formulare", "forminhfeld", "kvnrue", "laborxeingel", "laborxus", "laborxbakt", "laborxwert", "laborxwert", "laborxleist", "laborxsaetze", "liuez", "unbekannte_kennungen"
   Case Else
   aktTbk = Tbk(i)
   Print #257, ""
   Print #257, "Public FUNCTION ro" & aktTbk & "Zuw(i&, j&)"
   Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(aktTbn), Empty))
   Do While Not rsAdSc.EOF
    Print #257, " ro" & aktTbk & "(i)." & rsAdSc!COLUMN_NAME & " = r" & aktTbk & "(j)." & rsAdSc!COLUMN_NAME
    rsAdSc.MoveNext
   Loop
   Print #257, "End FUNCTION ' ro" & aktTbk & "Zuw"
   Print #257, ""
   Print #257, "Public FUNCTION " & aktTbk & "ZUnt%(i&, j&)"
   Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(aktTbn), Empty))
   Do While Not rsAdSc.EOF
    Print #257, " IF ro" & aktTbk & "(i)." & rsAdSc!COLUMN_NAME & " <> r" & aktTbk & "(j)." & rsAdSc!COLUMN_NAME & " THEN gosub unter"
    rsAdSc.MoveNext
   Loop
   Print #257, " Exit Function"
   Print #257, "unter:"
   Print #257, " " & aktTbk & "ZUnt = " & aktTbk & "ZUnt + 1"
   Print #257, " Return"
   Print #257, "End FUNCTION ' " & aktTbk & "ZUnt"
   Print #257, ""
   pText = "Public FUNCTION " & aktTbn & "Laden("
   Print #257, pText & ")"
   Print #257, " Dim pid$, rs As New Recordset, akt&"
   Print #257, " ON Error GoTo fehler"
   Print #257, " pid = rNa(0).Pat_id"
   raText = ""
   pText = " sql = ""SELECT "
   Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(aktTbn), Empty))
   Dim izaehl%, iazaehl%
   izaehl = 0
   iazaehl = 0
 syscmd 4, "Mache Typen (9) ..."
   Do While Not rsAdSc.EOF
'    pText = pText & "COALESCE(" & rsAdSc!COLUMN_NAME & "," & IIf(rsAdSc!data_type.Value = "varchar" OR rsAdSc!data_type.Value = "longtext" OR rsAdSc!data_type.Value = "mediumtext" OR rsAdSc!data_type.Value = "char" OR rsAdSc!data_type.Value = "text" OR rsAdSc!data_type.Value = "varbinary", "''", "0") & ") " & rsAdSc!COLUMN_NAME
    Select Case rsAdSc!data_type.Value
     Case "date", "datetime", "timestamp"
'      pText = pText & "IF(" & rsAdSc!COLUMN_NAME & " IS NULL OR " & rsAdSc!COLUMN_NAME & "=0,CONVERT('18991230',DATE)," & rsAdSc!COLUMN_NAME & ") " & rsAdSc!COLUMN_NAME
      pText = pText & "COALESCE(" & rsAdSc!COLUMN_NAME & " - INTERVAL 0 DAY,CONVERT('18991230',DATE)) " & rsAdSc!COLUMN_NAME
     Case "varchar", "longtext", "mediumtext", "char", "varbinary"
      pText = pText & "COALESCE(" & rsAdSc!COLUMN_NAME & ",'') " & rsAdSc!COLUMN_NAME
     Case Else
      pText = pText & "COALESCE(" & rsAdSc!COLUMN_NAME & ",0) " & rsAdSc!COLUMN_NAME
    End Select
    izaehl = izaehl + 1
    If izaehl = 4 Then
     iazaehl = iazaehl + 1
     pText = pText + """"
     If iazaehl = 25 Then
      pText = pText & vbCrLf & "sql = sql & """
      iazaehl = 0
     Else
      pText = pText + " & _" & vbCrLf & """"
     End If
     izaehl = 0
    End If
    ' "varchar", "longtext", "mediumtext", "char", "text", "varbinary"
    raText = raText & "   ro" & aktTbk & "(akt)." & rsAdSc!COLUMN_NAME & " = " & IIf(rsAdSc!data_type.Value <> "varchar" And rsAdSc!data_type.Value <> "longtext" And rsAdSc!data_type.Value <> "mediumtext" And rsAdSc!data_type.Value <> "char" And rsAdSc!data_type.Value <> "text" And rsAdSc!data_type.Value <> "varbinary", "rs!" & rsAdSc!COLUMN_NAME, "doUmwfSQL(rs!" & rsAdSc!COLUMN_NAME & ", lies.obMySQL, False)") & vbCrLf
'    replace$(doUmwfSQL(rs!diagtext,lies.obMySQL,false), "≥", "¸")
    rsAdSc.MoveNext
    If Not rsAdSc.EOF Then pText = pText & ","
   Loop
   raText = raText & "   rs.MoveNext"
   pText = pText & " FROM `" & LCase$(tbn(i)) & "` WHERE Pat_ID="" & pid & "" ORDER BY `" & ZPFeld(i) & "`"
   Print #257, pText
   Print #257, " myFrag rs, sql"
   If aktTbn = "namen" Then
    Print #257, " ReDim ro" & aktTbk & "(1)"
    Print #257, " If Not rs.EOF Then"
   Else ' aktTbn = "namen"
    Print #257, " If rs.EOF Then"
    Print #257, "  ReDim ro" & aktTbk & "(0)"
    Print #257, " Else ' rs.EOF Then"
    Print #257, "  ReDim ro" & aktTbk & "(1)"
   End If ' aktTbn = "namen" else
 syscmd 4, "Mache Typen (10) ..."
   Print #257, "  Do While Not rs.EOF"
   Print #257, "   akt = UBound(ro" & aktTbk & ")"
   Print #257, raText
   Print #257, "   IF Not rs.EOF THEN ReDim Preserve ro" & aktTbk & "(UBound(ro" & aktTbk & ") + 1)"
   Print #257, "  Loop ' While Not rs.EOF"
   Print #257, " End If ' If rs.EOF"
   Print #257, " Exit Function"
   Print #257, "fehler:"
   Print #257, " Dim AnwPfad$"
   Print #257, " #If VBA6 THEN"
   Print #257, " AnwPfad = currentDB.Name"
   Print #257, " #Else"
   Print #257, " AnwPfad = App.Path"
   Print #257, " #END IF"
   Print #257, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in " & aktTbn & "Laden/"" + AnwPfad)"
   Print #257, "  Case vbAbort: Call MsgBox("" 2Hˆre auf ""): Progende"
   Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
   Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
   Print #257, " END SELECT"
   Print #257, "End FUNCTION ' " & aktTbn & "Laden"
   Print #257, ""
   Print #257, "Function " & aktTbn & "Einf"
   Print #257, " Dim rbeg&, roendpe&, ri&, roi&" & IIf(aktTbn = "faelle", ", jj&, fazu%()", "")
If aktTbn = "faelle" Then
    Print #257, " Dim fzu() As fzu"
    Print #257, " On Error GoTo fehler"
    Print #257, " If UBound(rFa) > 0 Then ' wenn neue F‰lle da"
    Print #257, "  For ri = 1 To UBound(rFa)"
    Print #257, "   If rFa(ri).Fanf >= qbeg Then ' aktqanf()"
    Print #257, "    rbeg = ri ' dann bezeichnet rbeg den ersten zu importierenden Fall"
    Print #257, "    For roendpe = 0 To UBound(roFa)"
    Print #257, "     If roFa(roendpe).Fanf >= qbeg Then"
    Print #257, "      Exit For ' ... und roFa den ersten nicht mehr zu verwendenden bestehenden Fall,"
    Print #257, "     End If ' roFa(roendpe).Fanf >= qbeg Then"
    Print #257, "    Next roendpe"
    Print #257, "    Exit For"
    Print #257, "   End If ' rFa(ri).Fanf >= qbeg Then ' aktqanf()"
    Print #257, "  Next ri"
    Print #257, " End If ' UBound(rFa) > 0"
    Print #257, " If rbeg = 0 Then"
    Print #257, "  roendpe = UBound(roFa) + 1 ', ggf. alle aus roFa zu verwenden, und nur die,"
    Print #257, " Else ' rbeg = 0 Then"
    Print #257, " ' ermitteln, welcher Fallnummer aus rFa jeder zu lˆschende Fall aus rFo entspricht"
    Print #257, "  If roendpe <= UBound(roFa) Then"
    Print #257, "   ReDim fazu(roendpe To UBound(roFa))"
    Print #257, "   For roi = roendpe To UBound(roFa)"
    Print #257, "    For ri = rbeg To UBound(rFa) - 1"
    Print #257, "     If roFa(roi).Fanf >= rFa(ri).Fanf And roFa(roi).Fanf < rFa(ri).BhFE Then"
    Print #257, "      fazu(roi) = ri"
    Print #257, "      GoTo Fertig"
    Print #257, "     End If ' roFa(roi).Fanf >= rFa(ri).Fanf And roFa(roi).Fanf < rFa(ri).BhFE"
    Print #257, "    Next ri"
    Print #257, "    fazu(roi) = UBound(rFa)"
    Print #257, "Fertig:"
    Print #257, "   Next roi"
    Print #257, " End If ' If roendpe <= UBound(roFa) Then"
    Print #257, "' die Fallnummern zu aus roFa zu lˆschenden S‰tze durch  die aus rFa ersetzen"
    Print #257, "  For roi = roendpe To UBound(ro" & Tbk(i) & ")"
    For ii = 0 To TbZ1
     fat = 0
     If Not rsAdSc Is Nothing Then If rsAdSc.State = 1 Then rsAdSc.Close
     Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(ii)), Empty))
     Do While Not rsAdSc.EOF
      If UCase$(rsAdSc!COLUMN_NAME) = "FID" Then fat = -1: Exit Do
      rsAdSc.Move 1
     Loop
     If fat And LCase$(Tbk(ii)) <> "fa" Then
      Print #257, "   For jj = 1 To UBound(ro" & Tbk(ii) & ")"
      Print #257, "    IF ro" & Tbk(ii) & "(jj).FID = roFa(roi).FID THEN ro" & Tbk(ii) & "(jj).FID = rFa(fazu(roi)).FID"
      Print #257, "   Next jj"
     End If ' fat
    Next ii
   Print #257, "  Next roi"
   Print #257, "  ReDim Preserve roFa(roendpe - 1) ' roFa k¸rzen ..."
   Print #257, "  ReDim Preserve roFa(roendpe + UBound(rFa) - rbeg) ' ... und erweitern ..."
   Print #257, "  For ri = rbeg To UBound(rFa)"
   Print #257, "   Call roFaZuw(roendpe + ri - rbeg, ri) ' und die F‰lle aus rFa an roFa anh‰ngen"
   Print #257, "  Next ri"
   Print #257, " End If ' rbeg = 0 Then Else"
   Print #257, " ReDim fzu(UBound(roFa)) ' dann fzu f¸llen"
   Print #257, " For ri = 1 To UBound(roFa)"
   Print #257, "  fzu(ri).falt = roFa(ri).FID"
   Print #257, "  fzu(ri).fneu = patanffid + ri - 1"
   Print #257, " Next ri"
   Print #257, " For ri = UBound(roFa) To 1 Step -1 ' dann die k¸nftigen Fallnummern statt den aktuellen verwenden"
   Print #257, "  If roFa(ri).FID <> fzu(ri).falt Then"
   Print #257, "   MsgBox ""Fehler bei "" & rNa(0).Pat_id & "", ri: "" & ri & "", "" & roFa(ri).FID & "" <> "" & fzu(ri).falt"
   Print #257, "  End If ' roFa(ri).FID <> fzu(ri).falt Then"
    For ii = 0 To TbZ1
     fat = 0
     If Not rsAdSc Is Nothing Then If rsAdSc.State = 1 Then rsAdSc.Close
     Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(ii)), Empty))
     Do While Not rsAdSc.EOF
      If UCase$(rsAdSc!COLUMN_NAME) = "FID" Then fat = -1: Exit Do
      rsAdSc.Move 1
     Loop
     If fat And LCase$(Tbk(ii)) <> "fa" Then
      Print #257, "  For jj = 1 To UBound(ro" & Tbk(ii) & ")"
      Print #257, "   If ro" & Tbk(ii) & "(jj).FID = roFa(ri).FID Then ro" & Tbk(ii) & "(jj).FID = fzu(ri).fneu"
      Print #257, "  Next jj"
     End If ' fat
    Next ii
   Print #257, "  roFa(ri).FID = fzu(ri).fneu"
   Print #257, " Next ri"

Else ' aktTbn = "faelle"
   Print #257, " On Error GoTo fehler"
   Print #257, " IF UBound(r" & Tbk(i) & ") > 0 THEN"
'   Print #257, "  rbeg = 0"
   Print #257, "  For ri = 1 To UBound(r" & Tbk(i) & ")"
   Print #257, "   IF r" & Tbk(i) & "(ri)." & ZPFeld(i) & " >= qbeg Then ' aktqanf()"
   Print #257, "    rbeg = ri"
   Print #257, "    Exit For"
   Print #257, "   END IF"
   Print #257, "  Next ri"
   Print #257, "  IF rbeg <> 0 THEN"
   Print #257, "   For roendpe = 0 To UBound(ro" & Tbk(i) & ")"
   Print #257, "    IF ro" & Tbk(i) & "(roendpe)." & ZPFeld(i) & " >= qbeg THEN"
   Print #257, "     Exit For"
   Print #257, "    END IF"
   Print #257, "   Next roendpe"
   Print #257, "   IF roendpe <= UBound(roFa) THEN"
' Folgendes durch weiter oben stehenden Teil veraltet
'   If aktTbn = "faelle" Then
'    Print #257, "    ReDim fazu(roendpe To UBound(ro" & Tbk(i) & "))"
'    Print #257, "    For roi = roendpe To UBound(ro" & Tbk(i) & ")"
'    Print #257, "     For ri = rbeg To UBound(r" & Tbk(i) & ") - 1"
'    Print #257, "      IF ro" & Tbk(i) & "(roi)." & ZPFeld(i) & " >= r" & Tbk(i) & "(ri)." & ZPFeld(i) & " AND ro" & Tbk(i) & "(roi)." & ZPFeld(i) & " < r" & Tbk(i) & "(ri).BhFE THEN"
'    Print #257, "       fazu(roi) = ri"
'    Print #257, "       GoTo fertig"
'    Print #257, "      END IF"
'    Print #257, "     Next ri"
'    Print #257, "     fazu(roi) = UBound(r" & Tbk(i) & ")"
'    Print #257, "fertig:"
'    Print #257, "    Next roi"
'    Print #257, "    For roi = roendpe To UBound(ro" & Tbk(i) & ")"
'    For ii = 0 To TbZ1
'     fat = 0
'     If Not rsAdSc Is Nothing Then If rsAdSc.State = 1 Then rsAdSc.Close
'     Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(ii)), Empty))
'     Do While Not rsAdSc.EOF
'      If UCase$(rsAdSc!COLUMN_NAME) = "FID" Then fat = -1: Exit Do
'      rsAdSc.Move 1
'     Loop
'     If fat And LCase$(Tbk(ii)) <> "fa" Then
'      Print #257, "     For jj = 1 To UBound(ro" & Tbk(ii) & ")"
'      Print #257, "      IF ro" & Tbk(ii) & "(jj).FID = roFa(roi).FID THEN ro" & Tbk(ii) & "(jj).FID = rFa(fazu(roi)).FID"
'      Print #257, "     Next jj"
'     End If ' fat
'    Next ii
'    Print #257, "    Next roi"
'   End If ' aktTbn = "faelle" THEN
   Print #257, "    ReDim Preserve ro" & Tbk(i) & "(roendpe - 1)"
   Print #257, "   END IF ' IF roendpe <= UBound(roFa) THEN"
   Print #257, "   ReDim Preserve ro" & Tbk(i) & "(roendpe + UBound(r" & Tbk(i) & ") - rbeg)"
   Print #257, "   For ri = rbeg To UBound(r" & Tbk(i) & ")"
   Print #257, "    Call ro" & Tbk(i) & "Zuw(roendpe + ri - rbeg, ri)"
   Print #257, "   Next ri"
   Print #257, "  END IF ' IF rbeg <> 0 THEN"
   Print #257, " END IF ' IF UBound(rFa) > 0 THEN"
End If ' aktTbn = "faelle" Then
   
   Print #257, " r" & Tbk(i) & " = ro" & Tbk(i) & ""
   Print #257, " Exit Function"
   Print #257, "fehler:"
   Print #257, " Dim AnwPfad$"
   Print #257, " #If VBA6 THEN"
   Print #257, " AnwPfad = CurrentDb.name"
   Print #257, " #Else"
   Print #257, " AnwPfad = App.path"
   Print #257, " #END IF"
   Print #257, " Select Case MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in " & aktTbn & "Einf/"" + AnwPfad)"
   Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
   Print #257, "  Case vbRetry: Call MsgBox("" Versuche nochmal""): Resume"
   Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
   Print #257, " END Select"
   Print #257, "End FUNCTION ' " & aktTbn & "Einf"
  End Select ' aktTbn <> "forminhaltform_abk" AND aktTbn <> "formulare" AND aktTbn <> "forminhFeld" AND aktTbn <> "unbekannte_kennungen" THEN
  Print #257, ""
  pText = "Public FUNCTION " & aktTbn & "Speichern(SammelInsert%, BezfSp%"
  Select Case aktTbn
   Case "laborxus" ' , "laborxwert", "laborxbakt", "laborxleist"
    pText = pText + ", j&"
  End Select
  Print #257, pText & ")"
  Print #257, " Dim i&, rAF&, Pid$, m%, sfnr%, altmode$ ',sql0$" ', DBCn As New adodb.Connection ' SpeicherFehler-Nr."
  If aktTbn = "faelle" Then
   Print #257, "Dim j%"
  End If
  Print #257, " Dim csql0 As New CString, csql As New CString"
  Print #257, " Dim rs As New ADODB.recordset"
  Print #257, " T1 = Timer"
  Print #257, " On Error Resume Next"
  Print #257, " Pid = rNa(0).Pat_id"
  If aktTbn = "unbekannte_kennungen" Then
  Else
   Print #257, " ON Error GoTo fehler"
  End If
  Print #257, " syscmd 4, pid & "": Speichere "" & Ubound(r" & Tbk(i) & ")+" & IIf(aktTbn = "namen", 1, 0) & " & "" S‰tze in `" & aktTbn & "`"""
'  Print #257, " DBCnOpen CSStr"
'  Print #257, " Call myEFrag(""use quelle1"")"
'  Print #257, " IF lese.obmysql THEN ON Error GoTo fehler ELSE ON Error Resume Next"
  Dim InsBefFld$, NobAI% ' True = es existiert ein Feld ohne AutoIncrement
  InsBefFld = doMachSQL0(tbn(i), NobAI)
   Print #257, "' sql0 =" & IIf(Not SammelIns, " "" INSERT "" & sqlignore & ", "") & " ""INTO `" & LCase$(tbn(i)) & "`" & _
              IIf(aktTbn = "forminhfeld" Or NobAI, InsBefFld, "") + " VALUES"
   Print #257, " Call csql0.AppVar(Array("" INSERT "", sqlIgnore, "" INTO `" & LCase$(tbn(i)) & "`" & IIf(aktTbn = "forminhfeld" Or NobAI, InsBefFld, ""), " VALUES""))"
   If SammelIns Then
    Print #257, " IF lese.obmysql THEN sql" & IIf(SammelIns, vNS, "0") & " = ""INSERT IGNORE "" & sql0 ELSE sql0 = ""INSERT "" & sql0"
   End If
  Dim ianf$, iend$
  Select Case aktTbn
   Case "namen", "laborxsaetze": ianf = "0"
   Case "formulare": ianf = "rFo1 + 1"
   Case "forminhaltform_abk": ianf = "rFi1 + 1"
   Case "unbekannte_kennungen": ianf = "rUn1 + 1"
   Case "laborxus": ianf = "j"
'   Case "laborxwert", "laborxbakt", "laborxleist": iAnf = "j"
   Case Else: ianf = "1"
  End Select
  Select Case aktTbn
   Case "laborxus": iend = "j"
'   case "laborxwert", "laborxbakt", "laborxleist": iend = "j"
   Case Else: iend = "ubound(r" + Tbk(i) + ")"
  End Select
 syscmd 4, "Mache Typen (14) ..."
  If aktTbn = "forminhkopf" Then
   Print #257, " FoIDv = 0"
   Print #257, "erneut:"
  End If
  Print #257, "sql:"
  Print #257, " csql.m_Len = 0"
 syscmd 4, "Mache Typen (17) ..."
  If i <= TbZ1 Then
   Print #257, " IF NOT Allepat THEN"
'  Print #257, "  ON error resume next"
   Select Case aktTbn
    Case "formulare", "forminhaltform_abk", "forminhfeld", "unbekannte_kennungen"
    Case Else
' Folgendes aus Performancegr¸nden vor Compactdatabase gesammelt
     Select Case aktTbn
      Case "forminhkopf"
        Print #257, "'    sql = ""DELETE FROM `forminhfeld` WHERE foid IN (SELECT foid FROM `forminhkopf` WHERE pat_id = "" & CStr(rNa(0).Pat_ID) & "")"""
        Print #257, "'    Call myEFrag(sql)"
     End Select
'     If aktTbn <> "namen" Then
'      Print #257, "'    sql = ""SELECT pat_id FROM `" & LCase$(tbn(i)) & "` WHERE Pat_ID = """ & " & CStr(rNa(0).Pat_ID)"
'      Print #257, "'    myFrag rs, sql"
'      Print #257, "'    IF Not rs.BOF THEN"
'      Print #257, "'     SET rs = nothing"
'     End If
     Print #257, "   sql = ""DELETE FROM `" & LCase$(tbn(i)) & "` WHERE Pat_ID = """ & " & CStr(rNa(0).Pat_ID)"
     Print #257, "   Call myEFrag(sql)"
     If aktTbn <> "namen" Then
'      Print #257, "    END IF ' Not rs.BOF THEN"
     End If
   End Select
   Print #257, " END IF ' not AllePat"
  End If ' i <= tbz1
  Print #257, " For i = " & ianf & " to " & iend
  Print #257, "'  r" + Tbk(i) + "(i).AktZeit = now()"
  If i <= TbZ1 Then
   Select Case aktTbn
    Case "forminhfeld", "forminhaltform_abk"
    Case Else
     Print #257, "  r" + Tbk(i) + "(i).StByte = CStr(AktByte)"
   End Select
  End If ' i <= TbZ1 Then
  'sql = "  sql = ""SELECT pat_id FROM " + lcase(Tbn(i)) + " WHERE "
  'Call doMachSQL0(Tbn(i), sql)
  'sql = LEFT(sql, Len(sql) - 5) + ""
  'Print #257, sql
  'Print #257, " IF rs.State = 1 THEN rs.Close"
  'Print #257, " myFrag rs, sql"
  'Print #257, " IF rs.bof THEN"
  
 syscmd 4, "Mache Typen (15) ..."
  If SammelIns Then
   sql = "  sql = iif(lese.obmysql,sql,sql0) & ""("
  Else ' SammelIns Then
   sql = "  sql = sql0 & ""("
  End If ' SammelIns Then
#If False Then
  Call doMachSQL1(i, sql)
  Print #257, "' " & sql
#End If

  Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(i)), Empty))
  If noDup(i) Then ' zu Zeit laborxwert
   Print #257, "    Dim j&"
   Print #257, "    For j = 1 To i - 1"
   Do While Not rsAdSc.EOF
    Print #257, "    IF r" & Tbk(i) & "(i)." & rsAdSc!COLUMN_NAME & " <> r" & Tbk(i) & "(j)." & rsAdSc!COLUMN_NAME & " THEN GoTo nextj"
    rsAdSc.MoveNext
   Loop
   Print #257, "     GoTo nexti"
   Print #257, "nextj:"
   Print #257, "     Dim rsdop As New ADODB.Recordset"
   Print #257, "     SET rsdop = Nothing"
'   Dim dupsql As New CString
'   dupsql.Clear
'   dupsql.AppVar Array("rsdop.Open ""SELECT 0 FROM `", LCase$(tbn(i)), "` WHERE ")
'   Print #257, "     myFrag rsdop, ""SELECT 0 FROM laborxwert WHERE refnr = "" & rLw(i).RefNr & "" AND abk¸ = '"" & rLw(i).Abk¸ & ""'"""
   sql = "     myFrag rsdop, ""SELECT 0 FROM `" & LCase$(tbn(i)) & "` WHERE "
   Call doMachSQL3(i, sql)
   Print #257, sql
   Print #257, "     IF Not rsdop.EOF THEN GoTo nexti"
   Print #257, "    Next j"
  End If ' noDup(i)
  
 syscmd 4, "Mache Typen (16) ..."
  Print #257, "  IF SammelInsert = 0 OR i = " & ianf & " THEN"
  Print #257, "   csql.Append csql0"
  Print #257, "  END IF"
  
  sql = "  csql.AppVar Array(""("
  Call doMachSQL2(i, sql)
 syscmd 4, "Mache Typen (17) ..."
  Print #257, sql & ")"
  
  Print #257, "  IF SammelInsert <> 0 AND i < " & iend & " THEN csql.Append "","""
  Print #257, "  IF SammelInsert = 0 OR i = " & iend & " THEN"
  
  If SammelIns Then Print #257, "   IF lese.obmysql THEN"
  If SammelIns Then Print #257, "    IF i < " & iend & " THEN sql = sql & "","""
  If SammelIns Then Print #257, "   ELSE"
'  Print #257, "  Call myEFrag(sql,rAf)', , adAsyncExecute)"
  If aktTbn = "faelle" Then
   Print #257, "'   IF Not obForK THEN ForeignNo0"
  End If ' aktTbn = "faelle" Then
  Print #257, "'    altmode = DBCn.Execute(""SELECT @@global.sql_mode"").Fields(0)"
  Print #257, "    altmode = myEFrag(""SELECT @@global.sql_mode"", , DBCn).Fields(0)"
  Print #257, "'   Call DBCn.Execute(""SET GLOBAL sql_mode='STRICT_TRANS_TABLES'"") ' NO_ENGINE_SUBSTITUTION"
  Print #257, "    myEFrag ""SET GLOBAL sql_mode='STRICT_TRANS_TABLES'"", , DBCn ' NO_ENGINE_SUBSTITUTION"
  Print #257, "'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23"
  Print #257, "   Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)"
  If aktTbn = "faelle" Then
   Print #257, "'   IF Not obFork THEN ForeignYes0"
   Print #257, "   IF rAF = 0 THEN"
   Print #257, "    Err.Raise 998, , ""Fehler in " & aktTbn & "Speichern b.Pat. "" & rFa(i).Pat_id & "", Err.Number "" & Err.Number & "", err.description: "" & Err.Description"
   Print #257, "   END IF ' rAF = 0 THEN"
  End If
  Print #257, "   IF obforK THEN"
  Print #257, "    Call ForeignYes0"
  Print #257, "    Call ForeignYes1"
  Print #257, "   END IF ' obforK THEN"
'  Print #257, "   IF lese.obmysql AND obmitAlterTab THEN"
'  Print #257, "    SET rs = myEFrag(""SHOW WARNINGS"")"
'  Print #257, "    IF not rs.BOF() THEN"
'  Print #257, "     IF rs!code = 1265 THEN"
'  Print #257, "      Err.Raise -2147217833"
'  Print #257, "     END IF ' rs!code = 1265 THEN"
'  Print #257, "    END IF ' lese.obmysql AND obmitAlterTab THEN"
'  Print #257, "   END IF ' lese.obmysql AND obmitAlterTab"
  If aktTbn = "faelle" Then
   Print #257, "   IF SammelInsert = 0 THEN"
   Print #257, "   'Hier gibts mit Sammelins noch ein Problem ..."
   Print #257, "'    set rs = nothing"
   Print #257, "'    For j = 2 To 2"
   Print #257, "'     If j = 1 Then"
   Print #257, "'      Set rs = myEFrag(""SELECT LAST_INSERT_ID() FID"") ' session-spezifisch '27.8.23: liefert in Schleife immer die erste Zahl, auch mit Commit zwischendrin"
   Print #257, "'     Else ' j = 1 Then"
   Print #257, "      Set rs = myEFrag(""SELECT MAX(fid) FID FROM `faelle` WHERE pat_id = "" & rFa(i).Pat_id & "" AND quartal = '"" & rFa(i).Quartal & ""' AND bhfb = "" & DatFor_k(rFa(i).BhFB) & "" AND bhfe1 = "" & DatFor_k(rFa(i).BhFE1) & "" AND ausgst = "" & DatFor_k(rFa(i).ausgst))"
   Print #257, "'     End If"
   Print #257, "'     If Not rs.BOF Then If rs.Fields(0) <> 0 Then Exit For"
   Print #257, "'    Next j"
   Print #257, "    IF rs.BOF Then"
   Print #257, "     Err.Raise 999, , ""Fehler bei der Fallaktualisierung b.Pat. "" & rFa(i).Pat_ID & "", FID "" & rFa(i).FID"
   Print #257, "    ElseIf rs!FID = 0 Then"
   Print #257, "     MsgBox ""Fehler in faellespeichern:"" & vbCrLf & rs.source"
   Print #257, "     GoTo sql"
   Print #257, "    Else ' rs.BOF Then"
   Print #257, "     neufid = rs!FID"
   Print #257, "     If neufid <> rFa(i).FID Then"
   Print #257, "      Lese.Ausgeb ""ƒnderung bei der FallID  bei Pat. "" & rFa(i).Pat_ID & "", FID "" & rFa(i).FID & "" -> "" & neufid & "" in zu speichernden Tabellen mit fallid"", True "
'   Print #257, "      neufid = neufid + 1"
   Print #257, "      Dim jjj&"
 syscmd 4, "Mache Typen (19) ..."
   For ii = 0 To TbZ1
    If Tbk(ii) <> "Fa" Then
     fat = 0
     If Not rsAdSc Is Nothing Then If rsAdSc.State = 1 Then rsAdSc.Close
     Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(ii)), Empty))
     Do While Not rsAdSc.EOF
      If UCase$(rsAdSc!COLUMN_NAME) = "FID" Then fat = -1: Exit Do
      rsAdSc.Move 1
     Loop
     If fat Then
      Print #257, "      For jjj = 1 To UBound(r" & Tbk(ii) & ")"
      Print #257, "       IF r" & Tbk(ii) & "(jjj).FID = rFa(i).FID THEN"
      Print #257, "        r" & Tbk(ii) & "(jjj).FID = neufid"
      Print #257, "       END IF ' r" & Tbk(ii) & "(jjj).FID = rFa(i).FID THEN"
      Print #257, "      Next jjj"
     End If ' fat Then
    End If ' Tbk(ii) <> "Fa" Then
   Next ii
   Print #257, "     END IF ' neufid <> rFa(i).FID Then"
   Print #257, "     neufid = neufid + 1 ' f¸r den n‰chsten Patienten"
   Print #257, "    END IF ' rs.BOF Then"
   Print #257, "    csql.m_len = 0"
   Print #257, "   END IF ' IF SammelInsert = 0 Then"
  End If ' aktTbn = "faelle" Then
  Print #257, "  END IF ' SammelInsert = 0 OR i = " & iend
  If SammelIns Then Print #257, "  END IF"
  If noDup(i) Then
   Print #257, "nexti:"
  End If
  Print #257, "  DoEvents"
  Print #257, " Next i"
  Dim iiAnf$
  Select Case aktTbn
   Case "formulare": iiAnf = "rFo1"
   Case "forminhaltform_abk": iiAnf = "rFi1"
   Case "unbekannte_kennungen": iiAnf = "rUn1"
   Case Else: iiAnf = vNS
  End Select
  If iiAnf <> "" Then
   Print #257, " " + iiAnf + " = UBound(r" + Tbk(i) + ")"
  End If ' iiAnf <> "" Then
  If SammelIns Then Print #257, " IF lese.obmysql THEN IF ubound(r" & Tbk(i) & ")>0 THEN call myEFrag(sql)', , adAsyncExecute)"
  Print #257, " syscmd 5"
  Print #257, " Exit Function"
  Print #257, "fehler:"
  Print #257, "ErrDescr = Err.Description"
  Print #257, "ErrNumber = Err.Number"
  Print #257, "syscmd 4, ""r" + Tbk(i) + "("" & i & ""/"" & UBound(r" + Tbk(i) + ") & ""):   "" & ErrDescr"
 syscmd 4, "Mache Typen (21) ..."
  If aktTbn = "forminhkopf" Then
   Print #257, "If ErrNumber = -2147217900 And ErrDescr Like ""*Duplicate entry * for key 'PRIMARY'"" Then"
   Print #257, " Dim schl¸ssel$, pos&, iiru&, jjru&"
   Print #257, " pos = InStr(ErrDescr, ""'"")"
   Print #257, " schl¸ssel = Mid$(ErrDescr, pos + 1, InStr(pos + 2, ErrDescr, ""'"") - pos - 1)"
   Print #257, "' Debug.Print schl¸ssel"
   Print #257, " For iiru = 1 To UBound(rFr)"
   Print #257, "  If rFr(iiru).Foid = schl¸ssel Then"
'   Print #257, "   If FoIDv = 0 Then FoIDv = DBCn.Execute(""SELECT (MAX(foid)+1) FROM forminhkopf"").fields(0) Else FoIDv = FoIDv + 1"
   Print #257, "   If FoIDv = 0 Then FoIDv = MyEfrag(""SELECT (MAX(foid)+1) FROM forminhkopf"", , DBCn).fields(0) Else FoIDv = FoIDv + 1"
   Print #257, "   rFr(iiru).Foid = FoIDv"
   Print #257, "   For jjru = 1 To UBound(rFm)"
   Print #257, "    If rFm(jjru).Foid = schl¸ssel Then rFm(jjru).Foid = FoIDv"
   Print #257, "   Next jjru"
   Print #257, "   Exit For"
   Print #257, "  End If ' rFr(iiru).Foid = schl¸ssel Then"
   Print #257, " Next iiru"
   Print #257, " csql = """""
   Print #257, " Resume erneut"
   Print #257, "End If ' ErrNumber = -2147217900 And ErrDescr Like ""*Duplicate entry * for key 'PRIMARY'"" Then"
  End If 'aktTbn = "forminhkopf" Then
  Print #257, "sfnr = sfnr + 1"
  Print #257, "If sfnr > 10 then "
  Print #257, " Lese.Ausgeb sfnr & "" Fehler in """"" & aktTbn & "Speichern()"""" bei Pat. "" & rNa(0).Pat_id & "", gebe auf (ErrDes: "" & ErrDescr & "")"", True"
  Print #257, " sfnr = 0"
  Print #257, " Resume Next"
  Print #257, "End if ' sfnr > 10"
  Print #257, "IF ErrNumber = -2147217900 AND InStrB(ErrDescr, "" Doppelter; Eintrag; "") <> 0 THEN"
  Print #257, " Call Shell(App.path + ""\..\nachricht\nachricht.exe "" & App.EXEName & "" Doppelter Eintrag bei: "" & vbCrLf & csql.Value)"
  Print #257, " Resume Next"
  Print #257, "ElseIf ErrNumber = -2147467259 AND InStrB(ErrDescr, ""Daten zu lang"") = 0 AND InStrB(ErrDescr, ""Data too long"") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails"
  Print #257, " IF InStrB(ErrDescr, ""'READ-COMMITTED'"") <> 0 THEN"
  Print #257, "  myEFrag ""SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ"", rAF"
  Print #257, " Else"
  Print #257, "  Call doBezFeh(csql.Value, BezfSp, ErrDescr)"
  Print #257, " END IF"
  Print #257, " Resume"
  Print #257, "ElseIf ErrNumber = -2147217833 OR InStrB(ErrDescr, ""Daten zu lang"") <> 0 OR InStrB(ErrDescr, ""Data too long"") <> 0 THEN"
  Print #257, " Dim rsc As Adodb.Recordset, maxi%(), k%"
 syscmd 4, "Mache Typen (22) ..."
  Dim SpZ%, j%
  SpZ = 0
  Set rsAdSc = Nothing
  Set rsAdSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(tbn(i)), Empty))
  Do While Not rsAdSc.EOF
   Select Case rsAdSc!data_type
    Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
     SpZ = SpZ + 1
   End Select
   rsAdSc.Move 1
  Loop
  Print #257, " redim maxi(" & SpZ - 1 & ")"
  rsAdSc.MoveFirst
  j = 0
  Print #257, " for k = iif(SammelInsert<>0," & ianf & ",i) to iif(SammelInsert<>0," & iend & ",i)"
  Do While Not rsAdSc.EOF
   Select Case rsAdSc!data_type
    Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
     Print #257, "  IF Len(r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & ") > maxi(" & j & ") THEN maxi(" & j & ") = Len(r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & ")"
     j = j + 1
   End Select
   rsAdSc.Move 1
  Loop ' While Not rsAdSc.EOF
  Print #257, " next k"
  If i <= TbZ1 Then
   Print #257, " If obTrans <> 0 Then If myEFrag(""SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()"", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0"
'  Print #257, " IF obTrans <>0 Then If DBCn.Execute(""SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()"").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0"
  End If ' i <= TbZ1 Then
 syscmd 4, "Mache Typen (23) ..."
  Print #257, " nochmal:"
  Print #257, " SET rsc = New ADODB.Recordset"
  Print #257, " SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, """ & LCase$(tbn(i)) & """, Empty))"
  Print #257, " m = 0"
  Print #257, " Do While Not rsc.EOF"
  Print #257, "  SELECT CASE rsc!data_type"
  Print #257, "   Case ""varchar"", ""longtext"", ""mediumtext"", ""char"", ""text"", ""varbinary"", 8, 129, 130, 200, 201, 202, 203, ""set"", ""enum"", ""blob"", ""longblob"", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205"
  Print #257, "    maxL = SpMod(maxi(m), """ & LCase$(tbn(i)) & """, rsc)"
  Print #257, "    IF maxL > 0 THEN"
  Print #257, "     For k = IIf(SammelInsert <> 0," & ianf & ", i) To IIf(SammelInsert <> 0," & iend & ", i)"
  Print #257, "      SELECT CASE m"
  j = 0
  rsAdSc.MoveFirst
  Do While Not rsAdSc.EOF
   Select Case rsAdSc!data_type
    Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
     Print #257, "       Case " & j & ": Lese.Ausgeb ""   Verk¸rze Inhalt von r" & Tbk(i) & "." & colW(rsAdSc!COLUMN_NAME) & ": '"" & r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & " & ""' -> '"" & Left$(r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & ", maxL)  & ""'"",true : r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & " = Left$(r" & Tbk(i) & "(k)." & colW(rsAdSc!COLUMN_NAME) & ", maxL)"
     j = j + 1
   End Select
   rsAdSc.Move 1
  Loop ' While Not rsAdSc.EOF
  Print #257, "      END SELECT"
  Print #257, "     Next"
  Print #257, "    elseif maxl < 0 THEN"
  Print #257, "     goto nochmal"
  Print #257, "    END IF"
  Print #257, "    m = m + 1"
  Print #257, "  END SELECT"
  Print #257, "  IF rsc.State = 0 THEN Exit Do"
  Print #257, "  rsc.Move 1"
  Print #257, " Loop"
  Print #257, " Call ForeignNo0" ' 29.10.08
  Print #257, " Call ForeignNo1"
 syscmd 4, "Mache Typen (24) ..."
  
    
'  Print #257, "  SELECT CASE rsc!column_name"
'  rsAdSc.MoveFirst
'  Do While Not rsAdSc.EOF
'   SELECT CASE rsAdSc!data_type
'    Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
'     Print #257, "   Case """ & colW(rsAdSc!column_name) & """:  IF SpMod(r" & Tbk(i) & "(i)." & colW(rsAdSc!column_name) & ", """ & LCase$(Tbn(i)) & """, rsc) THEN Exit Do"
'   END SELECT
'   rsAdSc.Move 1
'  Loop
'  Print #257, "  END SELECT"
'  Print #257, "  rsc.Move 1"
'  Print #257, " Loop"

'  IF i <= TbZ1 THEN ' Kommentar 29.10.08
'   Print #257, " DBCn.BeginTrans"
'  END IF
'  Print #257, " IF lese.obMysql THEN Resume next ELSE resume" ' ge‰ndert 20.8.23
  Print #257, " resume"
  Print #257, "ElseIf InStrB(1, ErrDescr, ""gone away"", vbTextCompare) <> 0 Or InStrB(ErrDescr, ""ost connection"") <> 0 Then"
  Print #257, " DBCnOpen"
  Print #257, " Resume"
  If aktTbn = "namen" Then
   Print #257, "ElseIf ErrNumber = -2147217871 OR ErrNumber = -2147217859 OR ErrNumber = -2147467259 THEN"
   Print #257, " For i = 0 To 10"
   Print #257, "  Call ForeignYes0"
   Print #257, "  Call ForeignYes1"
   Print #257, " Next i"
   Print #257, " Call ForeignNo0"
   Print #257, " Call ForeignNo1"
   Print #257, " Resume"
  End If ' aktTbn = "namen" Then
  Print #257, "END IF ' ErrNumber = "
  If aktTbn = "faelle" Then
   Print #257, "IF ErrNumber = -2147467259 THEN"
   Print #257, " Dim sqlquer$"
   Print #257, " sqlquer = ""INSERT INTO `kassenliste`(name,kurzname,`GO`,`VKNR`,`IK`,`eingef`,pid) VALUES ("" & ""'"" & rFa(I).kasse & ""', '"" & rFa(I).kkasse_2 & ""', '"" & rFa(I).GOƒKatName & ""', '"" & rFa(I).VKNr & ""', '"" & rFa(I).IK & ""',"" & Format(NOW(), ""yyyymmddHHMMSS"") & "","" & rFa(I).Pat_id & "")"""
   Print #257, " InsKorr DBCn, DBCnS, sqlquer, rAF"
   Print #257, " Resume"
   Print #257, "END IF ' ErrNumber = -2147467259 THEN"
  End If ' lcase(Tbn(I)) = "faelle" THEN
  Print #257, " Dim AnwPfad$"
  Print #257, "#If VBA6 THEN"
  Print #257, " AnwPfad = currentDB.Name"
  Print #257, "#Else"
  Print #257, " AnwPfad = App.Path"
  Print #257, "#END IF"
  Print #257, " SELECT CASE MsgBox(""FNr: "" + CStr(ErrNumber) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in " & aktTbn & "Speichern" + "/"" + AnwPfad)"
  Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
  Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): If ErrNumber = 998 Then Resume sql Else Resume"
  Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
  Print #257, " END SELECT"
  Print #257, "End FUNCTION ' " & aktTbn & "Speichern"
 syscmd 4, "Mache Typen (25) ..."
 Next i ' For i = 0 to TbZ2 - 1
 syscmd 4, "Mache Typen (26) ..."
 Print #257, ""
 Print #257, "Public FUNCTION tuLaden"
 For i = 0 To TbZ1
  aktTbn = LCase$(REPLACE$(tbn(i), " ", "_"))
  If aktTbn <> "forminhaltform_abk" And aktTbn <> "formulare" And aktTbn <> "forminhfeld" And aktTbn <> "unbekannte_kennungen" And aktTbn <> "kvnrue" Then
   Print #257, " call " & aktTbn & "Laden"
  End If
  Select Case aktTbn
  Case "faelle", "forminhaltform_abk"
  Print #257, "'  IF not lese.obmysql THEN"
  Print #257, "'   IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0"
  Print #257, "'    Call DBCn.BeginTrans: obTrans = 1"
  Print #257, "'  END IF ' not lese.obmysql "
  Print #257, "   wechsTrans"
  End Select
 Next i
 syscmd 4, "Mache Typen (27) ..."
 Print #257, "End FUNCTION ' tuLaden"
 Print #257, ""
 Print #257, "Public FUNCTION tuSpeichern(frm AS Lese, SI%, BfS%) ' frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern"
 Print #257, " Dim rAf&, altsi$,altsam%"
 Print #257, " altsi = sqlIGNORE"
 Print #257, " sqlIGNORE = """""
 Print #257, " altsam = SI"
 Print #257, " sqlIGNORE = """""
 Print #257, " ON Error GoTo fehler"
 For i = 0 To TbZ1
  aktTbn = LCase$(REPLACE$(tbn(i), " ", "_"))
  If aktTbn = "formulare" Then
   Print #257, " sqlIGNORE = "" IGNORE """
  ElseIf aktTbn = "faelle" Then
   Print #257, " SI = 0"
   Print #257, " ComTrans"
  End If
  Print #257, " call " & aktTbn & "Speichern(SI, BfS)"
  If aktTbn = "faelle" Then
   Print #257, " SI = altsam"
  End If
  If aktTbn = "formulare" Then
   Print #257, " sqlIGNORE = """""
  End If
  Select Case aktTbn
   Case "faelle", "forminhaltform_abk"
'    Print #257, "   IF not lese.obmysql THEN"
'    Print #257, "    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0"
'    Print #257, "    Call DBCn.BeginTrans: obTrans = 1"
'    Print #257, "   END IF ' not lese.obmysql"
    Print #257, " wechsTrans"
  End Select
 Next i
 syscmd 4, "Mache Typen (28) ..."
 Print #257, " Call myEFrag(""UPDATE `namen` SET aktZeit = "" & DatFor_k(rNa(0).AktZeit) & "" WHERE pat_id = "" & rNa(0).Pat_ID,rAf)"
 Print #257, " IF rAf <> 1 THEN "
 Print #257, "  frm.Ausgeb ""Fehler bei der Setzung des Aktualisierungsdatum bei "" & rNa(0).Pat_ID & "" "" & rNa(0).Nachname & "" "" & rNa(0).Vorname, true"
 Print #257, " END IF ' rAf <> 0"
 Print #257, " sqlIGNORE = altsi"
 Print #257, " Exit Function"
 Print #257, "fehler:"
 Print #257, " Dim AnwPfad$"
 Print #257, "#If VBA6 THEN"
 Print #257, " AnwPfad = currentDB.Name"
 Print #257, "#Else"
 Print #257, " AnwPfad = App.Path"
 Print #257, "#END IF"
 Print #257, " ErrDescr = Err.Description"
 Print #257, " IF InStrB(ErrDescr, ""'READ-COMMITTED'"") <> 0 THEN"
 Print #257, "  myEFrag ""SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ"", rAF"
 Print #257, "  Resume"
 Print #257, " END IF"
 Print #257, " SELECT CASE MsgBox(""FNr: "" & CStr(Err.Number) & vbCrLf & ""LastDLLError: "" & CStr(Err.LastDllError) & vbCrLf & ""Source: "" & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) & vbCrLf & ""Description: "" & Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in tuSpeichern/"" + AnwPfad)"
 Print #257, "  Case vbAbort: Call MsgBox("" Hˆre auf ""): Progende"
 Print #257, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #257, "  Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
 Print #257, " END SELECT"
 Print #257, "End FUNCTION ' tuSpeichern"
 Print #257, ""
 
 syscmd 4, "Mache Typen (29) ..."
 Close #257
 frm.Ausgeb "Fertig mit MacheTypen", True
 syscmd 5
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in MacheTypen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MacheTypen

Public Function HolReg(hlese As Lese)
' wird aufgerufen aus MDIFormLoad und allen Abbruchvorg‰ngen
 Dim reg
 On Error GoTo fehler
 obStart = True
 RegStelle = RegWurzel & App.EXEName
 hlese.Ziel = getReg(1, RegStelle, "MdB")
 reg = getReg(1, RegStelle, "obAcc")
 If reg = "" Or reg = "Falsch" Or reg = "False" Then
  hlese.obAcc = False
 Else
    hlese.obAcc = reg
 End If
 hlese.obMySQL = Not hlese.obAcc
 reg = getReg(1, RegStelle, "obMyQuelle")
 If reg = "" Or reg = "Falsch" Or reg = "False" Then
 Else
    hlese.MyDB = "quelle"
 End If
 reg = getReg(1, RegStelle, "obMyQuelle1")
 If reg = "" Or reg = "Falsch" Or reg = "False" Then
 Else
    hlese.MyDB = "quelle1"
 End If
 reg = getReg(1, RegStelle, "obMyQuell2")
 If reg = "" Or reg = "Falsch" Or reg = "False" Then
 Else
    hlese.MyDB = "quelle2"
 End If
 If hlese.MyDB = "" Then
  hlese.MyDB = "quelle"
 End If
' hlese.dbv.ConStr = getReg(1, RegStelle, "ConStri")
 Dim dbvConstr
 dbvConstr = getReg(1, RegStelle, "ConStri")
 If VarType(dbvConstr) = vbString Then hlese.dbv.Constr = dbvConstr
 hlese.ConStri = hlese.dbv.Constr
 hlese.dlg.ConStrLabel = getReg(1, RegStelle, "ConStrLabel")
 hlese.dlg.ConstrCn = getReg(1, RegStelle, "ConStr")
 hlese.dbv.CnStr = hlese.dlg.ConstrCn
 hlese.dbv.Constr = hlese.dlg.ConStrLabel
 hlese.dlg.MdB = hlese.Ziel
 hlese.dlg.obAcc = hlese.obAcc
 If hlese.obAcc Then
  hlese.dlg.obAcc = True
 Else
  hlese.dlg.obMySQL = True
 End If
 obStart = False
' Call hlese.ConstrFestleg(0, hlese)
 Dim pServ$
 Call acon(quelleT, , , , hlese.dlg.ConstrCn, True)
 imAufbauDialog = True
 Call hlese.HolEinstvonDB
 imAufbauDialog = False
 
 hlese.dlg.EmDatei = getReg(1, RegStelle, "Email-Datei")
 reg = getReg(1, RegStelle, "Datei f¸r nachzuholende Laborimporte")
 DateiNachzuholen = uVerz & "anamnese\Nachholen.txt"
 If reg <> "" Then hlese.snst.DateiNachzuholen = reg Else hlese.snst.DateiNachzuholen = DateiNachzuholen
 
 reg = getReg(1, RegStelle, "Quelle f¸r Anamnesebˆgen")
 If reg = "" Then reg = QmdB ' in Importiere.bas: uverz & "Anamnese\Quelle.mdb"
 hlese.snst.QuelleAnamneseBˆgen = reg
 
 ' reg = getLDatei(hlese.snst.QuelleAnamneseBˆgen) ' Kommentar 13.2.22
 reg = getReg(1, RegStelle, "Debugdatei")
 DebugDatei = uVerz & "anamnese\debug.txt"
 If reg <> "" Then hlese.snst.DebugDatei = reg Else hlese.snst.DebugDatei = DebugDatei
 obStart = False
' dlg.hlese.Ziel = dlg.MdB
' Call hlese.ZeigDateien
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in HolReg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolReg(dlg AS Dialog)

Public Function PutReg(hlese As Lese)
 RegStelle = RegWurzel & App.EXEName
 Call fStSpei(HCU, RegStelle, "obAcc", hlese.obAcc)
 Call fStSpei(HCU, RegStelle, "obMyQuelle", hlese.obMySQL And hlese.MyDB = "quelle")
 Call fStSpei(HCU, RegStelle, "obMyQuelle1", hlese.obMySQL And hlese.MyDB = "quelle1")
 Call fStSpei(HCU, RegStelle, "obMyQuelle2", hlese.obMySQL And hlese.MyDB = "quelle2")
 Call fStSpei(HCU, RegStelle, "MdB", hlese.Ziel)
 Call fStSpei(HCU, RegStelle, "Email-Datei", hlese.dlg.EmDatei)
 Call fStSpei(HCU, RegStelle, "Datei f¸r nachzuholende Laborimporte", hlese.snst.DateiNachzuholen)
 Call fStSpei(HCU, RegStelle, "Quelle f¸r Anamnesebˆgen", hlese.snst.QuelleAnamneseBˆgen)
 Call fStSpei(HCU, RegStelle, "DebugDatei", hlese.snst.DebugDatei)
 Call fStSpei(HCU, RegStelle, "ConStri", hlese.dbv.Constr)
 Call fStSpei(HCU, RegStelle, "ConStr", hlese.dlg.ConstrCn)
 Call fStSpei(HCU, RegStelle, "ConStrLabel", hlese.dlg.ConStrLabel)
' Call hlese.ZeigDateien
End Function ' PubReg

Public Function HolKRein(hlese As Lese)
 Dim rs As ADODB.Recordset
 KRein = myFrag(rs, "SELECT COALESCE((SELECT ob FROM koricht ORDER BY datum DESC LIMIT 1),0)").Fields(0)
End Function ' HolKRein(hlese As Lese)


Function BooleanFld(frm As Lese)
 Dim i As ConDtb, j&, rs As ADODB.Recordset, rs1 As ADODB.Recordset
 Dim AfN&
 On Error GoTo fehler
 For i = qDtb To q2Dtb
'  Call frm.ConstrFestleg(i)
' Call hlese.ConstrFestleg(0, hlese)
  Call acon(quelleT, i)
  Set rs = myEFrag("SHOW TABLES;") 'full columns FROM " & DBCn.DefaultDatabase)
  Do While Not rs.EOF
'   rs.Move 1
   Set rs1 = myEFrag("SHOW FULL COLUMNS FROM `" & rs.Fields(0) & "`")
   Do While Not rs1.EOF
    If rs1!Type = "tinyint(1)" Then
'     Call myEFrag("ALTER TABLE `" & DBCn.DefaultDatabase & "`.`" & rs.Fields(0) & "` " & sqlALTER & " COLUMN `" & rs1!Field & "` Bit(1) DEFAULT NULL COMMENT '" & rs1!Comment & "', ENGINE = InnoDB;", AfN)
     Call myEFrag("ALTER TABLE `" & rs.Fields(0) & "` " & sqlALTER & " COLUMN `" & rs1!Field & "` Bit(1) DEFAULT " & IIf(IsNull(rs1!Default), "NULL", rs1!Default) & " COMMENT '" & rs1!Comment & "', ENGINE = InnoDB;", AfN)
'     Debug.Print AfN
    End If
    rs1.Move 1
   Loop
   rs.Move 1
  Loop
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in BooleanFld/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' BooleanFld(frm AS Lese)

Public Function VergleichTab(frm As Lese) ' Vergleich der Datenbankf¸llungen
 Const SpBr% = 20
 Dim ErrNr&, ErrDes$
 Dim SL As New SortierListe, sF As sFeld, cat As New ADOX.Catalog, Cn(3) As ADODB.Connection, CNs$(3)
' Dim aktCn As New ADODB.Connection
 Dim fld$(), i%, j%, n%, cnA As New ADODB.Recordset
' dim sc As New ADODB.Recordset
 Call machAbK¸
 For i = 0 To 3
'  Call frm.ConstrFestleg(i + 1, frm)
  Set Cn(i) = acon(quelleT, i + 1)
  CNs(i) = Lese.dbv.CnStr
  cat.ActiveConnection = DBCn
  For j = 0 To cat.Tables.COUNT - 1
   If cat.Tables(j).Type = "TABLE" Then
    Set sF = New sFeld
    sF.Feld = cat.Tables(j).name
    Call SL.sCAdd(sF, True)
   End If
  Next j
 Next i
 ReDim Preserve fld(3, SL.COUNT)
 For i = 0 To 3
'  Call frm.ConstrFestleg(i + 1, frm)
'  Call acon(quelleT, i + 1)
   SetDBCn Cn(i), CNs(i)
'  IF i > 0 THEN
'  IF Not cn Is Nothing THEN SET cn = Nothing 'If cn.State = 1 THEN cn.Close
'  SET aktCn = Nothing
'    aktCn.Open DBCn.ConnectionString
'  END IF
  ReDim Preserve fld(3, SL.COUNT)
  For j = 1 To SL.COUNT
   If i > 0 Then
'    ON Error Resume Next
'    myEFrag ("ALTER TABLE " & lcase(SL.item(j).Feld) & " change column abspos AbsPos integer UNSIGNED",, QCn(dbknr))
'    ON Error GoTo fehler
   End If
   If Not cnA Is Nothing Then If cnA.State = 1 Then cnA.Close
   On Error Resume Next
   myFrag cnA, "SELECT COUNT(0) ct FROM " & "`" & LCase$(SL.Item(j).Feld) & "`", , , , , , True, ErrNr, ErrDes
   If ErrNr <> 0 Then GoTo weiter
   On Error GoTo fehler
'   SET sc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, lcase(SL.item(j).Feld), Empty))
   fld(i, j) = cnA!ct
weiter:
  Next j
 Next i
 frm.Ausgabe = frm.Ausgabe & vbCrLf & Right$(Space$(20) & Right$(frm.dlg.MdB, 20), 20) & _
 Right$(Space$(SpBr) & "quelle", SpBr) & _
 Right$(Space$(SpBr) & "quelle1", SpBr) & _
 Right$(Space$(SpBr) & "quelle2", SpBr) & _
 vbCrLf
 frm.Ausgabe = frm.Ausgabe & vbCrLf & "Name Type Attributes DefinedSize" & vbCrLf
  For j = 1 To SL.COUNT
   frm.Ausgabe = frm.Ausgabe & Right$(Space$(20) & LCase$(SL.Item(j).Feld), 20)
    frm.Ausgabe = frm.Ausgabe & _
    Left$(Right$(Space$(16) & fld(0, j), 15) & Space$(SpBr), SpBr) & _
    Left$(Right$(Space$(16) & fld(1, j), 15) & Space$(SpBr), SpBr) & _
    Left$(Right$(Space$(16) & fld(2, j), 15) & Space$(SpBr), SpBr) & _
    Left$(Right$(Space$(16) & fld(3, j), 15) & Space$(SpBr), SpBr) & vbCrLf
    If fld(0, j) = "" And fld(1, j) = "" And fld(2, j) = "" Then GoTo weiter2
weiter2:
  Next j
 frm.Ausgabe = frm.Ausgabe & vbCrLf & "Fertig mit Vergleichen!"
 altAusgabe = frm.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in Vergleiche/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VergleichTab

Public Function Vergleiche(frm As Lese) ' Vergleich der Datenbankstrukturen
 Const SpBr% = 34
 Dim ErrNr&, ErrDes$
 Dim DSo$, Prov$, DSo2$
 Dim aktCn As New ADODB.Connection
 Dim fld$(), i%, j%, n%, k%, l%, cnA As New ADODB.Recordset, maxF%, Cn As New ADODB.Connection
' dim sc As New ADODB.Recordset
 Call machAbK¸
 For i = 0 To 3
'  Call frm.ConstrFestleg(i + 1, frm)
  If i = 0 Then
'   Call acon(quelleT, accDtb, "c:\fjdkalfjdk.mdb")
'  ElseIf i = 1 THEN
   Call acon(quelleT, accDtb, QmdB)
  ElseIf i = 1 Then
   Call acon(quelleT, qDtb, vNS, "quelle4")
  Else
   Call acon(quelleT, i + 0) ' + 1
  End If
'  IF i > 0 THEN
'  IF Not cn Is Nothing THEN SET cn = Nothing 'If cn.State = 1 THEN cn.Close
  Set aktCn = Nothing
  aktCn.Open DBCnS ' DBCn.ConnectionString
'  END IF
  For j = 0 To TbZ2
   If i > 0 Then
'    ON Error Resume Next
'    myEFrag ("ALTER TABLE " & lcase(Tbn(j)) & " change column abspos AbsPos integer UNSIGNED",,cn)
'    ON Error GoTo fehler
   End If
   If Not cnA Is Nothing Then If cnA.State = 1 Then cnA.Close
   On Error Resume Next
   myFrag cnA, "SELECT " & IIf(lies.obMySQL, vNS, "top 1 ") & "* FROM " & "`" & LCase$(tbn(j)) & "`" & IIf(lies.obMySQL, " LIMIT 1", ""), adOpenStatic, aktCn, adLockReadOnly, , , True, ErrNr, ErrDes
   If ErrNr <> 0 Then GoTo weiter
   On Error GoTo fehler
   If cnA.Fields.COUNT > maxF Then: maxF = cnA.Fields.COUNT: ReDim Preserve fld(3, TbZ2, 3, maxF)
'   SET sc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, lcase(Tbn(j)), Empty))
   For k = 0 To cnA.Fields.COUNT - 1
    fld(i, j, 0, k) = cnA.Fields(k).name
    fld(i, j, 1, k) = cnA.Fields(k).Type
    fld(i, j, 2, k) = cnA.Fields(k).Attributes
    fld(i, j, 3, k) = cnA.Fields(k).DefinedSize
   Next k
weiter:
  Next j
  DSo = DSo & Left$(DefDB(aktCn) & Space$(SpBr), SpBr) ' "Data Source Name"
  Prov = Prov & Left$(aktCn.Provider & Space$(SpBr), SpBr)
 Next i
 
' For j = 0 To TbZ2
'  For k = 0 To cnA.Fields.COUNT - 1
'   For l = 0 To 3
'    If fld(1, j, l, k) <> fld(2, j, l, k) Then
'     If l < 3 Then
'     Debug.Print tbn(j), fld(1, j, 0, k), l, fld(1, j, l, k), fld(2, j, l, k)
'     End If
'    End If
'   Next l
'  Next k
' Next j

' frm.Ausgabe = LEFT(frm.dlg.MdB + Space$(SpBr), SpBr) & " " & LEFT("quelle" + Space$(SpBr), SpBr) & " " & LEFT("quelle1" + Space$(SpBr), SpBr) & " " & LEFT("quelle2" + Space$(SpBr), SpBr) + vbCrLf + vbCrLf
 frm.Ausgabe = DSo & vbCrLf & Prov & vbCrLf
 
 frm.Ausgabe = frm.Ausgabe & vbCrLf & "Name Type Attributes DefinedSize" & vbCrLf
  For j = 0 To TbZ2
   frm.Ausgabe = frm.Ausgabe & LCase$(tbn(j)) + vbCrLf
   For k = 0 To maxF
    frm.Ausgabe = frm.Ausgabe & _
    Left$(Left$(fld(0, j, 0, k) + Space$(16), 15) & " " & Right$(Space$(3) & fld(0, j, 1, k), 3) & " " & fld(0, j, 2, k) & " " & Right$(Space$(3) & fld(0, j, 3, k), 3) & " " & Space$(SpBr), SpBr) & _
    Left$(Left$(fld(1, j, 0, k) + Space$(16), 15) & " " & Right$(Space$(3) & fld(1, j, 1, k), 3) & " " & fld(1, j, 2, k) & " " & Right$(Space$(3) & fld(1, j, 3, k), 3) & " " & Space$(SpBr), SpBr) & _
    Left$(Left$(fld(2, j, 0, k) + Space$(16), 15) & " " & Right$(Space$(3) & fld(2, j, 1, k), 3) & " " & fld(2, j, 2, k) & " " & Right$(Space$(3) & fld(2, j, 3, k), 3) & " " & Space$(SpBr), SpBr) & _
    Left$(Left$(fld(3, j, 0, k) + Space$(16), 15) & " " & Right$(Space$(3) & fld(3, j, 1, k), 3) & " " & fld(3, j, 2, k) & " " & Right$(Space$(3) & fld(3, j, 3, k), 3) & " " & Space$(SpBr), SpBr) & vbCrLf
    If fld(0, j, 0, k) = "" And fld(1, j, 0, k) = "" And fld(2, j, 0, k) = "" Then GoTo weiter2
   Next k
weiter2:
  Next j
 frm.Ausgabe = frm.Ausgabe & vbCrLf & "Fertig mit Vergleichen!"
 altAusgabe = frm.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in Vergleiche/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Vergleiche

Function TabFuellSnSh() ' Tabf¸ll; Tabellenf¸llungen ermitteln
 Dim rs As New ADODB.Recordset, Zp As Date, zps$, rsct As New ADODB.Recordset, rAf&, ct&
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 myEFrag "CREATE TABLE IF NOT EXISTS `tabfuell` (id integer(10) auto_increment key, zp datetime, tabname varchar(90), ds integer(10), index tabname(tabname), index zp(zp))"
' rs.Open "SELECT table_name, table_type FROM information_schema.tables WHERE table_schema = '" & CurDB(DBCn) & "'", DBCn, adOpenDynamic, adLockOptimistic
 myFrag rs, "SELECT table_name, table_type FROM information_schema.tables WHERE table_schema = '" & CurDB(DBCn) & "'"
 Zp = Now()
 zps = DatFor_k(Zp)
 Do While Not rs.EOF
  Set rsct = Nothing
  On Error Resume Next
  myFrag rsct, "SELECT COUNT(0) ct FROM `" & rs!table_name & "`", , , , , , True, ErrNr, ErrDes
  If ErrNr <> 0 Then
'   ErrDes = Err.Description
   InsKorr DBCn, DBCnS, "INSERT INTO `tabfuell`(zp,tabname,tabtype,fehler) VALUES(" & zps & ",'" & rs!table_name & "','" & rs!table_type & "','" & ErrDes & "')", rAf
  Else
   InsKorr DBCn, DBCnS, "INSERT INTO `tabfuell`(zp,tabname,tabtype,ds) VALUES(" & zps & ",'" & rs!table_name & "','" & rs!table_type & "'," & rsct!ct & ")", rAf
  End If
  On Error GoTo fehler
  rs.MoveNext
 Loop
 MsgBox "Fertig mit Datensatzz‰hlen f¸r `" & CurDB(DBCn) & "` um: " & zps
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in Tabf¸ll/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TabFuellSnSh
Function Zeichensatz$(Roh$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = vNS
For i = 1 To Len(Roh)
 bq = Mid$(Roh, i, 1)
 Select Case bq
  Case "Ä"
   bz = "_"
  Case "Ø", "Ó"
   bz = "ﬂ"
  Case "ı"
   bz = "‰"
  Case "˜"
   bz = "ˆ"
  Case "≥"
   bz = "¸"
'  Case "-" ' f¸hrt bei den Labor-Langbezeichnungen zu schlechten Ergebnissen
'   bz = "ƒ"
  Case "Õ"
   bz = "÷"
  Case "_"
   bz = "‹"
  Case "¡"
   bz = "µ"
  Case "ﬂ"
   bz = "·"
  Case "”"
   bz = "‡"
  Case "⁄"
   bz = "È"
  Case "ﬁ"
   bz = "Ë"
  Case "›"
   bz = "Ì"
  Case "˝"
   bz = "Ï"
  Case "æ"
   bz = "Û"
  Case "="
   bz = "Ú"
  Case "∑"
   bz = "˙"
  Case "®"
   bz = "˘"
  Case "‘"
   bz = "‚"
  Case "€"
   bz = "Í"
  Case "Ø"
   bz = "Ó"
  Case "∂"
   bz = "Ù"
  Case "π"
   bz = "˚"
  Case "¡"
   bz = "µ"
  Case "¶"
   bz = "≤"
  Case "¶"
   bz = "≥"
  Case "∞"
   bz = "ß"
  Case "«"
   bz = "Ä"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
Zeichensatz = ZS
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in Zeichensatz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Zeichensatz
Function ZSU1$(Roh$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = vNS
For i = 1 To Len(Roh)
 bq = Mid$(Roh, i, 1)
 Select Case bq
  Case "·"
   bz = "ﬂ"
  Case "Ñ"
   bz = "‰"
  Case "î"
   bz = "ˆ"
  Case "Å"
   bz = "¸"
  Case "é"
   bz = "ƒ"
  Case "ô"
   bz = "÷"
  Case "ö"
   bz = "‹"
  Case "†"
   bz = "·"
  Case "Ö"
   bz = "‡"
  Case "Ç"
   bz = "È"
  Case "ä"
   bz = "Ë"
  Case "°"
   bz = "Ì"
  Case "ç"
   bz = "Ï"
  Case "¢"
   bz = "Û"
  Case "ï"
   bz = "Ú"
  Case "£"
   bz = "˙"
  Case "ó"
   bz = "˘"
  Case "É"
   bz = "‚"
  Case "à"
   bz = "Í"
  Case "å"
   bz = "Ó"
  Case "ì"
   bz = "Ù"
  Case "ñ"
   bz = "˚"
  Case "Ê"
   bz = "µ"
  Case "˝"
   bz = "≤"
  Case "¸"
   bz = "≥"
  Case "ı"
   bz = "ß"
'  Case "_" ' Ä wird zwar als _ dargestellt, aber _ auch
'   bz = "Ä"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
ZSU1 = ZS
'If ZSU1 <> roh THEN
' Debug.Print "ZSU: " + ZSU1 + " -> " + roh
'END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ZSU1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU1

#If False Then
'' Stellt in der Datenbank programmst‰nde.Lauf den Wert Abbruch auf 'obabbruch', falls er nicht schon darauf steht und falls nicht "Nicht‰ndern" eingegeben wurde
'' Ansonsten liefert die Funktion "true"
'Public FUNCTION ProgrammLauf(obAbbruch%, Optional Cpt$, Optional Nichtƒndern%) AS Boolean
' Dim i%, plConn As New ADODB.Connection, rsA As Adodb.Recordset, lcConStr$
' Dim cat As New ADOX.Catalog
' Dim Pfad$
' IF DBCn.State = 1 THEN
' ON Error Resume Next
' For i = 1 To 6
'' a) Verbindung aufbauen
'  IF lies.obMySQL THEN
'   plConn.Open DBCn.ConnectionString
'   plConn.Execute ("use programmstaende;")
'  Else
'   Pfad = Mid$(DBCn, InStr(DBCn, "Source=") + 7)
'   Pfad = LEFT(Pfad, InStr(Pfad, ";"))
'   Do While Right$(Pfad, 1) <> "\"
'    Pfad = LEFT(Pfad, Len(Pfad) - 1)
'   Loop
'   lcConStr = ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
'               "Data Source=" & Pfad + "programmstaende" + ".mdb;" & _
'               "Jet OLEDB:Engine Type=5")
'   plConn.Open lcConStr
'  END IF
'' b) falls nˆtig, Datenbank erstellen
'  IF Err.Number <> 0 THEN
'   IF lies.obMySQL THEN
'      Call plConn.Execute("CREATE DATABASE programmstaende;")
'   Else
'      cat.Create lcConStr
'#If False THEN
'      Dim WS, DBE
'      SET DBE = New DAO.DBEngine
'      SET WS = DBE.Workspaces(0)
'      Call WS.CreateDatabase(Pfad + "programmstaende" + ".mdb", dbLangGeneral)
'#END IF
'   END IF
'  ELSE ' err.number <> 0 nach connection
'   SET rsA = New ADODB.Recordset
'   Call rsA.Open("SELECT Computer,Programm,-Abbruch AS j_abbruch FROM Lauf WHERE programm = ""DateiLese"";", plConn, adOpenDynamic, adLockOptimistic)
'   IF Err.Number = 0 AND NOT rsA.BOF THEN
'    Exit For
'   ELSE ' Err.Number = 0 AND NOT rsA.BOF THEN
'    Err.Clear
'    Call plConn.Execute("INSERT INTO Lauf (Programm) VALUES (""DateiLese"");")
'    IF Err.Number = 0 THEN
'     Exit For
'    ELSE '  Err.Number = 0 THEN
'     Err.Clear
'     Call plConn.Execute("CREATE TABLE Lauf (Programm varchar(30), abbruch bit, Computer varchar(50))")
'     IF Err.Number <> 0 THEN
'      MsgBox "Unbehebbarer Fehler bei Programmlauf"
'     END IF
'    END IF ' Err.Number = 0 THEN
'   END IF ' Err.Number = 0 AND NOT rsA.BOF THEN
'  END IF ' err.number <> 0 nach connection
' Next i
' ON Error GoTo fehler
' Cpt = IIf(ISNULL(rsA!Computer), vns, rsA!Computer)
' IF Abs(rsA!j_Abbruch) = Abs(obAbbruch) THEN
'  ProgrammLauf = True
' ElseIf Not Nichtƒndern THEN
'   IF obAbbruch THEN BrichAb = True
'   rsA!Computer = CptName
'   rsA!j_Abbruch = obAbbruch
'   rsA.Update
' END IF
' END IF ' DBCn.state
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ProgrammLauf/" + AnwPfad)
' Case vbAbort: Call MsgBox("Hˆre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' ProgrammLauf(obAbbruch%, Optional Nichtƒndern%) AS Boolean
#End If

Function adoxtest(dlg As Dialog)
 Dim i As ConDtb, j&, k&, l&, CC&
 Dim cat As New ADOX.Catalog
 Dim Tb As New ADOX.Table
  On Error GoTo fehler
  For i = qDtb To qDtb
'  Call dlg.hlese.ConstrFestleg(i, dlg.hlese)
  Call acon(quelleT, i)
  cat.ActiveConnection = DBCn
  For k = 1 To cat.Tables.COUNT - 1
   CC = 0
   On Error Resume Next
   CC = cat.Tables(k).Columns.COUNT
   On Error GoTo fehler
   If CC > 0 Then
    For j = 1 To CC - 1
'     Debug.Print cat.Tables(k).Columns(j).name, cat.Tables(k).Columns(j).Attributes, cat.Tables(k).Columns(j).Properties.COUNT
'     For l = 0 To cat.Tables(k).Columns(j).Properties.COUNT - 1
'      Debug.Print cat.Tables(k).Columns(j).Properties(l).name, cat.Tables(k).Columns(j).Properties(l).Value
'     Next l

'   Debug.Print cat.Tables(1).Columns(j).SortOrder
'    Debug.Print cat.Tables(k).Columns(j).Properties.Count
'    Debug.Print cat.Tables(k).Columns(j).Attributes
    Next j
   End If
  Next k
 Next
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in adoxtest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' adoxtest

Function getDokPfad$(Optional Abschnitt$)
 On Error GoTo fehler
 Dim idt As TMIniDatei, erg$
 If LenB(Abschnitt) <> 0 Or IsNull(PcDokPfad) Or LenB(PcDokPfad) = 0 Then
  If LenB(Abschnitt) = 0 Then Abschnitt = "Dokumente"
  Set idt = New TMIniDatei
  getDokPfad = idt.GetProp("Verzeichnisse/TurboMed/" & Abschnitt, "Pfad")
  If getDokPfad = "" Then
   MsgBox "Bitte c:\turbomed\Programm\local.ini auf Existenz pr¸fen und NVIni laufen lassen!"
   ProgEnde
  End If
  Dim p1%
  p1 = InStr(getDokPfad, "\\\")
  If p1 <> 0 Then getDokPfad = Mid$(getDokPfad, p1 + 1)
  erg = Dir(getDokPfad, vbDirectory)
  If LenB(erg) = 0 Then
   getDokPfad = REPLACE$(LCase$(getDokPfad), "x:", LiServer & "turbomed")
   erg = Dir(getDokPfad, vbDirectory)
  End If
  If LenB(erg) <> 0 Then If (GetAttr(getDokPfad) And vbDirectory) <> vbDirectory Then erg = vNS
  If LenB(erg) = 0 Or InStrB(getDokPfad, "\") = 0 Then
   getDokPfad = idt.GetProp("Verzeichnisse/TurboMed", "Pfad") & getDokPfad
  End If
  If Abschnitt = "Dokumente" Then PcDokPfad = getDokPfad
 Else
  getDokPfad = PcDokPfad
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in getDokPfad/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getDokPfad

Public Function Key(KeyCode%, Shift%, frm As Form, Optional ctrl$)
 Static keyzl%
 Dim frmart As ArtTyp
 On Error GoTo fehler
' IF KeyCode = 18 AND Shift = 4 THEN Exit Function
If KeyCode = 27 Then
' ƒnderungen verwerfen, indem Ursprungswerte aus der Registry geholt werden,
' bevor die aktuellen Werte ¸ber form_unload wieder zur¸ckgespeichert werden
    Select Case frm.name
     Case "MachDatenbank", "AbrechFehler"
      Unload frm
      GoTo Ende
     Case "Lese": ' Call HolReg(frm)
      Unload frm
      End
     Case "PatAuswahl":
     Case "PatListe":
       If frm.Text1.Visible Then
        frm.Text1.Visible = False
        frm.mfg_entercell
        GoTo Ende
       ElseIf frm.Li1.Visible Then
        frm.Li1.Visible = False
        frm.mfg_entercell
        GoTo Ende
       End If
     Case Else: ' Call HolReg(frm.hlese)
    End Select
    frm.Hide
    lies.Show
 ElseIf KeyCode = 13 Then
  If frm.name = "PatAuswahl" And (frm.ActiveControl.name = "PatAuswahl" Or frm.ActiveControl.name = "Pat_id") Then
'  frm.Pat_id = frm.getPat_id(frm.PatName)
  ElseIf frm.name = "PatListe" Then
   If frm.PLArt = artDiag Then
    Call frm.MFG_Click
   End If
  End If
 ElseIf KeyCode = 70 And ((Shift And vbCtrlMask) > 0) And frm.name = "PatListe" Then
  Call frm.Command1_Click(3) ' Suchen
  KeyCode = 0
 ElseIf KeyCode = 71 And ((Shift And vbCtrlMask) > 0) And frm.name = "PatListe" Then
  Call frm.Command1_Click(4) ' Weitersuchen
' ElseIf KeyCode = 67 THEN ' Ctrl+C, in MFG bei Patliste bereits bearbeitet
 ElseIf (KeyCode = 113 Or KeyCode >= Asc("a") And KeyCode <= Asc("z")) Or (KeyCode >= Asc("A") And KeyCode <= Asc("Z") Or (KeyCode >= Asc("0") And KeyCode <= Asc("9"))) And ((Shift And vbCtrlMask) = 0) And frm.name = "PatListe" And frm.ActiveControl.name <> "Li1" And frm.ActiveControl.name <> "Text1" Then
  On Error Resume Next
  frmart = frm.PLArt
  On Error GoTo fehler
  If frmart = artDiag Then
   If KeyCode <> 113 Then frm.AnfCode = Asc(LCase$(Chr$(KeyCode)))
   Call frm.MFG_Click
   KeyCode = 0
  End If
 ElseIf KeyCode = 32 Then
  If frm.name = "AbrechFehler" Then
   Call frm.MFG_Click
  ElseIf frm.name = "PatListe" Then
  If frm.PLArt = artPat And frm.MFG.col = zahlcol Then
    If keyzl = 0 Then keyzl = 1 Else keyzl = 0
    If keyzl = 1 Then
     frm.MFG.CellFontBold = Not frm.MFG.CellFontBold
     Dim bc&
     bc = frm.MFG.CellBackColor
     If bc = vbYellow Then
       frm.MFG.CellBackColor = vbWhite
       frm.GesZl = frm.GesZl - 1
       Dim iz&
       For iz = 1 To frm.GesColl.COUNT
        If frm.GesColl(iz) = frm.MFG.Text Then
         frm.GesColl.Remove (iz)
         Exit For
        End If
       Next iz
     Else
       frm.MFG.CellBackColor = vbYellow
       frm.GesZl = frm.GesZl + 1
       frm.GesColl.Add frm.MFG.Text
     End If
     frm.Command1(14).Caption = frm.GesZl & " &Ges"
    End If
   End If
  End If
 ElseIf KeyCode <> 16 And KeyCode <> 17 And KeyCode <> 18 Then
'  Debug.Print "KeyCode in Haupt.Key: " & KeyCode
 End If
' IF KeyCode = 33 THEN Call doR¸ckw‰rtsCmd(frm)
' IF KeyCode = 34 THEN Call doVorw‰rtsCmd(frm) <- stellt den aktuellen Feldinhalt falsch ein!
Ende:
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in key/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Key

Function ForeignNo0()
 Dim runde%
 On Error GoTo fehler
 If ForeignKAus0 = 0 Then
  If lies.obMySQL Then
'   DBCn.Execute "SET foreign_key_checks = 0"
   Call myEFrag("SET foreign_key_checks = 0")
  Else
   ZielDbS = Lese.dlg.MdB
   Call BezLˆschA
  End If
  ForeignKAus0 = 1
 End If ' ForeignKAus0 = 0 Then
 Exit Function
fehler:
ErrNumber = Err.Number
runde = runde + 1
If runde < 10 And Err.Number = -2147467259 Then ' The Server has gone away
 DBCnS = Lese.dbv.CnStr
' IF LenB(DBCnS) = 0 THEN DBCnS = DBCn.ConnectionString
' SetDBCn Nothing
 DBCn.Close ' 28.8.23
 DBCnOpen DBCnS
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in ForeignNo0/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ForeignNo()

Function ForeignNo1()
 If ForeignKAus1 = 0 Then
  If lies.obMySQL Then
'   Dim rAf&
'   DBCn.Execute "SET foreign_key_checks = 0", rAf
   Call myEFrag("SET foreign_key_checks = 0", rAf)
  Else
   ZielDbS = Lese.dlg.MdB
   Call BezLˆschA
  End If ' lies.obMySQL Then
  ForeignKAus1 = 1
 End If ' ForeignKAus1 = 0 Then
End Function ' ForeignNo()

Function ForeignYes0()
 If ForeignKAus0 = 1 Then
  If lies.obMySQL Then
   Call myEFrag("SET foreign_key_checks = 1")
'   DBCn.Execute "SET foreign_key_checks = 1"
  Else
   ZielDbS = Lese.dlg.MdB
   Call BezHerstA
  End If ' lies.obMySQL Then
  ForeignKAus0 = 0
 End If ' ForeignKAus0 = 1 Then
End Function ' ForeignNo()

Function ForeignYes1()
 If ForeignKAus1 = 1 Then
  If lies.obMySQL Then
   Call myEFrag("SET foreign_key_checks = 1")
'   DBCn.Execute "SET foreign_key_checks = 1"
  Else
   ZielDbS = Lese.dlg.MdB
   Call BezHerstA
  End If
  ForeignKAus1 = 0
 End If
End Function ' ForeignNo()

' 21.9.08 aus AccDBMach kopiert
Public Sub BezLˆschA()
 Dim rel, i%, j% ', ZielDB AS DAO.Database
 On Error GoTo fehler
#If Not ohneDAO Then
 Dim ZielDB As DAO.Database
 Set ZielDB = OpenDatabase(ZielDbS)
 For i = 1 To 10
  For Each rel In ZielDB.Relations
   ZielDB.Relations.Delete rel.name
  Next rel
 Next i
 ZielDB.Close
#Else
 Dim cat As New ADOX.Catalog
 Dim Tbl As ADOX.Table
' Dim fk As New ADOX.Key
' SetDBCn Nothing
 QuelleMDB = uVerz & "anamnese\quelle.mdb"
 DBCnS = aCStr(quelleT, accDtb, QuelleMDB, "")
 DBCnOpen DBCnS
 Call Zinit(obMySQL:=False)
 cat.ActiveConnection = DBCn ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & ZielDbS & "'"
 For Each Tbl In cat.Tables
  If Tbl.Type = "TABLE" Then
   For j = Tbl.Keys.COUNT To 1 Step -1
    If Tbl.Keys(j - 1).Type = adKeyForeign Then
     Call Tbl.Keys.Delete(j - 1)
    End If
   Next j
  End If
 Next Tbl
#End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in BezLˆschA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' BezLˆschA()

Public Function BezHerstA()
 ' 1: 1:1-Beziehung (dbRelationUnique)
 ' 2: ref.Integr.nicht erzwungen (DbRelationDontEnforce)
 ' 4: Bez zw 2 verkn Tab aus AND DB (dbRelationInherited)
 ' 256: Aktualisierungen weitergegeben (dbRelationUpdateCascade)
 ' 4096: Lˆschvorg‰nge weitergegebn (dbRelationDeleteCascade)
 ' 16777216: LEFT JOIN standardm angez (dbRelationLeft)
 ' 33554432: Righ JOIN standardm angez (dbRelationRight)
 
' Dim rel, ZielDB AS DAO.Database
' ON Error GoTo fehler
' Call syscmd(acSysCmdSetStatus, "Beziehungen herstellen")
' SET ZielDB = OpenDatabase(ZielDbS)
' ZielDB.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in BezHerstA/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' BezHerstA()

#If False Then
Function holAB(frm As Lese) ' kommt vor in AnamnesebogenHolen_Click
 Dim rq As New ADODB.Recordset ', f1 AS ADODB.Field, f2 AS ADODB.Field,
 Dim runde&, Pat_id&
 frm.BytesBez = "Datens‰tze:"
 myFrag rq, "SELECT COUNT(0) ct FROM `anamnesebogen`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
 frm.GesBytes = rq!ct
 On Error GoTo fehler
' DBCn.BeginTrans: obTrans = 1
 BegTrans
 Call ForeignNo0
 Call ForeignNo1
 rq.Close
 myFrag rq, "SELECT * FROM `anamnesebogen`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
 Do While Not rq.EOF
  runde = runde + 1
  If Not rsAnm Is Nothing Then If rsAnm.State = 1 Then rsAnm.Close
  rsAnm.CursorLocation = adUseClient
  myFrag rsAnm, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & rq!Pat_id, adOpenStatic, DBCn, adLockOptimistic
  If Not rsAnm.BOF Then
   Call anaUpd(rq, rsAnm)
  Else
   Call anaIns(rq)
  End If
  frm.Bytes = runde
  DoEvents
'  ON Error GoTo f1
'  rsAnm.Update
'  ON Error GoTo f0
  If BrichAb Then Exit Do
  rq.Move 1
 Loop
 Call ForeignYes0
 Call ForeignYes1
 If BrichAb Then
  frm.Ausgeb "Kopieren der Anamnesebˆgen von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "' abgebrochen.", True
 Else
  frm.Ausgeb "Fertig mit Kopieren der Anamnesebˆgen von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "'", True
 End If
 On Error Resume Next
' If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
 ComTrans
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in holAB/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' holAb
#End If

Function doLdFD() ' Liste der fehlenden Dokumente
 Dim Datei$
 Datei = uVerz & "Fehlende Dokumente.txt"
 Dim Ers$(20), i%, Zp$, tStr$
 Dim aktDP$
 Ers(0) = xVerz & "Dokumente"
 Ers(1) = "\\Anmeld2\TMServ\Dokumente"
 Ers(2) = "\\anmeld2\E\TurboMed\Dokumente"
 Ers(3) = "\\Sono\Daten (F)\daten\Turbomed\Dokumente"
 Ers(4) = "\\Sono\Daten (F)\daten\Papierkorb\Turbomed\Dokumente"
 Ers(5) = zVerz & "Papierkorb\Turbomed\Dokumente"
 Ers(6) = "\\server\TurboMed\Dokumente"
 Ers(7) = "c:\turbomed\dokumente"
 Ers(8) = "c:\opt\turbomed\dokumente"
 Ers(9) = LiServer & "shome\turbomed\Dokumente"
 Ers(10) = "\\server\TurboMed\Dokumentekaputt"
 Ers(11) = "\\labor\WD USB2 (E)\Turbomed gespiegelt\Dokumente\"
 Ers(12) = "\\labor\WD USB2 (E)\Turbomed gespiegelt\Dokumentekaputt\"
 Ers(13) = "\\sono\daten (f)\daten\shome\turbomed\dokumente"
 Ers(14) = "\\sono\daten (f)\daten\gerald\turbomed\dokumente"
 Ers(15) = "\\sono\daten (f)\ausgelagert\Papierkorb\turbomed\dokumente"
 Ers(16) = "\\mitte\Turbomed\Dokumente\Video"
 Ers(17) = "\\anmeld2\Volume (F)\Turbomed gespiegelt\Dokumente"
 Ers(18) = "\\anmeld2\E\TurboMed\Dokumente"
 If LenB(PcDokPfad) = 0 Then getDokPfad
 Dim rAb As New ADODB.Recordset
 Open Datei For Output As #322
' rAb.Open "SELECT * FROM dokumente", DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, "SELECT * FROM dokumente"
 Do While Not rAb.EOF
  aktDP = REPLACE$(LCase$(rAb!DokPfad), "$\turbomed\dokumente", PcDokPfad)
  If Not FSO.FileExists(aktDP) Then
   For i = 16 To 18
    Zp = REPLACE$(aktDP, PcDokPfad, Ers(i))
    If FSO.FileExists(Zp) Then
     FileCopy Zp, rAb!DokPfad
     Print #322, "Kopiere " & Zp & " -> " & aktDP
    End If
   Next i
'   Debug.Print aktDP
   Print #322, rAb!Pat_id, rAb!Zeitpunkt, aktDP, rAb!DokName
   Lese.Ausgabe = Lese.Ausgabe & vbCrLf & rAb!Pat_id & " " & rAb!Zeitpunkt & " " & aktDP & vbTab & rAb!DokName
   tStr = sucheinVerz(rAb!DokName, pVerz)
   If tStr <> "" Then
    Print #322, "     gefunden in:", tStr
    Lese.Ausgabe = Lese.Ausgabe & vbCrLf & "      gefunden in: " & tStr
   End If
   DoEvents
  End If
  rAb.Move 1
 Loop
 Close #322
 zeigan Datei
 MsgBox "Fertig!"
End Function ' doLdFD

Function sucheinVerz$(Datei$, Verz$)
 Dim SFold As Folder, tStr$
 If Right$(Verz, 1) <> "\" Then Verz = Verz & "\"
 If FSO.FileExists(Verz & Datei) Then
  sucheinVerz = Verz & Datei
 Else
  For Each SFold In FSO.GetFolder(Verz).SubFolders
   tStr = sucheinVerz(Datei, SFold.path)
   If tStr <> "" Then
    sucheinVerz = tStr
    Exit For
   End If
  Next SFold
 End If
End Function ' sucheinVerz
'Function doDMPListeAnzeigen(frm AS Lese)
' Const DN$ = uverz & "DMP-Liste.txt"
' Const sql$ = "SELECT MAX(zp) AS Zp, MIN(MID(abk,8,1)) AS Typ, MIN(erstd) AS Erstd, MAX(LEFT(nachname,13)) AS NachName, MAX(LEFT(vorname,13)) AS Vorname,MAX(pat_id) AS PAT_id FROM dmpreihe GROUP BY pat_id ORDER BY MIN(MID(abk,8,1)), pat_id"
' 'Const sql$ = "SELECT * FROM folgedokus"
' Dim rAb As New ADODB.Recordset, lfdnr&
' Open DN For Output AS #322
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
' Do While Not rAb.EOF
'  lfdnr = lfdnr + 1
'  Print #322, lfdnr, rAb!Typ, format$(rAb!zp, "dd/mm/yy"), rAb!ErstD, rAb!NachName, rAb!VorName, rAb!Pat_id
'  frm.Ausgabe = frm.Ausgabe & " " & lfdnr & " " & rAb!Typ & " " & format$(rAb!zp, "dd/mm/yy") & " " & rAb!ErstD & " " & rAb!NachName & " " & rAb!VorName & " " & rAb!Pat_id & vbCrLf
'  rAb.Move 1
' Loop
' Close #322
' Call Shell("notepad " & DN, vbMaximizedFocus)
' MsgBox "Fertig!"
'End Function
'
'Function testha()
' Lese.ProgStart
' Dim rs As New ADODB.Recordset
' myFrag rs, "SELECT ¸bwv, ¸bwvkvnr, ¸bwvlanr, ¸bwvbsnr FROM faelle"
' ' => alle genannten sind bisweilen "", aber nie null
' Do While Not rs.EOF
'  IF ISNULL(rs!‹bWVLANR) THEN
'   Stop
'  ElseIf rs!‹bWVLANR = "" THEN
''   Stop
'  END IF
'  rs.MoveNext
' Loop
'End Function

' nur in Sub Haus‰rzteBKK_Click()
Function doHABKK(frm As Lese)
 Const sql$ = "SELECT COUNT(0) Zahl, CONCAT(l.name,' ',l.vorname,' ',l.fachgruppe,' ',l.zulg,' ',l.ort) Hausarzt, l.telefon, l.fax, ¸bwr ¸w FROM `aktf` a LEFT JOIN `faelle` f ON a.fid = f.fid LEFT JOIN `kassenliste` k ON f.ik = k.ik AND f.vknr = k.vknr LEFT JOIN `aktlue` l ON ¸bwr = l.kvnro WHERE kateg = 'BKK' AND ¸bwr <> '' AND l.name <> 'Schade' GROUP BY ¸w ORDER BY Zahl DESC;"
 Dim rAb As New ADODB.Recordset
 Open uVerz & "Haus‰rzte BKK.txt" For Output As #322
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, sql
 TabAusgeb rAb, Lese, True, , , , Array(0, 90, 0, 0, 0, 0), True, uVerz & "HA_BKK.txt"
End Function ' doHABKK

Function doLdFHalt(frm As Lese) ' Liste der fehlenden Haus‰rzte
 Const sql$ = "SELECT n.Pat_id, n.Nachname, n.Vorname, n.Gebdat, Schgr FROM `namen` n LEFT JOIN `aktfv` ON `aktfv`.pat_id = n.pat_id WHERE n.kvnr = '' AND NOT ISNULL(`aktfv`.pat_id) ORDER BY `namen`.pat_id DESC;"
 Dim rAb As New ADODB.Recordset
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, sql
' TabAusgeb rAb, Lese, True, , , , , , "Fehlende Haus‰rzte.txt"
 Open uVerz & "Fehlende Haus‰rzte.txt" For Output As #322
 Do While Not rAb.EOF
  Print #322, rAb!Pat_id, rAb!Nachname, rAb!Vorname, rAb!GebDat
  frm.Ausgabe = frm.Ausgabe & rAb!Pat_id & " " & rAb!Nachname & " " & rAb!Vorname & " " & rAb!GebDat & vbCrLf
  rAb.Move 1
 Loop
 Close #322
 MsgBox "Fertig!"
End Function ' doldfH

' nur in Fehlende‹berweisungsscheine_Click()
Function doF‹wS(frm As Lese) ' fehlende ‹berweisungsscheine
 Dim sql$, AktQ$
 AktQ = ZQuart(Now() - Versp‰tung)
 sql = "SELECT n.Pat_id, gesname(n.pat_id) Name, GROUP_CONCAT(DISTINCT CAST(f.schgr AS char)) Schgr, LEFT(GROUP_CONCAT(DISTINCT CONCAT_WS(', ', l2.Name, LEFT(l2.vorname,1),l2.ort, l2.telefon)),31) ‹berweiser, LEFT(CONCAT_WS(', ',l.Name, LEFT(l.vorname,1),l.ort, l.telefon),31) Hausarzt FROM `aktf` f LEFT JOIN `briefe` b ON f.pat_id = b.pat_id AND b.name LIKE '%¸w%' AND b.name LIKE '%" & Left$(AktQ, 1) & "%' AND b.name LIKE '%" & Right$(AktQ, 2) & "%' LEFT JOIN `namen` n ON f.pat_id = n.pat_id LEFT JOIN `aktlue` l ON l.kvnr = n.kvnr LEFT JOIN `faelle` f2 ON f2.fid = f.fid AND ¸bwr <> '' LEFT JOIN `aktlue` l2 ON l2.kvnro = ¸bwr WHERE ISNULL(b.name) AND f.schgr<>'0' GROUP BY f.pat_id"
 Dim rAb As New ADODB.Recordset
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, sql
 TabAusgeb rAb, Lese, True, , , , , , "Fehlende ‹berweisungsscheine"
End Function ' doF‹wS

Function doLdFH(frm As Lese) ' Liste der fehlenden Haus‰rzte
' Const sql$ = "SELECT v.*,n.Nachname,n.Vorname,Gebdat, d.icd, IF(¸bwv = '', and¸w, ¸bwv) AS ¸w, k.kvnr, n.notiz FROM `aktfv` v LEFT JOIN `namen` n ON v.pat_id = n.pat_id LEFT JOIN `faelle` f ON v.fid = f.fid LEFT JOIN `diagnosen` d ON v.pat_id = d.pat_id AND d.icd LIKE 'E1%' LEFT JOIN kvnrue k ON v.pat_id = k.pat_id WHERE notiz = '' OR ISNULL(notiz) OR k.kvnr = '' OR ISNULL(k.kvnr) GROUP BY v.pat_id ORDER BY v.pat_id, k.lfdnr"
 Const sql$ = "SELECT n.Pat_id, n.Nachname, n.Vorname, n.Gebdat, `aktfv`.Schgr, KVNr FROM `namen` n LEFT JOIN `aktfv` ON `aktfv`.pat_id = n.pat_id WHERE n.kvnr IN ('" & KVNr & "') AND NOT ISNULL(`aktfv`.pat_id) ORDER BY n.kvnr, n.pat_id DESC"
 Dim rAb As New ADODB.Recordset
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, sql
 TabAusgeb rAb, Lese, True, , , , , , "PatientenMitUnsAlsHausarztOhneDieMitFehlendemHausarzt"
End Function ' doldfH

Function doSuchTel(frm As Lese) ' suche Telefonnummer
 Dim tel$
 tel = InputBox("Teile der gesuchten Tel'nr:", "Telefonnummernsuche")
 Dim sql$, telm$
 telm = "'%" & tel & "%'"
 sql$ = "SELECT n.Pat_id, gesname(n.Pat_id) `Name mit Telnr. " & tel & "`, PrivatTel, Privattel_2, Diensttel, PrivatMobil, PrivatFax FROM `namen` n WHERE privattel LIKE " & telm & " OR privattel_2 LIKE  " & telm & " OR privatfax LIKE  " & telm & " OR diensttel LIKE  " & telm & " OR privatmobil LIKE  " & telm
 Dim rAb As New ADODB.Recordset
' rAb.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rAb, sql
 TabAusgeb rAb, Lese, True, , , , , , "Suche Telefonnummer " + tel
End Function ' doSuchTel

Function ergEBM(frm As Lese)
 Dim QDat$, Text$, Spli$(), dszahl&, rAf&
' Zeilenumbr¸che f¸hren zu Fehlern, ‰hnlich bei ".csv" und ";" statt ".txt" und ";"
#If False Then
 QDat = getLDatei(uVerz, "Listenausgabe_EBM-Ziffern*.txt")
 Open QDat For Input As #281
 For i = 1 To 5
  Input #281, Text
 Next i
 Do While Not EOF(281)
  Input #281, Text
  Spli = Split(Text, vbTab)
  If UBound(Spli) > 0 Then
'   Debug.Print spli(0)
  End If
 Loop
 Close #281
#End If
 On Error Resume Next
 QDat = getLDatei(uVerz, "Listenausgabe_EBM-Ziffern*.xls")
 If Err.Number <> 0 Then
  MsgBox "Fehler beim Einladen der Datei " & uVerz & "Listenausgabe_EBM-Ziffern*.xls"
  Exit Function
 End If
 On Error GoTo fehler
' aus FUNCTION EmailsImport(EmDatei$)
 Dim con As New ADODB.Connection  ' Connection
 Dim rNa As New ADODB.Recordset
 Dim rEx As New ADODB.Recordset
 Dim rX As New ADOX.Catalog
' Const EmDatei$ = pverz & "Patienten¸bergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
 On Error GoTo fehler
 If LenB(DBCn) = 0 Or LenB(DBCn) = 0 Then
   Dim dlg As New Dialog
'   Call frm.ConstrFestleg(2)
   Call acon(quelleT, qDtb)
 End If
 If QDat <> "" Then
 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=1"";Data Source=" & QDat & ";" ' TABLE=Adressen$"
 Dim runde%, i%, zFeld$, lFeld$, obAnfang%, pNr%, pRoh
  rX.ActiveConnection = con
  rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", con ' Hier Excel, nicht lies.obmysql = 0!
  Do While Not rEx.EOF
'  Debug.Print runde
   If obAnfang Then
'    ON Error Resume Next
    pRoh = rEx.Fields(0)
'    ON Error GoTo fehler
    If pRoh <> "" Then
'      Call myEFrag("SELECT * FROM `ebm2000plus` WHERE leistung = '" & pRoh & "'", rAf)
      Set rNa = Nothing
      myFrag rNa, "SELECT 0 FROM `ebm2000plus` WHERE leistung IN ('" & pRoh & "','0" & pRoh & "')"
      If rNa.BOF Then
'       Debug.Print pRoh
       Dim euro#, punkte#, pos&, FldI
       euro = 0
       punkte = 0
       FldI = rEx.Fields(2)
        pos = InStr(FldI, "|")
        If pos > 0 Then
         FldI = Mid$(FldI, pos + 1)
        End If
       pos = InStr(FldI, " Punkte")
       If pos > 0 Then
        punkte = CDbl(Left$(FldI, pos - 1))
       Else
        pos = InStr(FldI, " Euro")
        If pos > 0 Then
         euro = CDbl(Left$(FldI, pos - 1))
        End If
       End If
       InsKorr DBCn, DBCnS, "INSERT INTO `ebm2000plus`(Leistung,Titel,punktwert,euro) VALUES('" & pRoh & "','" & REPLACE$(rEx.Fields(1), "'", "''") & "','" & REPLACE$(CStr(punkte), ",", ".") & "','" & REPLACE$(CStr(euro), ",", ".") & "')", rAf
       dszahl = dszahl + 1
      Else
'       Debug.Print pRoh, rNa!Leistung
      End If
    End If
'      Call myEFrag("UPDATE `namen` SET email = '" & rEx.Fields(zFeld) & "' WHERE pat_id = " & pNr, rAf)
 ' wenn Name noch nicht da
'     IF rAf <> 1 THEN Err.Raise 999, , "Fehler beim Einf¸gen der Email-Adresse '" & rEx.Fields(eFeld) & "' zum Patienten Nr. " & pNr
    ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) And Not IsNull(rEx.Fields(2)) Then
     obAnfang = True
   End If
  runde = runde + 1
  rEx.MoveNext
  Loop
 End If
 On Error Resume Next
 rEx.Close
 frm.Ausgeb "Fertig mit Nachtragen aus '" & QDat & "' von " & dszahl & " Datens‰tzen ", True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in EmailsImport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ergEBM(frm As Lese)

#If uralt Then
'' nur noch in holEBM2000plus_Click
'Function holAllg(frm As Lese, TabN$, Optional sF$, Optional obPS%) ' Allgemeines / Haus‰rzte ¸bertragen
'' 21.3.21: nur noch in holEBM2000plus_Click
'' kommt vor in: holLaborParameter_Click(), holEBM2000plus_Click(), holMedArten_Click(), KassenlisteKopieren_Click()
' Dim sqldel$, rAf&, GesZahl&
'' Const TabN$ = "kassenliste"
'' Const SF$ = "ID"
' Dim rq As New ADODB.Recordset ', f1 AS ADODB.Field, f2 AS ADODB.Field,
' Dim runde&, i%
' frm.BytesBez = "Datens‰tze:"
' AccOpen rq, "SELECT COUNT(0) AS ct FROM " & "`" & TabN & "`", frm.snst.QuelleAnamneseBˆgen
'' rq.Open "SELECT COUNT(0) AS ct FROM " & "`" & TabN & "`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
' frm.GesBytes = rq!ct
' If frm.GesBytes > 1 Then ' And LCase$(CStrAcc & frm.snst.QuelleAnamneseBˆgen) <> LCase$(DBCnS) Then ' DBCn.ConnectionString
'  On Error GoTo fehler
''  DBCn.BeginTrans: obTrans = 1
'  BegTrans
'  Call ForeignNo0
'  Call ForeignNo1
'  AccOpen rq, TabN, frm.snst.QuelleAnamneseBˆgen
''  rq.Close
''  rq.Open "`" & TabN & "`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
'  Call myEFrag("DELETE FROM `" & TabN & "`")
'
'  Do While Not rq.EOF
'   runde = runde + 1
'   GesZahl = GesZahl + TIns(TabN, sF, obPS, rq)
'   If BrichAb Then Exit Do
'   rq.Move 1
'  Loop
'  If BrichAb Then
'   frm.Ausgeb "Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "' abgebrochen. " & GesZahl & " Datens‰tze ¸bertragen.", True
'  Else
'   frm.Ausgeb "Fertig mit Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "'. " & GesZahl & " Datens‰tze ¸bertragen.", True
'  End If
'  Call ForeignYes0
'  Call ForeignYes1
''  If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
'  ComTrans
' ElseIf frm.GesBytes <= 1 Then
'  frm.Ausgeb "‹bertragung von " & TabN & " nicht gestartet. Zahl der Datens‰tze mit " & frm.GesBytes & " zu gering in " & frm.snst.QuelleAnamneseBˆgen, True
'' ElseIf LCase$(CStrAcc & frm.snst.QuelleAnamneseBˆgen) = LCase$(DBCnS) Then ' DBCn.ConnectionString
''  frm.Ausgeb "‹bertragung von " & TabN & " nicht gestartet. Quelle und Ziel identisch.", True
' End If
' Exit Function
'fehler:
'Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in holKassen/" + App.path)
' Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' holAllg
#End If

#If False Then
Function holDA(frm As Lese) ' kommt vor in DokumenteAbgehaktkopieren_Click()
 Dim sqldel$, rAf&, GesZahl&
 Const TabN$ = "dokumente abgehakt"
 Const sF$ = "dokpfad"
 Dim rq As New ADODB.Recordset ', f1 AS ADODB.Field, f2 AS ADODB.Field,
 Dim runde&, Pat_id&, i%
 frm.BytesBez = "Datens‰tze:"
 myFrag rq, "SELECT COUNT(0) ct FROM " & "`" & TabN & "`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
 frm.GesBytes = rq!ct
 On Error GoTo fehler
' DBCn.BeginTrans: obTrans = 1
 BegTrans
 Call ForeignNo0
 Call ForeignNo1
 rq.Close
 myFrag rq, "`" & TabN & "`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
 Do While Not rq.EOF
  runde = runde + 1
  Dim Wert$, WertRoh$, obSchonDa%
  obSchonDa = 0
  For i = 1 To 4
   WertRoh = rq(sF)
   If InStrB(WertRoh, "\\\\") > 0 Or InStrB(WertRoh, ":\\") > 0 Then WertRoh = REPLACE$(WertRoh, "\\", "\")
   If LenB(PcDokPfad) = 0 Then getDokPfad
   WertRoh = REPLACE$(LCase$(WertRoh), "$\turbomed\dokumente", PcDokPfad)
   Select Case i
    Case 1: Wert = WertRoh
    Case 2: Wert = REPLACE$(WertRoh, PcDokPfad, "$\TurboMed\Dokumente")
    Case 3: Wert = doUmwfSQL(WertRoh, lies.obMySQL)
    Case 4: Wert = doUmwfSQL(REPLACE$(WertRoh, PcDokPfad, "$\TurboMed\Dokumente"), lies.obMySQL)
   End Select
   If Not rsAnm Is Nothing Then If rsAnm.State = 1 Then rsAnm.Close
   rsAnm.CursorLocation = adUseClient
   myFrag rsAnm, "SELECT -abgehakt j_abgehakt, da.* FROM `" & TabN & "` da WHERE `" & sF & "` = '" & Wert & "'", adOpenStatic, DBCn, adLockOptimistic
   If Not rsAnm.BOF Then
    If rsAnm!j_abgehakt <> 0 Then obSchonDa = -1
    sqldel = "DELETE FROM " & "`" & TabN & "`" & " WHERE " & sF & " = '" & Wert & "'"
    Call myEFrag(sqldel, rAf)
   End If
  Next i
  GesZahl = GesZahl + TIns(TabN, sF, 0, rq, obSchonDa)
  frm.Bytes = runde
  DoEvents
'  ON Error GoTo f1
'  rsAnm.Update
'  ON Error GoTo f0
  If BrichAb Then Exit Do
  rq.Move 1
 Loop
 frm.Ausgeb "Fertig mit der ‹bertragung von " & TabN & " auf MySQL: " & GesZahl & " Datens‰tze ¸bertragen", True
 If BrichAb Then
  frm.Ausgeb "Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "' abgebrochen. " & GesZahl & " Datens‰tze ¸bertragen.", True
 Else
  frm.Ausgeb "Fertig mit Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "'. " & GesZahl & " Datens‰tze ¸bertragen.", True
 End If
 Call ForeignYes0
 Call ForeignYes1
' If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
 ComTrans
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in holDA/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' holAb
#End If

'Function holLab(frm AS Lese) ' kommt nirgends anders vor
' Dim sqldel$, rAF&, gesZahl&
' Const TabN$ = "laborparameter"
' Const sF$ = vNS
' Dim rq As New ADODB.Recordset ', f1 AS ADODB.Field, f2 AS ADODB.Field,
' Dim runde& ', i%
' frm.BytesBez = "Datens‰tze:"
' myFrag rq, "SELECT COUNT(0) AS ct FROM `" & TabN & "`", CStrAcc & frm.snst.QuelleAnamneseBˆgen
' frm.GesBytes = rq!ct
' IF frm.GesBytes > 1 AND LCase$(CStrAcc & frm.snst.QuelleAnamneseBˆgen) <> LCase$(DBCnS) THEN ' DBCn.ConnectionString
'  ON Error GoTo fehler
'  DBCn.BeginTrans
'  Call ForeignNo0
'  Call ForeignNo1
'  rq.Close
'  myFrag rq, "`" & TabN & "`", CStrAcc$ & frm.snst.QuelleAnamneseBˆgen
'  Call myEFrag("DELETE FROM `" & TabN & "`")
'  Do While Not rq.EOF
'   runde = runde + 1
'   gesZahl = gesZahl + TIns(TabN, sF, 0, rq)
'   IF BrichAb THEN Exit Do
'   rq.Move 1
'  Loop
'  IF BrichAb THEN
'   frm.Ausgeb "Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "' abgebrochen. " & gesZahl & " Datens‰tze ¸bertragen.", True
'  Else
'   frm.Ausgeb "Fertig mit Kopieren von '" & TabN & "' von '" & frm.snst.QuelleAnamneseBˆgen & "' nach: '" & frm.Ziel + "'. " & gesZahl & " Datens‰tze ¸bertragen.", True
'  END IF
'  Call ForeignYes0
'  Call ForeignYes1
'  DBCn.CommitTrans
' ElseIf frm.GesBytes <= 1 THEN
'  frm.Ausgeb "‹bertragung von `" & TabN & "` nicht gestartet. Zahl der Datens‰tze mit " & frm.GesBytes & " zu gering in " & frm.snst.QuelleAnamneseBˆgen, True
' ElseIf LCase$(CStrAcc & frm.snst.QuelleAnamneseBˆgen) = LCase$(DBCnS) THEN ' DBCn.ConnectionString
'  frm.Ausgeb "‹bertragung von " & TabN & " nicht gestartet. Quelle und Ziel identisch." & vbCrLf & altAusgabe, True
' END IF
' Exit Function
'fehler:
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in holKassen/" + App.path)
' Case vbAbort: Call MsgBox("Hˆre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function

Public Function doWSt0Erg()
 Dim rs As New ADODB.Recordset
 ' COALESCE(d.f6010,0)=0 AND
 sql = "SELECT pat_id, nachname, vorname, zpkt FROM (SELECT pat_id, nachname, vorname," & _
        " (SELECT MAX(fid) mfid FROM `faelle` WHERE pat_id = `namen`.pat_id) AS mmfid" & _
        " FROM `namen`" & _
        " WHERE EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.icd LIKE 'L89.%' AND NOT d.icd LIKE 'L89.1%' AND d.diagsicherheit <> 'A' AND " & _
                                              " pat_id = `namen`.pat_id and" & _
                                              " fid <> (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id))" & _
" AND NOT EXISTS (SELECT 0 FROM `diagview` d WHERE d.gICD LIKE 'L89.1%'" & _
                                              " pat_id = `namen`.pat_id)" & _
" AND NOT EXISTS (SELECT 0 FROM `diagview` d WHERE d.gICD LIKE 'L89.%'" & _
                                              " pat_id = `namen`.pat_id and" & _
                                              " fid = (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id)))" & _
" AS innen" & _
" LEFT JOIN" & _
" (SELECT MIN(zeitpunkt) AS zpkt, MIN(fid) AS mfid FROM `leistungen` GROUP BY fid) AS lei ON innen.mmfid = lei.mfid ORDER BY zpkt DESC"

' rs.Open "SELECT pat_id, nachname, vorname, zpkt FROM (SELECT pat_id, nachname, vorname, (SELECT MAX(fid) AS mfid FROM `faelle` WHERE pat_id = `namen`.pat_id) AS mmfid FROM `namen` WHERE exists (SELECT * FROM `diagnosen` WHERE icd LIKE 'L89.%' AND NOT icd LIKE 'L89.1%'  AND diagsicherheit NOT IN ('A','Z') AND COALESCE(d.f6010,0)=0 AND pat_id = `namen`.pat_id AND fid <> (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id)) AND NOT EXISTS (SELECT * FROM `diagnosen` WHERE icd LIKE 'L89.1%' AND diagsicherheit NOT IN ('A','Z') AND COALESCE(d.f6010,0)=0 AND pat_id = `namen`.pat_id)) AS innen LEFT JOIN (SELECT MIN(zeitpunkt) AS zpkt, MIN(fid) AS mfid FROM `leistungen` GROUP BY fid) AS lei ON innen.mmfid = lei.mfid ORDER BY zpkt DESC", DBCn, adOpenDynamic, adLockOptimistic
 Screen.MousePointer = vbHourglass
' rs.Open sql, DBCn, adOpenDynamic, adLockOptimistic
 myFrag rs, sql
 Do While Not rs.EOF
  Lese.Ausgeb Right$(Space$(6) & rs!Pat_id, 6) & " " & Format$(rs!zpkt, "dd.mm.yy") & " " & rs!Nachname & " " & rs!Vorname, True
  rs.Move 1
 Loop
 Forms(0).Ausgabe = vbCrLf & "Patienten, die bis zu ihrem vorletzten Fall ein DFS >= Wagner 1 hatten und keines mit Wagner 0 als Dauerdiagnose:" & vbCrLf & altAusgabe
 altAusgabe = Forms(0).Ausgabe
 ' COALESCE(d.f6010,0)=0 AND
 sql = "SELECT pat_id, nachname, vorname, zpkt FROM (SELECT pat_id, nachname, vorname," & _
        " (SELECT MAX(fid) mfid FROM `faelle` WHERE pat_id = `namen`.pat_id) AS mmfid" & _
        " FROM `namen`" & _
        " WHERE   exists (SELECT 0 FROM `diagnosen` d WHERE d.icd LIKE 'L89.%' AND NOT d.icd LIKE 'L89.1%'  AND d.diagsicherheit <> 'A' AND" & _
                                              " pat_id = `namen`.pat_id and" & _
                                              " fid <> (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id))" & _
" AND NOT EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.icd LIKE 'L89.%' AND NOT d.icd LIKE 'L89.1%' AND diagsicherheit = 'Z' AND" & _
                                              " pat_id = `namen`.pat_id and" & _
                                              " fid = (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id))" & _
" AND NOT EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.gICD LIKE 'L89.%' AND" & _
                                              " pat_id = `namen`.pat_id and" & _
                                              " fid = (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id)))" & _
" AS innen" & _
" LEFT JOIN" & _
" (SELECT MIN(zeitpunkt) AS zpkt, MIN(fid) AS mfid FROM `leistungen` GROUP BY fid) AS lei ON innen.mmfid = lei.mfid ORDER BY zpkt DESC"

' rs.Open "SELECT pat_id, nachname, vorname, zpkt FROM (SELECT pat_id, nachname, vorname, (SELECT MAX(fid) AS mfid FROM `faelle` WHERE pat_id = `namen`.pat_id) AS mmfid FROM `namen` WHERE exists (SELECT * FROM `diagnosen` WHERE icd LIKE 'L89.%' AND NOT icd LIKE 'L89.1%'  AND diagsicherheit <> 'A' AND COALESCE(f6010,0)=0 AND pat_id = `namen`.pat_id AND fid <> (SELECT MAX(fid) FROM `faelle` WHERE pat_id = `namen`.pat_id)) AND NOT EXISTS (SELECT * FROM `diagnosen` WHERE icd LIKE 'L89.1%' AND diagsicherheit NOT IN ('A','Z') AND COALESCE(f6010,0)=0 AND pat_id = `namen`.pat_id)) AS innen LEFT JOIN (SELECT MIN(zeitpunkt) AS zpkt, MIN(fid) AS mfid FROM `leistungen` GROUP BY fid) AS lei ON innen.mmfid = lei.mfid ORDER BY zpkt DESC", DBCn, adOpenDynamic, adLockOptimistic
 Screen.MousePointer = vbHourglass
 Set rs = Nothing
' rs.Open sql, DBCn, adOpenDynamic, adLockOptimistic
 myFrag rs, sql
 Do While Not rs.EOF
  Lese.Ausgeb Right$(Space$(6) & rs!Pat_id, 6) & " " & Format$(rs!zpkt, "dd.mm.yy") & " " & rs!Nachname & " " & rs!Vorname, True
  rs.Move 1
 Loop
 Forms(0).Ausgeb "Patienten, die bis zu ihrem vorletzten Fall ein DFS >= Wagner 1 hatten und keines mit Z.n. Wagner >= 1 als Dauerdiagnose:", True
 Screen.MousePointer = vbNormal
End Function ' doWSt0Erg()

' in DokPfadKorrigieren_Click
Function dokpfad‰nder(frm As Lese)
 On Error GoTo fehler
 Dim pos%, runde%, cdTB As ConDtb, db$, fld$, rs As New ADODB.Recordset, Cn As New ADODB.Connection, altCnS$, rAf&
' pu = InputBox("Ursprungsserver ohne Slashes:", "Dokpfad‰nder: Parametereingabe", "mitte")
' IF pu = "" THEN Exit Function
' pz = InputBox("Zielserver ohne Slashes:", "Dokpfad‰nder: Parametereingabe", lcase(frm.dbv.Cpt))
' IF pu = "" THEN Exit Function
 altCnS = DBCnS ' DBCn.ConnectionString
 For cdTB = accDtb To q2Dtb 'accDtb
  Set Cn = acon(quelleT, cdTB)
  For runde = 1 To 3
   Select Case runde
    Case 1: db = "dokumente": fld = "dokpfad"
    Case 2: db = "dokumente abgehakt": fld = "dokpfad"
    Case 3: db = "briefe": fld = "pfad"
   End Select
   If cdTB = accDtb Then
    Set rs = Nothing
'    rs.Open "SELECT * FROM " & "`" & db & "`" & " WHERE `" & fld & "` LIKE '%\turbomed\%' AND NOT (`" & fld & "` LIKE '$\%' AND NOT `" & fld & "` LIKE '$\\%')", Cn, adOpenDynamic, adLockOptimistic '
    myFrag rs, "SELECT * FROM " & "`" & db & "`" & " WHERE `" & fld & "` LIKE '%\turbomed\%' AND NOT (`" & fld & "` LIKE '$\%' AND NOT `" & fld & "` LIKE '$\\%')", adOpenDynamic, Cn
    rAf = 0
    Do While Not rs.EOF
     pos = InStr(1, rs(fld), "\turbomed\", vbTextCompare)
     If pos > 0 Then 'And rs!DokPfad LIKE "$\\*" THEN
      rs(fld) = "$\" & Mid$(REPLACE$(rs(fld), "\\", "\"), pos)
      rs.Update
      rAf = rAf + 1
     End If
     rs.Move 1
    Loop
   Else
    myEFrag "UPDATE `" & db & "` SET `" & fld & "` = CONCAT('$',mid$(replace$(`" & fld & "`, '\\\\', '\\'), instrb(lcase(replace$(`" & fld & "`,'\\\\','\\')), '\\turbomed\\'))) WHERE `" & fld & "` LIKE '%turbomed%'", rAf, Cn
   End If
   frm.Ausgeb rAf & " Datens‰tze in `" & db & "` ge‰ndert auf $\turbomed", True
  Next runde
 Next cdTB
' SetDBCn Nothing
 DBCnOpen altCnS
 DBCnS = altCnS
 MsgBox "Fertig mit ƒndern der Dokumentpfade"
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in dokPfadƒnder/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dokpfad‰nder

#If dp‰alt Then
'Function dokpfad‰nderAlt(frm As Lese)
' On Error GoTo fehler
' Dim pu$, pz$, runde%, cdTB As ConDtb, db$, rs As New ADODB.Recordset, Cn As New ADODB.Connection, altCnS$, rAf&
' pu = InputBox("Ursprungsserver ohne Slashes:", "Dokpfad‰nder: Parametereingabe", "mitte")
' If LenB(pu) = 0 Then Exit Function
' pz = InputBox("Zielserver ohne Slashes:", "Dokpfad‰nder: Parametereingabe", LCase$(frm.dbv.Cpt))
' If LenB(pu) = 0 Then Exit Function
' altCnS = DBCnS ' DBCn.ConnectionString
' For cdTB = accDtb To q2Dtb
'  Set Cn = acon(quelleT, cdTB)
'  For runde = 1 To 2
'   If runde = 1 Then db = "dokumente" Else db = "dokumente abgehakt"
'   If cdTB = accDtb Then
'    Set rs = Nothing
''    rs.Open "SELECT * FROM " & "`" & db & "`" & " WHERE dokpfad LIKE '%\\" & pu & "%'", Cn, adOpenDynamic, adLockOptimistic
'    myFrag rs, "SELECT * FROM " & "`" & db & "`" & " WHERE dokpfad LIKE '%\\" & pu & "%'", adOpenDynamic, Cn
'    rAf = 0
'    Do While Not rs.EOF
'     rs!DokPfad = REPLACE$(rs!DokPfad, "\\" & pu, "\\" & pz)
'     rs.Update
'     rAf = rAf + 1
'    Loop
'   Else
'    myEFrag "UPDATE " & "`" & db & "`" & " SET dokpfad = replace$(dokpfad,'\\" & pu & "','\\" & pz & "') WHERE dokpfad LIKE '%\\" & pu & "%'", rAf, Cn
'   End If
'   frm.Ausgeb rAf & " Datens‰tze in " & "`" & db & "`" & " ge‰ndert von \\" & pu & " nach \\ " & pz & " auf " & frm.dbv.Cpt, True
'  Next runde
' Next cdTB
'' SetDBCn Nothing
' DBCnOpen altCnS
' DBCnS = altCnS
' MsgBox "Fertig mit ƒndern der Dokumentpfade von " & pu & " nach " & pz & "!"
' Exit Function
'fehler:
'Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in dokPfadƒnder/" + App.path)
' Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' dokpfad‰nder
#End If

'Function DokPfadKorr(frm AS Lese) ' DokumentPfade korrigieren
' Dim rs As New ADODB.Recordset, sql$, rAF&, KZahl&
' IF lenb(PcDokPfad) = 0 THEN getDokPfad
' IF Not lies.obMySQL THEN
'  rs.Open "dokumente", DBCn, adOpenDynamic, adLockOptimistic
'  Do While Not rs.EOF
'   IF instrb(rs!DokPfad, "$\\TurboMed\\Dokumente") > 0 THEN
'    rs!DokPfad = replace$(rs!DokPfad, "\\", "\")
'    KZahl = KZahl + 1
'    rs.Update
'   END IF
'   IF instrb(rs!DokPfad, "$\TurboMed\Dokumente") > 0 THEN
'    rs!DokPfad = replace$(lcase(rs!DokPfad), "$\turbomed\dokumente", PcDokPfad)
'    KZahl = KZahl + 1
'    rs.Update
'   END IF
'   IF instrb(rs!DokPfad, ":\\") > 0 OR instrb(rs!DokPfad, "\\\\") > 0 THEN
'    rs!DokPfad = replace$(rs!DokPfad, "\\", "\")
'    KZahl = KZahl + 1
'    rs.Update
'   END IF
'   rs.Move 1
'  Loop
'  MsgBox KZahl & " Pfade in " & rs.ActiveConnection.ConnectionString & " / " & rs.Source & " korrigiert!"
''  sql = "UPDATE dokumente SET dokpfad = replace$(dokpfad,""\\"",""\"") WHERE instr(dokpfad,""$\\TurboMed\\Dokumente"")>0"
''  SET rs = myEFrag(sql, raf)
' Else
'  MsgBox "Nichts korrigiert, da MySQL gew‰hlt!"
' END IF
'End Function

Public Function anaUpd(rq As Recordset, rz As Recordset)
 Const TName$ = "anamnesebogen"
 Dim ErrNr&, ErrDes$
 Dim csql As New CString
 Dim i&, sql$, rAf&, Wert$
 Dim pos0&, neuverbindzahl%
 Dim fehlerzahl%
' Dim slen As Adodb.Recordset, maxlen&(), ordi%()
nochmal:
   On Error GoTo fehler
'   Dim rSc As New ADODB.Recordset
'   SET rSc = DBCnOSchema(adSchemaColumns, Array(DBCn.DefaultDatabase, Empty, "anamnesebogen", Empty))
'   ReDim maxlen(rSc.RecordCount)
'   ReDim ordi(rSc.RecordCount)
'   i = 1
'   Do While Not rSc.EOF
'    IF NOT ISNULL(rSc!character_maximum_length) THEN
'     IF rSc!character_maximum_length <= &HFFFFFFFF THEN
'      maxlen(i) = rSc!character_maximum_length
'     END IF
'    END IF
'    ordi(i) = rSc!ordinal_position
'    i = i + 1
'    rSc.Move 1
'   Loop

'Datentypen Recordset
'Const adEmpty = 0
'Const adSMALLINT = 2
'Const adInteger = 3
'Const adSingle = 4
'Const adDouble = 5
'Const adCurrency = 6
'Const adDate = 7
'Const adBSTR = 8
'Const adIDispatch = 9
'Const adError = 10
'Const adBoolean = 11
'Const adVariant = 12
'Const adIUnknown = 13
'Const adDecimal = 14
'Const adTinyInt = 16
'Const adUNSIGNEDTinyInt = 17
'Const adUNSIGNEDSMALLINT = 18
'Const adUNSIGNEDInt = 19
'Const adBigInt = 20
'Const adUNSIGNEDBigInt = 21
'Const adGUID = 72
'Const adBinary = 128
'Const adChar = 129
'Const adWChar = 130
'Const adNumeric = 131
'Const adUserDefined = 132
'Const adDBDate = 133
'Const adDBTime = 134
'Const adDBTimeStamp = 135
'Const adVarChar = 200
'Const adLongVarChar = 201
'Const adVarWChar = 202
'Const adLongVarWChar = 203
'Const adVarBinary = 204
'Const adLongVarBinary = 205

   csql = "UPDATE `" & TName & "` SET "
   For i = 0 To rq.Fields.COUNT - 1
    Select Case LCase$(rq.Fields(i).name)
     Case "pat_id", "prim"
     Case Else
'      IF (i Mod 15) = 0 THEN sql = sql & " _" & vbCrLf
'     sql = sql & "`" & rq.Fields(i).name & "` = "
      csql.AppVar Array("`", rq.Fields(i).name, "` = ")
      Select Case rq.Fields(i).Type
         Case adBoolean, adUnsignedTinyInt ' 11: in MySQL bit
'         Case 11, 17
           If IsNull(rq.Fields(i).Value) Then Wert = "0" Else If rq.Fields(i) Then If lies.obMySQL Then Wert = "1" Else Wert = "-1" Else Wert = "0"
         Case adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adBigInt, adUnsignedBigInt, _
             adNumeric, adVarNumeric, adCurrency, adDecimal
'         Case 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", ".")
         Case adUnsignedSmallInt, adUnsignedInt
'         Case 18, 19
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = Abs(REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", "."))
         Case adDate, adFileTime, adDBDate, adDBTime, adDBTimeStamp
'         Case 7, 64, 133, 134, 135
           Wert = DatFor_k(rq.Fields(i).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in anaUpd mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
         Case adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adEmpty, _
              adIDispatch, adVariant, adIUnknown, adGUID, adBinary, adUserDefined, adPropVariant, _
              adVarBinary, adLongVarBinary, adError, adArray
'         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205, 10, 8192
           Wert = "'" & doUmwfSQL(IIf(IsNull(rq.Fields(i).Value), vNS, rq.Fields(i).Value), lies.obMySQL) & "'"
'           IF 1 = 0 THEN
           pos0 = InStr(Wert, vbNullChar)
           If pos0 > 0 Then
            Wert = Left$(Wert, pos0 - 2) & "'"
           End If
           If rz.Fields(rq.Fields(i).name).Type = adVarChar Then ' 200
            If InStrB(UCase$(DBCn), "MYSQL") <> 0 Then
'             SET rSc = Nothing
'             SET rSc = DBCnOSchema(adSchemaColumns, Array(DBCn.DefaultDatabase, Empty, "anamnesebogen", rq.Fields(i).Name))
'             Feldl‰nge = rSc!character_maximum_length
'             SET slen = myEFrag("SHOW FULL COLUMNS FROM `anamnesebogen` WHERE field = '" & rq.Fields(i).Name & "'")
'             dim slV$
'             slV = slen.Fields(1).Value
'             IF slV LIKE "varchar(*" THEN
'              Feldl‰nge = mid$(slV, 9, InStr(9, slV, ")") - 9)
              Dim Feldl‰nge&
              Feldl‰nge = rz.Fields(rq.Fields(i).name).DefinedSize
              If Len(Wert) - 2 > Feldl‰nge Then ' Anf¸hrungszeichen
               Dim rsc As New ADODB.Recordset
               Set rsc = Nothing
               myFrag rsc, "SHOW FULL COLUMNS FROM `" & TName & "` WHERE field = '" & rq.Fields(i).name & "'"
               myEFrag ("ALTER TABLE `" & TName & "` " & sqlALTER & " COLUMN `" & rq.Fields(i).name & "` " & sqlText & "(" & Len(Wert) - 2 & ")" & " DEFAULT " & IIf(IsNull(rsc!Default), "NULL", rsc!Default) & IIf(IsNull(rsc!collation), vNS, " COLLATE " & rsc!collation) & " COMMENT '" & rsc!Comment & "'")
               Const tmp$ = "tmpMess.txt"
               Open tmp For Output As #322
               Print #322, "Vergrˆﬂere in Tabelle " & TName & " das Feld '" & rq.Fields(i).name & "' von " & Feldl‰nge & " auf " & Len(Wert) - 2 & " wegen:"
               Print #322, "'" & Wert & "'"
               Close #322
               zeigan tmp
'              END IF
              End If
             End If
'            END IF ' 1 = 0
'            Feldl‰nge = rz.Fields(rq.Fields(i).Name).Properties.Item("OCTETLENGTH")
           End If
         Case Else
           Err.Raise 999, , "Fehler in anaUpd mit unbekanntem Datentyp: " & rq.Fields(i).Type
       End Select
'       sql = sql & Wert & ","
       csql.AppVar Array(Wert, ",")
       Wert = vNS
    End Select
    If Not lies.obMySQL And i Mod 1000 = 999 Then
      'sql = LEFT(sql, len(sql) - 1) & " WHERE pat_id = " & rq!Pat_id
      csql.Cut (csql.length() - 1)
      csql.AppVar Array(" WHERE pat_id = ", rq!Pat_id)
      Call myEFrag(csql.Value, rAf)
      csql = "UPDATE `" & TName & "` SET "
    End If
   Next i
'   sql = LEFT(sql, Len(sql) - 1) & " WHERE pat_id = " & rq!Pat_id
    csql.Cut (csql.length() - 1)
    csql.AppVar Array(" WHERE pat_id = ", rq!Pat_id)
'   GoTo nochmal:
   Set rq = Nothing
   Dim rtok As New ADODB.Recordset
   myFrag rtok, "SHOW FULL PROCESSLIST"
   Do While Not rtok.EOF
    If rtok!Command = "Sleep" Then
     myEFrag ("KILL " & rtok!id)
    End If
    On Error GoTo schlaffehler
    rtok.MoveNext
   Loop
endeschlaf:
   On Error GoTo fehler
   Call myEFrag(csql.Value, rAf, , True, ErrNr)
   If rAf <> 1 And ErrNr <> 0 Then
'    Call myEFrag("COMMIT")
    ComTrans
    Call myEFrag(csql.Value, rAf, , True, ErrNr)
    If rAf <> 1 And ErrNr <> 0 Then
     Set rq = Nothing
     Call myEFrag(csql.Value, rAf, , True, ErrNr, ErrDes)
     If rAf <> 1 And ErrNr <> 0 Then
      Err.Raise 999, , ErrNr & " " & ErrDes & vbCrLf & "Fehler in anaUpd: Falsche Zahl an Datens‰tzen aktualisiert: " & rAf & vbCrLf & "bei: " & csql.Value
     End If ' rAF <> 1 And ErrNr <> 0 Then
    End If ' rAF <> 1 And ErrNr <> 0 Then
   End If ' rAF <> 1 And ErrNr <> 0 Then
Ende:
   Exit Function
schlaffehler:
  Resume endeschlaf
fehler:
fehlerzahl = fehlerzahl + 1
If fehlerzahl > 2 Then
 Resume Ende
End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Resume
End If
If Err.Number = -2147467259 Then ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.36-log]MySQL server has gone away
  Call acon(quelleT)
  neuverbindzahl = neuverbindzahl + 1
  If neuverbindzahl > 10 Then End
  Resume
End If
If Err.Number = -2147217871 Or Err.Number = -2147467259 Then
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in anaUpd/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' anaUpd(rq AS Recordset)

Public Function TUpd(TabN$, sF$, rq As Recordset)
 Dim i&, sql$, rAf&, Wert$
 Dim pos0&
 Dim rz As New ADODB.Recordset
 If InStrB(UCase$(DBCn), "MYSQL") <> 0 Then
  myFrag rz, "SELECT * FROM `" & TabN & "` LIMIT 1"
 Else
  myFrag rz, "SELECT top 1 * FROM `" & TabN
 End If
   On Error GoTo fehler
   sql = "UPDATE `" & TabN & "` SET "
   For i = 0 To rq.Fields.COUNT - 1
    Select Case LCase$(rq.Fields(i).name)
     Case sF, "prim"
     Case Else
      sql = sql & "`" & rq.Fields(i).name & "`" & " = "
      Select Case rq.Fields(i).Type
         Case adBoolean, adUnsignedTinyInt ' 11
           If IsNull(rq.Fields(i).Value) Then Wert = "0" Else If rq.Fields(i) Then If lies.obMySQL Then Wert = "1" Else Wert = "-1" Else Wert = "0"
         Case 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", ".")
         Case 17, 18, 19
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = Abs(REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", "."))
         Case 7, 64, 133, 134, 135
           Wert = DatFor_k(rq.Fields(i).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in TUpd mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
           Wert = "'" & doUmwfSQL(IIf(IsNull(rq.Fields(i).Value), vNS, rq.Fields(i).Value), lies.obMySQL) & "'"
'           IF instrb(sF, "pfad") > 0 THEN
'            IF lenb(PcDokPfad) = 0 THEN getDokPfad
'            Wert = replace$(lcase(Wert), "$\\turbomed\\dokumente", PcDokPfad)
'           END IF
           pos0 = InStr(Wert, vbNullChar)
           If pos0 > 0 Then
            Wert = Left$(Wert, pos0 - 2) & "'"
           End If
           If rz.Fields(rq.Fields(i).name).Type = adVarChar Then ' 200
            If InStrB(UCase$(DBCn), "MYSQL") <> 0 Then
'             SET rSc = Nothing
'             SET rSc = DBCnOSchema(adSchemaColumns, Array(DBCn.DefaultDatabase, Empty, "anamnesebogen", rq.Fields(i).Name))
'             Feldl‰nge = rSc!character_maximum_length
'             SET slen = myEFrag("SHOW FULL COLUMNS FROM `anamnesebogen` WHERE field = '" & rq.Fields(i).Name & "'")
'             dim slV$
'             slV = slen.Fields(1).Value
'             IF slV LIKE "varchar(*" THEN
'              Feldl‰nge = mid$(slV, 9, InStr(9, slV, ")") - 9)
              Dim rsc As New ADODB.Recordset
              Dim Feldl‰nge&
'              Feldl‰nge = CLng(Mid$(rsc!Type, 9, Len(rsc!Type) - 9))
              Feldl‰nge = rz.Fields(rq.Fields(i).name).DefinedSize
              If Len(Wert) - 2 > Feldl‰nge Then ' Anf¸hrungszeichen
               Set rsc = Nothing
               myFrag rsc, "SHOW FULL COLUMNS FROM `" & TabN & "` WHERE field = '" & rq.Fields(i).name & "'"
               myEFrag ("ALTER TABLE `" & TabN & "` " & sqlALTER & " COLUMN `" & rq.Fields(i).name & "` " & sqlText & "(" & Len(Wert) - 2 & ")" & " DEFAULT " & IIf(IsNull(rsc!Default), " NULL", rsc!Default) & IIf(IsNull(rsc!collation), vNS, " COLLATE " & rsc!collation) & " COMMENT '" & rsc!Comment & "'")
               Const tmp$ = "tmpMess.txt"
               Open tmp For Output As #322
               Print #322, "Vergrˆﬂere in Tabelle `" & TabN & "` das Feld '" & rq.Fields(i).name & "' von " & Feldl‰nge & " auf " & Len(Wert) - 2 & " wegen:"
               Print #322, "'" & Wert & "'"
               Close #322
               zeigan tmp
'              END IF
              End If
             End If
'            END IF ' 1 = 0
'            Feldl‰nge = rz.Fields(rq.Fields(i).Name).Properties.Item("OCTETLENGTH")
           End If
         Case Else
           Err.Raise 999, , "Fehler in TUpd mit unbekanntem Datentyp: " & rq.Fields(i).Type
       End Select
       sql = sql & Wert & ","
       Wert = vNS
    End Select
    If Not lies.obMySQL And i Mod 100 = 99 Then
      sql = Left$(sql, Len(sql) - 1) & " WHERE pat_id = " & rq!Pat_id
      Call myEFrag(sql, rAf)
      sql = "UPDATE " & TabN & " SET "
    End If
   Next i
   sql = Left$(sql, Len(sql) - 1) & " WHERE pat_id = " & rq!Pat_id
   Call myEFrag(sql, rAf)
   If rAf <> 1 Then Err.Raise 999, , "Fehler in TUpd: Falsche Zahl an Datens‰tzen aktualisiert: " & rAf
   Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in TUpd/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TUpd(rq AS Recordset)

Public Function anaIns(rq As Recordset)
 Const TName$ = "anamnesebogen"
 Dim i&, sql1$, sql2$, rAf&, Wert$, mprim&, rs As ADODB.Recordset, pos0&
   On Error GoTo fehler
   sql1 = "INSERT INTO `" & TName & "` ("
   sql2 = "values("
   Set rs = myEFrag("SELECT MAX(prim) mprim FROM `" & TName & "`")
   mprim = IIf(IsNull(rs!mprim), 0, rs!mprim) + 1
   For i = 0 To rq.Fields.COUNT - 1
    Select Case LCase$(rq.Fields(i).name)
     Case "pat_id"
      Wert = rq!Pat_id
     Case "prim"
      Wert = mprim
     Case Else
      Select Case rq.Fields(i).Type
         Case adBoolean, adUnsignedTinyInt ' 11
           If IsNull(rq.Fields(i).Value) Then Wert = "0" Else If rq.Fields(i) Then If lies.obMySQL Then Wert = "1" Else Wert = "-1" Else Wert = "0"
         Case 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", ".")
         Case 17, 18, 19
           If IsNull(rq.Fields(i)) Then Wert = 0 Else Wert = Abs(REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", "."))
         Case 7, 64, 133, 134, 135
           Wert = DatFor_k(rq.Fields(i).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in anaIns mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
           Wert = "'" & doUmwfSQL(IIf(IsNull(rq.Fields(i).Value), vNS, rq.Fields(i).Value), lies.obMySQL) & "'"
           pos0 = InStr(Wert, vbNullChar)
           If pos0 > 0 Then
            Wert = Left$(Wert, pos0 - 2) & "'"
           End If
         Case Else
           Err.Raise 999, , "Fehler in anaIns: Falsche Zahl an Datens‰tzen aktualisiert: " & rAf
       End Select
    End Select
    sql1 = sql1 & "`" & rq.Fields(i).name & "`" & ","
    sql2 = sql2 & Wert & ","
    Wert = vNS
   Next i
   sql2 = Left$(sql2, Len(sql2) - 1) & ")"
   sql1 = Left$(sql1, Len(sql1) - 1) & ") " & sql2
   InsKorr DBCn, DBCnS, sql1, rAf
   If rAf <> 1 Then Err.Raise 999, , "Fehler in anaIns: Falsche Zahl an Datens‰tzen aktualisiert: " & rAf
   Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in anaIns/" + App.path)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' anaIns(rq AS Recordset)

#Const turbomed = -1
Public Function TIns&(TabN$, sF$, obPS%, rq, Optional schonAbgehakt) ' ob Prim‰rschl¸ssel
 Dim i&, sql1$, sql2$, sqldel$, Wert$, mprim&
 Dim rz As New ADODB.Recordset, rs As ADODB.Recordset
   On Error GoTo fehler
   sql1 = "INSERT INTO `" & TabN & "` ("
   sql2 = "values("
   If obPS Then
    Set rs = myEFrag("SELECT MAX(" & sF & ") mprim FROM " & "`" & TabN & "`" & ";")
    mprim = IIf(IsNull(rs!mprim), 0, rs!mprim) + 1
   End If
   For i = 0 To rq.Fields.COUNT - 1
    Select Case LCase$(rq.Fields(i).name)
     Case sF
      If obPS Then
       Wert = mprim
      Else
       Wert = rq(sF)
       Wert = IIf(IsNull(rq(sF)), vNS, rq(sF))
#If turbomed Then
       If InStrB(sF, "pfad") > 0 Then
        Dim ob2%, obnetz%
        If Wert Like "??*\\*" Then
         ob2 = True
         If Wert Like "\\\\*" Then obnetz = True
        Else
         If Wert Like "\\*" Then obnetz = True
        End If
'        IF lenb(PcDokPfad) = 0 THEN getDokPfad
'        Wert = umwfSQL(replace$(lcase(replace$(Wert, "\\", "\")), "$\turbomed\dokumente", PcDokPfad))
        Wert = doUmwfSQL(REPLACE$(Wert, "\\", "\"), lies.obMySQL)
        If obnetz And Not ob2 Then Wert = "\\" & Wert
       End If
#End If
       Wert = "'" & Wert & "'"
      End If
     Case Else
      Select Case rq.Fields(i).Type
         Case adBoolean, adUnsignedTinyInt ' 11
           If IsNull(rq.Fields(i).Value) Then
            Wert = "0"
           Else
            If rq.Fields(i) Then
             Wert = "-1"
             If lies.obMySQL Then
              Set rz = Nothing
              myFrag rz, "SHOW FULL COLUMNS FROM `" & TabN & "` WHERE field = '" & rq.Fields(i).name & "'"
              If Not rz.BOF Then
               If rz!Type = "bit(1)" Then Wert = "1"
              End If
             End If
            Else
             Wert = "0"
            End If
           End If
           If Not IsError(schonAbgehakt) Then
            If LCase$(rq.Fields(i).name) = "abgehakt" And schonAbgehakt Then
             If lies.obMySQL Then Wert = "1" Else Wert = "-1"
            End If
           End If
         Case 16, 2, 3, 4, 5, 20, 21, 131, 139, 6, 14
           If IsNull(rq.Fields(i)) Then
            Wert = "null"
           Else
            Wert = REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", ".")
            If LenB(Wert) = 0 Then Wert = 0
           End If
         Case 17, 18, 19
           If IsNull(rq.Fields(i)) Then
            Wert = "null"
           Else
            Wert = Abs(REPLACE$(REPLACE$(rq.Fields(i), ".", ""), ",", "."))
            If LenB(Wert) = 0 Then Wert = 0
           End If
         Case 7, 64, 133, 134, 135
           Wert = DatFor_k(rq.Fields(i).Value)
           If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in TIns mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
         Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
           Wert = doUmwfSQL(IIf(IsNull(rq.Fields(i).Value), vNS, rq.Fields(i).Value), lies.obMySQL)
           If InStrB(sF, "pfad") > 0 Then
'            IF lenb(PcDokPfad) = 0 THEN getDokPfad
'            Wert = replace$(lcase(replace$(Wert, "\\", "\")), "$\turbomed\dokumente", PcDokPfad)
            Wert = REPLACE$(Wert, "\\", "\")
           End If
           Wert = "'" & Wert & "'"
         Case Else
           Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datens‰tzen aktualisiert: " & TIns
       End Select
    End Select
    sql1 = sql1 & "`" & rq.Fields(i).name & "`" & ","
    sql2 = sql2 & Wert & ","
    Wert = vNS
   Next i
   sql2 = Left$(sql2, Len(sql2) - 1) & ")"
   sql1 = Left$(sql1, Len(sql1) - 1) & ") " & sql2
   Dim rAf&
'   DBCn.CursorLocation = adUseClient
   InsKorr DBCn, DBCnS, sql1, rAf
   If rAf = 0 Or (rAf <> 1 And rAf < 100) Then Err.Raise 999, , "Fehler in TIns: Falsche Zahl an Datens‰tzen aktualisiert: " & rAf
   If rAf <> 1 Then TIns = 1
   Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in TIns/" & AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TIns(rq AS Recordset)

Private Function AusschlVorbel()
  Ausschl(0) = " "
  Ausschl(1) = ","
  Ausschl(2) = "."
  Ausschl(3) = ";"
  Ausschl(4) = 9
  Ausschl(5) = 10
  Ausschl(6) = 11
  Ausschl(7) = 13
End Function ' AusschlVorbel

'Liefert Anfang- und endposition des ersten (bzw. bei obumgekehrt des letzten) Vorkommens von s2 in s1, wobei die in ausschl hinterlegten Zeichen ignoriert werden
Public Function doSp(s1$, s2$, Optional obUmgekehrt = -1) As Vorkommen
  Dim iPos%
  Dim f1() As String * 1, f2() As String * 1
  Dim l1&, l2&
  Dim i%, j%, k%
  Dim paﬂt%, valid%, zwi%
  If Ausschl(0) = vbNullChar Then
   Call AusschlVorbel
  End If
  l1 = Len(s1)
  l2 = Len(s2)
  ReDim f1(l1)
  ReDim f2(l2)
  For i = 0 To l1 - 1
   If obUmgekehrt Then
    f1(l1 - 1 - i) = Mid$(s1, i + 1, 1)
   Else
    f1(i) = Mid$(s1, i + 1, 1)
   End If
  Next i
  For i = 0 To l2 - 1
   If obUmgekehrt Then
    f2(l2 - 1 - i) = Mid$(s2, i + 1, 1)
   Else
    f2(i) = Mid$(s2, i + 1, 1)
   End If
  Next i
  j = 0
  i = 0
  doSp.Beg = 1
  Do
   If f2(j) = f1(i) Then
    j = j + 1
    If j = l2 Then doSp.END = i + 1: paﬂt = -1: Exit Do
    Do
     i = i + 1
     If i = l1 Then Exit Do
     valid = -1
     For k = 0 To UBound(Ausschl)
      If f1(i) = Ausschl(k) And f1(i) <> f2(j) Then valid = 0: Exit For
     Next k
     If valid Then Exit Do
    Loop
   Else ' f2(j) = f1(i) THEN
    If j = 0 Then i = i + 1
    j = 0
    doSp.Beg = i + 1
   End If ' f2(j) = f1(i) THEN
   If i = l1 Then Exit Do
  Loop
  If Not paﬂt Then
   doSp.Beg = 0
  Else
   If obUmgekehrt Then
    zwi = doSp.Beg
    doSp.Beg = l1 - doSp.END + 1
    doSp.END = l1 - zwi + 1
   Else
   End If
  End If
End Function ' doSp

' aufgerufen in: Form_Activate
Function getLDatei$(Upfad$, Optional muster$)
 Dim Datei$, Pfad$ ' , nr&
 Dim objShell, objExecObject
 On Error GoTo fehler
 Pfad = getPfad(Upfad, Datei)
 If LenB(muster) = 0 Then muster = Datei
 If LenB(muster) = 0 Then muster = "*.bdt"
nochmal:
 Pfad = getPfad(Upfad)
' IF Muster = "*.bdt" THEN
If LenB(Pfad) <> 0 Then
'  erg = Shell("cmd / c dir '" & Pfad & muster & "' /b /o-d")
'  SET objShell = Wscript.CreateObject("WScript.Shell")
  Set objShell = New IWshShell_Class
#If True Then
   objShell.rUn "cmd /c dir """ & Pfad & muster & """ /b /o-d|clip", 0, True
   Dim strOutput$, pos&
'   strOutput = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
   strOutput = Dir(Pfad, vbNormal)
   Dim tw As Object
   strOutput = Dir(Pfad & muster, vbNormal)
   Set tw = CreateObject("htmlfile")
   If Not tw Is Nothing Then
    On Error Resume Next
    strOutput = tw.ParentWindow.ClipboardData.GetData("text")
    On Error GoTo fehler
    If strOutput = "" Then GoTo anders
    pos = InStr(strOutput, Chr(13))
    If pos Then Datei = Left$(strOutput, pos - 1) Else Datei = strOutput
    If Datei <> "" Then getLDatei = Pfad & Datei
   Else
anders:
    Dim DCol As New Collection
    Datei = Dir(Pfad & muster)
    Do While LenB(Datei) <> 0
     Call AddFileSortedByDateLastModified(DCol, FSO.GetFile(Pfad & Datei), kmName)
     Datei = Dir
    Loop
    If DCol.COUNT > 0 Then getLDatei = Pfad & DCol.Item(DCol.COUNT).name
   End If ' Not tw Is Nothing Then else
#ElseIf False Then
   Set objExecObject = objShell.Exec("cmd /c dir """ & Pfad & muster & """ /b /o-d")
   Do While Not objExecObject.StdOut.AtEndOfStream
    Datei = objExecObject.StdOut.ReadLine()
    getLDatei = Pfad & Datei
    Exit Do
   Loop
#Else
   Dim Folder As Folder
   Dim cur As File, DT As File
   Set Folder = FSO.GetFolder(Pfad)
   Set cur = FSO.GetFile(Pfad & Dir(Pfad)) 'cur (Neueste gefundene) auf erste setzen
   For Each DT In Folder.Files
    If UCase$(DT) Like UCase$(muster) Then
     If DT.DateCreated > cur.DateCreated Then
      Set cur = FSO.GetFile(Pfad & DT.name)
     End If
    End If
   Next DT
   If Not cur Is Nothing Then getLDatei = Pfad & cur.name
#End If
 End If ' LenB(pfad) <> 0 Then
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in getLDatei/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume nochmal
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' getLDatei

' aufgerufen nur in getLDatei
Public Function getPfad$(dname$, Optional getDat$)
 Dim i%, ldn%
 getDat = vNS
 ldn = Len(dname)
 For i = ldn To 1 Step -1
  If Mid$(dname, i, 1) = "\" Then
   getPfad = Left$(dname, i)
   If i = ldn Then getDat = vNS Else getDat = Right$(dname, ldn - i) ' 13.2.22: -1 gestrichen
   Exit For
  End If
 Next i
End Function ' getPfad

' #If False Then
' wurde nur verwendet in getLDatei
Public Function AddFileSortedByDateLastModified _
 (Collection As Collection, Item As File, _
 Optional ByVal KeyMode As KeyModeConstants = kmName) As Long
  Dim nCount&
  Dim nHigh As Long
  Dim nLow As Long
  Dim nTest As Long
  Dim nTestItem As Object
  Dim nKey As String
  Select Case KeyMode
    Case kmName
      nKey = Item.name
    Case kmPath
      nKey = Item.path
  End Select
  With Collection
    nCount = .COUNT
    If nCount Then
      If Item.DateLastModified < .Item(1).DateLastModified Then
        .Add Item, nKey, 1
        AddFileSortedByDateLastModified = 1
      ElseIf Item.DateLastModified > _
       .Item(nCount).DateLastModified Then
        .Add Item, nKey
        AddFileSortedByDateLastModified = nCount + 1
      Else
        nLow = 1
        nHigh = nCount
        Do
          nTest = (nLow + nHigh) \ 2
          If nTest = nLow Then
            Exit Do
          End If
          Set nTestItem = .Item(nTest)
          Select Case Item.DateLastModified
            Case Is < nTestItem.DateLastModified
              nHigh = nTest
            Case nTestItem.DateLastModified
              Exit Do
            Case Is > nTestItem.DateLastModified
              nLow = nTest
          End Select
        Loop
        If nTest < nCount Then
          Do While .Item(nTest + 1).DateLastModified = _
           Item.DateLastModified
            nTest = nTest + 1
            If nTest = nCount Then
              Exit Do
            End If
          Loop
        End If
'        ON Error Resume Next ' wenn Dateiname einmal mit kleinem und einmal mit groﬂem Buchstaben da 24.8.20
        .Add Item, nKey, , nTest
        AddFileSortedByDateLastModified = nTest + 1
'        ON Error GoTo 0
      End If
    Else
      .Add Item, nKey
      AddFileSortedByDateLastModified = 1
    End If
  End With
End Function ' AddFileSortedByDateLastModified
' #End If

' Analoga nach 29.9.06 verschrieben:
' SELECT DISTINCT rezepteintraege.pat_id, nachname, vorname, versicherung, kassenliste.name FROM (`quelle`.`rezepteintraege` LEFT JOIN quelle.`anamnesebogen` ON rezepteintraege.pat_id = `anamnesebogen`.pat_id) LEFT JOIN quelle.`kassenliste` ON `anamnesebogen`.versicherung = `kassenliste`.vknr WHERE (medikament LIKE "%umalog%" OR medikament LIKE "%ovorapid%" OR medikament LIKE "%iprolog%" OR medikament LIKE "%pidra%") AND zeitpunkt > '2006-09-28' AND Diabetestyp = 2 AND versicherungsart <> 90;
' SELECT DISTINCT rezepteintraege.pat_id, nachname, vorname, `kassenliste`.name, medikament FROM (`quelle`.`rezepteintraege` LEFT JOIN quelle.`anamnesebogen` ON rezepteintraege.pat_id = `anamnesebogen`.pat_id) LEFT JOIN quelle.`kassenliste` ON `anamnesebogen`.versicherung = `kassenliste`.vknr WHERE (medikament LIKE "%umalog%" OR medikament LIKE "%ovorapid%" OR medikament LIKE "%iprolog%" OR medikament LIKE "%pidra%") AND zeitpunkt > '2006-09-28' AND Diabetestyp = 2 AND versicherungsart <> 90 AND versicherung <> 71101 ORDER BY versicherung;
'SELECT DISTINCT rezepteintraege.pat_id, nachname, vorname, `kassenliste`.name, medikament FROM (`quelle`.`rezepteintraege` LEFT JOIN quelle.`anamnesebogen` ON rezepteintraege.pat_id = `anamnesebogen`.pat_id) LEFT JOIN quelle.`kassenliste` ON `anamnesebogen`.versicherung = `kassenliste`.vknr _
   WHERE (medikament LIKE "%umalog%" OR medikament LIKE "%ovorapid%" OR medikament LIKE "%iprolog%" OR medikament LIKE "%pidra%") AND zeitpunkt > '2006-09-28' AND Diabetestyp = 2 AND versicherungsart <> 90 AND NOT name LIKE "aok%" AND NOT (name LIKE "%Barmer%" AND medikament LIKE "%Humalog%") AND NOT (name LIKE "LKK%" AND (nachname LIKE "Kreitmair%" OR nachname = "Zeiner")) AND NOT (name LIKE "Techniker Krankenkasse%" AND medikament NOT LIKE "%Liprolog%") AND NOT (name LIKE "neue BKK%" AND (medikament LIKE "Novorapid%" OR medikament LIKE "Humalog%")) AND NOT (name LIKE "Deutsche BKK%" AND medikament LIKE "Humalog%") AND NOT name LIKE "dak%" AND NOT (name LIKE "GEK%" AND nachname = "schˆnwetter") AND NOT (name LIKE "bkk man%"
' AND medikament LIKE "novorapid%") ORDER BY versicherung;
Function ‹bertragFormulare(lies As Lese, AccForm$, FName$, Datenbank$, MyTab$)
' Const DatenBank$ = uverz & "zugriff.mdb"
' Const AccForm$ = "Anamnesebogen"
' Const MyTab$ = "anamnesebogen"
 Dim CS$
 Const obDatenS‰tze% = -1
 Dim MdB As New Access.Application ' createobject("Access.Application.9")
 On Error GoTo fehler
 lies.Ausgeb "÷ffne Datenbank " & Datenbank, False
 Call MdB.OpenCurrentDatabase(Datenbank)
 Dim i&, j&, k&, m&
#If False Then
 Call MdB.DoCmd.OpenForm(assform, acDesign)
 For i = 0 To MdB.Forms(0).Properties.COUNT - 1
  On Error Resume Next
'  Debug.Print MdB.Forms(0).Properties(i).name & " " & MdB.Forms(0).Properties(i)
  On Error GoTo fehler
 Next i
 Call MdB.DoCmd.Close(acForm, AccForm)
 Call MdB.DoCmd.OpenForm(AccForm)
 For i = 0 To MdB.Forms(0).Properties.COUNT - 1
  On Error Resume Next
'  Debug.Print MdB.Forms(0).Properties(i).name & " " & MdB.Forms(0).Properties(i)
  On Error GoTo fehler
 Next i
 Call MdB.DoCmd.Close(acForm, AccForm)
#End If
 lies.Ausgeb "÷ffne Formular " & AccForm, False
 Call MdB.DoCmd.OpenForm(AccForm, acDesign)
 Dim MaxHˆhe&(2), aktHˆhe&, Sekt%, GesHˆhe& ' f¸r die Sektionen: 0 = Haupt, 2 = Fuﬂ
 Dim MaxWeite&, aktWeite&
' Dim TabI() As TabInd
' ReDim TabI(0)
 Dim TabI%(), aktTabi%
 ReDim TabI(0)
 aktTabi = 0
 For i = 0 To MdB.Forms(0).COUNT - 1
  ReDim Preserve TabI(aktTabi)
  Select Case MdB.Forms(0)(i).ControlType
   Case 100
    If MdB.Forms(0)(i).Parent.name = AccForm Then
     TabI(aktTabi) = aktTabi + 1
    ElseIf MdB.Forms(0)(i).Parent.name = MdB.Forms(0)(i - 1).name Then
     TabI(aktTabi) = TabI(aktTabi - 1)
     TabI(aktTabi - 1) = TabI(aktTabi - 1) + 1
    Else
     MsgBox "Stop 1 bei ‹bertragFormulare; " & vbCrLf & "MdB.Forms(0).parent.name: " & MdB.Forms(0).Parent.name & "i: " & i
     Stop
    End If
   Case Else: TabI(aktTabi) = aktTabi + 1
  End Select
  aktTabi = aktTabi + 1
 Next i
 For i = 0 To MdB.Forms(0).COUNT - 1
'  IF MdB.Forms(0)(i).ControlType <> 100 THEN
'   TabI(UBound(TabI)).Ctl = MdB.Forms(0)(i).Name
'   TabI(UBound(TabI)).TI = MdB.Forms(0)(i).TabIndex
'   ReDim TabI(UBound(TabI) + 1)
'  END IF
  Sekt = MdB.Forms(0)(i).Section
  aktHˆhe = MdB.Forms(0)(i).Top + MdB.Forms(0)(i).Height
  If aktHˆhe > MaxHˆhe(Sekt) Then MaxHˆhe(Sekt) = aktHˆhe
  aktWeite = MdB.Forms(0)(i).Left + MdB.Forms(0)(i).Width
  If aktWeite > MaxWeite Then MaxWeite = aktWeite
 Next i
 GesHˆhe = MaxHˆhe(0) + MaxHˆhe(1) + MaxHˆhe(2) + IIf(obDatenS‰tze, 600, 0)
 
 lies.Ausgeb "÷ffne Datei " & FName & ".frm", False
 Open App.path & "\" & FName & ".frm" For Output As #309
 Open App.path & "\" & FName & "IndexZuordnung.txt" For Output As #310
 Print #309, "VERSION 5.00"
 Print #309, "Begin VB.Form " & FName
 Print #309, "   BorderStyle     =   3  'Fester Dialog"
 Print #309, "   Caption         =   """ & FName & """"
 Print #309, "   ClientHeight    =   " & GesHˆhe ' 13230" '& MdB.Forms(0).WindowHeight
 Print #309, "   ClientLeft      =   45"
 Print #309, "   ClientTop       =   330"
 Print #309, "   ClientWidth     =   " & MaxWeite ' 19110" '& MdB.Forms(0).WindowWidth
 Print #309, "   KeyPreview = -1         'True"
 Print #309, "   LinkTopic = """ & FName & """"
 Print #309, "   MaxButton = 0           'False"
 Print #309, "   MinButton = 0           'False"
 Print #309, "   ScaleHeight = " & MdB.Forms(0).WindowHeight
 Print #309, "   ScaleWidth = " & MdB.Forms(0).WindowWidth
 Dim iLab&, iTextB&, iCheckB&, iCommandB&, iComboB&, iOptionB&, sName$, PunktZahl&
  Dim CtlArt$(6)
  CtlArt(1) = "vLab"
  CtlArt(2) = "vTextB"
  CtlArt(3) = "vCheckb"
  CtlArt(4) = "vCommandB"
  CtlArt(5) = "vComboB"
  CtlArt(6) = "vOptionB"
 Dim CtlNm$()
 ReDim CtlNm(UBound(CtlArt), 0) ' Control-Namen, Reihenfolg s.u.
 iLab = 1
 iTextB = 1
 iCheckB = 1
 iCommandB = 1
 iComboB = 1
 iOptionB = 1
 PunktZahl = 0
 lies.Ausgeb "‹bertrage Kontrollelemente ... ", False
 lies.Ausgeb vbCrLf, False
 For i = 0 To MdB.Forms(0).COUNT - 1
  lies.Ausgeb ".", False
  PunktZahl = PunktZahl + 1
  If PunktZahl > 50 Then
   PunktZahl = 0
   lies.Ausgeb vbCrLf, False
  End If
  On Error GoTo fehler
  sName = REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(MdB.Forms(0)(i).name, " ", "_"), "ﬂ", "ss"), "ˆ", "oe"), "¸", "ue"), "?", ""), "-", "_"), ",", ""), "(", ""), ")", "")
  Dim LTabInd&
  Select Case MdB.Forms(0)(i).ControlType
   Case 100
    Print #309, "   Begin VB.Label " & CtlArt(1) ' vLab" '& SName
    Print #309, "      Caption         =   """ & MdB.Forms(0)(i).Caption & """"
'    Print #309, "      TabIndex        =   " & LTabInd
    Print #309, "      Index           =   " & iLab
'    Dim VglName$
'    VglName = MdB.Forms(0)(i).Parent.Name
'    For j = 0 To UBound(TabI) - 1
'     IF TabI(j).Ctl = VglName THEN
'      Print #309, "      TabIndex        =   " & TabI(j).TI
'      Exit For
'     END IF
'    Next j
    Print #310, "vLab(" & iLab & ") <- " & MdB.Forms(0)(i).name
    If UBound(CtlNm, 2) < iLab Then ReDim Preserve CtlNm(UBound(CtlArt), iLab)
    CtlNm(1, iLab) = MdB.Forms(0)(i).name
    iLab = iLab + 1
   Case 109
    Print #309, "   Begin VB.TextBox " & CtlArt(2) ' vTextB" '& SName
    CS = REPLACE$(MdB.Forms(0)(i).ControlSource, "`", "`")
    CS = REPLACE$(CS, "`", "`")
    CS = REPLACE$(CS, "`", "")
    If InStrB(CS, "=") > 0 Then
       If InStrB(CS, "grˆﬂe") > 0 Then
        CS = "bmi"
       ElseIf InStrB(CS, "nachname") > 0 Then
        CS = "gesname"
       Else
        CS = vNS
       End If
    End If
    Print #309, "      DataField = """ & CS & """"
    If MdB.Forms(0)(i).Height > 280 Then
     Print #309, "      MultiLine       =   -1  'True"
     Print #309, "      ScrollBars      =   2  'Vertikal"
    End If
    Print #309, "      Index           =   " & iTextB
    Print #310, "vTextBox(" & iTextB & ") <- " & MdB.Forms(0)(i).name
    If UBound(CtlNm, 2) < iTextB Then ReDim Preserve CtlNm(UBound(CtlArt), iTextB)
    CtlNm(2, iTextB) = MdB.Forms(0)(i).name
    iTextB = iTextB + 1
'    LTabInd = MdB.Forms(0)(i).TabIndex
'    Print #309, "      TabIndex        =   " & LTabInd
   Case 106
    Print #309, "   Begin VB.CheckBox " & CtlArt(3) ' vCheckb" '& SName
    CS = REPLACE$(MdB.Forms(0)(i).ControlSource, "`", "`")
    CS = REPLACE$(CS, "`", "`")
    CS = REPLACE$(CS, "`", "")
    Print #309, "      Caption         =   """ & MdB.Forms(0)(i).ControlName & """"
'    LTabInd = MdB.Forms(0)(i).TabIndex
'    Print #309, "      TabIndex        =   " & LTabInd
    Print #309, "      Index           =   " & iCheckB
    Print #309, "      DataField = """ & CS & """"
    Print #310, "vCheckb(" & iCheckB & ") <- " & MdB.Forms(0)(i).ControlName
    If UBound(CtlNm, 2) < iCheckB Then ReDim Preserve CtlNm(UBound(CtlArt), iCheckB)
    CtlNm(3, iCheckB) = MdB.Forms(0)(i).ControlName
    iCheckB = iCheckB + 1
   Case 104
    Print #309, "   Begin VB.CommandButton " & CtlArt(4) ' vCommandB" '& SName
    Print #309, "      Caption         =   """ & MdB.Forms(0)(i).ControlName & """"
'    LTabInd = MdB.Forms(0)(i).TabIndex
'    Print #309, "      TabIndex        =   " & LTabInd
    Print #309, "      Index           =   " & iCommandB
    Print #310, "vCommandB(" & iCommandB & ") <- " & MdB.Forms(0)(i).ControlName
    If UBound(CtlNm, 2) < iCommandB Then ReDim Preserve CtlNm(UBound(CtlArt), iCommandB)
    CtlNm(4, iCommandB) = MdB.Forms(0)(i).name
    iCommandB = iCommandB + 1
   Case 111
    Print #309, "   Begin VB.ComboBox " & CtlArt(5) ' vComboB" '& SName
    Print #309, "      text = """"" ' & MdB.Forms(0)(i).Caption"
    Print #309, "      Index           =   " & iComboB
'    LTabInd = MdB.Forms(0)(i).TabIndex
'    Print #309, "         TabIndex        =   " & LTabInd
    Print #310, "vComboB(" & iComboB & ") <- " & MdB.Forms(0)(i).ControlName
    If UBound(CtlNm, 2) < iComboB Then ReDim Preserve CtlNm(UBound(CtlArt), iComboB)
    CtlNm(5, iComboB) = MdB.Forms(0)(i).name
    iComboB = iComboB + 1
   Case 122
    Print #309, "      Begin VB.CheckBox " & CtlArt(6) ' vOptionB" '& SName
    Print #309, "         Caption         =   """ & MdB.Forms(0)(i).Caption & """"
    Print #309, "         Style = 1              'Grafisch"
'    LTabInd = MdB.Forms(0)(i).TabIndex
'    Print #309, "         TabIndex        =   " & LTabInd
    Print #309, "         Index           =   " & iOptionB
    Print #309, "      DataField = """ & CS & """"

    Print #310, "vOptionB(" & iOptionB & ") <- " & MdB.Forms(0)(i).ControlName
    If UBound(CtlNm, 2) < iOptionB Then ReDim Preserve CtlNm(UBound(CtlArt), iOptionB)
    CtlNm(6, iOptionB) = MdB.Forms(0)(i).name
    iOptionB = iOptionB + 1
   Case Else
'    LTabInd = MdB.Forms(0)(i).TabIndex
    MsgBox "Stop 2 bei ‹bertragformulare:" & vbCrLf & "MdB.Forms(0)(i).ControlType: " & MdB.Forms(0)(i).ControlType & vbCrLf & "i: " & i
    Stop
  End Select
'  If 1 = 0 Then
'    For j = 0 To MdB.Forms(0)(i).Properties.COUNT - 1
'     Debug.Print MdB.Forms(0)(i).Properties(j).name & " "; MdB.Forms(0)(i).Properties(j)
'    Next j
'  End If
  On Error Resume Next
  Print #309, "      TabIndex        =   " & TabI(i)
  Print #309, "      BackColor       =   " & MdB.Forms(0)(i).BackColor
  On Error GoTo fehler
  Print #309, "      Height          =   " & MdB.Forms(0)(i).Height
  Print #309, "      Left            =   " & MdB.Forms(0)(i).Left
  Print #309, "      Top             =   " & (IIf(MdB.Forms(0)(i).Section = 2, MaxHˆhe(0), 0) + MdB.Forms(0)(i).Top)
  Print #309, "      Width           =   " & MdB.Forms(0)(i).Width
  Print #309, "   End"
  On Error GoTo fehler
 Next i
 Dim CtlTxt$() ' Programm-Text hinter den Controls, 0.Dimension:
 Dim CtlTp$(9)
 CtlTp(0) = "Click"
' CtlTp(1) = "DblClick"
' CtlTp(2) = "Enter"
' CtlTp(3) = "Exit"
 CtlTp(1) = "GotFocus"
 CtlTp(2) = "KeyDown"
 CtlTp(3) = "KeyPress"
 CtlTp(4) = "KeyUp"
 CtlTp(5) = "LostFocus"
 CtlTp(6) = "MouseDown"
 CtlTp(7) = "MouseMove"
 CtlTp(8) = "MouseUp"
 CtlTp(9) = "Change"
 
 'Click,DblClick,Enter,Exit,GotFocus,KeyDown,KeyPress,KeyUp,LostFocus,MouseDown,MouseMove,MouseUp
 ReDim CtlTxt$(UBound(CtlTp), UBound(CtlArt), UBound(CtlNm, 2))
 
 If obDatenS‰tze Then
  Print #309, "   Begin VB.PictureBox picButtons"
  Print #309, "      Align = 2              'Unten ausrichten"
  Print #309, "      Appearance = 0         '2D"
  Print #309, "      BorderStyle = 0        'Kein"
  Print #309, "      ForeColor = &H80000008"
  Print #309, "      Height = 300"
  Print #309, "      Left = 0"
  Print #309, "      ScaleHeight = 300"
  Print #309, "      ScaleWidth = 7875"
  Print #309, "      TabIndex = 386"
  Print #309, "      Top = 10515"
  Print #309, "      Width = 7875"
  Print #309, "      Begin VB.CommandButton cmdAdd"
  Print #309, "         Caption = ""Hinzuf¸gen"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 59"
  Print #309, "         TabIndex = 1387"
  Print #309, "         Top = 0"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdEdit"
  Print #309, "         Caption = ""Bearbeiten"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 1213"
  Print #309, "         TabIndex = 1388"
  Print #309, "         Top = 0"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdDelete"
  Print #309, "         Caption = ""Lˆschen"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 2367"
  Print #309, "         TabIndex = 1389"
  Print #309, "         Top = 0"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdRefresh"
  Print #309, "         Caption = ""neu lesen"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 3521"
  Print #309, "         TabIndex = 1390"
  Print #309, "         Top = 0"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdClose"
  Print #309, "         Caption = ""S&chlieﬂen"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 4675"
  Print #309, "         TabIndex = 1391"
  Print #309, "         Top = 0"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdUpdate"
  Print #309, "         Caption = ""Aktualisieren"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 59"
  Print #309, "         TabIndex = 1392"
  Print #309, "         Top = 0"
  Print #309, "         Visible = 0             'False"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdCancel"
  Print #309, "         Caption = ""Abbrechen"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 1213"
  Print #309, "         TabIndex = 1393"
  Print #309, "         Top = 0"
  Print #309, "         Visible = 0             'False"
  Print #309, "         Width = 1095"
  Print #309, "      End"
  Print #309, "   End"
  Print #309, "   Begin VB.PictureBox picStatBox"
  Print #309, "      Align = 2              'Unten ausrichten"
  Print #309, "      Appearance = 0         '2D"
  Print #309, "      BorderStyle = 0        'Kein"
  Print #309, "      ForeColor = &H80000008"
  Print #309, "      Height = 300"
  Print #309, "      Left = 0"
  Print #309, "      ScaleHeight = 300"
  Print #309, "      ScaleWidth = 7875"
  Print #309, "      TabIndex = 380"
  Print #309, "      Top = 10815"
  Print #309, "      Width = 7875"
  Print #309, "      Begin VB.CommandButton Suchen"
  Print #309, "         Caption = ""&Suchen"""
  Print #309, "         Height = 255"
  Print #309, "         Left = 4920"
  Print #309, "         TabIndex = 440"
  Print #309, "         Top = 0"
  Print #309, "         Width = 735"
  Print #309, "      End"
  Print #309, "      Begin VB.TextBox suche"
  Print #309, "         Height = 285"
  Print #309, "         Left = 5640"
  Print #309, "         TabIndex = 441"
  Print #309, "         Top = 0"
  Print #309, "         Width = 2655"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdFirst"
  Print #309, "         Height = 300"
  Print #309, "         Left = 0"
  Print #309, "         Picture         =   ""AnTest.frx"":0000"
  Print #309, "         Style = 1              'Grafisch"
  Print #309, "         TabIndex = 1381"
  Print #309, "         Top = 0"
  Print #309, "         UseMaskColor = -1       'True"
  Print #309, "         Width = 345"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdPrevious"
  Print #309, "         Caption = ""&¸"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 345"
  Print #309, "         Picture         =   ""AnTest.frx"":0684"
  Print #309, "         Style = 1              'Grafisch"
  Print #309, "         TabIndex = 1382"
  Print #309, "         Top = 0"
  Print #309, "         UseMaskColor = -1       'True"
  Print #309, "         Width = 345"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdNext"
  Print #309, "         Caption = ""&‰"""
  Print #309, "         Height = 300"
  Print #309, "         Left = 4200"
  Print #309, "         Picture         =   ""AnTest.frx"":0342"
  Print #309, "         Style = 1              'Grafisch"
  Print #309, "         TabIndex = 1383"
  Print #309, "         Top = 0"
  Print #309, "         UseMaskColor = -1       'True"
  Print #309, "         Width = 345"
  Print #309, "      End"
  Print #309, "      Begin VB.CommandButton cmdLast"
  Print #309, "         Height = 300"
  Print #309, "         Left = 4545"
  Print #309, "         Picture         =   ""AnTest.frx"":09C6"
  Print #309, "         Style = 1              'Grafisch"
  Print #309, "         TabIndex = 1384"
  Print #309, "         Top = 0"
  Print #309, "         UseMaskColor = -1       'True"
  Print #309, "         Width = 345"
  Print #309, "      End"
  Print #309, "      Begin VB.Label lblStatus"
  Print #309, "         BackColor = &HFFFFFF   'weiss"
  Print #309, "         BorderStyle = 1        'Fest Einfach"
  Print #309, "         Height = 285"
  Print #309, "         Left = 690"
  Print #309, "         TabIndex = 1385"
  Print #309, "         Top = 0"
  Print #309, "         Width = 3360"
  Print #309, "      End"
  Print #309, "   End"
 End If
 Print #309, "End"
 Print #309, "Attribute VB_Name = """ & FName & """"
 Print #309, "Attribute VB_GlobalNameSpace = False"
 Print #309, "Attribute VB_Creatable = False"
 Print #309, "Attribute VB_PredeclaredId = True"
 Print #309, "Attribute VB_Exposed = False"
 Print #309, "Public WithEvents anaRS As Recordset"
 Print #309, "Attribute anaRS.VB_VarHelpID = -1"
 Print #309, "Dim mbChangedByCode As Boolean"
 Print #309, "Dim mvBookMark As Variant"
 Print #309, "Dim mbEditFlag As Boolean"
 Print #309, "Dim mbAddNewFlag As Boolean"
 Print #309, "Public mbDataChanged As Boolean"
 If obDatenS‰tze Then
  Print #309, "Dim altsuche$, SuchStringGe‰ndert%"
 End If
 Print #309, "Private Sub anaRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)"
 Print #309, "  'Hierdurch wird die aktuelle Datensatzposition f¸r diese Datensatzgruppe angezeigt"
 Print #309, "   lblStatus.Caption = CStr(anaRS.AbsolutePosition)"
 Print #309, "   Call do_Form_Current" & "_" & FName & "(Me)"
 Print #309, "End Sub 'anaRS_MoveComplete"
 Print #309, ""
 Print #309, "Private Sub anaRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As Adodb.Recordset)"
 Print #309, "  'Hier kˆnnen Sie Code zur ‹berpr¸fung einf¸gen"
 Print #309, "  'Dieses Ereignis wird aufgerufen, wenn die folgenden Aktionen eintreten"
 Print #309, "  Dim bCancel As Boolean"
 Print #309, "  SELECT CASE adReason"
 Print #309, "  Case adRsnAddNew"
 Print #309, "  Case adRsnClose"
 Print #309, "  Case adRsnDelete"
 Print #309, "  Case adRsnFirstChange"
 Print #309, "  Case adRsnMove"
 Print #309, "  Case adRsnRequery"
 Print #309, "  Case adRsnResynch"
 Print #309, "  Case adRsnUndoAddNew"
 Print #309, "  Case adRsnUndoDelete"
 Print #309, "  Case adRsnUndoUpdate"
 Print #309, "  Case adRsnUpdate"
 Print #309, "  END SELECT"
 Print #309, "  IF bCancel THEN adStatus = adStatusCancel"
 Print #309, "End Sub 'anaRS_WillChangeRecord"
 Print #309, ""
 Print #309, "Public Sub cmdAdd_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  With anaRS"
 Print #309, "    IF Not (.BOF AND .EOF) THEN"
 Print #309, "      mvBookMark = .Bookmark"
 Print #309, "    END IF"
 Print #309, "    .AddNew"
 Print #309, "    lblStatus.Caption = ""Datensatz; hinzuf¸gen"""
 Print #309, "    mbAddNewFlag = True"
 Print #309, "    SetButtons False"
 Print #309, "  END With"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdAdd_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub 'cmdAdd_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdDelete_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  With anaRS"
 Print #309, "    .Delete"
 Print #309, "    .MoveNext"
 Print #309, "    IF .EOF THEN .MoveLast"
 Print #309, "  END With"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdDelete_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' Private Sub cmdDelete_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdRefresh_Click()"
 Print #309, "  'Dies wird nur f¸r Mehrbenutzeranwendungen benˆtigt"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  anaRS.Requery"
 Print #309, " Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdRefresh_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' Private Sub cmdRefresh_Click()"""
 Print #309, ""
 Print #309, "Public Sub cmdEdit_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  lblStatus.Caption = ""Datensatz bearbeiten"""
 Print #309, "  mbEditFlag = True"
 Print #309, "  SetButtons False"
 Print #309, " Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdEdit_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub 'cmdEdit_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdCancel_Click()"
 Print #309, "  ON Error Resume Next"
 Print #309, "  SetButtons True"
 Print #309, "  mbEditFlag = False"
 Print #309, "  mbAddNewFlag = False"
 Print #309, "  anaRS.CancelUpdate"
 Print #309, " IF mvBookMark > 0 THEN"
 Print #309, "    anaRS.Bookmark = mvBookMark"
 Print #309, "  Else"
 Print #309, "    anaRS.MoveFirst"
 Print #309, "  END IF"
 Print #309, "  mbDataChanged = False"
 Print #309, "End Sub ' cmdCancel_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdUpdate_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  anaRS.UpdateBatch adAffectAll"
 Print #309, "  IF mbAddNewFlag THEN"
 Print #309, "    anaRS.MoveLast              'Zu neuem Datensatz gehen"
 Print #309, "  END IF"
 Print #309, "  mbEditFlag = False"
 Print #309, "  mbAddNewFlag = False"
 Print #309, "  SetButtons True"
 Print #309, "  mbDataChanged = False"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdUpdate_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' cmdUpdate_Click()"
 Print #309, ""
 Print #309, "Private Sub cmdClose_Click()"
 Print #309, "  Unload Me"
 Print #309, "End Sub"
 Print #309, ""
 Print #309, "Public Sub cmdFirst_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  anaRS.MoveFirst"
 Print #309, "  mbDataChanged = False"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdFirst_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' cmdFirst_Click"
 Print #309, ""
 Print #309, "Public Sub cmdLast_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  anaRS.MoveLast"
 Print #309, "  mbDataChanged = False"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdLast_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' cmdLast_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdNext_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  IF Not anaRS.EOF THEN anaRS.MoveNext"
 Print #309, "  IF anaRS.EOF AND anaRS.RecordCount > 0 THEN"
 Print #309, "    T¸t 1760,1000"
 Print #309, "     'Ende der Zeile wurde erreicht; zur¸ck zum Zeilenanfang"
 Print #309, "    anaRS.MoveLast"
 Print #309, "  END IF"
 Print #309, "  'Aktuellen Datensatz anzeigen"
 Print #309, "  mbDataChanged = False"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdNext_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' cmdNext_Click()"
 Print #309, ""
 Print #309, "Public Sub cmdPrevious_Click()"
 Print #309, "  ON Error GoTo fehler"
 Print #309, "  IF Not anaRS.BOF THEN anaRS.MovePrevious"
 Print #309, "  IF anaRS.BOF AND anaRS.RecordCount > 0 THEN"
 Print #309, "    T¸t 1760,1000"
 Print #309, "    'Ende der Zeile wurde erreicht; zur¸ck zum Zeilenanfang"
 Print #309, "    anaRS.MoveFirst"
 Print #309, "  END IF"
 Print #309, "  'Aktuellen Datensatz anzeigen"
 Print #309, "  mbDataChanged = False"
 Print #309, "  Exit Sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in cmdPrevious_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub ' cmdPrevious_Click()"
 Print #309, ""
 Print #309, "Public Sub SetButtons(bVal As Boolean)"
 Print #309, "  cmdAdd.Visible = bVal"
 Print #309, "  cmdEdit.Visible = bVal"
 Print #309, "  cmdUpdate.Visible = Not bVal"
 Print #309, "  cmdCancel.Visible = Not bVal"
 Print #309, "  cmdDelete.Visible = bVal"
 Print #309, "  cmdClose.Visible = bVal"
 Print #309, "  cmdRefresh.Visible = bVal"
 Print #309, "  cmdNext.Enabled = bVal"
 Print #309, "  cmdFirst.Enabled = bVal"
 Print #309, "  cmdLast.Enabled = bVal"
 Print #309, "  cmdPrevious.Enabled = bVal"
 Print #309, "End Sub ' SetButtons"
 Print #309, "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)"
 Print #309, " Call doKeyDown(Me, KeyCode, Shift)"
 Print #309, "End Sub"
 
' IF AccName = "Anamnesebogen" THEN
 If obDatenS‰tze Then
 Print #309, ""
 Print #309, "Public Sub Suchen_Click()"
 Print #309, " ON Error GoTo fehler"
 Print #309, " IF SuchStringGe‰ndert THEN"
 Print #309, "  altsuche = Me.suche"
 Print #309, "  IF IsNumeric(Me.suche) THEN"
 Print #309, "   anaRS.Find "" Pat_id = "" & Me.suche, 0, adSearchForward"
 Print #309, "  Else"
 Print #309, "   IF anaRS.EOF THEN anaRS.MoveFirst"
 Print #309, "   anaRS.Find ""gesname LIKE '"" & Me.suche & ""*'"", 1, adSearchForward"
 Print #309, "  END IF"
 Print #309, " Else"
 Print #309, "  Me.suche.SetFocus"
 Print #309, "  SuchStringGe‰ndert = True"
 Print #309, " END IF"
 Print #309, " exit sub"
 Print #309, "fehler:"
 Print #309, " Dim AnwPfad$"
 Print #309, "#If VBA6 THEN"
 Print #309, " AnwPfad = currentDB.Name"
 Print #309, "#Else"
 Print #309, " AnwPfad = App.Path"
 Print #309, "#END IF"
 Print #309, " SELECT CASE MsgBox(""FNr: "" + CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.Source), """", CStr(Err.Source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""aufgefangener Fehler in Suchen_Click/"" + AnwPfad)"
 Print #309, "  Case vbAbort: Call MsgBox(""Hˆre auf""): Progende"
 Print #309, "  Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
 Print #309, "  Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
 Print #309, " END SELECT"
 Print #309, "End Sub 'Suchen_Click()"
 End If ' obDatenS‰tze


 Dim zeile$, obFL%, obRs%, obFuL%, obKD%
 obFL = 0
 obRs = 0
 obFuL = 0
 obKD = 0
 Dim obZeileVerwertet%
 On Error Resume Next
 
 Dim Modul$
 Modul = MdB.Modules("Form_" & AccForm).name
 On Error GoTo fehler
 If LenB(Modul) <> 0 Then
 For i = 1 To MdB.Modules("Form_" & AccForm).CountOfLines
  zeile = MdB.Modules("Form_" & AccForm).Lines(i, 1)
  If zeile Like "Option Compare*" Then GoTo weiter
  
  obZeileVerwertet = 0
  If zeile Like "Private Sub *" Then
   For j = 1 To UBound(CtlArt)
    For k = 1 To UBound(CtlNm, 2)
     If LenB(CtlNm(j, k)) = 0 Then Exit For
     If zeile Like "Private Sub " & CtlNm(j, k) & "_*" Then
      For m = 0 To UBound(CtlTp)
       If zeile Like "*_" & CtlTp(m) & "(*" Then
        obZeileVerwertet = True
        Do
         i = i + 1
         If i > MdB.Modules("Form_" & AccForm).CountOfLines Then Exit Do
         zeile = MdB.Modules("Form_" & AccForm).Lines(i, 1)
         If zeile Like "End Sub*" Then Exit Do
         If m = 3 Then zeile = REPLACE$(zeile, "KeyCode", "KeyAscii")
         CtlTxt(m, j, k) = CtlTxt(m, j, k) & IIf(LenB(CtlTxt(m, j, k)) = 0, vNS, vbCrLf) & "  " & zeile
        Loop
       End If
       If obZeileVerwertet Then Exit For
      Next m
     End If
     If obZeileVerwertet Then Exit For
    Next k
    If obZeileVerwertet Then Exit For
   Next j
  End If
  If Not obZeileVerwertet Then
   Print #309, zeile
   If InStrB(zeile, "Form_Load()") > 0 And Not obFL Then
    Call FormLoadEinschub(MyTab, iTextB, iCheckB)
    obFL = -1
   End If
   If InStrB(zeile, "Form_Resize()") > 0 And Not obRs Then
    Call FormResizeEinschub
    obRs = -1
   End If
   If InStrB(zeile, "Form_Unload(") > 0 And Not obFuL Then
    Call FormUnLoadEinschub
    obFuL = -1
   End If
   If InStrB(zeile, "Form_KeyDown(") > 0 And Not obKD Then
    Call FormKDEinschub
    obKD = -1
   End If
  End If ' not obZeileVerwertet
weiter:
 Next i
 End If ' lenb(Modul)<>0
 Dim obBeg%, obSubBeg%, ‹berS$ ' ob ¸berhaupt / mit Sub begonnen
 For j = 1 To UBound(CtlArt) ' vLab, vTextB, vCheckb
  For m = 0 To UBound(CtlTp) ' Click, GotFocus
   For k = 1 To UBound(CtlNm, 2) ' Control-Namen
    If LenB(CtlNm(j, k)) = 0 Then Exit For ' dann keine weiteren Control-Namen mehr bef¸llt
    If LenB(CtlTxt(m, j, k)) <> 0 Then
     If Not obSubBeg Then
      If Not obBeg Then
       obBeg = True
      Else
       Print #309, " END SELECT ' Index"
       Print #309, "End Sub '" & ‹berS
       Print #309, ""
      End If
      ‹berS = "Private Sub " & CtlArt(j) & "_" & CtlTp(m) & "(Index As Integer"
      Select Case m
       Case 0, 1, 5, 9: ‹berS = ‹berS & ")"
       Case 2, 4: ‹berS = ‹berS & ", KeyCode As Integer, Shift As Integer)"
       Case 3: ‹berS = ‹berS & ", KeyAscii As Integer)"
       Case 6, 7, 8: ‹berS = ‹berS & ", Button As Integer, Shift As Integer, X As Single, Y As Single)"
      End Select
      Print #309, ‹berS
      Print #309, " SELECT CASE Index"
      obSubBeg = True
     End If
     Print #309, "  Case " & k & " ' " & CtlNm(j, k)
     Print #309, CtlTxt(m, j, k)
    End If
   Next k
   obSubBeg = 0
  Next m
 Next j
 If obBeg Then
  Print #309, " END SELECT ' Index"
  Print #309, "End Sub '" & ‹berS
  Print #309, ""
 End If
 
 If Not obFL Then
  Print #309, ""
  Print #309, "  Private Sub Form_Load()"
  Call FormLoadEinschub(MyTab, iTextB, iCheckB)
  Print #309, "  END Sub 'Private Sub Form_Load()"
 End If
 If Not obRs Then
  Print #309, ""
  Print #309, "Private Sub Form_Resize()"
  Call FormResizeEinschub
  Print #309, "End Sub"
 End If
 If Not obFuL Then
  Print #309, ""
  Print #309, "Private Sub Form_Unload(Cancel As Integer)"
  Call FormUnLoadEinschub
  Print #309, "End Sub 'Private Sub Form_Unload()"
 End If
 If Not obKD Then
  Print #309, "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)"
  Print #309, "End Sub"
 End If
 Dim SL As New SortierListe, SSt As SortierString
 Set SL = Nothing
 
   For j = 1 To UBound(CtlArt)
    For k = 1 To UBound(CtlNm, 2)
     If LenB(CtlNm(j, k)) = 0 Then Exit For
     Print #309, "' " & Left$(CtlArt(j) & "(" & k & ")" & Space$(15), 15) & ":" & CtlNm(j, k)
     Set SSt = New SortierString
     SSt.Stri = "'" & Left$(CtlNm(j, k) & ":" & Space$(40), 40) & CtlArt(j) & "(" & k & ")"
     SL.sCAdd SSt
    Next k
   Next j
   
 Print #309, ""
 For i = 1 To SL.COUNT
  Print #309, SL.Item(i).Stri
 Next i
 
 Close #309
 Close #310
 zeigan App.path & "\" & FName & ".frm"
 zeigan App.path & "\" & FName & "IndexZuordnung.txt"
 Call MdB.CloseCurrentDatabase
 lies.Ausgeb "Fertig mit Erstellen des Formulars " & FName, True
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ‹bertragFormulare/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function                ' ‹bertragFormulare

Private Function FormLoadEinschub(MyTab$, iTextB&, iCheckB&)
   Print #309, "  Call doForm_Load(Me)"
''   Const ConString$ = """PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=mitte;uid=praxis;pwd=***REMOVED***;database=quelle;"""
'   Dim SelString$
'   Print #309, "  Dim db As Connection"
'   Print #309, "  SET db = New Connection"
'   Print #309, "  db.CursorLocation = adUseClient"
'   Call lies.dbv.cnVorb("quelle", "anamnesebogen", "Patientendaten")
'   Print #309, "  db.Open dbcn" '""" & lies.dbv.wCn.ConnectionString & """" ' ConString
'   Print #309, "  SET anaRS = New Recordset"
'' =Wenn(`Grˆﬂe`=0;"";`Gewicht`/`Grˆﬂe`/`Grˆﬂe`*Wenn(`Grˆﬂe`>3;10000;1))
'    SELECT CASE MyTab
'    Case "anamnesebogen"
''     SelString = """SELECT --tkz AS j_tkz, --insulinpumpe AS j_insulinpumpe, `" & MyTab & "`.*,IF(grˆﬂe=0,'',gewicht/grˆﬂe/grˆﬂe*if(grˆﬂe>3,10000,1)) AS bmi, CONCAT(nachname°' '° vorname) AS gesname  FROM `" & MyTab & "`"", db, adOpenStatic, adLockOptimistic"
'     SelString = """SELECT *,IF(grˆﬂe=0,'',gewicht/grˆﬂe/grˆﬂe*if(grˆﬂe>3,10000,1)) AS bmi, CONCAT(nachname°' '° vorname) AS gesname  FROM " & MyTab & """, db, adOpenStatic, adLockOptimistic"
'     IF lies.obMySQL THEN
'      SelString = replace$(SelString, "°", ",")
'     Else
'      SelString = replace$(replace$(SelString, "concat", ""), "°", " & ")
'     END IF
'    Case Else
'     SelString = """SELECT * FROM " & MyTab & """, db, adOpenStatic, adLockOptimistic"
'   END SELECT
''   Dim rSc As New ADODB.Recordset
''   SET rSc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "" & MyTab & vns, Empty))
''   Do While Not rSc.EOF
''   Loop
'   Print #309, "  anaRS.Open " & SelString
'   IF iTextB > 0 THEN
'    Print #309, "  Dim oText AS TextBox"
'    Print #309, "  'Textfelder an Datenprovider binden"
'    Print #309, "  For Each oText In Me.vTextB"
'    Print #309, "    SET oText.DataSource = anaRS"
'    Print #309, "  Next"
'   END IF
'   IF iCheckB > 0 THEN
'    Print #309, "  Dim oCheck As CheckBox"
'    Print #309, "  'Kontrollk‰stchen an Datenprovider binden"
'    Print #309, "  For Each oCheck In Me.vCheckb"
'    Print #309, "    SET oCheck.DataSource = anaRS"
'    Print #309, "  Next"
'   END IF
'   Print #309, "  mbDataChanged = False"
End Function ' FormLoadEinschub

Private Function FormResizeEinschub()
 Print #309, "  ON Error Resume Next"
 Print #309, "'  lblStatus.Width = Me.Width - 1500"
 Print #309, "'  cmdNext.Left = lblStatus.Width + 700"
 Print #309, "'  cmdLast.Left = cmdNext.Left + 340"
End Function ' FormResizeEinschub

Private Function FormUnLoadEinschub()
 Print #309, "  Screen.MousePointer = vbDefault"
End Function ' FormUnLoadEinschub

Private Function FormKDEinschub()
 Print #309, "  IF mbEditFlag OR mbAddNewFlag THEN Exit Sub"
 Print #309, "  SELECT CASE KeyCode"
 Print #309, "    Case vbKeyEscape"
 Print #309, "      cmdClose_Click"
 Print #309, "    Case vbKeyEnd"
 Print #309, "      cmdLast_Click"
 Print #309, "    Case vbKeyHome"
 Print #309, "      cmdFirst_Click"
 Print #309, "    Case vbKeyUp, vbKeyPageUp"
 Print #309, "      IF Shift = vbCtrlMask THEN"
 Print #309, "        cmdFirst_Click"
 Print #309, "      Else"
 Print #309, "        cmdPrevious_Click"
 Print #309, "      END IF"
 Print #309, "    Case vbKeyDown, vbKeyPageDown"
 Print #309, "      IF Shift = vbCtrlMask THEN"
 Print #309, "        cmdLast_Click"
 Print #309, "      Else"
 Print #309, "        cmdNext_Click"
 Print #309, "      END IF"
 Print #309, "  END SELECT"
End Function ' FormKDEinschub

'Function umwfSQL$(q$) ' 11.7.09 jetzt in QuelleDB
' IF lies.obMySQL THEN
'  umwfSQL = replace$(replace$(Trim$(q), "\", "\\"), "'", "\'")
' Else
'  umwfSQL = q
' END IF
'End FUNCTION ' umwfSQL$(q$)

' Make the FlexGrid's columns big enough to hold all values.
Sub SizeColumns(ByVal flx As MSHFlexGrid, frm As Form, Optional allevoll%, Optional maxpix&, Optional abSpalte%)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Long
Dim r As Long
Dim c As Integer
On Error GoTo fehler
    max_row = flx.Rows - 1
    For c = abSpalte To flx.cols - 1
        max_wid = 0
        For r = 0 To max_row
            wid = frm.TextWidth(Left$(REPLACE$(REPLACE$(flx.TextMatrix(r, c), vbCrLf, "--"), Chr(10), "-"), 6129)) ' sonst ‹berlauf
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = (max_wid + 400) * IIf(allevoll, 0.9, IIf(c > 5, 1, 0.7))
        If maxpix <> 0 Then If flx.ColWidth(c) > maxpix Then flx.ColWidth(c) = maxpix
    Next c
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "aufgefangener Fehler in SizeColumns/" + App.path)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SizeColumns

Function doWirt()
 Dim DNam$
 DNam$ = pVerz & "Wirt.txt"
 Dim Pat_id$, rs As New ADODB.Recordset, Faelle As New ADODB.Recordset
' Dim lies As New Lese
 Load Lese
 Lese.obMySQL = True
 Pat_id = 283
 Open DNam For Output As #303
 Call acon(quelleT)
' Call Faelle.Open("SELECT DISTINCT pat_id FROM `faelle` WHERE quartal LIKE '_2005' AND schgr <> '90'", DBCn, adOpenDynamic, adLockReadOnly)
 myFrag Faelle, "SELECT DISTINCT pat_id FROM `faelle` WHERE quartal LIKE '_2005' AND schgr <> '90'"
 Do While Not Faelle.EOF
  Pat_id = CStr(Faelle!Pat_id)
'  SET rs = Nothing
'  myFrag rs, "SELECT * FROM `namen` WHERE pat_id = " & Pat_id
  Set rs = Nothing
'  Call rs.Open("SELECT * FROM `Formular` F WHERE FormVorl LIKE '%Kassenrezept%' AND zeitpunkt BETWEEN '2005-01-01' AND '2006-01-01' AND pat_id = " & Pat_id & " AND feld IN ('Datum','Dosierung','Medikamente') AND feldinh <> '-  -  -  -' ORDER BY foid", DBCn, adOpenDynamic, adLockReadOnly)
  myFrag rs, "SELECT * FROM `Formular` F WHERE FormVorl LIKE '%Kassenrezept%' AND zeitpunkt BETWEEN '2005-01-01' AND '2006-01-01' AND pat_id = " & Pat_id & " AND feld IN ('Datum','Dosierung','Medikamente') AND feldinh <> '-  -  -  -' ORDER BY foid"
  If Not rs.BOF Then
   Print #303, vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Patient: " & CStr(rs!Pat_id) & vbCrLf
   Dim DiagTab() As CString
   Print #303, "Diagnosen: " & REPLACE$(DiagString$(Pat_id, DiagTab, #1/1/2005#, True), vbVerticalTab, vbCrLf)
  End If
  Do While Not rs.EOF
   If rs!Feld = "Datum" Then Print #303, vbCrLf
   Print #303, rs!Feld, rs!FeldInh
   rs.Move 1
  Loop
  DoEvents
  Faelle.Move 1
 Loop
 Close #303
 zeigan DNam, vbNormalFocus
End Function ' doWirt

Function doHilfsmittelklassifikationen(frm As Lese)
 Dim sql$, rs As New ADODB.Recordset, i%, FNr&, Zahl&(3), summe&, DatumS$, Datum As Date
 DatumS = "1.6.08"
 Do
  DatumS = InputBox("Bitte Ab-Datum eingeben:", "Datumseingabe", DatumS)
  If IsDate(DatumS) Then
   Datum = CDate(DatumS)
   Exit Do
  End If
 Loop
 summe = 0
 Screen.MousePointer = vbHourglass
' Nadeln, nicht EnaDura
 myFrag rs, "SELECT MID(kvdt.feldinh,14) art, COUNT(0) zahl FROM ((SELECT pat_id, foid, feldinh, zeitpunkt FROM `formular` f WHERE zeitpunkt > " & DatFor_k(Datum) & " AND feld = 'Medikamente' AND (feldinh LIKE '%Fine%' OR (feldinh LIKE '%Nad%' AND NOT feldinh LIKE '%enad%') OR feldinh LIKE '%Lancet%' OR feldinh LIKE '%Lanzet%')) f LEFT JOIN (SELECT * FROM `formular` WHERE feld = 'KVDTVerordnenderArzt') kvdt ON f.foid = kvdt.foid) LEFT JOIN `namen` n ON f.pat_id = n.pat_id WHERE kvdt.feldinh LIKE 'Gerald Schade%' GROUP BY kvdt.feldinh ORDER BY art"
 altAusgabe = vbCrLf & altAusgabe
 Do While Not rs.EOF
  Select Case rs!art
   Case ".": Zahl(0) = rs!Zahl
   Case ",": Zahl(1) = rs!Zahl
   Case ";": Zahl(2) = rs!Zahl
   Case "-": Zahl(3) = rs!Zahl
  End Select
  summe = summe + rs!Zahl
  rs.Move 1
 Loop
 DoEvents
 frm.Ausgabe = "Lanzetten, Nadeln seit " & Format$(Datum, "short date") & ":" & vbCrLf & _
               "DiaReal      :" & Right$(Space$(6) & Zahl(0), 6) & " = " & Right$(Space$(2) & Round(Zahl(0) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "DiaExpress   :" & Right$(Space$(6) & Zahl(1), 6) & " = " & Right$(Space$(2) & Round(Zahl(1) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "Ratsapotheke :" & Right$(Space$(6) & Zahl(2), 6) & " = " & Right$(Space$(2) & Round(Zahl(2) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "kein Anbieter:" & Right$(Space$(6) & Zahl(3), 6) & " = " & Right$(Space$(2) & Round(Zahl(3) / summe * 100, 0), 2) & "%" & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Set rs = Nothing
 summe = 0
 myFrag rs, "SELECT MID(kvdt.feldinh,14) art, COUNT(0) zahl FROM (SELECT pat_id, foid, feldinh, zeitpunkt FROM `formular` f WHERE zeitpunkt > " & DatFor_k(Datum) & " AND feld = 'Medikamente' AND (feldinh LIKE '%TTR%' OR feldinh LIKE '%Sensor%' OR feldinh LIKE '%teststr%')) f LEFT JOIN (SELECT * FROM `formular` WHERE feld = 'KVDTVerordnenderArzt') kvdt USING (foid) LEFT JOIN `namen` n ON f.pat_id = n.pat_id WHERE kvdt.feldinh LIKE 'Gerald Schade%' GROUP BY kvdt.feldinh ORDER BY art"
 altAusgabe = vbCrLf & altAusgabe
 Do While Not rs.EOF
  Select Case rs!art
   Case ".": Zahl(0) = rs!Zahl
   Case ",": Zahl(1) = rs!Zahl
   Case ";": Zahl(2) = rs!Zahl
   Case "-": Zahl(3) = rs!Zahl
  End Select
  summe = summe + rs!Zahl
  rs.Move 1
 Loop
 DoEvents
 frm.Ausgabe = "Teststreifen seit " & Format$(Datum, "short date") & ":" & vbCrLf & _
               "DiaReal      :" & Right$(Space$(6) & Zahl(0), 6) & " = " & Right$(Space$(2) & Round(Zahl(0) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "DiaExpress   :" & Right$(Space$(6) & Zahl(1), 6) & " = " & Right$(Space$(2) & Round(Zahl(1) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "Ratsapotheke :" & Right$(Space$(6) & Zahl(2), 6) & " = " & Right$(Space$(2) & Round(Zahl(2) / summe * 100, 0), 2) & "%" & vbCrLf & _
               "kein Anbieter:" & Right$(Space$(6) & Zahl(3), 6) & " = " & Right$(Space$(2) & Round(Zahl(3) / summe * 100, 0), 2) & "%" & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Set rs = Nothing
 myFrag rs, "SELECT f.pat_id, n.nachname, n.vorname, n.gebdat, f.feldinh, f.zeitpunkt FROM (SELECT pat_id, foid, feldinh, zeitpunkt FROM `formular` f WHERE zeitpunkt > " & DatFor_k(Datum) & " AND feld = 'Medikamente' AND (feldinh LIKE '%TTR%' OR feldinh LIKE '%Fine%' OR feldinh LIKE '%Sensor%' OR (feldinh LIKE '%Nad%' AND NOT feldinh LIKE '%enadu%') OR feldinh LIKE '%Lancet%' OR feldinh LIKE '%Lanzet%' OR feldinh LIKE '%teststr%')) f LEFT JOIN (SELECT * FROM `formular` WHERE feld = 'KVDTVerordnenderArzt') kvdt USING (foid) LEFT JOIN `namen` n ON f.pat_id = n.pat_id WHERE kvdt.feldinh = 'Gerald Schade'"
 altAusgabe = vbCrLf & altAusgabe
 If rs.BOF Then
  frm.Ausgabe = "keine" & vbCrLf & altAusgabe
  altAusgabe = frm.Ausgabe
 Else
  Do While Not rs.EOF
   frm.Ausgabe = Right$(Space$(4) & rs!Pat_id, 4) & " " & rs!Nachname & " " & rs!Vorname & " " & rs!GebDat & " " & rs!FeldInh & " " & rs!Zeitpunkt & vbCrLf & altAusgabe
   altAusgabe = frm.Ausgabe
   rs.Move 1
  Loop
 End If
 DoEvents
 frm.Ausgabe = "fehlende Klassifikationen seit " & Format$(Datum, "short date") & ":" & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Screen.MousePointer = vbNormal
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doHilfsmittelklassifikationen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doHilfsmittelklassifikationen

Function fzsfuell(frm As Lese, abstand&, Optional obgestern) ' Abstand: 999 => unbekannt, 9999 => ohne abstand
  Dim sql$, vrs As New ADODB.Recordset, runde%, FNr&, rAf&
  Dim s0&, s1&, k0&, k1&, QT$
  Dim csql0 As CString
  Dim rsse() As rstype
  On Error GoTo fehler
  Screen.MousePointer = vbHourglass
  If abstand = 999 Then ' code f¸r unbekannt
   s0 = 0: s1 = 0: k0 = 0: k1 = 0
   sql = "SELECT to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) abstand"
   myFrag vrs, sql
   abstand = vrs!abstand
   Set vrs = Nothing
  End If ' abstand = 999
  frm.Ausgeb "Berechne Fallzahlst‰nde f¸r Abstand: " & abstand & ", Tage: " & obgestern, True
  Dim QBed$, QBedi$, VQBed$, zpQBed$
  If abstand = 9999 Then
   QBed = vNS
   QBedi = vNS
   VQBed = vNS
   zpQBed = vNS
  Else
   QBed = "AND TO_DAYS(fanf)-TO_DAYS((CONCAT(MID(quartal,2,4),'-',(LEFT(quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand
   QBedi = "AND TO_DAYS(fanf)-TO_DAYS((CONCAT(MID(i.quartal,2,4),'-',(LEFT(i.quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand
   VQBed = "AND TO_DAYS(fanf)-TO_DAYS((CONCAT(MID(vorquartal,2,4),'-',(LEFT(vorquartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand
   zpQBed = "AND TO_DAYS(e.zeitpunkt)-TO_DAYS((CONCAT(YEAR(e.zeitpunkt),'-',MONTH(e.zeitpunkt)-((MONTH(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand
  End If ' abstand = 9999
  ' n‰ Zeile 23.4.22
  myEFrag ("DELETE fz FROM fallzahlstand fz LEFT JOIN" & vbCrLf & _
     "(SELECT quartal, COUNT(DISTINCT pat_id) zahl" & vbCrLf & _
     "FROM faelle f" & vbCrLf & _
     "WHERE schgr <> '90' " & QBed & vbCrLf & _
     "AND CONCAT(MID(quartal,2),LEFT(quartal,1))>'20042'" & vbCrLf & _
     "GROUP BY quartal" & vbCrLf & _
     ") i" & vbCrLf & _
     "ON fz.quartal=i.quartal AND fz.kassenpat<>i.zahl" & vbCrLf & _
     "WHERE fz.tage=" & abstand & " AND (NOT ISNULL(i.zahl) OR (SELECT COUNT(0) FROM (SELECT * FROM fallzahlstand) i WHERE quartal=fz.quartal AND tage=fz.tage)>1);")
  sql = "SELECT i.quartal," & abstand & " tage,i.kassenpat,fz.kassenpat " & _
   "FROM (SELECT Quartal, COUNT(0) kassenpat " & _
   "FROM (SELECT quartal " & _
   "FROM `faelle` f " & _
   "WHERE schgr <> '90' " & QBed & " AND substr(quartal,2,4)> '2003' AND quartal <> '12004' AND quartal <> '22004' GROUP BY quartal, pat_id) i " & _
   "GROUP BY quartal ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) DESC) i " & _
   "LEFT JOIN `fallzahlstand` fz ON i.quartal = fz.quartal AND fz.tage = " & abstand & " WHERE i.kassenpat<>fz.kassenpat OR ISNULL(fz.kassenpat) OR fz.arbt=0"
  Set csql0 = SqlU(sql, ((LVobMySQL)))
  myFrag vrs, csql0.Value
  If Not vrs.BOF Then
   Do While Not vrs.EOF
    If SafeArrayGetDim(rsse) = 0 Then ReDim rsse(0) Else ReDim Preserve rsse(UBound(rsse) + 1)
    rsse(UBound(rsse)).kassenpat = IIf(IsNull(vrs!kassenpat), 0, vrs!kassenpat)
    rsse(UBound(rsse)).Quartal = vrs!Quartal
    rsse(UBound(rsse)).tage = vrs!tage
    vrs.MoveNext
   Loop
   Dim ilauf&
   For ilauf = 0 To UBound(rsse)
    QT = rsse(ilauf).Quartal
    frm.Ausgeb "Berechne den Fallzahlstand f¸r Quartal: " & QT & IIf(rsse(ilauf).tage = 9999, " gesamt ...", " f¸r " & rsse(ilauf).tage & " Tage nach Quartalsbeginn ... "), True
'    Debug.Print QT, rsse(ilauf).tage, rsse(ilauf).kassenpat
    If ilauf = 0 Or myEFrag("SELECT 0 FROM fallzahlstand WHERE quartal = " & QT & " AND tage = " & abstand).BOF Then
'    sql = "DELETE FROM `fallzahlstand` WHERE quartal = " & Qt & " AND tage = " & abstand
'    myEFrag sql, rAF
    Dim qbg$, qed$, vqu$, vqbg$, vqed$
    qbg = Format(QAnf(QT), "yyyymmdd")
    qed = Format(IIf(abstand = 9999, DateAdd("m", 3, QAnf(QT)), QAnf(QT) + abstand + 1) - 1 / 60 / 24 / 24, "'yyyy-mm-dd HH:MM:SS'")
    vqu = (((Left$(QT, 1) + 2) Mod 4) + 1) & (Mid$(QT, 2) - IIf(Left$(QT, 1) = 1, 1, 0))
    vqbg = Format(QAnf(vqu), "yyyymmdd")
    vqed = Format(IIf(abstand = 9999, DateAdd("m", 3, QAnf(vqu)), QAnf(vqu) + abstand + 1) - 1 / 60 / 24 / 24, "'yyyy-mm-dd HH:MM:SS'")
    sql = "INSERT INTO `fallzahlstand`(Quartal,Tage,Kassenpat,KassenpatRel,KassenpatSchade,KassenpatSchadeRel,KassenpatKothny,KassenpatKothnyRel,DM,DmRel,GDM,Neue,NeueDM,NeueDmRel,NeueSchade,NeueKothny,Doppler,Duplex,Sonos,Schul,aktzeit,arbt,arbtKothny,arbtSchade,BriefeSchade,BriefeKothny,dmBriefeSchade,dmBriefeKothny) " & vbCrLf & _
     "SELECT i.quartal," & abstand & " Tage,kassenpat " & vbCrLf & _
     ",IF(Kassenpatvor=0,'0',ROUND(ROUND(kassenpat / Kassenpatvor,3)*100-100,1)) KassenpatRel" & vbCrLf & _
     ",davonSchade KassenpatSchade" & vbCrLf & _
     ",IF(Kassenpat=0,'0',ROUND(ROUND(davonSchade / Kassenpat,3)*100,1)) KassenpatSchadeRel" & vbCrLf & _
     ",davonKothny KassenpatKothny" & vbCrLf & _
     ",IF(Kassenpat=0,'0',ROUND(ROUND(davonKothny / Kassenpat,3)*100,1)) KassenpatKothnyRel" & vbCrLf & _
     ",DM" & vbCrLf & _
     ",IF(Kassenpat=0,'0',ROUND(ROUND(DM / Kassenpat,3)*100,1)) DMRel" & vbCrLf & _
     ",GDM,neue,neuedm" & vbCrLf & _
     ",IF(neue=0,'0',ROUND(ROUND(neuedm / neue,3)*100,1)) NeueDMRel" & vbCrLf & _
     ",neueSchade,neueKothny,Doppler,Duplex,Sonos,Schul,now(),arbt,arbtKothny,arbtSchade,BriefeSchade,BriefeKothny,dmBriefeSchade,dmBriefeKothny " & vbCrLf
     '  AND COALESCE(d.f6010,0)=0
    sql = sql & "FROM (" & vbCrLf & _
    "SELECT " & QT & " quartal," & abstand & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & "),0) Kassenpat" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle WHERE quartal=" & vqu & " AND schgr<>'90' AND fanf BETWEEN " & vqbg & " AND " & vqed & "),0) Kassenpatvor" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.pat_id = f.pat_id AND d.icd RLIKE '^E1[0-4]\.' AND d.diagsicherheit IN ('G',' ') AND obdauer<>0) GROUP BY quartal),0) DM" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND (art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') AND NOT inhalt LIKE '%(wd)%' AND NOT inhalt LIKE '%(ah)%') GROUP BY quartal),0) davonSchade" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND (art IN ('tk') OR inhalt LIKE '%(tk)%')) GROUP BY quartal),0) davonKothny" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM diagnosen WHERE icd='O24.4' AND diagdatum BETWEEN " & qbg & " AND " & qed & " AND diagsicherheit IN ('G',' ') AND f6010=0),0) GDM" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND (SELECT MIN(fanf) FROM faelle WHERE pat_id=f.pat_id)=f.fanf) GROUP BY quartal),0) neue" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND (SELECT MIN(fanf) FROM faelle WHERE pat_id=f.pat_id)=f.fanf) AND EXISTS (SELECT 0 FROM diagview d WHERE d.pat_id = f.pat_id AND d.gICD RLIKE '^E1[0-4]\.' AND obdauer<>0) GROUP BY quartal),0) neueDM" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND ((art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') AND NOT inhalt LIKE '%(wd)%' AND NOT inhalt LIKE '%(ah)%') AND (SELECT MIN(fanf) FROM faelle WHERE pat_id=f.pat_id)=f.fanf) GROUP BY quartal),0) neueSchade" & vbCrLf & _
    ",COALESCE((SELECT COUNT(DISTINCT pat_id) FROM faelle f WHERE quartal=" & QT & " AND schgr<>'90' AND fanf BETWEEN " & qbg & " AND " & qed & " AND EXISTS (SELECT 0 FROM eintraege WHERE pat_id= f.pat_id AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " AND (art IN ('tk') OR inhalt LIKE '%(tk)%') AND (SELECT MIN(fanf) FROM faelle WHERE pat_id=f.pat_id)=f.fanf) GROUP BY quartal),0) neueKothny" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'doppler' AND zeitpunkt BETWEEN " & qbg & " AND " & qed & "),0) Doppler" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'duplex' AND zeitpunkt BETWEEN " & qbg & " AND " & qed & "),0) Duplex" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'sono' AND zeitpunkt BETWEEN " & qbg & " AND " & qed & "),0) Sonos" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'schul' AND zeitpunkt BETWEEN " & qbg & " AND " & qed & "),0) Schul" & vbCrLf
    sql = sql & _
    ",COALESCE((SELECT COUNT(0) FROM (SELECT COUNT(0) zahl FROM eintraege WHERE art IN (" & artSpezBerat & "," & artSpezMA & ") AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " GROUP BY DATE(zeitpunkt)) i WHERE zahl>10),0) arbt" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM (SELECT COUNT(0) zahl FROM eintraege WHERE art IN ('tk') AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " GROUP BY DATE(zeitpunkt)) i WHERE zahl>5),0) arbtKothny" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM (SELECT COUNT(0) zahl FROM eintraege WHERE art IN ('gs') AND zeitpunkt BETWEEN " & qbg & " AND " & qed & " GROUP BY DATE(zeitpunkt)) i WHERE zahl>5),0) arbtSchade" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM briefe WHERE autor='gs' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND quelldatum BETWEEN " & qbg & " AND " & qed & "),0) BriefeSchade" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM briefe WHERE autor='tk' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND quelldatum BETWEEN " & qbg & " AND " & qed & "),0) BriefeKothny" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM briefe b INNER JOIN diagview d ON d.pat_id=b.pat_id AND d.gicd RLIKE '^E1[0-4]\.' AND obdauer<>0 WHERE autor='gs' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND quelldatum BETWEEN " & qbg & " AND " & qed & "),0) dmBriefeSchade" & vbCrLf & _
    ",COALESCE((SELECT COUNT(0) FROM briefe b INNER JOIN diagview d ON d.pat_id=b.pat_id AND d.gicd RLIKE '^E1[0-4]\.' AND obdauer<>0 WHERE autor='tk' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND quelldatum BETWEEN " & qbg & " AND " & qed & "),0) dmBriefeKothny" & vbCrLf & _
    ") i"
'   sql = sql & _
    " FROM (SELECT qu" & vbCrLf & _
    ",CONCAT(IF(LEFT(qu,1)=1,4,LEFT(qu,1)-1),IF(LEFT(qu,1)+3=1,MID(qu,2,4)-1,MID(qu,2,4))) vqu" & vbCrLf & _
    ",CONCAT(MID(qu,2),'-',(LEFT(qu,1)-1)*3+1,'-01') qbeg" & vbCrLf & _
    ",CONCAT(MID(qu,2),'-',(LEFT(qu,1)-1)*3+1,'-01') + INTERVAL " & abstand & " DAY + INTERVAL 1 DAY - INTERVAL 1 SECOND qend" & vbCrLf & _
    ",CONCAT(MID(qu,2)-if(LEFT(qu,1)='1',1,0),'-',(((LEFT(qu,1)+2)MOD 4))*3+1,'-01') vqbeg" & vbCrLf & _
    ",CONCAT(MID(qu,2)-if(LEFT(qu,1)='1',1,0),'-',(((LEFT(qu,1)+2)MOD 4))*3+1,'-01') + INTERVAL " & abstand & " DAY + INTERVAL 1 DAY - INTERVAL 1 SECOND vqed" & vbCrLf & _
    "FROM (SELECT " & Qt & " qu) i) i) i" & vbCrLf & _
    ";"

 '    sql = sql & vbCrLf & _
     "FROM (SELECT Quartal, COUNT(0) kassenpat,  " & vbCrLf & _
     "(SELECT COUNT(DISTINCT pat_id) FROM `faelle` WHERE schgr <> '90' AND quartal = vorquartal " & VQBed & " GROUP BY quartal) Kassenpatvor,  " & vbCrLf & _
     "(SELECT COUNT(DISTINCT pat_id) FROM `faelle` f WHERE schgr <> '90' AND quartal = i.quartal " & QBedi & " AND EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.pat_id = f.pat_id AND d.icd RLIKE '^E1[0-4]\.|^O24\.' AND d.diagsicherheit RLIKE '[^Z^A^V]' AND COALESCE(d.f6010,0)=0) GROUP BY quartal) Diabetes, " & vbCrLf & _
     "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('" & IIf(QAnf(Quartal) > #1/1/2006#, "gs", "gs','notiz") & "','doppler','duplex') OR inhalt LIKE '%(gs)%') " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) davonSchade, " & vbCrLf & _
     "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art = 'tk' OR inhalt LIKE '%(tk)%') " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) davonKothny, " & vbCrLf & _
     "(SELECT COUNT(DISTINCT d.pat_id) Zahl FROM `diagnosen` d LEFT JOIN `faelle` f ON d.fid = f.fid WHERE d.icd = 'O24.4' AND d.diagsicherheit RLIKE '[^V^Z^A]' AND COALESCE(d.f6010,0)=0 AND f.quartal = i.quartal) GDM, " & vbCrLf & _
     "SUM(erst) neue, SUM(erstdm) neuedm, " & vbCrLf & _
     "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) neueSchade, " & vbCrLf & _
     "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('tk') OR inhalt LIKE '%(tk)%') " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) neueKothny, " & vbCrLf & _
     "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'doppler' " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Doppler, " & vbCrLf & _
     "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'duplex' " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Duplex, " & vbCrLf & _
     "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'sono' " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Sonos, " & vbCrLf & _
     "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'schul' " & zpQBed & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Schul " & vbCrLf & _
     "FROM (SELECT COUNT(0) zi, quartal, " & vbCrLf & _
     "CONCAT(IF(LEFT(quartal,1)=1,4,LEFT(quartal,1)-1),IF(LEFT(quartal,1)=1,MID(quartal,2,4)-1,MID(quartal,2,4))) Vorquartal, " & vbCrLf & _
     "(SELECT COUNT(0)=0 FROM `faelle` f1 WHERE f1.pat_id = f.pat_id AND f1.fanf < MIN(f.fanf)) erst, " & vbCrLf & _
     "(SELECT COUNT(0)<>0 FROM `faelle` f1 WHERE f1.fid = f.fid AND f1.fanf = (SELECT MIN(fanf) FROM `faelle` f2 WHERE f2.pat_id = f1.pat_id) AND EXISTS (SELECT 0 FROM `diagnosen` d WHERE d.pat_id = f1.pat_id AND d.icd RLIKE '^E1[0-4]\.|^O24\.' AND d.diagsicherheit RLIKE '[^Z^A^V]' AND COALESCE(d.f6010,0)=0)) erstdm " & vbCrLf & _
     "FROM `faelle` f " & vbCrLf & _
     "WHERE quartal = " & Qt & " AND schgr <> '90' " & QBed & " GROUP BY quartal, pat_id) i " & vbCrLf & _
     "GROUP BY quartal) i"
     ' aus vorletzter Zeile vor group: AND substr(quartal,2,4)> '2008' AND quartal <> '12009'
     ' aus letzter Zeile vor Klammer: ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) desc
    Set csql0 = SqlU(sql, ((LVobMySQL)))
    myEFrag csql0.Value, rAf
' das folgende wurde mit Hilfe des bash-Programms briefautor.sh unnˆtig:
#If False Then
  Dim RsI As New ADODB.Recordset
'    sql = "UPDATE `fallzahlstand` SET arbt = (SELECT COUNT(0) arbt FROM (SELECT DATE(zeitpunkt) tag, COUNT(0) zahl FROM `eintraege` e WHERE zeitpunkt BETWEEN STR_TO_DATE(CONCAT(MID(" & Qt & ",2),(LEFT(" & Qt & ",1)-1)*3+1,'/1'),'%Y%c/%d') AND adddate(STR_TO_DATE(CONCAT(MID(" & Qt & ",2),(LEFT(" & Qt & ",1)-1)*3+1,'/1'),'%Y%c/%d')," & abstand & ") AND art IN ('tk','gs','notiz') GROUP BY DATE(zeitpunkt)) i WHERE zahl > 5) WHERE quartal = " & Qt & " AND tage = " & abstand
'    sql = "SELECT eart, COUNT(0) zahl FROM (SELECT * FROM (SELECT name,pat_id,DATE(zeitpunkt) zp,(SELECT MAX(art) FROM `eintraege` e WHERE e.pat_id = `briefe`.pat_id AND art IN ('tk','gs') AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `eintraege` e WHERE e.pat_id = `briefe`.pat_id AND art IN ('tk','gs') AND zeitpunkt < `briefe`.zeitpunkt)) eart FROM `briefe` WHERE not name LIKE '%dmp-%' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND art IN ('wbr','brief')) i WHERE NOT ISNULL(eart) GROUP BY pat_id, eart, zp) i WHERE zp BETWEEN '" & Format(QAnf(quartal), "yyyy-mm-dd") & "' AND '" & IIf(abstand = 9999, Format(qend(quartal), "yyyy-mm-dd"), Format(QAnf(quartal) + abstand, "yyyy-mm-dd")) & "' GROUP BY eart"
    
    ' ktag fehlerhaft
    csql0 = "UPDATE `fallzahlstand` SET "
    '" arbt = (SELECT COUNT(0) arbt FROM (SELECT DATE(zeitpunkt), COUNT(0) zahl FROM `eintraege` WHERE zeitpunkt BETWEEN " & DatFor_k(qanf(Qt)) & " AND " & IIf(abstand = 9999, DatFor_k(qend(Qt) + 1), DatFor_k(qanf(Qt) + abstand + 1)) & " AND art IN ('tk','gs','notiz') GROUP BY DATE(zeitpunkt)) i WHERE zahl > 5), " & _
                                     "arbtkothny = (SELECT COUNT(0) arbt FROM (SELECT DATE(zeitpunkt), COUNT(0) zahl FROM `eintraege` WHERE zeitpunkt BETWEEN " & DatFor_k(qanf(Qt)) & " AND " & IIf(abstand = 9999, DatFor_k(qend(Qt) + 1), DatFor_k(qanf(Qt) + abstand + 1)) & " AND art IN ('tk') GROUP BY DATE(zeitpunkt)) i WHERE zahl > 5), " & _
                                     "arbtschade = (SELECT COUNT(0) arbt FROM (SELECT DATE(zeitpunkt), COUNT(0) zahl FROM `eintraege` WHERE zeitpunkt BETWEEN " & DatFor_k(qanf(Qt)) & " AND " & IIf(abstand = 9999, DatFor_k(qend(Qt) + 1), DatFor_k(qanf(Qt) + abstand + 1)) & " AND art IN ('" & IIf(qanf(Qt) > #1/1/2006#, "gs", "gs','notiz") & "') GROUP BY DATE(zeitpunkt)) i WHERE zahl > 5) "
' dann unten weiter ' 19.8.21: quelldatum statt zeitpunkt, wegen Index
'  AND COALESCE(d.f6010,0)=0
    sql = "SELECT eart, dm, COUNT(0) zahl FROM (" & _
       "SELECT name,pat_id,zp,eart,dm FROM (" & _
            "SELECT name,b.pat_id,b.quelldatum zp," & _
            "(SELECT MAX(art) FROM `eintraege` e " & _
             "WHERE e.pat_id = b.pat_id AND art IN ('tk','gs') AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `eintraege` e WHERE e.pat_id = b.pat_id AND art IN ('tk','gs') AND zeitpunkt < b.zeitpunkt) " & _
            ") eart," & _
         "NOT ISNULL(icd) dm " & _
       "FROM `briefe` b " & _
       "LEFT JOIN diagview d ON b.pat_id = d.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.f6010=0 AND d.diagsicherheit IN ('G',' ')))" & _
       "WHERE not name LIKE '%dmp-%' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND art IN ('wbr','brief') " & _
                ") i WHERE NOT ISNULL(eart) GROUP BY pat_id, eart, zp, dm " & _
                ") i WHERE zp BETWEEN " & DatFor_k(QAnf(QT)) & " AND " & IIf(abstand = 9999, DatFor_k(QEnd(QT) + 1), DatFor_k(QAnf(QT) + abstand + 1)) & " GROUP BY eart,dm "
    Dim obCnOffen%
    If Not DBCn Is Nothing Then If DBCn.State = 1 Then obCnOffen = -1
    If obCnOffen Then DBCn.Close
    SetDBCn Nothing, vNS
    myFrag RsI, sql
    
    If Not RsI.BOF Then
     Do While Not RsI.EOF
      Select Case RsI!eArt
       Case "gs": If RsI!dm = 0 Then s0 = RsI!Zahl Else s1 = RsI!Zahl
       Case "tk": If RsI!dm = 0 Then k0 = RsI!Zahl Else k1 = RsI!Zahl
      End Select
      RsI.MoveNext
     Loop
     csql0.Append " BriefeSchade = " & s1 + s0 & ",BriefeKothny=" & k1 + k0 & ",dmBriefeSchade=" & s1 & ",dmBriefeKothny=" & k1
     csql0.Append " WHERE quartal = '" & QT & "' AND tage = " & abstand
     myEFrag csql0.Value, rAf
    End If ' Not RsI.BOF THEN
    
    Set RsI = Nothing
#End If
    DoEvents
    End If ' ilauf = 0 OR myEFrag(...).bof
   Next ilauf ' While Not vrs.EOF
  End If ' Not vrs.BOF THEN
  Screen.MousePointer = vbNormal
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fzsfuell/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' fzsfuell(Optional obgestern)

' auskommentiert 6.8.11
'Function nachfuell()
' Dim rs As New ADODB.Recordset, RsI As New ADODB.Recordset, sql$, abstand&, rAF&
' Dim s0&, s1&, k0&, k1&
' Lese.ProgStart
' myFrag, vrs, "SELECT quartal,tage FROM `fallzahlstand`"
' IF Not vrs.BOF THEN
'  Do While Not vrs.EOF
'    abstand = vrs!tage
'    s0 = 0: s1 = 0: k0 = 0: k1 = 0
'#If False THEN
'    sql = "UPDATE `fallzahlstand` SET arbt = (SELECT COUNT(0) arbt FROM (SELECT DATE(zeitpunkt) tag, COUNT(0) zahl FROM `eintraege` e WHERE zeitpunkt BETWEEN '" & Format(QAnf(quartal), "yyyy-mm-dd") & "' AND '" & IIf(abstand = 9999, Format(qend(quartal), "yyyy-mm-dd"), Format(QAnf(quartal) + abstand, "yyyy-mm-dd")) & "' AND art IN ('tk','gs','notiz') GROUP BY DATE(zeitpunkt)) i WHERE zahl > 5) WHERE quartal = '" & Qt & "' AND tage = " & abstand
'    myEFrag sql, rAF
'#END IF
''    sql = "SELECT eart, COUNT(0) zahl FROM (SELECT * FROM (SELECT name,pat_id,DATE(zeitpunkt) zp,(SELECT MAX(art) FROM `eintraege` e WHERE e.pat_id = briefe.pat_id AND art IN ('tk','gs') AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `eintraege` e WHERE e.pat_id = `briefe`.pat_id AND art IN ('tk','gs') AND zeitpunkt < `briefe`.zeitpunkt)) eart FROM `briefe` WHERE not name LIKE '%dmp-%' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND art IN ('wbr','brief')) i WHERE NOT ISNULL(eart) GROUP BY pat_id, eart, zp) i WHERE zp BETWEEN '" & Format(QAnf(quartal), "yyyy-mm-dd") & "' AND '" & IIf(abstand = 9999, Format(qend(quartal), "yyyy-mm-dd"), Format(QAnf(quartal) + abstand, "yyyy-mm-dd")) & "' GROUP BY eart"
'    sql = "SELECT eart, dm, COUNT(0) zahl FROM (" & _
'       "SELECT name,pat_id,zp,eart,dm FROM (" & _
'            "SELECT name,b.pat_id,DATE(b.zeitpunkt) zp," & _
'            "(SELECT MAX(art) FROM `eintraege` e " & _
'             "WHERE e.pat_id = b.pat_id AND art IN ('tk','gs') AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `eintraege` e WHERE e.pat_id = b.pat_id AND art IN ('tk','gs') AND zeitpunkt < b.zeitpunkt) " & _
'            ") eart," & _
'         "NOT ISNULL(icd) dm " & _
' "FROM `briefe` b " & _
' "LEFT JOIN `diagnosen` d ON b.pat_id = d.pat_id AND d.icd RLIKE '^E1[0-4]\.|^O24\.' AND d.diagsicherheit IN ('G',' ') AND COALESCE(d.f6010,0)=0 " & _
' "WHERE not name LIKE '%dmp-%' AND (name LIKE '%brief%' OR name LIKE '%nachricht%') AND art IN ('wbr','brief') " & _
'                ") i WHERE NOT ISNULL(eart) GROUP BY pat_id, eart, zp, dm " & _
'                ") i WHERE zp BETWEEN '" & Format(QAnf(quartal), "yyyy-mm-dd") & "' AND '" & IIf(abstand = 9999, Format(qend(quartal), "yyyy-mm-dd"), Format(QAnf(quartal) + abstand, "yyyy-mm-dd")) & "' GROUP BY eart,dm "
'
'    myFrag RsI, sql
'    IF Not RsI.BOF THEN
'     Do While Not RsI.EOF
'      SELECT CASE RsI!eart
'       Case "gs": IF RsI!dm = 0 THEN s0 = RsI!Zahl ELSE s1 = RsI!Zahl
'       Case "tk": IF RsI!dm = 0 THEN k0 = RsI!Zahl ELSE k1 = RsI!Zahl
'      END SELECT
'      RsI.MoveNext
'     Loop
'     myEFrag "UPDATE `fallzahlstand` SET BriefeSchade = " & s1 + s0 & ",BriefeKothny=" & k1 + k0 & ",dmBriefeSchade=" & s1 & ",dmBriefeKothny=" & k1 & " WHERE quartal = '" & Qt & "' AND tage = " & abstand, rAF
'    END IF
'    SET RsI = Nothing
'    vrs.MoveNext
'  Loop
' END IF
'End FUNCTION ' nachfuell

Function dofallzahlstand(frm As Lese, Optional obgestern$) ' in GesLies und Fallzahlstand_Click
 Dim sql$, rs As New ADODB.Recordset, runde%, FNr&, rAf&
 Dim abstand&, heute As Date
 Dim csql As CString
 myFrag rs, "SELECT DATE(NOW()) jetzt"
 myFrag rs, "SELECT DATE(NOW()) jetzt"
 heute = rs!jetzt
 Set rs = Nothing
 If IsNumeric(obgestern) Then
  sql = "SELECT to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) abstand"
  myFrag rs, sql
  abstand = rs!abstand
  Set rs = Nothing
 Else
  abstand = 9999
  obgestern = 0
 End If ' IsNumeric(obgestern) Then
 Call fzsfuell(frm, abstand, obgestern)
 sql = "SELECT CONCAT(LEFT(Quartal,1),'/',RIGHT(quartal,2)) `Quart`, ArbT, CONCAT(RIGHT(CONCAT('    ',CAST(Kassenpat AS char)),5),' ', RIGHT(CONCAT('     ',CAST(KassenPatRel AS char)),5),'% ',CAST(round(Kassenpat/Arbt,1) AS char)) `Kassenpat., dav.:`, CONCAT(RIGHT(CONCAT('   ',CAST(ArbTSchade AS char)),3),' ', RIGHT(CONCAT('   ',CAST(KassenpatSchade AS char)),4),' ', RIGHT(CONCAT('     ',CAST(KassenpatSchadeRel AS char)),5),'% ',IF(ArbtSchade=0,'',CAST(round(KassenpatSchade/ArbtSchade,1) AS char))) `Schade,`, CONCAT(RIGHT(CONCAT('   ',CAST(ArbTKothny AS char)),3),' ', RIGHT(CONCAT('   ',CAST(KassenpatKothny AS char)),4),' ', RIGHT(CONCAT('     ',FORMAT(KassenPatKothnyRel,1)),5),'% ',IF(ArbtKothny=0,'',CAST(round(KassenpatKothny/ArbtKothny,1) AS char))) `Kothny,`, " & _
       "CONCAT(Dm,' (',RIGHT(CONCAT('     ',FORMAT(DmRel,1)),5),'%)') `Dm;`, GDM `GDM,`, CAST(RIGHT(CONCAT('   ',FORMAT(Neue,0)),3) AS SIGNED) `Neue,`, NeueDM `dav: Dm,`, CAST(RIGHT(CONCAT('   ',FORMAT(NeueSchade,0)),3) AS SIGNED) `S,`, CAST(RIGHT(CONCAT('   ',FORMAT(NeueKothny,0)),3) AS SIGNED) `K;`, Doppler Dop, Duplex Dup, Sonos Son, Schul `Schul;`, BriefeSchade `ABri: S,`, BriefeKothny `K,`,dmBriefeSchade `dav.D.m.:S,`,dmBriefeKothny `K:` FROM `fallzahlstand` f WHERE Tage = " & abstand & " ORDER BY MID(quartal, 2) DESC, quartal DESC"
       ' RIGHT(CONCAT('     ',format(NeueDmRel,0),'%'),5)
 Set rs = Nothing
 Set csql = SqlU(sql, ((LVobMySQL)))
 myFrag rs, csql.Value
 TabAusgeb rs, frm, True
 frm.Ausgeb "Fallzahlstand " & IIf(abstand = 9999, "ganzes Quartal", abstand & " Tage nach Quartalsbeginn") & ", errechnet f¸r " & Format(heute, "dd/mm/yy") & IIf(abstand = 9999, ":", " minus " & obgestern & " Tage:"), 1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dofallzahlstand/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' fallzahlstand

' For runde = 3 To 3
''   sql = "SELECT Quartal, COUNT(0) - COUNT(erst) AS `neue Pat.`, COUNT(0) AS `F‰lle` FROM `faelleverschiedenneu` f " & _
  "WHERE schgr <> '90' AND to_days(fanf)-to_days(cdate(CONCAT(MID(quartal,2,4)°'-'°(LEFT(quartal,1)-1)*3+1°'-01'))) " & _
  "BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days(cdate(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))°'-'°intacc((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)divmy 3)*3+1°'-01'))) AND " & IIf(runde = 1, "(pat_id < 3044 OR pat_id > 50000)", IIf(runde = 2, "pat_id > 3044", "true")) & _
  " AND substr(quartal,2,4)> '2008' AND quartal <> '12009' " & _
  " GROUP BY quartal ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) DESC"
' Wenn schon Eintrag in der Tabelle `fallzahlstand`
'  myFrag rs, "SELECT COUNT(0) Fehlende, Quartal FROM (SELECT f.quartal, f1.tage FROM `fallzahlstand` f LEFT JOIN `fallzahlstand` f1 ON f1.quartal = f.quartal AND f1.tage = " & abstand & " WHERE ISNULL(f1.tage) GROUP BY f.quartal) i"
'  IF Not rs.BOF THEN
'   IF rs!fehlende <> 0 THEN
'    Screen.MousePointer = vbHourglass
'    sql = "SELECT i.quartal, CONCAT(CAST(kassenpat AS char),' (',CONCAT(IF(kassenpat>kassenpatvor,'+',''),CAST(round(round(kassenpat / Kassenpatvor,3)*100-100,1) AS char),'%'),')') `Kassenpat.,`, CONCAT(CAST(davonSchade AS char),' (',CONCAT(CAST(round(round(davonSchade / Kassenpat,3)*100,1) AS char),'%'),')') `davon Schade,`, CONCAT(CAST(davonKothny AS char),' (',CONCAT(CAST(round(round(davonKothny / Kassenpat,3)*100,1) AS char),'%'),')') `Kothny,`, CONCAT(CAST(round(round(Diabetes / kassenpat,3)*100,1) AS char),'%') `D.m.,`, GDM `GDM,`, neue `Neue,`, '' `dav.: `, CONCAT(CAST(round(round(neuedm / neue,3)*100,1) AS char),'%') `D.m.,`, neueSchade `Schade,`, neueKothny `Kothny;`, Doppler `Kasse+Privat: Doppler`, Duplex, Sonos, Schul " & _
          "FROM (SELECT Quartal, COUNT(0) kassenpat,  " & _
          "(SELECT COUNT(DISTINCT pat_id) FROM `faelle` WHERE schgr <> '90' AND quartal = vorquartal AND to_days(fanf)-to_days((CONCAT(MID(vorquartal,2,4),'-',(LEFT(vorquartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand & " GROUP BY quartal) Kassenpatvor,  " & _
          "(SELECT COUNT(DISTINCT pat_id) FROM `faelle` f WHERE schgr <> '90' AND quartal = i.quartal AND to_days(fanf)-to_days((CONCAT(MID(i.quartal,2,4),'-',(LEFT(i.quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand & " AND exists (SELECT 0 FROM `diagnosen` d WHERE pat_id = f.pat_id AND d.icd RLIKE '^E1[0-4]\.|^O24\.' AND diagsicherheit RLIKE '[^Z^A^V]' AND COALESCE(f6010,0)=0) GROUP BY quartal) Diabetes, " & _
          "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') AND to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) davonSchade, " & _
          "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art = 'tk' OR inhalt LIKE '%(tk)%') AND to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) davonKothny, " & _
          "(SELECT COUNT(DISTINCT d.pat_id) Zahl FROM `diagnosen` d LEFT JOIN `faelle` f ON d.fid = f.fid WHERE d.icd = 'O24.4' AND d.diagsicherheit RLIKE '[^V^Z^A]' AND COALESCE(d.f6010,0)=0 AND f.quartal = i.quartal) GDM, " & _
          "SUM(erst) neue, SUM(erstdm) neuedm, " & _
          "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') AND to_days(e.zeitpunkt)- to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) neueSchade, " & _
          "(SELECT COUNT(DISTINCT e.pat_id) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.pat_id = f.pat_id WHERE schgr <> '90' AND (art IN ('tk') OR inhalt LIKE '%(tk)%') AND to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) neueKothny, " & _
          "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'doppler' AND to_days(e.zeitpunkt)- to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Doppler, " & _
          "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'duplex' AND to_days(e.zeitpunkt)- to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Duplex, " & _
          "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'sono' AND to_days(e.zeitpunkt)- to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Sonos, " & _
          "(SELECT COUNT(0) FROM `eintraege` e WHERE e.art = 'schul' AND to_days(e.zeitpunkt)- to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " AND CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) Schul " & _
          "FROM (SELECT COUNT(0) zi, quartal, " & _
          "CONCAT(IF(LEFT(quartal,1)=1,4,LEFT(quartal,1)-1),IF(LEFT(quartal,1)=1,MID(quartal,2,4)-1,MID(quartal,2,4))) Vorquartal, " & _
          "(SELECT COUNT(0)=0 FROM `faelle` f1 WHERE f1.pat_id = f.pat_id AND f1.fanf < MIN(f.fanf)) erst, " & _
          "(SELECT COUNT(0)<>0 FROM `faelle` f1 WHERE f1.fid = f.fid AND f1.fanf = (SELECT MIN(fanf) FROM `faelle` f2 WHERE f2.pat_id = f1.pat_id) AND exists (SELECT 0 FROM `diagnosen` d WHERE pat_id = f1.pat_id AND d.icd RLIKE '^E1[0-4]\.|^O24\.' AND d.diagsicherheit RLIKE '[^Z^A^V]' AND COALESCE(d.f6010,0)=0 )) erstdm " & _
          "FROM `faelle` f WHERE schgr <> '90' AND to_days(fanf)-to_days((CONCAT(MID(quartal,2,4),'-',(LEFT(quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand & " AND substr(quartal,2,4)> '2008' AND quartal <> '12009'  GROUP BY quartal, pat_id) i GROUP BY quartal ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) DESC) i"

'' sql = _
 "SELECT Quartal, COUNT(0) `Kassenpat.,`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('gs','doppler','duplex') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) `davon Schade,`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art = 'tk' " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) `Kothny;`, " & _
 "SUM(erst) `neue Kassenpat.,`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('gs','doppler','duplex') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) `davon Schade,`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('tk') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND " & abstand & " " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) `Kothny` " & _
"FROM " & _
"(SELECT COUNT(0) zi, quartal, (SELECT COUNT(0)=0 FROM `faelle` f1 WHERE f1.pat_id = f.pat_id AND f1.fanf < MIN(f.fanf)) erst " & _
"FROM `faelle` f WHERE schgr <> '90' AND to_days(fanf)-to_days((CONCAT(MID(quartal,2,4),'-',(LEFT(quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND " & abstand & " AND substr(quartal,2,4)> '2008' AND quartal <> '12009'  GROUP BY quartal, pat_id) i GROUP BY quartal ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) DESC"
  
''sql = "SELECT Quartal, COUNT(0) `Kassenpat.`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('gs','doppler','duplex') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) `davon Schade`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art = 'tk' " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal) `Kothny`, " & _
 "SUM(erst) `neue Kassenpat.`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('gs','doppler','duplex') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) `davon Schade`, " & _
 "(SELECT COUNT(distinct(e.pat_id)) Zahl FROM `eintraege` e LEFT JOIN `faelle` f ON e.fid = f.fid WHERE schgr <> '90' AND art IN ('tk') " & _
 "and to_days(e.zeitpunkt)-to_days((CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01'))) BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) " & _
 "and CONCAT((month(e.zeitpunkt)+2) div 3, YEAR(e.zeitpunkt)) = i.Quartal AND (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) = (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id AND fanf >= (CONCAT(YEAR(e.zeitpunkt),'-',month(e.zeitpunkt) - ((month(e.zeitpunkt)-1) mod 3),'-01')) )) `Kothny` " & _
"FROM " & _
"(SELECT COUNT(0) zi, quartal, (SELECT COUNT(0)=0 FROM `faelle` f1 WHERE f1.pat_id = f.pat_id AND f1.fanf < MIN(f.fanf)) erst " & _
"FROM `faelle` f WHERE schgr <> '90' AND to_days(fanf)-to_days((CONCAT(MID(quartal,2,4),'-',(LEFT(quartal,1)-1)*3+1,'-01'))) BETWEEN 0 AND to_days(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-to_days((CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY)),'-',((month(SUBDATE(NOW(),INTERVAL " & obgestern & " DAY))-1)div 3)*3+1,'-01'))) AND substr(quartal,2,4)> '2008' AND quartal <> '12009'  GROUP BY quartal, pat_id) i GROUP BY quartal ORDER BY substr(Quartal,2,4) DESC,LEFT(Quartal,1) DESC "
' "and (SELECT MIN(fanf) FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) fanfmin = (SELECT fanf FROM `faelle` f1 WHERE f1.pat_id = e.pat_id) fanfakt ) `Schade neu` " & _

 '  END IF
'  END IF ' not rs.bof
'  SET rs = Nothing
'  Call DtbCreateQueryDef("fallzahlstand " & runde, sql)
'  SET rs = Nothing
'  myFrag rs, "SELECT * FROM `fallzahlstand " & runde & "`"
'  Do While Not rs.EOF
'   frm.Ausgabe = Right$(Space$(4) & rs!Zahl, 4) & " " & Right$(Space$(3) & rs!z0, 3) & " " & rs!Quartal + vbCrLf + altAusgabe
'   altAusgabe = frm.Ausgabe
'   rs.Move 1
'  Loop
'  frm.Ausgeb "F‰lle verschiedener Patienten/ neue Patienten in Praxis " & IIf(runde = 1, "Schade", IIf(runde = 2, "Dr. Kothny", "gesamt")) & ":", True
' Next runde
' Screen.MousePointer = vbNormal


'Function doPortoalt(frm AS Lese)
' Dim z$, op$, sql$, rs As New ADODB.Recordset, rAF&, Zp As Date
' ON Error GoTo fehler
' sql = "SELECT time(zp) AS uzeit, i2.* FROM (SELECT IF(minleizp > zeitpunkt, minleizp, zeitpunkt) AS zp, innen.* FROM (SELECT b.pat_id, (SELECT adddate(min(zeitpunkt),INTERVAL 1 minute) FROM `leistungen` WHERE fid = f.fid) AS minleizp, b.zeitpunkt AS zeitpunkt, f.fid, '40120' AS leistung, schgr,GOƒKatName,GOƒKatNr,‹WZiel,And‹w,‹bwV,BhFB,BhFE1,TMFNr,AbrGb,GebOr,KVKserg,KVKs,IK,AbrAr,lVorl,KtrAbrB, name FROM `briefe` b LEFT JOIN lfaellev f using(pat_id) LEFT JOIN (SELECT * FROM `leistungen` WHERE leistung LIKE '4012%') AS l ON l.pat_id = b.pat_id AND DATE(l.zeitpunkt)= DATE(b.zeitpunkt) WHERE (name LIKE '%brief%' OR name LIKE '%DMP-Daten%') AND b.zeitpunkt BETWEEN CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1)div 3)*3+1,'-01') AND adddate(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 14 DAY)),'-',((month(SUBDATE(NOW(),INTERVAL 14 DAY))-1)div 3)*3+1,'-01'), INTERVAL 3 MONTH) AND ISNULL(leistung) AND schgr <> '90') AS innen) AS i2x;"
' myFrag rs, sql
' IF Not rs.BOF THEN
' z = hVerz + "LEIST " + Format$(Now, "dd/mm/yy HH.MM") + ".BDT"
' Open z For Output AS #310
'  Print #310, "01380000020"
'  Print #310, "014810000082"
'  Print #310, "0169100" & KVNR
'  Print #310, "017910309042005"
'  Print #310, "0129105001"
'  Print #310, "01091062"
'  Print #310, "01380000022"
'  Print #310, "014810000107"
'  Print #310, "014921001/99"
'  Print #310, "014921302/94"
'  Print #310, "01096001"
'  Print #310, "025960101011980" + Format$(Now, "ddmmyyyy")
'  Print #310, "017960214290200"
'  Print #310, "01380000010"
'  Print #310, "014810000315"
'  Print #310, "0160101A001011"
'  Print #310, "0260102TurboMed EDV GmbH"
'  Print #310, "0250103TurboMed@Windows"
'  Print #310, "0180104IBM PC/AT"
'  Print #310, "0160201" & KVNR
'  Print #310, "01002021"
'  Print #310, "0220203Gerald Schade"
'  Print #310, "0500204FA Innere und Allgemeinmedizin (Hausarzt)"
'  Print #310, "0290205Mittermayerstraﬂe 13"
'  Print #310, "014021585221"
'  Print #310, "0150216Dachau"
'  Print #310, "024020808131 / 616 380"
'  Print #310, "024020908131 / 616 381"
'
' Do While Not rs.EOF
''  IF rs!Pat_id = 42 THEN
'  Print #310, "0138000" + IIf(rs!SchGr = "90", "0190", "0102") ' Satzidentifikation
' ' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
'' bei 190 entsteht ein neuer Privatfall, bei 6100 l‰uft alles ohne Fehler durch, aber keine Leistung steht drin
'' bei 6200 entsteht ein neuer Kassenfall
'  Print #310, "014810000845"
' '     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
''     Print #310, zsu(op)
''     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
''     Print #310, zsu(op)
'     op = Format$(3 + 4 + Len(CStr(rs!Pat_id)), "000") + "3000" + CStr(rs!Pat_id)
'     Print #310, ZSU(op)
'     IF NOT ISNULL(rs!KtrAbrB) THEN  ' bei Privaten
'      op = Format$(3 + 4 + Len(rs!KtrAbrB), "000") + "4106" + rs!KtrAbrB
'      Print #310, ZSU(op)
'     END IF
'     IF NOT ISNULL(rs!AbrAr) THEN  ' bei Privaten
'      op = Format$(3 + 4 + Len(rs!AbrAr), "000") + "4107" + rs!AbrAr
'      Print #310, ZSU(op)
'     END IF
'     IF rs!SchGr <> "90" THEN
'      op = Format$(3 + 4 + 8, "000") + "4109" + Format$(rs!lVorl, "ddmmyyyy")
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + 4, "000") + "4110" + Format$(rs!lVorl, "hhmm")
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + Len(rs!IK), "000") + "4111" + rs!IK
'      Print #310, ZSU(op)
'      IF NOT ISNULL(rs!KVKs) THEN ' bei Pat_id 43
'       op = Format$(3 + 4 + Len(rs!KVKs), "000") + "4112" + rs!KVKs
'       Print #310, ZSU(op)
'      END IF
'      IF NOT ISNULL(rs!KVKserg) THEN
'       op = Format$(3 + 4 + Len(rs!KVKserg), "000") + "4113" + rs!KVKserg
'       Print #310, ZSU(op)
'      END IF
'      END IF
'     IF NOT ISNULL(rs!GebOr) THEN ' bei Privaten
'      op = Format$(3 + 4 + Len(rs!GebOr), "000") + "4121" + rs!GebOr
'      Print #310, ZSU(op)
'     END IF
'     IF NOT ISNULL(rs!AbrGb) THEN ' bei Privaten
'      op = Format$(3 + 4 + Len(rs!AbrGb), "000") + "4122" + rs!AbrGb
'      Print #310, ZSU(op)
'     END IF
''     14.4.08: folgende beiden Zeilen f¸hren zur Fall-Neuanlage
''     op = format$(3 + 4 + Len("TM#" + IIf(ISNULL(rs!TMFNr), space$(11), rs!TMFNr)), "000") + "4144" + "TM#" + IIf(ISNULL(rs!TMFNr), space$(11), rs!TMFNr)
''     Print #310, ZSU(op)
'     op = Format$(3 + 4 + 8, "000") + "4150" + IIf(rs!BhFB = 0, "00000000", Format$(rs!BhFB, "ddmmyyyy"))
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + 8, "000") + "4151" + IIf(rs!BhFE1 = 0, "00000000", Format$(rs!BhFE1, "ddmmyyyy"))
'     Print #310, ZSU(op)
''     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rs!BhFE2 = 0, "00000000", format$(rs!BhFE2, "ddmmyyyy"))
''     Print #310, ZSU(op)
'     Dim ‹w$
'     ‹w = IIf(ISNULL(rs!‹bwV), vns, rs!‹bwV)
'     IF ‹w <> "" THEN
'      op = Format$(3 + 4 + Len(‹w), "000") + "4218" + ‹w
'      Print #310, ZSU(op)
'     END IF
'     ‹w = IIf(ISNULL(rs!And‹w), vns, rs!And‹w)
'     IF ‹w <> "" THEN
'      op = Format$(3 + 4 + Len(‹w), "000") + "4219" + ‹w
'      Print #310, ZSU(op)
'     END IF
'     IF NOT ISNULL(rs!‹WZiel) AND rs!‹WZiel <> "" THEN
'      op = Format$(3 + 4 + Len(rs!‹WZiel), "000") + "4220" + rs!‹WZiel
'      Print #310, ZSU(op)
'     END IF
'     op = Format$(3 + 4 + 2, "000") + "4239" + rs!SchGr
'     Print #310, ZSU(op)
'     IF (NOT ISNULL(rs!GOƒKatNr) AND rs!GOƒKatNr <> "") OR (NOT ISNULL(rs!GOƒKatName) AND rs!GOƒKatName <> "") THEN
'      ‹w = "TM#" + CStr(rs!GOƒKatNr) + "#" + CStr(rs!GOƒKatName)
'      op = Format$(3 + 4 + Len(‹w), "000") + "4580" + ‹w
'      Print #310, ZSU(op)
'     END IF
'leistungdirekt:
'     op = Format$(3 + 4 + 8, "000") + "5000" + Format$(rs!Zp, "ddmmyyyy")
'     Print #310, ZSU(op)
'     Dim UZeit$
'     IF rs!Leistung = "9995" THEN
'      UZeit = "0001"
'     ElseIf ISNULL(rs!UZeit) THEN
'      UZeit = "1900"
'     ElseIf rs!UZeit = CDate(0) THEN
'      UZeit = "1900"
'     Else
'      UZeit = Format$(rs!UZeit, "hhmm")
'     END IF
'     op = Format$(3 + 4 + 4, "000") + "6201" + UZeit
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + Len(rs!Leistung), "000") + "5001" + CStr(rs!Leistung)
'     Print #310, ZSU(op)
''     q.Edit
''     q!Status = ¸bertragen
''     q.Update
'     Call myEFrag("INSERT INTO " & "`" & "leistungen exportiert" & "`" & "(datum,pat_id,schgr,leistung,¸bertragen) VALUES(" & DatFor_k(rs!Zp) & "," & rs!Pat_id & "," & rs!SchGr & "," & rs!Leistung & ", " & DatFor_k(NOW()) & ")", rAF)
''    END IF ' pat_id =
'    rs.Move 1
' Loop
' Close #310
' zeigan z & vNS, vbnormalmodus
' END IF
'Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
' SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in doPorto/" + AnwPfad)
'  Case vbAbort: Call MsgBox("Hˆre auf"): Progende
'  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
'  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
' END SELECT
'End Function

'Function test8()
' Dim Cn As New ADODB.Connection
'' cn.Open "DRIVER={MySQL ODBC 5.1 Driver};database=office;server=" & LiName & ";uid=praxis;pwd=...;option=3"
' Cn.Open "DRIVER={" & ODBCStr & "};database=office;server=" & LiName & ";uid=praxis;pwd=...;option=3"
' Dim rs As New ADODB.Recordset
' rs.Open "SELECT * FROM test", Cn, adOpenStatic, adLockReadOnly
' Do While Not rs.EOF
'  Debug.Print rs!t
'  rs.Move 1
' Loop
'End Function ' test8

Function fallzeig()
 Dim sql$, AuS$, i%, rs As New ADODB.Recordset
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 sql = "SELECT * FROM (SELECT TOP 100 pat_id, CONCAT(nachname° ' '° vorname) AS name, intaccdatemy(fanf) AS fanf, Schgr, fid, intaccdatemy(ausgst) AS ausgst, aktzeit, quartal, lvorl, GebOr FROM `faelle` ORDER BY fanf DESC, fid DESC LIMIT 100) i ORDER BY fanf, fid"
 sql = SqlU(sql, Lese.obMySQL)
 myFrag rs, sql, , , , , , True, ErrNr, ErrDes
 If ErrNr <> 0 Then Exit Function
 Do While Not rs.EOF
  AuS = Right$(Space$(6) & rs!Pat_id, 6) & "  " & Left$(rs!name & Space$(20), 20) & "  " & Left$(rs!Fanf & Space$(10), 10) & "  " & Left$(rs!SchGr & Space$(4), 4) & "" & Right$(Space$(6) & rs!FID, 6) & "  " & Left$(rs!ausgst & Space$(10), 10) & "  " & Left$(rs!AktZeit & Space$(19), 19) & "  " & Left$(rs!Quartal & Space$(7), 7) & "  " & Left$(rs!lVorl & Space$(16), 16) & "  " & Left$(rs!GebOr & Space$(5), 5)
'  For i = 0 To rs.Fields.Count - 1
'   AuS = IIf(lenb(AuS) = 0, vns, AuS & " ") & Left$(rs.Fields(i) & Space$(14), 14)
'  Next i
'  Ausgeb (AuS)
  Lese.Ausgeb AuS, True
  rs.Move 1
 Loop
 AuS = "Pat_id  Name                  Fanf        Schgr  Fid  Ausgst      Aktzeit              Quartal  lVorl             GebOr"
' For i = 0 To rs.Fields.Count - 1
'  AuS = IIf(lenb(AuS) = 0, vns, AuS & " ") & Left$(rs.Fields(i).Name & Space$(14), 14)
' Next i
 Ausgeb (AuS)
 Exit Function
fehler:
 Exit Function
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fallzeig/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fallzeig

#If zutesten Then
 ' 28.10.18: nirgends aufgerufen
Public Function labtest()
 Dim erg$, rAf&
 Lese.ProgStart
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT (SELECT COUNT(0) FROM `laborxus` u WHERE datid = e.datid) zahl,datid,pfad,name FROM laborxeingel e"
 Do While Not rs.EOF
  erg = Dir(rs!Pfad)
  If LenB(erg) = 0 Then
   erg = Dir("\\anmeldr\biowinbackup\" & rs!name)
   If LenB(erg) <> 0 Then
    myEFrag "UPDATE laborxeingel SET pfad = CONCAT('\\\\anmeldr\\biowinbackup\\',name) WHERE datid = " & rs!DatID, rAf
'    Debug.Print rAF & " bei ƒnderung von: " & rs!name
   Else
'    Debug.Print rs!Zahl & " " & rs!Pfad
   End If
  End If
  rs.MoveNext
 Loop
' Debug.Print "Fertig!"
End Function ' labtest
#End If

Function getHatest(pid&)
 Dim infos$()
 Lese.ProgStart
' getHausarztAlt 52022, Infos
 Dim rFa() As Faelle
 Dim rKv1() As kvnrue
 getHausarzt1 infos, rFa, rKv1, , pid
 Stop
End Function ' getHatest

' aufgerufen in: gewEintrag, DiesQ, MDIForm_Load, tuAbstand, qtAnf, qtEnd
Public Function holFrist()
 Dim rv As New ADODB.Recordset
 Dim FristS
 On Error GoTo fehler
 Lese.ProgStart
 If myEFrag("SHOW FULL TABLES WHERE TABLE_TYPE LIKE '%VIEW%' AND Tables_in_" & DBCn.DefaultDatabase & "='aktfv'").EOF Then
  doViewsErstellen
 End If ' myEFrag(
 myEFrag ("USE " & DBCn.DefaultDatabase)
 myFrag rv, "SHOW CREATE VIEW `aktf`"
 If (1) Then
  FristS = rv.Fields(1)
  Set rv = Nothing
  FristS = Mid$(FristS, InStr(FristS, "interval ") + 9) ' interval muss klein sein
  FristS = Left$(FristS, InStr(FristS, " ") - 1)
  If Not IsNumeric(FristS) Then
   MsgBox "Ungeeignete Abfrage `aktfv`, evtl. erst Views erstellen"
   Exit Function
  End If
 Else
  FristS = 21
 End If
 holFrist = FristS
 Versp‰tung = FristS
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in holFrist/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' holFrist

#If vonnirgends Then
' aufgerufen nirgends
Function gewEintrag()
 Dim rn As New ADODB.Recordset, rf As New ADODB.Recordset, rE As New ADODB.Recordset, gew0 As New CString, gewStr As New CString, buch$, i&, gew#, za%, kz%, ZZ%, kzm% ' Ziffer (als letztes) angeh‰ngt, Kommazahl, Ziffernzahl, keine Ziffer mehr
 Lese.ProgStart
 Dim FristS$
 FristS = holFrist()
 myFrag rn, "SELECT pat_id FROM namen"
 Do While Not rn.EOF
  myFrag rf, "SELECT 0 FROM `faelle` WHERE pat_id = " & rn!Pat_id
  If Not rf.EOF Then
   Do While Not rf.EOF
    myFrag rE, "SELECT zeitpunkt, inhalt FROM `eintraege` e WHERE pat_id = " & rn!Pat_id & " AND art IN ('gew','gewicht') AND zeitpunkt BETWEEN " & lQAnfuEnd(FristS) & " ORDER BY zeitpunkt DESC"
    If Not rE.EOF Then
     gew0 = rE!Inhalt
     gewStr.Clear
     kz = 0
     ZZ = 0
     kzm = 0
     For i = 1 To Len(gew0)
      buch = Mid$(gew0, i, 1)
      If IsNumeric(buch) Then
       If kzm = 0 Then
        gewStr.Append buch
        za = True
        ZZ = 1
       End If
      ElseIf (buch = "." Or buch = ",") Then
       If za And kz = 0 Then
        gewStr.Append ","
        za = False
        kz = 1
       End If
      Else
       gewStr.Append " "
       za = False
       If ZZ = 1 Then kzm = True
      End If
     Next i
     If IsNumeric(gewStr) Then
      gew = CDbl(gewStr)
     Else
      gew = 0
     End If
     If gew <> gew0.ToNumber Then
      MsgBox "Fehler in gewEintrag: gew <> gew0.ToNumber"
      Stop
     End If
    End If
    Set rE = Nothing
    rf.MoveNext
   Loop
  End If
  Set rf = Nothing
  rn.MoveNext
 Loop
End Function ' gewEintrag()
#End If ' vonnirgends

Public Function ZahlEintrag() ' &Zahlen aus Eintr‰gen ermitteln
 Dim rn As New ADODB.Recordset, rf As New ADODB.Recordset, rE As New ADODB.Recordset, gew0 As New CString, gewStr As New CString, buch$, i&, gew#, za%, kz%, ZZ%, kzm% ' Ziffer (als letztes) angeh‰ngt, Kommazahl, Ziffernzahl, keine Ziffer mehr
 Dim bearbeitet&, vorhanden&, sql$, rafSum&, rAf&
 Lese.ProgStart
 sql = "SELECT id,inhalt FROM `eintraege` WHERE ISNULL(inhnum) OR inhnum=0 ORDER BY pat_id DESC"
 myFrag rE, "SELECT COUNT(0) zl FROM (" & sql & ") i"
 vorhanden = rE!zl
 Set rE = Nothing
 myFrag rE, sql
 Do While Not rE.EOF
  gew0 = rE!Inhalt
  myEFrag "UPDATE `eintraege` SET inhnum = " & REPLACE(gew0.ToNumber, ",", ".") & " WHERE id = " & rE!id, rAf
  rafSum = rafSum + rAf
  rE.MoveNext
  bearbeitet = bearbeitet + 1
  Lese.Ausgeb bearbeitet & "/" & vorhanden & " Datens‰tze bearbeitet, " & rafSum & " ge‰ndert", 0
 Loop
End Function ' zahlEintrag()

#If False Then
Function AllRisk(Age, Duration, Gender As String, AtrialFibrillation As String, Ethnic, Smoking, SysBP, HbA1c, LipidRatio, Time, hprec, bprec, lprec) As Variant
    Dim aAnswer(1 To 12)
    Dim vAnswer As Variant
    Dim nValue, nLower, nUpper
    With oPatient
        .AgeDiagDiabetes = Age
        .Female = (UCase$(Trim$(Gender)) = "F")
        .EthnicGroup = Ethnic
        .SmokingStatus = Smoking
        .BloodPressure = SysBP
        .HbA1c = HbA1c
        .LipidRatio = LipidRatio
        .Time = Time
        .AtrialFibrillation = (UCase$(Trim$(AtrialFibrillation)) = "Y")
        .DurationDiagnosedDiabetes = Duration
        .HbA1c_Precision = hprec
        .BloodPressure_Precision = bprec
        .LipidRatio_Precision = lprec
        
        '
        ' CHD
        '
        nValue = .CardiacRisk
        nLower = .CardiacLower
        nUpper = .CardiacUpper
        aAnswer(1) = nValue
        aAnswer(2) = nLower
        aAnswer(3) = nUpper
        '
        ' fCHD
        '
        nValue = .CardiacFatalRisk
        nLower = .CardiacFatalLower
        nUpper = .CardiacFatalUpper
        aAnswer(4) = nValue
        aAnswer(5) = nLower
        aAnswer(6) = nUpper
        '
        ' Stroke
        '
        nValue = .StrokeRisk
        nLower = .StrokeLower
        nUpper = .StrokeUpper
        aAnswer(7) = nValue
        aAnswer(8) = nLower
        aAnswer(9) = nUpper
        '
        ' Stroke
        '
        nValue = .StrokeFatalRisk
        nLower = .StrokeFatalLower
        nUpper = .StrokeFatalUpper
        aAnswer(10) = nValue
        aAnswer(11) = nLower
        aAnswer(12) = nUpper
    End With
    '
    ' We can't pass back an array but we can pass back
    ' a variant containing the array. So we assign the
    ' array to the variant AND return that AS the result
    '
    Let vAnswer = aAnswer
    AllRisk = vAnswer
End Function ' AllRisk

Function testrisk()
 Call AllRisk(55, 0, "F", "N", 1, 0, 140, 6, 4, 2, 6, 2, 10)
 Debug.Print oPatient.CardiacFatalRisk
End Function ' testrisk
#End If

Public Sub LiesKassen()
Call liesExcel(pVerz & "Patienten¸bergreifendes\Listenausgabe_Krankenkassen.xls", 5, "kassen")
myEFrag "DROP TABLE IF EXISTS kassenliste"
sql = "CREATE TABLE `kassenliste` (" & vbCrLf & _
  "`ID` int(10) NOT NULL AUTO_INCREMENT," & vbCrLf & _
  "`VKNR` varchar(6)  DEFAULT NULL," & vbCrLf & _
  "`IK` varchar(8)  DEFAULT NULL," & vbCrLf & _
  "`Name` varchar(100)  DEFAULT NULL," & vbCrLf & _
  "`Kateg` varchar(3)  DEFAULT NULL COMMENT 'Kategorie'," & vbCrLf & _
  "`AnzahlIK` int(4) UNSIGNED DEFAULT NULL," & vbCrLf & _
  "`AnzahlKTUG` int(4) UNSIGNED DEFAULT NULL," & vbCrLf & _
  "`G¸ltigVon` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'," & vbCrLf & _
  "`G¸ltigBis` datetime NOT NULL DEFAULT '0000-00-00 00:00:00' ," & vbCrLf & _
  "`GO` varchar(15)  DEFAULT NULL," & vbCrLf & _
  "`Kurzname` varchar(60)  DEFAULT NULL," & vbCrLf & _
  "`rName` varchar(80)  DEFAULT NULL COMMENT 'Kurzname, falls nicht verf¸gbar: Name'," & vbCrLf & _
  "`Lantus2` tinyint(1) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Mehrwertvertrag f¸r Lantus f¸r Typ-2-D.m.'," & vbCrLf & _
  "`Levemir2` tinyint(1) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Mehrwertvertrag f¸r Levemir f¸r Typ-2-D.m.'," & vbCrLf & _
  "`Humalog` tinyint(1) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Rabattvertrag f¸r Humalog'," & vbCrLf & _
  "`Liprolog` tinyint(1) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Rabattvertrag f¸r Liprolog'," & vbCrLf & _
  "`Novorapid` tinyint(7) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Rabattvertrag f¸r Novorapid'," & vbCrLf & _
  "`Apidra` tinyint(1) UNSIGNED NOT NULL DEFAULT 0 COMMENT 'Rabattvertrag f¸r Apidra'," & vbCrLf & _
  "`eingef` datetime NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'einf¸gedatum'," & vbCrLf & _
  "`geaen` datetime NOT NULL DEFAULT '0000-00-00 00:00:00' COMMENT 'zuletzt ge‰ndert'," & vbCrLf & _
  "`pid` int(10) DEFAULT NULL COMMENT 'pat_id, f¸r den Satz eingef¸gt wurde'," & vbCrLf & _
  "PRIMARY KEY (`ID`)," & vbCrLf & _
  "KEY `IK` (`IK`)," & vbCrLf & _
  "KEY `VKIK` (`VKNR`,`IK`)," & vbCrLf
  sql = sql & " KEY `kateg` (`Kateg`)" & vbCrLf & _
  ") ENGINE=InnoDB AUTO_INCREMENT=10422 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC"
  myEFrag sql
  myEFrag "INSERT INTO kassenliste(kurzname,ik,vknr,AnzahlIK,AnzahlKTUG,G¸ltigVon,G¸ltigBis,GO) SELECT kurzname,ik,vknr,`Anzahl IK`,`Anzahl KTUG`,`G¸ltig von`,`G¸ltig bis`,`GO` FROM kassen"
  DoKassenkategorienBestimmen
  Forms(0).Ausgeb "Fertig mit Lieskassen!", True
End Sub ' LiesKassen

Public Sub RufKassenKategorienBestimmen()
 Call Lese.ProgStart
 DoKassenkategorienBestimmen
 Call Lese.ProgEnde
End Sub ' RufKassenKategorienBestimmen

' in doEinles, liesKassen und RufKassenKategorienBestimmen
Public Sub DoKassenkategorienBestimmen()
 Dim namen$()
 syscmd 4, "Bestimme Kassenkategorien"
 ReDim namen(80)
 namen(0) = "aok": Call doKassKat("AOK", namen)
 namen(0) = "ikk": Call doKassKat("IKK", namen)
 namen(0) = "lkk":
 namen(1) = "g‰rtner"
 namen(2) = "wehr" ' 16.7.21: nach R¸cksprache Frau Sturm/ KV ¸ber Diabetesvereinbarung abrechenbar
 Call doKassKat("LKK", namen)
 namen(0) = "bkn": namen(1) = "knapp": namen(2) = "": Call doKassKat("BKN", namen)
 namen(0) = "SHV"
 namen(1) = "Amt"
 namen(2) = "Asyl"
 namen(3) = "KSA "
 namen(4) = "Bay. LI Gesundheit"
 namen(5) = "as stadt"
 namen(6) = "ba f"
 namen(7) = "migr"
 namen(8) = "pol "
 namen(9) = "kreis"
 namen(10) = "bgs"
 namen(11) = "jva"
 namen(12) = "ja "
 namen(13) = "job"
 namen(14) = "jugend"
 namen(15) = "polizei"
 namen(16) = "sa "
 namen(17) = "sht "
 namen(18) = "stadt"
 namen(19) = "ziv"
 namen(20) = "sht "
 namen(21) = "mdk "
 namen(22) = "medizinischer dienst "
 namen(23) = "aus "
 namen(24) = "bezirk "
 namen(25) = "BAGSFI"
 namen(26) = "institut"
 namen(27) = "behˆrde"
 namen(28) = "bezirk "
 namen(29) = "minist"
 namen(30) = "bg "
 namen(31) = "bpo "
 namen(32) = "aus "
 namen(33) = "bza "
 namen(34) = "stelle"
 namen(35) = "diakonie"
 namen(36) = "verwaltung"
 namen(37) = "unterst"
 namen(38) = "papierlose"
 namen(39) = "ref."
 namen(40) = "einrichtung"
 namen(41) = "f¸rsorge"
 namen(42) = "hilfs"
 namen(43) = "justiz"
 namen(44) = "feuerw"
 namen(45) = "sozial"
 namen(46) = "kvno"
 namen(47) = "kv "
 namen(48) = "lab "
 namen(49) = "landesamt"
 namen(50) = "kvwl"
 namen(51) = "lasv "
 namen(52) = "lra "
 namen(53) = "lva "
 namen(54) = "jug. "
 namen(55) = "med.dienst"
 namen(56) = "musterung"
 namen(57) = "bez¸ge"
 namen(58) = "obdach"
 namen(59) = "pr‰sidium"
 namen(60) = "rp "
 namen(61) = "schwulen"
 namen(62) = "sengpg"
 namen(63) = "smi n‰v"
 namen(64) = "anderer kt"
 namen(65) = "BAS"
 Call doKassKat("SHV", namen, True)
 ReDim namen(30)
 namen(0) = "BARMER"
 namen(1) = "DAK"
 namen(2) = "TKK"
 namen(3) = "KKH"
 namen(4) = "Kaufm‰nnische"
 namen(5) = "HEK"
 namen(6) = "HMK"
 namen(7) = "HKK"
 namen(8) = "GEK"
 namen(9) = "Gm¸nder Ersatz"
 namen(10) = "HZK"
 namen(11) = "KEH"
 namen(12) = "HAMBURG-M‹NCHENER"
 namen(13) = "Handelskrankenkasse"
 namen(14) = "Techniker"
 namen(15) = "HANSEATISCHE"
 namen(16) = "AUS "
 namen(17) = "EK"
 namen(18) = "TK t"
 namen(19) = "TK Rheinland"
 namen(20) = "buchdrucker"
 Call doKassKat("EK", namen, False)
 ReDim namen(20)
 namen(0) = "bkk": namen(1) = "betriebsk": namen(2) = "sbk"
 namen(3) = "viactiv"
 namen(4) = "actimonda"
 namen(5) = "heimat "
 Call doKassKat("BKK", namen)
 ReDim namen(20)
 namen(0) = "PBeaKK"
 Call doKassKat("PBe", namen)
 syscmd 4, "Fertig mit Kassenkategorienbestimmen"
' Debug.Print "Fertig mit Kassenkategorienbestimmen"
End Sub ' DoKassenkategorienBestimmen

Sub doKassKat(Kateg$, namen$(), Optional isn%)
 Dim rs As ADODB.Recordset, rAf&, sql$, i%
 sql = "UPDATE `kassenliste` SET geaen=" & Format(Now(), "yyyymmddHHMMSS") & ",Kateg='" & Kateg & "' WHERE ("
 i = 0
 Do
  If LenB(namen(i)) = 0 Then Exit Do
  sql = sql & "name LIKE '%" & namen(i) & "%' OR kurzname LIKE '%" & namen(i) & "%'"
  namen(i) = vNS
  i = i + 1
  If LenB(namen(i)) <> 0 Then
   sql = sql & "OR "
  End If
 Loop
 sql = sql & ") AND (Kateg <> '" & Kateg & "' OR Kateg='' OR ISNULL(Kateg))"
 If isn <> 0 Then
  sql = sql & " AND (ISNULL(Kateg) OR Kateg='')"
 End If
 Set rs = myEFrag(sql, rAf)
 If rAf <> 0 Then
  Lese.Ausgabe = Lese.Ausgabe & "In die Kategorie '" & Kateg & "' wurden " & rAf & " Kassen eingeteilt." & vbCrLf
  altAusgabe = Lese.Ausgabe
 End If
End Sub ' doKassKat


Public Sub liesExcel(Datei$, ‹Zeile%, Tbl$)
 Const XStra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
 Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
 Dim rX As New ADOX.Catalog, sql$, rAf&
 Dim XCon As New ADODB.Connection
 Dim rEx As New ADODB.Recordset, rs As New ADODB.Recordset
 Dim InS$
 FNr = 19
 On Error GoTo fehler
 DBCnS = DBCn
 Set XCon = Nothing
 XCon.Open XStra & Datei & XStrb
 Set rX = Nothing
 rX.ActiveConnection = XCon
 rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", XCon ' Hier Excel, nicht obmysql = 0!
 Dim zeile&
 Dim j&
 zeile = 1
 Dim ArtM() ' 0=numeric, 1=date, 2=string
 Dim LenM()
 Dim btlm() ' Bruchteill‰ngen-Matrix
 Dim SpNm()
 Dim pos&, btlmakt&, Inh$
 Do While Not rEx.EOF
'  Debug.Print Zeile, ": ", rEx.Fields.Count
  If zeile = ‹Zeile Then
   ReDim ArtM(rEx.Fields.COUNT - 1)
   ReDim LenM(rEx.Fields.COUNT - 1)
   ReDim SpNm(rEx.Fields.COUNT - 1)
   ReDim btlm(rEx.Fields.COUNT - 1)
   For j = 0 To rEx.Fields.COUNT - 1
    SpNm(j) = rEx.Fields(j)
   Next j
  ElseIf zeile > ‹Zeile Then
   For j = 0 To rEx.Fields.COUNT - 1
    If IsNull(rEx.Fields(j)) Then Exit For
    Inh = IIf(rEx.Fields(j) = "*", "", rEx.Fields(j))
    pos = InStr(Inh, ".")
    If pos = 0 Then pos = InStr(Inh, ",")
    If pos > 0 Then
     btlmakt = Len(Inh) - pos
     If btlmakt > btlm(j) Then btlm(j) = btlmakt
    End If
    If Len(Inh) > LenM(j) Then LenM(j) = Len(Inh)
    If ArtM(j) < 2 Then
     If ArtM(j) = 0 Then
      If Not IsNumeric(Inh) Then ArtM(j) = 1
     End If
     If ArtM(j) = 1 Then
      If Inh <> "" And Not IsDate(Inh) Then ArtM(j) = 2
     End If
    End If
'    Debug.Print rEx.Fields(j)
   Next
  End If
  zeile = zeile + 1
  rEx.Move 1
 Loop
' Dim maxsp&
' maxsp = rEx.Fields.Count
 On Error Resume Next
 myEFrag ("DROP TABLE IF EXISTS `" & Tbl & "`")
 On Error GoTo fehler
 sql = "CREATE TABLE `" & Tbl & "` (id INTEGER(10) AUTO_INCREMENT KEY"
 For j = 0 To UBound(SpNm)
  If ArtM(j) = 0 Then
   If btlm(j) > 0 Then
    sql = sql & ",`" & SpNm(j) & "` DECIMAL(" & MINvb(LenM(j) + 1 + btlm(j), 65) & "," & MINvb(btlm(j), 30) & ")"
   Else
    sql = sql & ",`" & SpNm(j) & "` INTEGER(" & LenM(j) & ")"
   End If
  ElseIf ArtM(j) = 1 Then
   sql = sql & ",`" & SpNm(j) & "` DATETIME"
  Else
   sql = sql & ",`" & SpNm(j) & "` VARCHAR(" & LenM(j) & ")"
  End If
  InS = InS & "`" & SpNm(j) & "`"
  If j <> UBound(SpNm) Then InS = InS & ","
 Next
 sql = sql & ")"
 myEFrag sql, rAf
 rEx.MoveFirst
 zeile = 1
 Do While Not rEx.EOF
  If zeile > ‹Zeile Then
    If IsNull(rEx.Fields(1)) Then Exit Do
    sql = "INSERT INTO `" & Tbl & "`(" & InS & ") VALUES("
    For j = 0 To UBound(SpNm)
     Inh = IIf(rEx.Fields(j) = "*", "", rEx.Fields(j))
     If ArtM(j) = 1 Then
      If Inh = "" Then sql = sql & "0" Else sql = sql & Format(DateValue(Inh), "yyyymmdd")
     Else
      If ArtM(j) = 0 And btlm(j) <> 0 Then
       sql = sql & REPLACE(Inh, ",", ".")
      Else
       If ArtM(j) = 2 Then sql = sql & "'"
       If ArtM(j) = 0 And Inh = "" Then sql = sql & "0" Else sql = sql & UmwfSQL(Inh)
       If ArtM(j) = 2 Then sql = sql & "'"
      End If
     End If
     If j <> UBound(SpNm) Then sql = sql & ","
    Next j
    sql = sql & ")"
    InsKorr DBCn, DBCnS, sql, rAf
  End If
  zeile = zeile + 1
  rEx.Move 1
 Loop
 Forms(0).Ausgeb "Datei `" & Datei & "` in Tabelle `" & Tbl & "` eingelesen", True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lieseExcel/" + AnwPfad)
  Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' liesExcel

' in Kontrolllisten_f¸r_DMP_HA_Click und tubriefStandalone
Public Function umlweg$(ByRef q$)
' IF InStrB(q, "‹") OR InStrB(q, "ﬂ") THEN Stop
 umlweg = REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(q, "‰", "\'e4"), "ƒ", "\'c4"), "ˆ", "\'f6"), "÷", "\'d6"), "¸", "\'fc"), "‹", "\'dc"), "ﬂ", "\'df")
End Function ' umlweg$

