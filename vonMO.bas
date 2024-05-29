Attribute VB_Name = "vonMo"
Option Explicit
Option Compare Text
Dim aru&

Const MoSer$ = "wser"
Const MoSzn$ = "szn4"
Const parsemotxt$ = "v:\Parsememo31.txt"
Const mestausg$ = "v:\mestr.txt"
Public Const MOCStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSer & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
Public Const MOsStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSzn & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
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

Public Sub testeb()
 Dim eb() As ebType
 ReDim eb(6)
 eb(0).nr = 0: eb(0).Wert = 0
 eb(1).nr = 1: eb(1).Wert = 1
 eb(2).nr = 2: eb(2).Wert = 2
 eb(3).nr = 3: eb(3).Wert = 3
 eb(4).nr = 4: eb(4).Wert = 4
 eb(5).nr = 5: eb(5).Wert = 5
 eb(6).nr = 6: eb(6).Wert = 6
 CopyMemoryPtr ByVal VarPtr(eb(4)), VarPtr(eb(5)), 4 * (UBound(eb) - 4)
End Sub

Public Sub ebaltloe(ByRef eb() As ebType, ByVal pos&, obloe%)
 Static loezahl%
 Dim i&
 If pos >= 0 Then
  If pos <= UBound(eb) Then
   If pos < UBound(eb) Then
    For i = pos + 1 To UBound(eb)
     eb(i - 1) = eb(i)
    Next i
   End If
   loezahl = loezahl + 1
  End If
 End If
 If obloe And loezahl Then
  ReDim Preserve eb(UBound(eb) - loezahl)
  loezahl = 0
 End If
End Sub ' ebaltloe

Public Sub memoaltloe(ByRef memo() As memoType, ByVal pos&, obloe%)
 Static loezahl%
 Dim i&
 If pos >= 0 Then
  If pos <= UBound(memo) Then
   If pos < UBound(memo) Then
    For i = pos + 1 To UBound(memo)
     memo(i - 1) = memo(i)
    Next i
   End If
   loezahl = loezahl + 1
  End If
 End If
 If obloe And loezahl Then
  ReDim Preserve memo(UBound(memo) - loezahl)
  loezahl = 0
 End If
End Sub ' memoaltloe

Public Sub ebloe(ByRef AryVar() As ebType, ByVal RemoveWhich&, obloe%)
    '// The size of the array elements
    '// In the case of string arrays, they are
    '// simply 32 bit pointers to BSTR's.
    Dim byteLen As Byte
    Static loezahl%
    ' byteLen = 4   '// String pointers are 4 bytes
    '// The copymemory operation is not necessary unless
    '// we are working with an array element that is not
    '// at the end of the array
    If RemoveWhich >= 0 And RemoveWhich < UBound(AryVar) Then
        byteLen = VarPtr(AryVar(1)) - VarPtr(AryVar(0))
        '// Copy the block of string pointers starting at the position after the removed item back one spot.
        CopyMemoryPtr ByVal VarPtr(AryVar(RemoveWhich)), ByVal VarPtr(AryVar(RemoveWhich + 1)), (byteLen) * (UBound(AryVar) - RemoveWhich)
        loezahl = loezahl + 1
    End If
   '// If we are removing the last array element
    '// just deinitialize the array
    '// otherwise chop the array down by one.
    If UBound(AryVar) = LBound(AryVar) Then
        Erase AryVar
        loezahl = 0
    ElseIf obloe Then
        ReDim Preserve AryVar(UBound(AryVar) - loezahl)
        loezahl = 0
    End If
End Sub ' ebloe

Public Sub memoloe(ByRef AryVar() As memoType, ByVal RemoveWhich&, obloe%)
    '// The size of the array elements
    '// In the case of string arrays, they are
    '// simply 32 bit pointers to BSTR's.
    Dim byteLen As Byte
    Static loezahl%
    ' byteLen = 4   '// String pointers are 4 bytes
    '// The copymemory operation is not necessary unless
    '// we are working with an array element that is not
    '// at the end of the array
    If RemoveWhich >= 0 And RemoveWhich < UBound(AryVar) Then
        byteLen = VarPtr(AryVar(1)) - VarPtr(AryVar(0))
        '// Copy the block of string pointers starting at the position after the removed item back one spot.
        CopyMemoryPtr ByVal VarPtr(AryVar(RemoveWhich)), ByVal VarPtr(AryVar(RemoveWhich + 1)), (byteLen) * (UBound(AryVar) - RemoveWhich)
        loezahl = loezahl + 1
    End If
   '// If we are removing the last array element
    '// just deinitialize the array
    '// otherwise chop the array down by one.
    If UBound(AryVar) = LBound(AryVar) Then
        Erase AryVar
        loezahl = 0
    ElseIf obloe Then
        ReDim Preserve AryVar(UBound(AryVar) - loezahl)
        loezahl = 0
    End If
End Sub ' memoloe
    
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
Public Function ebQuickSort(ByRef pvarArray() As ebType, Optional ByVal plngLeft&, Optional ByVal plngRight&) As ebType()
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

' BLOB-Felder aus Medical Office parsen
Public Function ParseMemo(FMemo$, MeStr() As memoType, rFa() As Faelle, Optional obDebug%, Optional ÜSchr$) ' pNr&, FSur&,
 Dim jj&, iru&
 Dim gl&, pos&, MAX&, aktmax&, altmax&, tlen&, ie&, altie&, mznr%, i&
 Dim obDruck%, txt$, ebS$
 Dim eb() As ebType
 On Error GoTo fehler
 aru = aru + 1
 Erase MeStr
 If FMemo <> "" Then
  pos = 1
  ie = 0
  If obDebug Then
   Open parsemotxt For Append As #255
   Print #255, ÜSchr & ":"
  End If ' obDebug
  iru = 0
  Do ' Schleife labt
   iru = iru + 1
   obDruck = 0
   tlen = Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1))
   If pos = 1 Then gl = tlen: MAX = gl
   altmax = MAX
   aktmax = tlen + pos + 1 ' aktuell angegebene Länge aus dem und dem nä Byte
   If aktmax >= 0 And aktmax <= gl + 2 Then ' wenn die angegebene Länge vertretbar
   ' und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen
   ' wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde:
    Dim obnaechst%
    obnaechst = 0
    If pos > MAX Then
     obnaechst = True
    Else
     If (Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Then
      obnaechst = True
     End If
    End If ' pos > MAX
    If obnaechst Then
     altie = ie ' letzte Ebene
     If aktmax <= MAX Then ie = ie + 1 ' im ersten Fall wird die Ebene erhöht
     If pos > MAX Then ' im zweiten Fall wird zurückgegriffen
      mznr = -1
      ie = 0
      If SafeArrayGetDim(MeStr) <> 0 Then
       For i = 0 To UBound(MeStr)
        If MeStr(i).mx >= aktmax And MeStr(i).znr > mznr Then mznr = MeStr(i).znr ' MeStr(i).patnr = pNr And
       Next i
       For i = 0 To UBound(MeStr)
        If MeStr(i).mx >= aktmax And MeStr(i).znr = mznr And MeStr(i).ebn > ie Then ie = MeStr(i).ebn ' MeStr(i).patnr = pNr And
       Next i
      End If ' SafeArrayGetDim(MeStr) <> 0 Then
      ie = ie + 1
     End If ' pos > MAX
     If ie <= altie Then ' wenn die Ebene nicht erhöht worden ist
      If ie < altie Then ' wenn sie vielmehr reduziert wurde
       For i = UBound(eb) To 0 Step -1
        If eb(i).nr > ie Then
         Call ebaltloe(eb, i, False) ' dann werden die höheren Einträge wieder gelöscht
        End If
       Next i
      End If ' ie < AltID Then
      For i = 0 To UBound(eb)
       If eb(i).nr = ie Then
        eb(i).Wert = eb(i).Wert + 1 ' dann wird die Zählung der akt. Ebene erhöht
        Exit For
       End If
      Next i
      Call ebaltloe(eb, -1, True)
     Else ' ie < AltID Then
      ' sonst muss ein neuer Eintrag für die hohe Ebene erstellt werden
      If SafeArrayGetDim(eb) = 0 Then ReDim eb(0) Else ReDim Preserve eb(UBound(eb) + 1)
      eb(UBound(eb)).nr = ie
      eb(UBound(eb)).Wert = 1
     End If ' ie <= altie Then
     If obDebug Then
      Dim kk&
      Print #255, "eb:"
      For kk = 0 To UBound(eb)
       Print #255, kk, eb(kk).nr, eb(kk).Wert
      Next kk
     End If
     MAX = aktmax ' neue vorbestehende Länge
     obDruck = 1 ' und drucken
    End If ' ((Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Or pos > MAX) Then
   End If ' aktmax >= 0 And aktmax <= gl Then
   If obDruck = 1 Then ' Or pos = 1 Then ' or pos=1 or True then
'   If True Then
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
     If SafeArrayGetDim(MeStr) = 0 Then
       ReDim MeStr(0)
     Else ' SafeArrayGetDim(MeStr) = 0 Then
       ' das Folgende geht leider nicht, weil dann oben eb() falsch befüllt wird
       ' zum Debuggen das Folgende mit "or True" ergänzen
'       If Not (InStr(ebS, MeStr(UBound(MeStr)).enr & ".") = 1 And ebS <> MeStr(UBound(MeStr)).enr) Then
        ReDim Preserve MeStr(UBound(MeStr) + 1)
'       End If
     End If ' SafeArrayGetDim(MeStr) = 0 Then else
'     MeStr(UBound(MeStr)).patnr = pNr
'     MeStr(UBound(MeStr)).FSur = FSur
     MeStr(UBound(MeStr)).znr = pos
     MeStr(UBound(MeStr)).mx = MAX
     MeStr(UBound(MeStr)).ebn = ie
     MeStr(UBound(MeStr)).enr = ebS
     If Asc(Right$(txt, 1)) = 0 Then txt = Left$(txt, Len(txt) - 1)
     MeStr(UBound(MeStr)).Text = txt
'     If obDebug Then
'      Print #255, pos & "|" & MAX & "|" & ie & "|" & ebS & "|" & Mid$(FMemo, pos, 1) & "|" & Asc(Mid$(FMemo, pos, 1)) & "|" & Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1)) & "| aktMax: " & aktmax & "| altMax: " & altmax & "| Laenge: " & Len(txt) & "| EndByte: " & Asc(Mid$(FMemo, aktmax, 1)) & "|" & txt & vbCrLf
'     End If ' obDebug Then
    Else ' txt <> String$(Chr$(0), tlen) Then ' 0-Strings nicht auffieseln
     pos = pos + tlen
    End If ' txt <> String$(Chr$(0), tlen) Then
   End If ' obdruck = 1 Then ' or pos=1 or true
   pos = pos + 1
   If pos >= gl Then Exit Do
  Loop
  ' folende Zeilen zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
  For i = UBound(MeStr) - 1 To 0 Step -1
   If InStr(MeStr(i + 1).enr, MeStr(i).enr & ".") = 1 And MeStr(i + 1).enr <> MeStr(i).enr Then ' MeStr(i).patnr = MeStr(i + 1).patnr And MeStr(i).FSur = MeStr(i + 1).FSur And
    Call memoaltloe(MeStr, i, False)
   End If
  Next i
  Call memoaltloe(MeStr, -1, True)
  If obDebug Then
   For i = 0 To UBound(MeStr)
    Dim endse$, endsz$
    If MeStr(i).znr <= Len(FMemo) Then
     endse = Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     endse = "!: " & MeStr(i).znr & ">" & Len(FMemo)
    End If
    If MeStr(i).znr <= Len(FMemo) - 1 Then
     endsz = Asc(Mid$(FMemo, MeStr(i).znr + 1, 1)) * 256& + Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     endsz = "!: " & MeStr(i).znr - 1 & ">" & Len(FMemo)
    End If
    Print #255, MeStr(i).znr & "|" & MeStr(i).mx & "|" & MeStr(i).ebn & "|" & MeStr(i).enr & "|" & IIf(endse <> "10", Mid$(FMemo, MeStr(i).znr, 1), "") & "|" & endse & "|" & endsz & "| Laenge: " & Len(MeStr(i).Text) & "|" & IIf(Right$(MeStr(i).Text, 1) = Chr$(10), Left$(MeStr(i).Text, IIf(Len(MeStr(i).Text) = 0, 1, Len(MeStr(i).Text)) - 1), MeStr(i).Text)
   Next i
  End If ' obdebug
  If obDebug Then
   Print #255, vbCrLf
   Close #255
  End If ' obDebug
 End If ' FMEMO<>""
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ParseMemo/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ParseMemo

' einige Male in doPatvonMO auskommentiert ausgerufen
Function MeStDruck(Eig$, MeStr() As memoType)
 Dim i&
 Open mestausg For Append As #245
 Print #245, vbCrLf & Eig & ":"
 Print #245, "i znr  ebn  enr                 mx text"
 For i = 0 To UBound(MeStr)
  Print #245, CStr(i) & Space$(4 - Len(CStr(i))) & _
  MeStr(i).znr & Space$(5 - Len(CStr(MeStr(i).znr))) & _
  MeStr(i).ebn & Space$(4 - Len(MeStr(i).ebn)) & _
  MeStr(i).enr & Space$(19 - Len(MeStr(i).enr)) & _
  MeStr(i).mx & Space$(5 - Len(MeStr(i).mx)) & _
  MeStr(i).Text
 Next i
 Close #245
End Function ' MestDruck

Public Function testfloat()
 Dim d#
 d = 2.1
 Dim dptr&, cvptr&, cptr&
 dptr = VarPtr(d)
 
 Dim c As String * 10
 cvptr = VarPtr(c)
 cptr = StrPtr(c)
 cvptr = VarPtr(c)
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
 Dim rt As New ADODB.Recordset
 rt.Open "select text from tmpmpatfall where patnr=53119 and enr='5.22'", MOCon, adOpenStatic, adLockReadOnly
' Debug.Print rt.Fields(0)
' CopyMemoryPtr VarPtr(d), VarPtr(rt.Fields(0)), 8
 CopyMemoryPtr cptr - 4, VarPtr(10), 4
 CopyMemoryPtr cptr, dptr, 8
 CopyMemoryPtr VarPtr(c) + 3, dptr, 6
 Debug.Print c
End Function

' String zu double; Formel experimentell ermittelt über Datei
Public Function stzd#(s$)
 stzd = (1 + (Asc(Mid$(s, 7, 1)) Mod 16) / 16 + Asc(Mid$(s, 6, 1)) / 4096 + Asc(Mid$(s, 5, 1)) / 2 ^ 20 + Asc(Mid$(s, 4, 1)) / 2 ^ 28 + Asc(Mid$(s, 3, 1)) / 2 ^ 36 + Asc(Mid$(s, 2, 1)) / 2 ^ 44 + Asc(Mid$(s, 1, 1)) / 2 ^ 52) * 2 ^ (Int(Asc(Mid$(s, 7, 1)) / 16) + 1 + 16 * (Asc(Mid$(s, 8, 1)) - 64))
End Function ' stzd

' String zu Kalender; Formel experimentell ermittelt über Datei
Public Function stzk(s$) As Date
 stzk = CDate(CStr(Asc(Mid$(s, 4))) & "." & CStr(Asc(Mid$(s, 3))) & "." & CStr(256 * Asc(Mid(s, 2)) + Asc(Mid(s, 1))))
End Function ' stzk


' in PatvonMO_Click
Public Function doPatvonMO(pNr&)
 Const obDebug% = True
 Static lfdfl&
 Dim pid&, pos&, SchGr%, jj%, rAf&, aktZeit As Date
 Dim rsNa As New ADODB.Recordset, rsFa As New ADODB.Recordset
 Dim NaStr() As memoType, FaStr() As memoType, rsfaru%, j&
 pNr& = 68927 ' 151 ' 225 ' 68316 ' 65405 ' 45 ' 64659 ' 45 ' 69367 ' 69377 ' 53119 ' 51630 ' 105 ' 18 ' 246 ' 59152 ' 1394 ' 2112
 aktZeit = Now()
 pid = pNr + 100000
 Call LöschePat(pid)
 On Error Resume Next
 If obDebug Then FSO.DeleteFile parsemotxt
 On Error GoTo fehler
 Tinit
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
 ' konnte nicht genau rausfinden, wann FMemo richtig übermittelt wird, evtl. erst als zweites Feld, evtl. nicht unter dem Namen FMemo
 ' unter Ado muss es auf latin1 übersetzt werden für die Zahlen > 128
 sql = "SELECT fsurogat nix, COALESCE(CONVERT(FMemo USING latin1),'') Fm, p.* from patstamm p WHERE FSurogat=" & pNr
 rsNa.Open sql, MOCon, adOpenStatic, adLockReadOnly
 If Not rsNa.BOF Then
 
  rNa(0).aktZeit = aktZeit
  rNa(0).Pat_id = pid ' = pNr
  rNa(0).lfdnr = -1 ' Import aus MO
  rNa(0).Nachname = rsNa!FNachname
  rNa(0).NVorsatz = rsNa!FNamensvorsatz
  rNa(0).NVors = rsNa!FNamenszusatz
  rNa(0).Titel = rsNa!FTitel
  rNa(0).Vorname = rsNa!fVorname
  Select Case rsNa!FGeschlecht:  Case "2": rNa(0).geschlecht = "w":  Case "1": rNa(0).geschlecht = "m":  Case Else: rNa(0).geschlecht = rsNa!FGeschlecht: End Select
  rNa(0).GebDat = DateSerial(Mid$(rsNa!fgeburtsdatum, 1, 4), Mid$(rsNa!fgeburtsdatum, 5, 2), Mid$(rsNa!fgeburtsdatum, 7, 2))
  rNa(0).Versichertennummer = rsNa!FAktversichertennr
  ' fehlt rna(0).KarGen: aus ' ',0,2,3
  If rsNa!FErstkontakt <> 0 Then rNa(0).AufnDat = CDate("1.1.1890") + rsNa!FErstkontakt
 ' rna(0).kAufDat=
  rNa(0).Straße = rsNa!FStrasse
  If Not IsNull(rsNa!fHausnr) Then If Not rsNa!fHausnr = "" Then rNa(0).Straße = rNa(0).Straße & " " & rsNa!fHausnr
  rNa(0).Hausnr = rsNa!fHausnr
  rNa(0).plz = rsNa!FPlz
  rNa(0).ort = rsNa!FOrt
  rNa(0).Lkz = rsNa!FLaendercode
  ' nicht enthalten/nicht befüllt: Anschrzus, PFPlz,anschrzus_2,postfach_2,lk_2,postfach,beruf,Weggeldzone,
  rNa(0).WeggzZahl = IIf(rsNa!FEntfernung = "", 0, rsNa!FEntfernung) ' bei Pat. 1219 doch nicht die Weggeldzone als Zahl
  rNa(0).Titel = rsNa!FTitel
 
  rNa(0).PrivatTel = rsNa!FTelefonprivat
  pos = InStr(rNa(0).PrivatTel, " (")
  If pos <> 0 Then rNa(0).PrivatTel = Left$(rNa(0).PrivatTel, pos - 1)
 
  If InStrB(rsNa!ftelefonmobil, "dienstlich") <> 0 Then
   rNa(0).DienstTel = rsNa!ftelefonmobil
   pos = InStr(rNa(0).DienstTel, " (")
   If pos <> 0 Then rNa(0).DienstTel = Left$(rNa(0).DienstTel, pos - 1)
  ElseIf InStrB(rsNa!ftelefonmobil, "Funktelefon") <> 0 Then
   rNa(0).PrivatMobil = rsNa!ftelefonmobil
   pos = InStr(rNa(0).PrivatMobil, " (")
   If pos <> 0 Then rNa(0).PrivatMobil = Left$(rNa(0).PrivatMobil, pos - 1)
  Else
   rNa(0).PrivatTel_2 = rsNa!ftelefonmobil
   pos = InStr(rNa(0).PrivatTel_2, " (")
   If pos <> 0 Then rNa(0).PrivatTel_2 = Left$(rNa(0).PrivatTel_2, pos - 1)
  End If
 
  If rsNa!ftelefondienst <> "" Then
   rNa(0).PrivatMobil = rsNa!ftelefondienst
   pos = InStr(rNa(0).PrivatMobil, " (")
   If pos <> 0 Then rNa(0).PrivatMobil = Left$(rNa(0).PrivatMobil, pos - 1)
  End If
 
  rNa(0).PrivatFax = rsNa!ffax
  pos = InStr(rNa(0).PrivatFax, " (")
  If pos <> 0 Then rNa(0).PrivatFax = Left$(rNa(0).PrivatFax, pos - 1)
 
  rNa(0).email = rsNa!femail
  pos = InStr(rNa(0).email, " (")
  If pos <> 0 Then rNa(0).email = Left$(rNa(0).email, pos - 1)
  
  
  On Error Resume Next
  FSO.DeleteFile mestausg
  On Error GoTo fehler
  If rsNa!fm <> "" Then
   Call ParseMemo(rsNa!fm, NaStr(), rFa(), obDebug, "FMemo von rsNa, Pat-id: " & pNr)
'   Call MeStDruck(CStr(pNr), NaStr)
   For j = 0 To UBound(NaStr)
'     Debug.Print NaStr(j).enr, NaStr(j).Text
    Select Case NaStr(j).enr
      Case "18"
         rNa(0).Notiz = REPLACE$(LTrim$(REPLACE$(NaStr(j).Text, "Notiz:", "")), Chr$(10), vbCrLf)
      Case "21.1": ' asc( Zahl der eingetragenen Kinder
      Case "25"
        If NaStr(j).Text <> 0 Then
         SchGr = 90
        End If
    End Select
   Next j
  End If ' rsNa!fm
'  Call MeStDruck(CStr(pNr), NaStr)
  
  rsFa.Open "SELECT f.fsurogat nix, COALESCE(CONVERT(f.FMemo USING latin1),'') Fm,CONCAT(f.fpatnr,', ',18900101 + INTERVAL f.fvon DAY,' - ',18900101 + INTERVAL f.fbis DAY) ueschr, f.*,a.FBezeichnung, le.FNachname FROM patfall f LEFT JOIN abrechner a ON f.FArztnr=a.FSurogat LEFT JOIN lstgerb le USING (FLstgerbnr) WHERE fpatnr=" & pNr & " ORDER BY FVon DESC", MOCon, adOpenStatic, adLockReadOnly
'  rsfaru = 0
  If Not rsFa.BOF Then
   Do While Not rsFa.EOF
    lfdfl = lfdfl + 1
    ReDim Preserve rFa(UBound(rFa) + 1)
    rFa(UBound(rFa)).aktZeit = aktZeit
    rFa(UBound(rFa)).lfdnr = lfdfl
    rFa(UBound(rFa)).Pat_id = pid
    rFa(UBound(rFa)).AbrAr = ""
    rFa(UBound(rFa)).VermiArt = 0
    rFa(UBound(rFa)).bPerG = "0"
'     rFa(UBound(rFa)).GebOr = 2 ' in patfall und tmpmpatfall bei Pat. 151 kein Unterscheidungskriterium gefunden
'     rFa(UBound(rFa)).AbrAr = 2 ' in patfall und tmpmpatfall bei Pat. 151 kein Unterscheidungskriterium gefunden
    If rsFa!fvon <> 0 Then rFa(UBound(rFa)).BhFB = CDate("1.1.1890") + rsFa!fvon ' wenngleich nie 0
    rFa(UBound(rFa)).Quartal = ZQuart(rFa(UBound(rFa)).BhFB)
    If Not IsNumeric(rFa(UBound(rFa)).Quartal) Or Len(rFa(UBound(rFa)).Quartal) <> 5 Then
     Stop
    End If
    If rsFa!fbis <> 0 Then rFa(UBound(rFa)).BhFE1 = CDate("1.1.1890") + rsFa!fbis ' wenngleich nie 0
    If rsFa!fabgerechnet <> 0 Then rFa(UBound(rFa)).BhFE2 = CDate("1.1.1890") + rsFa!fabgerechnet
    rFa(UBound(rFa)).BhFE = rFa(UBound(rFa)).BhFE2
    rFa(UBound(rFa)).IK = Right$(rsFa!Fik, 7)
    rFa(UBound(rFa)).abrArzt = rsFa!FBezeichnung
    If rsFa!FVorhanden <> " " And rsFa!FVorhanden <> "" Then
      rFa(UBound(rFa)).KartBes = rsFa!FVorhanden
    End If
'    rFa(UBound(rFa)).lanrid = IIf(rsFa!FLstgerbnr = 3, 2, 1) ' 2 = Schade, 3 = Kothny
    rFa(UBound(rFa)).lanrid = IIf(InStrB(rsFa!FNachname, "Kothny") <> 0, 2, 1) ' 2 = Schade, 3 = Kothny
    Select Case rsFa!fscheintyp
     Case 1
      rFa(UBound(rFa)).SchGr = "90"
      Select Case rsFa!fscheinart
       Case 0: rFa(UBound(rFa)).GOÄKatNr = "01": rFa(UBound(rFa)).GOÄKatName = "Privat"
       Case 1: rFa(UBound(rFa)).GOÄKatNr = "02": rFa(UBound(rFa)).GOÄKatName = "Postbeamte B"
       Case 2: rFa(UBound(rFa)).GOÄKatNr = "03": rFa(UBound(rFa)).GOÄKatName = "Bundesbahn (KVB I-III)"
       Case 3: rFa(UBound(rFa)).GOÄKatNr = "04": rFa(UBound(rFa)).GOÄKatName = "Postbeamte Dienstunfälle"
       Case 4: rFa(UBound(rFa)).GOÄKatNr = "05": rFa(UBound(rFa)).GOÄKatName = "Bundesbahn Dienstunfälle"
       Case 5: rFa(UBound(rFa)).GOÄKatNr = "06": rFa(UBound(rFa)).GOÄKatName = "Knappschaft"
       Case 6: rFa(UBound(rFa)).GOÄKatNr = "07": rFa(UBound(rFa)).GOÄKatName = "Privat mit MwSt."
       Case 7: rFa(UBound(rFa)).GOÄKatNr = "08": rFa(UBound(rFa)).GOÄKatName = "Basistarif"
       Case 8: rFa(UBound(rFa)).GOÄKatNr = "09": rFa(UBound(rFa)).GOÄKatName = "Einfachtarif"
       Case 9: rFa(UBound(rFa)).GOÄKatNr = "10": rFa(UBound(rFa)).GOÄKatName = "Studententarif"
       Case 10: rFa(UBound(rFa)).GOÄKatNr = "11": rFa(UBound(rFa)).GOÄKatName = "IGeL-Leistungen"
       Case 11: rFa(UBound(rFa)).GOÄKatNr = "12": rFa(UBound(rFa)).GOÄKatName = "Standardtarif"
      End Select ' Case rsFa!fscheinart
     Case 0
      Select Case rsFa!ftarif
        Case 1     ' eigene Behandlung
         rFa(UBound(rFa)).SchGr = "00"
        Case 2     ' Überweisungsfall
         Select Case rsFa!fScheinuntergruppe
          Case 1    ' Selbstausstellung
           rFa(UBound(rFa)).SchGr = "20"
          Case 2    ' Auftragsleistung
           rFa(UBound(rFa)).SchGr = "21"
          Case 4    ' Konsiliaruntersuchung
           rFa(UBound(rFa)).SchGr = "23"
          Case 5    ' Mit- und Weiterbehandlung
           rFa(UBound(rFa)).SchGr = "24"
          Case Else
           rFa(UBound(rFa)).SchGr = "29" ' frei erfunden
         End Select
        Case 3     ' Belegärztlich
' 43 = Notfall 41 = ÄND, 92 = nur Musterwoman
        Case 4     ' Notfall/Vertretung
         Select Case rsFa!fScheinuntergruppe
          Case 1    ' Ärztlicher Notfalldienst
           rFa(UBound(rFa)).SchGr = "41"
          Case 2    ' Urlaubsvertretung
           rFa(UBound(rFa)).SchGr = "42"
          Case 3    ' Notfall
           rFa(UBound(rFa)).SchGr = "43"
          Case 4    ' Notfalldienst mit Taxi
           rFa(UBound(rFa)).SchGr = "44"
          Case 5    ' Notfall-Rettungswagen
           rFa(UBound(rFa)).SchGr = "45"
          Case 6    ' Zentraler Notfalldienst
           rFa(UBound(rFa)).SchGr = "46"
          Case Else
           rFa(UBound(rFa)).SchGr = "49" ' frei erfunden
         End Select
        Case Else
         rFa(UBound(rFa)).SchGr = "99" ' frei erfunden
      End Select ' ftarif
     End Select ' fscheintyp
    Call ParseMemo(rsFa!fm, FaStr(), rFa(), obDebug, rsFa!ueschr)  ' rsFa!fpatnr, rsFa!fsurogat,
'    Call MeStDruck(pNr & " " & rsFa!ueschr, FaStr)
'  For jj = 1 To UBound(rFa)
'   If Not IsNumeric(rFa(jj).Quartal) Or Len(rFa(jj).Quartal) <> 5 Then Stop
'  Next jj
'    If rsfaru = 0 Then
     For j = 0 To UBound(FaStr)
'      If lfdfl = 1 Then
'       Select Case FaStr(j).enr
'        Case "3.2.2.2":        rNa(0).Versichertennummer = Trim$(FaStr(j).Text) ' Versichertennummer, s.o.
'        Case "3.2.2.3.2":      rNa(0).GebDat = CDate(Format$(FaStr(j).Text, "####\.##\.##")), s.o.
'        Case "3.2.2.3.3":      rNa(0).geschlecht = IIf(FaStr(j).Text = "2" Or FaStr(j).Text = "w", "w", IIf(FaStr(j).Text = "1" Or FaStr(j).Text = "m", "m", " ")), s.o.
'        Case "3.2.2.3.6.2":    rNa(0).plz = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.3":    rNa(0).ort = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.4":    rNa(0).Lkz = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.5":    rNa(0).Straße = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.6"
'           rNa(0).Hausnr = FaStr(j).Text, s.o.
'           If rNa(0).Hausnr <> "" Then rNa(0).Straße = rNa(0).Straße & " " & rNa(0).Hausnr, s.o.
'       End Select ' FaStr(0).enr
'      End If ' lfdfl = 1 Then
      
'     Debug.Print FaStr(j).enr, FaStr(j).Text
      Select Case FaStr(j).enr
'      case "2.2":           ' Namenszusatz des Hauptversicherten
'      Case "2.3":           ' Nachname des Hauptversicherten
'      Case "2.4":           ' Vorname  des Hauptversicherten
'      case "2.5":           ' Geburtsdatum des Hauptversicherten mit stzk(
'      case "2.6":           ' Titel des Hauptversicherten
'      case "2.7.3":           ' Straße des Hauptversicherten
'      case "2.7.4":         ' Hausnr.
'      case "2.7.5":         ' PLZ
'      case "2.7.6":         ' Ort
'      case "2.8":           ' Geschlecht: "1" = männlich, "2" = weiblich
'      case "2.9":           ' Vorsatzwort des Hauptversicherten
       Case "3.2.2.3.4.4":    rFa(UBound(rFa)).Nachname = FaStr(j).Text
       Case "3.2.2.3.4.3":    rFa(UBound(rFa)).Vorname = FaStr(j).Text
       Case "3.2.2.4.2": ' gibt es nur auf szn4
           rFa(UBound(rFa)).VschBeg = CDate(Format$(FaStr(j).Text, "##\.##\.####"))
       Case "3.2.2.4.3": ' gibt es nur auf wser
            rFa(UBound(rFa)).VschEnd = CDate(Format$(FaStr(j).Text, "####\.##\.##")) '
'           rFa(UBound(rFa)).VschEnd = CDate(Left$(FaStr(j).Text, 4) & "." & Mid$(FaStr(j).Text, 5, 2) & "." & Mid$(FaStr(j).Text, 7, 2))
        Case "3.2.2.4.4.4": ' ist wieder auf beiden gleich
            rFa(UBound(rFa)).KKasse_2 = FaStr(j).Text
      Case "3.2.2.5.4"        ' KV-Bereich, BDT 3116, wird nicht aus Turbomed übertragen
            rFa(UBound(rFa)).Kasse = FaStr(j).Text ' wird dann noch ergänzt unter 4.5
'      Case "3.2.2.5.5":      ' Rechtskreis, BDT ?, 1 = normal, 9 = Ost, vermutlich nicht in Turbomed
       Case "3.2.2.5.6"
            rFa(UBound(rFa)).Status = FaStr(j).Text
            rNa(0).KVKStatus = FaStr(j).Text
'       Case "3.2.2.5.7" ' ?, kann fehlen, meist 1, selten 2, auch bei gl.Pat.+Versich.(z.B. 59465), weder AbrAr noch GebOr
       Case "3.2.2.5.8"
           rFa(UBound(rFa)).bPerG = Right$(FaStr(j).Text, 1) ' besondere Personengruppe: MO: Kostenträger->Versicherungsdaten,Zusatzinformationen; Turbomed: Verwalten -> Allgemeine Behandlungsfalldaten
       Case "3.2.2.5.9"
           rFa(UBound(rFa)).DMPKnZ = Right$(FaStr(j).Text, 2) ' DMP-Klass: 00=-, 1=T2Dm, 2=Brustkr, 3=KHK, 4=T1Dm, 5=Asthma, 6=COPD, 7=Herzins, 8=Depr, 9=Rückensz, 10=Rheuma, 11=Osteoporose,
       Case "3.2.4":
            rNa(0).eGKSchVer = FaStr(j).Text ' CDM-Version, z.B. 5.1.0, 5.2.0
       Case "3.2.5.2":        rFa(UBound(rFa)).DtlOnlPfg = BDTtoDateTime(Left$(FaStr(j).Text, 14))
       Case "3.2.5.3":        rFa(UBound(rFa)).ErgbdOnlP = Asc(Left$(FaStr(j).Text, 1))
       Case "3.2.5.4":        rFa(UBound(rFa)).ErrorCode = Left$(FaStr(j).Text, 5) ' bis jetzt kein Beispiel
       Case "3.2.5.5":        rFa(UBound(rFa)).PrüfZdFd = FaStr(j).Text
       Case "3.3":            rNa(0).KarGen = FaStr(j).Text ' Kartentyp, 0, 2, 3
       Case "3.4":
            rFa(UBound(rFa)).lVorl = stzk(FaStr(j).Text)
            rFa(UBound(rFa)).ausgst = rFa(UBound(rFa)).lVorl
'            Call VorstellSetz(rFa(UBound(rFa)).lVorl)
       Case "4.2":            rFa(UBound(rFa)).VKNr = FaStr(j).Text
'       Case "4.3":             ' Kostenträgergruppe BDT 2018, in Turbomed nicht in der Falldatei
       Case "4.4":            rFa(UBound(rFa)).KtrAbrB = Trim$(Left$(FaStr(j).Text, 2)) ' BDT 4106 unter 80000 Fällen fast immer 00, sonst 0, 00, 01, 06, 08, 1 und 2
       Case "4.5":            rFa(UBound(rFa)).Kasse = Left$(FaStr(j).Text & Space$(27), 27) & " " & rFa(UBound(rFa)).Kasse
       Case "4.6":            If Trim$(FaStr(j).Text) = "1" Then rFa(UBound(rFa)).UnfFlg = "1"
       Case "4.7":            rFa(UBound(rFa)).AbrGb = Format$(FaStr(j).Text, String(2, "0"))
'                   If FaStr(j).Text <> "7" Then Stop
       Case "4.8":           rFa(UBound(rFa)).PersKreis = Format$(Asc(FaStr(j).Text), "00") ' Personenkreis, BDT 9402 bzw. 4123
       Case "4.10":          rFa(UBound(rFa)).SKtZusatz = FaStr(j).Text ' bei Pat. 53776
       Case "4.13":          ' gültig von, stzk( ' BDT 4150
       Case "4.14":          ' gültig bis, stzk( ' BDT 4151 = 4152
       Case "4.15":          rFa(UBound(rFa)).SktBem = FaStr(j).Text ' Skt-Bemerkung, BDT 4126
       Case "4.21": rFa(UBound(rFa)).VermiArt = Left$(FaStr(j).Text, 1) ' 4.22: Zusatzinfo, 4.23: Vermittlungscode, 4.24: Datum d. Terminvermittulung
           ' dabei (Vermittlungsart 6) auch noch FAmbulStat auf 1 gesetzt (?)
       Case "4.22":          rFa(UBound(rFa)).VermiZusatz = Trim$(FaStr(j).Text)
       Case "4.23":          rFa(UBound(rFa)).VermiCode = Trim$(FaStr(j).Text)
       Case "4.24":          rFa(UBound(rFa)).VermiDatum = stzd(FaStr(j).Text)
       Case "4.25": ' Unfalltag zu 4.6, nicht in Turbomed
'       Case "5"                ' immer Ascii 4, auf szn4 und wser
' falsch:    If Asc(Left$(FaStr(j).Text, 1)) > 31 Then
'             rFa(UBound(rFa)).Kasse = Left$(Left$(rFa(UBound(rFa)).Kasse, 24) & String(24, " "), 24) & FaStr(j).Text
'            End If
       Case "5.2.11":
           rFa(UBound(rFa)).privVers = FaStr(j).Text
       Case "5.22"
           rFa(UBound(rFa)).FaktPers = stzd(FaStr(j).Text)
       Case "5.23"
           rFa(UBound(rFa)).FaktTechn = stzd(FaStr(j).Text)
       Case "5.24"
           rFa(UBound(rFa)).FaktLabor = stzd(FaStr(j).Text)
       Case "5.3" ' ?
           Debug.Print stzd(FaStr(j).Text)
       Case "5.4" ' ? (identisch mit 5.3, nicht Rechnungsbetrag)
           Debug.Print stzd(FaStr(j).Text)
'       Case "5.26" ' Punktwert 5,82873
'           Debug.Print stzd(FaStr(j).Text)
'       Case "5.27" ' Punktwert 5,82873
'           Debug.Print stzd(FaStr(j).Text)
'       Case "5.28" ' Punktwert 5,82873
'           Debug.Print stzd(FaStr(j).Text)
'       Case "6"                ' immer Ascii 4, auf szn4 und wser
'       Case "7"                ' immer Ascii 4, auf szn4 und wser
'       Case "8"                ' immer Ascii 4, auf szn4 und wser
'       Case "9.2"              ' immer "0", auf szn4 und wser
'       Case "9.8"              ' immer "0", auf szn4 und wser
       Case "10.2"
           rFa(UBound(rFa)).ÜbWVBSNR = FaStr(j).Text
       Case "10.3"
           rFa(UBound(rFa)).ÜWZiel = FaStr(j).Text
       Case "10.4"
           rFa(UBound(rFa)).ÜbwLANR = FaStr(j).Text
           rFa(UBound(rFa)).ÜbWVLANR = FaStr(j).Text
       Case "10.5" ' <BSNR>#<LANR>#<epraxis.fsurogat>#<earzt.fsurogat>
             ' Feld kommt aber auf wser nur einmal vor und auf szn4 nur viermal
       Case "10.6"
            Dim han$, hav$, hatit$, hapos%
            hapos = InStr(FaStr(j).Text, "med.")
            If hapos <> 0 Then han = Trim$(Mid$(FaStr(j).Text, hapos + 4)): hatit = Left$(FaStr(j).Text, hapos + 4) Else han = FaStr(j).Text: hatit = ""
            hapos = InStr(han, ",")
            If hapos <> 0 Then hav = Mid$(han, hapos + 1): han = Left$(han, hapos - 1)
            rFa(UBound(rFa)).ÜWNaN = han
            rFa(UBound(rFa)).ÜWTit = hatit
            rFa(UBound(rFa)).ÜWVor = hav
       Case "10.7"
           Debug.Print stzk(FaStr(j).Text)
       Case "10.8": rFa(UBound(rFa)).Verdacht = FaStr(j).Text
       Case "10.9": rFa(UBound(rFa)).Auftrag = FaStr(j).Text
'       Case "10.10"   ' kommt selten vor, immer "0"
       Case "10.13": rFa(UBound(rFa)).Befund = FaStr(j).Text
'       Case "13"      ' immer  Ascii 2
'       Case "14"      ' kommt selten vor, Ascii 1 oder 4
'       Case "15"      ' kommt selten vor, Datum in der Nähe der Behandlungsdaten stzk(, Bedeutung konnte ich nicht ermitteln
'       Case "16"      ' kommt selten vor, Ascii 1 oder 4
'       Case "17":      rFa(UBound(rFa)).aktZeit = stzk(FaStr(j).Text) ' wohl Importdatum als stzk
' bPerG, DMPKnZ
'        Exit For
      End Select ' Case FaStr(j).enr
     Next j
' nicht übertragen:
     Dim rhar As New ADODB.Recordset
     Dim aktkvnr$
     myFrag rhar, "SELECT farztnralt kvnr FROM earzt a LEFT JOIN epraxis p ON p.FSurogat = a.FExtpraxisnr WHERE fbetriebsnr='" & rFa(UBound(rFa)).ÜbWVBSNR & "' AND (farztnr='" & rFa(UBound(rFa)).ÜbwLANR & "' OR " & IIf(rFa(UBound(rFa)).ÜbwLANR = "", "TRUE", "FALSE") & ")", adOpenStatic, MOCon
     If rhar Is Nothing Then
'      Stop
     Else
      If Not rhar.BOF Then
       aktkvnr = Format$(rhar.Fields(0), String(10, "0"))
       rFa(UBound(rFa)).ÜbWVKVNR = aktkvnr
       rFa(UBound(rFa)).Übwr = aktkvnr
       rFa(UBound(rFa)).ÜWNNr = aktkvnr
       Dim uewv As New ADODB.Recordset
       For j = 1 To 2
        myFrag uewv, "SELECT ID,Nachname,Vorname,Titel FROM ueberwvon WHERE kvnr=" & aktkvnr, adOpenStatic, DBCn
        If Not uewv Is Nothing Then
         If uewv.BOF Then
          sql = "INSERT INTO ueberwvon(KVNr,Titel,Vorname,Zusatz,Nachname) VALUES(" & aktkvnr & ",'" & hatit & "','" & hav & "','','" & han & "')"
          InsKorr DBCn, DBCnS, sql, rAf
         Else ' uewv.BOF Then
          rFa(UBound(rFa)).üwvid = uewv.Fields(0)
          rFa(UBound(rFa)).ÜWNaN = uewv.Fields(1)
          rFa(UBound(rFa)).ÜWVor = uewv.Fields(2)
          rFa(UBound(rFa)).ÜWTit = uewv.Fields(3)
          Exit For
         End If ' uewv.BOF Then else
        End If ' Not uewv Is Nothing Then
       Next j
      End If ' Not rhar.BOF Then
     End If ' Not rhar Is Nothing Then
'     rsfaru = 1
'    End If ' rsfaru = 0 Then
    rsFa.MoveNext
   Loop ' While Not rsFa.EOF
  End If ' Not rsFa.BOF Then
  lfdfl = 0 ' für nä Pat
  
  Dim rsha As New ADODB.Recordset, rslue  As New ADODB.Recordset
  ' -34 Überweiser, -40 Hausarzt, -32 Arzt
  rsha.Open "SELECT FArztnralt, FAdresse, farztgruppe, FNachname, FVorname, FRelationtyp FROM patrelation r LEFT JOIN earzt a ON r.freferenzid = a.fsurogat AND freferenztyp=2 LEFT JOIN epraxis p ON a.FExtpraxisnr = p.fsurogat WHERE fpatid=" & pNr & " AND FRelationtyp IN (-34,-40,-32)", MOCon, adOpenStatic, adLockReadOnly
  If Not rsha.BOF Then
   Dim haru%, KVNr$
   haru = 0
   KVNr = ""
   Do While Not rsha.EOF
    haru = haru + 1
    If rsha!FArztnralt <> "" Then
     KVNr = rsha!FArztnralt
    ElseIf Not IsNull(rsha!Fadresse) Then ' mit COALESCE kommt trotzdem eine Fehlermeldung raus
'    Set rslue = myEFrag(, , DBCn)
     Dim NN$, VN$, Adr$
     NN = rsha!FNachname
     VN = rsha!fVorname
     Adr = REPLACE$(rsha!Fadresse, "'", "")
     On Error Resume Next
     rslue.Open "SELECT kvnr FROM aktlue a WHERE nameo ='" & NN & "' AND vno='" & VN & "' AND '" & Adr & "' LIKE CONCAT('%',a.strasse,'%')", DBCn, adOpenStatic, adLockReadOnly
     If Not rslue.BOF Then
      KVNr = rslue!KVNr
     End If
     On Error GoTo fehler
    End If
    If KVNr = "" Then
     Set rslue = myEFrag("SELECT kvnr FROM aktlue a WHERE nameo ='" & rsha!FNachname & "' AND vno='" & rsha!fVorname & "' AND fachgruppe='" & rsha!farztgruppe & "'", , DBCn)
     If Not rslue.BOF Then
      KVNr = rslue!KVNr
     Else
      Set rslue = myEFrag("SELECT kvnr FROM aktlue a WHERE nameo ='" & rsha!FNachname & "' AND vno='" & rsha!fVorname & "'", , DBCn)
      If Not rslue.BOF Then
       KVNr = rslue!KVNr
      End If
     End If
    End If
    Select Case haru
     Case 1: rNa(0).KVNr = KVNr
        rNa(0).getHA0 = IIf(KVNr = "", 0, KVNr)
        rNa(0).fnHA0 = "(" & IIf(rsha!frelationtyp = -40 Or rsha!frelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA0 = rNa(0).fnHA0 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA0 = rNa(0).fnHA0 & ")"
     Case 2: rNa(0).KVNr2 = KVNr
        rNa(0).getHA1 = KVNr
        rNa(0).fnHA1 = "(" & IIf(rsha!frelationtyp = -40 Or rsha!frelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA1 = rNa(0).fnHA1 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA1 = rNa(0).fnHA1 & ")"
     Case 3: rNa(0).KVNr3 = KVNr
        rNa(0).getHA2 = KVNr
        rNa(0).fnHA2 = "(" & IIf(rsha!frelationtyp = -40 Or rsha!frelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA2 = rNa(0).fnHA2 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA2 = rNa(0).fnHA2 & ")"
     Case 4: rNa(0).KVNr4 = KVNr
    End Select
    rsha.MoveNext
   Loop
  End If ' Not rsha.BOF Then
  
  For j = 0 To UBound(NaStr)
   If NaStr(j).enr Like "21.*" And NaStr(j).enr <> "21.1" Then
    ReDim Preserve rSw(UBound(rSw) + 1)
    rSw(UBound(rSw)).Pat_id = rNa(0).Pat_id
    rSw(UBound(rSw)).FormTitel = "ssd"
    rSw(UBound(rSw)).vorET = stzk(NaStr(j).Text)
    rSw(UBound(rSw)).lR = rSw(UBound(rSw)).vorET - 280
    rSw(UBound(rSw)).MB = rSw(UBound(rSw)).vorET - 42
    rSw(UBound(rSw)).aktZeit = aktZeit
    For jj = 1 To UBound(rFa)
     If rFa(jj).BhFB < rSw(UBound(rSw)).vorET And rFa(jj).BhFE1 > rSw(UBound(rSw)).vorET - 268 Then
      rFa(jj).vorET = rSw(UBound(rSw)).vorET
      rFa(jj).letzteRegel = rSw(UBound(rSw)).lR
     End If
    Next jj
   End If
  Next j
 
 ' KVnr: fpatrelation, dort fpatid= fpatnr, freferenztyp 2 = Hausarzt (0=Arbeitgeber), freferenzid = earzt.fsurogat,
 ' dort FExtpraxisnr = epraxis.fsurogat
  
  
' Set rsFa = Nothing
' Set rsFa = Nothing ' wirkt witzigerweise erst beim zweiten Mal (!?)
  Call rFaDump
  Call alleSpeichern(lies, vonMo:=True)
  syscmd 4, "Fertig mit doPatvonMO " & pNr & " auf '" & MOCon.Properties("Server Name") & "'"
 Else
  syscmd 4, "doPatvonMO: " & pNr & " nicht in patstamm auf '" & MOCon.Properties("Server Name") & "' gefunden!"
 End If '  If Not rsNa.BOF Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPatvonMO/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doPatvonMO(pNr&, pid&)

' kommt nirgends vor
Public Function suchfi(pNr&, fI$, Optional obszn4%)
 Dim altt$, gefu%, i&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 If obszn4 Then
  MOCon.Open MOsStr
 Else
  MOCon.Open MOCStr
 End If
 rst.Open "SELECT c.TABLE_NAME tn, c.COLUMN_NAME cn -- , a.column_name cn" & vbCrLf & _
          "FROM information_schema.columns c" & vbCrLf & _
          "-- LEFT JOIN information_schema.columns a ON c.TABLE_CATALOG=a.TABLE_CATALOG AND c.TABLE_SCHEMA=a.TABLE_SCHEMA AND c.table_name=a.TABLE_NAME" & vbCrLf & _
          "WHERE c.table_catalog='def' AND c.table_schema='medoff' AND (c.COLUMN_NAME IN ('fpatnr','patnr','fpatid') OR (c.TABLE_NAME='patstamm' AND c.COLUMN_NAME='fsurogat')) GROUP BY c.table_name ORDER BY c.table_name", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
  If rst!Tn <> altt Then
   Set rsu = Nothing
   rsu.Open "SELECT * FROM `" & rst!Tn & "` WHERE `" & rst!Cn & "`=" & pNr, MOCon, adOpenStatic, adLockReadOnly
   gefu = 0
   Do While Not rsu.EOF
    For i = 0 To rsu.Fields.COUNT - 1
     If Not (IsNull(rsu.Fields(i))) Then
'      If CStr(rsu.Fields(i)) = fI Then
      If InStrB(LCase$(rsu.Fields(i)), LCase$(fI)) <> 0 Then
       Debug.Print rst!Tn, rsu.Fields(i).name, rsu.Fields(i)
       gefu = True
      End If
     End If
    Next i
    If gefu Then GoTo weiter
   rsu.MoveNext
  Loop ' While Not rsu.EOF
weiter:
  End If
  rst.MoveNext
 Loop ' While Not rst.EOF
 syscmd 4, "Fertig mit suchfi " & pNr& & " " & fI$
End Function ' suchfi(pNr&, fI$)


' kommt nirgends vor
Public Function suchal(fI$, Optional NotObRlike%)
 Dim altt$, j&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 MOCon.Open MOCStr
#If True Then
 rst.Open "SELECT TABLE_NAME tn, GROUP_CONCAT(CONCAT('CAST(',column_name,' AS CHAR)" & IIf(Not NotObRlike, " RLIKE ", "=") & "''" & fI & "''') SEPARATOR ' OR ') cn FROM information_schema.columns c WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_NAME NOT RLIKE 'fsurogat' group by table_name" _
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
 "and not exists (select 0 from information_schema.columns where  table_catalog='def' AND table_schema='medoff' AND TABLE_NAME =t.table_name and column_Name in ('fpatnr','patnr','fpatid'))" & vbCrLf & _
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
 syscmd 4, "Fertig mit suchfal " & fI & " " & NotObRlike%
End Function ' suchal


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
    Key zuloe(PatNr, enr)
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
  IF aktmax BETWEEN 0 AND gl+2 AND -- wenn die angegebene Länge vertretbar
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
 SELECT PatNr, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ascii(mid(TEXT,2))b2, ASCII(mid(TEXT,3))b3, ascii(mid(TEXT,4))b4,
 CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmpatstamm WHERE pnr=0 OR patnr=pnr ORDER BY znr;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmpatfall`(
    IN `pnr` INT
)
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FMemo aus patfall in die Tabelle tmpfmemo aus'
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
    Key zuloe(PatNr, fsur, enr)
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
  IF aktmax BETWEEN 0 AND gl+2 AND -- wenn die angegebene Länge vertretbar
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
 SELECT PatNr, FSur, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
  CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmpatfall WHERE pnr=0 OR patnr=pnr ORDER BY fsur DESC,znr;
End

- Versuch mit Cursor:
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(1024);
DECLARE vsur,vpatnr INT(11);
DECLARE vmemo LONGBLOB;
DECLARE done TINYINT;
DECLARE cpf CURSOR FOR
SELECT fsurogat fsur, fpatnr, fmemo FROM patfall WHERE (pnr=0 OR fpatnr=pnr) AND fmemo IS NOT NULL ORDER BY fvon DESC;
DECLARE CONTINUE HANDLER FOR NOT FOUND SET done=TRUE;
DECLARE EXIT HANDLER FOR SQLEXCEPTION
 Begin
   GET DIAGNOSTICS CONDITION 1 @sqlstat = RETURNED_SQLSTATE, @errno = MYSQL_ERRNO, @etext = MESSAGE_TEXT;
   SELECT @sqlstat, @errno, @etext;
 END;
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
    Key zuloe(PatNr, fsur, enr)
);
DELETE FROM tmpmpatfall WHERE pnr=0 OR patnr=pnr;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
SET done=0;
OPEN cpf;
laba: Loop
 FETCH NEXT FROM cpf INTO vsur, vpatnr, vmemo;
-- SELECT pnr, done,vsur,vpatnr,vmemo;
-- leave laba;
 IF done THEN LEAVE laba; ELSE
-- laba: FOR r IN (SELECT fsurogat fsur, fpatnr, fmemo FROM patfall WHERE (pnr=0 OR fpatnr=pnr) AND fmemo IS NOT NULL ORDER BY fvon DESC) DO
 SET pos=1;
 SET ie=0;
 TRUNCATE tmpeb;
labt:  Loop
  SET obdruck=0;
  SET tlen=CONV(HEX(CONCAT(MID(vmemo,pos+1,1),MID(vmemo,pos,1))),16,10);
--  SELECT length(vmemo),"pos: ",pos,CONV(HEX(MID(vmemo,pos+1,1)),16,10),CONV(HEX(MID(vmemo,1,1)),16,10),CONV(HEX(CONCAT(MID(vmemo,pos+1,1),MID(vmemo,pos,1))),16,10);
--  leave laba;
  IF pos=1 then SET gl=tlen; SET MAX=gl; END IF;
  SET altmax=MAX;
  SET aktMAX=tlen+pos+1; -- aktuell angegebene Länge aus dem und dem nä Byte
  IF aktmax BETWEEN 0 AND gl+2 AND -- wenn die angegebene Länge vertretbar
    (( CONV(HEX(MID(vmemo,aktMax,1)),16,10)=0 AND -- und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen
     aktmax<=MAX) OR pos>MAX) then -- wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde
    SET altie=ie; -- letzte Ebene
    IF aktmax<=MAX then SET ie=ie+1; END IF; -- im ersten Fall wird die Ebene erhöht
    IF pos>MAX then -- im zweiten Fall wird zurückgegriffen
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmpatfall WHERE patnr=vpatnr and fsur=vsur AND mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmpatfall WHERE patnr=vpatnr and fsur=vsur AND mx>=aktMAX);
     END IF;
    IF ie<=altie then -- wenn die Ebene nicht erhöht worden ist
     IF ie<altie then -- wenn sie vielmehr reduziert wurde
      DELETE FROM tmpeb WHERE nr>ie; -- dann werden die höheren Einträge wieder gelöscht
     END IF;
     UPDATE tmpeb SET wert=wert+1 WHERE nr=ie; -- dann wird die Zählung der akt. Ebene erhöht
    Else
     INSERT INTO tmpeb(nr,wert) VALUES(ie,1); -- sonst muss ein neuer Eintrag für die hohe Ebene erstellt werden
    END IF;
    SET MAX=aktmax; -- neue vorbestehende Länge
    SET obdruck=1; -- und drucken
  END IF;
  IF obdruck=1
--   OR pos=1 -- zum Debuggen der ersten Gesamt-Zeile
   OR TRUE -- zum Debuggen aller Zeilen
   THEN
    SET txt=MID(vmemo,pos+2,tlen);
    If txt <> Repeat(Chr(0), tlen) Then
     INSERT INTO tmpmpatfall(patnr,fsur,znr,mx,ebn,enr,TEXT) VALUES(vpatnr,vsur,pos,max,ie,
       (SELECT GROUP_CONCAT(wert ORDER BY nr SEPARATOR '.') FROM tmpeb),
      CONCAT(MID(vmemo,pos,1),'|',HEX(MID(vmemo,pos,1)),'|',CONV(HEX(MID(vmemo,pos,1)),16,10),'|',CONV(HEX(CONCAT(MID(vmemo,pos+1,1),MID(vmemo,pos,1))),16,10),'|','aktMax:',aktmax, '|','Max:',altMAX,'|','Laenge: ',LENGTH(txt),'|','Endbyte:',COALESCE(CONV(HEX(MID(vmemo,aktMax,1)),16,10),'-'),
     CONCAT(IF(obdruck<>1,'- ',''),'"',
      txt
      ,'"'))
       );
     ELSE SET pos=pos+tlen; -- 0-Strings nicht auffieseln
     END IF;
  END IF; -- obdruck=1
  SET pos=pos+1;
  IF pos>=gl then LEAVE labt; END IF;
 END loop labt;
 END IF; -- done ELSE
END LOOP laba;
CLOSE cpf;
 -- folende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
-- DELETE FROM tmpmpatfall WHERE (pnr=0 OR patnr=pnr) AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpmpatfall i WHERE i.patnr=tmpmpatfall.patnr AND i.fsur=tmpmpatfall.fsur AND INSTR(i.enr,CONCAT(tmpmpatfall.enr,'.'))=1 AND i.enr<>tmpmpatfall.enr);
-- folgende Variante bringt, aus der Prozedur aufgerufen, mariadb zum Absturz:
-- DELETE a FROM tmpmpatfall a LEFT JOIN (SELECT LEAD(enr) OVER(PARTITION BY patnr,fsur ORDER BY znr) nenr, patnr,fsur,znr FROM tmpmpatfall) i USING (patnr,fsur,znr) WHERE INSTR(i.nenr,a.enr)=1 AND i.nenr<>a.enr;
-- COMMIT;
-- LEAVE tp;
 SELECT PatNr, FSur, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
  CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmpatfall WHERE pnr=0 OR patnr=pnr ORDER BY fsur DESC,znr;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `proczerlegmemo`(
    IN `Feld` BLOB
)
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'gibt die ascii-Werte eines Strings aus'
Begin
DECLARE pos INTEGER DEFAULT 0;
DECLARE erg BLOB DEFAULT '';
schl: Loop
 SET pos=pos+1;
 IF pos>LENGTH(feld) THEN LEAVE schl; END IF;
 SET erg=CONCAT(erg, ' ',ASCII(MID(feld,pos,1)));
-- SELECT ASCII(MID(feld,pos,1));
END loop schl;
SELECT erg;
End

#End If
