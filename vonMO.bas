Attribute VB_Name = "vonMo"
Option Explicit
Option Compare Text
Const Fakt& = 256
' Public Const pidoffs& = 100000
' Public Const obszn4% = True
'Public Const MoWServ$ = "wser"
'Public Const MoSzn4$ = "szn4"
Public Const MOAnfStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;server="
Public Const gesnamegmo$ = "COALESCE(CONCAT(FTitel,IF(FTitel<>'',' ',''),FVorname,IF(FVorname<>'',' ',''),FNamenszusatz,IF(FNamenszusatz<>'',' ',''),IF(FNamensvorsatz<>FNamenszusatz,FNamensvorsatz,''),IF(IF(FNamensvorsatz<>FNamenszusatz,FNamensvorsatz,'')<>'',' ',''),FNachname,', *',DATE_FORMAT(FGeburtsdatum,'%e.%c.%Y')),'') gesnameg"
' "COALESCE((SELECT CONCAT(FTitel,IF(FTitel<>'',' ',''),FVorname,IF(FVorname<>'',' ',''),FNamenszusatz,IF(FNamenszusatz<>'',' ',''),IF(FNamensvorsatz<>FNamenszusatz,FNamensvorsatz,''),IF(IF(FNamensvorsatz<>FNamenszusatz,FNamensvorsatz,'')<>'',' ',''),FNachname,', *',DATE_FORMAT(FGeburtsdatum,'%e.%c.%Y')) gesnameg"
Public MOCon As New ADODB.Connection
' Const parsemotxt$ = "v:\Parsememo31.txt"
#If mitmestdruck Then
Const mestausg$ = "v:\mestr.txt"
#End If
'Public Const MOCStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSer & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
'Public Const MOsStr$ = "DRIVER={MySQL ODBC 8.0 Unicode Driver};server=" & MoSzn & ";option=0;database=medoff;uid=medoff;pwd=medoff;port=2020;"
' Public MOCon As New ADODB.Connection
  
Type memoType ' zur Übertragung der FMemo-Felder aus Medical Office
' patnr As Long
' FSur As Long
 znr As Long
 mx As Long
 ebn As Integer
 ENr As String
 Text As String
 endse As String
 endsz As String
End Type ' memoType

Type ebType
 nr As Integer
 Wert As Integer
End Type ' ebtype

Type Abrtyp ' Tabelle abrechner
 fS As Integer
 BSNR As Long
End Type

Declare Sub CopyMemoryPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination&, ByVal Sourc&, ByVal length&)
Dim aru&

Public Function explor()
Const Filename$ = "P:\dok\69926\"
Const wart& = 100
Dim cmd$
On Error GoTo fehler
cmd = "C:\Windows\explorer.exe /select,""" & Filename & """"
cmd = "C:\Windows\explorer.exe """ & Filename & """"
Dim objShell As Object
Set objShell = CreateObject("WScript.Shell")
With objShell
    .rUn cmd, 1, False
    Sleep 1000
'    .Sendkeys "% x"
'    Sleep wart
    DoEvents
    .Sendkeys "%"
    Sleep wart
    DoEvents
    .Sendkeys "a"
'    Sleep wart
    .Sendkeys "o{Down}{Enter}"
End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in exp/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function

' wird nicht verwendet
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
End Sub ' testeb()

' in ParseMemo (2x)
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
  ReDim Preserve eb(IIf(UBound(eb) - loezahl > 0, UBound(eb) - loezahl, 0))
  loezahl = 0
 End If
End Sub ' ebaltloe

' in ParseMemo (2x)
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

' wird nicht verwendet
Public Sub ebloe(ByRef AryVar() As ebType, ByVal RemoveWhich&, obloe%)
    '// The size of the array elements
    '// In the case of string arrays, they are
    '// simply 32 bit pointers to BSTR's.
    Dim ByteLen As Byte
    Static loezahl%
    ' byteLen = 4   '// String pointers are 4 bytes
    '// The copymemory operation is not necessary unless
    '// we are working with an array element that is not
    '// at the end of the array
    If RemoveWhich >= 0 And RemoveWhich < UBound(AryVar) Then
        ByteLen = VarPtr(AryVar(1)) - VarPtr(AryVar(0))
        '// Copy the block of string pointers starting at the position after the removed item back one spot.
        CopyMemoryPtr ByVal VarPtr(AryVar(RemoveWhich)), ByVal VarPtr(AryVar(RemoveWhich + 1)), (ByteLen) * (UBound(AryVar) - RemoveWhich)
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

' wird nicht verwendet
Public Sub memoloe(ByRef AryVar() As memoType, ByVal RemoveWhich&, obloe%)
    '// The size of the array elements
    '// In the case of string arrays, they are
    '// simply 32 bit pointers to BSTR's.
    Dim ByteLen As Byte
    Static loezahl%
    ' byteLen = 4   '// String pointers are 4 bytes
    '// The copymemory operation is not necessary unless
    '// we are working with an array element that is not
    '// at the end of the array
    If RemoveWhich >= 0 And RemoveWhich < UBound(AryVar) Then
        ByteLen = VarPtr(AryVar(1)) - VarPtr(AryVar(0))
        '// Copy the block of string pointers starting at the position after the removed item back one spot.
        CopyMemoryPtr ByVal VarPtr(AryVar(RemoveWhich)), ByVal VarPtr(AryVar(RemoveWhich + 1)), (ByteLen) * (UBound(AryVar) - RemoveWhich)
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
    
' wird nicht verwendet
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

' in ParseMemo
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

' in zeigmosystem (5x), doPatvonMo (4x), MoAusgeb
' BLOB-Felder aus Medical Office parsen
Public Function ParseMemo(FMemo$, MeStr() As memoType, Optional obDebug%, Optional ÜSchr$) ' pNr&, FSur&,
 Dim jj&, iru&
 Dim gl&, pos&, MAX&, aktmax&, altmax&, tlen&, ie&, altie&, mznr&, i&
 Dim obDruck%, txt$, ebS$
 Dim eb() As ebType
 On Error GoTo fehler
 aru = aru + 1
 Erase MeStr
 If FMemo <> "" Then
  pos = 1
  ie = 0
  If obDebug Then
'   Open parsemotxt For Append As #255
   Open "\\linux1\daten\down\" & ÜSchr & ".txt" For Output As #255
   Print #255, ÜSchr & ":"
  End If ' obDebug
  iru = 0
  Do ' Schleife labt
   iru = iru + 1
   obDruck = 0
   On Error Resume Next
   tlen = Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1))
   If Err.Number <> 0 Then tlen = 0
   On Error GoTo fehler
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
     If aktmax <= MAX Then
     ' zuvor:   If (Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Then obnaechst = true
      If aktmax > Len(FMemo) Then
      ElseIf Asc(Mid$(FMemo, aktmax, 1)) = 0 Then
       obnaechst = True
      End If
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
     MeStr(UBound(MeStr)).ENr = ebS
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
  ' folgende Zeilen zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
  If SafeArrayGetDim(MeStr) <> 0 Then
   For i = UBound(MeStr) - 1 To 0 Step -1
    If InStr(MeStr(i + 1).ENr, MeStr(i).ENr & ".") = 1 And MeStr(i + 1).ENr <> MeStr(i).ENr Then ' MeStr(i).patnr = MeStr(i + 1).patnr And MeStr(i).FSur = MeStr(i + 1).FSur And
     Call memoaltloe(MeStr, i, False)
    End If
   Next i
   Call memoaltloe(MeStr, -1, True)
   For i = 0 To UBound(MeStr)
    If MeStr(i).znr <= Len(FMemo) Then
     MeStr(i).endse = Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     MeStr(i).endse = "!: " & MeStr(i).znr & ">" & Len(FMemo)
    End If
    If MeStr(i).znr <= Len(FMemo) - 1 Then
     MeStr(i).endsz = Asc(Mid$(FMemo, MeStr(i).znr + 1, 1)) * 256& + Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     MeStr(i).endsz = "!: " & MeStr(i).znr - 1 & ">" & Len(FMemo)
    End If
    If obDebug Then
     Print #255, MeStr(i).znr & "|" & MeStr(i).mx & "|" & MeStr(i).ebn & "|" & MeStr(i).ENr & "|" & IIf(MeStr(i).endse <> "10", Mid$(FMemo, MeStr(i).znr, 1), "") & "|" & MeStr(i).endse & "|" & MeStr(i).endsz & "| Laenge: " & Len(MeStr(i).Text) & "|" & IIf(Right$(MeStr(i).Text, 1) = Chr$(10), Left$(MeStr(i).Text, IIf(Len(MeStr(i).Text) = 0, 1, Len(MeStr(i).Text)) - 1), MeStr(i).Text)
    End If ' obDebug
   Next i
  End If ' SafeArrayGetDim(MeStr)
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

Public Function testwm()
 Dim MeStr() As memoType
 Call WechsMemo("patfall", 14327, "FMemo", "3.2.2.4.2", "20140101", 1, MeStr())
End Function ' testwm()

' wird nirgends aufgerufen
Public Function fbumdreh() ' Fallbeginnumdreh
 Dim sql$, rsMO As New ADODB.Recordset
 Dim MeStr() As memoType
 Call MOConInit(, "Fallbeginnumdreh")
 sql = "SELECT fsurogat f FROM patfall WHERE fpatnr=68012"
 rsMO.Open sql, MOCon, adOpenStatic, adLockReadOnly
 If Not rsMO.BOF Then
  Do While Not rsMO.EOF
   Call WechsMemo("patfall", rsMO!F, "FMemo", "3.2.2.4.2", "", 1, MeStr(), , , True)
   rsMO.MoveNext
  Loop ' While Not rsMO.EOF
 End If ' Not rsMO.BOF Then
End Function ' fbumdreh()

' in BLOB-Feldern aus Medical Office Teile korrigieren
' art: 0=Zahl, 1=String, 2=Datum
Public Function WechsMemo(TabName$, snr&, mfeld$, anENr$, neu$, art$, MeStr() As memoType, Optional obDebug%, Optional ÜSchr$, Optional obUmdreh%) ' pNr&, FSur&,
 Dim aktenr$, neus$, sqls$, FMemo$, rsMO As New ADODB.Recordset
 Dim jj&, iru&
 Dim gl&, pos&, MAX&, aktmax&, altmax&, tlen&, ie&, altie&, reppos&, nlen&, mznr%, i&
 Dim obDruck%, txt$
 Dim eb() As ebType
 Dim rAf&
 On Error GoTo fehler
 aru = aru + 1
 Erase MeStr ' entspricht tmpmwechs
 If rsMO Is Nothing Then Call MOConInit(, "WechsMemo(" & TabName & "," & snr & "," & mfeld & "," & anENr & "," & neu & "," & art)
 sqls = "SELECT CONVERT(`" & mfeld & "` USING latin1)M FROM `" & TabName & "` WHERE fsurogat=" & snr
 rsMO.Open sqls, MOCon, adOpenStatic, adLockReadOnly
 If Not rsMO.BOF Then
 FMemo = rsMO.Fields(0)
' Debug.Print FMemo
 If FMemo <> "" Then
  pos = 1
  ie = 0
  If obDebug Then
'   Open parsemotxt For Append As #255
   Open "\\linux1\daten\down\" & ÜSchr & ".txt" For Output As #255
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
     If aktmax <= MAX Then
     ' zuvor:   If (Asc(Mid$(FMemo, aktmax, 1)) = 0 And aktmax <= MAX) Then obnaechst = true
      If aktmax > Len(FMemo) Then
      ElseIf Asc(Mid$(FMemo, aktmax, 1)) = 0 Then
       obnaechst = True
      End If
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
     aktenr = ""
     Dim eb2() As ebType
     If SafeArrayGetDim(eb) <> 0 Then
      eb2 = ebQuickSort(eb)
      For i = 0 To UBound(eb)
       aktenr = aktenr & eb(i).Wert & "."
      Next i
      If Right$(aktenr, 1) = "." Then aktenr = Left$(aktenr, Len(aktenr) - 1)
     End If ' SafeArrayGetDim(eb) <> 0 Then
     If aktenr = anENr Then
      Debug.Print "txt: '" & Left$(txt, 8) & "'"
      reppos = pos
      If obUmdreh Then
        If Mid$(txt, 5, 1) = "2" Or Mid$(txt, 5, 2) = "19" Then
         neus = Mid$(txt, 5, 4) & Left$(txt, 4)
'         Debug.Print neu
        Else ' Mid$(
         neus = txt
        End If ' Mid$
        nlen = 8
      Else ' obUmdreh Then
       Select Case art
        Case 0 ' Zahl
            nlen = 1
            neus = "LEFT(CHR(" & neu & ")," & nlen & ")"
            WechsMemo = 1
        Case 1 ' String
            nlen = Len(neu)
            neus = "LEFT('" & neu & "'," & nlen & ")"
            WechsMemo = 1
        Case 2 ' Datum
            nlen = 5
            neus = "CHAR(YEAR(" & neu & ")MOD 256,YEAR(" & neu & ")DIV 256,MONTH(" & neu & "),DAY(" & neu & "),0)"
            WechsMemo = 1
       End Select
      End If ' obUmdreh Then Else
      If Left$(neus, nlen) <> Left$(txt, nlen) Then
       sqls = "UPDATE `" & TabName & "` SET `" & mfeld & "`=CONCAT(LEFT(`" & mfeld & "`," & reppos & "+1)," & neus & ",MID(`" & mfeld & "`," & reppos & "+2+" & nlen & ")) WHERE fsurogat=" & snr
       Debug.Print sqls
       Call MOCon.Execute(sqls, rAf)
       Debug.Print rAf & " Datensätze geändert"
       Exit Do
      End If ' Left$(neu,
     End If ' aktenr = anENr Then
     If SafeArrayGetDim(MeStr) = 0 Then
       ReDim MeStr(0)
     Else ' SafeArrayGetDim(MeStr) = 0 Then
       ' das Folgende geht leider nicht, weil dann oben eb() falsch befüllt wird
       ' zum Debuggen das Folgende mit "or True" ergänzen
'       If Not (InStr(aktenr, MeStr(UBound(MeStr)).enr & ".") = 1 And aktenr <> MeStr(UBound(MeStr)).enr) Then
        ReDim Preserve MeStr(UBound(MeStr) + 1)
'       End If
     End If ' SafeArrayGetDim(MeStr) = 0 Then else
'     MeStr(UBound(MeStr)).patnr = pNr
'     MeStr(UBound(MeStr)).FSur = FSur
     MeStr(UBound(MeStr)).znr = pos
     MeStr(UBound(MeStr)).mx = MAX
     MeStr(UBound(MeStr)).ebn = ie
     MeStr(UBound(MeStr)).ENr = aktenr
     If Asc(Right$(txt, 1)) = 0 Then txt = Left$(txt, Len(txt) - 1)
     MeStr(UBound(MeStr)).Text = txt
'     If obDebug Then
'      Print #255, pos & "|" & MAX & "|" & ie & "|" & aktenr & "|" & Mid$(FMemo, pos, 1) & "|" & Asc(Mid$(FMemo, pos, 1)) & "|" & Asc(Mid$(FMemo, pos + 1, 1)) * 256& + Asc(Mid$(FMemo, pos, 1)) & "| aktMax: " & aktmax & "| altMax: " & altmax & "| Laenge: " & Len(txt) & "| EndByte: " & Asc(Mid$(FMemo, aktmax, 1)) & "|" & txt & vbCrLf
'     End If ' obDebug Then
    Else ' txt <> String$(Chr$(0), tlen) Then ' 0-Strings nicht auffieseln
     pos = pos + tlen
    End If ' txt <> String$(Chr$(0), tlen) Then
   End If ' obdruck = 1 Then ' or pos=1 or true
   pos = pos + 1
   If pos >= gl Then Exit Do
  Loop
  ' folgende Zeilen zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
  If SafeArrayGetDim(MeStr) <> 0 Then
   For i = UBound(MeStr) - 1 To 0 Step -1
    If InStr(MeStr(i + 1).ENr, MeStr(i).ENr & ".") = 1 And MeStr(i + 1).ENr <> MeStr(i).ENr Then ' MeStr(i).patnr = MeStr(i + 1).patnr And MeStr(i).FSur = MeStr(i + 1).FSur And
     Call memoaltloe(MeStr, i, False)
    End If
   Next i
   Call memoaltloe(MeStr, -1, True)
   For i = 0 To UBound(MeStr)
    If MeStr(i).znr <= Len(FMemo) Then
     MeStr(i).endse = Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     MeStr(i).endse = "!: " & MeStr(i).znr & ">" & Len(FMemo)
    End If
    If MeStr(i).znr <= Len(FMemo) - 1 Then
     MeStr(i).endsz = Asc(Mid$(FMemo, MeStr(i).znr + 1, 1)) * 256& + Asc(Mid$(FMemo, MeStr(i).znr, 1))
    Else
     MeStr(i).endsz = "!: " & MeStr(i).znr - 1 & ">" & Len(FMemo)
    End If
    If obDebug Then
     Print #255, MeStr(i).znr & "|" & MeStr(i).mx & "|" & MeStr(i).ebn & "|" & MeStr(i).ENr & "|" & IIf(MeStr(i).endse <> "10", Mid$(FMemo, MeStr(i).znr, 1), "") & "|" & MeStr(i).endse & "|" & MeStr(i).endsz & "| Laenge: " & Len(MeStr(i).Text) & "|" & IIf(Right$(MeStr(i).Text, 1) = Chr$(10), Left$(MeStr(i).Text, IIf(Len(MeStr(i).Text) = 0, 1, Len(MeStr(i).Text)) - 1), MeStr(i).Text)
    End If ' obDebug
   Next i
  End If ' SafeArrayGetDim(MeStr)
  If obDebug Then
   Print #255, vbCrLf
   Close #255
  End If ' obDebug
 End If ' FMEMO<>""
 End If ' not rMO.BOF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WechsMemo/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WechsMemo

#If mitmestdruck Then
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
  MeStr(i).ENr & Space$(19 - Len(MeStr(i).ENr)) & _
  MeStr(i).mx & Space$(5 - Len(MeStr(i).mx)) & _
  MeStr(i).Text
 Next i
 Close #245
End Function ' MestDruck
#End If

'Public Function testfloat()
' Dim d#
' d = 2.1
' Dim dptr&, cvptr&, cptr&
' dptr = VarPtr(d)
'
' Dim c As String * 10
' cvptr = VarPtr(c)
' cptr = StrPtr(c)
' cvptr = VarPtr(c)
' Dim MOCon As New ADODB.Connection
' MOCon.Open MOAnfStr & MoWServ
' Dim rt As New ADODB.Recordset
' rt.Open "select text from tmpmpatfall where patnr=53119 and enr='5.22'", MOCon, adOpenStatic, adLockReadOnly
'' Debug.Print rt.Fields(0)
'' CopyMemoryPtr VarPtr(d), VarPtr(rt.Fields(0)), 8
' CopyMemoryPtr cptr - 4, VarPtr(10), 4
' CopyMemoryPtr cptr, dptr, 8
' CopyMemoryPtr VarPtr(c) + 3, dptr, 6
' Debug.Print c
'End Function

' String zu double; Formel experimentell ermittelt über Datei
Public Function stzd#(s$)
 On Error Resume Next
 stzd = (1 + (Asc(Mid$(s, 7, 1)) Mod 16) / 16 + Asc(Mid$(s, 6, 1)) / 4096 + Asc(Mid$(s, 5, 1)) / 2 ^ 20 + Asc(Mid$(s, 4, 1)) / 2 ^ 28 + Asc(Mid$(s, 3, 1)) / 2 ^ 36 + Asc(Mid$(s, 2, 1)) / 2 ^ 44 + Asc(Mid$(s, 1, 1)) / 2 ^ 52) * 2 ^ (Int(Asc(Mid$(s, 7, 1)) / 16) + 1 + 16 * (Asc(Mid$(s, 8, 1)) - 64))
End Function ' stzd

' String zu Datum ("Kalender"); Formel experimentell ermittelt über Datei
Public Function stzk(s$) As Date
 Dim dats$
 On Error Resume Next
 dats = CStr(Asc(Mid$(s, 4))) & "." & CStr(Asc(Mid$(s, 3))) & "." & CStr(256 * Asc(Mid(s, 2)) + Asc(Mid(s, 1)))
 If IsDate(dats) Then stzk = CDate(dats) Else stzk = 0
End Function ' stzk

' in Übertragung_aus_MO_Click, zeigmosystem, doPatvonMO
Public Function MOConInit(Optional MServ$, Optional anzeige$)
  On Error GoTo fehler
  If MServ = "" Then MServ = Lese.MOServer
  syscmd 4, anzeige & IIf(anzeige = "", "", ", ") & "öffne Verbindung zur Medical Office-Datenbank ..."
  If MOCon Is Nothing Or MOCon = "" Then MOCon.Open MOAnfStr & MServ
  syscmd 4, anzeige & " ..."
'  If MOCon = "" Then
'   If obszn4 Then
'    MOCon.Open MOAnfStr & MoSzn4
'   Else ' obszn4
'    MOCon.Open MOAnfStr & MoWServ
'   End If ' obszn4 else
'  End If ' MOCon = "" Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in moconinit/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' MOConInit

' zum Aufruf im Direktfenster
Public Function zeigmosystem(Optional obszn4%)
 Const obDebug% = False
 Dim kat() As memoType, tKat() As memoType, Abl() As memoType, fAuft() As memoType, FMem() As memoType
 Dim rsMO As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 Dim j&
 Dim EintS As SortierEintr
 Dim EinL As New SortierListe
' MOConInit
 Dim MServ$
 MServ = Lese.MOServer
 If MOCon Is Nothing Or MOCon = "" Then MOCon.Open MOAnfStr & MServ
 sql = "SELECT COALESCE(CONVERT(FKategorieliste USING latin1),'') FKat, COALESCE(CONVERT(Ftextkategorieliste USING latin1),'') Ftk, COALESCE(CONVERT(fAblageliste USING latin1),'') FAb, COALESCE(CONVERT(fAuftragstypenliste USING latin1),'') FAuf, COALESCE(CONVERT(FMemo USING latin1),'') Fm FROM mosystem"
 rsMO.Open sql, MOCon, adOpenStatic, adLockReadOnly
 If Not rsMO.BOF Then
  If rsMO!fkat <> "" Then
   Call ParseMemo(rsMO!fkat, kat(), obDebug, "FMemo von fKategorieliste aus mosystem")
  End If
  If rsMO!ftk <> "" Then
   Call ParseMemo(rsMO!ftk, tKat(), obDebug, "FMemo von fTextKategorieliste aus mosystem")
  End If
  If rsMO!fAb <> "" Then
   Call ParseMemo(rsMO!fAb, Abl(), obDebug, "FMemo von fAblageliste aus mosystem")
  End If
  If rsMO!fAuf <> "" Then
   Call ParseMemo(rsMO!fAuf, fAuft(), obDebug, "FMemo von fAuftragstypenliste aus mosystem")
  End If
  If rsMO!fm <> "" Then
   Call ParseMemo(rsMO!fm, FMem(), obDebug, "FMemo von fMemo aus mosystem")
   For j = 0 To UBound(FMem)
    If FMem(j).ENr Like "*.2" And FMem(j).ENr <> "1.2" Then
     Set EintS = New SortierEintr
     EintS.TypNr = Asc(FMem(j).Text)
     If Len(FMem(j).Text) > 1 Then
      EintS.TypNr = EintS.TypNr + Fakt * Asc(Mid(FMem(j).Text, 2))
     End If
    ElseIf FMem(j).ENr Like "*.3" And FMem(j).ENr <> "1.3" Then
     EintS.art = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.5" And FMem(j).ENr <> "1.5" Then
     EintS.EKür = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.8" And FMem(j).ENr <> "1.8" Then
     EintS.name = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.10" And FMem(j).ENr <> "1.10" Then
     EintS.Kürz = FMem(j).Text
     EinL.sCAdd EintS
    End If
   Next j
  End If
 End If ' Not rsMo.BOF Then
' Suche, z.B.:
'  Set EintS = New SortierEintr
'  EintS.TypNr = 1002
'  Set EintS = EinL.GetItem(EintS)
'  If Not EintS Is Nothing Then
'   Debug.Print EintS.Name
'  End If
 Debug.Print "Fertig mit zeigmosystem"
End Function ' zeigmosystem()

' in Markierungen_Click, doPatvonMO
Public Function doMarkierungen(Optional FPatNr&, Optional nurfrag%)
 Dim rMo As ADODB.Recordset, rDl As ADODB.Recordset, rAf&, mZl&, altPNr&
 Dim meld$
 Call MOConInit(, "doMarkierungen(" & FPatNr & "," & CStr(nurfrag) & ")")
' Dim rNa() As namen
' ReDim rNa(0)
'sql = _
"SELECT FPatnr," & vbCrLf & _
"CONCAT('UPDATE namen SET obk=0,obs=0,obh=0,antikoag=0,dmt1=0,gdm=0,kdm=0,cgm=0,insAnw=0'," & vbCrLf & _
"GROUP_CONCAT(setz SEPARATOR ''),' WHERE pat_id=',FPatnr) sqls" & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT pm.FPatnr, ROW_NUMBER() OVER(ORDER BY fpatnr DESC,FMarkiernr) rg, COUNT(0) OVER() zl, LAST_VALUE(FMarkiernr) OVER(PARTITION BY FPatnr)=FMarkiernr speichern, pm.FMarkiernr,m.FText" & vbCrLf & _
", CASE FText " & vbCrLf & _
"WHEN 'Arzt Kothny' THEN ',obk=1'" & vbCrLf & _
"WHEN 'Arzt Schade' THEN ',obs=1'" & vbCrLf & _
"WHEN 'Arzt Hammerschmidt' THEN ',obh=1'" & vbCrLf & _
"WHEN 'Antikoagulation' THEN ',antikoag=1'" & vbCrLf & _
"WHEN 'Diabetes Typ I' THEN ',dmt1=1'" & vbCrLf & _
"WHEN 'Gdm' THEN ',gdm=1'" & vbCrLf & _
"WHEN 'kein D.m.' THEN ',kdm=1'" & vbCrLf & _
"WHEN 'Libre-Handy' THEN ',cgm=1'" & vbCrLf & _
"WHEN 'Libre Lesegerät' THEN ',cgm=2'" & vbCrLf & _
"WHEN 'Dexcom-Clarity' THEN ',cgm=3'" & vbCrLf & _
"WHEN 'Dexcom Lesegerät' THEN ',cgm=4'" & vbCrLf & _
"WHEN 'Medtronic Simplera' THEN ',cgm=5'" & vbCrLf & _
"WHEN 'Eversense' THEN ',cgm=6'" & vbCrLf & _
"WHEN 'Novopen' THEN ',insanw=1'" & vbCrLf & _
"WHEN 'Accu Chek Spirit Combo' THEN ',insAnw=2'" & vbCrLf & _
"WHEN 'AccuChek Insight Yourloops' THEN ',insAnw=3'" & vbCrLf & _
"WHEN 'Kaleido Glooko' THEN ',insAnw=4'" & vbCrLf & _
"WHEN 'Medtronic 780 pdf Patient' THEN ',insAnw=5'" & vbCrLf
'sql = sql & _
"WHEN 'Omnipod 5 Glooko' THEN ',insAnw=6'" & vbCrLf & _
"WHEN 'Omnipod Dash auslesen' THEN ',insAnw=7'" & vbCrLf & _
"WHEN 'TSlim-Glooko Backoffice auslesen' THEN ',insAnw=8'" & vbCrLf & _
"WHEN 'Ypsopump-Glooko' THEN ',insAnw=9'" & vbCrLf & _
"WHEN 'Dana' THEN ',insAnw=10'" & vbCrLf & _
"ELSE ''" & vbCrLf & _
"END setz" & vbCrLf & _
"FROM patmark pm" & vbCrLf & _
"JOIN markier m ON pm.FMarkiernr=m.FSurogat" & vbCrLf & _
"WHERE " & FPatNr & "=0 OR fpatnr=" & FPatNr & vbCrLf & _
") i" & vbCrLf & _
"GROUP BY fpatnr" & vbCrLf & _
"ORDER BY fpatnr DESC" & vbCrLf

sql = _
"SELECT FPatnr,CONCAT('UPDATE namen SET obk=0,obs=0,obh=0,antikoag=0,dmt1=0,gdm=0,kdm=0,cgm=0,insAnw=0,insdat=NULL'," & vbCrLf & _
"GROUP_CONCAT(setz SEPARATOR ''),' WHERE pat_id=',FPatnr) sqls" & vbCrLf & _
"FROM (" & vbCrLf & _
" SELECT FPatnr,zp,FUserID,FText,FFarbe, CONCAT(i.setz,IF(setz LIKE',insanw=%',CONCAT(',insdat=''',i.zp,''''),''))setz" & vbCrLf & _
" FROM (" & vbCrLf & _
"  SELECT FPatnr,zp,FUserID,FText,FFarbe" & vbCrLf & _
"  , CASE FText" & vbCrLf & _
"  WHEN 'Arzt Kothny' THEN ',obk=1'" & vbCrLf & _
"  WHEN 'Arzt Schade' THEN ',obs=1'" & vbCrLf & _
"  WHEN 'Arzt Hammerschmidt' THEN ',obh=1'" & vbCrLf & _
"  WHEN 'Antikoagulation' THEN ',antikoag=1'" & vbCrLf & _
"  WHEN 'Diabetes Typ I' THEN ',dmt1=1'" & vbCrLf & _
"  WHEN 'Gdm' THEN ',gdm=1'" & vbCrLf & _
"  WHEN 'kein D.m.' THEN ',kdm=1'" & vbCrLf & _
"  WHEN 'Libre-Handy' THEN ',cgm=1'" & vbCrLf & _
"  WHEN 'Libre Lesegerät' THEN ',cgm=2'" & vbCrLf & _
"  WHEN 'Dexcom-Clarity' THEN ',cgm=3'" & vbCrLf & _
"  WHEN 'Dexcom Lesegerät' THEN ',cgm=4'" & vbCrLf & _
"  WHEN 'Medtronic Simplera' THEN ',cgm=5'" & vbCrLf & _
"  WHEN 'Eversense' THEN ',cgm=6'" & vbCrLf & _
"  WHEN 'Novopen' THEN ',insanw=1'" & vbCrLf & _
"  WHEN 'Accu Chek Spirit Combo' THEN ',insAnw=2'" & vbCrLf & _
"  WHEN 'AccuChek Insight Yourloops' THEN ',insAnw=3'" & vbCrLf
sql = sql & _
"  WHEN 'Kaleido Glooko' THEN ',insAnw=4'" & vbCrLf & _
"  WHEN 'Medtronic 780 pdf Patient' THEN ',insAnw=5'" & vbCrLf & _
"  WHEN 'Omnipod 5 Glooko' THEN ',insAnw=6'" & vbCrLf & _
"  WHEN 'Omnipod Dash auslesen' THEN ',insAnw=7'" & vbCrLf & _
"  WHEN 'TSlim-Glooko Backoffice auslesen' THEN ',insAnw=8'" & vbCrLf & _
"  WHEN 'Ypsopump-Glooko' THEN ',insAnw=9'" & vbCrLf & _
"  WHEN 'Dana' THEN ',insAnw=10'" & vbCrLf & _
"  ELSE ''" & vbCrLf & _
"  END setz" & vbCrLf & _
"  FROM (" & vbCrLf & _
"   SELECT ROW_NUMBER()OVER(PARTITION BY d.FPatnr,d.FTablename,d.FPrimarykey ORDER BY zp DESC) rn," & vbCrLf & _
"      18900101+INTERVAL d.FDatum DAY+INTERVAL d.FUhrzeit SECOND zp,FPatnr,FUserID,FTyp,FPrimarykey" & vbCrLf & _
"     FROM dbsprot d" & vbCrLf & _
"     WHERE (" & FPatNr & "=0 OR fpatnr=" & FPatNr & ") AND d.FTablename='patmark'" & vbCrLf & _
"  )i" & vbCrLf & _
"  LEFT JOIN markier m ON i.fprimarykey=CONCAT('""',i.fpatnr,'""""',m.fsurogat,'""')" & vbCrLf & _
"  WHERE i.rn=1 AND i.ftyp=0" & vbCrLf & _
"  )i" & vbCrLf & _
" )i" & vbCrLf & _
"GROUP BY fpatnr" & vbCrLf & _
"ORDER BY fpatnr DESC" & vbCrLf & _
";"

 Set rMo = myEFrag(sql, rAf, MOCon)
 If Not rMo.BOF Then
  Do While Not rMo.EOF
   myEFrag rMo!sqls, rAf, DBCn
   mZl = mZl + rAf
   rMo.MoveNext
  Loop
 End If
 meld = "Markierungen bei " & mZl & " Patienten durch Übertragung von MO auf PraxisDB geändert."
 syscmd 4, meld
 Debug.Print meld
 
 Exit Function
#If False Then
' If SafeArrayGetDim(rNa) = 0 Then ReDim rNa(0)
'' sql = _
' "SELECT pm.FPatnr, ROW_NUMBER() OVER(ORDER BY fpatnr DESC,FMarkiernr) rg, COUNT(0) OVER() zl, LAST_VALUE(FMarkiernr) OVER(PARTITION BY FPatnr)=FMarkiernr speichern, pm.FMarkiernr,m.FText " & vbCrLf & _
' "FROM patmark pm " & vbCrLf & _
' "JOIN markier m ON pm.FMarkiernr=m.FSurogat" & vbCrLf
'' If FPatNr <> 0 Then
''  sql = sql & _
''  "WHERE fpatnr=" & FPatNr & vbCrLf
'' End If ' fpatnr <> 0 Then
' sql = _
' "SELECT FPatnr,zp,FUserID,FText,FFarbe FROM (" & vbCrLf & _
' "SELECT ROW_NUMBER()OVER(PARTITION BY d.FPatnr,d.FTablename,d.FPrimarykey ORDER BY zp DESC) rn," & vbCrLf & _
' "       18900101+INTERVAL d.FDatum DAY+INTERVAL d.FUhrzeit SECOND zp,FPatnr,FUserID,FTyp,FPrimarykey" & vbCrLf & _
' "         FROM dbsprot d" & vbCrLf & _
' "         WHERE " & IIf(FPatNr = 0, "", "FPatnr=" & FPatNr & " AND ") & "d.FTablename='patmark'" & vbCrLf & _
' ")i" & vbCrLf & _
' "LEFT JOIN markier m ON i.fprimarykey=CONCAT('""',i.fpatnr,'""""',m.fsurogat,'""')" & vbCrLf & _
' "WHERE i.rn=1 AND i.ftyp=0" & vbCrLf & _
' "ORDER BY fpatnr DESC;"
' Set rMo = myEFrag(sql, rAf, MOCon)
' If Not rMo.BOF Then
'  Do While Not rMo.EOF
'   If rMo!FPatNr <> altPNr Then ReDim rNa(0)
'   rNa(0).Pat_id = rMo!FPatNr
'   markAuswert rNa, rMo!FText
'   If nurfrag Then Exit Function
'    sql = ""
'    sql = sql & ",obk=" & IIf(rNa(0).obk = 0, "0", "1")
'    sql = sql & ",obs=" & IIf(rNa(0).obs = 0, "0", "1")
'    sql = sql & ",obh=" & IIf(rNa(0).obh = 0, "0", "1")
'    sql = sql & ",antikoag=" & IIf(rNa(0).antikoag = 0, "0", "1")
'    sql = sql & ",dmt1=" & IIf(rNa(0).dmt1 = 0, "0", "1")
'    sql = sql & ",gdm=" & IIf(rNa(0).gdm = 0, "0", "1")
'    sql = sql & ",kdm=" & IIf(rNa(0).kdm = 0, "0", "1")
'    sql = sql & ",cgm=" & IIf(rNa(0).cgm = 0, "0", rNa(0).cgm)
'    If rNa(0).insanw <> 0 Then sql = sql & ",insdat=" & Format(rMo!Zp, "yymmddhhmmss")
'    sql = sql & ",insAnw=" & IIf(rNa(0).insanw = 0, "0", rNa(0).insanw)
'    If sql <> "" Then
'     sql = "UPDATE namen SET " & Mid$(sql, 2) & " WHERE pat_id=" & rNa(0).Pat_id
'     myEFrag sql, rAf, DBCn
'     syscmd 4, "Markierung " & rMo!rg & " von " & rMo!zl & " eingetragen (geänderte Sätze: " & rAf & ")"
'     mZl = mZl + rAf
'    End If ' sql <> "" Then
'    ReDim rNa(0)
'   End If ' rMo!speichern <> 0 Then
'   altPNr = rMo!FPatNr
'   rMo.MoveNext
'  Loop
'  meld = "Markierungen bei " & mZl & " Patienten durch Übertragung von MO auf PraxisDB geändert."
'  syscmd 4, meld
'  Debug.Print meld
' End If ' Not rMo.BOF Then
 #End If
End Function ' doMarkierungen()

' in Notizen_Click(), doPatvonMO()
' "Infos" aus dem Krankenblatt (= FEintragsart 1105)
Public Sub doNotizen(Optional fPtNr& = 0, Optional mitSpeichern% = True)
 Dim rMo As ADODB.Recordset, rDl As ADODB.Recordset, rAf&, mZl&, FDet$
 Dim ErrNr&, ErrDes$
 If SafeArrayGetDim(rNa) = 0 Then
  ReDim rNa(0)
  rNa(0).Pat_ID = fPtNr
 End If
 Call MOConInit(, "doNotzizen(" & fPtNr & "," & CStr(mitSpeichern) & ")")
' Dim rNa() As namen
' ReDim rNa(0)
 sql = _
 "SELECT FPatnr, ROW_NUMBER() OVER(ORDER BY fpatnr DESC) rg, COUNT(0) OVER() zl" & vbCrLf & _
 ", REGEXP_REPLACE(GROUP_CONCAT(REPLACE(IF(INSTR(FDetails,'text ""'),MID(FDetails,LOCATE('text',FDetails)+LENGTH('text')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('text',FDetails)+LENGTH('text')+2)-LOCATE('text',FDetails)-LENGTH('text')-2),FText),'''','\''') separator '\r\n'),'^\\n','') FDet " & vbCrLf & _
 " FROM ltag l" & vbCrLf & _
 "WHERE (" & fPtNr & " = 0 OR fPatNr=" & fPtNr & ") AND" & vbCrLf & _
 " FEintragsart=1105" & vbCrLf & _
 " GROUP BY fpatnr -- ,fsurogat"
 Set rMo = myEFrag(sql, rAf, MOCon)
 If Not rMo.BOF Then
  Do While Not rMo.EOF
   FDet = REPLACE$(rMo!FDet, "\r\n", vbCrLf)
   obhierdmpfn FDet, rNa(0).NZNr, rNa(0).dmpklass, rNa(0).dmpbeg, rNa(0).dmpkhkklass, rNa(0).dmpkhkbeg, rNa(0).dmpcopdklass, rNa(0).dmpcopdbeg, rNa(0).dmpabklass, rNa(0).dmpabbeg, rNa(0).HzV, rNa(0).HzVbeg, rNa(0).Ds, rNa(0).DSbeg
   If mitSpeichern Then
    sql = "UPDATE namen SET NZNr=" & rNa(0).NZNr & ",dmpklass=" & rNa(0).dmpklass & ",dmpbeg='" & Format(rNa(0).dmpbeg, "yyyymmdd") & "',dmpkhkklass=" & rNa(0).dmpkhkklass & ",dmpkhkbeg='" & Format(rNa(0).dmpkhkbeg, "yyyymmdd") & "',dmpcopdklass=" & rNa(0).dmpcopdklass & ",dmpcopdbeg='" & Format(rNa(0).dmpcopdbeg, "yyyymmdd") & "',dmpabklass=" & rNa(0).dmpabklass & ",dmpabbeg='" & Format(rNa(0).dmpabbeg, "yyyymmdd") & "',HzV=" & rNa(0).HzV & ",HzVbeg='" & Format(rNa(0).HzVbeg, "yyyymmdd") & "',DS=" & rNa(0).Ds & ",DSbeg='" & Format(rNa(0).DSbeg, "yyyymmdd") & "' WHERE pat_id=" & rMo!FPatNr
    myEFrag sql, rAf, DBCn, , ErrNr, ErrDes
    If ErrNr <> 0 Then
     syscmd 4, "Bei PatNr: " & rMo!FPatNr & ": Fehler & " & ErrNr & ": " & ErrDes
    Else
     syscmd 4, "Notiz " & rMo!rg & " von " & rMo!zl & " eingetragen (geänderte Sätze: " & rAf & ")"
    End If ' ErrNr <> 0 Then
    mZl = mZl + rAf
    ReDim rNa(0)
   End If
   rMo.MoveNext
  Loop
  syscmd 4, "Notizen bei " & mZl & " Patienten durch Übertragung von MO auf PraxisDB geändert."
 End If ' Not rMo.BOF Then
End Sub ' doÜbertrag

' in HATrans
Public Function tbtrans(Tbl$)
  Static pwd$
  On Error GoTo fehler
  syscmd 4, "HATrans: Übertrage " & Tbl
  If pwd = "" Then pwd = InputBox("bitte Passwort für root auf Linux eingeben", "Passworteingabe")
  Shell """c:\program files\putty\plink.exe"" -pw " & pwd & " -batch root@linux1 ""mariadb-dump --defaults-extra-file=~/.modbpwd medoff " & Tbl & "|mariadb --defaults-extra-file=~/.mariadbpwd quelle"""
#If False Then
  Dim rMo As ADODB.Recordset, rq As ADODB.Recordset, ErrNr&, ErrDes$, einzeln%, sp0$, sp1$, iru%
  Call Lese.ProgStart
  Call MOConInit(, "tbtrans(" & Tbl & ")")
 
 sql = "SELECT GROUP_CONCAT(COLUMN_NAME), GROUP_CONCAT(CONCAT('COALESCE(',IF(column_type='longblob',CONCAT('REPLACE(LEFT(CONVERT(',COLUMN_NAME,' USING latin1),10000000),''\'',''\\'')'),COLUMN_NAME),','''')')) FROM information_schema.columns WHERE table_schema='medoff' and TABLE_NAME='" & Tbl & "' ORDER BY ordinal_position"
' sql = "SELECT GROUP_CONCAT(COLUMN_NAME), GROUP_CONCAT(CONCAT('COALESCE(REPLACE(REPLACE(',COLUMN_NAME,',''\n'',''''),''\r'',''''),'''')')) FROM information_schema.columns WHERE table_schema='medoff' and TABLE_NAME='" & tbl & "' ORDER BY ordinal_position"
 myFrag rMo, sql, adOpenStatic, MOCon, adLockReadOnly, 10000000, rAf, , ErrNr, ErrDes
 If Not rMo.BOF Then
  sp0 = rMo.Fields(0)
  sp1 = rMo.Fields(1)
  myFrag rq, "TRUNCATE `" & Tbl & "`", adOpenStatic, DBCn, adLockReadOnly, 10000000, rAf
  For iru = 1 To 2
   If iru = 2 Then einzeln = True
   sql = _
   "SELECT CONCAT('INSERT IGNORE INTO `" & Tbl & "`(" & vbCrLf & _
   sp0 & vbCrLf & _
    ") VALUES('''" & vbCrLf & _
   "," & IIf(einzeln, "", "GROUP_CONCAT(") & "CONCAT_WS(''','''," & sp1 & ")" & IIf(einzeln, "", " SEPARATOR '''),(''')") & vbCrLf & _
   ",''')') FROM `" & Tbl & "`"
'   "SELECT CONVERT(CONCAT('INSERT IGNORE INTO `" & tbl & "`(" & vbCrLf & _
'   ",''')') USING utf8mb4) FROM `" & tbl & "`"

   myFrag rMo, sql, adOpenStatic, MOCon, adLockReadOnly, 10000000, rAf
   If Not rMo.EOF Then
    Do While Not rMo.EOF
'     DBCn.Execute rMo.Fields(0), rAf
     myFrag rq, rMo.Fields(0), , DBCn, adLockReadOnly, 10000000, rAf, , ErrNr, ErrDes
     If ErrNr Then
      Debug.Print ErrNr, ErrDes
'      Exit Do
     End If
     rMo.MoveNext
    Loop
    If ErrNr = 0 Then Exit For
   End If 'Not rMo.EOF Then
  Next iru
 End If ' Not rMo.BOF Then
#End If
 syscmd 4, "Fertig mit Übertragung von " & Tbl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tbtrans/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' tbtrans(tbl$)

' für Arzt -> Hausärzte_von_MO_nach_Linux_übertragen_Click()
Public Function HATrans()
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 syscmd 4, "HATrans(), werfe die Datenbankverbindungen an"
 Lese.ProgStart
 Call MOConInit(, "HATrans()")
 If MsgBox("soll procmepraxis(0) aufgerufen werden (dauert mind. 7 min)?", vbYesNo, "Rückfrage") = vbYes Then
  syscmd 4, "HATrans: Extrahierte FAdresse (kann mindestens 7 Minuten dauern)"
  myEFrag "TRUNCATE tmpmepraxis", rAf, MOCon, , ErrNr, ErrDes, 10000000
  myEFrag "call procmepraxis(0)", rAf, MOCon, , ErrNr, ErrDes, 10000000
  If ErrNr Then
   syscmd 4, "Fehler " & ErrNr & ": " & ErrDes
  End If ' ErrNr
 End If ' true
 If ErrNr = 0 Then
  tbtrans "earzt"
  tbtrans "tmpmepraxis"
  tbtrans "epraxis"
  tbtrans "patrelation"
  syscmd 4, "Fertig mit HATrans"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HATrans/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' hatrans

' war nur zum einmaligen Aufruf
'Public Function inakt()
' Dim rMo As ADODB.Recordset
' Call Lese.ProgStart
' MOConInit
' Set rMo = myEFrag("SELECT fsurogat, 18900101+INTERVAL finaktivseit DAY seit FROM patstamm WHERE finaktivgrund=1;", rAf, MOCon)
' Do While Not rMo.EOF
'  myEFrag "UPDATE namen SET SDatum='" & rMo.Fields(1) & "', inaktiv=1 WHERE pat_id=" & rMo.Fields(0), rAf
'  Debug.Print rMo.Fields(0), rAf, rMo.Fields(1)
'  rMo.MoveNext
' Loop
'End Function
'
' war nur zum einmaligen Aufruf
'Public Function sdatum()
' Dim zeile$, pid$, pos&, ldat&
' MOConInit
' Open "u:\tmexport\sterbedatum.txt" For Input As #244
' Do While Not EOF(244)
'  Line Input #244, zeile
'  If InStrB(zeile, "Sterbedatum") Then
'   ldat = CDate(Mid(zeile, 14, 10)) - CDate("1.1.1890")
'   Call myEFrag("UPDATE patstamm SET FInaktivseit=" & ldat & " WHERE FSurogat=" & pid, rAf, MOCon)
'   Debug.Print pid, ldat, rAf
'  Else
'   pid = zeile
'  End If
' Loop
'End Function

' in PatvonMO_Click
Public Function doPatvonMO(fPtNr&, Optional obmitFormularen%, Optional obpruef%, Optional ohneLabor%, Optional obtranspa%, Optional oblabla%)
 Const obDebug% = False
 Dim pid&, pos&, pneu&, SchGr%, j&, jj%, rAf&, Puls$, Bem$, ErrNr&, ErrDes$ ' , rInh$, aktZeit As Date
 Dim meldTxt$
 Dim LaborLangsam%
 Dim ij&, rj& ' Laufvariable und zu befüllender Satz in rDM
 '    Veriablen für die rDm-Befüllung:
      Dim DMPArt%, uDat As Date, DokuDatum As Date, Druckdatum As Date, exportiert As Date
      Dim testdat As Date
 LaborLangsam = oblabla
' LaborLangsam = True
abermals:
'  fPtNr& = 68393  ' 69618 ' 63635 ' 67180 ' 63635 ' 64800 ' 69333 ' 68316 ' 65405 ' 45 ' 64659 ' 45 ' 69367 ' 69377 ' 53119 ' 51630 ' 105 ' 18 ' 246 ' 59152 ' 1394 ' 2112 ' 151 ' 225 '
 pid = setzPid(fPtNr)
 Static lfdfl&
 Dim rsFa As New ADODB.Recordset, rsMO As New ADODB.Recordset
 Dim FMem() As memoType ', Kat() As memoType, tKat() As memoType, Abl() As memoType, fAuft() As memoType
 Dim NaStr() As memoType, FaStr() As memoType, rsfaru%
 Dim EintS As SortierEintr
 Static EinL As New SortierListe, EinK As New SortierListe
 Call MOConInit(, "Übertragung aus MO von Pat. " & fPtNr)
 If obpruef Then ' prüft, ob Import Neues brächte
  Dim rab As ADODB.Recordset, raz As ADODB.Recordset
' dbsprot.fPrimaryKey ist varchar(35), ltag.FSurogat int(11), somit laesst sich kein vorhandere Index verwenden
' vorlaeufig wird ein Index dd_fpatnr erstellt für FPatnr, FTablename
'        "-- AND 18900101 + INTERVAL d.FDatum DAY + INTERVAL d.FUhrzeit SECOND<NOW()-INTERVAL 30 SECOND" & vbCrLf
  sql = "SELECT" & vbCrLf & _
        "18900101 + INTERVAL d.FDatum DAY + INTERVAL d.FUhrzeit SECOND laend/*, d.*,p.**/" & vbCrLf & _
        "FROM dbsprot d" & vbCrLf & _
        "LEFT JOIN dbsprot p ON d.FTablename='ltag' AND p.FTablename='extauftr' AND d.FPrimarykey=p.FPrimarykey" & vbCrLf & _
        "WHERE d.FPatnr=" & fPtNr & vbCrLf & _
        "AND d.FTablename IN ('ltag','termin')" & vbCrLf & _
        "AND p.FSurogat IS NULL" & vbCrLf & _
        "AND (d.FXmlinhalt IS NULL OR d.FXmlinhalt NOT RLIKE 'arztbrief|Erledigtdatum')" & vbCrLf & _
        "ORDER BY d.FDatum DESC,d.FUhrzeit DESC" & vbCrLf & _
        "LIMIT 1"
' AND ftablename NOT IN ('datafile','d2dmail','med95ini','mail','nutzerneu','markier','earzt','epraxis','tzone','ldtarc','globalitems','zertifikat','nutzerzugriff','patfall','patrelation')
  myFrag rab, sql, adOpenStatic, MOCon, adLockReadOnly
  If Not rab.EOF Then
   If Not IsNull(rab!laend) Then
    Set raz = Nothing ' notwendig bei Pat. 3776
    myFrag raz, "SELECT COALESCE(aktzeit,18990101) aktzeit FROM namen WHERE pat_id=" & fPtNr, , DBCn, adLockReadOnly, , rAf
    If Not raz.EOF Then
     If raz!aktZeit > rab!laend Then
'      If obtranspa Then
'       MsgBox sql & vbCrLf & "Patient wurde in MO zuletzt geändert: " & Format(rab!laend, "dd.mm.yyyy HH:MM:SS") & "," & vbCrLf & _
       "zuletzt importiert: " & Format(raz!aktZeit, "dd.mm.yyyy HH:MM:SS") & " => braucht nicht übertragen zu werden"
'      End If
      syscmd 4, "die Überprüfung ergab: keine Übertragung bei PtNr." & fPtNr & " notwendig!"
      Exit Function
'     Else ' raz!aktZeit > rab!laend Then
'      If obtranspa Then
'       MsgBox sql & vbCrLf & "Patient wurde in MO zuletzt geändert: " & Format(rab!laend, "dd.mm.yyyy HH:MM:SS") & "," & vbCrLf & _
       "zuletzt importiert: " & Format(raz!aktZeit, "dd.mm.yyyy HH:MM:SS") & " => wird übertragen"
'      End If
     End If ' raz!aktZeit > rab!laend Then
    End If ' Not raz.EOF Then
   End If ' Not IsNull(rab!laend) Then
  End If ' not rab.eof
 End If ' obpruef
 aktZeit = Now()
 syscmd acSysCmdSetStatus, "Übertrage Daten aus MO zu Pat. " & fPtNr
' BegTrans
 Call Tinit
 Call doTabVorb(Lese, 0, 0) ' obVorber, obmitFormularen)
' ComTrans
 syscmd 4, "richte desktopkop her"
 myEFrag "DELETE FROM desktopkop WHERE pat_id=" & pid
 Dim spal$
 spal = myEFrag("SELECT GROUP_CONCAT(COLUMN_NAME) FROM information_schema.columns c WHERE TABLE_NAME='desktop' AND table_catalog='def' AND TABLE_schema='quelle' AND column_key<>'PRI'", , , , , , 15000).Fields(0)
 myEFrag "INSERT INTO desktopkop(" & spal & ") SELECT " & REPLACE$(spal, "Pat_ID", "Pat_ID+" & CStr(Lese.pidoffs) & "") & " FROM desktop WHERE pat_id=" & pid
 Dim rdesk As ADODB.Recordset
 Dim aDesk() As desktop
 Set rdesk = myEFrag("SELECT CONCAT(tooltiptext,'\\n',titel) tit, d.* FROM desktop d WHERE pat_id=" & pid & " ORDER BY erstZP", rAf)
 Do While Not rdesk.EOF
  If SafeArrayGetDim(aDesk) Then ReDim Preserve aDesk(UBound(aDesk) + 1) Else ReDim aDesk(0)
  aDesk(UBound(aDesk)).erstZP = rdesk!erstZP
  aDesk(UBound(aDesk)).absPos = rdesk!absPos
  aDesk(UBound(aDesk)).aktZeit = rdesk!aktZeit
  aDesk(UBound(aDesk)).exoL = rdesk!exoL
  aDesk(UBound(aDesk)).hideT = rdesk!hideT
  aDesk(UBound(aDesk)).iconPath = rdesk!iconPath
  aDesk(UBound(aDesk)).IDS = rdesk!IDS
  aDesk(UBound(aDesk)).noteBkColor = rdesk!noteBkColor
  aDesk(UBound(aDesk)).noteFgColor = rdesk!noteFgColor
  aDesk(UBound(aDesk)).Pat_ID = pid
  aDesk(UBound(aDesk)).positionBottom = rdesk!positionBottom
  aDesk(UBound(aDesk)).positionLeft = rdesk!positionLeft
  aDesk(UBound(aDesk)).positionRight = rdesk!positionRight
  aDesk(UBound(aDesk)).positionTop = rdesk!positionTop
  aDesk(UBound(aDesk)).showAsNote = rdesk!showAsNote
  aDesk(UBound(aDesk)).syncInfoList = rdesk!syncInfoList
  aDesk(UBound(aDesk)).Titel = rdesk!tit
  aDesk(UBound(aDesk)).toolTipText = rdesk!toolTipText
  aDesk(UBound(aDesk)).verankert = rdesk!verankert
  rdesk.MoveNext
 Loop
 Call LöschePat(pid, , ohneLabor:=ohneLabor)
' On Error Resume Next
' If obDebug Then FSO.DeleteFile parsemotxt
 On Error GoTo fehler
' Tinit
 If EinL.COUNT = 0 Then
  syscmd 4, "Lade MOSystem"
 ' unter Ado müssen die Memo-Felder auf latin1 übersetzt werden für die Zahlen > 128 (nicht in HeidiSQL)
  sql = "SELECT COALESCE(CONVERT(FKategorieliste USING latin1),'') FKat, COALESCE(CONVERT(Ftextkategorieliste USING latin1),'') Ftk, COALESCE(CONVERT(fAblageliste USING latin1),'') FAb, COALESCE(CONVERT(fAuftragstypenliste USING latin1),'') FAuf, COALESCE(CONVERT(FMemo USING latin1),'') Fm " & vbCrLf & _
  "FROM mosystem"
  Dim FmS$
' rsMO.Open sql, MOCon, adOpenStatic, adLockReadOnly
  Set rsMO = myEFrag(sql, rAf, MOCon)
  On Error Resume Next
  FmS = rsMO!fm
  On Error GoTo fehler
  If FmS <> "" Then
   Call ParseMemo(FmS, FMem(), obDebug, "FMemo aus mosystem")
   For j = 0 To UBound(FMem)
    If FMem(j).ENr Like "*.2" And FMem(j).ENr <> "1.2" Then
     Set EintS = New SortierEintr
     EintS.TypNr = Asc(FMem(j).Text)
     If Len(FMem(j).Text) > 1 Then
      EintS.TypNr = EintS.TypNr + Fakt * Asc(Mid(FMem(j).Text, 2))
     End If
    ElseIf FMem(j).ENr Like "*.3" And FMem(j).ENr <> "1.3" Then
     EintS.art = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.5" And FMem(j).ENr <> "1.5" Then
     EintS.EKür = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.8" And FMem(j).ENr <> "1.8" Then
     EintS.name = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.10" And FMem(j).ENr <> "1.10" Then
     EintS.Kürz = FMem(j).Text
     EinL.sCAdd EintS
    End If
   Next j
  End If ' FmS <> "" Then
  On Error Resume Next
  FmS = rsMO!ftk
  On Error GoTo fehler
  If FmS <> "" Then
   Call ParseMemo(FmS, FMem(), obDebug, "FTextkategorie aus mosystem")
   For j = 0 To UBound(FMem)
    If FMem(j).ENr Like "*.2" Then
     Set EintS = New SortierEintr
     EintS.name = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.6" Then
     EintS.Kürz = FMem(j).Text
    ElseIf FMem(j).ENr Like "*.7" Then
     EintS.TypNr = Fakt * Asc(Mid(FMem(j).Text, 2)) + Asc(FMem(j).Text)
     EinK.sCAdd EintS
    End If
   Next j
  End If ' FmS <> "" Then
 End If ' einl.count=0
 
' Call MONamen(fPtNr)
 Dim rsNa As New ADODB.Recordset
' On Error GoTo fehler
 syscmd 4, "befülle Tabelle rna"
 ' konnte nicht genau rausfinden, wann FMemo richtig übermittelt wird, evtl. erst als zweites Feld, evtl. nicht unter dem Namen FMemo
 sql = "SELECT COALESCE(CONVERT(p.FMemo USING latin1),'') Fm, COALESCE(m.FText,'') mftxt, 18900101+INTERVAL FInaktivseit DAY SDatum, p.* " & vbCrLf & _
 "FROM patstamm p " & vbCrLf & _
 "LEFT JOIN patmark pm ON p.FSurogat=pm.FPatnr " & vbCrLf & _
 "LEFT JOIN markier m ON pm.FMarkiernr=m.FSurogat WHERE p.FSurogat=" & fPtNr
 rsNa.Open sql, MOCon, adOpenStatic, adLockReadOnly
 If Not rsNa.BOF Then
  rNa(0).aktZeit = 0 ' aktZeit ' erst am Schluss, s.u.
  rNa(0).Pat_ID = pid ' = fPtNr
  rNa(0).TM_Pat_ID = TMPid(pid) ' 6.4.25
  rNa(0).lfdnr = -1 ' Import aus MO
  rNa(0).Nachname = doUmwfSQL(rsNa!fnachname, True)
  rNa(0).NVorsatz = doUmwfSQL(rsNa!FNamensvorsatz, True)
  rNa(0).NVors = doUmwfSQL(rsNa!FNamenszusatz, True)
  rNa(0).Titel = rsNa!FTitel
  rNa(0).Vorname = rsNa!FVorname
  Select Case rsNa!FGeschlecht:  Case "2": rNa(0).geschlecht = "w":  Case "1": rNa(0).geschlecht = "m":  Case Else: rNa(0).geschlecht = rsNa!FGeschlecht: End Select
  If rsNa!FGeburtsdatum <> "" Then rNa(0).GebDat = DateSerial(Mid$(rsNa!FGeburtsdatum, 1, 4), Mid$(rsNa!FGeburtsdatum, 5, 2), Mid$(rsNa!FGeburtsdatum, 7, 2))
  rNa(0).Versichertennummer = rsNa!FAktversichertennr
  ' fehlt rna(0).KarGen: aus ' ',0,2,3
  If rsNa!FErstkontakt <> 0 Then rNa(0).AufnDat = CDate("1.1.1890") + rsNa!FErstkontakt
 ' rna(0).kAufDat=
  rNa(0).Straße = doUmwfSQL(rsNa!FStrasse, True)
  If Not IsNull(rsNa!FHausnr) Then If Not rsNa!FHausnr = "" Then rNa(0).Straße = rNa(0).Straße & " " & rsNa!FHausnr
  rNa(0).Hausnr = rsNa!FHausnr
  rNa(0).plz = rsNa!FPlz
  rNa(0).ort = rsNa!FOrt
  rNa(0).Lkz = rsNa!FLaendercode
  ' nicht enthalten/nicht befüllt: Anschrzus, PFPlz,anschrzus_2,postfach_2,lk_2,postfach,beruf,Weggeldzone,
  rNa(0).WeggzZahl = IIf(rsNa!FEntfernung = "", 0, rsNa!FEntfernung) ' bei Pat. 1219 doch nicht die Weggeldzone als Zahl
  rNa(0).Titel = rsNa!FTitel
  rNa(0).Swz = rsNa!FSchangerzahl
  rNa(0).Gbz = rsNa!FGeburtzahl
  rNa(0).Kiz = rsNa!FKinderzahl
  If rsNa!FInaktivgrund <> "" Then
   rNa(0).inaktiv = rsNa!FInaktivgrund
   If rNa(0).inaktiv = 1 Then rNa(0).SDatum = rsNa!SDatum
  End If ' rsNa!FInaktivgrund = 1 Then
 
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
#If mitmestdruck Then
  FSO.DeleteFile mestausg
#End If
  On Error GoTo fehler
  If rsNa!fm <> "" Then
'   If fPtNr = 70338 Then Stop
   Call ParseMemo(rsNa!fm, NaStr(), obDebug, "FMemo von rsNa (patstamm), Pat-id " & fPtNr)
'   Call MeStDruck(CStr(fPtNr), NaStr)
   If SafeArrayGetDim(NaStr) <> 0 Then
    For j = 0 To UBound(NaStr)
'     Debug.Print NaStr(j).enr, NaStr(j).Text
    Select Case NaStr(j).ENr
      Case "2" ' Notiz in der Patientenbearbeitung
         rNa(0).notiz = UmwfSQL(NaStr(j).Text)
      Case "18" ' Notiz im dynamic view
         rNa(0).notiz = UmwfSQL(REPLACE$(REPLACE$(LTrim$(IIf(Left$(LTrim$(NaStr(j).Text), 6) = "Notiz:", Mid$(LTrim$(NaStr(j).Text), 7), NaStr(j).Text)), vbCrLf, Chr$(10)), Chr$(10), vbCrLf) & IIf(rNa(0).notiz = "", "", vbCrLf & rNa(0).notiz))
      Case "21.1": ' asc( Zahl der eingetragenen Kinder
         rNa(0).ZdeK = Asc(NaStr(j).Text)
'     case "22": ' 71101, 02602, 68415, , 72601, 95301
      Case "25"
        If NaStr(j).Text <> 0 Then
         SchGr = 90
        End If
    End Select ' Case NaStr(j).ENr
   Next j ' j = 0 To UBound(NaStr)
   End If ' SafeArraygetdim
  End If ' rsNa!fm
'  Call MeStDruck(CStr(fPtNr), NaStr)
' ein Stück kommt noch hinter Fälle
    syscmd 4, "befülle Tabelle rfa"
'  rsFa.Open "SELECT f.fsurogat nix, COALESCE(CONVERT(f.FMemo USING latin1),'') Fm,CONCAT(f.fpatnr,', ',18900101 + INTERVAL f.fvon DAY,' - ',18900101 + INTERVAL f.fbis DAY) ueschr, f.*,a.FBezeichnung, le.FNachname FROM patfall f LEFT JOIN abrechner a ON f.FArztnr=a.FSurogat LEFT JOIN lstgerb le USING (FLstgerbnr) WHERE fpatnr=" & fPtNr & " ORDER BY FVon DESC", MOCon, adOpenStatic, adLockReadOnly
 sql = "SELECT f.fsurogat nix, COALESCE(CONVERT(f.FMemo USING latin1),'') Fm" & vbCrLf & _
 ", CONCAT(f.fpatnr,', ',18900101 + INTERVAL f.fvon DAY,' - ',18900101 + INTERVAL f.fbis DAY) ueschr" & vbCrLf & _
 ", f.*,a.FBezeichnung, le.FNachname" & vbCrLf & _
 "FROM patfall f " & vbCrLf & _
 "LEFT JOIN lstgerb le USING (FLstgerbnr)" & vbCrLf & _
 "LEFT JOIN patfall p0 ON f.fpatnr=p0.FPatnr AND  p0.fscheintyp=0 AND p0.ftarif=0 AND p0.fvon=49308" & vbCrLf & _
 "LEFT JOIN abrechner a ON f.FArztnr=a.FSurogat" & vbCrLf & _
 "WHERE f.fpatnr=" & fPtNr & vbCrLf & _
 "AND NOT (f.fvon=49308 AND ((f.fscheintyp=0 AND f.ftarif=0) OR (p0.FPatnr IS NOT NULL AND f.FScheinart=7)))" & vbCrLf & _
 "ORDER BY FVon DESC" ' Ausschluss der unbekannt-Fälle des 1.Quartals 2025 nach der Migration und der begeleitenden Privatfälle
  myFrag rsFa, sql, adOpenStatic, MOCon, adLockReadOnly
'  rsfaru = 0
  If Not rsFa.BOF Then
   Do While Not rsFa.EOF
    lfdfl = lfdfl + 1
    ReDim Preserve rFa(UBound(rFa) + 1)
    rFa(UBound(rFa)).aktZeit = aktZeit
    rFa(UBound(rFa)).lfdnr = lfdfl
    rFa(UBound(rFa)).Pat_ID = pid
    rFa(UBound(rFa)).AbrAr = ""
    rFa(UBound(rFa)).VermiArt = 0
    rFa(UBound(rFa)).bPerG = "0"
'     rFa(UBound(rFa)).GebOr = 2 ' in patfall und tmpmpatfall bei Pat. 151 kein Unterscheidungskriterium gefunden
'     rFa(UBound(rFa)).AbrAr = 2 ' in patfall und tmpmpatfall bei Pat. 151 kein Unterscheidungskriterium gefunden
    If rsFa!fvon <> 0 Then rFa(UBound(rFa)).BhFB = CDate("1.1.1890") + rsFa!fvon ' wenngleich nie 0
    rFa(UBound(rFa)).Quartal = ZQuart(rFa(UBound(rFa)).BhFB)
    If Not IsNumeric(rFa(UBound(rFa)).Quartal) Or Len(rFa(UBound(rFa)).Quartal) <> 5 Then
     MsgBox "IsNumeric(rFa(UBound(rFa)).Quartal) Or Len(rFa(UBound(rFa)).Quartal) <> 5 bei rFa(Ubound(rfa)).Quartal=" & rFa(UBound(rFa)).Quartal
     Stop
    End If ' Not IsNumeric(rFa(UBound(rFa)).Quartal) Or Len(rFa(UBound(rFa)).Quartal) <> 5 Then
    If rsFa!fbis <> 0 Then rFa(UBound(rFa)).BhFE1 = CDate("1.1.1890") + rsFa!fbis ' wenngleich nie 0
    ' FAbgerechnet 3652 = gesperrt, da nicht abrechenbar
    If rsFa!fabgerechnet <> 0 Then rFa(UBound(rFa)).BhFE2 = CDate("1.1.1890") + rsFa!fabgerechnet
    rFa(UBound(rFa)).BhFE = rFa(UBound(rFa)).BhFE2
    rFa(UBound(rFa)).IK = Right$(rsFa!Fik, 7)
    If Not IsNull(rsFa!FBezeichnung) Then rFa(UBound(rFa)).abrArzt = rsFa!FBezeichnung
    If rsFa!FVorhanden <> " " And rsFa!FVorhanden <> "" Then
      rFa(UBound(rFa)).KartBes = rsFa!FVorhanden
    End If
    
'    rFa(UBound(rFa)).lanrid = IIf(rsFa!FLstgerbnr = 3, 2, 1) ' 2 = Schade, 3 = Kothny
'    rFa(UBound(rFa)).lanrid = IIf(InStrB(rsFa!fnachname, "Kothny") <> 0, 2, IIf(InStrB(rsFa!fnachname, "Schade") <> 0, 2, 5)) ' 2 = Schade, 3 = Kothny, 5 = Hammerschmidt
    Select Case rsFa!fnachname
     Case "Schade": rFa(UBound(rFa)).lanrid = 1
     Case "Kothny": rFa(UBound(rFa)).lanrid = 2
     Case "Hammerschmidt": rFa(UBound(rFa)).lanrid = 5
     Case Else: rFa(UBound(rFa)).lanrid = 4
    End Select
    
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
     Call ParseMemo(rsFa!fm, FaStr(), obDebug, "FMemo von rsFa (patfall), Pat-id " & rsFa!ueschr)  ' rsFa!fpatnr, rsFa!fsurogat,
'    Call MeStDruck(fPtNr & " " & rsFa!ueschr, FaStr)
'  For jj = 1 To UBound(rFa)
'   If Not IsNumeric(rFa(jj).Quartal) Or Len(rFa(jj).Quartal) <> 5 Then Stop
'  Next jj
'    If rsfaru = 0 Then
     Dim buch$, ob71010%, Spli ' 5.9.25 aktuell nur bei Pat. 71010
     ob71010 = 0
     For j = 0 To UBound(FaStr)
'      If lfdfl = 1 Then
'       Select Case FaStr(j).enr
'        Case "3.2.2.2","1.3.2.2.2":          rNa(0).Versichertennummer = Trim$(FaStr(j).Text) ' Versichertennummer, s.o.
'        Case "3.2.2.3.2","1.3.2.2.3.2":      rNa(0).GebDat = CDate(Format$(FaStr(j).Text, "####\.##\.##")), s.o.
'        Case "3.2.2.3.3", "1.3.2.2.3.3":     rNa(0).geschlecht = IIf(FaStr(j).Text = "2" Or FaStr(j).Text = "w", "w", IIf(FaStr(j).Text = "1" Or FaStr(j).Text = "m", "m", " ")), s.o.
'        case "1.3.2.2.3.5":                  <evtl. Bitfeld>
'        Case "3.2.2.3.6.2","1.3.2.2.3.6.2":  rNa(0).plz = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.3","1.3.2.2.3.6.3":  rNa(0).ort = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.4","1.3.2.2.3.6.4":  rNa(0).Lkz = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.5","1.3.2.2.3.6.5":  rNa(0).Straße = FaStr(j).Text, s.o.
'        Case "3.2.2.3.6.6","1.3.2.2.3.6.6"
'           rNa(0).Hausnr = FaStr(j).Text, s.o.
'           If rNa(0).Hausnr <> "" Then rNa(0).Straße = rNa(0).Straße & " " & rNa(0).Hausnr, s.o.
'       End Select ' FaStr(0).enr
'      End If ' lfdfl = 1 Then
      
'     Debug.Print FaStr(j).enr, FaStr(j).Text
      Select Case FaStr(j).ENr
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
'         ' letzteres bei Pat. 70326:
       Case "3.2.2.3.4.4", "1.3.2.2.3.4.4": rFa(UBound(rFa)).Nachname = doUmwfSQL(FaStr(j).Text, True)
       Case "3.2.2.3.4.3", "1.3.2.2.3.4.3": rFa(UBound(rFa)).Vorname = FaStr(j).Text
       Case "3.2.2.4.2", "1.3.2.2.4.2": ' Versicherungsschutzbeginn ' letzteres bei Pat. 70326
          If FaStr(j).Text Like "20######" Then
            On Error Resume Next
            rFa(UBound(rFa)).VschBeg = CDate(Format$(FaStr(j).Text, "####\.##\.##"))
            If rFa(UBound(rFa)).VschBeg = 0 Then rFa(UBound(rFa)).VschBeg = CDate(Format$(FaStr(j).Text, "##\.##\.####"))
            On Error GoTo fehler
          Else ' FaStr(j).Text Like "20######" Then
            On Error Resume Next
            rFa(UBound(rFa)).VschBeg = CDate(Format$(FaStr(j).Text, "##\.##\.####"))
            If rFa(UBound(rFa)).VschBeg = 0 Then rFa(UBound(rFa)).VschBeg = CDate(Format$(FaStr(j).Text, "####\.##\.##"))
            On Error GoTo fehler
           End If ' FaStr(j).Text Like "20######" Then else
       Case "3.2.2.4.3": ' VK gültig bis ' gibt es nur auf wser
           If FaStr(j).Text Like "20######" Then
            On Error Resume Next
            rFa(UBound(rFa)).VschEnd = CDate(Format$(FaStr(j).Text, "####\.##\.##"))
            If rFa(UBound(rFa)).VschEnd = 0 Then rFa(UBound(rFa)).VschEnd = CDate(Format$(FaStr(j).Text, "##\.##\.####"))
            On Error GoTo fehler
           Else
            On Error Resume Next
            rFa(UBound(rFa)).VschEnd = CDate(Format$(FaStr(j).Text, "##\.##\.####"))
            If rFa(UBound(rFa)).VschEnd = 0 Then rFa(UBound(rFa)).VschEnd = CDate(Format$(FaStr(j).Text, "####\.##\.##"))
            On Error GoTo fehler
           End If
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
       Case "3.2.4", "1.3.2.4":
            rNa(0).eGKSchVer = FaStr(j).Text ' CDM-Version, z.B. 5.1.0, 5.2.0
       Case "3.2.5.2", "1.3.2.5.2": rFa(UBound(rFa)).DtlOnlPfg = BDTtoDateTime(Left$(FaStr(j).Text, 14))
       Case "3.2.5.3", "1.3.2.5.3": rFa(UBound(rFa)).ErgbdOnlP = Asc(Left$(FaStr(j).Text, 1))
       Case "3.2.5.4":
                              buch = Mid$(FaStr(j).Text, 2)
                              If buch = "" Then buch = Chr$(0)
                              rFa(UBound(rFa)).ErrorCode = Asc(buch) * 256& + Asc(Mid$(FaStr(j).Text, 1)) ' Pat. 60726 ("G/")
       Case "3.2.5.5", "1.3.2.5.5": rFa(UBound(rFa)).PrüfZdFd = FaStr(j).Text
       Case "3.3":            rNa(0).KarGen = FaStr(j).Text ' Kartentyp, 0, 2, 3
       Case "3.4", "1.3.4": ' letzeres Pat. 70326
            rFa(UBound(rFa)).lVorl = stzk(FaStr(j).Text)
            rFa(UBound(rFa)).ausgst = rFa(UBound(rFa)).lVorl ' wird nachher ggf. von 10.7 überschrieben
'            Call VorstellSetz(rFa(UBound(rFa)).lVorl)
       Case "4.2", "1.4.2":   rFa(UBound(rFa)).VKNr = FaStr(j).Text ' letzeres Pat. 70326
'       Case "4.3", "1.4.3": ' Kostenträgergruppe BDT 2018, in Turbomed nicht in der Falldatei
       Case "4.4", "1.4.4":   rFa(UBound(rFa)).KtrAbrB = Trim$(Left$(FaStr(j).Text, 2)) ' BDT 4106 unter 80000 Fällen fast immer 00, sonst 0, 00, 01, 06, 08, 1 und 2
       Case "4.5", "1.4.5":   rFa(UBound(rFa)).Kasse = Left$(FaStr(j).Text & Space$(27), 27) & " " & rFa(UBound(rFa)).Kasse ' letzters bei Pat. 70326
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
       Case "4.23":
                             rFa(UBound(rFa)).VermiCode = Trim$(FaStr(j).Text)
       Case "4.24":
'                             On Error Resume Next
                             rFa(UBound(rFa)).VermiDatum = stzk(FaStr(j).Text)
'                             On Error GoTo fehler
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
           Debug.Print "rFa " & FaStr(j).ENr & ": " & stzd(FaStr(j).Text)
       Case "5.4" ' ? (identisch mit 5.3, nicht Rechnungsbetrag)
           Debug.Print "rFa " & FaStr(j).ENr & ": " & stzd(FaStr(j).Text)
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
       Case "2.2"
        If IsNumeric(FaStr(j).Text) Then
         ob71010 = True
         GoTo m102
        End If
       Case "2.4"
        If ob71010 Then GoTo m104
       Case "2.5"
        If ob71010 Then GoTo m105
       Case "10.2", "1.10.2" ' letzteres bei Pat. 70326
m102:
           rFa(UBound(rFa)).ÜbWVBSNR = FaStr(j).Text ' Überweiser
       Case "10.3", "1.10.3"
           rFa(UBound(rFa)).ÜWZiel = FaStr(j).Text
       Case "10.4", "1.10.4"
m104:
           rFa(UBound(rFa)).ÜbwLANR = FaStr(j).Text
           rFa(UBound(rFa)).ÜbWVLANR = FaStr(j).Text
       Case "10.5", "1.10.5" ' <BSNR>#<LANR>#<epraxis.fsurogat>#<earzt.fsurogat>
m105:
            If rFa(UBound(rFa)).ÜbWVBSNR = "" Or rFa(UBound(rFa)).ÜbwLANR = "" Then
             Spli = Split(FaStr(j).Text, "#")
             If rFa(UBound(rFa)).ÜbWVBSNR = "" Then rFa(UBound(rFa)).ÜbWVBSNR = Spli(0)
             If rFa(UBound(rFa)).ÜbwLANR = "" Then rFa(UBound(rFa)).ÜbwLANR = Spli(1)
            End If ' ebsnr = "" Or elanr = "" Then
             ' Feld kommt aber auf wser nur einmal vor und auf szn4 nur viermal
       Case "10.6", "1.10.6"
            Dim han$, hav$, hatit$, hapos%
            hapos = InStr(FaStr(j).Text, "med.")
            If hapos <> 0 Then han = Trim$(Mid$(FaStr(j).Text, hapos + 4)): hatit = Left$(FaStr(j).Text, hapos + 4) Else han = FaStr(j).Text: hatit = ""
            hapos = InStr(han, ",")
            If hapos <> 0 Then hav = Mid$(han, hapos + 1): han = Left$(han, hapos - 1)
            rFa(UBound(rFa)).ÜWNaN = han
            rFa(UBound(rFa)).ÜWTit = hatit
            rFa(UBound(rFa)).ÜWVor = hav
       Case "10.7", "1.10.7" ' Ausstellungsdatum
           rFa(UBound(rFa)).ausgst = stzk(FaStr(j).Text)
'           Debug.Print "rFa " & FaStr(j).ENr & ": " & stzk(FaStr(j).Text)
       Case "10.8", "1.10.8": rFa(UBound(rFa)).Verdacht = doUmwfSQL(FaStr(j).Text, True)
       Case "10.9", "1.10.9": rFa(UBound(rFa)).Auftrag = doUmwfSQL(FaStr(j).Text, True)
'       Case "10.10"   ' kommt selten vor, immer "0"
       Case "10.13", "1.10.13": rFa(UBound(rFa)).Befund = doUmwfSQL(FaStr(j).Text, True)
'       Case "13"      ' immer  Ascii 2
'       Case "14"      ' kommt selten vor, Ascii 1 oder 4
'       Case "15"      ' kommt selten vor, Datum in der Nähe der Behandlungsdaten stzk(, Bedeutung konnte ich nicht ermitteln
'       Case "16"      ' kommt selten vor, Ascii 1 oder 4
'       Case "17":      rFa(UBound(rFa)).aktZeit = stzk(FaStr(j).Text) ' wohl Importdatum als stzk
' bPerG, DMPKnZ
'        Exit For
      End Select ' Case FaStr(j).enr
     Next j
' Schritte zur Korrektur falsch importierter Gültigkeiten:
' SELECT t.*
' , ASCII(TEXT)b1, ASCII(mid(TEXT,2))b2, ASCII(mid(TEXT,3))b3, ascii(mid(TEXT,4))b4
' , IF(ASCII(mid(TEXT,2))=7,CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)),'') Datum
' FROM tmpmpatfall t WHERE patnr=2169 /* and fsur IN (196030,145766) */ and TEXT = 20250228 ORDER BY cast(enr AS INTEGER), fsur;
' SELECT
' IF(ASCII(mid(TEXT,2))=7,CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)),'') Datum
' , t.* FROM tmppatfall t WHERE enr='3.2.2.4.3' AND TEXT=20250228;
' SELECT 18900101+interval fbis day bis,fbis,fsurogat,fpatnr FROM patfall WHERE fpatnr = 53814 AND fmemo RLIKE 20250228 AND fbis=49397;
' CALL procmwechs('patfall',(SELECT fsurogat FROM patfall WHERE fpatnr = 65474 AND fmemo RLIKE 20250228 AND fbis=49397),'FMemo','3.2.2.4.3','        ',1);
' CALL procmwechs('patfall',194815,'FMemo','3.2.2.4.3','        ',1);

     
' nicht übertragen:
     Dim rhar As New ADODB.Recordset
     Dim aktKVNr$
     sql = "SELECT farztnralt kvnr FROM earzt a" & vbCrLf & _
           "LEFT JOIN epraxis p ON p.FSurogat = a.FExtpraxisnr" & vbCrLf & _
           "WHERE farztnralt<>'' AND FBetriebsnr='" & rFa(UBound(rFa)).ÜbWVBSNR & "'" & vbCrLf & _
           "AND (farztnr='" & rFa(UBound(rFa)).ÜbwLANR & "' OR " & IIf(rFa(UBound(rFa)).ÜbwLANR = "", "TRUE", "FALSE") & ")"
     myFrag rhar, sql, adOpenStatic, MOCon
     If rhar Is Nothing Then
'       MsgBox "rhar is nothing mit: " & sql
'      Stop
     Else
      If Not rhar.BOF Then
       aktKVNr = Format$(rhar.Fields(0), String(10, "0"))
       rFa(UBound(rFa)).ÜbWVKVNR = aktKVNr
       rFa(UBound(rFa)).Übwr = aktKVNr
       rFa(UBound(rFa)).ÜWNNr = aktKVNr
       Dim uewv As New ADODB.Recordset
       For j = 1 To 2
        myFrag uewv, "SELECT ID,Nachname,Vorname,Titel FROM ueberwvon WHERE kvnr=" & aktKVNr, adOpenStatic, DBCn
        If Not uewv Is Nothing Then
         If uewv.BOF Then
          sql = "INSERT INTO ueberwvon(KVNr,Titel,Vorname,Zusatz,Nachname) VALUES(" & aktKVNr & ",'" & hatit & "','" & hav & "','','" & han & "')"
          InsKorr DBCn, sql, rAf
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
  
' 26.10.25: musste hinter Fallbefüllung gelegt werden
  If rsNa!fm <> "" Then
   If SafeArrayGetDim(NaStr) <> 0 Then
   For j = 0 To UBound(NaStr)
    If NaStr(j).ENr Like "21.*" And NaStr(j).ENr <> "21.1" Then
     ReDim Preserve rSw(UBound(rSw) + 1)
     rSw(UBound(rSw)).Pat_ID = pid
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
    End If ' NaStr(j).ENr Like "21.*" And NaStr(j).ENr <> "21.1" Then
   Next j ' j = 0 to UBound(NaStr)
   End If ' SafeArrayGetDim
  End If ' rsNa!fm <> ""
  Do While Not rsNa.EOF
   Call markAuswert(rNa, rsNa!mftxt)
   rsNa.MoveNext
  Loop
    
  lfdfl = 0 ' für nä Pat
  
' hier rFa fertig
  Dim infos As InfoTyp
  Call holHAausMO(infos, fPtNr)
' hier rNa fertig

 
 ' KVnr: fpatrelation, dort fpatid= FPatnr, FReferenztyp 2 = Hausarzt (0=Arbeitgeber), freferenzid = earzt.fsurogat,
 ' dort FExtpraxisnr = epraxis.fsurogat
  
  
' Set rsFa = Nothing
' Set rsFa = Nothing ' wirkt witzigerweise erst beim zweiten Mal (!?)
'  Call rFaDump
  
  Dim art$, neuart%, a1$, a2$, apos&
  Dim rsEi As New ADODB.Recordset
  Dim rFoNeu%, FormAbk$, FormBez$, lFormID&, i&, nextFormID&, rFm_Nr&
  Dim FoIDv& ' Pseudo-Foid
  Dim mt$, mdat$
  
  syscmd 4, "bearbeite Formulare"
'  FBehgrundnr>0: Diagnosen
'   FStatus -32767: meist Eintrag, oder Notiz
'                0: Diagnose oder Eintrag, 1: AU (=FEintragsart 19)
'  FStatus 2:     FEintragsart 5,7: Laborwerte, 50: Link, 148: Word-Brief, 151: Link, 166: Link,169: Brief 598: Dokumente
'  FStatus 3 und 4:  FEintragsart 20: Überweisung,
'  FStatus 40 und 41: (FEintragsart 12): Leistung (bei 41 evtl. GOÄ)
'  FStatus 100: FEintragsart 41: Unfallmeldung
' Eintragsarten: 1=Diagnose akut,
'                2=Diagnose inaktiv,
'                5: bei FStatus 0 meist Eintrag (außer FICDC..: "PATNRALPHA": Patientennummer numerisch, "LAR", "PLAR", "UTXT" (Überweisungstext), "MHNG" (Mahnung), "MED" (Medikamenteneintrag, wohl eMP-Eintrag, FDetails: "((EText ..."); wenn FStatus 2: dann Laborwert
'                   bei Gewicht Eintrag als Formular
'                7: Labor (Pat. 70595)
'                8=Freitext (tf), Verwandtschaftsverhältnisse und Notizen,
'                9=Anamnese,
'                10=Text-Befund (tb)
'                11 kommt nicht vor
'                12=Leistungen (FDetails: {(Gnrliste [{( ..."),
'                13=Medikamente (mit PZN in FICD..) und Hilfsmittel (FDetails: "{(Handelsname ..."
'                14=Med., ohne PZN,
'                15=Medikament (md)
'                16: Medikament (mh) mit Dosierung
'                17: Hilfsmittel (FDetails: {(Bezeichnung .."
'                18: Heilmittel (hp)
'                19: AU,
'                20: Üw,
'                21: Khs-Einweisung,
'                29: Text-Allergie (tl)
'                50 u. 148: Link auf Datei oder Word-Dokument aus Turbomed
'                151: z.T. Einträge, z.T. PDF-Dateien (meist: "ePDF: ...", "pdf: ..."), "bild: ...", oder Links ("link: ..."), "brief: ", alle FStatus 2
'                166: "link: ...", 169: "brief: ...", "wbr: ..."; 501 u. 598: jpg und tif, ohne Vorsilben, z.T. mit "link: " bei Sono-Bildern
'                1001: Eintrag (auch, aber nicht nur: sono) Zeile abgeschnitten, 1002: Eintrag aug, 1003: Blutabnahme, 1004: Einträge
'                1004: Text (Text)
'                1005: Desktop-Notizen, 1006: Einträge
'                1019: cr
'                1028: Icon aus Turbomed
'                1045: tk (tk)
'                1046: Text - tn (tn)
'                1053: Überweisungstexte
'                1085: Langrezept aus Turbomed
'                1099: Text - Cave (tc)
'                1105: Notiz ("Infos") (ti)
'                1110: Text-Notiz (tn)
'                1118: Text-Hinweis (th)
'                1144: taille (taille)
'                1148: Taille (tai)
'                1119: Trop-Test (trop)
'                2013: Markierung gesetzt
'                2017: Diagnosen Dauer
' Labor: FEintragsart 5 oder 7, FStatus immer 2, FStatusergaenzung immer 0, FBehgrundnr immer 0,
'        FDurchfNutzernr immer -2147483647, FEintragsnr immer -2147483647
'        FLstGerbNr 1,2 oder 4
' Feintragsart: (diesbezüglich in ltag und beschein gleich)
' 27143: DMP-Doku COPD
' 27144: DMP-Doku Typ 1
' 27188: DMP-Doku Typ 2
' 27193: Asthma bronchiale
' 27216: Herzinsuffizienz
' 27217: chronischer Rückenschmerz

  sql = _
"SELECT 18900101+INTERVAL l.FDatum DAY+INTERVAL l.FZeit SECOND Zp, na.FUsername ua, nb.FUsername ub, l.*, IF(INSTR(FText,':') BETWEEN 1 AND 6,LEFT(FText,INSTR(FText,':')-1),'') art, IF(INSTR(FText,':') BETWEEN 1 AND 6,TRIM(MID(FText,INSTR(FText,':')+1)),FText) ename" & vbCrLf & _
", l.FStatus lFSt, FText" & vbCrLf & _
", COALESCE(CONVERT(b.FMemo USING latin1),'') BFMemo, l.FEintragsart lFE, b.FEintragsart bFE, b.FSurogat bFSu, b.*" & vbCrLf & _
", l.FEintragsart IN(13,14,16,17,18,1085,2004,2005,2006,2007,2029) obRezE" & vbCrLf & _
", IF(FText RLIKE '^[ ]*[0-9]+[ ]*x.*',SUBSTRING_INDEX(FText,'x',1),1) Anz" & vbCrLf & _
", REGEXP_REPLACE(REGEXP_REPLACE(FText,'^([ ]*[0-9]+[ ]*x[ ]*)?(.*)[ ]*$','\2'),'([ ]*\(.*\)[ ]*)*$','') Med" & vbCrLf & _
", REGEXP_REPLACE(FText,'^([ ]*[0-9]+[ ]*x[ ]*)?.*\((.*)\)[ ]*$','\2') Rezkl" & vbCrLf & _
", REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(FText,'^[^(]*',''),'[ ]*\([^)]*\)[ ]*$',''),'\(([^)]*)\)','\1') Rkl0" & vbCrLf & _
", IF(INSTR(l.FDetails,'(Nonoutidem ""'),MID(l.FDetails,INSTR(l.FDetails,'(Nonoutidem ""')+LENGTH('(Nonoutidem ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'(Nonoutidem ""',-1),'"")')-1),'') nonoi" & vbCrLf & _
", IF(INSTR(l.FDetails,'(Anzahl '),MID(l.FDetails,INSTR(l.FDetails,'(Anzahl ')+LENGTH('(Anzahl '),INSTR(SUBSTRING_INDEX(l.FDetails,'(Anzahl ',-1),')')-1),'') Anzahl" & vbCrLf & _
", IF(INSTR(l.FDetails,'(Packungszahl '),MID(l.FDetails,INSTR(l.FDetails,'(Packungszahl ')+LENGTH('(Packungszahl '),INSTR(SUBSTRING_INDEX(l.FDetails,'(Packungszahl ',-1),')')-1),'') Packungszahl" & vbCrLf & _
", IF(INSTR(l.FDetails,'(Rezeptart '),MID(l.FDetails,INSTR(l.FDetails,'(Rezeptart ')+LENGTH('(Rezeptart '),INSTR(SUBSTRING_INDEX(l.FDetails,'(Rezeptart ',-1),')')-1),'') Rezeptart" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation|^dmp(dm|dtyp|khk)|^edmp(dm|khk|ab|copd)|DMP Teilnahmeerklärung' obdr" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation Diabetes|^dmp(dm|dtyp)|^edmp(dm)' obdmr" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE 'Typ II' obt2r" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation koronare|^dmp(khk)|^edmp(khk)' obkhr" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation COPD' obcor" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation Asthma' obasr" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation chronische H' obhir" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation chronischer R' obrsr" & vbCrLf & _
",(FIcdcode LIKE'%dmp%'OR FIcdcode='')AND b.FSurogat IS NOT NULL AND FText RLIKE '(Erst|Verlaufs)-Dokumentation Brustkrebs' obbkr" & vbCrLf
sql = sql & _
"FROM (SELECT l.*" & vbCrLf & _
"FROM ltag l) l" & vbCrLf & _
"LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat" & vbCrLf & _
"LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat" & vbCrLf & _
"LEFT JOIN beschein b ON b.FSurogat = l.FEintragsnr" & vbCrLf & _
"WHERE l.fpatnr=" & fPtNr & vbCrLf & _
"HAVING (NOT ISNULL(bFE)OR(obRezE))" & vbCrLf & _
""
' ", IF(FText RLIKE '^[ ]*[0-9]+[ ]*x.*',MID(FText,INSTR(FText,'x')+1),FText) Med" & vbCrLf & _
", IF(FText RLIKE '.*(.)',LEFT(SUBSTRING_INDEX(FText,'(',-1),INSTR(SUBSTRING_INDEX(FText,'(',-1),')')-1),0) Rezkl" & vbCrLf & _

  myFrag rsEi, sql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
  ' das BTM in "... (BTM) (K)" noch verwerten
   Dim rseiru&
   rseiru = 0
   Do While Not rsEi.EOF
'    Debug.Print rsEi!fsurogat, rsEi.Fields(3)
'    If rsEi.Fields(3) = 11699396 Then Stop
'    If rseiru = 298 Then Stop
    rseiru = rseiru + 1
    If rsEi!bfmemo <> "" Then
'     If rsEi!fsurogat = 16045 Then Stop
'     If rseiru = 3 Then Stop
     Call ParseMemo(rsEi!bfmemo, FMem(), obDebug, "FMemo aus beschein") ' Pat. 59535, rseiru 298: FMem kann danach auch leer bleiben!
    End If ' rsEi!BFMemo <> ""
    'dmpreihe (1); Doppeleinträge gemäß Index "eindeutig" vermeiden
    If rsEi!obdr Then
      DMPArt = 0
      uDat = 0
      DokuDatum = 0
      Druckdatum = 0
      exportiert = 0
      testdat = 0
'       Debug.Print rsEi!FText
       If SafeArrayGetDim(FMem) <> 0 Then
        If rsEi!obdmr <> 0 Then
         If rsEi!obt2r <> 0 Then ' Typ 2
          DMPArt = 2
          For j = 0 To UBound(FMem)
           Select Case FMem(j).ENr
            Case "75": ' vermutlich Formularversion
                uDat = stzk(FMem(j).Text)
            Case "96": ' bei Typ 1: 6
                If Not Len(FMem(j).Text) = 1 And Asc(FMem(j).Text) = 1 Then
                   DokuDatum = stzk(FMem(j).Text)
                End If
'           Case "121":
'                Stop
            Case "115", "117", "118", "136", "137" ' bei Typ 1: 91
                testdat = stzk(FMem(j).Text)
                If testdat Then
                 If FMem(j).ENr = "117" Then
                   If Druckdatum = 0 Then
                    Druckdatum = testdat
                   Else
                    exportiert = testdat
                   End If
                 Else
                   Druckdatum = testdat
                 End If
                End If
            Case "119", "120", "138", "139" ' ' bei Typ 1: 104; 136/138 = Pat. 1339; 115/117 = 2885
                testdat = stzk(FMem(j).Text)
                If testdat Then exportiert = testdat
           End Select
          Next j
         Else ' rsEi!obt2r <> 0 Then: Typ 1
          DMPArt = 1
          For j = 0 To UBound(FMem)
           Select Case FMem(j).ENr
            Case "4" ' vermutlich Formularversion, bei Typ 1: 4
                 uDat = stzk(FMem(j).Text)
            Case "6"
                If Not Len(FMem(j).Text) = 1 And Asc(FMem(j).Text) = 1 Then
                   DokuDatum = stzk(FMem(j).Text)
                End If
'           Case "121":
'                Stop
            Case "91"
                testdat = stzk(FMem(j).Text)
                If testdat Then Druckdatum = testdat
            Case "104"
                testdat = stzk(FMem(j).Text)
                If testdat Then exportiert = testdat
           End Select
          Next j
         End If ' rsEi!obt2r <> 0 Then
        ElseIf rsEi!obKHr <> 0 Then 'koronare Herz
         DMPArt = 3
         For j = 0 To UBound(FMem)
          Select Case FMem(j).ENr
           Case "4": ' vermutlich Formularversion
                uDat = stzk(FMem(j).Text)
           Case "19":
                If Not Len(FMem(j).Text) = 1 And Asc(FMem(j).Text) = 1 Then
                   DokuDatum = stzk(FMem(j).Text)
                End If
'         Case "121":
'                Stop
           Case "66":
                Druckdatum = stzk(FMem(j).Text)
           Case "72" ' , "91":
'                If IsDate(stzk(FMem(j).Text)) Then
                  exportiert = stzk(FMem(j).Text)
'                End If
          End Select
         Next j
        ElseIf rsEi!obcor <> 0 Then ' COPD
         DMPArt = 4
         For j = 0 To UBound(FMem)
          Select Case FMem(j).ENr
           Case "4": ' vermutlich Formularversion
                uDat = stzk(FMem(j).Text)
           Case "9":
                If Not Len(FMem(j).Text) = 1 And Asc(FMem(j).Text) = 1 Then
                   DokuDatum = stzk(FMem(j).Text)
                End If
' Rest muss noch überprüft werden
          End Select
         Next j
        ElseIf rsEi!obasr Then ' Asthma
         DMPArt = 5
         For j = 0 To UBound(FMem)
          Select Case FMem(j).ENr
           Case "4": ' vermutlich Formularversion
                uDat = stzk(FMem(j).Text)
           Case "13":
                If Not Len(FMem(j).Text) = 1 And Asc(FMem(j).Text) = 1 Then
                   DokuDatum = stzk(FMem(j).Text)
                End If
' Rest muss noch überprüft werden
          End Select
         Next j
        ElseIf rsEi!obhir Then ' Herzinsuffizienz
         DMPArt = 6
' muss noch überprüft werden
        ElseIf rsEi!obrsr Then ' Rückenschmerz
         DMPArt = 7
' muss noch überprüft werden
        ElseIf rsEi!obbkr Then ' Brustkrebs
         DMPArt = 8
' muss noch überprüft werden (4 gibts)
        Else
         Debug.Print "noch was anderes"
        End If
       End If ' SafeArryGetDim(FMem)
       If DokuDatum <> 0 And DokuDatum <> CDate(rsEi!Zp) Then
        Debug.Print "Unterschied bei " & pid & " im Dokudatum: " & DokuDatum & " -> " & CDate(rsEi!Zp)
       End If
       DokuDatum = CDate(rsEi!Zp)
    
       For ij = 1 To UBound(rDm) ' um dem eindeutigen Index gerecht zu werden
        If rDm(ij).Pat_ID = pid And rDm(ij).DMPArt = DMPArt And rDm(ij).DokuDatum = DokuDatum Then
         rj = ij
         GoTo gefunden
        End If
       Next ij
       rj = UBound(rDm) + 1
       ReDim Preserve rDm(rj)
       rDm(rj).Pat_ID = pid
       rDm(rj).DMPArt = DMPArt
       rDm(rj).DokuDatum = DokuDatum
gefunden:
       rDm(rj).exportiert = exportiert
       rDm(rj).aktZeit = aktZeit
'       pos = InStr(rsEi!FText, "#")
'       If pos > 0 Then rDm(rj).Abk = Left$(rsEi!FText, pos - 1)
       rDm(rj).art = IIf(InStrB(rsEi!FText, "Erst"), "ED", "FD")
       rDm(rj).Abk = rsEi!FText
       rDm(rj).Ok = rsEi!lFSt
'       rDm(rj).ausgedruckt = IIf(InStrB(rsEi!erg, "ausgedruckt"), 1, 0)
'       pos = 1
'       Do
'        pneu = InStr(pos + 1, rsEi!erg, "exportiert am")
'        If pneu = 0 Then Exit Do Else pos = pneu
'       Loop
'       If pos > 1 Then rDm(rj).exportiert = CDate(Mid$(rsEi!erg, pos + 14, 10))
       rDm(rj).KarteiDatum = DokuDatum ' CDate(rsEi!Zp)
'       rDm(rj).DokuDatum = rDm(rj).KarteiDatum ' ist schon oben
'       rDm(rj).obvoll = IIf(InStrB(rsEi!erg, "vollständig"), 1, 0)
'       rDm(rj).Ok = IIf(InStrB(rsEi!erg, "(ok"), 1, 0)
'       rDm(rj).lanrid = rsEi!farztnr
       Select Case rsEi!FLstgerbnr
        Case 2: rDm(rj).lanrid = 1 ' Schade
        Case 3: rDm(rj).lanrid = 2 ' Kothny
        Case 4: rDm(rj).lanrid = 5 ' Hammerschmidt
        Case Else: rDm(rj).lanrid = 4 ' unbek
       End Select
       rDm(rj).Nachname = rNa(0).Nachname
       rDm(rj).Vorname = rNa(0).Vorname
       rDm(rj).GebDat = rNa(0).GebDat
    ElseIf rsEi!obRezE Then ' Rezepteintrag
     ReDim Preserve rRe(UBound(rRe) + 1)
     rRe(UBound(rRe)).aktZeit = aktZeit
     rRe(UBound(rRe)).Pat_ID = pid
     rRe(UBound(rRe)).Zeitpunkt = rsEi!Zp
     rRe(UBound(rRe)).Medikament = doUmwfSQL(rsEi!Med, True)
     rRe(UBound(rRe)).PZN = rsEi!FIcdcode
'     rRe(UBound(rRe)).lanrid = rsEi!farztnr
     Select Case rsEi!FLstgerbnr
      Case 2: rRe(UBound(rRe)).lanrid = 1 ' Schade
      Case 3: rRe(UBound(rRe)).lanrid = 2 ' Kothny
      Case 4: rRe(UBound(rRe)).lanrid = 5 ' Hammerschmidt
      Case Else: rRe(UBound(rRe)).lanrid = 4 ' unbek
     End Select
     rRe(UBound(rRe)).anzl = Switch(rsEi!anz <> "" And IsNumeric(rsEi!anz), rsEi!anz, rsEi!Anzahl <> "", rsEi!Anzahl, rsEi!packungszahl <> "", rsEi!packungszahl, True, 0)
     rRe(UBound(rRe)).FEintragsart = rsEi!lFE
     rRe(UBound(rRe)).kbez = rsEi!rkl0
     Select Case rsEi!RezKl
      Case "K": rRe(UBound(rRe)).Rezklkurz = "rp": rRe(UBound(rRe)).Rezkllang = "Kassenrp"
      Case "P": rRe(UBound(rRe)).Rezklkurz = "prp": rRe(UBound(rRe)).Rezkllang = "Privatrp"
      Case "S": rRe(UBound(rRe)).Rezklkurz = "sp": rRe(UBound(rRe)).Rezkllang = "Sprechstundenbedarf"
      Case "M": rRe(UBound(rRe)).Rezklkurz = "mu": rRe(UBound(rRe)).Rezkllang = "Muster"
      Case "F": rRe(UBound(rRe)).Rezklkurz = "fr": rRe(UBound(rRe)).Rezkllang = "Fremd"
      Case "V": rRe(UBound(rRe)).Rezklkurz = "vk": rRe(UBound(rRe)).Rezkllang = "Verkauf"
      Case "SM": rRe(UBound(rRe)).Rezklkurz = "sm": rRe(UBound(rRe)).Rezkllang = "Sebstmedikation"
      Case "BTM": rRe(UBound(rRe)).Rezklkurz = "btm": rRe(UBound(rRe)).Rezkllang = "BTM-Rezept"
     End Select
     If IsNumeric(rsEi!Rezeptart) Then rRe(UBound(rRe)).Rezeptart = rsEi!Rezeptart
     If rsEi!nonoi <> "" Then rRe(UBound(rRe)).auti = rsEi!nonoi
    
     If rsEi!bfmemo <> "" Then
'      rRe(UBound(rRe)).auti = 1 ' manchmal in Turbomed auch 2
      For j = 0 To UBound(FMem)
       Select Case FMem(j).ENr
        Case "6.2.3"
         rRe(UBound(rRe)).anzl = Asc(FMem(j).Text)
        Case "6.2.4"
         If Asc(FMem(j).Text) = 1 Then rRe(UBound(rRe)).auti = 1
        Case "6.2.5"
         rRe(UBound(rRe)).Medikament = doUmwfSQL(FMem(j).Text, True)
        Case "6.2.8"
         If rRe(UBound(rRe)).PZN = "" And FMem(j).Text <> "0" Then rRe(UBound(rRe)).PZN = FMem(j).Text
       End Select
      Next j
     End If ' rsEi!BFMemo<>""
    Else ' rsEi!obRezE Then ' Rezepteintrag
    Select Case rsEi!lFE
     Case 21 ' Krankenhauseinweisung
      ReDim Preserve rKh(UBound(rKh) + 1)
      rKh(UBound(rKh)).Pat_ID = pid
      rKh(UBound(rKh)).Zeitpunkt = rsEi!Zp
      rKh(UBound(rKh)).aktZeit = aktZeit
      For j = 0 To UBound(FMem)
       mt = FMem(j).Text & String(4, Chr(0))
       mdat = Asc(Mid(mt, 4, 1)) & "." & Asc(Mid(mt, 3, 1)) & "." & Asc(Mid(mt, 2, 1)) * 256& + Asc(mt)
       If IsDate(mdat) Then
        mt = mdat
       Else
        mt = Left$(mt, Len(mt) - 4)
        Do
         If LenB(mt) = 0 Then Exit Do
         If Asc(Right$(mt, 1)) < 10 Then mt = Left$(mt, Len(mt) - 1) Else Exit Do ' 3 kommt auch vor
        Loop
        If Right$(mt, 1) = Chr$(10) Then mt = Left$(mt, Len(mt) - 1)
        mt = doUmwfSQL(mt, True)
       End If ' IsDate(mdat) Then else
       Select Case FMem(j).ENr
        Case "3":
         rKh(UBound(rKh)).obNot = Asc(FMem(j).Text) Mod 2
         rKh(UBound(rKh)).obBeleg = Asc(FMem(j).Text) / 2
        Case "4":
         rKh(UBound(rKh)).Diagnose = mt
        Case "5":
         rKh(UBound(rKh)).Ziel = mt
        Case "6":
         rKh(UBound(rKh)).Befund = doUmwfSQL(mt, True)
        Case "7":
         rKh(UBound(rKh)).BisMas = doUmwfSQL(mt, True)
        Case "8":
         rKh(UBound(rKh)).FraStel = doUmwfSQL(mt, True)
        Case "9":
         rKh(UBound(rKh)).MitBef = doUmwfSQL(mt, True)
       End Select
      Next j
     Case Else ' Formular
      rFoNeu = -1
      FormAbk = rsEi!lFE
      FormBez = doUmwfSQL(rsEi!FText, True)
      For i = 1 To UBound(rFo)
       If rFo(i).Form_Abk = FormAbk And rFo(i).FormBez = FormBez Then
        rFoNeu = 0
        lFormID = rFo(i).FormID
        GoTo fgefunden
       End If ' rFo(i).Form_Abk =
      Next i ' = 1 To UBound(rFo)
fgefunden:
      If rFoNeu And Not obmitFormularen Then
       Dim rsf As ADODB.Recordset
       Call myFrag(rsf, "SELECT * FROM formulare WHERE Form_Abk='" & FormAbk & "' AND FormBez='" & FormBez & "' ORDER BY FormID", adOpenStatic)
       If Not rsf.BOF Then
        rFoNeu = 0
        ReDim Preserve rFo(UBound(rFo) + 1)
        rFo(UBound(rFo)).absPos = rsf!absPos
        rFo(UBound(rFo)).aktZeit = rsf!aktZeit
        rFo(UBound(rFo)).Form_Abk = IIf(IsNull(rsf!Form_Abk), vNS, rsf!Form_Abk)
        rFo(UBound(rFo)).FormBez = rsf!FormBez
        rFo(UBound(rFo)).FormID = rsf!FormID
        lFormID = rsf!FormID
        rFo(UBound(rFo)).StByte = rsf!StByte
       End If ' Not rsf.BOF Then
      End If ' rFoNeu And Not obmitFormularen Then
      If rFoNeu Then
       ReDim Preserve rFo(UBound(rFo) + 1)
       rFo(UBound(rFo)).Form_Abk = FormAbk
       rFo(UBound(rFo)).aktZeit = aktZeit
       rFo(UBound(rFo)).FormBez = FormBez
       If nextFormID = 0 Then nextFormID = -myEFrag("SELECT MAX(formid)+1 FROM formulare").Fields(0) Else nextFormID = nextFormID - 1
       rFo(UBound(rFo)).FormID = nextFormID
       ' rFo(UBound(rFo) - 1).FormID 1   ' muss noch in `forminhkopf` angpasst werden, wird deshalb dort negativ gespeichert
       lFormID = rFo(UBound(rFo)).FormID
       Call myFrag(rsf, "SELECT Form_AbkVW FROM forminhaltform_abk WHERE Form_Abk ='" & FormAbk & "' AND FormBez='" & FormBez & "'", adOpenStatic)
       If rsf.BOF Then
        InsKorr DBCn, "INSERT INTO forminhaltform_abk(Form_Abk,FormBez) VALUES('" & FormAbk & "','" & FormBez & "')", rAf
       End If
      End If ' rFoNeu Then
      rFm_Nr = rFm_Nr + 1
      
      ReDim Preserve rFr(UBound(rFr) + 1)
      rFr(UBound(rFr)).aktZeit = aktZeit
      rFr(UBound(rFr)).Form_ID = lFormID '-lFormID ' negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFo angepaßt werden muss
      rFr(UBound(rFr)).Pat_ID = pid
      rFr(UBound(rFr)).Zeitpunkt = rsEi!Zp
'      rFr(UBound(rFr)).lanrid = IIf(rsEi!FLstgerbnr = 3, 2, 1) ' 2 = Schade, 3 = Kothny
      Select Case rsEi!FLstgerbnr
       Case 2: rFr(UBound(rFr)).lanrid = 1 ' Schade
       Case 3: rFr(UBound(rFr)).lanrid = 2 ' Kothny
       Case 4: rFr(UBound(rFr)).lanrid = 5 ' Hammerschmidt
       Case Else: rFr(UBound(rFr)).lanrid = 4 ' ?
      End Select
      If FoIDv = 0 Then
       FoIDv = myEFrag("SELECT (COALESCE(MAX(foid))+1) mfoid FROM `forminhkopf`").Fields(0)
      End If
      rFr(UBound(rFr)).Foid = FoIDv ' Pseudo-Foid
      FoIDv = FoIDv + 1
   
    ' FormVorl unbekannt
'      If rsEi!bfmemo <> "" Then ' ParseMemo oben
      If SafeArrayGetDim(FMem) Then ' ParseMemo oben
       For j = 0 To UBound(FMem)
        ReDim Preserve rFm(UBound(rFm) + 1)
        rFm(UBound(rFm)).nr = rFm_Nr
        rFm_Nr = rFm_Nr + 1
        rFm(UBound(rFm)).Feld = FMem(j).ENr
        mt = FMem(j).Text & String(4, Chr(0))
        mdat = Asc(Mid(mt, 4, 1)) & "." & Asc(Mid(mt, 3, 1)) & "." & Asc(Mid(mt, 2, 1)) * 256& + Asc(mt)
        If IsDate(mdat) Then
         mt = mdat
        Else
         mt = Left$(mt, Len(mt) - 4)
         Do
          If LenB(mt) = 0 Then Exit Do
'          If Asc(Right$(mt, 1)) < 10 Then mt = Left$(mt, Len(mt) - 1) Else Exit Do ' 3 kommt auch vor
          If Asc(Right$(mt, 1)) = 0 Then mt = Left$(mt, Len(mt) - 1) Else Exit Do ' 3 kommt auch vor
         Loop
         If Right$(mt, 1) = Chr$(10) Then mt = Left$(mt, Len(mt) - 1)
        End If
        rFm(UBound(rFm)).FeldInh = mt
        rFm(UBound(rFm)).FeldNr = j
        rFm(UBound(rFm)).Foid = FoIDv - 1 '-FoID 'negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFr angepaßt werden muss
       Next j
      End If ' rsEi!BFMemo <> "" Then
    End Select ' Case rsEi!FEintragsart
    End If ' rsEi!obRezE Then ' Rezepteintrag else
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then


  syscmd 4, "bearbeite Briefe"
'  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, FICdcode Art, MID(fdetails,INSTR(fdetails,'ext ""')+5,LENGTH(fdetails)-2-INSTR(fdetails,'ext ""')-5) FText, FEintragsart, f.* FROM ltag f WHERE fpatnr = " & fPtNr & " AND ((FEintragsart=5 and FStatus=0) OR FEintragsart IN (8,10,11,151,1001,1002,1003,1004,1006)) AND fbehgrundnr<=0"
' Dokumente
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, na.FUsername ua, nb.FUsername ub, l.*, d.*,IF(INSTR(FText,':') BETWEEN 1 AND 6,LEFT(FText,INSTR(FText,':')-1),'') art, IF(INSTR(FText,':') BETWEEN 1 AND 6,TRIM(MID(FText,INSTR(FText,':')+1)),FText) ename " & vbCrLf & _
  "FROM ltag l " & vbCrLf & _
  "INNER JOIN datafile d ON l.FSurogat=d.FReferenznr " & vbCrLf & _
  "LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat " & vbCrLf & _
  "LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat " & vbCrLf & _
  "WHERE l.fpatnr=" & fPtNr & " AND feintragsart IN (5,50,51,52,148,151,166,169,501,598,1033) " & vbCrLf & _
  "ORDER BY l.FSurogat, Zp"
  myFrag rsEi, sql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
' Typ As String 'Typ varchar '
    ReDim Preserve rBr(UBound(rBr) + 1)
    rBr(UBound(rBr)).Pat_ID = pid
    rBr(UBound(rBr)).aktZeit = aktZeit
    rBr(UBound(rBr)).Zeitpunkt = rsEi!Zp
    rBr(UBound(rBr)).name = doUmwfSQL(rsEi!EName, True)
    rBr(UBound(rBr)).autor = rsEi!ua
    rBr(UBound(rBr)).art = rsEi!art
    rBr(UBound(rBr)).DokGroe = rsEi!FFilesize
    rBr(UBound(rBr)).DokAenD = rsEi!FCreationtime
    rBr(UBound(rBr)).Pfad = doUmwfSQL(rsEi!FFilename, True)
    rBr(UBound(rBr)).QS = ZQSort(rBr(UBound(rBr)).Zeitpunkt)
    rBr(UBound(rBr)).QT = ZQuart(rBr(UBound(rBr)).Zeitpunkt)
    rBr(UBound(rBr)).Quelldatum = doQuelldatum(rBr(UBound(rBr)).name, rBr(UBound(rBr)).DokAenD)
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
  
  syscmd 4, "bearbeite Medpläne"
' Medplan
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, COALESCE(na.FUsername,'') ua, REPLACE(FDosiertext,CHR(0),'') FDt, REPLACE(FHinweise,CHR(0),'') FHw, REPLACE(FBemerkung,CHR(0),'') FBm, l.*, m.*, COALESCE(CONVERT(m.FMemo USING latin1),'') Fm " & vbCrLf & _
        "FROM ltag l INNER JOIN meddosis m ON l.FText='Dosierplan' AND l.FSurogat=m.FDosierplannr " & vbCrLf & _
        "LEFT JOIN nutzerneu na ON FNutzernr= na.FSurogat " & vbCrLf & _
        "WHERE l.fpatnr=" & fPtNr & " AND FVerbindungsschluessel<>'#ZWTEXT#' " & vbCrLf & _
        "ORDER BY l.FSurogat, FDosierplanpos"
' '#ZWTEXT#' sind Zwischenüberschriften, z.B. "Bedarfsmedikation"; der numerische Verbindungsschlüssel scheint ansonsten redundant gegenüber l.FSurogat=m.FDosierplannr
  myFrag rsEi, sql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
   Dim MPNr&, Fldnr%, altlFSur&
   Do While Not rsEi.EOF
    If altlFSur <> rsEi!FDosierplannr Then MPNr = MPNr + 1: Fldnr = 1 Else Fldnr = Fldnr + 1
    ReDim Preserve rMe(UBound(rMe) + 1)
    rMe(UBound(rMe)).Pat_ID = pid
    rMe(UBound(rMe)).Nutzer = rsEi!ua
    rMe(UBound(rMe)).aktZeit = aktZeit
    rMe(UBound(rMe)).Zeitpunkt = rsEi!Zp
    rMe(UBound(rMe)).Datum = rsEi!Zp
    rMe(UBound(rMe)).MPNr = MPNr
    rMe(UBound(rMe)).FeldNr = Fldnr
    If rsEi!FVerbindungsschluessel <> "#TEXT#" Or rsEi!FHinweise <> "" Then
     rMe(UBound(rMe)).Medikament = doUmwfSQL(rsEi!FMedikamentname, True)
     rMe(UBound(rMe)).MedAnfang = GetMed(rMe(UBound(rMe)).Medikament, 0).Value
    Else
     rMe(UBound(rMe)).Bemerkung = doUmwfSQL(rsEi!FMedikamentname, True) ' Inhalt einer Freitextzeile
    End If
    rMe(UBound(rMe)).Wirkstoff = doUmwfSQL(rsEi!FWirkstoff, True)
    If rsEi!FPzn > 0 Then rMe(UBound(rMe)).PZN = rsEi!FPzn
    rMe(UBound(rMe)).Grund = doUmwfSQL(rsEi!FGrund, True)
    rMe(UBound(rMe)).Einheit = rsEi!FPackeinheit
    rMe(UBound(rMe)).Form = rsEi!FDosierform
    rMe(UBound(rMe)).Stärke = rsEi!FWirkstaerke
    On Error Resume Next
    rMe(UBound(rMe)).Menge = rsEi!FStartmenge
    On Error GoTo fehler
    altlFSur = rsEi!FDosierplannr
    If rsEi!fm <> "" Then
     Call ParseMemo(rsEi!fm, FMem(), obDebug, "FMemo aus meddosis")
     For j = 0 To UBound(FMem)
      Select Case FMem(j).ENr
       Case "2": rMe(UBound(rMe)).Einheit = FMem(j).Text
       Case "4": rMe(UBound(rMe)).mo = FMem(j).Text
       Case "5": rMe(UBound(rMe)).mi = FMem(j).Text
       Case "6": rMe(UBound(rMe)).nm = FMem(j).Text
       Case "7": rMe(UBound(rMe)).ab = FMem(j).Text
       Case "8": rMe(UBound(rMe)).Zn = FMem(j).Text
       Case "9": rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & "geb.ZZ: " & doUmwfSQL(FMem(j).Text, True)
       Case "12": rMe(UBound(rMe)).Datum = stzk(FMem(j).Text)
       Case "14": rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & "Komm.: " & doUmwfSQL(FMem(j).Text, True)
       Case "18": rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & "Abghw: " & doUmwfSQL(FMem(j).Text, True)
      End Select
     Next j
    End If ' rsEi!fm <> ""
    If rsEi!fdt <> "" Then rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & doUmwfSQL(rsEi!fdt, True)
    If rsEi!FHw <> "" Then rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & doUmwfSQL(rsEi!FHw, True)
    If rsEi!FBm <> "" Then rMe(UBound(rMe)).Bemerkung = IIf(rMe(UBound(rMe)).Bemerkung = "", "", rMe(UBound(rMe)).Bemerkung & "; ") & "Bem: " & doUmwfSQL(rsEi!FBm, True)
    
    If rMe(UBound(rMe)).mo = "" Then If rsEi!FModosis <> -1.7E+308 Then rMe(UBound(rMe)).mo = rsEi!FModosis
    If rMe(UBound(rMe)).mi = "" Then If rsEi!Fmidosis <> -1.7E+308 Then rMe(UBound(rMe)).mi = rsEi!Fmidosis
    If rMe(UBound(rMe)).nm = "" Then If rsEi!Fnmdosis <> -1.7E+308 Then rMe(UBound(rMe)).nm = rsEi!Fnmdosis
    If rMe(UBound(rMe)).ab = "" Then If rsEi!Fabdosis <> -1.7E+308 Then rMe(UBound(rMe)).ab = rsEi!Fabdosis
    If rMe(UBound(rMe)).Zn = "" Then If rsEi!FNadosis <> -1.7E+308 Then rMe(UBound(rMe)).Zn = rsEi!FNadosis
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then

  Call MOLeistungen(fPtNr, pid)
  
  syscmd 4, "bearbeite AUen"
' AU
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, FICdcode Art," & _
  "REPLACE(COALESCE(REPLACE(MID(fdetails,INSTR(fdetails,'ext ""')+5,LENGTH(fdetails)-2-INSTR(fdetails,'ext ""')-5),'\n','; '),FText),'''','\''') FText," & _
  "FEintragsart, 18900101+INTERVAL FAnorddatum DAY+INTERVAL FAnordzeit SECOND AnZp," & _
  "COALESCE(na.FInitialen,'') ua, COALESCE(nb.FInitialen,'') ub, MID(ftext,4,1) Art, IF(MID(ftext,4,1)='E',MID(ftext,INSTR(ftext,' ')+1,8),'') von, MID(ftext,INSTR(ftext,'- ')+2,8) bis, MID(ftext,INSTR(ftext,': ')+2) diag " & _
  "FROM ltag f " & vbCrLf & _
  "LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat " & _
  "LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat " & _
  "WHERE FPatnr = " & fPtNr & _
  " AND FEintragsart=19" & _
  " AND fbehgrundnr<=0"
  myFrag rsEi, sql, adOpenStatic, MOCon ' AU
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
    ReDim Preserve rAu(UBound(rAu) + 1)
    rAu(UBound(rAu)).aktZeit = aktZeit
    rAu(UBound(rAu)).Pat_ID = pid
    rAu(UBound(rAu)).Zeitpunkt = rsEi!Zp ' rsEi!anzp ' 29.6.25 korrigiert
    rAu(UBound(rAu)).Ersteller = rsEi!ua
    rAu(UBound(rAu)).Änderer = rsEi!ub
    rAu(UBound(rAu)).art = rsEi!art
    rAu(UBound(rAu)).Beginn = Left$(rsEi!VoN, 2) & Mid(rsEi!VoN, 4, 2) & "20" & Mid$(rsEi!VoN, 7, 2)
    rAu(UBound(rAu)).Ende = Left$(rsEi!Bis, 2) & Mid(rsEi!Bis, 4, 2) & "20" & Mid$(rsEi!Bis, 7, 2)
    rAu(UBound(rAu)).ICDs = rsEi!Diag
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
  
' Blutdruck
#If False Then
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, REGEXP_REPLACE(FICdcode,'(\w+)#\1','\1') Art, FText" & vbCrLf & _
  ", REGEXP_REPLACE(FText,'^(?>[^0-9]|[4-9](?![0-9])|[0-2](?![0-9]{2}))*\b((?:[4-9][0-9]|[0-3][0-9]{2})(?:-[0-9]{2,3}){0,2}) */? *(?:über )?((?:[3-9][0-9]|[0-2][0-9]{2})(?:-[0-9]{2,3}){0,2})?(?:(?:[^PH]|H(?!F))*(?:Puls|P(?=[0-9 :.])|HF))?:? *([0-9]{1,3}(?:-[0-9]{2,3})?)? *(.*)','\1‡\2‡\3‡\4') Erg" & vbCrLf & _
  ", REGEXP_REPLACE(ftext,'^(?:[^l]|l(?!e))*(?:let?zten *(\d{1,3}))?.*$','\1') zahl" & vbCrLf & _
  ", FEintragsart, 18900101+INTERVAL FAnorddatum DAY+INTERVAL FAnordzeit SECOND AnZp" & vbCrLf & _
  ", COALESCE(na.FInitialen,'') ua, COALESCE(nb.FInitialen,'') ub " & vbCrLf & _
  "FROM ltag f " & vbCrLf & _
  "LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat " & vbCrLf & _
  "LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat " & vbCrLf & _
  "WHERE FPatnr = " & fPtNr & vbCrLf & _
  " AND REGEXP_REPLACE(FICdcode,'(\w+)#\1','\1') IN ('RR','RRVGL')" & vbCrLf & _
  " AND ((FEintragsart=5 AND FStatus=0))" & vbCrLf & _
  " AND fbehgrundnr<=0"
  myFrag rsEi, sql, adOpenStatic, MOCon
  Dim Spl$()
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
    messDatum = rsEi!Zp ' rsEi!anzp
    art = rsEi!art
    Call aufSplit(rsEi!erg, "‡")
    Call RREintr
    rRr(UBound(rRr)).RR = rsEi!FText ' REPLACE$(rsEi!erg, "‡", " ")
    rRr(UBound(rRr)).RRsyst = Arra(0)
    rRr(UBound(rRr)).RRdiast = Arra(1)
    rRr(UBound(rRr)).Puls = IIf(Arra(2) = "", 0, Arra(2))
    rRr(UBound(rRr)).Bemerkung = doUmwfSQL(Arra(3), True)
    rRr(UBound(rRr)).Quelle = "MO"
    If rsEi!Zahl <> "" Then rRr(UBound(rRr)).RRzahl = rsEi!Zahl
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
#End If

' Notizen
'  doNotizen fPtNr, False ' kommen in allespeichern

' andere Einträge, Desktop-Notizen, RR, DMP-Reihe (2) (Schluss ist ca. nach 275 Zeilen)
' zuvor (hat Bruchteile verschluckt): '^.*?(?:Ewert ""?(?:Dieser Eintrag wurde manuell erzeugt.|([\d,]+(?:\b|\.?\d+?))\.?0*\b(?:\\n(?:\\""|\\|[^""])*)?)""?(?#<- hintere 0er löschen).*(?:Einheit( )""((?:\\""|\\|[^""])*)"")?|\((?:T|Et)ext ""((?:\\""|\\|[^""])*)"").*$'
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp" & vbCrLf & _
  ", IF (FText RLIKE '^(\w+)#\1:', REGEXP_REPLACE(FText,'(\w+)#\1:.*','\1'),REGEXP_REPLACE(FICdcode,'(\w+)#\1','\1')) Art" & vbCrLf & _
  ", REPLACE(REPLACE(REPLACE(IF(rWert=FDetails OR rWert IS NULL,FDet,rWert),'\t',''),'\r',''),'\n',' ') Wert, FDet, FICDCode, FEintragsart, 18900101+INTERVAL FAnorddatum DAY+INTERVAL FAnordzeit SECOND AnZp" & vbCrLf & _
  ", COALESCE(na.FInitialen,'') ua, COALESCE(nb.FInitialen,'') ub, l.FLstgerbnr, FText, FDetails" & vbCrLf & _
  ", REGEXP_REPLACE(FDet,'^(?>[^0-9]|[4-9](?![0-9])|[0-2](?![0-9]{2}))*\b((?:[4-9][0-9]|[0-3][0-9]{2})(?:-[0-9]{2,3}){0,2}) */? *(?:über )?((?:[3-9][0-9]|[0-2][0-9]{2})(?:-[0-9]{2,3}){0,2})?(?:(?:[^PH]|H(?!F))*(?:Puls|P(?=[0-9 :.])|HF))?:? *([0-9]{1,3}(?:-[0-9]{2,3})?)? *(.*)','\1‡\2‡\3‡\4') FArray" & vbCrLf & _
  ", REGEXP_REPLACE(FDet,'^(?:[^l]|l(?!e))*(?:let?zten *(\d{1,3}))?.*$','\1') zahl" & vbCrLf & _
  ", IF(INSTR(FDet,'exportiert')=0,0,REGEXP_REPLACE(FDet,'.*am ([^)]*)\).*','\1')) exp" & vbCrLf & _
  "FROM (" & vbCrLf & _
  " SELECT REPLACE(IF(INSTR(FDetails,'text ""'),MID(FDetails,LOCATE('text',FDetails)+LENGTH('text')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('text',FDetails)+LENGTH('text')+2)-LOCATE('text',FDetails)-LENGTH('text')-2),FText),'''','\''') FDet" & vbCrLf & _
  ", ltag.*" & vbCrLf & _
  ",REGEXP_REPLACE(FDetails,'^.*?(?:Ewert ""?(?:Dieser Eintrag wurde manuell erzeugt.|(\d+(?:[^0-9,.]|[.,]\d+?))0*\b(?:\\n(?:\\""|\\|[^""])*)?)""?(?#<- hintere 0er löschen).*(?:Einheit( )""((?:\\""|\\|[^""])*)"")?|\((?:T|Et)ext ""((?:\\""|\\|[^""])*)"").*$','\1\2\3\4') rWert" & vbCrLf & _
  " FROM ltag) l" & vbCrLf & _
  "LEFT JOIN nutzerneu na ON FAnordnutzernr = na.FSurogat" & vbCrLf & _
  "LEFT JOIN nutzerneu nb ON FAusfnutzernr = nb.FSurogat" & vbCrLf & _
  "WHERE FPatnr = " & fPtNr & vbCrLf & _
  " AND ((FEintragsart=5 AND FStatus=0) OR FEintragsart IN (8,9,10,11,151) OR FEintragsart>1000)" & vbCrLf & _
  "-- AND REGEXP_REPLACE(FICdcode,'(\w+)#\1','\1') IN ('RR','RRVGL')" & vbCrLf & _
  " AND FBehgrundnr<=0"
' FEintragsart>1000 sind die selbst definierten Kategorien in mosystem
  myFrag rsEi, sql, adOpenStatic, MOCon ' Einträge
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
'    If rsEi!Wert Like "Barthel-Index bei Schade, Gerald, geb. 17.12.1962*" Then Stop
'    If rsEi!Wert Like "Time-up-and-Go-Test bei Schade, Gerald, *" Then Stop
'    Debug.Print "Eintragsart: " & rsEi!FEintragsart
    messDatum = rsEi!Zp ' rsEi!anzp ' umgestellt 29.6.25
    art = rsEi!art
    neuart = 0
'    If rsEi!FEintragsart = 1129 Then Stop
'    If rsEi!FEintragsart = 1105 Then Stop
'    If rsEi!FEintragsart = 1058 Then Stop ' andm2
    If art = "" Then ' bei Kategorien die art aus mosystem holen
     Set EintS = New SortierEintr
     EintS.TypNr = rsEi!FEintragsart
     Set EintS = EinL.GetItem(EintS)
     If Not EintS Is Nothing Then
      art = EintS.art
      If art = "dak" Then
       If rsEi!FDet Like "Wie oft am Tag müssen Sie Wasser lassen*" Then
        art = "dakluts"
       ElseIf rsEi!FDet Like "Bekommen Sie regelmäßig nach bestimmten Gehstrecken*" Then
        art = "dakap"
       ElseIf rsEi!FDet Like "Schweißtest: *" Then
        art = "daknp"
       End If ' rsEi!FDet Like "Wie
      End If ' art = "dak" Then
      neuart = True
'     ElseIf rsEi!FEintragsart = 1053 Then ' Überweisungstext ' ist bei EinK dabei => "utxt"
'      art = "ütxt"
'     ElseIf rsEi!FEintragsart = 1105 Then ' Infos
'      art = "info"
     Else ' Not EintS Is Nothing Then
      Set EintS = New SortierEintr
      EintS.TypNr = rsEi!FEintragsart ' Fehler bei: 1138
      Set EintS = EinK.GetItem(EintS)
'      Debug.Print EintS.Kürz
'      On Error Resume Next ' , Fehler bei 70328
      If EintS Is Nothing Then
       Debug.Print "Eintragsart nicht gefunden: " & rsEi!FEintragsart
      Else ' EintS Is Nothing Then
       Select Case EintS.TypNr
        Case 1138
        Case Else
         art = EintS.Kürz
       End Select
      End If ' EintS Is Nothing Then Else
'      On Error GoTo fehler
     End If ' Not EintS Is Nothing Then
    End If ' rEi(UBound(rEi)).Art = "" Then
'    If Art Like "VKGD*" Then
'     ReDim Preserve rVk(UBound(rVk) + 1)
'    Else
    Select Case UCase$(art)
     Case "ti"
       rNa(0).notiz = IIf(rNa(0).notiz = "", "", rNa(0).notiz & vbCrLf) & UmwfSQL(REPLACE$(REPLACE$(rsEi!FDet, "\n", ""), "\r", vbCrLf))
     Case "ICON" ' Desktop-Notiz
      ReDim Preserve rDe(UBound(rDe) + 1)
      rDe(UBound(rDe)).Pat_ID = pid
      rDe(UBound(rDe)).aktZeit = aktZeit
      rDe(UBound(rDe)).Titel = UmwfSQL(rsEi!FDet)
      rDe(UBound(rDe)).erstZP = rsEi!Zp
      If SafeArrayGetDim(aDesk) <> 0 Then
       Dim k&
       For k = 0 To UBound(aDesk)
        If pid = aDesk(k).Pat_ID And Format$(rDe(UBound(rDe)).erstZP, "yyyymmddhhmm") = Format$(aDesk(k).erstZP, "yyyymmddhhmm") Then
'        And InStrB(rDe(UBound(rDe)).Titel, aDesk(k).Titel) <> 0 Then
         rDe(UBound(rDe)).absPos = aDesk(k).absPos
         rDe(UBound(rDe)).erstZP = aDesk(k).erstZP
         rDe(UBound(rDe)).Titel = UmwfSQL(aDesk(k).Titel)
         rDe(UBound(rDe)).aktZeit = aDesk(k).aktZeit
         rDe(UBound(rDe)).exoL = aDesk(k).exoL
         rDe(UBound(rDe)).hideT = aDesk(k).hideT
         rDe(UBound(rDe)).iconPath = aDesk(k).iconPath
         rDe(UBound(rDe)).IDS = aDesk(k).IDS
         rDe(UBound(rDe)).noteBkColor = aDesk(k).noteBkColor
         rDe(UBound(rDe)).noteFgColor = aDesk(k).noteFgColor
         rDe(UBound(rDe)).positionBottom = aDesk(k).positionBottom
         rDe(UBound(rDe)).positionLeft = aDesk(k).positionLeft
         rDe(UBound(rDe)).positionRight = aDesk(k).positionRight
         rDe(UBound(rDe)).positionTop = aDesk(k).positionTop
         rDe(UBound(rDe)).showAsNote = aDesk(k).showAsNote
         rDe(UBound(rDe)).syncInfoList = aDesk(k).syncInfoList
         rDe(UBound(rDe)).toolTipText = aDesk(k).toolTipText
         rDe(UBound(rDe)).verankert = aDesk(k).verankert
        End If
       Next k
      End If ' SafeArrayGetDim(aDesk) <> 0 Then
     Case "RR", "RRVGL"
      Call aufSplit(rsEi!FArray, "‡")
      Call RREintr
      rRr(UBound(rRr)).art = art
      rRr(UBound(rRr)).RR = doUmwfSQL(rsEi!FDet, True) ' REPLACE$(rsEi!FArray, "‡", " ")
      If IsNumeric(Arra(0)) Then rRr(UBound(rRr)).RRsyst = Arra(0)
      If ArraInd > 0 Then If IsNumeric(Arra(1)) Then rRr(UBound(rRr)).RRdiast = Arra(1)
      On Error Resume Next
      If ArraInd > 1 Then rRr(UBound(rRr)).Puls = IIf(Arra(2) = "", 0, Arra(2)) ' 73-84
      On Error GoTo fehler
      If ArraInd > 2 Then rRr(UBound(rRr)).Bemerkung = doUmwfSQL(CStr(Arra(3)), True)
      rRr(UBound(rRr)).Quelle = "MO"
      If rsEi!Zahl <> "" Then rRr(UBound(rRr)).RRzahl = rsEi!Zahl
'      rInh = rsEi!Wert
'      rRr(UBound(rRr)).Bemerkung = doUmwfSQL(rsEi!FDet, True); Kommentar 21.9.25, s. .RR
      
'      Debug.Print ""
'      Debug.Print "FDet: " & rsEi!FDet
'      Debug.Print "Wert: " & rsEi!Wert
'      Debug.Print "FArray:" & rsEi!FArray
'      Debug.Print "Zahl:" & rsEi!Zahl
'      Debug.Print "FText:" & rsEi!FText
'      Debug.Print "FDetails: " & rsEi!FDetails
'      Debug.Print "FICDCode: " & rsEi!FICdcode
'      Debug.Print "Art: " & rsEi!art
'      Debug.Print "rInh: " & rInh
'      Puls = holPuls(rInh, Bem) ' ändert u.U. rInh
'      Debug.Print "Puls: " & Puls
'      rInh = rsEi!FDet
'      Bem = rInh
'      If InStrB(rInh, "P") <> 0 Then Stop
'      RREintr
'      rRr(UBound(rRr)).RR = rInh
'      If Bem <> rInh And InStr(Bem, rInh) = 1 Then Bem = Mid$(Bem, Len(rInh) + 1)
'      rRr(UBound(rRr)).Bemerkung = Bem
'      If IsNumeric(Puls) Then rRr(UBound(rRr)).Puls = Puls
     Case Else ' dmpreihe, Eintraege
     'dmpreihe (2); Doppeleinträge gemäß Index "eindeutig" vermeiden
      DMPArt = 0
      If (rsEi!FIcdcode Like "*dmp*" And rsEi!FIcdcode <> "DMPERG") Or _
      (UCase$(art) = "TEXT" And InStrB(rsEi!FDet, "dokumentation") <> 0 And InStrB(rsEi!FText, "dmp") <> 0) Then
       Select Case rsEi!art
        Case "DMPDTYP1", "EDMPDM1": DMPArt = 1
        Case "DMPDTYP2", "EDMPDM2": DMPArt = 2
        Case "DMPKHK", "EDMPKHK": DMPArt = 3
        Case "EDMPCOPD": DMPArt = 4
        Case "EDMPAB": DMPArt = 5
        Case "DMPKHK": DMPArt = 3
        Case "DMPKHK": DMPArt = 3
       End Select
       
       For ij = 1 To UBound(rDm) ' um dem eindeutigen Index gerecht zu werden
        If rDm(ij).Pat_ID = pid And rDm(ij).DMPArt = DMPArt And rDm(ij).DokuDatum = messDatum Then
         rj = ij
         GoTo gef2
        End If
       Next ij
       rj = UBound(rDm) + 1
       ReDim Preserve rDm(rj)
       rDm(rj).Pat_ID = pid
       rDm(rj).DMPArt = DMPArt
       rDm(rj).DokuDatum = messDatum
gef2:
       rDm(rj).aktZeit = aktZeit
       rDm(rj).Pat_ID = pid
       
       If rsEi!FIcdcode Like "*dmp*" And rsEi!FIcdcode <> "DMPERG" Then
'        Debug.Print rsEi!ficdcode, rsEi!Wert
        rDm(rj).Abk = art
        If InStrB(rsEi!FDet, "Erst") Then rDm(rj).art = "ED" Else rDm(rj).art = "FD"
        rDm(rj).exportiert = CDate(rsEi!Exp)
'        rDm(rj).DokuDatum = messDatum ' ist schon oben
        rDm(rj).KarteiDatum = messDatum
        rDm(rj).Ok = InStrB(rsEi!FDet, "(ok")
        rDm(rj).ausgedruckt = InStrB(rsEi!FDet, "ausgedruckt")
       ElseIf UCase$(art) = "TEXT" And InStrB(rsEi!FDet, "dokumentation") <> 0 And InStrB(rsEi!FText, "dmp") <> 0 Then
        pos = InStr(rsEi!FArray, "#")
        If pos > 0 Then rDm(rj).Abk = Left$(rsEi!FArray, pos - 1)
        rDm(rj).art = IIf(InStrB(rsEi!FDet, "Erst"), "ED", "FD")
        rDm(rj).ausgedruckt = IIf(InStrB(rsEi!FDet, "ausgedruckt"), 1, 0)
        pos = 1
        Do
         pneu = InStr(pos + 1, rsEi!FDet, "exportiert am")
         If pneu = 0 Then Exit Do Else pos = pneu
        Loop
        If pos > 1 Then rDm(rj).exportiert = CDate(Mid$(rsEi!FDet, pos + 14, 10))
        rDm(rj).KarteiDatum = CDate(rsEi!Zp)
        rDm(rj).obvoll = IIf(InStrB(rsEi!FDet, "vollständig"), 1, 0)
        rDm(rj).Ok = IIf(InStrB(rsEi!FDet, "(ok"), 1, 0)
 '       rDm(rj).lanrid = rsEi!farztnr
        Select Case rsEi!FLstgerbnr
         Case 2: rDm(rj).lanrid = 1 ' Schade
         Case 3: rDm(rj).lanrid = 2 ' Kothny
         Case 4: rDm(rj).lanrid = 5 ' Hammerschmidt
         Case Else: rDm(rj).lanrid = 4 ' unbek
        End Select
        rDm(rj).Nachname = rNa(0).Nachname
        rDm(rj).Vorname = rNa(0).Vorname
        rDm(rj).GebDat = rNa(0).GebDat
       End If ' rsEi!FIcdcode Like "*dmp*" And rsEi!FIcdcode <> "DMPERG" Then else
       
      Else ' (rsEi!FIcdcode Like "*dmp*" And rsEi!FIcdcode <> "DMPERG") Or _
      (UCase$(art) = "TEXT" And InStrB(rsEi!FDet, "dokumentation") <> 0 And InStrB(rsEi!FText, "dmp") <> 0) Then
      ' Einträge
       ReDim Preserve rEi(UBound(rEi) + 1)
       rEi(UBound(rEi)).aktZeit = aktZeit
       rEi(UBound(rEi)).Pat_ID = pid
       rEi(UBound(rEi)).Zeitpunkt = messDatum
       rEi(UBound(rEi)).QS = ZQSort(rEi(UBound(rEi)).Zeitpunkt)
       rEi(UBound(rEi)).QT = ZQuart(rEi(UBound(rEi)).Zeitpunkt)
       rEi(UBound(rEi)).Ersteller = rsEi!ua
       rEi(UBound(rEi)).Änderer = rsEi!ub
'      If InStrB(rsEi!fdet, "Lexotanil und") <> 0 Then Stop
'       rEi(UBound(rEi)).Inhalt = doUmwfSQL(REPLACE$(REPLACE$(rsEi!FDet, "\n", " "), "\r", ""), True)
       Debug.Print rsEi!Wert
       rEi(UBound(rEi)).Inhalt = doUmwfSQL(IIf(rsEi!Wert = "", rsEi!FDet, rsEi!Wert), True)
       rEi(UBound(rEi)).absPos = IIf(neuart <> 0, -1, 1)
'      If art = "usdm2" Then Stop
       If art <> "" Then
        If art = "TEXT" Then
         If rEi(UBound(rEi)).Inhalt Like "*[#]*:*" Then
          apos = InStr(rEi(UBound(rEi)).Inhalt, "#")
          If apos > 1 Then ' z.B. LF#LF nach Datenübertragung
           a1 = Left$(rEi(UBound(rEi)).Inhalt, apos - 1)
           a2 = Mid$(rEi(UBound(rEi)).Inhalt, apos + 1, InStr(apos + 1, rEi(UBound(rEi)).Inhalt, ":") - apos - 1)
           If a1 = a2 Then
            rEi(UBound(rEi)).art = a1
            apos = InStr(rEi(UBound(rEi)).Inhalt, ":")
'           If InStrB(rsEi!fdet, "Lexotanil und") <> 0 Then Stop
            rEi(UBound(rEi)).Inhalt = Trim$(Mid$(rEi(UBound(rEi)).Inhalt, apos + 1))
           End If ' a1 = a2 Then
          End If ' pos > 1 Then
         End If ' Inhalt Like "*[#]*:*" Then
        Else ' art = "TEXT" Then
         apos = InStr(art, "#")
         If apos > 1 Then ' z.B. LF#LF nach Datenübertragung
          a1 = Left$(art, apos - 1)
          a2 = Mid$(art, apos + 1)
          If a1 = a2 Then rEi(UBound(rEi)).art = a1
         End If ' pos > 1 Then
        End If ' art = "TEXT" Then else
        If rEi(UBound(rEi)).art = "" Then rEi(UBound(rEi)).art = art
       End If ' art<>""
'      If InStrB(rsEi!fdet, "Lexotanil und") <> 0 Then Stop
       If rEi(UBound(rEi)).art = "GEWICHT" And IsNumeric(rEi(UBound(rEi)).Inhalt) Then rEi(UBound(rEi)).Inhalt = rEi(UBound(rEi)).Inhalt & " kg"
       If rEi(UBound(rEi)).art = "dak" Then
            ' Update Ignore
            ' namen n JOIN eintraege e ON e.pat_id=n.Pat_ID AND e.art='dak'
            ' SET n.dakab=IF(DAYNAME(REPLACE(REPLACE(MID(Inhalt, 30, INSTR(Inhalt,',')-30),',', '.'),':','.'))IS NULL,18991230,STR_TO_DATE(REPLACE(REPLACE(MID(Inhalt,30,InStr(Inhalt,',')-30),',','.'),':','.'),'%d.%m.%y'))
            ' WHERE dakab=18991230 AND IF(DAYNAME(REPLACE(REPLACE(MID(Inhalt, 30, INSTR(Inhalt,',')-30),',', '.'),':','.'))IS NULL,18991230,STR_TO_DATE(REPLACE(REPLACE(MID(Inhalt,30,InStr(Inhalt,',')-30),',','.'),':','.'),'%d.%m.%y'))<>18991230
            ' ;
' Einverständnisse:
' find . -regextype posix-extended -iregex ".*/[^/]*(dak|kkh|tk)[^/]*pdf" -iregex ".*(Zus(s|a)tzt?pg?r?o?gramm|Erklärung|modul|besondere versorg|einverständnis|modell).*" -not -iregex ".*(Fragebogen|dmp-teilnahme|teilnahmebestätigung dmp|krankenpflege|ende.*teilnahme).*"
' dazwischen:
' find . -regextype posix-extended -iregex ".*/[^/]*dak[^/]*pdf" -not -iregex ".*(Zus(s|a)tzt?pg?r?o?gramm|Erklärung|Teilnahme|modul|besondere versorg|änzung zu gesundheitsprogramm|einverständnis|modell).*" -iregex '.*(budak|daknama|gerdakeidel|aidakupinic|karydaki|schweißtest|kostenü|ablehnung|bescheinigung|fragen?bogen|ersatz(karte|beha)|libre|spange|reha|versicherungs|genehmigung|nachweis|krankenpflege|gutachten|befundbericht|hilfsmittel|antwortschreiben|ende der versich|fehlende doku|gültig|fragen|anfrage|zuzahlung|rtcgm|verlaufsdokumentation|unterlagen|attest|ausschreibung|eversense|folgedoku|lesegerät|wegen cgm|behandlungsausweis|sensoren|kh dak onko|cgm system|bestätigung für|medakation|pflegeversicherung|bearbeitet|schutzschuhe|vhk|dmp[ -] ?einschreibung|einschreibung dak|(wg.|bzgl.) dmp|fehlende dmp|(dmp|dak)[ -]beend|schreiben von dak vorsorge|doppeleinschreib|arztwechsel|dmp diverse|bzgl. extra-vorsorgeleistungen|erstdoku|schreiben dak|einschreibung dmp|dmp dak einschreibung|dak( dmp| gesundheit|)
' schreiben|dak dmp|dmp ende|dak te |dak wegen dmp|dak-dmp|anforderung).*'
            If rNa(0).dakab = 0 Then
             If InStr(rEi(UBound(rEi)).Inhalt, ",") > 30 Then
              If IsDate(REPLACE$(REPLACE$(Mid$(rEi(UBound(rEi)).Inhalt, 30, InStr(rEi(UBound(rEi)).Inhalt, ",") - 30), ",", "."), ":", ".")) Then
               rNa(0).dakab = CDate(REPLACE$(REPLACE$(Mid$(rEi(UBound(rEi)).Inhalt, 30, InStr(rEi(UBound(rEi)).Inhalt, ",") - 30), ",", "."), ":", "."))
              End If ' IsDate(replace$(Mid$(rei(ubound(rei)).inhalt, 30, InStr(rei(ubound(rei)).inhalt, ",") - 30), ",", ".")) Then
             End If ' If InStr(rei(ubound(rei)).inhalt, ",") > 30 Then
            End If ' rNa(0).dakab = 0 Then
       End If ' rei(UBound(rei)).art = "dak" Then
      End If ' (rsEi!FIcdcode Like "*dmp*" And rsEi!FIcdcode <> "DMPERG") Or _
      (UCase$(art) = "TEXT" And InStrB(rsEi!Wert, "dokumentation") <> 0 And InStrB(rsEi!FText, "dmp") <> 0) Then
    End Select ' ucase$(art)
'    End If ' Art like ...
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then

  syscmd 4, "bearbeite Erinnerungen"
' Erinnerungen => Desktop
  sql = "SELECT 18900101+INTERVAL FDatum DAY datum, COALESCE(CONVERT(FMemo USING latin1),'') Fm, e.* FROM erinnerung e WHERE FPatnr = " & fPtNr
  myFrag rsEi, sql, adOpenStatic, MOCon ' Einträge
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
    ReDim Preserve rDe(UBound(rDe) + 1)
    rDe(UBound(rDe)).Pat_ID = pid
    rDe(UBound(rDe)).erstZP = rsEi!Datum
    rDe(UBound(rDe)).Titel = doUmwfSQL(rsEi!FText, True)
    If rsEi!fm <> "" Then
     Call ParseMemo(rsEi!fm, FMem(), obDebug, "FMemo von fMemo aus Erinnerung")
     For j = 0 To UBound(FMem)
       Select Case FMem(j).ENr
        Case "2.2.5" ' Datum
         rDe(UBound(rDe)).erstZP = stzk(FMem(j).Text)
        Case "2.2.6": ' Uhrzeit
         'Debug.Print FMem(j).Text
         rDe(UBound(rDe)).erstZP = rDe(UBound(rDe)).erstZP + Asc(Mid(FMem(j).Text, 1)) / 24 + Asc(Mid(FMem(j).Text, 2)) / 24 / 60
'         rDe(UBound(rDe)).erstZP
       End Select
     Next j
    End If ' rsEi!FMemo <> "" Then
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
' Brief an Pat. => Patienteninfo kommt auch noch dazu
' termin in zutun bzw. termin

  syscmd 4, "bearbeite Diagnosen"
  Call MODiagnosen(fPtNr, pid)
  
  syscmd 4, "bearbeite Termine und Zutunliste"
' Termine und Zutunlisten
  sql = "SELECT 18900101 + INTERVAL t.FDatumvon DAY + INTERVAL t.FZeitvon SECOND ab," & vbCrLf & _
        "18900101 + INTERVAL t.FDatumbis DAY + INTERVAL t.FZeitbis SECOND bis," & vbCrLf & _
        "t.fbemerkung, COALESCE(CONVERT(t.FMemo USING latin1),'') Memo" & vbCrLf & _
        ", tz.fname, IF(tz.fvorlauf=-32767,0,tz.FVorlauf)TArt, tz.ftyp, t.*, tz.*" & vbCrLf & _
        "FROM termin t" & vbCrLf & _
        "LEFT JOIN tzone tz ON t.FZonenid=tz.FSurogat" & vbCrLf & _
        "WHERE fpatnr = " & fPtNr & vbCrLf & _
        "ORDER BY ab,bis;"
  myFrag rsEi, sql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
   sql = ""
   myEFrag "DELETE FROM termine WHERE pid = " & pid
   Do While Not rsEi.EOF
    Dim behgr$
    behgr = ""
    If rsEi!memo <> "" Then
     Call ParseMemo(rsEi!memo, FMem(), obDebug, "FMemo von Termine aus mosystem")
     For j = 0 To UBound(FMem)
      Select Case FMem(j).ENr
       Case "20": behgr = FMem(j).Text ' Straße
      End Select
     Next j
    End If ' rsMo!fm <> "" Then
    If sql = "" Then
     sql = "INSERT INTO termine(pid,zp,raum,zusatz,FVorlauf,aktzeit)VALUES"
    Else ' sql = "" Then
     sql = sql & ","
    End If ' sql = "" Then
    sql = sql & "(" & pid & ",'" & rsEi!ab & "','" & rsEi!FName & "','" & doUmwfSQL(rsEi!FBemerkung & IIf(behgr = "", "", "; " & behgr), True) & "'," & rsEi!tart & ",'" & Format(aktZeit, "yyyy-mm-dd HH:MM:SS") & "')"
    rsEi.MoveNext
   Loop
   InsKorr DBCn, sql, rAf
  End If ' Not rsEi.BOF Then

Dim labsql$
' ", REPLACE(IF(IF(INSTR(l.FDetails,'Testname ""'),MID(l.FDetails,INSTR(l.FDetails,'Testname ""')+LENGTH('Testname ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'Testname ""',-1),'""')-1),'')<>'',IF(INSTR(l.FDetails,'Testname ""')<>0,MID(l.FDetails,INSTR(l.FDetails,'Testname ""')+LENGTH('Testname ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'Testname ""',-1),'""')-1),''),IF(INSTR(l.FDetails,'(Text ""')<>0,MID(l.FDetails,INSTR(l.FDetails,'(Text ""')+LENGTH('(Text ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'(Text ""',-1),'""')-1),'')),'''','\''') Testname" & vbCrLf
labsql = _
"SELECT" & vbCrLf & _
" IF(INSTR(FDetails,'Normwertug '),REGEXP_REPLACE(FDetails,'^.*Normwertug ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\) \(Normwertog ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\).*$','\1')," & vbCrLf & _
" IF(REGEXP_REPLACE(REGEXP_REPLACE(normtext,'^- ([^0-9])','\1'),'^(?:[^ ]* ?[^ :]* ?[^:]*:? *)?bis +([0-9.,:]*[0-9.,]+).*$','\1') RLIKE '^[^>]*>=? *([0-9.,]*) *= *pos.*$|^[^-]*- *([0-9.,]*).*$|(1: )?.*<=? *([^ ]*)|^[^>]*>=?.*$',REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(normtext,'^- ([^0-9])','\1'),'.*(grenzwertig|Graubereich|Warnbereich).*','0'),'^([^ ]* ?[^ :]* ?[^:]*:? *)?bis +([0-9.,:]*[0-9.,]+).*$','0'),'^[^>]*>=? *([0-9.,]*) *= *pos.*$','0'),'^([^0-9]*|[^:]*:) *([0-9.,]*) *-.*[0-9].*$','\2'),'(1: )?.*<=?.*','\10'),'.*>=? *([^ ]*).*$','\1'),'')) Normwertug" & vbCrLf & _
",IF(INSTR(FDetails,'Normwertog '),REGEXP_REPLACE(FDetails,'^.*Normwertug ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\) \(Normwertog ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\).*$','\2')," & vbCrLf & _
" IF(REGEXP_REPLACE(REGEXP_REPLACE(normtext,'^- ([^0-9])','\1'),'^(?:[^ ]* ?[^ :]* ?[^:]*:? *)?bis +([0-9.,:]*[0-9.,]+).*$','\1') RLIKE '^[^>]*>=? *([0-9.,]*) *= *pos.*$|^[^-]*- *([0-9.,]*).*$|(1: )?.*<=? *([^ ]*)|^[^>]*>=?.*$',REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(normtext,'^- ([^0-9])','\1'),'^(?:[^ ]* ?[^ :]* ?[^:]*:? *)?bis +([0-9.,:]*[0-9.,]+).*$','\1'),'^[^>]*>=? *([0-9.,]*) *= *pos.*$','\1'),'^[^-]*- *([0-9.,]*).*$','\1'),'(1: )?.*<=? *([^ ]*)','\1\2'),'^[^>]*>=?.*$',CONVERT('8'USING utf8mb4)),'')) Normwertog" & vbCrLf & _
",CONCAT(testhinweis, IF(testid='',IF(INSTR(testhinweis,'manuell'),'',CONCAT(IF(testhinweis='','',', '),'(manuell eingegeben)')),'')) Kommentar" & vbCrLf & _
", i.* FROM(" & vbCrLf & _
"SELECT" & vbCrLf & _
" 18900101+INTERVAL l.FDatum DAY+INTERVAL l.FZeit SECOND Zp" & vbCrLf & _
",REPLACE(IF(INSTR(FDetails,'Testid'),MID(FDetails,LOCATE('Testid',FDetails)+LENGTH('Testid')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('Testid',FDetails)+LENGTH('Testid')+2)-LOCATE('Testid',FDetails)-LENGTH('Testid')-2),''),'''','\''') testid" & vbCrLf & _
",REPLACE(REGEXP_REPLACE(l.FDetails,'.*\(Te(?:stname|xt) ""((?:\\""|\\|[^""])*)"".*','\1'),'''','\''') Testname" & vbCrLf & _
",REPLACE(IF(INSTR(l.FDetails,'Einheit ""'),MID(l.FDetails,INSTR(l.FDetails,'Einheit ""')+LENGTH('Einheit ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'Einheit ""',-1),'""')-1),''),'''','\''') Einheit" & vbCrLf & _
",REPLACE(IF(INSTR(l.FDetails,'Gwi ""'),MID(l.FDetails,INSTR(l.FDetails,'Gwi ""')+LENGTH('Gwi ""'),INSTR(SUBSTRING_INDEX(l.FDetails,'Gwi ""',-1),'""')-1),''),'''','\''') Gwi" & vbCrLf & _
",REPLACE(REPLACE(IF(INSTR(l.FDetails,'Ewert '),MID(l.FDetails,INSTR(l.FDetails,'Ewert ')+LENGTH('Ewert '),INSTR(SUBSTRING_INDEX(l.FDetails,'Ewert ',-1),')')-1),   SUBSTRING_INDEX(IF(INSTR(FDetails,'Etext'),MID(FDetails,LOCATE('Etext',FDetails)+LENGTH('Etext')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('Etext',FDetails)+LENGTH('Etext')+2)-LOCATE('Etext',FDetails)-LENGTH('Etext')-2),''),'\n',1)   ),'''','\'''),',','.') EWert" & vbCrLf & _
",REGEXP_REPLACE(REPLACE(CASE WHEN l.FDetails LIKE '%Normtext%' THEN REGEXP_REPLACE(l.FDetails,'^.*Normtext [""]([^""]*)[""].*$','\1') WHEN l.FDetails like '%Normwertug%Normwertog%' THEN REGEXP_REPLACE(l.FDetails,'^.*Normwertug ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\) \(Normwertog ([0-9]*(?:[.,][0-9]*[1-9]+?)?)[.,]?0*\).*$','\1-\2') ELSE '' END,'\r\n',' '),'\\(?=(?:[^'']|$))','\\\\') Normtext" & vbCrLf & _
",REGEXP_REPLACE(REPLACE(REPLACE(IF(INSTR(FDetails,'Testhinweis'),MID(FDetails,LOCATE('Testhinweis',FDetails)+LENGTH('Testhinweis')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('Testhinweis',FDetails)+LENGTH('Testhinweis')+2)-LOCATE('Testhinweis',FDetails)-LENGTH('Testhinweis')-2),''),'''','\'''),'\r\n',' '),'\\(?=(?:[^'']|$))','\\\\') Testhinweis" & vbCrLf & _
",REGEXP_REPLACE(REPLACE(REPLACE(IF(INSTR(FDetails,'Etext'),MID(FDetails,LOCATE('Etext',FDetails)+LENGTH('Etext')+2,LOCATE('""',REPLACE(FDetails,'\""','\'''),LOCATE('Etext',FDetails)+LENGTH('Etext')+2)-LOCATE('Etext',FDetails)-LENGTH('Etext')-2),''),'''','\'''),'\r\n',' '),'\\(?=(?:[^'']|$))','\\\\') Etext" & vbCrLf & _
",l.FDetails,l.FText,l.FSurogat,l.FPatnr,l.FIcdcode " & vbCrLf & _
"-- ,l.*,na.FUsername ua, nb.FUsername ub" & vbCrLf & _
"FROM ltag l" & vbCrLf & _
"-- LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat" & vbCrLf & _
"-- LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat" & vbCrLf
labsql = labsql & _
"WHERE FEintragsart IN (5,7)" & vbCrLf & _
"-- AND fstatus IN (0,2)" & vbCrLf & _
" AND l.fpatnr=" & fPtNr & vbCrLf & _
") i" & vbCrLf & _
" WHERE (testid<>'' OR (testid='' AND INSTR(FDetails,'Etext ""')=0 " & vbCrLf & _
"    AND i.ftext NOT RLIKE 'Bltdruck|Blutdruck|Gewicht|Puls|Größe|umfang|temperatur|caro|sono|Body-Mass|angd|aufgd|bzvgl'))" & vbCrLf & _
" ORDER BY i.FSurogat,Zp" & vbCrLf

If LaborLangsam Then
LaborLangsam:
 If Not ohneLabor Then
  syscmd 4, "bearbeite Labor einzeln"
' Änderungen des folgenden SQL-Textes koordinieren mit "bearbeite Labor gesammelt"

' witzigerweise verzichtet das Programm mit dieser vorausgeschalteten Schleife auf lange Wartezeit (?!)
' dazu ist sowohl der doppelte Aufruf als auch der doppelte Durchlauf nötig
'  myFrag rsEi, labsql, adOpenStatic, MOCon
'  If Not rsEi.BOF Then
'   Do While Not rsEi.EOF
'    rsEi.MoveNext
'   Loop
'  End If
' P.S.: noch witzigererweise fehlt dann das ganze Labor!
  Set rsEi = Nothing
  myFrag rsEi, labsql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
    Dim ls&
    ReDim Preserve rLa(UBound(rLa) + 1): ls = UBound(rLa)
    rLa(ls).Pat_ID = pid
    rLa(ls).Zeitpunkt = rsEi!Zp
'    If Int(rLa(ls).Zeitpunkt) = #12/3/2024# Then Stop
    rLa(ls).FertigStGrad = "E" ' ergänzt 26.3.25
'   rLa(ls).Labor = AbküLabor
    rLa(ls).Abkü = IIf(rsEi!testid = "", rsEi!FIcdcode, rsEi!testid) ' nauftrag->FSchluessel
    rLa(ls).aktZeit = aktZeit
'    rLa(ls).FID = rFa(UBound(rFa)).FID
'    rLa(ls).absPos = absPos
    rLa(ls).Langtext = rsEi!testname
    rLa(ls).LangtextVW = LTEinfüg&(rLa(ls).Langtext)
    rLa(ls).Kommentar = rsEi!testhinweis
    If rsEi!testid = "" Then ' manueller Eintrag
     If InStrB(rLa(ls).Kommentar, "manuell") = 0 Then
      rLa(ls).Kommentar = rLa(ls).Kommentar & IIf(rLa(ls).Kommentar = "", "", ", ") & "(manuell eingegeben)"
     End If ' InStrB(rLa(ls).Kommentar, "manuell") = 0 Then
    End If ' rsEi!testid = "" Then ' manueller Eintrag
    rLa(ls).KommentarVW = KomEinfüg&(rLa(ls).Kommentar)
    rLa(ls).Einheit = rsEi!Einheit
    rLa(ls).Normber = Trim(REPLACE$(rsEi!normtext, "\", ""))
'    If Left$(rLa(ls).Normber, 1) = "<" Then Stop
    rLa(ls).uNm = Trim(REPLACE$(rsEi!Normwertug, "\", ""))
    rLa(ls).oNm = Trim(REPLACE$(rsEi!Normwertog, "\", ""))
    rLa(ls).NormberVW = nbEinfüg&(rLa(ls).Normber, rLa(ls).uNm, rLa(ls).oNm)
    rLa(ls).AbschlZl = rsEi!etext
    rLa(ls).AbschlZlVW = AZEinfüg&(rLa(ls).AbschlZl)
    rLa(ls).obpath = rsEi!Gwi
'    rLa(ls).Wert = rsEi!ewert
'    While Right$(rLa(ls).Wert, 1) = "0" And Right$(rLa(ls).Wert, 2) <> ".0"
'     rLa(ls).Wert = Left$(rLa(ls).Wert, Len(rLa(ls).Wert) - 1)
'    Wend
    Dim ewert$, ewz$, komz%
'    If IsNumeric(rsEi!ewert) Then
'      rLa(UBound(rLa)).Wert = CDbl(REPLACE$(rsEi!ewert, ".", ",", , 1))
    ewert = rsEi!ewert
'    If InStrB(ewert, "1:10") <> 0 Then Stop
    ewz = ""
    komz = 0
    For i = 1 To Len(ewert)
     buch = Mid(ewert, i, 1)
     If InStrB("0123456789<>:", buch) Then
        ewz = ewz & buch
     ElseIf buch = "," Or buch = "." Then
        komz = komz + 1
        If komz > 1 Then Exit For
        ewz = ewz & ","
     Else
      Exit For
     End If
    Next i
'    If IsNumeric(rsEi!ewert) Then
'      rLa(UBound(rLa)).Wert = CDbl(REPLACE$(rsEi!ewert, ".", ",", , 1))
'    If ewz <> "" And ewz <> "," Then
    If IsNumeric(ewz) Then
     rLa(UBound(rLa)).Wert = CDbl(ewz)
    Else
     rLa(UBound(rLa)).Wert = rsEi!ewert
'     rLa(UBound(rLa)).Wert = ewz
    End If
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
 End If ' not ohnelabor then
 If Not LaborLangsam Then
  Return
 End If ' not laborlangsam
End If ' laborlangsam
 
 Call VorstellSetz(rAna(0).Vorgestellt)
 For i = 1 To UBound(rEi)
  If rEi(i).Zeitpunkt > #6/30/2004# And (VorStDat = 0 Or VorStDat > rEi(i).Zeitpunkt) Then
   VorStDat = rEi(i).Zeitpunkt
  End If
 Next i
 Call alleSpeichern(lies, vonMo:=True, ohneAktDat:=True, ohneLabor:=ohneLabor)
 Call doMarkierungen(fPtNr) ' 20.9.25, statt in PLZ
 
'#If Not laborlangsam Then
If Not LaborLangsam Then
 If Not ohneLabor Then
  syscmd 4, "bearbeite Labor gesammelt"
' Änderungen des folgenden SQL-Textes koordinieren mit "bearbeite Labor einzeln"
' 26.10.25: bei Normwertug und Normwertog wurde jeweils in den zweiten Zeilen die Bedingung eingebaut,
' dass nur bei Passen zu einer der Ersetzungen von Normwertog die Ersetzung vorgenommen wird, sonst '' eingesetzt wird
' weiterhin wurde \r\n bei Normtext, Testhinweis und EText durch '' ersetzt
sql = _
"SELECT " & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborkommentar(kommentar) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT tn))) tn, " & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborkommentar(kommentar) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT kom))) kom, " & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborabschlzl(AbschlZl) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT AbschlZl))) AbschlZl," & vbCrLf & _
"CONCAT('INSERT IGNORE INTO labornormber(uNm,oNm,Normber) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT normt))) normt," & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborparameter(Abkü,Langtext,Einheit,uNm,oNm,NBm,AktZeit) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT part))) part," & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborlangtext(Langtext) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(DISTINCT langt))) langt," & vbCrLf & _
"CONCAT('INSERT IGNORE INTO laborneu(FID,Pat_ID,ZeitPunkt,FertigStGrad,Abkü,LangtextVW,Wert,Einheit,obpath,KommentarVW,AktZeit,AbschlZlVW,NormberVW) VALUES ',TRIM(LEADING ',' FROM GROUP_CONCAT(ges))) ges " & vbCrLf & _
"FROM (" & vbCrLf & _
"SELECT " & vbCrLf & _
"   IF(testname<>'',CONCAT('(''',testname,''')'),'') tn," & vbCrLf & _
"   IF(kommentar<>'',CONCAT('(''',Kommentar,''')'),'') kom," & vbCrLf & _
"   IF(etext<>'',CONCAT('(''',etext,''')'),'') AbschlZl," & vbCrLf & _
"   CONCAT('(''',normwertug,''',''',normwertog,''',''',normtext,''')') Normt," & vbCrLf & _
"   CONCAT('(''',IF(testid='',FICdcode,testid),''',''',Testname,''',''',Einheit,''',''',normwertug,''',''',normwertog,''',''',normtext,''',''',DATE_FORMAT(NOW(),'%y-%m-%d %H:%i:%S'),''')') part," & vbCrLf & _
"   CONCAT('(''',Testname,''')') langt," & vbCrLf & _
"   CONCAT('((SELECT fid FROM faelle WHERE pat_id=',FPatNr,' AND ''',zp,'''>bhfb ORDER BY bhfb DESC LIMIT 1),',FPatNr,',''',COALESCE(zp,'NULL'),''',''E'',''',IF(testid='',FICdcode,testid)," & vbCrLf & _
"    ''',(SELECT LangtextVW FROM laborlangtext WHERE Langtext=''',Testname,''' LIMIT 1),''',IF(CAST(EWert AS DOUBLE)=0 OR EWert RLIKE '[<>:]',CONCAT('',EWert,''),CAST(EWert AS DOUBLE)),''',''',Einheit,''',''',Gwi," & vbCrLf & _
"    ''',(SELECT KommentarVW FROM laborkommentar WHERE Kommentar=''',Kommentar,''' LIMIT 1),'''," & vbCrLf & _
"    DATE_FORMAT(NOW(),'%y-%m-%d %H:%i:%S'),''',(SELECT AbschlZlVW FROM laborabschlzl WHERE AbschlZl=''',etext,''' LIMIT 1),'" & vbCrLf & _
"    '(SELECT NormberVW FROM labornormber WHERE uNm=''',Normwertug,''' AND oNm=''',normwertog,''' AND Normber=''',normtext,''' LIMIT 1)'," & vbCrLf & _
"    ')') ges" & vbCrLf & _
", i.*" & vbCrLf & _
"FROM(" & vbCrLf
sql = sql & labsql & _
") i" & vbCrLf & _
") i" & vbCrLf & _
";"
  Set rsEi = Nothing
  myFrag rsEi, sql, adOpenStatic, MOCon, adLockReadOnly, "9992147483647", rAf
  If Not rsEi.BOF Then
   For i = 0 To 6
    If Not IsNull(rsEi.Fields(i)) Then
     If Right$(rsEi.Fields(i), 7) <> "VALUES " Then
      meldTxt = "Labor gesammelt: Füge Sätze aus " & rsEi.Fields(i).name & " ein ...."
      syscmd 4, meldTxt
      myEFrag rsEi.Fields(i), rAf, DBCn, , ErrNr, ErrDes
      If ErrNr Then
       InsKorr DBCn, rsEi.Fields(i), rAf, ErrDes, , ErrNr
      End If
      meldTxt = "Labor gesammelt: " & rAf & " Sätze aus " & rsEi.Fields(i).name & " eingefügt mit Fehler Nr. " & ErrNr & IIf(ErrNr <> 0, " " & ErrDes, "")
      Debug.Print meldTxt
      syscmd 4, meldTxt
      If rAf = 0 And i = 6 Then
'        doPatvonMO = 3
'        Exit Function
        LaborLangsam = True
        Set rsNa = Nothing ' essentiell
        Set rsFa = Nothing
        Set rsMO = Nothing ' eher überflüssig
        Erase aDesk
        GoTo abermals
#If gehtnicht Then
'        MsgBox (ErrNr & " " & ErrDes & vbCrLf & vbCrLf & "vermutlich reicht Group-concat nicht bei Pat." & pid)
        Open "v:\sqllaborimport.txt" For Output As #250
        Print #250, rsEi.Fields(i)
        Close #250
        If SafeArrayGetDim(roLa) = 0 Then
         Call faelleLaden
         Call laborneuLaden
        End If
        GoSub LaborLangsam
        Call laborparameterSpeichern
        Call laborneuEinf
        Call laborneuSpeichern(lies.dlg.SammelInsert, lies.dlg.BeziehungsfehlerSpeichern)
#End If
'        Stop
      End If
     End If
    End If
   Next i
  End If ' Not rsEi.BOF Then
  End If ' not ohneLabor
 End If ' not laborlangsam
'#End If ' not laborlangsam
  myEFrag "UPDATE namen SET aktzeit=" & Format(aktZeit, "yyyymmddHHMMSS") & " WHERE pat_id=" & rNa(0).Pat_ID, rAf, DBCn, , ErrNr, ErrDes
  If rAf = 1 Then
   syscmd 4, "Fertig mit doPatvonMO " & fPtNr & " auf '" & MOCon.Properties("Server Name") & "'"
  Else
   syscmd 4, "! Fehler bei doPatvonMO " & fPtNr & " auf '" & MOCon.Properties("Server Name") & "' !"
  End If
 Else
  syscmd 4, "doPatvonMO: auf Server '" & Lese.MOServer & "' " & fPtNr & " nicht in patstamm auf '" & MOCon.Properties("Server Name") & "' gefunden!"
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
End Function      ' doPatvonMO(pNr&, pid&)

' in doPatvonMO, holHAausMO
Function setzPid(fPtNr&)
 If Lese.pidoffs = 0 Then
' 6.4.25: wir drehen es um: die fPatNr aus MO wird zur Pid, die Turbomed-Pat_id zu TM_Pat_ID
'  On Error Resume Next
'  pid = DBCn.Execute("SELECT Pat_ID FROM namen WHERE FPatnr=" & fPtNr).Fields(0)
'  On Error GoTo fehler
'  If pid = 0 Then pid = fPtNr
  setzPid = fPtNr
 Else
  setzPid = fPtNr + Lese.pidoffs
 End If ' Lese.pidoffs = 0 Then
End Function ' setzPid(fPtNr&)

' in doPatvonMO und GetHausarzt1
Function holHAausMO(inf As InfoTyp, fPtNr&, Optional satznr%)
  Const obDebug% = True
  Dim rsHa As New ADODB.Recordset, rslue  As New ADODB.Recordset, j&, rAf&
  Dim rHa As ADODB.Recordset
  Dim Fd$
  Dim iru&
  Dim FMem() As memoType ', Kat() As memoType, tKat() As memoType, Abl() As memoType, fAuft() As memoType
  On Error GoTo fehler
  ' -34 Überweiser, -40 Hausarzt, -32 Arzt
  sql = "SELECT COALESCE(p.FArztnralt,'')FArztnralt,COALESCE(p.FAdresse,'')FAdresse,COALESCE(a.FArztgruppe,'')FArztgruppe," & vbCrLf & _
  "     COALESCE(a.FAnrede,'')FAnrede, COALESCE(a.FTitel,'')FTitel," & vbCrLf & _
  "     COALESCE(a.FArztnr,'')FArztnr, COALESCE(a.FNachname,'')FNachname, COALESCE(a.FVorname,'')FVorname," & vbCrLf & _
  "     COALESCE(r.FRelationtyp,'')FRelationtyp,COALESCE(CONVERT(p.FAdresse USING latin1),'')Adr," & vbCrLf & _
  "     COALESCE(IF(a.FAnrede='\N','',a.FAnrede),'')anrk" & vbCrLf & _
      "FROM patrelation r" & vbCrLf & _
      "JOIN earzt a ON r.FReferenztyp=2 AND (a.FSurogat = r.FReferenzid" & vbCrLf & _
      "OR (a.FSurogat <> r.FReferenzid" & vbCrLf & _
      "AND a.FNachname=REGEXP_REPLACE(fdata,'^.*Nachname ""([^""]*)"".*$','\1')" & vbCrLf & _
      "AND a.FVorname =REGEXP_REPLACE(fdata,'^.*Vorname ""([^""]*)"".*$','\1')" & vbCrLf & _
      "AND a.FArztgruppe =REGEXP_REPLACE(fdata,'^.*Namenszusatz ""([^""]*)"".*$','\1')" & vbCrLf & _
      "   ))" & vbCrLf & _
      "LEFT JOIN epraxis p ON a.FExtpraxisnr = p.FSurogat" & vbCrLf & _
      "WHERE fpatid=" & fPtNr & " AND r.FRelationtyp IN (-34,-40,-32)" & _
      "ORDER BY r.FRelationtyp=-40 DESC LIMIT " & (satznr + 1)
  rsHa.Open sql, MOCon, adOpenStatic, adLockReadOnly
  Do While satznr <> 0
   If rsHa.BOF Or rsHa.EOF Then Exit Do
   rsHa.MoveNext
   satznr = satznr - 1
  Loop
  ' wenn nichts gefunden
  If rsHa.BOF Or rsHa.EOF Then
#If altbezieh Then
  ' aus Beziehungspuffer nehmen
   myFrag rHa, "SELECT FData FROM patrelation pr WHERE pr.fpatid=" & fPtNr, adOpenStatic, MOCon
   If Not rHa.BOF Then
    inf.Funktion = "HA"
    Fd = rHa!fdata
    If Fd <> "{}" Then
     Dim datf, pr
     datf = Split(Fd, ") (")
     If UBound(datf) > 0 Then
      datf(0) = REPLACE$(datf(0), "{(", "")
      datf(UBound(datf)) = REPLACE$(datf(UBound(datf)), ")}", "")
      For iru = 0 To UBound(datf)
       pr = Split(datf(iru), " """)
       If UBound(pr) > 0 Then
        pr(1) = Left$(pr(1), Len(pr(1)) - 1)
        Select Case pr(0)
         Case "Nachname": inf.Nachname = pr(1)
         Case "Vorname":  inf.Vorname = pr(1)
         Dim rGs As ADODB.Recordset
         myFrag rGs, "SELECT IF(FAnrede='\N','',FAnrede) anrk FROM earzt WHERE FVorname='" & pr(1) & "' LIMIT 1", adOpenStatic, MOCon
         If Not rGs.EOF Then
          inf.anrede = rGs!anrk
          If inf.anrede <> "" Then
           If inf.anrede = "Frau" Then inf.Überschr = "Sehr geehrte Frau Kollegin" Else inf.Überschr = "Sehr geehrter Herr Kollege"
          End If ' inf.anrede <> "" Then
         End If ' Not rGs.EOF Then
        Case "Titel": inf.gesName = pr(1)
        Case "Geschlecht"
             If pr(1) = "1" Then
               inf.anrede = "Herrn"
               inf.Überschr = "Sehr geehrter Herr Kollege"
             Else
               inf.anrede = "Frau"
               inf.Überschr = "Sehr geehrte Frau Kollegin"
             End If
         Case "Strasse": inf.Straße = pr(1)
         Case "Plz": inf.plz = pr(1)
         Case "Ort": inf.ort = pr(1)
         Case "Fax": inf.Faxnr = pr(1)
         Case "Telefon": inf.TelNr = pr(1)
         Case "Namenszusatz": inf.Fachrtg = pr(1)
        End Select
       End If ' UBound(pr) > 0 Then
      Next iru
      If inf.gesName = "" Then inf.gesName = "Dr. med."
      inf.gesName = Trim$(inf.gesName) & " " & inf.Vorname & " " & inf.Nachname
     End If ' UBound(datf) > 0 Then
     If inf.anrede = "" Then inf.anrede = "Frau/Herrn"
     If inf.Überschr = "" Then inf.Überschr = "Sehr geehrte/r Frau/Herr Kollege/in"
    End If ' Fd <> "{}" Then
   End If ' Not rHa.bof Then
#End If ' altbezieh
  Else ' rHa.BOF Then -> not
   Dim haru%
   haru = 0
   inf.KVNr = ""
 '  0: Frau/Herrn
 '  1: Titel+Vorn+Nachn,
 '  2: Straße,
 '  3: PLZ+Ort,
 '  4: Faxnr,
 '  5: S.g./Liebe,
 '  6: DMPTyp2,
 '  7: DMPTyp1,
 '  8: Niederlassungsgebiet 3. Feld für einmal austauschen,
 '  9: Vorname,
 ' 10: Funktion ("Üw 207, HA"),
 ' 11: Fachrichtung
 ' 12: KV-Nummer
 ' 13: Tel'nr.
 ' 14: Nachname,
 ' 15: Email
   
   Do While Not rsHa.EOF
    haru = haru + 1
    Dim Adr$, Lkz$
    On Error Resume Next
    inf.Nachname = rsHa!fnachname
    inf.Vorname = rsHa!FVorname
    inf.Fachrtg = rsHa!FArztgruppe
    inf.gesName = "Dr.med. " & inf.Vorname & " " & inf.Nachname
    On Error GoTo fehler
    inf.anrede = rsHa!anrk
    If inf.anrede <> "" Then
     inf.Überschr = Switch(inf.anrede Like "Liebe*", inf.anrede, inf.anrede = "Frau", "Sehr geehrte Frau Kollegin", True, "Sehr geehrter Herr Kollege")
    End If
    inf.Funktion = "HA"
    Adr = REPLACE$(REPLACE$(doUmwfSQL(rsHa!Adr, True), "\n", ""), "\r", "")
    If rsHa!farztnralt <> "" Or rsHa!farztnr <> "" Or rsHa!Adr <> "" Then ' 29.9.25
     inf.KVNr = rsHa!farztnralt
     If inf.KVNr = "" Then inf.KVNr = rsHa!farztnr
     If rsHa!Adr <> "" Then
      Lkz = ""
      Call ParseMemo(rsHa!Adr, FMem(), obDebug, "FAdresse aus epraxis")
      For j = 0 To UBound(FMem)
       Select Case FMem(j).ENr
        Case "3": inf.Straße = Trim$(FMem(j).Text) ' Straße
        Case "4": inf.Straße = inf.Straße & " " & FMem(j).Text ' Hausnummer
        Case "5":
            inf.plz = FMem(j).Text ' Postleitzahl
        Case "6":
        
            inf.ort = FMem(j).Text ' Ort
        Case "8": ' "5" ?
        Case "9.1": ' unbekannte ascii-Ziffer
        Case "9.2":
            inf.TelNr = FMem(j).Text ' Telefonnummer
        Case "10.1": ' unbekannte ascii-Ziffer
        Case "10.2":
            inf.Faxnr = REPLACE$(FMem(j).Text, "/", "") ' Faxnummer
        Case "12": Lkz = FMem(j).Text ' Länderkennzeichen
        Case "6.2.3"
'          rRe(UBound(rRe)).anzl = Asc(FMem(j).Text)
        Case Else
'          Debug.Print FMem(j).ENr, FMem(j).Text
       End Select
      Next j ' j = 0 To UBound(FMem)
     End If ' rsHa!Adr <> "" Then
     sql = "SELECT COUNT(0) OVER() zahl, a.* FROM liuez a WHERE kvnr='" & inf.KVNr & "' ORDER BY ID"
     Set rslue = myEFrag(sql, rAf, DBCn)
     If rslue.BOF Then
      sql = "INSERT INTO liuez(kvnr,name,vorname,titelt,fachgruppe,strasse,plz,ort,telefon,fax,anrede,lanr,ursp,aktzeit) VALUES('" & _
           inf.KVNr & "','" & inf.Nachname & "','" & inf.Vorname & "','" & rsHa!FTitel & "','" & rsHa!FArztgruppe & "','" & inf.Straße & "','" & inf.plz & "','" & inf.ort & "','" & inf.TelNr & "','" & inf.Faxnr & "','" & IIf(InStrB(rsHa!fanrede, "Frau") <> 0, "Frau", "Herr") & "','" & rsHa!farztnr & "','" & "MO" & "'," & DatFor_k(aktZeit) & ")"
      InsKorr DBCn, sql, rAf
     Else ' rslue.BOF Then
       Dim rsluezahl&
     ' Update
       rsluezahl = rslue!Zahl
       If rslue!name <> inf.Nachname Or rslue!Vorname <> inf.Vorname Or rslue!titelt <> rsHa!FTitel Or rslue!fachgruppe <> rsHa!FArztgruppe Or rslue!strasse <> inf.Straße Or rslue!plz <> inf.plz Or rslue!ort <> inf.ort Or rslue!telefon <> inf.TelNr Or rslue!fax <> inf.Faxnr Or rslue!anrede <> IIf(InStrB(rsHa!fanrede, "Frau") <> 0, "Frau", "Herr") Or rslue!Lanr <> rsHa!farztnr Then
        sql = "UPDATE liuez SET name='" & inf.Nachname & "',vorname='" & inf.Vorname & "',titelt='" & rsHa!FTitel & "',fachgruppe='" & rsHa!FArztgruppe & "',strasse='" & inf.Straße & "',plz='" & inf.plz & "',ort='" & inf.ort & "',telefon='" & inf.TelNr & "',fax='" & inf.Faxnr & "',anrede='" & IIf(InStrB(rsHa!fanrede, "Frau") <> 0, "Frau", "Herr") & "',lanr='" & rsHa!farztnr & "',ursp='MO',aktzeit=" & DatFor_k(aktZeit) & "" & vbCrLf & _
       "WHERE kvnr='" & inf.KVNr & "'"
       Set rslue = Nothing
       Call myEFrag(sql, rAf, DBCn)
      End If ' rs!name <> inf.NachName
      If rsluezahl > 1 Then
       sql = "DELETE d from liuez d INNER JOIN (SELECT kvnr, id, RANK() OVER(PARTITION BY kvnr ORDER BY id) rang FROM liuez) b ON d.kvnr=b.kvnr AND b.rang=1 AND d.id>b.id WHERE d.kvnr=" & inf.KVNr
       Call myEFrag(sql, rAf, DBCn)
      End If ' rslue!Zahl > 1 Then
     End If ' rslue.BOF Then
    ElseIf Not IsNull(rsHa!FAdresse) Then ' mit COALESCE kommt trotzdem eine Fehlermeldung raus
'    Set rslue = myEFrag(, , DBCn)
     On Error Resume Next
     Set rslue = myEFrag("SELECT kvnr FROM aktlue a WHERE ((nameo ='" & inf.Nachname & "' AND vno='" & inf.Nachname & "') OR (name ='" & inf.Nachname & "' AND vorname='" & inf.Nachname & "')) AND '" & Adr & "' LIKE CONCAT('%',a.strasse,'%')", , DBCn)
     If Not rslue.BOF Then
      inf.KVNr = rslue!KVNr
     End If
     On Error GoTo fehler
    End If ' rsHa!FArztnralt <> "" Then
    If inf.KVNr = "" Then
     Set rslue = myEFrag("SELECT kvnr FROM aktlue a WHERE ((nameo ='" & inf.Nachname & "' AND vno='" & inf.Vorname & "') OR (name ='" & inf.Nachname & "' AND vorname='" & inf.Vorname & "')) AND fachgruppe='" & rsHa!FArztgruppe & "'", , DBCn)
     If Not rslue.BOF Then
      If Not IsNull(rslue!KVNr) Then inf.KVNr = rslue!KVNr
     Else
      Set rslue = myEFrag("SELECT kvnr FROM aktlue a WHERE ((nameo ='" & inf.Nachname & "' AND vno='" & inf.Vorname & "') OR (name ='" & inf.Nachname & "' AND vorname='" & inf.Vorname & "'))", , DBCn)
      If Not rslue.BOF Then
       If Not IsNull(rslue!KVNr) Then inf.KVNr = rslue!KVNr Else inf.KVNr = 0
      End If
     End If
    End If ' KVNr = "" Then
    Select Case haru
     Case 1: rNa(0).KVNr = inf.KVNr
        rNa(0).getHA0 = IIf(inf.KVNr = "", 0, inf.KVNr)
        rNa(0).fnHA0 = "(" & IIf(rsHa!FRelationtyp = -40 Or rsHa!FRelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA0 = rNa(0).fnHA0 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA0 = rNa(0).fnHA0 & ")"
     Case 2: rNa(0).KVNr2 = inf.KVNr
        On Error Resume Next
        rNa(0).getHA1 = inf.KVNr
        On Error GoTo fehler
        rNa(0).fnHA1 = "(" & IIf(rsHa!FRelationtyp = -40 Or rsHa!FRelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA1 = rNa(0).fnHA1 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA1 = rNa(0).fnHA1 & ")"
     Case 3: rNa(0).KVNr3 = inf.KVNr
        If IsNumeric(KVNr) Then rNa(0).getHA2 = inf.KVNr
        rNa(0).fnHA2 = "(" & IIf(rsHa!FRelationtyp = -40 Or rsHa!FRelationtyp = -32, "HA", "Üw")
        If UBound(rFa) <> 0 Then rNa(0).fnHA2 = rNa(0).fnHA2 & " " & Left$(rFa(1).Quartal, 1) & Right$(rFa(1).Quartal, 2)
        rNa(0).fnHA2 = rNa(0).fnHA2 & ")"
     Case 4: rNa(0).KVNr4 = inf.KVNr
    End Select
    If inf.KVNr <> "" And SafeArrayGetDim(rKv) <> 0 Then
     Call addierrKV(setzPid(fPtNr), inf.KVNr, aktZeit, absPos:=0)
    End If ' inf.KVNr <> "" Then
    Exit Do ' 7.9.25
    rsHa.MoveNext
   Loop ' While Not rsHa.EOF
  End If ' Not rsHa.BOF Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in holHAausMO/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' holHAausMO(infos)

' in doPatvonMo, doMarkierungen
Public Sub markAuswert(ByRef rNa() As namen, txt$)
   Select Case txt
    Case "Arzt Schade"
     rNa(0).obs = 1
    Case "Arzt Kothny"
     rNa(0).obk = 1
    Case "Arzt Hammerschmidt"
     rNa(0).obh = 1
    Case "Antikoagulation"
     rNa(0).antikoag = 1
    Case "Diabetes Typ I"
     rNa(0).dmt1 = 1
    Case "Gdm"
     rNa(0).gdm = 1
    Case "kein D.m."
     rNa(0).kdm = 1
' 1=Libre Handy, 2=Libre Gerät, 3=Dexcom Handy, 4=Dexcom Gerät, 5=Simplera
    Case "Libre-Handy"
     rNa(0).cgm = 1
    Case "Libre Lesegerät"
     rNa(0).cgm = 2
    Case "Dexcom-Clarity"
     rNa(0).cgm = 3
    Case "Dexcom Lesegerät"
     rNa(0).cgm = 4
    Case "Medtronic Simplera"
     rNa(0).cgm = 5
    Case "Eversense"
     rNa(0).cgm = 6
' 1=Novopen, 2=Combo, 3=Insight, 4=Kaleido, 5=Medt.780, 6=Omnipod 5, 7=Dash, 8=TSlim, 9=Ypsopump
    Case "Novopen"
     rNa(0).insanw = 1
    Case "Accu Chek Spirit Combo"
     rNa(0).insanw = 2
    Case "AccuChek Insight Yourloops"
     rNa(0).insanw = 3
    Case "Kaleido Glooko"
     rNa(0).insanw = 4
    Case "Medtronic 780 pdf Patient"
     rNa(0).insanw = 5
    Case "Omnipod 5 Glooko"
     rNa(0).insanw = 6
    Case "Omnipod Dash auslesen"
     rNa(0).insanw = 7
    Case "TSlim-Glooko Backoffice auslesen"
     rNa(0).insanw = 8
    Case "Ypsopump-Glooko"
     rNa(0).insanw = 9
    Case "Dana"
     rNa(0).insanw = 10
   End Select ' Txt
End Sub ' markAuswert

' in richtdiag_Click
Public Sub turichtdiag()
 Dim rPt As ADODB.Recordset, rPid As ADODB.Recordset, rAf&, rAf2&, aktz&, dzahl&
 Dim ErrNr&, ErrDes$, altsi$
 On Error GoTo fehler
 Call Lese.ProgStart
 Call MOConInit(, "turichtdiag()")
 altsi = sqlIGNORE
 sqlIGNORE = ""
 ReDim rNa(0)
' ReDim rDi(0)
' Call ForeignNo0
' Call ForeignNo1
 Dim patn$
 Const limit& = 20000 ' 100
' sql = "SELECT GROUP_CONCAT(fpatnr) FROM (SELECT fpatnr FROM behgrund WHERE fstatus<>3 order by fpatnr desc LIMIT " & limit & ")i"
' myFrag rPt, sql, adOpenStatic, MOCon, adLockReadOnly, 1000000, rAf
' If Not rPt.EOF Then
'  patn = rPt.Fields(0)
'  myFrag rPid, "DELETE from diagnosen WHERE pat_id IN (" & patn & ")", adOpenStatic, DBCn, adLockReadOnly, , rAf, , ErrNr, ErrDes
'  AllePat = -1
' End If ' Not rPt.EOF Then
' DBCn.BeginTrans
 myFrag rPt, "SELECT COUNT(0) OVER() zahl, FPatnr FROM behgrund/* WHERE FStatus<>3*/ GROUP BY FPatnr ORDER BY FPatnr DESC LIMIT " & limit, adOpenStatic, MOCon
 If Not rPt.BOF Then
  Do While Not rPt.EOF
   rNa(0).Pat_ID = rPt!FPatNr
   aktz = aktz + 1
'   myFrag rPid, "SELECT 0 FROM faelle WHERE pat_id=" & rNa(0).Pat_ID & " LIMIT 1", adOpenStatic
'   If Not rPid.BOF() Then
    myFrag rPid, "DELETE from diagnosen WHERE pat_id =" & rNa(0).Pat_ID, adOpenStatic, DBCn, adLockReadOnly, , rAf, , ErrNr, ErrDes
    MODiagnosen rNa(0).Pat_ID
    If UBound(rDi) <> 0 Then
     diagnosenSpeichern True, Lese.dlg.BeziehungsfehlerSpeichern, rAf, True
     sql = "UPDATE diagnosen d FORCE INDEX (auswahl) LEFT JOIN faelle f FORCE INDEX (auswahl) ON d.Pat_ID=f.pat_id AND d.diagdatum BETWEEN bhfb AND bhfe1 SET d.fid=f.fid WHERE d.pat_id= " & rNa(0).Pat_ID & " ORDER BY diagdatum DESC;"
     myEFrag sql, rAf2, DBCn, , ErrNr, ErrDes
     ReDim rDi(0)
    End If
'    Debug.Print aktz & "/" & rPt!Zahl, rNa(0).Pat_ID, rAf, rAf2
    Lese.Ausgeb "-> " & aktz & "/" & rPt!Zahl & " " & rNa(0).Pat_ID & " " & rAf & " " & rAf2, 0
    dzahl = dzahl + rAf2
'   End If
   rPt.MoveNext
  Loop
 End If ' Not rPt.BOF Then
 Lese.Ausgeb "Zu " & aktz & " Patienten insgesamt " & dzahl & " Diagnosen übertragen.", True
' DBCn.CommitTrans
 Call ForeignYes0
 Call ForeignYes1
 sqlIGNORE = altsi
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in richtleist/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' richtleist


' in doPatvonMO und richtdiag
Sub MODiagnosen(fPtNr&, Optional pid&)
  On Error GoTo fehler
  If pid = 0 Then pid = fPtNr
' Diagnosen
' FStatus: 1=akut, 5=Dauer, 2=anamnest, 3=historisch, 4=abgeschlossen
' FKlasse: 2=gesichert, 1=V.a., 3=Z.n., 4=Ausschluss
' FEintragsart:
'    1:                ak,       ,A,G,V,Z         H            (711)
'    2:                ab,       ,A,G,V,Z         H         (114107)
'    3:                an,            G,Z         H             (10)
'    4:                hi,          G,V,Z         H             (71)
'   13:    ab,ak,an,da,hi(1x),   ,A,G,V,Z         H           (4630)
'   14:             ak,da,        A,G,V,Z         H            (162)
'   15:                da,              G         H              (2)
'   17:          ab,ak,da,        A,G,V,Z         H            (590)
'   18:          ak,an,da,          G,V,Z         H            (141)
'   19:             ak,da,            G,Z         H             (11)
'   20:             ak,da,         ,G,V,Z         H            (323)
'   21:             ak,da,            G,V         H             (16)
'   22:                da,          G,V,Z         H              (8)
'   23:             ak,da,            G,Z         H             (26)
' 2004:             ak,da,         ,G,V,Z         H            (274)
' 2005:          ak,da,hi,        A,G,V,Z         H            (309)
' 2006:             ak,da,              G         H              (5)
' 2017:                da,       ,A,G,V,Z       H,N         (138268)
'20002:                da,              G         H              (1)
'20004:                da,              G         H              (1)
'20012:                da,            G,V         H              (3)
'27144:             ak,da,           ,G,Z         H             (46)
'27187:                da,              G         H              (5)
'27188:             ak,da,         ,G,V,Z         H            (378)
'27193:                da,              G         H              (1)
'27216:                da,              G         H              (1)
'27217:                da,          G,V,Z         H              (7)
'29957:                da,            G,Z         H             (18)
'30465:                da,              G         H              (3)
'30535:                da,              G         H              (1)
'30540:                da,              G         H              (1)
'30543:                da,              G         H              (1)
' über:
'WITH sel AS (SELECT fbehgrundnr fb,FEintragsart ea from ltag /*WHERE fpatnr=197 *//* AND fbehgrundnr>0*/)
'SELECT GROUP_CONCAT(zusi SEPARATOR '') gzus FROM (
'SELECT
'CONCAT('''',LPAD(ea,5),': '
',LPAD(GROUP_CONCAT(DISTINCT CASE FStatus WHEN 1 THEN 'ak' WHEN 2 THEN 'an' WHEN 3 THEN 'hi' WHEN 4 THEN 'ab' WHEN 5 THEN 'da' ELSE ' ' END),17),','
',LPAD(GROUP_CONCAT(DISTINCT CASE FKlasse MOD 15 WHEN 1 THEN 'V' WHEN 2 THEN 'G' WHEN 3 THEN 'Z' WHEN 4 THEN 'A' ELSE ' ' END),15)
',LPAD(GROUP_CONCAT(DISTINCT CASE FKlasse div 15 WHEN 0 THEN 'H' ELSE 'N' END),10)
',LPAD(CONCAT('(',COUNT(0),')'),17),CHR(13)
') zusi,
'COUNT(0) OVER() gesz, COUNT(0) zahl, FPatnr, 18900101 + INTERVAL FDatum DAY + INTERVAL FZeit SECOND Diagdat,
'CASE F4201 WHEN 1 THEN 'R' WHEN 2 THEN 'L' WHEN 3 THEN 'B' ELSE ' ' END Seite,
'GROUP_CONCAT(DISTINCT CASE FKlasse MOD 15 WHEN 1 THEN 'V' WHEN 2 THEN 'G' WHEN 3 THEN 'Z' WHEN 4 THEN 'A' ELSE ' ' END) Sich,
'GROUP_CONCAT(DISTINCT CASE FKlasse div 15 WHEN 0 THEN 'H' ELSE 'N' END) Kard,
'FIcdcode ICD, lt.ea, COALESCE(IF(RIGHT(FText,1)=0,LEFT(FText,LENGTH(FText)-1),FText),'') FText,
'GROUP_CONCAT(DISTINCT CASE FStatus WHEN 1 THEN 'ak' WHEN 2 THEN 'an' WHEN 3 THEN 'hi' WHEN 4 THEN 'ab' WHEN 5 THEN 'da' ELSE ' ' END) Stat,
'COALESCE(IF(RIGHT(FErlaeuterung,1)=0,LEFT(FErlaeuterung,LENGTH(FErlaeuterung)-1),FErlaeuterung),'') Zus, FNutzernr, FID, FAusnahme
'FROM sel lt INNER JOIN behgrund b on lt.fb=b.FSurogat
'-- WHERE NOT EXISTS (SELECT bi.* FROM sel lti INNER JOIN behgrund bi ON lti.fb=bi.FSurogat
'--   WHERE ficdcode=b.ficdcode AND ftext=b.ftext AND fklasse=b.FKlasse AND f4201=b.f4201 AND ((fstatus IN(1,5)AND b.fstatus IN(2,3,4))OR(fstatus=b.fstatus AND fsurogat>b.fsurogat)))
'GROUP BY ea ORDER BY cast(ea AS INTEGER)
') i;

 Dim rsDi As New ADODB.Recordset
'myFrag rsFa, "SELECT f.fsurogat nix, COALESCE(CONVERT(f.FMemo USING latin1),'') Fm,CONCAT(f.fpatnr,', ',18900101 + INTERVAL f.fvon DAY,' - ',18900101 + INTERVAL f.fbis DAY) ueschr, f.*,a.FBezeichnung, le.FNachname FROM patfall f LEFT JOIN abrechner a ON f.FArztnr=a.FSurogat LEFT JOIN lstgerb le USING (FLstgerbnr) WHERE fpatnr=" & fPtNr & " ORDER BY FVon DESC", adOpenStatic, MOCon, adLockReadOnly
  sql = "WITH sel AS (SELECT fbehgrundnr fb,FEintragsart ea,FIcdcode ICD FROM ltag WHERE fpatnr=" & fPtNr & ")" & vbCrLf & _
  "SELECT FPatnr, 18900101 + INTERVAL FDatum DAY + INTERVAL FZeit SECOND Diagdat," & vbCrLf & _
  "CASE F4201 WHEN 1 THEN 'R' WHEN 2 THEN 'L' WHEN 3 THEN 'B' ELSE ' ' END Seite," & vbCrLf & _
  "CASE (FKlasse MOD 15)MOD 10 WHEN 1 THEN 'V' WHEN 2 THEN 'G' WHEN 3 THEN 'Z' WHEN 4 THEN 'A' ELSE ' ' END Sich," & vbCrLf & _
  "CASE FKlasse DIV 15 WHEN 0 THEN 'H' ELSE 'N' END Kard," & vbCrLf & _
  "ICD, COALESCE(IF(RIGHT(FText,1)=0,LEFT(FText,LENGTH(FText)-1),FText),'') FText," & vbCrLf & _
  "CASE FStatus WHEN 1 THEN 'ak' WHEN 2 THEN 'an' WHEN 3 THEN 'hi' WHEN 4 THEN 'ab' WHEN 5 THEN 'da' ELSE ' ' END Stat," & vbCrLf & _
  "COALESCE(IF(RIGHT(FErlaeuterung,1)=0,LEFT(FErlaeuterung,LENGTH(FErlaeuterung)-1),FErlaeuterung),'') Zus, FNutzernr, FID, FAusnahme, ea" & vbCrLf & _
  "FROM sel lt INNER JOIN behgrund b ON lt.fb=b.FSurogat" & vbCrLf & _
  "WHERE NOT ((FKlasse MOD 15)MOD 10=1 AND FStatus IN(3,4)) AND lt.ICD=b.FIcdcode" & vbCrLf & _
  ";"
  ' => Z.n. V.a. lassen wir weg
  ' akut  =      FStatus 1, stat ak, ea 1
  ' inaktiv =    FStatus 4, stat ab, ea 2
  ' dauer =      FStatus 5, stat da, ea 2017
  ' historisch = FStatus 3, Stat hi, ea 4
  ' abgeschlossen = FStatus 4, Stat ab, ea 2
'   "WHERE NOT EXISTS (SELECT bi.* FROM sel lti INNER JOIN behgrund bi ON lti.fb=bi.FSurogat" & vbCrLf & _
  "  WHERE ficdcode=b.ficdcode AND ftext=b.ftext AND fklasse=b.FKlasse AND f4201=b.f4201 AND ((fstatus IN(1,5)AND b.fstatus IN(2,3,4))OR(fstatus=b.fstatus AND fsurogat>b.fsurogat))" & vbCrLf & _

  ReDim rDi(0)
  myFrag rsDi, sql, adOpenStatic, MOCon
  If Not rsDi.BOF Then
   Do While Not rsDi.EOF
    ReDim Preserve rDi(UBound(rDi) + 1)
    rDi(UBound(rDi)).aktZeit = aktZeit
    rDi(UBound(rDi)).Pat_ID = pid
    rDi(UBound(rDi)).DiagDatum = rsDi!diagdat
    rDi(UBound(rDi)).DiagSicherheit = rsDi!sich
    rDi(UBound(rDi)).DiagText = doUmwfSQL(rsDi!FText, True)
    rDi(UBound(rDi)).DiagSeite = rsDi!Seite
    rDi(UBound(rDi)).DiagAttr = doUmwfSQL(rsDi!Zus, True)
    rDi(UBound(rDi)).ICD = rsDi!ICD
'    If rDi(UBound(rDi)).ICD = "E11.41" Then Stop
    rDi(UBound(rDi)).obDauer = IIf(rsDi!Stat = "ak", 0, 1)
    If rsDi!ea = 2 Or rsDi!Stat = "hi" Or rsDi!Stat = "ab" Then rDi(UBound(rDi)).DiagSicherheit = "Z" ' 11.4.25 ' 20.7.25: abgeschlossene Diagnosen sind auch die Inaktivierten
    rDi(UBound(rDi)).obKasse = IIf(rsDi!Stat = "ak" Or rsDi!Stat = "da", 1, 0)
    rsDi.MoveNext
   Loop
  End If ' Not rsDi.BOF Then
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MODiagnosen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MODiagnosen(fPtNr&)

' nirgends
Sub richtleist()
 Dim rPt As ADODB.Recordset, rPid As ADODB.Recordset, rAf&, aktz&
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 syscmd 4, "richtleist, Lese.ProgStart()"
 Call Lese.ProgStart
 Call MOConInit(, "richtleist()")
 ReDim rNa(0)
 ReDim rLe(0)
 
' Call ForeignNo0
' Call ForeignNo1
' DBCn.BeginTrans
 Call myEFrag("SET GLOBAL foreign_key_checks = 0")
 myFrag rPt, "SELECT COUNT(0) OVER() zahl, FPatnr FROM ltag WHERE FEintragsart=12 GROUP BY FPatnr ORDER BY FPatnr DESC", adOpenStatic, MOCon
 If Not rPt.BOF Then
  Do While Not rPt.EOF
   rNa(0).Pat_ID = rPt!FPatNr ' 139 ' 59284 ' rPt!fPatNr
   aktz = aktz + 1
   myFrag rPid, "SELECT 0 FROM faelle WHERE pat_id=" & rNa(0).Pat_ID & " LIMIT 1", adOpenStatic
   If Not rPid.BOF() Then
    MOLeistungen (rNa(0).Pat_ID)
     If UBound(rLe) <> 0 Then
      leistungenSpeichern True, Lese.dlg.BeziehungsfehlerSpeichern, rAf
      sql = "UPDATE leistungen l FORCE INDEX (pid_zp) LEFT JOIN faelle f FORCE INDEX (auswahl) ON l.Pat_ID=f.pat_id AND l.ZeitPunkt BETWEEN bhfb AND bhfe1 SET l.fid=f.fid WHERE l.pat_id= " & rNa(0).Pat_ID & " ORDER BY zeitpunkt DESC;"
      myEFrag sql, rAf, DBCn, , ErrNr, ErrDes
      ReDim rLe(0)
     End If
    Debug.Print aktz & "/" & rPt!Zahl, rNa(0).Pat_ID, rAf
   End If
   rPt.MoveNext
  Loop
 End If ' Not rPt.BOF Then
 On Error Resume Next
 Call myEFrag("SET GLOBAL foreign_key_checks = 1")
' DBCn.CommitTrans
' Call ForeignYes0
' Call ForeignYes1
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in richtleist/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' richtleist

' wird nirgends aufgerufen
Sub callMOLei()
 Dim rsl As New ADODB.Recordset, i&
 Call MOConInit(, "callMOLei()")
' rsL.Open "select count(0) over() zahl, fpatnr from ltag where feintragsart=12 and fpatnr<68608 group by fpatnr order by fpatnr desc", MOCon, adOpenStatic, adLockReadOnly
 rsl.Open "select count(0) over() zahl, fpatnr from ltag where feintragsart=12 and fpatnr<53194 group by fpatnr order by fpatnr desc", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rsl.EOF
  i = i + 1
  ReDim rLe(0)
  Debug.Print "-----------> " & i & "/" & rsl!Zahl, rsl!FPatNr
  MOLeistungen rsl!FPatNr
  rsl.MoveNext
 Loop
 Debug.Print "Fertig!"
End Sub ' callMOLei()

' in doPatvonMo
Sub MOLeistungen(fPtNr&, Optional pid& = -1)
  Dim rsEi As ADODB.Recordset
  Dim abz%
  Dim rAf&
  On Error GoTo fehler
  syscmd 4, "bearbeite Leistungen"
  
  Dim rab() As Abrtyp
  Dim rAbr As ADODB.Recordset
  myFrag rAbr, "SELECT FSurogat,FBetriebsnr FROM abrechner", adOpenStatic, MOCon ' AU
  Do While Not rAbr.EOF
   If SafeArrayGetDim(rab) = 0 Then ReDim rab(0) Else ReDim Preserve rab(UBound(rab) + 1)
    rab(UBound(rab)).fS = rAbr!fsurogat
    rab(UBound(rab)).BSNR = rAbr!FBetriebsnr
   rAbr.MoveNext
  Loop
  
  If pid = -1 Then pid = fPtNr
  syscmd 4, "Übertrage Leistungen für " & fPtNr
'  Leistungen
  sql = "SELECT 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND Zp, MID(18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND,12,5) uhrz, FICdcode Art," & vbCrLf & _
  "REPLACE(COALESCE(REPLACE(MID(FDetails,INSTR(FDetails,'ext ""')+5,LENGTH(FDetails)-2-INSTR(FDetails,'ext ""')-5),'\n','; '),FText),'''','\''') FText," & vbCrLf & _
  "FEintragsart, 18900101+INTERVAL FAnorddatum DAY+INTERVAL FAnordzeit SECOND AnZp," & vbCrLf & _
  "COALESCE(na.FInitialen,'') ua, COALESCE(nb.FInitialen,'') ub, REPLACE(REPLACE(REPLACE(FDetails,'{(Gnrliste [',''),'])}',''),'''','\''') Lei " & vbCrLf & _
  "FROM ltag l " & vbCrLf & _
  "LEFT JOIN nutzerneu na ON FAnordnutzernr= na.FSurogat " & vbCrLf & _
  "LEFT JOIN nutzerneu nb ON FAusfnutzernr= nb.FSurogat " & vbCrLf & _
  "LEFT JOIN patfall pf ON l.FScheinnr=pf.FSurogat" & vbCrLf & _
  "WHERE l.FPatnr = " & fPtNr & vbCrLf & _
  " AND FEintragsart=12" & vbCrLf & _
  " AND FAbgerechnet<>3652" ' gesperrt
'  " AND 18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND >CONVERT(240601, DATE)"
'  "AND FDetails RLIKE '(""97271E""|""92392F""|""92392B"")'" & vbCrLf & _
'  "AND CONVERT(18900101+INTERVAL FDatum DAY+INTERVAL FZeit SECOND, date) BETWEEN 20250101 AND 20250401" & vbCrLf & _

  myFrag rsEi, sql, adOpenStatic, MOCon
  If Not rsEi.BOF Then
   Do While Not rsEi.EOF
    Dim Lei$, lp&, EKT%, GKT%, RKT%, w As Byte, BDTz$, Inh$, ia As Byte, b As String * 1 ' eckige Klammertiefe, runde Klammertiefe, iA= in Anführungszeichen, Buchstabe
    Dim neuW As Byte, obWert As Byte ' Beginn eines neuen Wortes, aktuelles Wort ist Wert
    GKT = 0
    EKT = 0
    RKT = 0
    Lei = rsEi!Lei
    For lp = 1 To Len(Lei)
     b = Mid$(Lei, lp)
     Select Case b
      Case "{"
        If Not ia Then
         GKT = GKT + 1
         If GKT = 1 Then
          ReDim Preserve rLe(UBound(rLe) + 1)
          rLe(UBound(rLe)).Pat_ID = pid
          rLe(UBound(rLe)).aktZeit = aktZeit
          rLe(UBound(rLe)).Zeitpunkt = rsEi!Zp
          rLe(UBound(rLe)).QS = ZQSort(rLe(UBound(rLe)).Zeitpunkt)
          rLe(UBound(rLe)).QT = ZQuart(rLe(UBound(rLe)).Zeitpunkt)
          rLe(UBound(rLe)).LUhrz = rsEi!Uhrz
          rLe(UBound(rLe)).Ersteller = rsEi!ua
          rLe(UBound(rLe)).Änderer = rsEi!ub
         End If ' GKT = 1 Then
        End If ' Not ia Then
      Case "}": If Not ia Then GKT = GKT - 1
      Case "[": If Not ia Then EKT = EKT + 1
      Case "]": If Not ia Then EKT = EKT - 1
      Case "(": If Not ia Then RKT = RKT + 1: If RKT = 1 Then w = 0
      Case ")"
       If Not ia Then
        RKT = RKT - 1
        If EKT = 0 And RKT = 0 Then
         Select Case BDTz
          Case "5001"
           rLe(UBound(rLe)).Leistung = Inh
          Case "5002"
           rLe(UBound(rLe)).ArtdUs = REPLACE$(Inh, "ArtdUs: ", "")
          Case "5005"
           rLe(UBound(rLe)).LAnzl = Inh
          Case "5006"
           rLe(UBound(rLe)).LUhrz = Format(Inh, "##:##:##")
          Case "5018"
           rLe(UBound(rLe)).Zone = Inh
          Case "Lkz":
           Select Case Inh
            Case "K", "tk": rLe(UBound(rLe)).Lanr = "933284903": rLe(UBound(rLe)).lanrid = 2
            Case "S", "gs": rLe(UBound(rLe)).Lanr = "889690003": rLe(UBound(rLe)).lanrid = 1
            Case "H", "ah": rLe(UBound(rLe)).Lanr = "177828303": rLe(UBound(rLe)).lanrid = 5
            Case Else: rLe(UBound(rLe)).Lanr = "999999900": rLe(UBound(rLe)).lanrid = 4
           End Select
          Case "Lstgerbnr"
           rLe(UBound(rLe)).Lstgerbnr = Inh
          Case "Position"
           rLe(UBound(rLe)).Position = Inh
          Case "Eignung"
           rLe(UBound(rLe)).Eignung = Inh
          Case "Pruefzeit"
           rLe(UBound(rLe)).Pruefzeit = Inh
          Case "Kalkzeit"
           rLe(UBound(rLe)).Kalkzeit = Inh
          Case "organliste"
            rLe(UBound(rLe)).LOrgan = Inh
          Case "begruendungsliste"
           rLe(UBound(rLe)).LfBegr = Inh
          Case "Sachkostentyp"
           rLe(UBound(rLe)).Sachkostentyp = Inh
          Case "sachkostenliste", "sachkostenliste5001"
           rLe(UBound(rLe)).Sachkbez = Inh
          Case "5061"
             rLe(UBound(rLe)).Punkte = REPLACE$(Inh, ".", ",")
          Case "5062"
             rLe(UBound(rLe)).Faktor = CDbl(REPLACE$(Inh, ".", ","))
          Case "Chargennummer"
            rLe(UBound(rLe)).Charge = REPLACE$(Inh, "Chargennummer", "")
          Case "Bsnr"
           rLe(UBound(rLe)).BSNR = Inh
           If SafeArrayGetDim(rab) <> 0 Then
            For abz = 0 To UBound(rab)
             If rab(abz).fS = Inh Then
              rLe(UBound(rLe)).LBSNR = rab(abz).BSNR
              Exit For
             End If
            Next abz
           End If
          Case "Kennzeichen_arzt_technik" ' "A"
           rLe(UBound(rLe)).Kennzeichen_arzt_technik = Inh
          Case "Auftragsschluessel" ' "EXP"
           rLe(UBound(rLe)).Auftragsschluessel = Inh
          Case "Auftragstext" ' experimenteller Auftrag
           rLe(UBound(rLe)).Auftragstext = Inh
          Case "Medikanr"
           rLe(UBound(rLe)).Medikanr = Inh
          Case "Abrechnungssperre"
           rLe(UBound(rLe)).Abrechnungssperre = Inh
          Case "Arztliste"
           rLe(UBound(rLe)).Arztliste = Inh
          Case "Usegebordtext"
           rLe(UBound(rLe)).Usegebordtext = Inh
          Case "Laborkosten"
           rLe(UBound(rLe)).Laborkosten = Inh
          Case Else
           Debug.Print "Leistungen: BDTz: " & BDTz & ": " & Inh
         End Select
         BDTz = ""
         Inh = ""
        End If '  RKT = 1 Then
       End If ' Not ia And Not EKT Then
      Case """": ia = Not ia
      Case " ":
       If Not ia Then
        If EKT = 0 And RKT = 1 Then
         w = 1
        Else
         If Inh <> "" Then Inh = Inh & b
        End If
       Else
        Inh = Inh & b
       End If
      Case Else
       If w Then Inh = Inh & b Else BDTz = BDTz & b
     End Select
    Next lp
    rsEi.MoveNext
   Loop ' while not rsEi.EOF
  End If ' Not rsEi.BOF Then
 Dim jj&
' DBCn.BeginTrans
' For jj = 1 To UBound(rLe)
'  DBCn.Execute "update leistungen set lanrid = " & rLe(jj).lanrid & ",lanr=" & rLe(jj).Lanr & " where pat_id=" & rLe(jj).Pat_ID & " and leistung = '" & rLe(jj).Leistung & "' and position=" & rLe(jj).Position & " and zeitpunkt='" & Format(rLe(jj).Zeitpunkt, "yyyymmddHHMMSS") & "'", rAf
''  Debug.Print jj, rLe(jj).lanrid, rLe(jj).Pat_ID, rLe(jj).Leistung, rLe(jj).Position, rLe(jj).Zeitpunkt, rAf
' Next jj
' DBCn.CommitTrans
 Ausgeb "Leistungen für " & fPtNr & " übertragen", True
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MOLeistungen/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MOLeistungen(FPatnr&)



' in MedOffTabZahl_Click (obsyst=true) und suchfi (obsyst=false)
Public Function moausgeb(MOCon As ADODB.Connection, tn$, obsyst%, Bedg$)
  Dim rcol As New ADODB.Recordset, raen As New ADODB.Recordset
  Dim runde%, tr&, colN$, colZ&, cols$, c2s$, sql$, Prim$, MOSvr$, ausgb$
  Dim MeStr() As memoType
  Dim Wt() ' Werte
  Dim mt$
  Dim Offs% ' Offset der ersten Spalte aus a.*
  On Error GoTo fehler
'  Dim obvers% ' wird beim Aufruf schon beachtet
'  obvers = MOCon.Execute("SELECT table_type='SYSTEM VERSIONED' FROM information_schema.tables WHERE table_schema='medoff' AND TABLE_NAME='" & Tn & "'").Fields(0)
'  If obvers = 0 And obsyst <> 0 Then Exit Function
  Prim = MOCon.Execute("SELECT GROUP_CONCAT(DISTINCT COLUMN_NAME) sp FROM information_schema.key_column_usage i WHERE CONSTRAINT_NAME='PRIMARY' AND table_schema='medoff' AND table_name='" & tn & "' AND column_name NOT IN ('row_start','row_end') GROUP BY table_catalog,table_schema,TABLE_NAME ORDER BY table_catalog,table_schema,table_name,ordinal_position").Fields(0)
  rcol.Open "SELECT GROUP_CONCAT(column_name) cns,COUNT(0) zl,GROUP_CONCAT(CONCAT('CONCAT(''',column_name,':'',CONVERT(COALESCE(LEFT(',column_name,',40),'''') USING latin1))') SEPARATOR ','' '',') cols,GROUP_CONCAT(CONCAT(IF(data_type='longblob',CONCAT('COALESCE(CONVERT(',COLUMN_NAME,' USING latin1),'''') ',column_name),column_name)) ORDER BY ordinal_position SEPARATOR ',') c2s FROM information_schema.columns WHERE table_schema = 'medoff' AND TABLE_NAME='" & tn & "' ORDER BY ordinal_position", MOCon, adOpenStatic, adLockReadOnly
'  GROUP_CONCAT(CONCAT('CONCAT(''',column_name,':'',CONVERT(COALESCE(LEFT(',column_name,',20),'''') USING ''utf8mb4''))') SEPARATOR ','' '',')
'  GROUP_CONCAT(CONCAT(IF(data_type='longblob',CONCAT('COALESCE(CONVERT(',column_name,' USING latin1),'''') ',column_name),column_name)) SEPARATOR ',')
  If Not rcol.EOF Then
   colN = rcol!CNs
   colZ = rcol!zl
   cols = rcol!cols
   c2s = rcol!c2s
   Set rcol = Nothing
   ReDim Wt(colZ)
   rcol.Open "SELECT column_name cn,DATA_TYPE='longblob' obm FROM information_schema.columns WHERE table_schema = 'medoff' AND TABLE_NAME='" & tn & "' ORDER BY ordinal_position", MOCon, adOpenStatic, adLockReadOnly
   If obsyst Then ' obvers
    sql = "SELECT * FROM (SELECT CONCAT(" & cols & ",' ',row_start,' ',row_end) sp, row_start, LEAD(row_start,1) OVER (PARTITION BY " & Prim & " ORDER BY row_start) nrs, LAG(row_start,1) OVER (PARTITION BY " & Prim & " ORDER BY row_start) vrs, " & c2s & _
         " FROM `" & tn & "`" & IIf(obsyst, " FOR system_time ALL a", "") & ") i WHERE " & Bedg & " ORDER BY " & Prim & ",row_start;"
    Offs = 4
   Else ' obvers
    sql = "SELECT * FROM (SELECT CONCAT(" & cols & ") sp, 0 vrs," & c2s & _
         " FROM `" & tn & "`) i WHERE " & Bedg & " ORDER BY " & Prim & ";"
    Offs = 2
   End If ' obvers else
   MOSvr = GetSvr(MOCon)
   raen.Open sql, MOCon, adOpenStatic, adLockReadOnly
   runde = 0 ' bei Versioning in Runde 0 die vorherigen Werte in wt() merken, in Runde 1 beide drucken
   Do While Not raen.EOF
'    If Tn = "beschein" Then Stop
    If runde = 0 Then
     tr = MOCon.Execute("SELECT COUNT(0) FROM `" & tn & "`").Fields(0)
     ausgb = tn & " ("
     If obsyst <> 0 Then ausgb = ausgb & DBCn.Execute("SELECT CONCAT('(-',MAX(datum),') ',table_rows) FROM moprot WHERE server='" & MOSvr & "' AND table_name='" & tn & "' AND datum=(SELECT MAX(datum) FROM moprot WHERE server='" & MOSvr & "' AND table_name='" & tn & "')").Fields(0) & " -> "
     ausgb = ausgb & tr & " Sätze):"
     Print #220, ausgb
     Print #220, colN & ":"
    End If ' runde = 0 Then
    Print #220, raen.Fields(0)
    rcol.MoveFirst
    colZ = 0
    Do While Not rcol.EOF
 '    If Tn = "patstamm" And raen.Fields(colz + Offs).name = "FMemo" Then Stop
'     If Tn = "beschein" And rcol.Fields(0) = "FMemo" Then Stop
     If colZ + Offs >= raen.Fields.COUNT Then Exit Do
     If runde = 0 Then
      If IsNull(raen.Fields(colZ + Offs)) Then Wt(colZ) = "NULL" Else Wt(colZ) = raen.Fields(colZ + Offs)
     End If ' runde = 0 Then
     If (runde <> 0 And Not IsNull(raen!vrs)) Or _
        (runde = 0 And IsNull(raen!vrs)) Then ' wenn Datensatz neu eingefügt wurde, Runde 0, sonst Runde 1
'      If IsNull(raen!vrs) Then Stop
      If IsNull(raen.Fields(colZ + Offs)) Then ausgb = "NULL" Else ausgb = REPLACE$(REPLACE$(raen.Fields(colZ + Offs), Chr(10), ""), Chr(13), "<nl>")
'      If rcol.Fields(0) = "FMemo" Then Stop
      If Not obsyst Or runde = 0 Or CStr(Wt(colZ)) <> CStr(IIf(IsNull(raen.Fields(colZ + Offs)), "NULL", raen.Fields(colZ + Offs))) Then
       Print #220, rcol.Fields(0) & ": " & IIf(runde = 0, "", IIf(obsyst <> 0, Wt(colZ) & vbCrLf & " -> ", "") & Space$(Len(rcol.Fields(0)) - 2)) & ausgb
       If raen.Fields(colZ + Offs).Type = adLongVarBinary Or raen.Fields(colZ + Offs).Type = adLongVarChar Then ' 205
        Dim wtcolz$, i&, pru%
        For pru = 0 To IIf(obsyst <> 0 And runde <> 0, 1, 0)
         If pru = 0 Then
          wtcolz = Wt(colZ)
         Else
          wtcolz = raen.Fields(colZ + Offs)
          Print #220, " -> "
         End If
         Call ParseMemo(wtcolz, MeStr)
         If SafeArrayGetDim(MeStr) <> 0 Then
          Print #220, "Znr.|mx|Ebn|ENr|wtcolz|endse|endsz|Länge|mt|Datum(mt)"
          For i = 0 To UBound(MeStr)
           mt = MeStr(i).Text & String(4, Chr(0))
           Print #220, MeStr(i).znr & "|" & MeStr(i).mx & "|" & MeStr(i).ebn & "|" & MeStr(i).ENr & "|" & IIf(MeStr(i).endse <> "10", Mid$(wtcolz, MeStr(i).znr, 1), "") & "|" & MeStr(i).endse & "|" & MeStr(i).endsz & "| Laenge: " & Len(mt) & "|" & IIf(Right$(mt, 1) = Chr$(10), Left$(mt, IIf(Len(mt) = 0, 1, Len(mt)) - 1), mt), _
           Asc(mt), Asc(Mid(mt, 2)), Asc(Mid(mt, 3)), Asc(Mid(mt, 4)), _
           Asc(Mid(mt, 4, 1)) & "." & Asc(Mid(mt, 3, 1)) & "." & Asc(Mid(mt, 2, 1)) * 256& + Asc(mt)
          Next i
         End If ' SafeArrayGetDim(MeStr) <> 0 Then
         If pru = 1 Then
          Print #220, ""
         End If ' pru = 1
        Next pru
       End If ' raen.Fields(colz + Offs).Type = adLongVarBinary Or raen.Fields(colz + Offs).Type = adLongVarChar Then ' 205
      End If ' CStr(Wt(colz)) <> CStr(IIf(IsNull(raen.Fields(colz + Offs)), "NULL", raen.Fields(colz + Offs))) Then
     End If ' runde = 0 else
     colZ = colZ + 1
     rcol.MoveNext
    Loop ' While Not rcol.EOF
    runde = runde + 1
    If obsyst <> 0 Or (obsyst = 0 And runde Mod 2 = 0) Then
     raen.MoveNext
    End If ' obsyst <> 0 Or (obsyst = 0 And runde Mod 2 = 0) Then
   Loop ' While Not raen.EOF
   If ausgb <> "" Then Print #220, ""
   Set raen = Nothing
  End If ' Not rcol.EOF Then
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 If Err.Number = -2147467259 Then
  DBCn.Close
  DBCn.Open
  Resume
 End If ' Err.Number = -2147467259 Then
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in moausgeb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' moausgeb


' nur zum Aufruf im Direktfenster
' und aus Start(MOSuch)
Public Function suchfi(pNr&, fI$, notObRlike%, MServ$)
 Dim altt$, gefu%, i&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset, RsI As New ADODB.Recordset
 Dim MOCon As New ADODB.Connection
 Dim D1$, fn$, ausgStr$, ausgTNr%
 D1 = "\\linux1\daten\down\suchfi_" & pNr & "_" & fI & "_" & MServ
 Open D1 For Output As #220
 MOCon.Open MOAnfStr & MServ
 
 rst.Open "SELECT c.TABLE_NAME tn, c.COLUMN_NAME cn -- , a.column_name cn" & vbCrLf & _
          "FROM information_schema.columns c" & vbCrLf & _
          "-- LEFT JOIN information_schema.columns a ON c.TABLE_CATALOG=a.TABLE_CATALOG AND c.TABLE_SCHEMA=a.TABLE_SCHEMA AND c.table_name=a.TABLE_NAME" & vbCrLf & _
          "WHERE c.table_catalog='def' AND c.table_schema='medoff' AND (c.COLUMN_NAME IN ('fpatnr','patnr','fpatid') OR (c.TABLE_NAME='patstamm' AND c.COLUMN_NAME='fsurogat')) GROUP BY c.table_name ORDER BY c.table_name", MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
  If rst!tn <> altt Then
   Set RsI = Nothing
   RsI.Open "SELECT k.column_name FROM information_schema.table_constraints t JOIN information_schema.key_column_usage k USING(constraint_name,table_schema,TABLE_NAME) WHERE t.constraint_type='PRIMARY KEY' AND t.table_schema='medoff' AND t.table_name='" & rst!tn & "'", MOCon, adOpenStatic, adLockReadOnly
   If Not RsI.EOF Then fn = RsI.Fields(0) Else fn = ""
   Set rsu = Nothing
   rsu.Open "SELECT " & MOCon.Execute("SELECT GROUP_CONCAT(CONCAT(IF(data_type='longblob',CONCAT('COALESCE(CONVERT(',COLUMN_NAME,' USING latin1),'''') ',COLUMN_NAME),COLUMN_NAME)) ORDER BY ordinal_position SEPARATOR ',') FROM information_schema.`COLUMNS` c WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_NAME='" & rst!tn & "'").Fields(0) & " FROM `" & rst!tn & "` WHERE `" & rst!Cn & "`=" & pNr, MOCon, adOpenStatic, adLockReadOnly
   gefu = 0
'   If LCase$(rst!Tn) = "beschein" Then Stop
   Do While Not rsu.EOF
    For i = 0 To rsu.Fields.COUNT - 1
'     If rsu.Fields(i).name = "FMemo" Then Stop
     If Not (IsNull(rsu.Fields(i))) Then
'      If CStr(rsu.Fields(i)) = fI Then
      If (notObRlike <> 0 And LCase$(CStr(rsu.Fields(i))) = LCase$(fI)) Or (notObRlike = 0 And InStrB(LCase$(rsu.Fields(i)), LCase$(fI))) <> 0 Then
       ausgTNr = ausgTNr + 1
       ausgStr = "Tb.: " & rst!tn & ", " & fn & ": " & rsu.Fields(fn) & ", " & rsu.Fields(i).name & ": " & REPLACE$(REPLACE$(rsu.Fields(i), Chr(10), ""), Chr(13), "<nl>")
       Debug.Print ausgStr
'       Print #220, IIf(ausgTNr <> 1, vbCrLf, "") & ausgStr
        Print #220, ausgStr
'       Call moausgeb(MOCon, rst!Tn, False, "`" & rst!Cn & "`=" & pNr)
       Call moausgeb(MOCon, rst!tn, False, "`" & fn & "`=" & rsu.Fields(fn))
       gefu = True
       GoTo weiter
      End If ' InStrB(LCase$(rsu.Fields(i)), LCase$(fI)) <> 0 Then
     End If ' Not (IsNull(rsu.Fields(i))) Then
    Next i ' i = 0 To rsu.Fields.COUNT - 1
'    If gefu Then GoTo weiter
   rsu.MoveNext
  Loop ' While Not rsu.EOF
weiter:
  End If
  rst.MoveNext
 Loop ' While Not rst.EOF
 Close #220
 zeigan D1
 syscmd 4, "Fertig mit suchfi " & pNr& & " " & fI$ & "," & MServ
 Debug.Print "Fertig mit Suchfi(" & pNr & "," & fI & "," & MServ & ")"
End Function ' suchfi(pNr&, fI$)


' nur zum Aufruf im Direktfenster
Public Function suchal(fI$, Optional notObRlike%, Optional MServ$)
 Dim altt$, j&
 Dim rst As New ADODB.Recordset, rsu As New ADODB.Recordset
' Dim MOCon As New ADODB.Connection
 Dim D1$
 D1 = "\\linux1\daten\down\suchal_" & fI & "_" & notObRlike & "_" & MServ
 Open D1 For Output As #318
 Call MOConInit(MServ, "suchal(" & fI & "," & CStr(notObRlike) & "," & MServ & ")")
' MOCon.Open MOAnfStr & MServ
#If True Then
' rst.Open "SELECT TABLE_NAME tn, GROUP_CONCAT(CONCAT('CAST(',column_name,' AS CHAR)" & IIf(Not notObRlike, " RLIKE ", "=") & "''" & fI & "''') SEPARATOR ' OR ') cn FROM information_schema.columns c WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_TYPE<>'SEQUENCE' GROUP BY table_name" _
 , MOCon, adOpenStatic, adLockReadOnly
 sql = "SELECT c.TABLE_NAME tn, GROUP_CONCAT(CONCAT('CAST(',COLUMN_NAME,' AS CHAR)" & IIf(Not notObRlike, " RLIKE ", "=") & "''" & fI & "''') SEPARATOR ' OR ') cn FROM information_schema.COLUMNS c" & vbCrLf & _
       "LEFT JOIN information_schema.TABLES t ON c.TABLE_CATALOG=t.TABLE_CATALOG AND c.TABLE_SCHEMA=t.TABLE_SCHEMA AND c.TABLE_NAME= t.TABLE_NAME" & vbCrLf & _
       "WHERE t.table_catalog='def' AND t.table_schema='medoff' AND t.TABLE_TYPE<>'SEQUENCE' GROUP BY t.TABLE_NAME"
 rst.Open sql, MOCon, adOpenStatic, adLockReadOnly
 Do While Not rst.EOF
'  Debug.Print rst!Tn, rst!Cn
  DoEvents
  Set rsu = Nothing
  rsu.Open "SELECT * from `" & rst!tn & "` WHERE " & rst!Cn, MOCon, adOpenStatic, adLockReadOnly
  If Not rsu.EOF Then
   Print #318, rst!tn & ": " & rsu.Fields(0).name & ": " & rsu.Fields(0)
   Debug.Print rst!tn & ": " & rsu.Fields(0).name & ": " & rsu.Fields(0)
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
  If rst!tn <> altt Then
   Set rsu = Nothing
   rsu.Open "SELECT * from `" & rst!tn & "`", MOCon, adOpenStatic, adLockReadOnly
   j = 0
   Do While Not rsu.EOF
    j = j + 1
    Print #318, rst!tn & " " & j
    Debug.Print rst!tn, j
    DoEvents
'    If j = 8374 Then Stop
    For i = 0 To rsu.Fields.COUNT - 1
     If Not (IsNull(rsu.Fields(i))) Then
      If CStr(rsu.Fields(i)) = fI Then
       Print #318, rsu.Fields(i).name
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
 Close #318
 zeigan D1
 syscmd 4, "Fertig mit suchal " & fI & " " & notObRlike% & "," & MServ
 Debug.Print "Fertig mit Suchal(" & fI & "," & notObRlike & "," & MServ & ")"
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
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmpatstamm (
    patnr INT(11) UNSIGNED NOT NULL,
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(patnr,znr),
    Key zuloe(PatNr, ENr)
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmpatstamm WHERE (pnr=0 OR patnr=pnr) AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpmpatstamm i WHERE i.patnr=tmpmpatstamm.patnr AND INSTR(i.enr,CONCAT(tmpmpatstamm.enr,'.'))=1 AND i.enr<>tmpmpatstamm.enr);
-- COMMIT;
-- LEAVE tp;
 SELECT PatNr, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(mid(TEXT,2))b2, ASCII(mid(TEXT,3))b3, ascii(mid(TEXT,4))b4,
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
DECLARE txt VARCHAR(10024);
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
    Key zuloe(PatNr, FSur, ENr)
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
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
DECLARE txt VARCHAR(10024);
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
    Key zuloe(PatNr, FSur, ENr)
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
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


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmTextKategorieliste`()
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FTextKategorieliste aus mosystem in die Tabelle tmpmTextKat aus'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmTextKat (
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(znr),
    Key zuloe(ENr)
);
DELETE FROM tmpmTextKat;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT FTextKategorieliste fmemo FROM mosystem WHERE FTextKategorieliste IS NOT NULL) DO
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
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmTextKat WHERE mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmTextKat WHERE mx>=aktMAX);
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
     INSERT INTO tmpmTextKat(znr,mx,ebn,enr,TEXT) VALUES(pos,max,ie,
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmTextKat WHERE enr<>'0' AND EXISTS (SELECT 0 FROM tmpmTextKat i WHERE INSTR(i.enr,CONCAT(tmpmTextKat.enr,'.'))=1 AND i.enr<>tmpmTextKat.enr);
-- COMMIT;
-- LEAVE tp;
 SELECT znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
 CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmTextKat ORDER BY znr;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmKategorieliste`()
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FKategorieliste aus mosystem in die Tabelle tmpmTextKat aus'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmKat (
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(znr),
    Key zuloe(ENr)
);
DELETE FROM tmpmTextKat;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT FKategorieliste fmemo FROM mosystem WHERE FKategorieliste IS NOT NULL) DO
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
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmTextKat WHERE mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmTextKat WHERE mx>=aktMAX);
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
     INSERT INTO tmpmTextKat(znr,mx,ebn,enr,TEXT) VALUES(pos,max,ie,
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmTextKat WHERE enr<>'0' AND EXISTS (SELECT 0 FROM tmpmTextKat i WHERE INSTR(i.enr,CONCAT(tmpmTextKat.enr,'.'))=1 AND i.enr<>tmpmTextKat.enr);
-- COMMIT;
-- LEAVE tp;
 SELECT znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
 CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmTextKat ORDER BY znr;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmAuftragstypenliste`()
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FAuftragstypenliste aus mosystem in die Tabelle tmpmAuftrTyp aus'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmAuftrTyp (
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(znr),
    Key zuloe(ENr)
);
DELETE FROM tmpmAuftrTyp;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT FAuftragstypenliste fmemo FROM mosystem WHERE FAuftragstypenliste IS NOT NULL) DO
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
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmAuftrTyp WHERE mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmAuftrTyp WHERE mx>=aktMAX);
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
     INSERT INTO tmpmAuftrTyp(znr,mx,ebn,enr,TEXT) VALUES(pos,max,ie,
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmAuftrTyp WHERE enr<>'0' AND EXISTS (SELECT 0 FROM tmpmAuftrTyp i WHERE INSTR(i.enr,CONCAT(tmpmAuftrTyp.enr,'.'))=1 AND i.enr<>tmpmAuftrTyp.enr);
-- COMMIT;
-- LEAVE tp;
 SELECT znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
 CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmAuftrTyp ORDER BY znr;
End


CREATE DEFINER=`medoff`@`%` PROCEDURE `procmAblageliste`()
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'liest das Feld FAblageliste aus mosystem in die Tabelle tmpmAbl aus'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmAbl (
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(znr),
    Key zuloe(ENr)
);
DELETE FROM tmpmAbl;
CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle
laba: FOR r IN (SELECT FAblageliste fmemo FROM mosystem WHERE FAblageliste IS NOT NULL) DO
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
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmAbl WHERE mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmAbl WHERE mx>=aktMAX);
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
     INSERT INTO tmpmAbl(znr,mx,ebn,enr,TEXT) VALUES(pos,max,ie,
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
 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmAbl WHERE enr<>'0' AND EXISTS (SELECT 0 FROM tmpmAbl i WHERE INSTR(i.enr,CONCAT(tmpmAbl.enr,'.'))=1 AND i.enr<>tmpmAbl.enr);
-- COMMIT;
-- LEAVE tp;
 SELECT znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
 CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum
   FROM tmpmAbl ORDER BY znr;
End


Begin
DECLARE sq LONGTEXT DEFAULT CONCAT("SELECT COLUMN_NAME,TABLE_NAME INTO @sp,@tb FROM information_schema.columns "
     "WHERE table_catalog='def' AND table_schema='medoff' AND TABLE_NAME ='", tab, "' AND COLUMN_NAME='",fld,"'");
 SET @sp=NULL;
 PREPARE smt FROM sq; EXECUTE smt; DEALLOCATE PREPARE smt;
 IF @sp IS NOT NULL THEN
   SET sq=CONCAT("CREATE OR REPLACE DEFINER=`medoff`@`%` PROCEDURE `procm",@tb,"`(IN `pnr` INT) LANGUAGE SQL NOT DETERMINISTIC CONTAINS SQL SQL SECURITY DEFINER ", CHR(10),
   "COMMENT 'liest das Feld ",@sp," aus ",@tb," in die Tabelle tmpm",@tb," aus'", CHR(10),
   "tp: Begin", CHR(10),
   "DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie INT(10) DEFAULT 0;", CHR(10),
   "DECLARE obdruck INT(1);", CHR(10),
   "DECLARE txt VARCHAR(10024);", CHR(10),
   "-- START TRANSACTION;", CHR(10),
   "CREATE TABLE IF NOT EXISTS tmpm",@tb," (", CHR(10),
   "    patnr INT(11) UNSIGNED NOT null,", CHR(10),
   "    fsur INT(11) UNSIGNED NOT NULL,", CHR(10),
   "    znr INT(2) UNSIGNED DEFAULT '0',", CHR(10),
   "    MX INT(3) UNSIGNED DEFAULT '0',", CHR(10),
   "    ebn INT(2) DEFAULT '0',", CHR(10),
   "    enr VARCHAR(128) DEFAULT '',", CHR(10),
   "    TEXT text DEFAULT '',", CHR(10),
   "    PRIMARY KEY Zugriff(patnr,fsur,znr),", CHR(10),
   "    Key zuloe(PatNr, fsur, enr)", CHR(10),
   ");", CHR(10),
   "DELETE FROM tmpm",@tb," WHERE pnr=0 OR patnr=pnr;", CHR(10),
   "CREATE TEMPORARY TABLE IF NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle", CHR(10),
   "laba: FOR r IN (SELECT fsurogat fsur, fpatnr, ",@sp," FROM ",@tb," WHERE (pnr=0 OR fpatnr=pnr) AND ",@sp," IS NOT NULL) DO", CHR(10),
   " SET pos=1;", CHR(10),
   " SET ie=0;", CHR(10),
   " TRUNCATE tmpeb;", CHR(10),
   "labt:  LOOP", CHR(10),
   "  SET obdruck=0;", CHR(10),
   "  SET tlen=CONV(HEX(CONCAT(MID(r.",@sp,",pos+1,1),MID(r.",@sp,",pos,1))),16,10);", CHR(10),
   "  IF pos=1 THEN SET gl=tlen; SET MAX=gl; END IF;", CHR(10),
   "  SET altmax=MAX;", CHR(10),
   "  SET aktMAX=tlen+pos+1; -- aktuell angegebene Länge aus dem und dem nä Byte", CHR(10),
   "  IF aktmax BETWEEN 0 AND gl+2 AND -- wenn die angegebene Länge vertretbar", CHR(10),
   "    (( CONV(HEX(MID(r.",@sp,",aktMax,1)),16,10)=0 AND -- und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen", CHR(10),
   "     aktmax<=MAX) OR pos>MAX) THEN -- wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde", CHR(10),
   "    SET altie=ie; -- letzte Ebene", CHR(10),
   "    IF aktmax<=MAX THEN SET ie=ie+1; END IF; -- im ersten Fall wird die Ebene erhöht", CHR(10),
   "    IF pos>MAX THEN -- im zweiten Fall wird zurückgegriffen", CHR(10),
   "     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpm",@tb," WHERE patnr=r.fpatnr and fsur=r.fsur AND mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpm",@tb," WHERE patnr=r.fpatnr and fsur=r.fsur AND mx>=aktMAX);", CHR(10),
   "     END IF;", CHR(10),
   "    IF ie<=altie THEN -- wenn die Ebene nicht erhöht worden ist", CHR(10),
   "     IF ie<altie THEN -- wenn sie vielmehr reduziert wurde", CHR(10),
   "      DELETE FROM tmpeb WHERE nr>ie; -- dann werden die höheren Einträge wieder gelöscht", CHR(10),
   "     END IF;", CHR(10),
   "     UPDATE tmpeb SET wert=wert+1 WHERE nr=ie; -- dann wird die Zählung der akt. Ebene erhöht", CHR(10),
   "    ELSE", CHR(10),
   "     INSERT INTO tmpeb(nr,wert) VALUES(ie,1); -- sonst muss ein neuer Eintrag für die hohe Ebene erstellt werden", CHR(10),
   "    END IF;", CHR(10),
   "    SET MAX=aktmax; -- neue vorbestehende Länge", CHR(10),
   "    SET obdruck=1; -- und drucken", CHR(10),
   "  END IF;", CHR(10),
   "  IF obdruck=1", CHR(10),
   "--   OR pos=1 -- zum Debuggen der ersten Gesamt-Zeile", CHR(10),
   "--   OR TRUE -- zum Debuggen aller Zeilen", CHR(10),
   "   THEN", CHR(10),
   "    SET txt=MID(r.",@sp,",pos+2,tlen);", CHR(10),
   "    IF txt <> Repeat(Chr(0), tlen) THEN", CHR(10),
   "     INSERT INTO tmpm",@tb,"(patnr,fsur,znr,mx,ebn,enr,TEXT) VALUES(r.fpatnr,r.fsur,pos,max,ie,", CHR(10),
   "       (SELECT GROUP_CONCAT(wert ORDER BY nr SEPARATOR '.') FROM tmpeb),", CHR(10),
   "--    CONCAT(MID(r.",@sp,",pos,1),'|',HEX(MID(r.",@sp,",pos,1)),'|',CONV(HEX(MID(r.",@sp,",pos,1)),16,10),'|',CONV(HEX(CONCAT(MID(r.",@sp,",pos+1,1),MID(r.",@sp,",pos,1))),16,10),'|','aktMax:',aktmax, '|','Max:',altMAX,'|','Laenge: ',LENGTH(txt),'|','Endbyte:',COALESCE(CONV(HEX(MID(r.",@sp,",aktMax,1)),16,10),'-'),", CHR(10),
   "--     CONCAT(IF(obdruck<>1,'- ',''),'""',", CHR(10),
   "      txt", CHR(10),
   "--    ,'""'))", CHR(10),
   "       );", CHR(10),
   "     ELSE SET pos=pos+tlen; -- 0-Strings nicht auffieseln", CHR(10),
   "     END IF;", CHR(10),
   "  END IF; -- obdruck=1", CHR(10),
   "  SET pos=pos+1;", CHR(10),
   "  IF pos>=gl THEN LEAVE labt; END IF;", CHR(10),
   " END LOOP labt;", CHR(10),
   "END FOR laba;", CHR(10),
   " -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)", CHR(10),
   "DELETE FROM tmpm",@tb," WHERE (pnr=0 OR patnr=pnr) AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpm",@tb," i WHERE i.patnr=tmpm",@tb,".patnr AND i.fsur=tmpm",@tb,".fsur AND INSTR(i.enr,CONCAT(tmpm",@tb,".enr,'.'))=1 AND i.enr<>tmpm",@tb,".enr);", CHR(10),
   "-- folgende Variante bringt, aus der Prozedur aufgerufen, mariadb zum Absturz:", CHR(10),
   "-- DELETE a FROM tmpm",@tb," a LEFT JOIN (SELECT LEAD(enr) OVER(PARTITION BY patnr,fsur ORDER BY znr) nenr, patnr,fsur,znr FROM tmpm",@tb,") i USING (patnr,fsur,znr) WHERE INSTR(i.nenr,a.enr)=1 AND i.nenr<>a.enr;", CHR(10),
   "-- COMMIT;", CHR(10),
   "-- LEAVE tp;", CHR(10),
   " SELECT PatNr, FSur, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,", CHR(10),
   "  CONCAT(ASCII(MID(TEXT,4,1)),'.',ASCII(MID(TEXT,3,1)),'.',ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum", CHR(10),
   "   FROM tmpm",@tb," WHERE pnr=0 OR patnr=pnr ORDER BY fsur DESC,znr;", CHR(10),
   "END"
    );
  PREPARE smt FROM sq; EXECUTE smt; DEALLOCATE PREPARE smt;
  SET sq=CONCAT("CALL procm",@tb,"(",pnr,")");
  PREPARE smt FROM sq; EXECUTE smt; DEALLOCATE PREPARE smt;
 END IF;
End
CREATE DEFINER=`medoff`@`%` PROCEDURE `procmwechs`(
    IN `tabname` VARCHAR(50),
    IN `snr` INT,
    IN `mfeld` VARCHAR(50),
    IN `anENr` VARCHAR(50),
    IN `neu` VARCHAR(50),
    IN `art` INT
)
Language sql
NOT DETERMINISTIC
CONTAINS sql
SQL SECURITY DEFINER
Comment 'tauscht einen string in einem Memofeld aus'
tp: Begin
DECLARE gl,pos,MAX,aktmax,altmax,tlen,ie,altie,reppos,nlen INT(10) DEFAULT 0;
DECLARE aktenr VARCHAR(50);
DECLARE neus VARCHAR(100);
DECLARE sqls VARCHAR(1000);
-- DECLARE mem,memf,fmemf VARCHAR(1024);
DECLARE obdruck INT(1);
DECLARE txt VARCHAR(10024);
-- START TRANSACTION;
CREATE TABLE IF NOT EXISTS tmpmwechs (
    fsur INT(11) UNSIGNED NOT NULL,
    znr INT(2) UNSIGNED DEFAULT '0',
    MX INT(3) UNSIGNED DEFAULT '0',
    ebn INT(2) DEFAULT '0',
    enr VARCHAR(128) DEFAULT '',
    TEXT text DEFAULT '',
    PRIMARY KEY Zugriff(fsur,znr),
    Key zuloe(FSur, ENr)
);
DELETE FROM tmpmwechs WHERE fsur=snr;
CREATE TEMPORARY TABLE if NOT EXISTS tmpeb(nr INT(2) PRIMARY KEY,wert INT(2)); -- temporäre Ebenentabelle

SET sqls=CONCAT("SELECT ",mfeld," INTO @mem FROM `", tabname,"` WHERE fsurogat=",snr);
PREPARE stm FROM sqls;
EXECUTE stm;
DEALLOCATE PREPARE stm;
 SET pos=1;
 SET ie=0;
 TRUNCATE tmpeb;
labt:  Loop
  SET obdruck=0;
  SET tlen=CONV(HEX(CONCAT(MID(@mem,pos+1,1),MID(@mem,pos,1))),16,10);
  if pos=1 then SET gl=tlen; SET MAX=gl; END if;
  SET altmax=MAX;
  SET aktMAX=tlen+pos+1; -- aktuell angegebene Länge aus dem und dem nä Byte
  IF aktmax BETWEEN 0 AND gl+2 AND -- wenn die angegebene Länge vertretbar
    (( CONV(HEX(MID(@mem,aktMax,1)),16,10)=0 AND -- und an der letzten Stelle 0 steht (dann könnte es eine Länge sein), da es hier aber Ausnahmen gibt, wurde re>Max davon ausgenommen
     aktmax<=MAX) OR pos>MAX) then -- wenn die akt.Länge nicht über die Vorbestehende hinausreicht oder d.vorbest.schon überschritten wurde
    SET altie=ie; -- letzte Ebene
    if aktmax<=MAX then SET ie=ie+1; END if; -- im ersten Fall wird die Ebene erhöht
    if pos>MAX then -- im zweiten Fall wird zurückgegriffen
     SELECT COALESCE(MAX(ebn),0)+1 INTO ie FROM tmpmwechs WHERE fsur=snr AND mx>=aktMAX AND znr=(SELECT MAX(znr) FROM tmpmwechs WHERE fsur=snr AND mx>=aktMAX);
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
   THEN
    SET txt=MID(@mem,pos+2,tlen);
    If txt <> Repeat(Chr(0), tlen) Then
     SELECT GROUP_CONCAT(wert ORDER BY nr SEPARATOR '.') INTO aktenr FROM tmpeb;
     If aktenr = anENr Then
      SELECT aktenr, pos;
      SET reppos=pos;
      CASE art when 0 then -- Zahl
--       SET neu=CHR(neu);
       SET nlen=1;
       SET neus=CONCAT("LEFT(CHR(",neu,"),",nlen,")");
      when 1 then   -- string
       SET nlen=LENGTH(neu);
       SET neus=CONCAT("LEFT('",neu,"',",nlen,")");
      when 2 then  -- datum
       SELECT (YEAR(neu)MOD 256) UNION SELECT (YEAR(neu)DIV 256) UNION SELECT (MONTH(neu)) UNION SELECT (DAY(neu));
--       SET neu=CONCAT(CHR(YEAR(neu)MOD 256),CHR(YEAR(neu)DIV 256),CHR(MONTH(neu)),CHR(DAY(neu)));
--       SET neu=CHAR(YEAR(neu)MOD 256,YEAR(neu)DIV 256,MONTH(neu),DAY(neu));
       SET nlen=5;
       SET neus=CONCAT("CHAR(YEAR(",neu,")MOD 256,YEAR(",neu,")DIV 256,MONTH(",neu,"),DAY(",neu,"),0)");
      END CASE;
--      SET sqls=CONCAT("SELECT length('",txt,"'), '",txt,"', ascii(mid('",txt,"',1,1)), ascii(mid('",txt,"',2,1)), ascii(mid('",txt,"',3,1)), ascii(mid('",txt,"',4,1)), ascii(mid('",txt,"',5,1)), ascii(mid('",txt,"',6,1)), length(`",mfeld,"`), 0,0,0, `",mfeld,"` FROM `",tabname,"` WHERE fsurogat=",snr," UNION SELECT length('",neu,"'), '",neu,"', ascii(mid('",neu,"',1,1)), ascii(mid('",neu,"',2,1)), ascii(mid('",neu,"',3,1)), ascii(mid('",neu,"',4,1)), ascii(mid('",neu,"',5,1)), ascii(mid('",txt,"',6,1)), length(CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),'",LEFT(neu,nlen),"',MID(`",mfeld,"`,",reppos,"+2+",nlen,")) ), length(LEFT(`",mfeld,"`,",reppos,"+1)), length('",LEFT(neu,nlen),"'), length(MID(`",mfeld,"`,",reppos,"+2+",nlen,")), CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),'",LEFT(neu,nlen),"',MID(`",mfeld,"`,",reppos,"+2+",nlen,")) FROM `",tabname,"` WHERE fsurogat=",snr);
      SET sqls=CONCAT("UPDATE `",tabname,"` SET `",mfeld,"`=CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),",neus,",MID(`",mfeld,"`,",reppos,"+2+",nlen,")) WHERE fsurogat=",snr);
--      SET sqls=CONCAT("UPDATE `",tabname,"` SET `",mfeld,"`=CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),'",CHR(233),CHR(7),CHR(3),CHR(1),"',MID(`",mfeld,"`,",reppos,"+2+",nlen,")) WHERE fsurogat=",snr);
      SELECT sqls;
--      LEAVE labt;
--      SET sqls=CONCAT("SELECT length(`",mfeld,"`), `",mfeld,"` FROM `",tabname,"` WHERE fsurogat=",snr," UNION SELECT length(CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),'",LEFT(neu,nlen),"',MID(`",mfeld,"`,",reppos,"+2+",nlen,"))), CONCAT(LEFT(`",mfeld,"`,",reppos,"+1),'",LEFT(neu,nlen),"',MID(`",mfeld,"`,",reppos,"+2+",nlen,")) FROM `",tabname,"` WHERE fsurogat=",snr);
      PREPARE stm FROM sqls;
      EXECUTE stm;
      DEALLOCATE PREPARE stm;
        LEAVE labt;
      END IF;
     INSERT INTO tmpmwechs(fsur,znr,mx,ebn,enr,TEXT) VALUES(snr,pos,max,ie,aktenr,
--    CONCAT(MID(@mem,pos,1),'|',HEX(MID(@mem,pos,1)),'|',CONV(HEX(MID(@mem,pos,1)),16,10),'|',CONV(HEX(CONCAT(MID(@mem,pos+1,1),MID(@mem,pos,1))),16,10),'|','aktMax:',aktmax, '|','Max:',altMAX,'|','Laenge: ',LENGTH(txt),'|','Endbyte:',COALESCE(CONV(HEX(MID(@mem,aktMax,1)),16,10),'-'),
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

 -- folgende Zeile zum Debuggen auskommentieren (in der Schleife laba ist das deutlich langsamer!)
DELETE FROM tmpmwechs WHERE fsur=snr AND enr<>'0' AND EXISTS (SELECT 0 FROM tmpmwechs i WHERE i.fsur=tmpmwechs.fsur AND INSTR(i.enr,CONCAT(tmpmwechs.enr,'.'))=1 AND i.enr<>tmpmwechs.enr);
-- folgende Variante bringt, aus der Prozedur aufgerufen, mariadb zum Absturz:
-- DELETE a FROM tmpmwechs a LEFT JOIN (SELECT LEAD(enr) OVER(PARTITION BY fsur ORDER BY znr) nenr, fsur,znr FROM tmpmwechs) i USING (fsur,znr) WHERE INSTR(i.nenr,a.enr)=1 AND i.nenr<>a.enr;
-- COMMIT;
-- LEAVE tp;
 SELECT FSur, znr, Mx, Ebn, ENr, TEXT, ASCII(TEXT)b1, ASCII(MID(TEXT,2))b2, ASCII(MID(TEXT,3))b3, ASCII(MID(TEXT,4))b4,
  CONCAT(ASCII(MID(TEXT,4,1)),".",ASCII(MID(TEXT,3,1)),".",ASCII(MID(TEXT,2,1))*256+ASCII(TEXT)) Datum,
  ROUND(1 + (ASCII(MID(TEXT, 7, 1)) Mod 16)/16 + ASCII(MID(TEXT, 6, 1))/4096 + ASCII(MID(TEXT, 5, 1))/POWER(2,20) + ASCII(MID(TEXT, 4, 1))/POWER(2,28)+ASCII(MID(TEXT, 3, 1))/POWER(2,36)+ASCII(MID(TEXT, 2, 1))/POWER(2,44) + ASCII(MID(TEXT, 1, 1))/POWER(2,52) * POWER(2,(FLOOR(ASCII(MID(TEXT, 7, 1)) / 16) + 1 + 16 * (ASCII(MID(TEXT, 8, 1)) - 64))),8) Bruch
   FROM tmpmwechs WHERE fsur = snr ORDER BY fsur DESC,znr;
End

#End If

