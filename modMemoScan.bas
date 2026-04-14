Attribute VB_Name = "modMemoScan"
Option Explicit

' ============================================================
' modMemoScan.bas  –  v1
' Universeller BLOB-Scanner für epraxis/beschein-Memo-Felder
' Äquivalent zu pgs_fn_memo_scan (MariaDB)
'
' Verwendung:
'   Dim ab()  As Byte
'   Dim nCnt  As Long
'   Dim aFld() As TMemoFeld
'   ab   = BlobFromField(rs.Fields("FMemo"))
'   aFld = MemoScan(ab, nCnt)
'   Dim i As Long
'   For i = 0 To nCnt - 1
'       Debug.Print aFld(i).sTyp, aFld(i).sENr, aFld(i).sWert
'   Next i
' ============================================================

' -- Win32: Speicherkopie für IEEE-754-Konversion ------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)


' ============================================================
' Ergebnis-Struktur
' ============================================================
Public Type TMemoFeld
    sTyp  As String   ' "DAT" | "FLT" | "TXT" | "NUM" | "TEL"
    lPos  As Long     ' Byte-Position im BLOB (1-basiert wie SQL)
    lMax  As Long     ' MAX-Tiefenmarke zum Erkennungszeitpunkt
    sENr  As String   ' Hierarchische Adresse z.B. "1.2.3"
    sWert As String   ' Wert als lesbarer String
End Type

' -- Interner Hilfstyp für Hierarchie-Stack ------------------
Private Type TStackElem
    lAktMax As Long   ' Grenzposition dieses Containers
    lIE     As Long   ' Tiefenstufe (1-basiert)
End Type


' ============================================================
' BlobFromField
' ADO-Feldwert (Variant/ByteArray oder String) ? Byte-Array
' ============================================================
Public Function BlobFromField(vField As Variant) As Byte()
    Dim ab() As Byte
    If IsNull(vField) Then ReDim ab(0): BlobFromField = ab: Exit Function
    Select Case VarType(vField)
        Case vbArray + vbByte:  ab = vField
        Case vbString:          ab = StrConv(CStr(vField), vbFromUnicode)
        Case Else:              ReDim ab(0)
    End Select
    BlobFromField = ab
End Function


' ============================================================
' MemoScan  –  Haupt-Scanner
' Gibt Array(0..nCount-1) von TMemoFeld zurück.
' nCount: Anzahl gefundener Felder (Out-Parameter)
' ============================================================
Public Function MemoScan(abBlob() As Byte, _
                         nCount As Long) As TMemoFeld()

    '-- Ergebnis-Array --------------------------------------
    Dim aRes()  As TMemoFeld
    Dim nAlloc  As Long
    ReDim aRes(15)
    nAlloc = 16
    nCount = 0

    '-- Scan-Variablen --------------------------------------
    Dim lBLen   As Long     ' Länge des BLOB
    Dim lPos    As Long     ' aktueller Scan-Zeiger (0-basiert)
    Dim lGL     As Long     ' Gesamtlänge (aus erstem LE-Word)
    Dim lMax    As Long     ' aktuelle Tiefenmarke
    Dim lAktMax As Long
    Dim lTLen   As Long
    Dim lEndByte As Long
    Dim lSkipTo As Long     ' Sprungmarke nach TEL-Container
    Dim lB0     As Long     ' erstes Payload-Byte
    Dim lB1     As Long     ' zweites Payload-Byte
    Dim lLL     As Long     ' Länge in TEL-Payloads
    Dim lPS     As Long     ' Payload-Startindex (0-basiert)

    '-- Hierarchie-Stack (wie tmpeb in der SP) ---------------
    '   aStack(1..32): je ein Element pro Tiefenstufe
    '   .lAktMax = Grenzposition, .lIE = Stufenzähler
    Dim aStack(1 To 32) As TStackElem
    Dim nSTop   As Long     ' oberster belegter Stack-Index
    Dim lIE     As Long     ' aktuelle Tiefenstufe
    Dim lAltIE  As Long
    Dim i       As Long

    lBLen = UBound(abBlob) + 1
    If lBLen < 4 Then
        ReDim aRes(0): MemoScan = aRes: Exit Function
    End If

    lPos = 0
    lSkipTo = 0
    lIE = 0
    nSTop = 0

    '--------------------------------------------------------
    Do While lPos < lBLen - 1

        '-- tlen: LE-Word aus Byte[lPos] und Byte[lPos+1] --
        lTLen = CLng(abBlob(lPos)) + CLng(abBlob(lPos + 1)) * 256

        '-- Erster Durchlauf: Gesamtlänge lesen ------------
        If lPos = 0 Then
            lGL = lTLen
            lMax = lGL
            lPos = 1
            GoTo NextPos
        End If

        '-- TEL-Container überspringen ----------------------
        If lPos < lSkipTo Then GoTo NextPos

        lAktMax = lTLen + lPos + 1

        '-- Endbyte prüfen (1-basiert: Index lAktMax) -------
        If lAktMax >= 1 And lAktMax < lBLen Then
            lEndByte = CLng(abBlob(lAktMax))
        Else
            lEndByte = -1
        End If

        '----------------------------------------------------
        '  Hauptbedingung (identisch zur SQL-Funktion)
        '----------------------------------------------------
        If lTLen > 0 _
        And lAktMax >= 1 And lAktMax <= lGL + 2 _
        And ((lEndByte = 0 And lAktMax <= lMax) Or lPos > lMax) Then

            '-- Hierarchie-Tracking (wie tmpeb-Logik in SP) -
            lAltIE = lIE

            If lAktMax <= lMax Then
                '  Noch innerhalb des aktuellen Containers: tiefer
                lIE = lIE + 1
            Else
                '  pos > MAX: aus Container herausgetreten
                '  Suche im Stack die tiefste Stufe mit lAktMax
                '  >= aktuellem lAktMax (wie SP-Subquery auf mx)
                lIE = 0
                For i = nSTop To 1 Step -1
                    If aStack(i).lAktMax >= lAktMax Then
                        lIE = aStack(i).lIE + 1
                        Exit For
                    End If
                Next i
                If lIE = 0 Then lIE = 1
            End If

            '-- Stack (= tmpeb) pflegen ----------------------
            If lIE <= lAltIE Then
                '  Gleiche oder höhere Ebene: tiefere Einträge löschen
                If lIE < lAltIE Then
                    For i = lIE + 1 To nSTop
                        aStack(i).lAktMax = 0
                        aStack(i).lIE = 0
                    Next i
                    nSTop = lIE
                End If
                '  Zähler auf dieser Ebene inkrementieren
                If nSTop >= lIE Then
                    aStack(lIE).lIE = aStack(lIE).lIE + 1
                End If
            Else
                '  Neue (tiefere) Ebene einfügen
                If lIE > nSTop Then nSTop = lIE
                If lIE <= 32 Then
                    aStack(lIE).lAktMax = lAktMax
                    aStack(lIE).lIE = 1
                End If
            End If

            '-- ENr als "1.2.3"-String aufbauen -------------
            Dim sENr As String: sENr = ""
            For i = 1 To nSTop
                If aStack(i).lIE > 0 Then
                    If sENr <> "" Then sENr = sENr & "."
                    sENr = sENr & CStr(aStack(i).lIE)
                End If
            Next i

            '-- MAX nach vorne schieben (wie in SQL) ---------
            lMax = lAktMax

            '-- Payload auslesen -----------------------------
            lPS = lPos + 2    ' 0-basierter Payload-Start

            '-- Null-Payload? ? überspringen -----------------
            Dim bAllNull As Boolean: bAllNull = True
            For i = 0 To lTLen - 1
                If abBlob(lPS + i) <> 0 Then bAllNull = False: Exit For
            Next i
            If bAllNull Then
                lPos = lPos + lTLen
                GoTo NextPos
            End If

            lB0 = CLng(abBlob(lPS))
            lB1 = CLng(abBlob(lPS + 1))

            '------------------------------------------------
            Select Case True

            '-- TEL Format A: 02 00 XX 00 LL 00 <String> --
            Case (lB0 = &H2 And lB1 = &H0 And lTLen >= 6)
                lLL = CLng(abBlob(lPS + 4)) + CLng(abBlob(lPS + 5)) * 256
                If lLL > 0 And lLL + 6 <= lTLen Then
                    Dim sTelA As String
                    sTelA = TrimNulls(BytesToStr(abBlob, lPS + 6, lLL))
                    If sTelA <> "" Then
                        AppendResult aRes, nCount, nAlloc, _
                            "TEL", lPos + 1, lMax, sENr, sTelA
                    End If
                End If
                lSkipTo = lPos + lTLen + 1
                lPos = lSkipTo
                GoTo NextPos

            '-- TEL Format B: 02 01 LL <String> -----------
            Case (lB0 = &H2 And lB1 = &H1 And lTLen >= 3)
                lLL = CLng(abBlob(lPS + 2))
                If lLL > 0 And lLL + 3 <= lTLen Then
                    Dim sTelB As String
                    sTelB = TrimNulls(BytesToStr(abBlob, lPS + 3, lLL))
                    If sTelB <> "" Then
                        AppendResult aRes, nCount, nAlloc, _
                            "TEL", lPos + 1, lMax, sENr, sTelB
                    End If
                End If
                lSkipTo = lPos + lTLen + 1
                lPos = lSkipTo
                GoTo NextPos

            '-- DAT: 4 Bytes LE, Byte2 = 0x07 -------------
            Case (lTLen = 4 And lB1 = &H7)
                Dim lMon As Long, lTag As Long, lJahr As Long
                lMon = CLng(abBlob(lPS + 2))
                lTag = CLng(abBlob(lPS + 3))
                lJahr = lB0 + &H700
                If lMon >= 1 And lMon <= 12 And lTag >= 1 And lTag <= 31 Then
                    AppendResult aRes, nCount, nAlloc, _
                        "DAT", lPos + 1, lMax, sENr, _
                        Format(lTag, "00") & "." & _
                        Format(lMon, "00") & "." & CStr(lJahr)
                End If

            '-- NUM: 2 Bytes LE ----------------------------
            Case (lTLen = 2)
                Dim lNum As Long
                lNum = lB0 + lB1 * 256
                If lNum >= 1 And lNum <= 65535 Then
                    AppendResult aRes, nCount, nAlloc, _
                        "NUM", lPos + 1, lMax, sENr, CStr(lNum)
                End If

            '-- FLT: 8 Bytes IEEE-754 LE, MSB 0x3F-0x45 --
            Case (lTLen = 8 And _
                  CLng(abBlob(lPS + 7)) >= &H3F And _
                  CLng(abBlob(lPS + 7)) <= &H45)
                Dim ab8(7) As Byte, dVal As Double
                For i = 0 To 7: ab8(i) = abBlob(lPS + i): Next i
                CopyMemory dVal, ab8(0), 8
                ' Trailing Zeros abschneiden wie in SQL
                Dim sFlt As String
                sFlt = CStr(CDec(Format(dVal, "0.######")))
                AppendResult aRes, nCount, nAlloc, _
                    "FLT", lPos + 1, lMax, sENr, sFlt

            '-- TXT: druckbarer Text -----------------------
            Case (lTLen >= 3 And lTLen <= 500 And lB0 >= &H20)
                Dim sTxt As String
                sTxt = RTrim(TrimNulls(BytesToStr(abBlob, lPS, lTLen)))
                If sTxt <> "" Then
                    AppendResult aRes, nCount, nAlloc, _
                        "TXT", lPos + 1, lMax, sENr, sTxt
                End If

            End Select
            '------------------------------------------------

        End If   ' Hauptbedingung

NextPos:
        lPos = lPos + 1
        If lPos >= lGL Then Exit Do
    Loop
    '--------------------------------------------------------

    If nCount > 0 Then
        ReDim Preserve aRes(nCount - 1)
    Else
        ReDim aRes(0)
    End If
    MemoScan = aRes

End Function


' ============================================================
' MemoFieldVal  –  n-tes Feld eines Typs aus Scan-Ergebnis
' Äquivalent zu pgs_fn_memo_field (SQL)
' n: 1-basiert; gibt "" zurück wenn nicht gefunden
' ============================================================
Public Function MemoFieldVal(aFld() As TMemoFeld, _
                             nCount As Long, _
                             sTyp As String, _
                             n As Long) As String
    Dim i       As Long
    Dim nFound  As Long
    nFound = 0
    For i = 0 To nCount - 1
        If aFld(i).sTyp = sTyp Then
            nFound = nFound + 1
            If nFound = n Then
                MemoFieldVal = aFld(i).sWert
                Exit Function
            End If
        End If
    Next i
    MemoFieldVal = ""
End Function


' ============================================================
' MemoPipe  –  wie pgs_memo_pipe: Pipe-getrennter String
' Slots: DAT1|DAT2|DAT3|FLT1|FLT2|TXT1|TXT2|TXT3|
'        NUM1|NUM2|NUM3|TEL1|TEL2
' ============================================================
Public Function MemoPipe(abBlob() As Byte) As String
    Dim aFld()  As TMemoFeld
    Dim nCount  As Long
    aFld = MemoScan(abBlob, nCount)
    MemoPipe = _
        MemoFieldVal(aFld, nCount, "DAT", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "DAT", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "DAT", 3) & "|" & _
        MemoFieldVal(aFld, nCount, "FLT", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "FLT", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "TXT", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "TXT", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "TXT", 3) & "|" & _
        MemoFieldVal(aFld, nCount, "NUM", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "NUM", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "NUM", 3) & "|" & _
        MemoFieldVal(aFld, nCount, "TEL", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "TEL", 2)
End Function


' ============================================================
' Private Hilfsfunktionen
' ============================================================

' Ergebnis-Array dynamisch erweitern und Eintrag hinzufügen
Private Sub AppendResult(aRes() As TMemoFeld, _
                         nCount As Long, _
                         nAlloc As Long, _
                         sTyp As String, _
                         lPos As Long, _
                         lMax As Long, _
                         sENr As String, _
                         sWert As String)
    If nCount >= nAlloc Then
        nAlloc = nAlloc + 16
        ReDim Preserve aRes(nAlloc - 1)
    End If
    With aRes(nCount)
        .sTyp = sTyp
        .lPos = lPos
        .lMax = lMax
        .sENr = sENr
        .sWert = sWert
    End With
    nCount = nCount + 1
End Sub

' Byte-Sequenz ? ANSI/Windows-1252-String
Private Function BytesToStr(abData() As Byte, _
                            lStart As Long, _
                            lLen As Long) As String
    Dim ab() As Byte
    ReDim ab(lLen - 1)
    Dim i As Long
    For i = 0 To lLen - 1
        ab(i) = abData(lStart + i)
    Next i
    BytesToStr = StrConv(ab, vbUnicode)
End Function

' Ersten Null-Byte abschneiden und Rest ignorieren
Private Function TrimNulls(s As String) As String
    Dim p As Long
    p = InStr(s, Chr(0))
    If p > 0 Then TrimNulls = left(s, p - 1) Else TrimNulls = s
End Function

