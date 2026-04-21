Attribute VB_Name = "modMemoScan"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Public Type TMemoFeld
    sTyp  As String
    lPos  As Long
    lMAX  As Long
    lGrp  As Long
    sENr  As String
    sWert As String
End Type

Private Type TStackElem
    lAktMax As Long
    lIE     As Long
End Type


Public Function BlobFromField(vField As Variant) As Byte()
    Dim ab() As Byte
    If IsNull(vField) Then
        ReDim ab(0)
        BlobFromField = ab
        Exit Function
    End If
    Select Case VarType(vField)
        Case vbArray + vbByte
            ab = vField
        Case vbString
            ab = StrConv(CStr(vField), vbFromUnicode)
        Case Else
            ReDim ab(0)
    End Select
    BlobFromField = ab
End Function


' ============================================================
' MemoScan  –  Scanner + AID/Grp-Auflösung in einem Schritt
' Am Ende werden NUM-Felder an Top-Level-Grenzen als AID
' markiert und lGrp sequenziell neu belegt.
' ============================================================
Public Function MemoScan(abBlob() As Byte, _
                         nCount As Long) As TMemoFeld()
    '-- Ergebnis-Array --------------------------------------
    Dim aRes()   As TMemoFeld
    Dim nAlloc   As Long
    ReDim aRes(15)
    nAlloc = 16
    nCount = 0

    Dim lBLen   As Long     ' Länge des BLOB
    Dim lPos    As Long     ' aktueller Scan-Zeiger (0-basiert)
    Dim lGL     As Long     ' Gesamtlänge (aus erstem LE-Word)
    Dim lMAX    As Long     ' aktuelle Tiefenmarke
    Dim lAktMax  As Long
    Dim lTLen    As Long
    Dim lEndByte As Long
    Dim lSkipTo As Long     ' Sprungmarke nach TEL-Container
    Dim lB0     As Long     ' erstes Payload-Byte
    Dim lB1     As Long     ' zweites Payload-Byte
    Dim lB2& ' drittes Payload-Byte
    Dim lLL     As Long     ' Länge in TEL-Payloads
    Dim lPS     As Long     ' Payload-Startindex (0-basiert)
    Dim lRawGrp  As Long

    '-- Hierarchie-Stack (wie tmpeb in der SP) ---------------
    '   aStack(1..32): je ein Element pro Tiefenstufe
    '   .lAktMax = Grenzposition, .lIE = Stufenzähler
    Dim aStack(1 To 64) As TStackElem
    Dim nSTop   As Long     ' oberster belegter Stack-Index
    Dim lIE     As Long     ' aktuelle Tiefenstufe
    Dim lAltIE   As Long
    Dim i        As Long

    lBLen = UBound(abBlob) + 1
    If lBLen < 4 Then
        ReDim aRes(0)
        MemoScan = aRes
        Exit Function
    End If

    lPos = 0
    lSkipTo = 0
    lIE = 0
    nSTop = 0

    lGL = CLng(abBlob(0)) + CLng(abBlob(1)) * 256
    lMAX = lGL
    lPos = 1

    Do While lPos < lBLen

        If lPos < lSkipTo Then
            lPos = lPos + 1
            GoTo CheckEnd
        End If

        If lPos + 1 >= lBLen Then GoTo CheckEnd

        '-- tlen: LE-Word aus Byte[lPos] und Byte[lPos+1] --
        lTLen = CLng(abBlob(lPos)) + CLng(abBlob(lPos + 1)) * 256
        lAktMax = lTLen + lPos + 1

        If lAktMax >= 1 And lAktMax < lBLen Then
            lEndByte = CLng(abBlob(lAktMax))
        Else
            lEndByte = -1
        End If

        If lTLen > 0 _
        And lAktMax >= 1 And lAktMax <= lGL + 2 _
        And ((lEndByte = 0 And lAktMax <= lMAX) Or lPos > lMAX) Then

            lAltIE = lIE
            If lAktMax <= lMAX Then
                lIE = lIE + 1
            Else
                lIE = 0
                For i = nSTop To 1 Step -1
                    If aStack(i).lAktMax >= lAktMax Then
                        lIE = aStack(i).lIE + 1
                        Exit For
                    End If
                Next i
                If lIE = 0 Then lIE = 1
            End If

            If lIE <= lAltIE Then
                If lIE < lAltIE Then
                    For i = lIE + 1 To nSTop
                        aStack(i).lAktMax = 0
                        aStack(i).lIE = 0
                    Next i
                    nSTop = lIE
                End If
                If nSTop >= lIE Then
                    aStack(lIE).lIE = aStack(lIE).lIE + 1
                End If
            Else
                If lIE > nSTop Then nSTop = lIE
                If lIE <= 64 Then
                    aStack(lIE).lAktMax = lAktMax
                    aStack(lIE).lIE = 1
                End If
            End If

            Dim sENr As String
            sENr = ""
            For i = 1 To nSTop
                If aStack(i).lIE > 0 Then
                    If sENr <> "" Then sENr = sENr & "."
                    sENr = sENr & CStr(aStack(i).lIE)
                End If
            Next i

            If nSTop >= 1 Then
                lRawGrp = aStack(1).lIE
            Else
                lRawGrp = 0
            End If

            lMAX = lAktMax
            lPS = lPos + 2
            If lPS + lTLen - 1 >= lBLen Then GoTo IncPos

            Dim bAllNull As Boolean
            bAllNull = True
            For i = 0 To lTLen - 1
                If abBlob(lPS + i) <> 0 Then
                    bAllNull = False
                    Exit For
                End If
            Next i
' -- Null-Payload nur für tlen >= 4 überspringen --
            If bAllNull And lTLen >= 4 Then
                lPos = lPos + lTLen
                GoTo CheckEnd
            End If
            
            lB0 = CLng(abBlob(lPS))
            lB1 = CLng(abBlob(lPS + 1))
            lB2 = CLng(abBlob(lPS + 2))

            If lB0 = &H2 And lB1 = &H0 _
               And lTLen >= 6 And lTLen <= 100 Then
                If lPS + 5 < lBLen Then
                    lLL = CLng(abBlob(lPS + 4)) + _
                          CLng(abBlob(lPS + 5)) * 256
                    If lLL > 0 And lLL + 6 <= lTLen Then
                        Dim sTelA As String
                        sTelA = TrimNulls(BytesToStr(abBlob, lPS + 6, lLL))
                        If IsPhoneStr(sTelA) Then
                            AppendRec aRes, nCount, nAlloc, _
                                "TEL", lPos + 1, lMAX, lRawGrp, sENr, sTelA
                            lSkipTo = lPos + lTLen + 2
                            lPos = lSkipTo - 1
                            GoTo CheckEnd
                        End If
                    End If
                End If

            ElseIf lB0 = &H2 And lB1 = &H1 _
                   And lTLen >= 3 And lTLen <= 100 Then
                If lPS + 2 < lBLen Then
                    lLL = CLng(abBlob(lPS + 2))
                    If lLL > 0 And lLL + 3 <= lTLen Then
                        Dim sTelB As String
                        sTelB = TrimNulls(BytesToStr(abBlob, lPS + 3, lLL))
                        If IsPhoneStr(sTelB) Then
                            AppendRec aRes, nCount, nAlloc, _
                                "TEL", lPos + 1, lMAX, lRawGrp, sENr, sTelB
                            lSkipTo = lPos + lTLen + 2
                            lPos = lSkipTo - 1
                            GoTo CheckEnd
                        End If
                    End If
                End If

            ElseIf lTLen = 4 _
                   And lPS + 3 < lBLen _
                   And CLng(abBlob(lPS + 3)) = 0 Then
                Select Case lB0: Case 65 To 90, 97 To 122, 196, 214, 220, 223, 228, 246, 252: Case Else: GoTo clr: End Select
                Select Case lB1: Case 65 To 90, 97 To 122, 196, 214, 220, 223, 228, 246, 252: Case Else: GoTo clr: End Select
                Select Case lB2: Case 65 To 90, 97 To 122, 196, 214, 220, 223, 228, 246, 252: Case Else: GoTo clr: End Select
                GoTo txt:
clr:
                AppendRec aRes, nCount, nAlloc, _
                    "CLR", lPos + 1, lMAX, lRawGrp, sENr, _
                    Right("0" & Hex(lB0), 2) & _
                    Right("0" & Hex(lB1), 2) & _
                    Right("0" & Hex(lB2), 2)

            ElseIf lTLen = 4 And lB1 = &H7 Then
                If lPS + 3 < lBLen Then
                    Dim lMon As Long
                    Dim lTag As Long
                    lMon = CLng(abBlob(lPS + 2))
                    lTag = CLng(abBlob(lPS + 3))
                    If lMon >= 1 And lMon <= 12 _
                       And lTag >= 1 And lTag <= 31 Then
                        AppendRec aRes, nCount, nAlloc, _
                            "DAT", lPos + 1, lMAX, lRawGrp, sENr, _
                            Format(lTag, "00") & "." & _
                            Format(lMon, "00") & "." & _
                            CStr(lB0 + &H700)
                    End If
                End If

            ' NUM: 2 Bytes LE
            ElseIf lTLen = 2 Then
                Dim lNum As Long
                lNum = lB0 + lB1 * 256
                If lNum >= 1 And lNum <= 65535 Then
                    ' Wert 8 = Texttyp-Indikator: als TYP speichern
                    ' (intern für Separator-Erkennung, nicht anzeigen)
                    Dim sNumTyp As String
                    If lNum = 8 Then
                        sNumTyp = "TYP"
                    Else
                        sNumTyp = "NUM"
                    End If
                    AppendRec aRes, nCount, nAlloc, _
                        sNumTyp, lPos + 1, lMAX, lRawGrp, sENr, CStr(lNum)
                End If
            ElseIf lTLen = 8 _
                   And lPS + 7 < lBLen _
                   And CLng(abBlob(lPS + 7)) >= &H3F _
                   And CLng(abBlob(lPS + 7)) <= &H45 Then
                Dim ab8(7) As Byte
                Dim dVal   As Double
                For i = 0 To 7
                    ab8(i) = abBlob(lPS + i)
                Next i
                CopyMemory dVal, ab8(0), 8
                Dim sFlt As String
                sFlt = CStr(CDec(Format(dVal, "0.######")))
                Do While Right(sFlt, 1) = "0" And InStr(sFlt, ".") > 0
                    sFlt = left(sFlt, Len(sFlt) - 1)
                Loop
                If Right(sFlt, 1) = "." Then
                    sFlt = left(sFlt, Len(sFlt) - 1)
                End If
                AppendRec aRes, nCount, nAlloc, _
                    "FLT", lPos + 1, lMAX, lRawGrp, sENr, sFlt

            ElseIf lTLen >= 3 And lTLen <= 2000 And lB0 >= &H20 Then
txt:
                Dim sTxt As String
                sTxt = RTrim(TrimNulls(BytesToStr(abBlob, lPS, lTLen)))
                If sTxt <> "" Then
                    AppendRec aRes, nCount, nAlloc, _
                        "TXT", lPos + 1, lMAX, lRawGrp, sENr, sTxt
                End If
            ' -- FLG: 1 Byte (einfaches Flag, auch 0) ---------
            ElseIf lTLen = 1 Then
                AppendRec aRes, nCount, nAlloc, _
                    "FLG", lPos + 1, lMAX, lRawGrp, sENr, _
                    CStr(CLng(abBlob(lPS)))
            End If
        End If

IncPos:
        lPos = lPos + 1
CheckEnd:
        If lPos >= lGL Then Exit Do
    Loop

    If nCount > 0 Then
        ReDim Preserve aRes(nCount - 1)
    Else
        ReDim aRes(0)
    End If

' -- Pass 1: AID-Erkennung --------------------------------
    Dim j        As Long
    Dim m        As Long
    Dim lVal     As Long
    Dim bIsPair  As Boolean
    Dim bHasSep  As Boolean
    Dim bLongNext As Boolean
    Dim bNextSh  As Boolean

For j = 0 To nCount - 1

        If aRes(j).sTyp <> "NUM" And aRes(j).sTyp <> "TYP" Then GoTo NextNum

        lVal = CLng(aRes(j).sWert)

        If aRes(j).sTyp = "TYP" Then
            bNextSh = False
            Dim k1 As Long
            k1 = j + 1
            Do While k1 < nCount
                If aRes(k1).sTyp = "FLG" Then
                    k1 = k1 + 1
                Else
                    Exit Do
                End If
            Loop
            If k1 < nCount Then
                If aRes(k1).sTyp = "TXT" _
                   And Len(aRes(k1).sWert) <= 6 Then
                    bNextSh = True
                End If
            End If
            If Not bNextSh Then GoTo NextNum
        End If

        If lVal >= 97 And lVal <= 127 Then GoTo NextNum

        ' Paar-Prüfung: FLG-Felder überspringen
        bIsPair = False
        Dim nPaar As Long
        If lVal <= 127 Then
            For m = 1 To 4
                nPaar = j + m
                Do While nPaar < nCount And aRes(nPaar).sTyp = "FLG"
                    nPaar = nPaar + 1
                Loop
                If nPaar >= nCount Then Exit For
                If aRes(nPaar).sTyp = "NUM" Or aRes(nPaar).sTyp = "TYP" Then
                    If aRes(nPaar).sWert = aRes(j).sWert Then bIsPair = True
                    Exit For
                Else
                    Exit For
                End If
            Next m
            If Not bIsPair Then
                For m = 1 To 4
                    nPaar = j - m
                    Do While nPaar >= 0 And aRes(nPaar).sTyp = "FLG"
                        nPaar = nPaar - 1
                    Loop
                    If nPaar < 0 Then Exit For
                    If aRes(nPaar).sTyp = "NUM" Or aRes(nPaar).sTyp = "TYP" Then
                        If aRes(nPaar).sWert = aRes(j).sWert Then bIsPair = True
                        Exit For
                    Else
                        Exit For
                    End If
                Next m
            End If
        Else
            Dim jM1 As Long, jP1 As Long
            jM1 = j - 1
            Do While jM1 >= 0 And aRes(jM1).sTyp = "FLG": jM1 = jM1 - 1: Loop
            jP1 = j + 1
            Do While jP1 < nCount And aRes(jP1).sTyp = "FLG": jP1 = jP1 + 1: Loop
            If jM1 >= 0 Then
                If aRes(jM1).sTyp = "NUM" _
                   And aRes(jM1).sWert = aRes(j).sWert Then bIsPair = True
            End If
            If Not bIsPair And jP1 < nCount Then
                If aRes(jP1).sTyp = "NUM" _
                   And aRes(jP1).sWert = aRes(j).sWert Then bIsPair = True
            End If
        End If

        ' Separator-Prüfung: FLG-Felder überspringen
        bHasSep = False
        If Not bIsPair Then
            Dim nSep As Long
            nSep = 0
            For m = 1 To j + 20
                If j + m >= nCount Then Exit For
                If aRes(j + m).sTyp = "FLG" Then
                    ' überspringen
                ElseIf aRes(j + m).sTyp = "NUM" Or aRes(j + m).sTyp = "TYP" Then
                    Exit For
                ElseIf aRes(j + m).sTyp = "TXT" Then
                    nSep = nSep + 1
                    If left(aRes(j + m).sWert, 3) = "---" Then
                        bHasSep = True
                        Exit For
                    End If
                    If nSep >= 3 Then Exit For
                End If
            Next m
        End If

        ' LongNext: nächstes Nicht-FLG-Feld
        bLongNext = False
        If Not bIsPair And Not bHasSep And lVal <= 127 Then
            Dim jNxt As Long
            jNxt = j + 1
            Do While jNxt < nCount And aRes(jNxt).sTyp = "FLG"
                jNxt = jNxt + 1
            Loop
            If jNxt >= nCount Then
                bLongNext = True
            ElseIf aRes(jNxt).sTyp = "TXT" Then
                If left(aRes(jNxt).sWert, 7) = "Text - " Then
                    bLongNext = True
                End If
            ElseIf aRes(jNxt).sTyp = "NUM" Then
                If CLng(aRes(jNxt).sWert) > 127 Then
                    bLongNext = True
                End If
            End If
        End If

        ' Gruppengrenze: nächstes Nicht-FLG-Feld
        If Not bIsPair And Not bHasSep And Not bLongNext Then
            Dim jNxt2 As Long
            jNxt2 = j + 1
            Do While jNxt2 < nCount And aRes(jNxt2).sTyp = "FLG"
                jNxt2 = jNxt2 + 1
            Loop
            If jNxt2 >= nCount Then
                aRes(j).sTyp = "AID"
            ElseIf aRes(jNxt2).lGrp <> aRes(j).lGrp Then
                aRes(j).sTyp = "AID"
            End If
        End If

NextNum:
    Next j

    ' Pass 2: lGrp
    Dim lCurGrp As Long
    lCurGrp = 0
    For j = 0 To nCount - 1
        If aRes(j).sTyp = "AID" Then lCurGrp = lCurGrp + 1
        aRes(j).lGrp = lCurGrp
    Next j
    
    MemoScan = aRes
End Function


' ============================================================
' MemoFieldVal  –  n-tes Feld eines Typs (1-basiert)
' ============================================================
Public Function MemoFieldVal(aFld() As TMemoFeld, _
                             nCount As Long, _
                             sTyp As String, _
                             n As Long) As String
    Dim i      As Long
    Dim nFound As Long
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
' MemoPipe  –  Pipe-String
' 1-3 DAT, 4-5 FLT, 6-8 TXT, 9-11 NUM,
' 12-13 TEL, 14-15 CLR, 16-17 AID
' ============================================================
Public Function MemoPipe(abBlob() As Byte) As String
    Dim aFld() As TMemoFeld
    Dim nCount As Long
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
        MemoFieldVal(aFld, nCount, "TEL", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "CLR", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "CLR", 2) & "|" & _
        MemoFieldVal(aFld, nCount, "AID", 1) & "|" & _
        MemoFieldVal(aFld, nCount, "AID", 2)
End Function


' ============================================================
' Private Hilfsfunktionen
' ============================================================

Private Sub AppendRec(aRes() As TMemoFeld, _
                      nCount As Long, _
                      nAlloc As Long, _
                      sTyp As String, _
                      lPos As Long, _
                      lMAX As Long, _
                      lRawGrp As Long, _
                      sENr As String, _
                      sWert As String)
    If nCount >= nAlloc Then
        nAlloc = nAlloc + 16
        ReDim Preserve aRes(nAlloc - 1)
    End If
    With aRes(nCount)
        .sTyp = sTyp
        .lPos = lPos
        .lMAX = lMAX
        .lGrp = lRawGrp
        .sENr = sENr
        .sWert = sWert
    End With
    nCount = nCount + 1
End Sub

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

Private Function TrimNulls(s As String) As String
    Dim p As Long
    p = InStr(s, Chr(0))
    If p > 0 Then
        TrimNulls = left(s, p - 1)
    Else
        TrimNulls = s
    End If
End Function

Private Function IsPhoneStr(s As String) As Boolean
    Dim i As Long
    Dim c As Long
    IsPhoneStr = False
    If Len(s) = 0 Then Exit Function
    c = Asc(left(s, 1))
    If c < 48 Or (c > 57 And c <> 43 And c <> 40) Then Exit Function
    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        Select Case c
            Case 48 To 57
            Case 43, 42, 35, 40, 41, 47, 32, 45
            Case Else
                Exit Function
        End Select
    Next i
    IsPhoneStr = True
End Function



' ============================================================
' LooksLikeName  –  ist s ein Anzeige-Name oder ein Kürzel?
' Name: Länge > 8, oder enthält " - ", oder enthält Leerzeichen
'       und beginnt mit Großbuchstabe
' ============================================================
Private Function LooksLikeName(s As String) As Boolean
    If Len(s) = 0 Then LooksLikeName = False: Exit Function

    Dim c As Long
    c = Asc(left(s, 1))

    ' Beginnt mit Großbuchstabe (inkl. Ä Ö Ü) ? Name
    If (c >= 65 And c <= 90) _
       Or c = 196 Or c = 214 Or c = 220 Then
        LooksLikeName = True
        Exit Function
    End If

    ' Länge > 8 ? Name
    If Len(s) > 8 Then LooksLikeName = True: Exit Function

    ' Enthält " - " ? Name
    If InStr(s, " - ") > 0 Then LooksLikeName = True: Exit Function

    LooksLikeName = False
End Function

' ============================================================
' MemoGrpFelder  –  v4
' ============================================================
Public Sub MemoGrpFelder(aFld() As TMemoFeld, _
                         nCount As Long, _
                         lGrp As Long, _
                         sAIDNr As String, _
                         sKuerzInt As String, _
                         sName As String, _
                         sKuerzExt As String, _
                         sKuerzDrk As String, _
                         sTaste As String, _
                         sFarbe As String, _
                         bHatText As Boolean, _
                         bVerwendung As Boolean)
    Dim i           As Long
    Dim bNameFound  As Boolean
    Dim bFlgSeen    As Boolean
    Dim bTypNachFlg As Boolean
    Dim nPostName   As Long

    sAIDNr = ""
    sKuerzInt = ""
    sName = ""
    sKuerzExt = ""
    sKuerzDrk = ""
    sTaste = ""
    sFarbe = ""
    bHatText = False
    bVerwendung = True
    bNameFound = False
    bFlgSeen = False
    bTypNachFlg = False
    nPostName = 0

    For i = 0 To nCount - 1
        If aFld(i).lGrp <> lGrp Then GoTo weiter
        Select Case aFld(i).sTyp
            Case "AID"
                sAIDNr = aFld(i).sWert

            Case "FLG"
                bFlgSeen = True

            Case "TYP"
                bHatText = True
                If Not bNameFound And bFlgSeen Then
                    bTypNachFlg = True
                End If
                
            Case "TXT"
                Dim bIsSub  As Boolean
                Dim bIsName As Boolean
                bIsSub = (InStr(aFld(i).sENr, ".") > 0)
                bIsName = IsNameTxt(aFld(i).sWert, bFlgSeen, bIsSub)

                If bIsName Then
                    If Not bNameFound Then
                        sName = aFld(i).sWert
                        bNameFound = True
                        bVerwendung = Not bTypNachFlg
                    End If
                Else
                    If sKuerzInt = "" And Not bNameFound Then
                        sKuerzInt = aFld(i).sWert
                    ElseIf bNameFound Then
                        nPostName = nPostName + 1
                        Select Case nPostName
                            Case 1: sKuerzExt = aFld(i).sWert
                            Case 2: sKuerzDrk = aFld(i).sWert
                        End Select
                    End If
                End If

            Case "NUM"
                If sTaste = "" Then
                    Dim lT As Long
                    lT = CLng(aFld(i).sWert)
                    If lT >= 32 And lT <= 127 Then
                        sTaste = Chr(lT)
                    End If
                End If

            Case "CLR"
                sFarbe = aFld(i).sWert

        End Select
weiter:
    Next i

    If Not bNameFound Then bVerwendung = True
End Sub


' ============================================================
' IsNameTxt  –  ist s ein Anzeige-Name?
'
' Regeln:
'   Sub-Container (bIsSub)  ? immer Name
'   Enthält " - "           ? immer Name
'   Nach FLG (bFlgSeen):
'     Beginnt mit Großbuchstabe ? Name
'     Länge > 8               ? Name
'   Vor FLG (Not bFlgSeen):
'     Länge > 8 UND Großbuchstabe ? Name
'     (kurze Großbuchstaben-Kürzel wie "UEBLABOR" len=8 ? Kürzel)
' ============================================================
Private Function IsNameTxt(s As String, _
                           bFlgSeen As Boolean, _
                           bIsSub As Boolean) As Boolean
    IsNameTxt = False
    If Len(s) = 0 Then Exit Function

    ' Sub-Container ? immer Name
    If bIsSub Then IsNameTxt = True: Exit Function

    ' Enthält " - " ? immer Name (Praxisformular, Text - X, etc.)
    If InStr(s, " - ") > 0 Then IsNameTxt = True: Exit Function

    Dim c As Long
    c = Asc(left(s, 1))
    Dim bGross As Boolean
    bGross = (c >= 65 And c <= 90) Or c = 196 Or c = 214 Or c = 220

    If bFlgSeen Then
        ' Nach FLG: Großbuchstabe oder Länge > 8
        If bGross Or Len(s) > 8 Then IsNameTxt = True: Exit Function
    Else
        ' Vor FLG: NUR wenn Länge > 8 UND Großbuchstabe
        ' (schließt "UEBLABOR" len=8, "Text" len=4 etc. aus)
        If Len(s) > 8 And bGross Then IsNameTxt = True: Exit Function
    End If
End Function

