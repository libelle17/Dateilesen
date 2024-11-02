VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sLpar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Zu sortierende Einheit"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Private mvarInhalt AS Variant 'lokale Kopie
Private mvAbkü$
Private mvLabor$
Public Langtext$
Private mvEinheit$
Private mvklEinheit$
Private mvNBm$
Private mvoNm$
Private mvuNm$
Private mvarIndex As Long 'lokale Kopie

Public Property Let Index(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.index = 5
    mvarIndex = vData
End Property

Public Property Get Index() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.index
    Index = mvarIndex
End Property ' Get Index()

Public Property Get SortVar()
 SortVar = mvAbkü
End Property ' Get SortVar()

Public Function istKleinerAls(ByRef E2) As Boolean
Attribute istKleinerAls.VB_Description = "Wahr, wenn das erste Element größer ist"
 On Error GoTo fehler
 If mvAbkü < E2.Abkü Or (mvAbkü = E2.Abkü And mvklEinheit < E2.klEinheit) Or _
 (mvAbkü = E2.Abkü And mvklEinheit = E2.klEinheit) And ( _
    (E2.oNm <> "" And mvoNm < E2.oNm) Or _
    ((E2.oNm <> "" And mvoNm = E2.oNm) And (E2.uNm <> "" And mvuNm < E2.uNm)) _
 ) Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in istKleinerAls/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function '  istKleinerAls

Public Function unGleich(ByRef E2) As Boolean
 On Error GoTo fehler
 If mvAbkü <> E2.Abkü Or (mvAbkü = E2.Abkü And mvklEinheit <> E2.klEinheit) Or _
 (mvAbkü = E2.Abkü And mvklEinheit = E2.klEinheit) And ( _
    ((E2.oNm <> "" And mvoNm <> E2.oNm)) Or _
    ((E2.oNm <> "" And mvoNm = E2.oNm) And (E2.uNm <> "" And mvuNm <> E2.uNm)) _
 ) Then
  unGleich = True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in unGleich/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' unGleich

Public Property Let Abkü(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInh = 5
 On Error GoTo fehler
 mvAbkü = UCase$(vData)
 Exit Property
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Let Abkü/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property

Public Property Get Abkü() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInh
'    IF IsObject(mvFeldInh) THEN
'        SET FeldInh = mvFeldInh
'    Else
       Abkü = mvAbkü
'    END IF
End Property

Public Property Let Labor(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInh = 5
 On Error GoTo fehler
 mvLabor = UCase$(vData)
 Exit Property
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Let Labor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property

Public Property Get Labor() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInh
'    IF IsObject(mvFeldInh) THEN
'        SET FeldInh = mvFeldInh
'    Else
       Labor = mvLabor
'    END IF
End Property

Public Property Let Einheit(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInhVW = 5
    mvEinheit = vData
    mvklEinheit = LCase$(vData)
End Property

Public Property Get Einheit() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInhVW
'    IF IsObject(mvFeldInhVW) THEN
'        SET FeldInhVW = mvFeldInhVW
'    Else
        Einheit = mvEinheit
'    END IF
End Property

Public Property Get klEinheit() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInhVW
'    IF IsObject(mvFeldInhVW) THEN
'        SET FeldInhVW = mvFeldInhVW
'    Else
        klEinheit = mvklEinheit
'    END IF
End Property

Public Property Let NBm(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInh = 5
 mvNBm = UCase$(vData)
 Exit Property
End Property

Public Property Get NBm() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInh
'    IF IsObject(mvFeldInh) THEN
'        SET FeldInh = mvFeldInh
'    Else
       NBm = mvNBm
'    END IF
End Property

Public Property Let oNm(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInh = 5
 mvoNm = UCase$(vData)
 Exit Property
End Property

Public Property Get oNm() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInh
'    IF IsObject(mvFeldInh) THEN
'        SET FeldInh = mvFeldInh
'    Else
       oNm = mvoNm
'    END IF
End Property

Public Property Let uNm(ByVal vData As Variant)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.FeldInh = 5
 mvuNm = UCase$(vData)
 Exit Property
End Property

Public Property Get uNm() As Variant
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.FeldInh
'    IF IsObject(mvFeldInh) THEN
'        SET FeldInh = mvFeldInh
'    Else
       uNm = mvuNm
'    END IF
End Property

