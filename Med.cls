VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediCl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvMed$
Private mvLT$
Private mvPat_ID&
Private mvIndex& 'lokale Kopie
Public mvAnzahl&
Public mvPatZahl&

Public Property Let Index(ByVal vData As Long)
    mvIndex = vData
End Property
Public Property Get Index() As Long
    Index = mvIndex
End Property
Public Property Let Anzahl(ByVal vData As Long)
    mvAnzahl = vData
End Property
Public Property Get Anzahl() As Long
    Anzahl = mvAnzahl
End Property

Public Function istKleinerAls(ByRef E2) As Boolean
 If mvMed < E2.Med Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
End Function
Public Function unGleich(ByRef E2) As Boolean
 If mvMed <> E2.Med Then unGleich = True
End Function
Public Property Get Med$()
    Med = mvMed
End Property

Public Property Let Med(ByVal vData$)
    mvMed = vData
End Property

Public Property Get LT$()
    LT = mvLT
End Property
Public Property Let LT(ByVal vData$)
    mvLT = vData
End Property

Public Property Get Pat_id&()
    Pat_id = mvPat_ID
End Property
Public Property Let Pat_id(ByVal vData&)
    mvPat_ID = vData
End Property

