
type diagnosen
 ID1 as long
 FID as long
 Pat_id as long
 GesName as string * 50
 DiagDatum as date
 DiagSicherheit as string * 1
 DiagText as string
 DiagSeite as string * 1
 DiagAttr as string * 60
 ICD as string * 10
 obDauer as boolean
 AktZeit as date
end type

type fälle
 FID as long
 Pat_ID as long
 Quartal as string * 5
 Nachname as string * 50
 Vorname as string * 50
 lfdnr as long
 TMFNr as string * 20
 VKNr as string * 10
 BhFB as date
 BhFE1 as date
 BhFE2 as date
 f4202 as string * 5
 ausgst as date
 f4106 as string * 5
 f4107 as string * 5
 lVorl as date
 IK as string * 50
 KVKs as string * 10
 KVKserg as string * 5
 f4121 as string * 5
 f4122 as string * 5
 f4206 as string * 15
 ÜwText as string
 f4210 as string * 5
 statNuller as string * 20
 ÜbwV as string * 10
 AndÜw as string * 10
 ÜWZiel as string * 50
 ÜWNNr as string * 10
 ÜWNaN as string * 50
 ÜWTit as string * 39
 ÜWVor as string * 25
 ÜWVsw as string * 10
 statKlasse as string * 5
 f4237 as string * 5
 statBehTage as long
 SchGr as string * 5
 Weiterbeh as string * 70
 PGeb as string * 70
 PGebErg as string * 50
 Mahnfrist as string * 8
 GOÄKatNr as string * 10
 GOÄKatName as string * 30
 AdNam as string * 20
 AdStr as string * 40
 AdPlz as string * 10
 AdOrt as string * 30
 BhFE as date
 s8000 as string * 4
 s8100 as string * 7
 AktZeit as date
 Fanf as date
 altQuart as string * 5
 qs as string * 5
 qt as string * 5
 qanf as date
 qend as string * 10
end type

type forminhaltneu
 FID as long
 Pat_ID as long
 Form_ID as long
 Form_AbkVW as long
 ZeitPunkt as date
 FeldNr as long
 FeldVW as long
 FeldInhVW as long
 AbsPos as long
 AktZeit as date
end type

type medplan
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Datum as date
 Medikament as string * 75
 MedAnfang as string * 35
 FeldNr as 
 mo as string * 10
 mi as string * 10
 nm as string * 10
 ab as string * 10
 zn as string * 10
 bBed as boolean
 Bemerkung as string
 AbsPos as long
 AktZeit as date
end type

type dokumente
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 DokPfad as string * 150
 DokArt as string * 40
 DokName as string * 170
 Quelldatum as date
 AbsPos as long
 AktZeit as date
 DokGroe as long
 qs as string * 5
 qt as string * 5
end type

type laborneu
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 FertigStGrad as string * 1
 Abkü as string * 15
 LangtextVW as long
 Wert as string * 15
 Einheit as string * 20
 KommentarVW as long
 AbsPos as long
 AktZeit as date
 Refnr as long
end type

type einträge
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Art as string * 15
 Inhalt as string
 AbsPos as long
 AktZeit as date
 qs as string * 5
 qt as string * 5
end type

type kheinweis
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Ziel as string * 100
 Diagnose as string
 AbsPos as long
 AktZeit as date
end type

type briefe
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Pfad as string * 128
 Art as string * 40
 Name as string * 150
 Typ as string * 50
 AktZeit as date
 DokGroe as long
 qt as string * 5
 qs as string * 5
end type

type leistungen
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Leistung as string * 10
 f5002 as string * 32
 f5005 as string * 2
 f5006 as string * 20
 f5009 as string * 80
 f5015 as string * 30
 f5016 as string * 28
 f5021 as string * 20
 f5026 as string * 20
 absPos as long
 AktZeit as date
 qs as string * 5
 qt as string * 5
end type

type rezepteinträge
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Rezept as string * 3
 Rezeptklasse as string * 5
 Medikament as string * 50
 PZN as string * 20
 absPos as long
 AktZeit as date
 qs as string * 5
 qt as string * 5
end type

type au
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Beginn as string * 8
 Ende as string * 8
 ICDs as string * 100
 absPos as long
 AktZeit as date
end type

type rr
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 RR as string
 absPos as long
 AktZeit as date
end type

type laboranforderungen
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Text as string
 absPos as long
 AktZeit as date
end type

type namen
 Pat_ID as long
 lfdnr as long
 NVorsatz as string * 10
 Nachname as string * 50
 Vorname as string * 50
 GebDat as date
 Straße as string * 50
 KVKStatus as string * 1
 Geschlecht as string * 1
 Plz as string * 20
 Ort as string * 70
 Weggeldzone as string * 1
 AufnDat as date
 Titel as string * 39
 Versichertennummer as string * 30
 PrivatTel as string * 100
 KVNr as string * 10
 PrivatTel_2 as string * 50
 PrivatFax as string * 50
 DienstTel as string * 50
 PrivatMobil as string * 50
 Email as string * 100
 AnAllgda as boolean
 An1da as boolean
 An2da as boolean
 Checkda as boolean
 DMTypaD as string * 1
 AktZeit as date
end type
