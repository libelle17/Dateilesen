Dim sql$


Public type au
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Beginn as string
 Ende as string
 ICDs as string
 absPos as long
 AktZeit as date
 StByte as byte
end type

Public type briefe
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Pfad as string
 Art as string
 Name as string
 Typ as string
 AktZeit as date
 DokGroe as long
 qt as string
 qs as string
 StByte as long
end type

Public type diagnosen
 ID1 as long
 FID as long
 Pat_id as long
 GesName as string
 DiagDatum as date
 DiagSicherheit as string
 DiagText as string
 DiagSeite as string
 DiagAttr as string
 ICD as string
 obDauer as boolean
 AktZeit as date
 StByte as long
end type

Public type dokumente
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 DokPfad as string
 DokArt as string
 DokName as string
 Quelldatum as date
 AbsPos as long
 AktZeit as date
 DokGroe as long
 qs as string
 qt as string
 StByte as long
end type

Public type einträge
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Art as string
 Inhalt as string
 AbsPos as long
 AktZeit as date
 qs as string
 qt as string
 StByte as long
end type

Public type fälle
 FID as long
 Pat_ID as long
 Quartal as string
 Nachname as string
 Vorname as string
 lfdnr as long
 TMFNr as string
 VKNr as string
 BhFB as date
 BhFE1 as date
 BhFE2 as date
 f4202 as string
 ausgst as date
 f4106 as string
 f4107 as string
 lVorl as date
 IK as string
 KVKs as string
 KVKserg as string
 f4121 as string
 f4122 as string
 f4206 as string
 ÜwText as string
 f4210 as string
 statNuller as string
 ÜbwV as string
 AndÜw as string
 ÜWZiel as string
 ÜWNNr as string
 ÜWNaN as string
 ÜWTit as string
 ÜWVor as string
 ÜWVsw as string
 statKlasse as string
 f4237 as string
 statBehTage as long
 SchGr as string
 Weiterbeh as string
 PGeb as string
 PGebErg as string
 Mahnfrist as string
 GOÄKatNr as string
 GOÄKatName as string
 AdNam as string
 AdStr as string
 AdPlz as string
 AdOrt as string
 BhFE as date
 s8000 as string
 s8100 as string
 AktZeit as date
 Fanf as date
 altQuart as string
 qs as string
 qt as string
 qanf as date
 qend as string
 StByte as long
end type

Public type forminhaltneu
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
 StByte as long
end type

Public type kheinweis
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Ziel as string
 Diagnose as string
 AbsPos as long
 AktZeit as date
 StByte as long
end type

Public type lbanforderungen
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Text as string
 absPos as long
 AktZeit as date
 StByte as long
end type

Public type laborneu
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 FertigStGrad as string
 Abkü as string
 LangtextVW as long
 Wert as string
 Einheit as string
 KommentarVW as long
 AbsPos as long
 AktZeit as date
 Refnr as long
 StByte as long
end type

Public type leistungen
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Leistung as string
 f5002 as string
 f5005 as string
 f5006 as string
 f5009 as string
 f5015 as string
 f5016 as string
 f5021 as string
 f5026 as string
 absPos as long
 AktZeit as date
 qs as string
 qt as string
 StByte as long
end type

Public type medplan
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Datum as date
 Medikament as string
 MedAnfang as string
 FeldNr as integer
 mo as string
 mi as string
 nm as string
 ab as string
 zn as string
 bBed as boolean
 Bemerkung as string
 AbsPos as long
 AktZeit as date
 StByte as long
end type

Public type namen
 Pat_ID as long
 lfdnr as long
 NVorsatz as string
 Nachname as string
 Vorname as string
 GebDat as date
 Straße as string
 KVKStatus as string
 Geschlecht as string
 Plz as string
 Ort as string
 Weggeldzone as string
 AufnDat as date
 intZOGP as integer
 Titel as string
 Versichertennummer as string
 PrivatTel as string
 KVNr as string
 PrivatTel_2 as string
 PrivatFax as string
 DienstTel as string
 PrivatMobil as string
 Email as string
 AnAllgda as boolean
 An1da as boolean
 An2da as boolean
 Checkda as boolean
 DMTypaD as string
 AktZeit as date
 StByte as long
end type

Public type rezepteinträge
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 Rezept as string
 Rezeptklasse as string
 Medikament as string
 PZN as string
 absPos as long
 AktZeit as date
 qs as string
 qt as string
 StByte as long
end type

Public type rr
 FID as long
 Pat_ID as long
 ZeitPunkt as date
 RR as string
 absPos as long
 AktZeit as date
 StByte as long
end type

Public rAu() as au
Public rBr() as briefe
Public rDi() as diagnosen
Public rDo() as dokumente
Public rEi() as einträge
Public rFä() as fälle
Public rFo() as forminhaltneu
Public rKh() as kheinweis
Public rLb() as lbanforderungen
Public rLa() as laborneu
Public rLe() as leistungen
Public rMe() as medplan
Public rNa() as namen
Public rRe() as rezepteinträge
Public rRr() as rr

Public Function init()
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 ReDim rFä(0)
 ReDim rFo(0)
 ReDim rKh(0)
 ReDim rLb(0)
 ReDim rLa(0)
 ReDim rLe(0)
 ReDim rMe(0)
 ReDim rNa(0)
 ReDim rRe(0)
 ReDim rRr(0)
End Function

Public Function auSpeichern
 Dim i%
 sql = "delete from au where pat_id = " + CStr(rAu(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rAu)
  sql = "insert into au values(" + cStr(clng(rAu(i).FID)) + "," + cStr(clng(rAu(i).Pat_ID)) + "," + "'" + format(rAu(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rAu(i).Beginn) + "'" + "," + "'" + trim(rAu(i).Ende) + "'" + "," + "'" + trim(rAu(i).ICDs) + "'" + "," + cStr(clng(rAu(i).absPos)) _
  + "," + "'" + format(rAu(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rAu(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function briefeSpeichern
 Dim i%
 sql = "delete from briefe where pat_id = " + CStr(rBr(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rBr)
  sql = "insert into briefe values(" + cStr(clng(rBr(i).FID)) + "," + cStr(clng(rBr(i).Pat_ID)) + "," + "'" + format(rBr(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rBr(i).Pfad) + "'" + "," + "'" + trim(rBr(i).Art) + "'" + "," + "'" + trim(rBr(i).Name) + "'" + "," + "'" + trim(rBr(i).Typ) + "'" _
  + "," + "'" + format(rBr(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rBr(i).DokGroe)) + "," + "'" + trim(rBr(i).qt) + "'" + "," + "'" + trim(rBr(i).qs) + "'" _
  + "," + cStr(clng(rBr(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function diagnosenSpeichern
 Dim i%
 sql = "delete from diagnosen where pat_id = " + CStr(rDi(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rDi)
  sql = "insert into diagnosen values(" + cStr(clng(rDi(i).ID1)) + "," + cStr(clng(rDi(i).FID)) + "," + cStr(clng(rDi(i).Pat_id)) + "," + "'" + trim(rDi(i).GesName) + "'" _
  + "," + "'" + format(rDi(i).DiagDatum, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rDi(i).DiagSicherheit) + "'" + "," + "'" + trim(rDi(i).DiagText) + "'" _
  + "," + "'" + trim(rDi(i).DiagSeite) + "'" + "," + "'" + trim(rDi(i).DiagAttr) + "'" + "," + "'" + trim(rDi(i).ICD) + "'" + "," + cStr(clng(rDi(i).obDauer)) _
  + "," + "'" + format(rDi(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rDi(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function dokumenteSpeichern
 Dim i%
 sql = "delete from dokumente where pat_id = " + CStr(rDo(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rDo)
  sql = "insert into dokumente values(" + cStr(clng(rDo(i).FID)) + "," + cStr(clng(rDo(i).Pat_ID)) + "," + "'" + format(rDo(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rDo(i).DokPfad) + "'" + "," + "'" + trim(rDo(i).DokArt) + "'" + "," + "'" + trim(rDo(i).DokName) + "'" + "," + "'" + format(rDo(i).Quelldatum, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + cStr(clng(rDo(i).AbsPos)) + "," + "'" + format(rDo(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rDo(i).DokGroe)) + "," + "'" + trim(rDo(i).qs) + "'" _
  + "," + "'" + trim(rDo(i).qt) + "'" + "," + cStr(clng(rDo(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function einträgeSpeichern
 Dim i%
 sql = "delete from einträge where pat_id = " + CStr(rEi(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rEi)
  sql = "insert into einträge values(" + cStr(clng(rEi(i).FID)) + "," + cStr(clng(rEi(i).Pat_ID)) + "," + "'" + format(rEi(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rEi(i).Art) + "'" + "," + "'" + trim(rEi(i).Inhalt) + "'" + "," + cStr(clng(rEi(i).AbsPos)) + "," + "'" + format(rEi(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rEi(i).qs) + "'" + "," + "'" + trim(rEi(i).qt) + "'" + "," + cStr(clng(rEi(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function fälleSpeichern
 Dim i%
 sql = "delete from fälle where pat_id = " + CStr(rFä(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rFä)
  sql = "insert into fälle values(" + cStr(clng(rFä(i).FID)) + "," + cStr(clng(rFä(i).Pat_ID)) + "," + "'" + trim(rFä(i).Quartal) + "'" + "," + "'" + trim(rFä(i).Nachname) + "'" _
  + "," + "'" + trim(rFä(i).Vorname) + "'" + "," + cStr(clng(rFä(i).lfdnr)) + "," + "'" + trim(rFä(i).TMFNr) + "'" + "," + "'" + trim(rFä(i).VKNr) + "'" _
  + "," + "'" + format(rFä(i).BhFB, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + format(rFä(i).BhFE1, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + format(rFä(i).BhFE2, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rFä(i).f4202) + "'" + "," + "'" + format(rFä(i).ausgst, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rFä(i).f4106) + "'" + "," + "'" + trim(rFä(i).f4107) + "'" _
  + "," + "'" + format(rFä(i).lVorl, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rFä(i).IK) + "'" + "," + "'" + trim(rFä(i).KVKs) + "'" + "," + "'" + trim(rFä(i).KVKserg) + "'" _
  + "," + "'" + trim(rFä(i).f4121) + "'" + "," + "'" + trim(rFä(i).f4122) + "'" + "," + "'" + trim(rFä(i).f4206) + "'" + "," + "'" + trim(rFä(i).ÜwText) + "'" _
  + "," + "'" + trim(rFä(i).f4210) + "'" + "," + "'" + trim(rFä(i).statNuller) + "'" + "," + "'" + trim(rFä(i).ÜbwV) + "'" + "," + "'" + trim(rFä(i).AndÜw) + "'" _
  + "," + "'" + trim(rFä(i).ÜWZiel) + "'" + "," + "'" + trim(rFä(i).ÜWNNr) + "'" + "," + "'" + trim(rFä(i).ÜWNaN) + "'" + "," + "'" + trim(rFä(i).ÜWTit) + "'" _
  + "," + "'" + trim(rFä(i).ÜWVor) + "'" + "," + "'" + trim(rFä(i).ÜWVsw) + "'" + "," + "'" + trim(rFä(i).statKlasse) + "'" + "," + "'" + trim(rFä(i).f4237) + "'" _
  + "," + cStr(clng(rFä(i).statBehTage)) + "," + "'" + trim(rFä(i).SchGr) + "'" + "," + "'" + trim(rFä(i).Weiterbeh) + "'" + "," + "'" + trim(rFä(i).PGeb) + "'" _
  + "," + "'" + trim(rFä(i).PGebErg) + "'" + "," + "'" + trim(rFä(i).Mahnfrist) + "'" + "," + "'" + trim(rFä(i).GOÄKatNr) + "'" + "," + "'" + trim(rFä(i).GOÄKatName) + "'" _
  + "," + "'" + trim(rFä(i).AdNam) + "'" + "," + "'" + trim(rFä(i).AdStr) + "'" + "," + "'" + trim(rFä(i).AdPlz) + "'" + "," + "'" + trim(rFä(i).AdOrt) + "'" _
  + "," + "'" + format(rFä(i).BhFE, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rFä(i).s8000) + "'" + "," + "'" + trim(rFä(i).s8100) + "'" + "," + "'" + format(rFä(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + format(rFä(i).Fanf, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rFä(i).altQuart) + "'" + "," + "'" + trim(rFä(i).qs) + "'" + "," + "'" + trim(rFä(i).qt) + "'" _
  + "," + "'" + format(rFä(i).qanf, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rFä(i).qend) + "'" + "," + cStr(clng(rFä(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function forminhaltneuSpeichern
 Dim i%
 sql = "delete from forminhaltneu where pat_id = " + CStr(rFo(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rFo)
  sql = "insert into forminhaltneu values(" + cStr(clng(rFo(i).FID)) + "," + cStr(clng(rFo(i).Pat_ID)) + "," + cStr(clng(rFo(i).Form_ID)) + "," + cStr(clng(rFo(i).Form_AbkVW)) _
  + "," + "'" + format(rFo(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rFo(i).FeldNr)) + "," + cStr(clng(rFo(i).FeldVW)) + "," + cStr(clng(rFo(i).FeldInhVW)) _
  + "," + cStr(clng(rFo(i).AbsPos)) + "," + "'" + format(rFo(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rFo(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function kheinweisSpeichern
 Dim i%
 sql = "delete from kheinweis where pat_id = " + CStr(rKh(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rKh)
  sql = "insert into kheinweis values(" + cStr(clng(rKh(i).FID)) + "," + cStr(clng(rKh(i).Pat_ID)) + "," + "'" + format(rKh(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rKh(i).Ziel) + "'" + "," + "'" + trim(rKh(i).Diagnose) + "'" + "," + cStr(clng(rKh(i).AbsPos)) + "," + "'" + format(rKh(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + cStr(clng(rKh(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function lbanforderungenSpeichern
 Dim i%
 sql = "delete from lbanforderungen where pat_id = " + CStr(rLb(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rLb)
  sql = "insert into lbanforderungen values(" + cStr(clng(rLb(i).FID)) + "," + cStr(clng(rLb(i).Pat_ID)) + "," + "'" + format(rLb(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rLb(i).Text) + "'" + "," + cStr(clng(rLb(i).absPos)) + "," + "'" + format(rLb(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rLb(i).StByte)) _
  + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function laborneuSpeichern
 Dim i%
 sql = "delete from laborneu where pat_id = " + CStr(rLa(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rLa)
  sql = "insert into laborneu values(" + cStr(clng(rLa(i).FID)) + "," + cStr(clng(rLa(i).Pat_ID)) + "," + "'" + format(rLa(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rLa(i).FertigStGrad) + "'" + "," + "'" + trim(rLa(i).Abkü) + "'" + "," + cStr(clng(rLa(i).LangtextVW)) + "," + "'" + trim(rLa(i).Wert) + "'" _
  + "," + "'" + trim(rLa(i).Einheit) + "'" + "," + cStr(clng(rLa(i).KommentarVW)) + "," + cStr(clng(rLa(i).AbsPos)) + "," + "'" + format(rLa(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + cStr(clng(rLa(i).Refnr)) + "," + cStr(clng(rLa(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function leistungenSpeichern
 Dim i%
 sql = "delete from leistungen where pat_id = " + CStr(rLe(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rLe)
  sql = "insert into leistungen values(" + cStr(clng(rLe(i).FID)) + "," + cStr(clng(rLe(i).Pat_ID)) + "," + "'" + format(rLe(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rLe(i).Leistung) + "'" + "," + "'" + trim(rLe(i).f5002) + "'" + "," + "'" + trim(rLe(i).f5005) + "'" + "," + "'" + trim(rLe(i).f5006) + "'" _
  + "," + "'" + trim(rLe(i).f5009) + "'" + "," + "'" + trim(rLe(i).f5015) + "'" + "," + "'" + trim(rLe(i).f5016) + "'" + "," + "'" + trim(rLe(i).f5021) + "'" _
  + "," + "'" + trim(rLe(i).f5026) + "'" + "," + cStr(clng(rLe(i).absPos)) + "," + "'" + format(rLe(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rLe(i).qs) + "'" _
  + "," + "'" + trim(rLe(i).qt) + "'" + "," + cStr(clng(rLe(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function medplanSpeichern
 Dim i%
 sql = "delete from medplan where pat_id = " + CStr(rMe(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rMe)
  sql = "insert into medplan values(" + cStr(clng(rMe(i).FID)) + "," + cStr(clng(rMe(i).Pat_ID)) + "," + "'" + format(rMe(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + format(rMe(i).Datum, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rMe(i).Medikament) + "'" + "," + "'" + trim(rMe(i).MedAnfang) + "'" _
  + "," + cStr(clng(rMe(i).FeldNr)) + "," + "'" + trim(rMe(i).mo) + "'" + "," + "'" + trim(rMe(i).mi) + "'" + "," + "'" + trim(rMe(i).nm) + "'" + "," + "'" + trim(rMe(i).ab) + "'" _
  + "," + "'" + trim(rMe(i).zn) + "'" + "," + cStr(clng(rMe(i).bBed)) + "," + "'" + trim(rMe(i).Bemerkung) + "'" + "," + cStr(clng(rMe(i).AbsPos)) + "," + "'" + format(rMe(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + cStr(clng(rMe(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function namenSpeichern
 Dim i%
 sql = "delete from namen where pat_id = " + CStr(rNa(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rNa)
  sql = "insert into namen values(" + cStr(clng(rNa(i).Pat_ID)) + "," + cStr(clng(rNa(i).lfdnr)) + "," + "'" + trim(rNa(i).NVorsatz) + "'" + "," + "'" + trim(rNa(i).Nachname) + "'" _
  + "," + "'" + trim(rNa(i).Vorname) + "'" + "," + "'" + format(rNa(i).GebDat, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rNa(i).Straße) + "'" + "," + "'" + trim(rNa(i).KVKStatus) + "'" _
  + "," + "'" + trim(rNa(i).Geschlecht) + "'" + "," + "'" + trim(rNa(i).Plz) + "'" + "," + "'" + trim(rNa(i).Ort) + "'" + "," + "'" + trim(rNa(i).Weggeldzone) + "'" _
  + "," + "'" + format(rNa(i).AufnDat, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rNa(i).intZOGP)) + "," + "'" + trim(rNa(i).Titel) + "'" + "," + "'" + trim(rNa(i).Versichertennummer) + "'" _
  + "," + "'" + trim(rNa(i).PrivatTel) + "'" + "," + "'" + trim(rNa(i).KVNr) + "'" + "," + "'" + trim(rNa(i).PrivatTel_2) + "'" + "," + "'" + trim(rNa(i).PrivatFax) + "'" _
  + "," + "'" + trim(rNa(i).DienstTel) + "'" + "," + "'" + trim(rNa(i).PrivatMobil) + "'" + "," + "'" + trim(rNa(i).Email) + "'" + "," + cStr(clng(rNa(i).AnAllgda)) _
  + "," + cStr(clng(rNa(i).An1da)) + "," + cStr(clng(rNa(i).An2da)) + "," + cStr(clng(rNa(i).Checkda)) + "," + "'" + trim(rNa(i).DMTypaD) + "'" + "," + "'" + format(rNa(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + cStr(clng(rNa(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function rezepteinträgeSpeichern
 Dim i%
 sql = "delete from rezepteinträge where pat_id = " + CStr(rRe(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rRe)
  sql = "insert into rezepteinträge values(" + cStr(clng(rRe(i).FID)) + "," + cStr(clng(rRe(i).Pat_ID)) + "," + "'" + format(rRe(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rRe(i).Rezept) + "'" + "," + "'" + trim(rRe(i).Rezeptklasse) + "'" + "," + "'" + trim(rRe(i).Medikament) + "'" + "," + "'" + trim(rRe(i).PZN) + "'" _
  + "," + cStr(clng(rRe(i).absPos)) + "," + "'" + format(rRe(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + "'" + trim(rRe(i).qs) + "'" + "," + "'" + trim(rRe(i).qt) + "'" _
  + "," + cStr(clng(rRe(i).StByte)) + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function rrSpeichern
 Dim i%
 sql = "delete from rr where pat_id = " + CStr(rRr(0).Pat_id)
  call mycn.Execute(sql)
 For i = 0 to ubound(rRr)
  sql = "insert into rr values(" + cStr(clng(rRr(i).FID)) + "," + cStr(clng(rRr(i).Pat_ID)) + "," + "'" + format(rRr(i).ZeitPunkt, "yyyy-mm-dd hh:mm:ss") + "'" _
  + "," + "'" + trim(rRr(i).RR) + "'" + "," + cStr(clng(rRr(i).absPos)) + "," + "'" + format(rRr(i).AktZeit, "yyyy-mm-dd hh:mm:ss") + "'" + "," + cStr(clng(rRr(i).StByte)) _
  + ")"
  call mycn.Execute(sql)
Next i
End Function

Public Function doalleSpeichern
 call auSpeichern
 call briefeSpeichern
 call diagnosenSpeichern
 call dokumenteSpeichern
 call einträgeSpeichern
 call fälleSpeichern
 call forminhaltneuSpeichern
 call kheinweisSpeichern
 call lbanforderungenSpeichern
 call laborneuSpeichern
 call leistungenSpeichern
 call medplanSpeichern
 call namenSpeichern
 call rezepteinträgeSpeichern
 call rrSpeichern
End Function
