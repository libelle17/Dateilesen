Dim sql$

Public type namen
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 An1da as integer 'An1da 11 'Anamnese S.1 da
 An2da as integer 'An2da 11 'Anamnese S.2 da
 AnAllgda as integer 'AnAllgda 11 'Anamnese allgemein da
 AufnDat as date 'AufnDat 7 '3610
 Cave as string 'Cave 130 '3654
 Checkda as integer 'Checkda 11 'Checkliste da
 DienstTel as string 'DienstTel 130 '3629
 DMTypaD as string 'DMTypaD 130 'aus Diagnosen
 Email as string 'Email 130 'Email
 GebDat as date 'GebDat 7 '3103
 Geschlecht as string 'Geschlecht 130 '3110
 intZOGP as string 'intZOGP 130 '3635, interne Zuordnung Arzt bei GP
 KVKStatus as string 'KVKStatus 130 '3108
 KVNr as string 'KVNr 130 '3630
 lfdnr as long 'lfdnr 3 'laufende Patientennummer
 Nachname as string 'Nachname 130 '3101
 Notiz as string 'Notiz 130 '3634
 NVorsatz as string 'NVorsatz 130 '3100
 Ort as string 'Ort 130 '3113
 Pat_ID as long 'Pat_ID 3 '3000
 Plz as string 'Plz 130 '3112
 PrivatFax as string 'PrivatFax 130 '3629
 PrivatMobil as string 'PrivatMobil 130 '3629
 PrivatTel as string 'PrivatTel 130 '3629
 PrivatTel_2 as string 'PrivatTel_2 130 '3629
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Straße as string 'Straße 130 '3107
 Titel as string 'Titel 130 '3104
 Versichertennummer as string 'Versichertennummer 130 '3105
 Verwandt as string 'Verwandt 130 '3632
 Vorname as string 'Vorname 130 '3102
 Weggeldzone as string 'Weggeldzone 130 '3631
 zubenach as string 'zubenach 130 '3633
end type

Public type faelle
 AbrAr as string 'AbrAr 130 '4107, Abrechnungsart (1 = Primärkassen)
 AbrGb as string 'AbrGb 130 '4122, Abrechnungsgebiet (07 = Diabetes)
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AdNam as string 'AdNam 130 '4602(1) Name bei Musterfrau
 AdOrt as string 'AdOrt 130 '4602(4) Ort bei Musterfrau
 AdPlz as string 'AdPlz 130 '4602(3) PLZ bei Musterfrau
 AdStr as string 'AdStr 130 '4602(2) Straße bei Musterfrau
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 altQuart as string 'altQuart 130 '
 AndÜw as string 'AndÜw 130 '4219, anderer Überweiser
 ausgst as date 'ausgst 7 '4102 ('ausgestellt am')
 BhFB as date 'BhFB 7 '4150
 BhFE as date 'BhFE 7 '4604, Behandlungsfall: Ende, bei Privatpatienten
 BhFE1 as date 'BhFE1 7 '4151
 BhFE2 as date 'BhFE2 7 '4152
 f4202 as string 'f4202 130 '4202
 f4206 as string 'f4206 130 '4206, ?
 f4210 as string 'f4210 130 '4210, unbekannt
 f4237 as string 'f4237 130 '4237, ? (nur bei Musterw)
 Fanf as date 'Fanf 7 'Fallanfang
 FID as long 'FID 3 'Primärschlüssel
 GebOr as string 'GebOr 130 '4121, Gebührenordnung (1 = BMÄ, 2)
 GOÄKatName as string 'GOÄKatName 130 '4580 (2): Privat-Abrechnungskatalog
 GOÄKatNr as string 'GOÄKatNr 130 '4580 (1): Katalog-Nummer
 IK as string 'IK 130 '4111 (auch patientenbezogen)
 KtrAbrB as string 'KtrAbrB 130 '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 KVKs as string 'KVKs 130 '4112
 KVKserg as string 'KVKserg 130 '4113
 lfdnr as long 'lfdnr 3 'laufende Fallnummer
 lVorl as date 'lVorl 7 '4109, letzte Vorlage
 Mahnfrist as string 'Mahnfrist 130 '4403, Mahnfrist bis
 Nachname as string 'Nachname 130 '3101
 Pat_ID as long 'Pat_ID 3 '3000
 PGeb as string 'PGeb 130 '4401, Praxisgebühr
 PGebErg as string 'PGebErg 130 '4402, ?
 QAnf as date 'QAnf 7 'Quartalsanfang
 QEnd as date 'QEnd 7 'Quartalsende
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 Quartal as string 'Quartal 130 '4101
 s8000 as string 's8000 130 '8000, ???
 s8100 as string 's8100 130 '8100, ???
 SchGr as string 'SchGr 130 '4239, Scheingruppe
 statBehTage as long 'statBehTage 3 '4238
 statKlasse as string 'statKlasse 130 '4236
 statNuller as string 'statNuller 130 '4216, nu bei Musterfrau 16 Nuller
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 TMFNr as string 'TMFNr 130 '4144 Fallnummer in Turbomed
 Übw as string 'Übw 130 '4218 oder 4219, je nachdem, was befüllt
 ÜbwV as string 'ÜbwV 130 '4218, überwiesen von
 ÜWNaN as string 'ÜWNaN 130 '4231(3): Nachname des Überweisers
 ÜWNNr as string 'ÜWNNr 130 '4231(4): KV-Nummer des Überweisers
 ÜwText as string 'ÜwText 130 '4209: Auftrags- / erläuternder Text zur Überweisung
 ÜWTit as string 'ÜWTit 130 '4231(3): Titel des Überweisers
 ÜWVor as string 'ÜWVor 130 '4231(2): Vorname des Überweisers
 ÜWVsw as string 'ÜWVsw 130 '4231(2b): Vorsatzwort des Überweisers
 ÜWZiel as string 'ÜWZiel 130 '4220
 VKNr as string 'VKNr 130 '4104
 Vorname as string 'Vorname 130 '3102
 Weiterbeh as string 'Weiterbeh 130 '4243, Weiterbehandelnder
end type

Public type au
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Beginn as string 'Beginn 130 '6285 1. Hälfte
 Ende as string 'Ende 130 '6285 2. Hälfte
 FID as long 'FID 3 'Fall-Bezug
 ICDs as string 'ICDs 130 '6286
 Pat_ID as long 'Pat_ID 3 '3000
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '6200 + 6201
end type

Public type briefe
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Art as string 'Art 130 '
 DokGroe as long 'DokGroe 3 'Größe der Datei
 FID as long 'FID 3 'Fall-Bezug
 Name as string 'Name 130 '
 Pat_ID as long 'Pat_ID 3 '
 Pfad as string 'Pfad 130 '
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Typ as string 'Typ 130 '
 ZeitPunkt as date 'ZeitPunkt 7 '
end type

Public type diagnosen
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 DiagAttr as string 'DiagAttr 130 '6006
 DiagDatum as date 'DiagDatum 7 '
 DiagSeite as string 'DiagSeite 130 '6004
 DiagSicherheit as string 'DiagSicherheit 130 '6003
 DiagText as string 'DiagText 130 '
 FID as long 'FID 3 'Fall-Bezug
 GesName as string 'GesName 130 '
 ICD as string 'ICD 130 '
 ID1 as long 'ID1 3 '
 obDauer as integer 'obDauer 11 'ob Dauerdiagnose
 Pat_id as long 'Pat_id 3 'Bezug auf Anamneseblattt
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type dokumente
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 DokArt as string 'DokArt 130 '
 DokGroe as long 'DokGroe 3 'Dokument-Größe
 DokName as string 'DokName 130 '
 DokPfad as string 'DokPfad 130 '
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 Quelldatum as date 'Quelldatum 7 'Datum, auf das sich das Dokument bezieht
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '
end type

Public type eintraege
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Art as string 'Art 130 '6330
 FID as long 'FID 3 'Fall-Bezug
 Inhalt as string 'Inhalt 130 '8480
 Pat_ID as long 'Pat_ID 3 '3000
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '
end type

Public type forminhaltform_abk
 Form_Abk as string 'Form_Abk 130 '
 Form_AbkVW as long 'Form_AbkVW 3 '
end type

Public type formulare
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 7 'Zeitpunkt der Aktualisierung
 Form_Abk as string 'Form_Abk 130 '
 FormBez as string 'FormBez 130 '
 FormID as long 'FormID 3 '
 FormVorl as string 'FormVorl 130 '
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type forminhkopf
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 FID as long 'FID 3 'Fall-Bezug
 FoID as long 'FoID 3 '
 Form_ID as long 'Form_ID 3 '
 Pat_ID as long 'Pat_ID 3 '
 Satzart as string 'Satzart 130 '8000
 Satzlänge as string 'Satzlänge 130 '8100
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '
end type

Public type forminhfeld
 FeldInhVW as long 'FeldInhVW 3 '
 FeldInh as string
 FeldNr as integer 'FeldNr 2 '
 FeldVW as long 'FeldVW 3 '
 Feld as string
 FoID as long 'FoID 3 '
 Nr as integer 'Nr 2 '
end type

Public type kheinweis
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Diagnose as string 'Diagnose 130 '6230
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '
 Ziel as string 'Ziel 130 '6291
end type

Public type lbanforderungen
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 AnfText as string 'AnfText 130 '6280
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 StByte as long 'StByte 3 'Statusbyte
 ZeitPunkt as date 'ZeitPunkt 7 '6200 + 6201
end type

Public type laborneu
 Abkü as string 'Abkü 130 '8410
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Einheit as string 'Einheit 130 '8421
 FertigStGrad as string 'FertigStGrad 130 '8401
 FID as long 'FID 3 'Fall-Bezug
 KommentarVW as long 'KommentarVW 3 '8480
 Kommentar as string
 LangtextVW as long 'LangtextVW 3 '8411
 Langtext as string
 Pat_ID as long 'Pat_ID 3 '3000
 Refnr as long 'Refnr 3 'Bezug auf LaborXUS
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Wert as string 'Wert 130 '8420
 ZeitPunkt as date 'ZeitPunkt 7 '
end type

Public type leistungen
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 f5002 as string 'f5002 130 '5002
 f5005 as string 'f5005 130 '5005
 f5006 as string 'f5006 130 '5006
 f5009 as string 'f5009 130 '5009
 f5015 as string 'f5015 130 '5015
 f5016 as string 'f5016 130 '5016
 f5021 as string 'f5021 130 '5021
 f5026 as string 'f5026 130 '5026
 FID as long 'FID 3 'Fall-Bezug
 Leistung as string 'Leistung 130 '5001
 Med as string 'Med 130 '5010
 Pat_ID as long 'Pat_ID 3 '3000
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '5000 + 6201
end type

Public type medplan
 ab as string 'ab 130 '
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 bBed as integer 'bBed 11 '
 Bemerkung as string 'Bemerkung 130 '
 Datum as date 'Datum 7 'Zeitpunkt aus dem Kopf des Medikamentenplans
 FeldNr as integer 'FeldNr 2 '
 FID as long 'FID 3 'Fall-Bezug
 MedAnfang as string 'MedAnfang 130 '
 Medikament as string 'Medikament 130 '
 mi as string 'mi 130 '
 mo as string 'mo 130 '
 MPNr as long 'MPNr 3 'Ordnungsziffer für Medikamentenplan
 nm as string 'nm 130 '
 Pat_ID as long 'Pat_ID 3 '3000
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 'Zeitpunkt, der Speicherung im Turbomed
 zn as string 'zn 130 '
end type

Public type rezepteintraege
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 FID as long 'FID 3 'Fall-Bezug
 Medikament as string 'Medikament 130 '3652(2), 6218(4)
 Pat_ID as long 'Pat_ID 3 '3000
 PZN as string 'PZN 130 '6210(2), 6218(3)
 QS as string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 130 'Quartal des Behandlungsfallbeginns
 Rezept as string 'Rezept 130 '6210, 3652(1), 6218(1)
 Rezeptklasse as string 'Rezeptklasse 130 '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 StByte as long 'StByte 3 'Statusbyte
 ZeitPunkt as date 'ZeitPunkt 7 '6200 + 6201
end type

Public type rr
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 RR as string 'RR 130 '6230
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ZeitPunkt as date 'ZeitPunkt 7 '6200 + 6201
end type

Public type kvnrue
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 7 'Zeit der Aktualisuerung aus der BDT-Datei
 KVNr as string 'KVNr 130 '
 lfdnr as long 'lfdnr 3 '
 Pat_ID as long 'Pat_ID 3 '
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type unbekannte_kennungen
 absPos as long 'absPos 3 '
 Kennung as string 'Kennung 130 '
 Pat_ID as long 'Pat_ID 3 'zugehöriger Patient für spätere Ermittlungen
 StByte as long 'StByte 3 '
end type

Public type laborxsaetze
 Arztname as string 'Arztname 130 '203 Arztname (Turbomed)
 Arztnr as string 'Arztnr 130 '201 Arztnummer (Turbomed)
 DatID as long 'DatID 3 'Bezug zu LaborEingelesen
 Erstellungsdatum as string 'Erstellungsdatum 130 '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge as string 'Gesamtlänge 130 '9202 Gesamtlänge des Datenpaketes (Turbomed)
 KBVPrüfnr as string 'KBVPrüfnr 130 '101 KBV-Prüfnummer (Turbomed)
 Kundenarztnr as string 'Kundenarztnr 130 '8312 Kundenarztnummer (Turbomed)
 Labor as string 'Labor 130 '8320 Labor
 OrtLabor as string 'OrtLabor 130 '8323 Ort der Laboradresse (Turbomed)
 OrtPraxis as string 'OrtPraxis 130 '216 Ort der Praxis (Turbomed)
 PLZLabor as string 'PLZLabor 130 '8322 PLZ der Laboradresse (Turbomed)
 PLZPraxis as string 'PLZPraxis 130 '215 PLZ der Praxis (Turbomed)
 Satzart as string 'Satzart 130 '8000 Satzart (Turbomed)
 SatzID as long 'SatzID 3 'zum Bezug für LaborUS
 Satzlänge as string 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 SatzlängeSchluss as string 'SatzlängeSchluss 130 '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 StraßeLabor as string 'StraßeLabor 130 '8321 Straße der Laboradresse (Turbomed)
 StraßePraxis as string 'StraßePraxis 130 '205 Straße der Praxis (Turbomed)
 VersionSatzb as string 'VersionSatzb 130 '9212 Version der Satzbeschreibung (Turbomed)
 Zeichensatz as string 'Zeichensatz 130 '9106 verwendeter Zeichensatz (Turbomed)
end type

Public type laborxeingel
 DatID as long 'DatID 3 'Bezug auf LaborEingelesen
 fertig as integer 'fertig 11 'ob Einlesen fertig
 Name as string 'Name 130 'Name der eingelesenen Labordatei ohne Endung
 Pfad as string 'Pfad 130 'Pfadname
 Zp as date 'Zp 7 'Einlesezeitpunkt
end type

Public type laborxus
 Abrechnungstyp as string 'Abrechnungstyp 130 '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 AfN as integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
 Auftragsnummer as string 'Auftragsnummer 130 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel as string 'Auftragsschlüssel 130 '8311 Anforderungsnr d Labors (Turbomed)
 AuftrHinw as string 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 BefArt as string 'BefArt 130 '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Berichtsdatum as string 'Berichtsdatum 130 '8302 Berichtsdatum
 DatID as long 'DatID 3 'Bezug auf LaborEingelesen
 Eingang as date 'Eingang 7 '8301 Eingangsdatum in Datumsform
 GebDat as string 'GebDat 130 '3103
 GebüOrd as string 'GebüOrd 130 '8403 Gebührenordnung (Turbomed)
 Geschlecht as string 'Geschlecht 130 '8407 Geschlecht (Turbomed)
 LWerte as string 'LWerte 130 'Laborwerte, die zur Zuordnung geführt haben
 Nachname as string 'Nachname 130 '3101
 NVorsatz as string 'NVorsatz 130 '3100
 Pat_id as long 'Pat_id 3 '
 Pat_idErwG as string 'Pat_idErwG 130 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB as string 'Pat_idErwGB 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL as string 'Pat_idErwGL 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idErwVN as string 'Pat_idErwVN 130 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwVNG as string 'Pat_idErwVNG 130 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idLaborNeu as string 'Pat_idLaborNeu 130 'Pat_ids von in Laborneu zuordnbaren Patienten
 Pat_idUrsp as string 'Pat_idUrsp 130 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Patienteninformation as string 'Patienteninformation 130 '8405 Patienteninformation (Turbomed)
 RefNr as long 'RefNr 3 'Bezug auf LaborWert
 Satzart as string 'Satzart 130 '8000 Satzart (Turbomed)
 SatzID as long 'SatzID 3 'Bezug auf LaborXSätze
 Satzlänge as string 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 Titel as string 'Titel 130 '3104
 verglichen as date 'verglichen 7 'Datum, zu dem Datensatz zuletzt verglichen wurde
 Vorname as string 'Vorname 130 '3102
 ZdiP as long 'ZdiP 3 'Zahl der infragekommenden Patienten
 ZdüP as integer 'ZdüP 2 'Zahl der verglichenen Parameter
 ZeitpunktLaborneu as date 'ZeitpunktLaborneu 7 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
end type

Public type laborxbakt
 AbnDat as date 'AbnDat 7 '8432 Abnahmedatum (Turbomed)
 Erklärung as string 'Erklärung 130 '
 Keimzahl as string 'Keimzahl 130 '
 Kommentar as string 'Kommentar 130 '8480 Ergebnistest (Turbomed)
 KuQu as string 'KuQu 130 '8428 Probenmaterial-Ident (Turbomed)
 QSpez as string 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 Quelle as string 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 RefNr as long 'RefNr 3 '
 Verf as string 'Verf 130 '
end type

Public type laborxwert
 Abkü as string 'Abkü 130 '8410 Test-Ident  (Turbomed)
 AbnDat as date 'AbnDat 7 '8432 Abnahmedatum (Turbomed)
 AuftrHinw as string 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 Einheit as string 'Einheit 130 '8421 Einheit (Turbomed)
 Erklärung as string 'Erklärung 130 '8470 Testbezogene Hinweise (Turbomed)
 Grenzwerti as string 'Grenzwerti 130 '8422 Grenzwertindikator (Turbomed)
 Kommentar as string 'Kommentar 130 '8480 Ergebnistext (Turbomed)
 Langname as string 'Langname 130 '8411 Testbezeichnung (Turbomed)
 Normbereich as string 'Normbereich 130 '8460 Normalwert-Text (Turbomed)
 NormO as string 'NormO 130 '8461 Normobergrenze
 NormU as string 'NormU 130 '8461 Normuntergrenze
 QSpez as string 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 Quelle as string 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 RefNr as long 'RefNr 3 'Bezug auf LaborUS
 Teststatus as string 'Teststatus 130 '8418 Teststatus (Turbomed)
 Wert as string 'Wert 130 '8420 Ergebniswert (Turbomed)
end type

Public type laborxleist
 Abkü as string 'Abkü 130 '8410 Test-Ident (Turbomed)
 Anzahl as string 'Anzahl 130 '5005
 EBM as string 'EBM 130 '5001 GNR (Turbomed)
 goä as string 'goä 130 '8406
 RefNr as long 'RefNr 3 'Bezug auf LaborUS
 Verf as string 'Verf 130 '8434
end type

Public type anamnesebogen
 AktZeit as date 'AktZeit 7 'Aktualisierungszeit
 Albumin_zuletzt as string 'Albumin_zuletzt 130 ', letztes Albumin:
 Alkohol as string 'Alkohol 130 '^:
 Ameisen_Ausmaß as string 'Ameisen_Ausmaß 130 '<, Ausmaß:
 Ameisenlaufen as string 'Ameisenlaufen 130 '^:
 Amputation as string 'Amputation 130 ',:
 andere_Nierenerkrankung as string 'andere_Nierenerkrankung 130 ', andere Nierenerkrankung:
 Angina_pectoris as string 'Angina_pectoris 130 ',:
 Anrede as string 'Anrede 130 '
 ASR as string 'ASR 130 ',:
 Aufschreiben as string 'Aufschreiben 130 '<, Dokumentation:
 Augensp_Befund as string 'Augensp_Befund 130 '<, Befund:
 Augensp_zuletzt as string 'Augensp_zuletzt 130 '^Letzte Augenspiegelung:
 Bauch as string 'Bauch 130 ', Abdomen:
 BDselbst as string 'BDselbst 130 '^Blutdruckselbstmessung:
 Beinbefund as string 'Beinbefund 130 '^:
 Beindurchblutungsstörung as string 'Beindurchblutungsstörung 130 '^:
 BeinödVen as string 'BeinödVen 130 ', Beinödeme/ Venenkrankheiten:
 Bewegungseinschränkungen as string 'Bewegungseinschränkungen 130 ',:
 Bewußtlos_pa as string 'Bewußtlos_pa 130 '<, bewußtlos deshalb:
 BHD_beh_mit as string 'BHD_beh_mit 130 '<, behandelt mit:
 BHD_seit as string 'BHD_seit 130 '<seit:
 Blutdruckwerte as string 'Blutdruckwerte 130 '^Blutdruckwerte:
 Bluthochdruck as string 'Bluthochdruck 130 '^Bluthochdruck:
 Broteinheiten_abends as string 'Broteinheiten_abends 130 '<, abends
 Broteinheiten_früh as string 'Broteinheiten_früh 130 '<, früh
 Broteinheiten_gesamt as string 'Broteinheiten_gesamt 130 '^Broteinheiten:gesamt
 Broteinheiten_mittags as string 'Broteinheiten_mittags 130 '<, mittags
 Broteinheiten_nachmittags as string 'Broteinheiten_nachmittags 130 '<, nachmittags
 Broteinheiten_nachts as string 'Broteinheiten_nachts 130 '<, nachts
 Broteinheiten_ZM_früh as string 'Broteinheiten_ZM_früh 130 '<, Zwischenmahlzeit vormittags
 Bypass_kardial as integer 'Bypass_kardial 11 ',:
 Bypaß_peripher as integer 'Bypaß_peripher 11 ',:
 Bypass_wann as string 'Bypass_wann 130 '<, wann:
 BZgr300_pM as string 'BZgr300_pM 130 ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 BZMessungen_p_W_nachts as string 'BZMessungen_p_W_nachts 130 '<, nachts:
 BZMessungen_pW as string 'BZMessungen_pW 130 '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE as string 'BZMessungen_pW_ndE 130 '<, davon nach dem Essen:
 BZMessungen_selbst as string 'BZMessungen_selbst 130 '^Blutzuckermessung:Selbstmessung?
 BZWerte_n_d_Essen as string 'BZWerte_n_d_Essen 130 '<, nach dem Essen:
 BZWerte_v_d_Essen as string 'BZWerte_v_d_Essen 130 '^Blutzuckerwerte vor dem Essen:
 Carotiden as string 'Carotiden 130 ', Halsschlagadern:
 Diabet_Nierenschaden as string 'Diabet_Nierenschaden 130 '^Diabetischer Nierenschaden:
 Diabetes_seit as string 'Diabetes_seit 130 '<seit
 DiabetesMedikament_1 as string 'DiabetesMedikament_1 130 '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge as string 'DiabetesMedikament_1_Menge 130 '<
 DiabetesMedikament_2 as string 'DiabetesMedikament_2 130 '<,
 DiabetesMedikament_2_Menge as string 'DiabetesMedikament_2_Menge 130 '<
 DiabetesMedikament_3 as string 'DiabetesMedikament_3 130 '<,
 DiabetesMedikament_3_Menge as string 'DiabetesMedikament_3_Menge 130 '<
 DiabetesMedikament_4 as string 'DiabetesMedikament_4 130 '<,
 DiabetesMedikament_4_Menge as string 'DiabetesMedikament_4_Menge 130 '<,
 Diabetestyp as string 'Diabetestyp 130 '^Diabetes Typ
 Diagnosen as string 'Diagnosen 130 '
 Dialyse as integer 'Dialyse 11 ',:
 Dialyse_seit as string 'Dialyse_seit 130 '<seit
 DMP as string 'DMP 130 'ob Pat. bei HA im DMP
 DMPhier as date 'DMPhier 7 'ob Pat hier im DMP
 DMSchL as integer 'DMSchL 2 'Zahl der abgerechneten DMP-Schulungen hier
 DMSchulz as integer 'DMSchulz 2 'Zahl der DMP-Schulungen hier
 Druckstellen as string 'Druckstellen 130 ',:
 Einlagen as string 'Einlagen 130 ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Entleerungsstörungen_Harnblase as string 'Entleerungsstörungen_Harnblase 130 ',:
 Entleerungsstörungen_Magen as string 'Entleerungsstörungen_Magen 130 '^:
 erhöht as string 'erhöht 130 '<, Befund:
 Essenszeit_abends as string 'Essenszeit_abends 130 '<, abends
 Essenszeit_früh as string 'Essenszeit_früh 130 '^Essenszeiten:früh
 Essenszeit_mittags as string 'Essenszeit_mittags 130 '<, mittags
 Essenszeit_nachmittags as string 'Essenszeit_nachmittags 130 '<, nachmittags
 Essenszeit_spät as string 'Essenszeit_spät 130 '<, spät
 Essenszeit_vormittags as string 'Essenszeit_vormittags 130 '<, vormittags
 Familienanamnese as string 'Familienanamnese 130 '^:
 Folgeerkrankungen_Haut as string 'Folgeerkrankungen_Haut 130 '^:
 Fremde_Hilfe_pa as string 'Fremde_Hilfe_pa 130 '<, fremde Hilfe deshalb nötig:
 Fußpflege as string 'Fußpflege 130 '^:
 GebDat as date 'GebDat 7 ', geb.
 gemessen_am as date 'gemessen_am 7 '<, gemessen
 Gerät as string 'Gerät 130 '<:
 Geschwür as string 'Geschwür 130 ',:
 Gewicht as double 'Gewicht 5 ',:
 Größe as double 'Größe 5 '^:
 Grund_für_Vorstellung as string 'Grund_für_Vorstellung 130 '^:
 HANr as string 'HANr 130 'mit "/"
 HANr2 as string 'HANr2 130 'mit "/"
 Hausarzt as string 'Hausarzt 130 'Hausarzt laut Anamnesebogen
 Herz as string 'Herz 130 '^:
 Herzinfarkt as string 'Herzinfarkt 130 ',:
 Herzinfarkt_wann as string 'Herzinfarkt_wann 130 '<, wann:
 Herzkrankheit as string 'Herzkrankheit 130 '^Herzkrankheit:
 Herzkrankheit_Beschreibung as string 'Herzkrankheit_Beschreibung 130 ', Beschreibung:
 Herzschwäche as string 'Herzschwäche 130 ',:
 Hirndurchblutungsstörung as string 'Hirndurchblutungsstörung 130 '^:
 Hyperkeratosen as string 'Hyperkeratosen 130 ',:
 Insulin_seit as string 'Insulin_seit 130 ', Insulin seit
 Insulinpumpe as integer 'Insulinpumpe 11 '^:
 Insulinpumpe_Marke as string 'Insulinpumpe_Marke 130 '<, Marke:
 Insulinpumpe_seit as string 'Insulinpumpe_seit 130 '<seit
 Jahr_letzte_Diabetesschulung as string 'Jahr_letzte_Diabetesschulung 130 '^Letzte Diabetesschulung:
 Kalt_Warm as string 'Kalt_Warm 130 ', Kalt-Warm-Diskrimination:
 Keto_pa as string 'Keto_pa 130 '^Zahl der Ketoazidosen pro Jahr:
 Kraft_Knie as string 'Kraft_Knie 130 '<, Knie:
 Kraft_Zehenbeuger as string 'Kraft_Zehenbeuger 130 '<, Zehenbeuger:
 Kraft_Zehenheber as string 'Kraft_Zehenheber 130 '^Kraft:Zehenheber
 letzte_Änderung as date 'letzte_Änderung 7 'Datum der letzten Änderung
 letztes_HbA1c as string 'letztes_HbA1c 130 '^Letztes HbA1c:
 Liphypertrophien_Abdomen as string 'Liphypertrophien_Abdomen 130 '^Liphypertrophien:Abdomen
 Liphypertrophien_Arme as string 'Liphypertrophien_Arme 130 '<, Arme:
 Liphypertrophien_Beine as string 'Liphypertrophien_Beine 130 '<, Beine:
 LK as string 'LK 130 ', Lymphknoten:
 Lunge as string 'Lunge 130 ',:
 Monofilamenttest as string 'Monofilamenttest 130 ',:
 Mundhöhle as string 'Mundhöhle 130 ',:
 Nachname as string 'Nachname 130 '-
 Netzhaut_gelasert as string 'Netzhaut_gelasert 130 ', Netzhaut schon gelasert:
 Neue_Fußkomplikationen as string 'Neue_Fußkomplikationen 130 '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Neuro_sonst as string 'Neuro_sonst 130 '^Sonstige neurologische Befunde:
 NL as string 'NL 130 ', Nierenlager:
 NNH as string 'NNH 130 ', Nasennebenhöhlen:
 NVorsatz as string 'NVorsatz 130 '
 ob as integer 'ob 11 'für verschiedene Aktionen
 obAn1eing as integer 'obAn1eing 11 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing as integer 'obAn2eing 11 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing as integer 'obAnAeing 11 'ob Anamneseblatt allgemein eingegeben wurde
 obBZausgew as integer 'obBZausgew 11 'ob Blutzuckergerät ausgewechselt
 obCheck as integer 'obCheck 11 'ob Checkliste vorliegt
 obDMPaufgekl as string 'obDMPaufgekl 130 'ob über DMP aufgeklärt
 Oberflächensensibilität as string 'Oberflächensensibilität 130 '^:
 obMBlAusgeh as integer 'obMBlAusgeh 11 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz as integer 'obMedNetz 11 'ob von Med. Netz geschickt
 obOSaufgek as integer 'obOSaufgek 11 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek as integer 'obPodAufgek 11 'ob über Podologie aufgeklärt
 obSchulaufgek as string 'obSchulaufgek 130 'ob über Schulung aufgeklärt
 Ort_Schulung as string 'Ort_Schulung 130 '<in
 Pat_id as long 'Pat_id 3 '
 pAVK_Beschreibung as string 'pAVK_Beschreibung 130 ', Beschreibung der Beinbeschwerden:
 Podologie as string 'Podologie 130 ',:
 Prim as long 'Prim 3 'Primärschlüssel
 PSR as string 'PSR 130 ',:
 PTCA_oder_Stent as string 'PTCA_oder_Stent 130 ',:
 Puls_Adp as string 'Puls_Adp 130 '<,Fußrücken:
 Puls_Atp as string 'Puls_Atp 130 '<,Innenknöchel:
 Puls_Kniekehle as string 'Puls_Kniekehle 130 '<,Kniekehle:
 Puls_Leiste as string 'Puls_Leiste 130 '^Pulse:Leiste
 QS as string 'QS 130 'Quartal sortiert von vorgestellt
 QT as string 'QT 130 'Quartal sortiert von vorgestellt
 RR as string 'RR 130 '^Blutdruck:
 RRSchulz as integer 'RRSchulz 2 'Zahl der Hypertonie-Schulungen hier
 RRTurboMed as string 'RRTurboMed 130 '
 Schaufensterkrankheit as string 'Schaufensterkrankheit 130 ',:
 Schlaganfall as string 'Schlaganfall 130 ',:
 Schulung as string 'Schulung 130 'ob Schulungsbedarf
 Schwanger as string 'Schwanger 130 '^Aktuelle Schwangerschaft:
 Schwanger_seit as string 'Schwanger_seit 130 '<, seit:
 Schwindel_Aufstehen as string 'Schwindel_Aufstehen 130 ',:
 SD as string 'SD 130 ', Schilddrüse:
 Sehminderung_unbehebbar as string 'Sehminderung_unbehebbar 130 ', mit Brille nicht behebbare Sehminderung:
 Sexualstörung as string 'Sexualstörung 130 '^:
 Sexualstörung_seit as string 'Sexualstörung_seit 130 '<seit
 Spritz_Eß_Abstand_abends as string 'Spritz_Eß_Abstand_abends 130 '<, abends
 Spritz_Eß_Abstand_früh as string 'Spritz_Eß_Abstand_früh 130 '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags as string 'Spritz_Eß_Abstand_mittags 130 '<, mittags
 Spritzstelle_abends as string 'Spritzstelle_abends 130 '<, abends
 Spritzstelle_früh as string 'Spritzstelle_früh 130 '^Spritzstellen:früh
 Spritzstelle_mittags as string 'Spritzstelle_mittags 130 '<, mittags
 Spritzstelle_nachts as string 'Spritzstelle_nachts 130 '<, nachts
 Tabak as string 'Tabak 130 ',:
 Tabletten_seit as string 'Tabletten_seit 130 ', Tabletten seit
 Tendenz as string 'Tendenz 130 '<, Tendenz
 Ther1 as string 'Ther1 130 'Diät, OAD, CT, ICT, CSII
 TherAkt as string 'TherAkt 130 'Diät, OAD, CT, ICT, CSII
 Titel as string 'Titel 130 '
 Tkz as integer 'Tkz 11 'Tod-Kennzeichen
 Ulcera as string 'Ulcera 130 ',:
 Unterzucker_pM as string 'Unterzucker_pM 130 '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig as string 'UZ_rechtzeitig 130 '<, rechtzeitig bemerkt:
 UZ_Tageszeit as string 'UZ_Tageszeit 130 '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Verformungen as string 'Verformungen 130 ',:
 Verformungen_Beschreibung as string 'Verformungen_Beschreibung 130 '<Beschreibung:
 Versicherung as string 'Versicherung 130 '
 Versicherungsart as string 'Versicherungsart 130 '
 Vibration_Großzehe as string 'Vibration_Großzehe 130 '<, Großzehe:
 Vibration_IK as string 'Vibration_IK 130 ', Vibrationsempfinden Innenknöchel:
 Vorgestellt as date 'Vorgestellt 7 'Erstvorstellung
 vorherige_Werte as string 'vorherige_Werte 130 '<, vorher:
 Vorname as string 'Vorname 130 '
 Weitere_Anamnese as string 'Weitere_Anamnese 130 '^:
 Weitere_Befunde as string 'Weitere_Befunde 130 ', weitere Befunde:
 Weitere_Medikation as string 'Weitere_Medikation 130 '^:
 WS as string 'WS 130 ', Wirbelsäule:
 Zähne as string 'Zähne 130 ',:
end type

Public rNa() as namen
Public rFa() as faelle
Public rAu() as au
Public rBr() as briefe
Public rDi() as diagnosen
Public rDo() as dokumente
Public rEi() as eintraege
Public rFi() as forminhaltform_abk
Public rFo() as formulare
Public rFr() as forminhkopf
Public rFm() as forminhfeld
Public rKh() as kheinweis
Public rLb() as lbanforderungen
Public rLa() as laborneu
Public rLe() as leistungen
Public rMe() as medplan
Public rRe() as rezepteintraege
Public rRr() as rr
Public rKv() as kvnrue
Public rUn() as unbekannte_kennungen
Public rLs() as laborxsaetze
Public rLg() as laborxeingel
Public rLu() as laborxus
Public rLo() as laborxbakt
Public rLw() as laborxwert
Public rLL() as laborxleist
Public rAna() as anamnesebogen

Public Function Tinit()
 static wdh%
 ReDim rAna(0)
 ReDim rNa(0)
 ReDim rFa(0)
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 if wdh = 0 then reDim rFi(0)
 if wdh = 0 then reDim rFo(0)
 ReDim rFr(0)
 ReDim rFm(0)
 ReDim rKh(0)
 ReDim rLb(0)
 ReDim rLa(0)
 ReDim rLe(0)
 ReDim rMe(0)
 ReDim rRe(0)
 ReDim rRr(0)
 ReDim rKv(0)
 ReDim rUn(0)
 wdh = -1
End Function ' Tinit

Public Function LabInit()
 static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 wdh = -1
End Function ' Tinit

Public Function doLösch(frm as lese, Tbl$)
 Set rs = dbcn.Execute("select count(*) as ct from " + kla + Tbl + klz)
 frm.Ausgabe = "Lösche: " & kla & Tbl & klz & " (" & rs!ct & " Datensätze)" & vbCrLf & altAusgabe
 altausgabe = frm.Ausgabe
 sql = sqldeletefrom & kla & Tbl & klz
 Call dbcn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doLösch

Public Function AllesLösch(frm as lese)
 Dim ct&, rs as new ADODB.recordset
 on error goto fehler
 call ForeignNo0
 call ForeignNo1
 call doLösch(frm, "unbekannte kennungen")
 call doLösch(frm, "kvnrue")
 call doLösch(frm, "rr")
 call doLösch(frm, "rezepteintraege")
 call doLösch(frm, "medplan")
 call doLösch(frm, "leistungen")
 call doLösch(frm, "laborneu")
 call doLösch(frm, "lbanforderungen")
 call doLösch(frm, "kheinweis")
 call doLösch(frm, "forminhfeld")
 call doLösch(frm, "forminhkopf")
 call doLösch(frm, "formulare")
 call doLösch(frm, "forminhaltform_abk")
 call doLösch(frm, "eintraege")
 call doLösch(frm, "dokumente")
 call doLösch(frm, "diagnosen")
 call doLösch(frm, "briefe")
 call doLösch(frm, "au")
 call doLösch(frm, "faelle")
 call doLösch(frm, "namen")
 call ForeignYes0
 call ForeignYes1
 exit function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' AllesLösch

Public Function LabLösch(frm as lese)
 Dim ct&, rs as new ADODB.recordset
 on error goto fehler
 call ForeignNo0
 call ForeignNo1
 call doLösch(frm, "laborxleist")
 call doLösch(frm, "laborxwert")
 call doLösch(frm, "laborxbakt")
 call doLösch(frm, "laborxus")
 call doLösch(frm, "laborxeingel")
 call doLösch(frm, "laborxsaetze")
 call ForeignYes0
 call ForeignYes1
 exit function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' LabLösch


Public Function namenSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "namen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "namen" & klz & " (absPos,AktZeit,An1da," & _
     "An2da,AnAllgda,AufnDat,Cave,Checkda,DienstTel,DMTypaD,Email,GebDat,Geschlecht," & _
     "intZOGP,KVKStatus,KVNr,lfdnr,Nachname,Notiz,NVorsatz,Ort,Pat_ID,Plz," & _
     "PrivatFax,PrivatMobil,PrivatTel,PrivatTel_2,StByte,Straße,Titel,Versichertennummer,Verwandt,Vorname," & _
     "Weggeldzone,zubenach) values
 For i = 0 to ubound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rNa(i).absPos & "," & datform( 0 ) & "," & cstr(cint(rNa(i).An1da)) & "," & cstr(cint(rNa(i).An2da)) & "," & cstr(cint(rNa(i).AnAllgda)) & "," & datform( _
   rNa(i).AufnDat) & ",'" & rNa(i).Cave & "'," & cstr(cint(rNa(i).Checkda)) & ",'" & rNa(i).DienstTel & "','" & rNa(i).DMTypaD & "','" &  _
   rNa(i).Email & "'," & datform(rNa(i).GebDat) & ",'" & rNa(i).Geschlecht & "','" & rNa(i).intZOGP & "','" & rNa(i).KVKStatus & "','" &  _
   rNa(i).KVNr & "'," & rNa(i).lfdnr & ",'" & rNa(i).Nachname & "','" & rNa(i).Notiz & "','" & rNa(i).NVorsatz & "','" & rNa(i).Ort & "'," &  _
   rNa(i).Pat_ID & ",'" & rNa(i).Plz & "','" & rNa(i).PrivatFax & "','" & rNa(i).PrivatMobil & "','" & rNa(i).PrivatTel & "','" & rNa(i).PrivatTel_2 & "'," &  _
   rNa(i).StByte & ",'" & rNa(i).Straße & "','" & rNa(i).Titel & "','" & rNa(i).Versichertennummer & "','" & rNa(i).Verwandt & "','" &  _
   rNa(i).Vorname & "','" & rNa(i).Weggeldzone & "','" & rNa(i).zubenach & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Cave":  If SpMod(rNa(i).Cave, "namen", rsc) Then Exit Do
   Case "DienstTel":  If SpMod(rNa(i).DienstTel, "namen", rsc) Then Exit Do
   Case "DMTypaD":  If SpMod(rNa(i).DMTypaD, "namen", rsc) Then Exit Do
   Case "Email":  If SpMod(rNa(i).Email, "namen", rsc) Then Exit Do
   Case "Geschlecht":  If SpMod(rNa(i).Geschlecht, "namen", rsc) Then Exit Do
   Case "intZOGP":  If SpMod(rNa(i).intZOGP, "namen", rsc) Then Exit Do
   Case "KVKStatus":  If SpMod(rNa(i).KVKStatus, "namen", rsc) Then Exit Do
   Case "KVNr":  If SpMod(rNa(i).KVNr, "namen", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rNa(i).Nachname, "namen", rsc) Then Exit Do
   Case "Notiz":  If SpMod(rNa(i).Notiz, "namen", rsc) Then Exit Do
   Case "NVorsatz":  If SpMod(rNa(i).NVorsatz, "namen", rsc) Then Exit Do
   Case "Ort":  If SpMod(rNa(i).Ort, "namen", rsc) Then Exit Do
   Case "Plz":  If SpMod(rNa(i).Plz, "namen", rsc) Then Exit Do
   Case "PrivatFax":  If SpMod(rNa(i).PrivatFax, "namen", rsc) Then Exit Do
   Case "PrivatMobil":  If SpMod(rNa(i).PrivatMobil, "namen", rsc) Then Exit Do
   Case "PrivatTel":  If SpMod(rNa(i).PrivatTel, "namen", rsc) Then Exit Do
   Case "PrivatTel_2":  If SpMod(rNa(i).PrivatTel_2, "namen", rsc) Then Exit Do
   Case "Straße":  If SpMod(rNa(i).Straße, "namen", rsc) Then Exit Do
   Case "Titel":  If SpMod(rNa(i).Titel, "namen", rsc) Then Exit Do
   Case "Versichertennummer":  If SpMod(rNa(i).Versichertennummer, "namen", rsc) Then Exit Do
   Case "Verwandt":  If SpMod(rNa(i).Verwandt, "namen", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rNa(i).Vorname, "namen", rsc) Then Exit Do
   Case "Weggeldzone":  If SpMod(rNa(i).Weggeldzone, "namen", rsc) Then Exit Do
   Case "zubenach":  If SpMod(rNa(i).zubenach, "namen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
ElseIf Err.Number = -2147217871 or err.number = -2147217859 Then
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "faelle" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "faelle" & klz & " (AbrAr,AbrGb,absPos," & _
     "AdNam,AdOrt,AdPlz,AdStr,AktZeit,altQuart,AndÜw,ausgst,BhFB,BhFE," & _
     "BhFE1,BhFE2,f4202,f4206,f4210,f4237,Fanf,GebOr,GOÄKatName,GOÄKatNr," & _
     "IK,KtrAbrB,KVKs,KVKserg,lfdnr,lVorl,Mahnfrist,Nachname,Pat_ID,PGeb," & _
     "PGebErg,QAnf,QEnd,QS,QT,Quartal,s8000,s8100,SchGr,statBehTage," & _
     "statKlasse,statNuller,StByte,TMFNr,Übw,ÜbwV,ÜWNaN,ÜWNNr,ÜwText,ÜWTit," & _
     "ÜWVor,ÜWVsw,ÜWZiel,VKNr,Vorname,Weiterbeh) values
 For i = 1 to ubound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rFa(i).AbrAr & "','" & rFa(i).AbrGb & "'," & rFa(i).absPos & ",'" & rFa(i).AdNam & "','" & rFa(i).AdOrt & "','" & rFa(i).AdPlz & "','" &  _
   rFa(i).AdStr & "'," & datform(rFa(i).AktZeit) & ",'" & rFa(i).altQuart & "','" & rFa(i).AndÜw & "'," & datform(rFa(i).ausgst) & "," & datform( _
   rFa(i).BhFB) & "," & datform(rFa(i).BhFE) & "," & datform(rFa(i).BhFE1) & "," & datform(rFa(i).BhFE2) & ",'" & rFa(i).f4202 & "','" & rFa(i).f4206 & "','" &  _
   rFa(i).f4210 & "','" & rFa(i).f4237 & "'," & datform(rFa(i).Fanf) & ",'" & rFa(i).GebOr & "','" & rFa(i).GOÄKatName & "','" &  _
   rFa(i).GOÄKatNr & "','" & rFa(i).IK & "','" & rFa(i).KtrAbrB & "','" & rFa(i).KVKs & "','" & rFa(i).KVKserg & "'," & rFa(i).lfdnr & "," & datform( _
   rFa(i).lVorl) & ",'" & rFa(i).Mahnfrist & "','" & rFa(i).Nachname & "'," & rFa(i).Pat_ID & ",'" & rFa(i).PGeb & "','" & rFa(i).PGebErg & "'," & datform( _
   rFa(i).QAnf) & "," & datform(rFa(i).QEnd) & ",'" & rFa(i).QS & "','" & rFa(i).QT & "','" & rFa(i).Quartal & "','" & rFa(i).s8000 & "','" &  _
   rFa(i).s8100 & "','" & rFa(i).SchGr & "'," & rFa(i).statBehTage & ",'" & rFa(i).statKlasse & "','" & rFa(i).statNuller & "'," & rFa(i).StByte & ",'" &  _
   rFa(i).TMFNr & "','" & rFa(i).Übw & "','" & rFa(i).ÜbwV & "','" & rFa(i).ÜWNaN & "','" & rFa(i).ÜWNNr & "','" & rFa(i).ÜwText & "','" &  _
   rFa(i).ÜWTit & "','" & rFa(i).ÜWVor & "','" & rFa(i).ÜWVsw & "','" & rFa(i).ÜWZiel & "','" & rFa(i).VKNr & "','" & rFa(i).Vorname & "','" &  _
   rFa(i).Weiterbeh & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_ID & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & datform(rFa(i).BhFB) & " and bhfe1 = " & datform(rFa(i).BhFE1) & " and ausgst = " & datform(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgabe = "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & rs!FID & vbCrLf & Lese.Ausgabe
    altAusgabe = Lese.ausgabe
    Dim jjj&
    For jjj = 1 To UBound(rAu)
     If rAu(jjj).FID = rFa(i).FID Then
      rAu(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rBr)
     If rBr(jjj).FID = rFa(i).FID Then
      rBr(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rDi)
     If rDi(jjj).FID = rFa(i).FID Then
      rDi(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rDo)
     If rDo(jjj).FID = rFa(i).FID Then
      rDo(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rEi)
     If rEi(jjj).FID = rFa(i).FID Then
      rEi(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rFr)
     If rFr(jjj).FID = rFa(i).FID Then
      rFr(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rKh)
     If rKh(jjj).FID = rFa(i).FID Then
      rKh(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLb)
     If rLb(jjj).FID = rFa(i).FID Then
      rLb(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLa)
     If rLa(jjj).FID = rFa(i).FID Then
      rLa(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLe)
     If rLe(jjj).FID = rFa(i).FID Then
      rLe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rMe)
     If rMe(jjj).FID = rFa(i).FID Then
      rMe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rRe)
     If rRe(jjj).FID = rFa(i).FID Then
      rRe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rRr)
     If rRr(jjj).FID = rFa(i).FID Then
      rRr(jjj).FID = rs!FID
     End If
    Next jjj
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AbrAr":  If SpMod(rFa(i).AbrAr, "faelle", rsc) Then Exit Do
   Case "AbrGb":  If SpMod(rFa(i).AbrGb, "faelle", rsc) Then Exit Do
   Case "AdNam":  If SpMod(rFa(i).AdNam, "faelle", rsc) Then Exit Do
   Case "AdOrt":  If SpMod(rFa(i).AdOrt, "faelle", rsc) Then Exit Do
   Case "AdPlz":  If SpMod(rFa(i).AdPlz, "faelle", rsc) Then Exit Do
   Case "AdStr":  If SpMod(rFa(i).AdStr, "faelle", rsc) Then Exit Do
   Case "altQuart":  If SpMod(rFa(i).altQuart, "faelle", rsc) Then Exit Do
   Case "AndÜw":  If SpMod(rFa(i).AndÜw, "faelle", rsc) Then Exit Do
   Case "f4202":  If SpMod(rFa(i).f4202, "faelle", rsc) Then Exit Do
   Case "f4206":  If SpMod(rFa(i).f4206, "faelle", rsc) Then Exit Do
   Case "f4210":  If SpMod(rFa(i).f4210, "faelle", rsc) Then Exit Do
   Case "f4237":  If SpMod(rFa(i).f4237, "faelle", rsc) Then Exit Do
   Case "GebOr":  If SpMod(rFa(i).GebOr, "faelle", rsc) Then Exit Do
   Case "GOÄKatName":  If SpMod(rFa(i).GOÄKatName, "faelle", rsc) Then Exit Do
   Case "GOÄKatNr":  If SpMod(rFa(i).GOÄKatNr, "faelle", rsc) Then Exit Do
   Case "IK":  If SpMod(rFa(i).IK, "faelle", rsc) Then Exit Do
   Case "KtrAbrB":  If SpMod(rFa(i).KtrAbrB, "faelle", rsc) Then Exit Do
   Case "KVKs":  If SpMod(rFa(i).KVKs, "faelle", rsc) Then Exit Do
   Case "KVKserg":  If SpMod(rFa(i).KVKserg, "faelle", rsc) Then Exit Do
   Case "Mahnfrist":  If SpMod(rFa(i).Mahnfrist, "faelle", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rFa(i).Nachname, "faelle", rsc) Then Exit Do
   Case "PGeb":  If SpMod(rFa(i).PGeb, "faelle", rsc) Then Exit Do
   Case "PGebErg":  If SpMod(rFa(i).PGebErg, "faelle", rsc) Then Exit Do
   Case "QS":  If SpMod(rFa(i).QS, "faelle", rsc) Then Exit Do
   Case "QT":  If SpMod(rFa(i).QT, "faelle", rsc) Then Exit Do
   Case "Quartal":  If SpMod(rFa(i).Quartal, "faelle", rsc) Then Exit Do
   Case "s8000":  If SpMod(rFa(i).s8000, "faelle", rsc) Then Exit Do
   Case "s8100":  If SpMod(rFa(i).s8100, "faelle", rsc) Then Exit Do
   Case "SchGr":  If SpMod(rFa(i).SchGr, "faelle", rsc) Then Exit Do
   Case "statKlasse":  If SpMod(rFa(i).statKlasse, "faelle", rsc) Then Exit Do
   Case "statNuller":  If SpMod(rFa(i).statNuller, "faelle", rsc) Then Exit Do
   Case "TMFNr":  If SpMod(rFa(i).TMFNr, "faelle", rsc) Then Exit Do
   Case "Übw":  If SpMod(rFa(i).Übw, "faelle", rsc) Then Exit Do
   Case "ÜbwV":  If SpMod(rFa(i).ÜbwV, "faelle", rsc) Then Exit Do
   Case "ÜWNaN":  If SpMod(rFa(i).ÜWNaN, "faelle", rsc) Then Exit Do
   Case "ÜWNNr":  If SpMod(rFa(i).ÜWNNr, "faelle", rsc) Then Exit Do
   Case "ÜwText":  If SpMod(rFa(i).ÜwText, "faelle", rsc) Then Exit Do
   Case "ÜWTit":  If SpMod(rFa(i).ÜWTit, "faelle", rsc) Then Exit Do
   Case "ÜWVor":  If SpMod(rFa(i).ÜWVor, "faelle", rsc) Then Exit Do
   Case "ÜWVsw":  If SpMod(rFa(i).ÜWVsw, "faelle", rsc) Then Exit Do
   Case "ÜWZiel":  If SpMod(rFa(i).ÜWZiel, "faelle", rsc) Then Exit Do
   Case "VKNr":  If SpMod(rFa(i).VKNr, "faelle", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rFa(i).Vorname, "faelle", rsc) Then Exit Do
   Case "Weiterbeh":  If SpMod(rFa(i).Weiterbeh, "faelle", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "au" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "au" & klz & " (absPos,AktZeit,Beginn," & _
     "Ende,FID,ICDs,Pat_ID,StByte,ZeitPunkt) values
 For i = 1 to ubound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rAu(i).absPos & "," & datform(rAu(i).AktZeit) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "'," & rAu(i).FID & ",'" & rAu(i).ICDs & "'," &  _
   rAu(i).Pat_ID & "," & rAu(i).StByte & "," & datform(rAu(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Beginn":  If SpMod(rAu(i).Beginn, "au", rsc) Then Exit Do
   Case "Ende":  If SpMod(rAu(i).Ende, "au", rsc) Then Exit Do
   Case "ICDs":  If SpMod(rAu(i).ICDs, "au", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "briefe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "briefe" & klz & " (absPos,AktZeit,Art," & _
     "DokGroe,FID,Name,Pat_ID,Pfad,QS,QT,StByte,Typ,ZeitPunkt) values
 For i = 1 to ubound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rBr(i).absPos & "," & datform(rBr(i).AktZeit) & ",'" & rBr(i).Art & "'," & rBr(i).DokGroe & "," & rBr(i).FID & ",'" & rBr(i).Name & "'," &  _
   rBr(i).Pat_ID & ",'" & rBr(i).Pfad & "','" & rBr(i).QS & "','" & rBr(i).QT & "'," & rBr(i).StByte & ",'" & rBr(i).Typ & "'," & datform( _
   rBr(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rBr(i).Art, "briefe", rsc) Then Exit Do
   Case "Name":  If SpMod(rBr(i).Name, "briefe", rsc) Then Exit Do
   Case "Pfad":  If SpMod(rBr(i).Pfad, "briefe", rsc) Then Exit Do
   Case "QS":  If SpMod(rBr(i).QS, "briefe", rsc) Then Exit Do
   Case "QT":  If SpMod(rBr(i).QT, "briefe", rsc) Then Exit Do
   Case "Typ":  If SpMod(rBr(i).Typ, "briefe", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "diagnosen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "diagnosen" & klz & " (absPos,AktZeit,DiagAttr," & _
     "DiagDatum,DiagSeite,DiagSicherheit,DiagText,FID,GesName,ICD,obDauer,Pat_id,StByte) values
 For i = 1 to ubound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDi(i).absPos & "," & datform(rDi(i).AktZeit) & ",'" & rDi(i).DiagAttr & "'," & datform(rDi(i).DiagDatum) & ",'" & rDi(i).DiagSeite & "','" &  _
   rDi(i).DiagSicherheit & "','" & rDi(i).DiagText & "'," & rDi(i).FID & ",'" & rDi(i).GesName & "','" & rDi(i).ICD & "'," & cstr(cint( _
   rDi(i).obDauer)) & "," & rDi(i).Pat_id & "," & rDi(i).StByte & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DiagAttr":  If SpMod(rDi(i).DiagAttr, "diagnosen", rsc) Then Exit Do
   Case "DiagSeite":  If SpMod(rDi(i).DiagSeite, "diagnosen", rsc) Then Exit Do
   Case "DiagSicherheit":  If SpMod(rDi(i).DiagSicherheit, "diagnosen", rsc) Then Exit Do
   Case "DiagText":  If SpMod(rDi(i).DiagText, "diagnosen", rsc) Then Exit Do
   Case "GesName":  If SpMod(rDi(i).GesName, "diagnosen", rsc) Then Exit Do
   Case "ICD":  If SpMod(rDi(i).ICD, "diagnosen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "dokumente" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "dokumente" & klz & " (absPos,AktZeit,DokArt," & _
     "DokGroe,DokName,DokPfad,FID,Pat_ID,QS,QT,Quelldatum,StByte,ZeitPunkt) values
 For i = 1 to ubound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDo(i).absPos & "," & datform(rDo(i).AktZeit) & ",'" & rDo(i).DokArt & "'," & rDo(i).DokGroe & ",'" & rDo(i).DokName & "','" & rDo(i).DokPfad & "'," &  _
   rDo(i).FID & "," & rDo(i).Pat_ID & ",'" & rDo(i).QS & "','" & rDo(i).QT & "'," & datform(rDo(i).Quelldatum) & "," & rDo(i).StByte & "," & datform( _
   rDo(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DokArt":  If SpMod(rDo(i).DokArt, "dokumente", rsc) Then Exit Do
   Case "DokName":  If SpMod(rDo(i).DokName, "dokumente", rsc) Then Exit Do
   Case "DokPfad":  If SpMod(rDo(i).DokPfad, "dokumente", rsc) Then Exit Do
   Case "QS":  If SpMod(rDo(i).QS, "dokumente", rsc) Then Exit Do
   Case "QT":  If SpMod(rDo(i).QT, "dokumente", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "eintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "eintraege" & klz & " (absPos,AktZeit,Art," & _
     "FID,Inhalt,Pat_ID,QS,QT,StByte,ZeitPunkt) values
 For i = 1 to ubound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rEi(i).absPos & "," & datform(rEi(i).AktZeit) & ",'" & rEi(i).Art & "'," & rEi(i).FID & ",'" & rEi(i).Inhalt & "'," & rEi(i).Pat_ID & ",'" &  _
   rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & "," & datform(rEi(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rEi(i).Art, "eintraege", rsc) Then Exit Do
   Case "Inhalt":  If SpMod(rEi(i).Inhalt, "eintraege", rsc) Then Exit Do
   Case "QS":  If SpMod(rEi(i).QS, "eintraege", rsc) Then Exit Do
   Case "QT":  If SpMod(rEi(i).QT, "eintraege", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "forminhaltform_abk" & klz & " (Form_Abk) values
 For i = rFi1 + 1 to ubound(rFi)
'  rFi(i).AktZeit = now()
  sql = sql0 & "('" & rFi(i).Form_Abk & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFi1 = ubound(rFi)
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFi(i).Form_Abk, "forminhaltform_abk", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "formulare" & klz & " (absPos,AktZeit,Form_Abk," & _
     "FormBez,FormID,FormVorl,StByte) values
 For i = rFo1 + 1 to ubound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFo(i).absPos & "," & datform(rFo(i).AktZeit) & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "'," & rFo(i).FormID & ",'" & rFo(i).FormVorl & "'," &  _
   rFo(i).StByte & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFo1 = ubound(rFo)
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFo(i).Form_Abk, "formulare", rsc) Then Exit Do
   Case "FormBez":  If SpMod(rFo(i).FormBez, "formulare", rsc) Then Exit Do
   Case "FormVorl":  If SpMod(rFo(i).FormVorl, "formulare", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
'  sql = "delete FROM " & kla & "forminhfeld" & klz & " where foid in (select foid from " & kla & "forminhkopf" & klz & " where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "delete from " & kla & "forminhkopf" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "forminhkopf" & klz & " (AbsPos,AktZeit,FID," & _
     "FoID,Form_ID,Pat_ID,Satzart,Satzlänge,StByte,ZeitPunkt) values
 For i = 1 to ubound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFr(i).AbsPos & "," & datform(rFr(i).AktZeit) & "," & rFr(i).FID & "," & rFr(i).FoID & "," & rFr(i).Form_ID & "," & rFr(i).Pat_ID & ",'" &  _
   rFr(i).Satzart & "','" & rFr(i).Satzlänge & "'," & rFr(i).StByte & "," & datform(rFr(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rFr(i).Satzart, "forminhkopf", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rFr(i).Satzlänge, "forminhkopf", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "forminhfeld" & klz & " (FeldInhVW,FeldNr,FeldVW," & _
     "FoID,Nr) values
 For i = 1 to ubound(rFm)
'  rFm(i).AktZeit = now()
  sql = sql0 & "(" & rFm(i).FeldInhVW & "," & rFm(i).FeldNr & "," & rFm(i).FeldVW & "," & rFm(i).FoID & "," & rFm(i).Nr & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "kheinweis" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "kheinweis" & klz & " (absPos,AktZeit,Diagnose," & _
     "FID,Pat_ID,StByte,ZeitPunkt,Ziel) values
 For i = 1 to ubound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKh(i).absPos & "," & datform(rKh(i).AktZeit) & ",'" & rKh(i).Diagnose & "'," & rKh(i).FID & "," & rKh(i).Pat_ID & "," & rKh(i).StByte & "," & datform( _
   rKh(i).ZeitPunkt) & ",'" & rKh(i).Ziel & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Diagnose":  If SpMod(rKh(i).Diagnose, "kheinweis", rsc) Then Exit Do
   Case "Ziel":  If SpMod(rKh(i).Ziel, "kheinweis", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "lbanforderungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "lbanforderungen" & klz & " (absPos,AktZeit,AnfText," & _
     "FID,Pat_ID,StByte,ZeitPunkt) values
 For i = 1 to ubound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLb(i).absPos & "," & datform(rLb(i).AktZeit) & ",'" & rLb(i).AnfText & "'," & rLb(i).FID & "," & rLb(i).Pat_ID & "," & rLb(i).StByte & "," & datform( _
   rLb(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AnfText":  If SpMod(rLb(i).AnfText, "lbanforderungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "laborneu" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "laborneu" & klz & " (Abkü,AbsPos,AktZeit," & _
     "Einheit,FertigStGrad,FID,KommentarVW,LangtextVW,Pat_ID,Refnr,StByte,Wert,ZeitPunkt) values
 For i = 1 to ubound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rLa(i).Abkü & "'," & rLa(i).AbsPos & "," & datform(rLa(i).AktZeit) & ",'" & rLa(i).Einheit & "','" & rLa(i).FertigStGrad & "'," &  _
   rLa(i).FID & "," & rLa(i).KommentarVW & "," & rLa(i).LangtextVW & "," & rLa(i).Pat_ID & "," & rLa(i).Refnr & "," & rLa(i).StByte & ",'" & rLa(i).Wert & "'," & datform( _
   rLa(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLa(i).Abkü, "laborneu", rsc) Then Exit Do
   Case "Einheit":  If SpMod(rLa(i).Einheit, "laborneu", rsc) Then Exit Do
   Case "FertigStGrad":  If SpMod(rLa(i).FertigStGrad, "laborneu", rsc) Then Exit Do
   Case "Wert":  If SpMod(rLa(i).Wert, "laborneu", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "leistungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "leistungen" & klz & " (absPos,AktZeit,f5002," & _
     "f5005,f5006,f5009,f5015,f5016,f5021,f5026,FID,Leistung,Med," & _
     "Pat_ID,QS,QT,StByte,ZeitPunkt) values
 For i = 1 to ubound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLe(i).absPos & "," & datform(rLe(i).AktZeit) & ",'" & rLe(i).f5002 & "','" & rLe(i).f5005 & "','" & rLe(i).f5006 & "','" & rLe(i).f5009 & "','" &  _
   rLe(i).f5015 & "','" & rLe(i).f5016 & "','" & rLe(i).f5021 & "','" & rLe(i).f5026 & "'," & rLe(i).FID & ",'" & rLe(i).Leistung & "','" &  _
   rLe(i).Med & "'," & rLe(i).Pat_ID & ",'" & rLe(i).QS & "','" & rLe(i).QT & "'," & rLe(i).StByte & "," & datform(rLe(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "f5002":  If SpMod(rLe(i).f5002, "leistungen", rsc) Then Exit Do
   Case "f5005":  If SpMod(rLe(i).f5005, "leistungen", rsc) Then Exit Do
   Case "f5006":  If SpMod(rLe(i).f5006, "leistungen", rsc) Then Exit Do
   Case "f5009":  If SpMod(rLe(i).f5009, "leistungen", rsc) Then Exit Do
   Case "f5015":  If SpMod(rLe(i).f5015, "leistungen", rsc) Then Exit Do
   Case "f5016":  If SpMod(rLe(i).f5016, "leistungen", rsc) Then Exit Do
   Case "f5021":  If SpMod(rLe(i).f5021, "leistungen", rsc) Then Exit Do
   Case "f5026":  If SpMod(rLe(i).f5026, "leistungen", rsc) Then Exit Do
   Case "Leistung":  If SpMod(rLe(i).Leistung, "leistungen", rsc) Then Exit Do
   Case "Med":  If SpMod(rLe(i).Med, "leistungen", rsc) Then Exit Do
   Case "QS":  If SpMod(rLe(i).QS, "leistungen", rsc) Then Exit Do
   Case "QT":  If SpMod(rLe(i).QT, "leistungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "medplan" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "medplan" & klz & " (ab,AbsPos,AktZeit," & _
     "bBed,Bemerkung,Datum,FeldNr,FID,MedAnfang,Medikament,mi,mo,MPNr," & _
     "nm,Pat_ID,StByte,ZeitPunkt,zn) values
 For i = 1 to ubound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rMe(i).ab & "'," & rMe(i).AbsPos & "," & datform(rMe(i).AktZeit) & "," & cstr(cint(rMe(i).bBed)) & ",'" & rMe(i).Bemerkung & "'," & datform( _
   rMe(i).Datum) & "," & rMe(i).FeldNr & "," & rMe(i).FID & ",'" & rMe(i).MedAnfang & "','" & rMe(i).Medikament & "','" & rMe(i).mi & "','" &  _
   rMe(i).mo & "'," & rMe(i).MPNr & ",'" & rMe(i).nm & "'," & rMe(i).Pat_ID & "," & rMe(i).StByte & "," & datform(rMe(i).ZeitPunkt) & ",'" &  _
   rMe(i).zn & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "ab":  If SpMod(rMe(i).ab, "medplan", rsc) Then Exit Do
   Case "Bemerkung":  If SpMod(rMe(i).Bemerkung, "medplan", rsc) Then Exit Do
   Case "MedAnfang":  If SpMod(rMe(i).MedAnfang, "medplan", rsc) Then Exit Do
   Case "Medikament":  If SpMod(rMe(i).Medikament, "medplan", rsc) Then Exit Do
   Case "mi":  If SpMod(rMe(i).mi, "medplan", rsc) Then Exit Do
   Case "mo":  If SpMod(rMe(i).mo, "medplan", rsc) Then Exit Do
   Case "nm":  If SpMod(rMe(i).nm, "medplan", rsc) Then Exit Do
   Case "zn":  If SpMod(rMe(i).zn, "medplan", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "rezepteintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "rezepteintraege" & klz & " (absPos,AktZeit,FID," & _
     "Medikament,Pat_ID,PZN,QS,QT,Rezept,Rezeptklasse,StByte,ZeitPunkt) values
 For i = 1 to ubound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRe(i).absPos & "," & datform(rRe(i).AktZeit) & "," & rRe(i).FID & ",'" & rRe(i).Medikament & "'," & rRe(i).Pat_ID & ",'" & rRe(i).PZN & "','" &  _
   rRe(i).QS & "','" & rRe(i).QT & "','" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "'," & rRe(i).StByte & "," & datform(rRe(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Medikament":  If SpMod(rRe(i).Medikament, "rezepteintraege", rsc) Then Exit Do
   Case "PZN":  If SpMod(rRe(i).PZN, "rezepteintraege", rsc) Then Exit Do
   Case "QS":  If SpMod(rRe(i).QS, "rezepteintraege", rsc) Then Exit Do
   Case "QT":  If SpMod(rRe(i).QT, "rezepteintraege", rsc) Then Exit Do
   Case "Rezept":  If SpMod(rRe(i).Rezept, "rezepteintraege", rsc) Then Exit Do
   Case "Rezeptklasse":  If SpMod(rRe(i).Rezeptklasse, "rezepteintraege", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "rr" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "rr" & klz & " (absPos,AktZeit,FID," & _
     "Pat_ID,RR,StByte,ZeitPunkt) values
 For i = 1 to ubound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRr(i).absPos & "," & datform(rRr(i).AktZeit) & "," & rRr(i).FID & "," & rRr(i).Pat_ID & ",'" & rRr(i).RR & "'," & rRr(i).StByte & "," & datform( _
   rRr(i).ZeitPunkt) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "RR":  If SpMod(rRr(i).RR, "rr", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 if not Allepat then
  sql = "delete from " & kla & "kvnrue" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call dbcn.Execute(sql)
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "kvnrue" & klz & " (absPos,AktZeit,KVNr," & _
     "Pat_ID,StByte) values
 For i = 1 to ubound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKv(i).absPos & "," & datform(rKv(i).AktZeit) & ",'" & rKv(i).KVNr & "'," & rKv(i).Pat_ID & "," & rKv(i).StByte & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "KVNr":  If SpMod(rKv(i).KVNr, "kvnrue", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern
 Dim i%, sql0$, rAf&
 on error resume next
 if not Allepat then
 End if ' not allePat
 sql0 = " insert " & sqlignore &  "into " & kla & "unbekannte kennungen" & klz & " (absPos,Kennung,Pat_ID," & _
     "StByte) values
 For i = rUn1 + 1 to ubound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rUn(i).absPos & ",'" & rUn(i).Kennung & "'," & rUn(i).Pat_ID & "," & rUn(i).StByte & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rUn1 = ubound(rUn)
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Kennung":  If SpMod(rUn(i).Kennung, "unbekannte kennungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function laborxsaetzeSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxsaetze" & klz & " (Arztname,Arztnr,DatID," & _
     "Erstellungsdatum,Gesamtlänge,KBVPrüfnr,Kundenarztnr,Labor,OrtLabor,OrtPraxis,PLZLabor,PLZPraxis,Satzart," & _
     "Satzlänge,SatzlängeSchluss,StraßeLabor,StraßePraxis,VersionSatzb,Zeichensatz) values
 For i = 0 to ubound(rLs)
'  rLs(i).AktZeit = now()
  sql = sql0 & "('" & rLs(i).Arztname & "','" & rLs(i).Arztnr & "'," & rLs(i).DatID & ",'" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "','" &  _
   rLs(i).KBVPrüfnr & "','" & rLs(i).Kundenarztnr & "','" & rLs(i).Labor & "','" & rLs(i).OrtLabor & "','" & rLs(i).OrtPraxis & "','" & rLs(i).PLZLabor & "','" &  _
   rLs(i).PLZPraxis & "','" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).StraßeLabor & "','" &  _
   rLs(i).StraßePraxis & "','" & rLs(i).VersionSatzb & "','" & rLs(i).Zeichensatz & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Arztname":  If SpMod(rLs(i).Arztname, "laborxsaetze", rsc) Then Exit Do
   Case "Arztnr":  If SpMod(rLs(i).Arztnr, "laborxsaetze", rsc) Then Exit Do
   Case "Erstellungsdatum":  If SpMod(rLs(i).Erstellungsdatum, "laborxsaetze", rsc) Then Exit Do
   Case "Gesamtlänge":  If SpMod(rLs(i).Gesamtlänge, "laborxsaetze", rsc) Then Exit Do
   Case "KBVPrüfnr":  If SpMod(rLs(i).KBVPrüfnr, "laborxsaetze", rsc) Then Exit Do
   Case "Kundenarztnr":  If SpMod(rLs(i).Kundenarztnr, "laborxsaetze", rsc) Then Exit Do
   Case "Labor":  If SpMod(rLs(i).Labor, "laborxsaetze", rsc) Then Exit Do
   Case "OrtLabor":  If SpMod(rLs(i).OrtLabor, "laborxsaetze", rsc) Then Exit Do
   Case "OrtPraxis":  If SpMod(rLs(i).OrtPraxis, "laborxsaetze", rsc) Then Exit Do
   Case "PLZLabor":  If SpMod(rLs(i).PLZLabor, "laborxsaetze", rsc) Then Exit Do
   Case "PLZPraxis":  If SpMod(rLs(i).PLZPraxis, "laborxsaetze", rsc) Then Exit Do
   Case "Satzart":  If SpMod(rLs(i).Satzart, "laborxsaetze", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLs(i).Satzlänge, "laborxsaetze", rsc) Then Exit Do
   Case "SatzlängeSchluss":  If SpMod(rLs(i).SatzlängeSchluss, "laborxsaetze", rsc) Then Exit Do
   Case "StraßeLabor":  If SpMod(rLs(i).StraßeLabor, "laborxsaetze", rsc) Then Exit Do
   Case "StraßePraxis":  If SpMod(rLs(i).StraßePraxis, "laborxsaetze", rsc) Then Exit Do
   Case "VersionSatzb":  If SpMod(rLs(i).VersionSatzb, "laborxsaetze", rsc) Then Exit Do
   Case "Zeichensatz":  If SpMod(rLs(i).Zeichensatz, "laborxsaetze", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxeingel" & klz & " (fertig,Name,Pfad," & _
     "Zp) values
 For i = 1 to ubound(rLg)
'  rLg(i).AktZeit = now()
  sql = sql0 & "(" & cstr(cint(rLg(i).fertig)) & ",'" & rLg(i).Name & "','" & rLg(i).Pfad & "'," & datform(rLg(i).Zp) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Name":  If SpMod(rLg(i).Name, "laborxeingel", rsc) Then Exit Do
   Case "Pfad":  If SpMod(rLg(i).Pfad, "laborxeingel", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxus" & klz & " (Abrechnungstyp,AfN,Auftragsnummer," & _
     "Auftragsschlüssel,AuftrHinw,BefArt,Berichtsdatum,DatID,Eingang,GebDat,GebüOrd,Geschlecht,LWerte," & _
     "Nachname,NVorsatz,Pat_id,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idErwVN,Pat_idErwVNG,Pat_idLaborNeu,Pat_idUrsp," & _
     "Patienteninformation,Satzart,SatzID,Satzlänge,Titel,verglichen,Vorname,ZdiP,ZdüP,ZeitpunktLaborneu) values
 For i = j to j
'  rLu(i).AktZeit = now()
  sql = sql0 & "('" & rLu(i).Abrechnungstyp & "'," & rLu(i).AfN & ",'" & rLu(i).Auftragsnummer & "','" & rLu(i).Auftragsschlüssel & "','" & rLu(i).AuftrHinw & "','" &  _
   rLu(i).BefArt & "','" & rLu(i).Berichtsdatum & "'," & rLu(i).DatID & "," & datform(rLu(i).Eingang) & ",'" & rLu(i).GebDat & "','" &  _
   rLu(i).GebüOrd & "','" & rLu(i).Geschlecht & "','" & rLu(i).LWerte & "','" & rLu(i).Nachname & "','" & rLu(i).NVorsatz & "'," & rLu(i).Pat_id & ",'" &  _
   rLu(i).Pat_idErwG & "','" & rLu(i).Pat_idErwGB & "','" & rLu(i).Pat_idErwGL & "','" & rLu(i).Pat_idErwVN & "','" & rLu(i).Pat_idErwVNG & "','" &  _
   rLu(i).Pat_idLaborNeu & "','" & rLu(i).Pat_idUrsp & "','" & rLu(i).Patienteninformation & "','" & rLu(i).Satzart & "'," & rLu(i).SatzID & ",'" &  _
   rLu(i).Satzlänge & "','" & rLu(i).Titel & "'," & datform(rLu(i).verglichen) & ",'" & rLu(i).Vorname & "'," & rLu(i).ZdiP & "," &  _
   rLu(i).ZdüP & "," & datform(rLu(i).ZeitpunktLaborneu) & ")"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abrechnungstyp":  If SpMod(rLu(i).Abrechnungstyp, "laborxus", rsc) Then Exit Do
   Case "Auftragsnummer":  If SpMod(rLu(i).Auftragsnummer, "laborxus", rsc) Then Exit Do
   Case "Auftragsschlüssel":  If SpMod(rLu(i).Auftragsschlüssel, "laborxus", rsc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLu(i).AuftrHinw, "laborxus", rsc) Then Exit Do
   Case "BefArt":  If SpMod(rLu(i).BefArt, "laborxus", rsc) Then Exit Do
   Case "Berichtsdatum":  If SpMod(rLu(i).Berichtsdatum, "laborxus", rsc) Then Exit Do
   Case "GebDat":  If SpMod(rLu(i).GebDat, "laborxus", rsc) Then Exit Do
   Case "GebüOrd":  If SpMod(rLu(i).GebüOrd, "laborxus", rsc) Then Exit Do
   Case "Geschlecht":  If SpMod(rLu(i).Geschlecht, "laborxus", rsc) Then Exit Do
   Case "LWerte":  If SpMod(rLu(i).LWerte, "laborxus", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rLu(i).Nachname, "laborxus", rsc) Then Exit Do
   Case "NVorsatz":  If SpMod(rLu(i).NVorsatz, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwG":  If SpMod(rLu(i).Pat_idErwG, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwGB":  If SpMod(rLu(i).Pat_idErwGB, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwGL":  If SpMod(rLu(i).Pat_idErwGL, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwVN":  If SpMod(rLu(i).Pat_idErwVN, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwVNG":  If SpMod(rLu(i).Pat_idErwVNG, "laborxus", rsc) Then Exit Do
   Case "Pat_idLaborNeu":  If SpMod(rLu(i).Pat_idLaborNeu, "laborxus", rsc) Then Exit Do
   Case "Pat_idUrsp":  If SpMod(rLu(i).Pat_idUrsp, "laborxus", rsc) Then Exit Do
   Case "Patienteninformation":  If SpMod(rLu(i).Patienteninformation, "laborxus", rsc) Then Exit Do
   Case "Satzart":  If SpMod(rLu(i).Satzart, "laborxus", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLu(i).Satzlänge, "laborxus", rsc) Then Exit Do
   Case "Titel":  If SpMod(rLu(i).Titel, "laborxus", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rLu(i).Vorname, "laborxus", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxbakt" & klz & " (AbnDat,Erklärung,Keimzahl," & _
     "Kommentar,KuQu,QSpez,Quelle,RefNr,Verf) values
 For i = 1 to ubound(rLo)
'  rLo(i).AktZeit = now()
  sql = sql0 & "(" & datform(rLo(i).AbnDat) & ",'" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "','" & rLo(i).Kommentar & "','" & rLo(i).KuQu & "','" &  _
   rLo(i).QSpez & "','" & rLo(i).Quelle & "'," & rLo(i).RefNr & ",'" & rLo(i).Verf & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Erklärung":  If SpMod(rLo(i).Erklärung, "laborxbakt", rsc) Then Exit Do
   Case "Keimzahl":  If SpMod(rLo(i).Keimzahl, "laborxbakt", rsc) Then Exit Do
   Case "Kommentar":  If SpMod(rLo(i).Kommentar, "laborxbakt", rsc) Then Exit Do
   Case "KuQu":  If SpMod(rLo(i).KuQu, "laborxbakt", rsc) Then Exit Do
   Case "QSpez":  If SpMod(rLo(i).QSpez, "laborxbakt", rsc) Then Exit Do
   Case "Quelle":  If SpMod(rLo(i).Quelle, "laborxbakt", rsc) Then Exit Do
   Case "Verf":  If SpMod(rLo(i).Verf, "laborxbakt", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxwert" & klz & " (Abkü,AbnDat,AuftrHinw," & _
     "Einheit,Erklärung,Grenzwerti,Kommentar,Langname,Normbereich,NormO,NormU,QSpez,Quelle," & _
     "RefNr,Teststatus,Wert) values
 For i = 1 to ubound(rLw)
'  rLw(i).AktZeit = now()
  sql = sql0 & "('" & rLw(i).Abkü & "'," & datform(rLw(i).AbnDat) & ",'" & rLw(i).AuftrHinw & "','" & rLw(i).Einheit & "','" & rLw(i).Erklärung & "','" &  _
   rLw(i).Grenzwerti & "','" & rLw(i).Kommentar & "','" & rLw(i).Langname & "','" & rLw(i).Normbereich & "','" & rLw(i).NormO & "','" & rLw(i).NormU & "','" &  _
   rLw(i).QSpez & "','" & rLw(i).Quelle & "'," & rLw(i).RefNr & ",'" & rLw(i).Teststatus & "','" & rLw(i).Wert & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLw(i).Abkü, "laborxwert", rsc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLw(i).AuftrHinw, "laborxwert", rsc) Then Exit Do
   Case "Einheit":  If SpMod(rLw(i).Einheit, "laborxwert", rsc) Then Exit Do
   Case "Erklärung":  If SpMod(rLw(i).Erklärung, "laborxwert", rsc) Then Exit Do
   Case "Grenzwerti":  If SpMod(rLw(i).Grenzwerti, "laborxwert", rsc) Then Exit Do
   Case "Kommentar":  If SpMod(rLw(i).Kommentar, "laborxwert", rsc) Then Exit Do
   Case "Langname":  If SpMod(rLw(i).Langname, "laborxwert", rsc) Then Exit Do
   Case "Normbereich":  If SpMod(rLw(i).Normbereich, "laborxwert", rsc) Then Exit Do
   Case "NormO":  If SpMod(rLw(i).NormO, "laborxwert", rsc) Then Exit Do
   Case "NormU":  If SpMod(rLw(i).NormU, "laborxwert", rsc) Then Exit Do
   Case "QSpez":  If SpMod(rLw(i).QSpez, "laborxwert", rsc) Then Exit Do
   Case "Quelle":  If SpMod(rLw(i).Quelle, "laborxwert", rsc) Then Exit Do
   Case "Teststatus":  If SpMod(rLw(i).Teststatus, "laborxwert", rsc) Then Exit Do
   Case "Wert":  If SpMod(rLw(i).Wert, "laborxwert", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern
 Dim i%, sql0$, rAf&
 on error goto fehler
 sql0 = " insert " & sqlignore &  "into " & kla & "laborxleist" & klz & " (Abkü,Anzahl,EBM," & _
     "goä,RefNr,Verf) values
 For i = 1 to ubound(rLL)
'  rLL(i).AktZeit = now()
  sql = sql0 & "('" & rLL(i).Abkü & "','" & rLL(i).Anzahl & "','" & rLL(i).EBM & "','" & rLL(i).goä & "'," & rLL(i).RefNr & ",'" & rLL(i).Verf & "')"
  Call dbcn.Execute(sql,rAf)', , adAsyncExecute)
  If lies.obmysql and obmitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLL(i).Abkü, "laborxleist", rsc) Then Exit Do
   Case "Anzahl":  If SpMod(rLL(i).Anzahl, "laborxleist", rsc) Then Exit Do
   Case "EBM":  If SpMod(rLL(i).EBM, "laborxleist", rsc) Then Exit Do
   Case "goä":  If SpMod(rLL(i).goä, "laborxleist", rsc) Then Exit Do
   Case "Verf":  If SpMod(rLL(i).Verf, "laborxleist", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obmysql then Resume next else resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function doSpeichern(frm As Lese)
 Dim rAf&
 on error goto fehler
 call namenSpeichern
 call faelleSpeichern
   if not lies.obmysql then
    Call dbcn.CommitTrans
    Call dbcn.BeginTrans
   end if
 call auSpeichern
 call briefeSpeichern
 call diagnosenSpeichern
 call dokumenteSpeichern
 call eintraegeSpeichern
 call forminhaltform_abkSpeichern
   if not lies.obmysql then
    Call dbcn.CommitTrans
    Call dbcn.BeginTrans
   end if
 call formulareSpeichern
 call forminhkopfSpeichern
 call forminhfeldSpeichern
 call kheinweisSpeichern
 call lbanforderungenSpeichern
 call laborneuSpeichern
 call leistungenSpeichern
 call medplanSpeichern
 call rezepteintraegeSpeichern
 call rrSpeichern
 call kvnrueSpeichern
 call unbekannte_kennungenSpeichern
 Call DBCn.Execute("update namen set aktZeit = " & datform(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_ID,rAf)
 If rAf <> 1 Then 
  frm.Ausgabe = "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_ID & " " & rNa(0).Nachname & " " & rNa(0).Vorname & vbcrfl & altausgabe
  altausgabe = frm.ausgabe
 end if
 exit function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

