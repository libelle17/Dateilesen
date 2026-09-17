Attribute VB_Name = "Typen"
Dim sql$

Public Type namen
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 An1da As Integer 'Anamnese S.1 da
 An2da As Integer 'Anamnese S.2 da
 AnAllgda As Integer 'Anamnese allgemein da
 AufnDat As Date '3610
 Checkda As Integer 'Checkliste da
 DienstTel As String '3629
 DMTypaD As String 'aus Diagnosen
 Email As String 'Email
 GebDat As Date '3103
 Geschlecht As String '3110
 intZOGP As String '3635, interne Zuordnung Arzt bei GP
 KVKStatus As String '3108
 KVNr As String '3630
 lfdnr As Long 'laufende Patientennummer
 Nachname As String '3101
 Notiz As String '3634
 NVorsatz As String '3100
 Ort As String '3113
 Pat_ID As Long '3000
 Plz As String '3112
 PrivatFax As String '3629
 PrivatMobil As String '3629
 PrivatTel As String '3629
 PrivatTel_2 As String '3629
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Straße As String '3107
 Titel As String '3104
 Versichertennummer As String '3105
 Verwandt As String '3632
 Vorname As String '3102
 Weggeldzone As String '3631
 zubenach As String '3633
End Type

Public Type faelle
 AbrAr As String '4107, Abrechnungsart (1 = Primärkassen)
 AbrGb As String '4122, Abrechnungsgebiet (07 = Diabetes)
 absPos As Long 'Zeile in der BDT-Datei
 AdNam As String '4602(1) Name bei Musterfrau
 AdOrt As String '4602(4) Ort bei Musterfrau
 AdPlz As String '4602(3) PLZ bei Musterfrau
 AdStr As String '4602(2) Straße bei Musterfrau
 AktZeit As Date 'Aktualisierungszeit
 altQuart As String '
 AndÜw As String '4219, anderer Überweiser
 ausgst As Date '4102 ('ausgestellt am')
 BhFB As Date '4150
 BhFE As Date '4604, Behandlungsfall: Ende, bei Privatpatienten
 BhFE1 As Date '4151
 BhFE2 As Date '4152
 f4202 As String '4202
 f4206 As String '4206, ?
 f4210 As String '4210, unbekannt
 f4237 As String '4237, ? (nur bei Musterw)
 Fanf As Date 'Fallanfang
 FID As Long 'Primärschlüssel
 GebOr As String '4121, Gebührenordnung (1 = BMÄ, 2)
 GOÄKatName As String '4580 (2): Privat-Abrechnungskatalog
 GOÄKatNr As String '4580 (1): Katalog-Nummer
 IK As String '4111 (auch patientenbezogen)
 KtrAbrB As String '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 KVKs As String '4112
 KVKserg As String '4113
 lfdnr As Long 'laufende Fallnummer
 lVorl As Date '4109, letzte Vorlage
 Mahnfrist As String '4403, Mahnfrist bis
 Nachname As String '3101
 Pat_ID As Long '3000
 PGeb As String '4401, Praxisgebühr
 PGebErg As String '4402, ?
 QAnf As Date 'Quartalsanfang
 QEnd As Date 'Quartalsende
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 Quartal As String '4101
 s8000 As String '8000, ???
 s8100 As String '8100, ???
 SchGr As String '4239, Scheingruppe
 statBehTage As Long '4238
 statKlasse As String '4236
 statNuller As String '4216, nu bei Musterfrau 16 Nuller
 StByte As Long 'Ordnungsnummer der Datenübertragung
 TMFNr As String '4144 Fallnummer in Turbomed
 ÜbwV As String '4218, überwiesen von
 ÜWNaN As String '4231(3): Nachname des Überweisers
 ÜWNNr As String '4231(4): KV-Nummer des Überweisers
 ÜwText As String '4209: Auftrags- / erläuternder Text zur Überweisung
 ÜWTit As String '4231(3): Titel des Überweisers
 ÜWVor As String '4231(2): Vorname des Überweisers
 ÜWVsw As String '4231(2b): Vorsatzwort des Überweisers
 ÜWZiel As String '4220
 VKNr As String '4104
 Vorname As String '3102
 Weiterbeh As String '4243, Weiterbehandelnder
End Type

Public Type au
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 Beginn As String '6285 1. Hälfte
 Ende As String '6285 2. Hälfte
 FID As Long 'Fall-Bezug
 ICDs As String '6286
 Pat_ID As Long '3000
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '6200 + 6201
End Type

Public Type briefe
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 Art As String '
 DokGroe As Long 'Größe der Datei
 FID As Long 'Fall-Bezug
 Name As String '
 Pat_ID As Long '
 Pfad As String '
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Typ As String '
 Zeitpunkt As Date '
End Type

Public Type Diagnosen
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 DiagAttr As String '6006
 DiagDatum As Date '
 DiagSeite As String '6004
 DiagSicherheit As String '6003
 DiagText As String '
 FID As Long 'Fall-Bezug
 GesName As String '
 ICD As String '
 ID1 As Long '
 obDauer As Integer 'ob Dauerdiagnose
 Pat_ID As Long 'Bezug auf Anamneseblattt
 StByte As Long 'Ordnungsnummer der Datenübertragung
End Type

Public Type dokumente
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 DokArt As String '
 DokGroe As Long 'Dokument-Größe
 DokName As String '
 DokPfad As String '
 FID As Long 'Fall-Bezug
 Pat_ID As Long '
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 Quelldatum As Date 'Datum, auf das sich das Dokument bezieht
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '
End Type

Public Type eintraege
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 Art As String '6330
 FID As Long 'Fall-Bezug
 Inhalt As String '8480
 Pat_ID As Long '3000
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '
End Type

Public Type forminhaltform_abk
 Form_Abk As String '
 Form_AbkVW As Long '
End Type

Public Type formulare
 absPos As Long 'Zeile in BDT-Datei
 AktZeit As Date 'Zeitpunkt der Aktualisierung
 Form_Abk As String '
 FormBez As String '
 FormID As Long '
 FormVorl As String '
 StByte As Long 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhkopf
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 FID As Long 'Fall-Bezug
 FoID As Long '
 Form_ID As Long '
 Pat_ID As Long '
 Satzart As String '8000
 Satzlänge As String '8100
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '
End Type

Public Type forminhfeld
 FeldInhVW As Long '
 FeldInh As String
 FeldNr As Integer '
 FeldVW As Long '
 Feld As String
 FoID As Long '
 Nr As Integer '
End Type

Public Type kheinweis
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 Diagnose As String '6230
 FID As Long 'Fall-Bezug
 Pat_ID As Long '3000
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '
 Ziel As String '6291
End Type

Public Type lbanforderungen
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 AnfText As String '6280
 FID As Long 'Fall-Bezug
 Pat_ID As Long '3000
 StByte As Long 'Statusbyte
 Zeitpunkt As Date '6200 + 6201
End Type

Public Type laborneu
 Abkü As String '8410
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 Einheit As String '8421
 FertigStGrad As String '8401
 FID As Long 'Fall-Bezug
 KommentarVW As Long '8480
 Kommentar As String
 LangtextVW As Long '8411
 Langtext As String
 Pat_ID As Long '3000
 RefNr As Long 'Bezug auf LaborXUS
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Wert As String '8420
 Zeitpunkt As Date '
End Type

Public Type leistungen
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 f5002 As String '5002
 f5005 As String '5005
 f5006 As String '5006
 f5009 As String '5009
 f5015 As String '5015
 f5016 As String '5016
 f5021 As String '5021
 f5026 As String '5026
 FID As Long 'Fall-Bezug
 Leistung As String '5001
 Med As String '5010
 Pat_ID As Long '3000
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '5000 + 6201
End Type

Public Type medplan
 ab As String '
 absPos As Long 'Zeile in der BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 bBed As Integer '
 Bemerkung As String '
 Datum As Date 'Zeitpunkt aus dem Kopf des Medikamentenplans
 FeldNr As Integer '
 FID As Long 'Fall-Bezug
 MedAnfang As String '
 Medikament As String '
 mi As String '
 mo As String '
 MPNr As Long 'Ordnungsziffer für Medikamentenplan
 nm As String '
 Pat_ID As Long '3000
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date 'Zeitpunkt, der Speicherung im Turbomed
 zn As String '
End Type

Public Type rezepteintraege
 absPos As Long 'Zeile in BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 FID As Long 'Fall-Bezug
 Medikament As String '3652(2), 6218(4)
 Pat_ID As Long '3000
 PZN As String '6210(2), 6218(3)
 QS As String 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'Quartal des Behandlungsfallbeginns
 Rezept As String '6210, 3652(1), 6218(1)
 Rezeptklasse As String '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 StByte As Long 'Statusbyte
 Zeitpunkt As Date '6200 + 6201
End Type

Public Type RR
 absPos As Long 'Zeile in BDT-Datei
 AktZeit As Date 'Aktualisierungszeit
 FID As Long 'Fall-Bezug
 Pat_ID As Long '3000
 RR As String '6230
 StByte As Long 'Ordnungsnummer der Datenübertragung
 Zeitpunkt As Date '6200 + 6201
End Type

Public Type kvnrue
 absPos As Long 'Zeile in BDT-Datei
 AktZeit As Date 'Zeit der Aktualisuerung aus der BDT-Datei
 KVNr As String '
 lfdnr As Long '
 Pat_ID As Long '
 StByte As Long 'Ordnungsnummer der Datenübertragung
End Type

Public Type unbekannte_kennungen
 absPos As Long '
 Kennung As String '
 StByte As Long '
End Type

Public Type laborxsaetze
 Arztname As String '203 Arztname (Turbomed)
 Arztnr As String '201 Arztnummer (Turbomed)
 DatID As Long 'Bezug zu LaborEingelesen
 Erstellungsdatum As String '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge As String '9202 Gesamtlänge des Datenpaketes (Turbomed)
 KBVPrüfnr As String '101 KBV-Prüfnummer (Turbomed)
 Kundenarztnr As String '8312 Kundenarztnummer (Turbomed)
 Labor As String '8320 Labor
 OrtLabor As String '8323 Ort der Laboradresse (Turbomed)
 OrtPraxis As String '216 Ort der Praxis (Turbomed)
 PLZLabor As String '8322 PLZ der Laboradresse (Turbomed)
 PLZPraxis As String '215 PLZ der Praxis (Turbomed)
 Satzart As String '8000 Satzart (Turbomed)
 SatzID As Long 'zum Bezug für LaborUS
 Satzlänge As String '8100 Satzlänge (Turbomed)
 SatzlängeSchluss As String '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 StraßeLabor As String '8321 Straße der Laboradresse (Turbomed)
 StraßePraxis As String '205 Straße der Praxis (Turbomed)
 VersionSatzb As String '9212 Version der Satzbeschreibung (Turbomed)
 Zeichensatz As String '9106 verwendeter Zeichensatz (Turbomed)
End Type

Public Type laborxeingel
 DatID As Long 'Bezug auf LaborEingelesen
 Name As String 'Name der eingelesenen Labordatei ohne Endung
 Pfad As String 'Pfadname
 Zp As Date 'Einlesezeitpunkt
End Type

Public Type laborxus
 Abrechnungstyp As String '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 AfN As Integer 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
 Auftragsnummer As String '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As String '8311 Anforderungsnr d Labors (Turbomed)
 AuftrHinw As String '8490 Auftragsbezogene Hinweise (Turbomed)
 BefArt As String '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Berichtsdatum As String '8302 Berichtsdatum
 DatID As Long 'Bezug auf LaborEingelesen
 Eingang As Date '8301 Eingangsdatum in Datumsform
 GebDat As String '3103
 GebüOrd As String '8403 Gebührenordnung (Turbomed)
 Geschlecht As String '8407 Geschlecht (Turbomed)
 LWerte As String 'Laborwerte, die zur Zuordnung geführt haben
 Nachname As String '3101
 Pat_ID As Long '
 Pat_idErwG As String 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB As String 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL As String 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idErwVN As String 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwVNG As String 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idLaborNeu As String 'Pat_ids von in Laborneu zuordnbaren Patienten
 Pat_idUrsp As String 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Patienteninformation As String '8405 Patienteninformation (Turbomed)
 RefNr As Long 'Bezug auf LaborWert
 Satzart As String '8000 Satzart (Turbomed)
 SatzID As Long 'Bezug auf LaborXSätze
 Satzlänge As String '8100 Satzlänge (Turbomed)
 Titel As String '3104
 verglichen As Date 'Datum, zu dem Datensatz zuletzt verglichen wurde
 Vorname As String '3102
 ZdiP As Long 'Zahl der infragekommenden Patienten
 ZdüP As Integer 'Zahl der verglichenen Parameter
 ZeitpunktLaborneu As Date 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
End Type

Public Type laborxbakt
 AbnDat As Date '8432 Abnahmedatum (Turbomed)
 Erklärung As String '
 Keimzahl As String '
 Kommentar As String '8480 Ergebnistest (Turbomed)
 KuQu As String '8428 Probenmaterial-Ident (Turbomed)
 QSpez As String '8431 Probenmaterial-Spezifikation (Turbomed)
 Quelle As String '8430 Probenmaterial-Bezeichnung (Turbomed)
 RefNr As Long '
 Verf As String '
End Type

Public Type laborxwert
 Abkü As String '8410 Test-Ident  (Turbomed)
 AbnDat As Date '8432 Abnahmedatum (Turbomed)
 AuftrHinw As String '8490 Auftragsbezogene Hinweise (Turbomed)
 Einheit As String '8421 Einheit (Turbomed)
 Erklärung As String '8470 Testbezogene Hinweise (Turbomed)
 Grenzwerti As String '8422 Grenzwertindikator (Turbomed)
 Kommentar As String '8480 Ergebnistext (Turbomed)
 Langname As String '8411 Testbezeichnung (Turbomed)
 Normbereich As String '8460 Normalwert-Text (Turbomed)
 NormO As String '8461 Normobergrenze
 NormU As String '8461 Normuntergrenze
 QSpez As String '8431 Probenmaterial-Spezifikation (Turbomed)
 Quelle As String '8430 Probenmaterial-Bezeichnung (Turbomed)
 RefNr As Long 'Bezug auf LaborUS
 Teststatus As String '8418 Teststatus (Turbomed)
 Wert As String '8420 Ergebniswert (Turbomed)
End Type

Public Type laborxleist
 Abkü As String '8410 Test-Ident (Turbomed)
 Anzahl As String '5005
 EBM As String '5001 GNR (Turbomed)
 goä As String '8406
 RefNr As Long 'Bezug auf LaborUS
 Verf As String '8434
End Type

Public Type anamnesebogen
 AktZeit As Date 'Aktualisierungszeit
 Albumin_zuletzt As String ', letztes Albumin:
 Alkohol As String '^:
 Ameisen_Ausmaß As String '<, Ausmaß:
 Ameisenlaufen As String '^:
 Amputation As String ',:
 andere_Nierenerkrankung As String ', andere Nierenerkrankung:
 Angina_pectoris As String ',:
 Anrede As String '
 ASR As String ',:
 Aufschreiben As String '<, Dokumentation:
 Augensp_Befund As String '<, Befund:
 Augensp_zuletzt As String '^Letzte Augenspiegelung:
 Bauch As String ', Abdomen:
 BDselbst As String '^Blutdruckselbstmessung:
 Beinbefund As String '^:
 Beindurchblutungsstörung As String '^:
 BeinödVen As String ', Beinödeme/ Venenkrankheiten:
 Bewegungseinschränkungen As String ',:
 Bewußtlos_pa As String '<, bewußtlos deshalb:
 BHD_beh_mit As String '<, behandelt mit:
 BHD_seit As String '<seit:
 Blutdruckwerte As String '^Blutdruckwerte:
 Bluthochdruck As String '^Bluthochdruck:
 Broteinheiten_abends As String '<, abends
 Broteinheiten_früh As String '<, früh
 Broteinheiten_gesamt As String '^Broteinheiten:gesamt
 Broteinheiten_mittags As String '<, mittags
 Broteinheiten_nachmittags As String '<, nachmittags
 Broteinheiten_nachts As String '<, nachts
 Broteinheiten_ZM_früh As String '<, Zwischenmahlzeit vormittags
 Bypass_kardial As Integer ',:
 Bypaß_peripher As Integer ',:
 Bypass_wann As String '<, wann:
 BZgr300_pM As String ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 BZMessungen_p_W_nachts As String '<, nachts:
 BZMessungen_pW As String '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE As String '<, davon nach dem Essen:
 BZMessungen_selbst As String '^Blutzuckermessung:Selbstmessung?
 BZWerte_n_d_Essen As String '<, nach dem Essen:
 BZWerte_v_d_Essen As String '^Blutzuckerwerte vor dem Essen:
 Carotiden As String ', Halsschlagadern:
 Diabet_Nierenschaden As String '^Diabetischer Nierenschaden:
 Diabetes_seit As String '<seit
 DiabetesMedikament_1 As String '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge As String '<
 DiabetesMedikament_2 As String '<,
 DiabetesMedikament_2_Menge As String '<
 DiabetesMedikament_3 As String '<,
 DiabetesMedikament_3_Menge As String '<
 DiabetesMedikament_4 As String '<,
 DiabetesMedikament_4_Menge As String '<,
 Diabetestyp As String '^Diabetes Typ
 Diagnosen As String '
 Dialyse As Integer ',:
 Dialyse_seit As String '<seit
 DMP As String 'ob Pat. bei HA im DMP
 DMPhier As Date 'ob Pat hier im DMP
 DMSchL As Integer 'Zahl der abgerechneten DMP-Schulungen hier
 DMSchulz As Integer 'Zahl der DMP-Schulungen hier
 Druckstellen As String ',:
 Einlagen As String ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Entleerungsstörungen_Harnblase As String ',:
 Entleerungsstörungen_Magen As String '^:
 erhöht As String '<, Befund:
 Essenszeit_abends As String '<, abends
 Essenszeit_früh As String '^Essenszeiten:früh
 Essenszeit_mittags As String '<, mittags
 Essenszeit_nachmittags As String '<, nachmittags
 Essenszeit_spät As String '<, spät
 Essenszeit_vormittags As String '<, vormittags
 Familienanamnese As String '^:
 Folgeerkrankungen_Haut As String '^:
 Fremde_Hilfe_pa As String '<, fremde Hilfe deshalb nötig:
 Fußpflege As String '^:
 GebDat As Date ', geb.
 gemessen_am As Date '<, gemessen
 Gerät As String '<:
 Geschwür As String ',:
 Gewicht As Single ',:
 Größe As Single '^:
 Grund_für_Vorstellung As String '^:
 HANr As String 'mit "/"
 HANr2 As String 'mit "/"
 Hausarzt As String 'Hausarzt laut Anamnesebogen
 Herz As String '^:
 Herzinfarkt As String ',:
 Herzinfarkt_wann As String '<, wann:
 Herzkrankheit As String '^Herzkrankheit:
 Herzkrankheit_Beschreibung As String ', Beschreibung:
 Herzschwäche As String ',:
 Hirndurchblutungsstörung As String '^:
 Hyperkeratosen As String ',:
 Insulin_seit As String ', Insulin seit
 Insulinpumpe As Integer '^:
 Insulinpumpe_Marke As String '<, Marke:
 Insulinpumpe_seit As String '<seit
 Jahr_letzte_Diabetesschulung As String '^Letzte Diabetesschulung:
 Kalt_Warm As String ', Kalt-Warm-Diskrimination:
 Keto_pa As String '^Zahl der Ketoazidosen pro Jahr:
 Kraft_Knie As String '<, Knie:
 Kraft_Zehenbeuger As String '<, Zehenbeuger:
 Kraft_Zehenheber As String '^Kraft:Zehenheber
 letzte_Änderung As Date 'Datum der letzten Änderung
 letztes_HbA1c As String '^Letztes HbA1c:
 Liphypertrophien_Abdomen As String '^Liphypertrophien:Abdomen
 Liphypertrophien_Arme As String '<, Arme:
 Liphypertrophien_Beine As String '<, Beine:
 LK As String ', Lymphknoten:
 Lunge As String ',:
 Monofilamenttest As String ',:
 Mundhöhle As String ',:
 Nachname As String '-
 Netzhaut_gelasert As String ', Netzhaut schon gelasert:
 Neue_Fußkomplikationen As String '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Neuro_sonst As String '^Sonstige neurologische Befunde:
 NL As String ', Nierenlager:
 NNH As String ', Nasennebenhöhlen:
 NVorsatz As String '
 ob As Integer 'für verschiedene Aktionen
 obAn1eing As Integer 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As Integer 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As Integer 'ob Anamneseblatt allgemein eingegeben wurde
 obBZausgew As Integer 'ob Blutzuckergerät ausgewechselt
 obCheck As Integer 'ob Checkliste vorliegt
 obDMPaufgekl As String 'ob über DMP aufgeklärt
 Oberflächensensibilität As String '^:
 obMBlAusgeh As Integer 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz As Integer 'ob von Med. Netz geschickt
 obOSaufgek As Integer 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As Integer 'ob über Podologie aufgeklärt
 obSchulaufgek As String 'ob über Schulung aufgeklärt
 Ort_Schulung As String '<in
 Pat_ID As Long '
 pAVK_Beschreibung As String ', Beschreibung der Beinbeschwerden:
 Podologie As String ',:
 Prim As Long 'Primärschlüssel
 PSR As String ',:
 PTCA_oder_Stent As String ',:
 Puls_Adp As String '<,Fußrücken:
 Puls_Atp As String '<,Innenknöchel:
 Puls_Kniekehle As String '<,Kniekehle:
 Puls_Leiste As String '^Pulse:Leiste
 QS As String 'Quartal sortiert von vorgestellt
 QT As String 'Quartal sortiert von vorgestellt
 RR As String '^Blutdruck:
 RRSchulz As Integer 'Zahl der Hypertonie-Schulungen hier
 RRTurboMed As String '
 Schaufensterkrankheit As String ',:
 Schlaganfall As String ',:
 Schulung As String 'ob Schulungsbedarf
 Schwanger As String '^Aktuelle Schwangerschaft:
 Schwanger_seit As String '<, seit:
 Schwindel_Aufstehen As String ',:
 SD As String ', Schilddrüse:
 Sehminderung_unbehebbar As String ', mit Brille nicht behebbare Sehminderung:
 Sexualstörung As String '^:
 Sexualstörung_seit As String '<seit
 Spritz_Eß_Abstand_abends As String '<, abends
 Spritz_Eß_Abstand_früh As String '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags As String '<, mittags
 Spritzstelle_abends As String '<, abends
 Spritzstelle_früh As String '^Spritzstellen:früh
 Spritzstelle_mittags As String '<, mittags
 Spritzstelle_nachts As String '<, nachts
 Tabak As String ',:
 Tabletten_seit As String ', Tabletten seit
 Tendenz As String '<, Tendenz
 Ther1 As String 'Diät, OAD, CT, ICT, CSII
 TherAkt As String 'Diät, OAD, CT, ICT, CSII
 Titel As String '
 Tkz As Integer 'Tod-Kennzeichen
 Ulcera As String ',:
 Unterzucker_pM As String '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig As String '<, rechtzeitig bemerkt:
 UZ_Tageszeit As String '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Verformungen As String ',:
 Verformungen_Beschreibung As String '<Beschreibung:
 Versicherung As String '
 Versicherungsart As String '
 Vibration_Großzehe As String '<, Großzehe:
 Vibration_IK As String ', Vibrationsempfinden Innenknöchel:
 Vorgestellt As Date 'Erstvorstellung
 vorherige_Werte As String '<, vorher:
 Vorname As String '
 Weitere_Anamnese As String '^:
 Weitere_Befunde As String ', weitere Befunde:
 Weitere_Medikation As String '^:
 WS As String ', Wirbelsäule:
 Zähne As String ',:
End Type

Public rNa() As namen
Public rFa() As faelle
Public rAu() As au
Public rBr() As briefe
Public rDi() As Diagnosen
Public rDo() As dokumente
Public rEi() As eintraege
Public rFi() As forminhaltform_abk
Public rFo() As formulare
Public rFr() As forminhkopf
Public rFm() As forminhfeld
Public rKh() As kheinweis
Public rLb() As lbanforderungen
Public rLa() As laborneu
Public rLe() As leistungen
Public rMe() As medplan
Public rRe() As rezepteintraege
Public rRr() As RR
Public rKv() As kvnrue
Public rUn() As unbekannte_kennungen
Public rLs() As laborxsaetze
Public rLg() As laborxeingel
Public rLu() As laborxus
Public rLo() As laborxbakt
Public rLw() As laborxwert
Public rLL() As laborxleist
Public rAna() As anamnesebogen

Public Function Tinit()
 Static wdh%
 ReDim rAna(0)
 ReDim rNa(0)
 ReDim rFa(0)
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 If wdh = 0 Then ReDim rFi(0)
 If wdh = 0 Then ReDim rFo(0)
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
 Static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 wdh = -1
End Function ' Tinit

Public Function doLösch(frm As Lese, Tbl$)
 Set rs = DBCn.Execute("select count(*) as ct from " + kla + Tbl + klz)
 frm.Ausgabe = "Lösche: " & kla & Tbl & klz & " (" & rs!ct & " Datensätze)" & vbCrLf & frm.Ausgabe
 If obMySQL Then
  sql = "truncate table " & kla & Tbl & klz
 Else
  sql = "delete from " & kla & Tbl & klz
 End If
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doLösch

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 0")
 Call doLösch(frm, "unbekannte kennungen")
 Call doLösch(frm, "kvnrue")
 Call doLösch(frm, "rr")
 Call doLösch(frm, "rezepteintraege")
 Call doLösch(frm, "medplan")
 Call doLösch(frm, "leistungen")
 Call doLösch(frm, "laborneu")
 Call doLösch(frm, "lbanforderungen")
 Call doLösch(frm, "kheinweis")
 Call doLösch(frm, "forminhfeld")
 Call doLösch(frm, "forminhkopf")
 Call doLösch(frm, "formulare")
 Call doLösch(frm, "forminhaltform_abk")
 Call doLösch(frm, "eintraege")
 Call doLösch(frm, "dokumente")
 Call doLösch(frm, "diagnosen")
 Call doLösch(frm, "briefe")
 Call doLösch(frm, "au")
 Call doLösch(frm, "faelle")
 Call doLösch(frm, "namen")
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 1")
 Exit Function
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
 End Select
End Function ' AllesLösch

Public Function LabLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 0")
 Call doLösch(frm, "laborxleist")
 Call doLösch(frm, "laborxwert")
 Call doLösch(frm, "laborxbakt")
 Call doLösch(frm, "laborxus")
 Call doLösch(frm, "laborxeingel")
 Call doLösch(frm, "laborxsaetze")
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 1")
 Exit Function
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
 End Select
End Function ' LabLösch


Public Function namenSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "namen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "namen" & klz & " (absPos," & _
     "AktZeit,An1da,An2da,AnAllgda,AufnDat,Checkda,DienstTel,DMTypaD,Email,GebDat," & _
     "Geschlecht,intZOGP,KVKStatus,KVNr,lfdnr,Nachname,Notiz,NVorsatz,Ort,Pat_ID," & _
     "Plz,PrivatFax,PrivatMobil,PrivatTel,PrivatTel_2,StByte,Straße,Titel,Versichertennummer,Verwandt," & _
     "Vorname,Weggeldzone,zubenach) values"
 Else
  sql0 = "insert into " & kla & "namen" & klz & " (absPos," & _
     "AktZeit,An1da,An2da,AnAllgda,AufnDat,Checkda,DienstTel,DMTypaD,Email,GebDat," & _
     "Geschlecht,intZOGP,KVKStatus,KVNr,lfdnr,Nachname,Notiz,NVorsatz,Ort,Pat_ID," & _
     "Plz,PrivatFax,PrivatMobil,PrivatTel,PrivatTel_2,StByte,Straße,Titel,Versichertennummer,Verwandt," & _
     "Vorname,Weggeldzone,zubenach) values"
 End If
 For i = 0 To UBound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rNa(i).absPos & "," & datForm(rNa(i).AktZeit) & "," & CStr(CInt(rNa(i).An1da)) & "," & CStr(CInt(rNa(i).An2da)) & "," & CStr(CInt(rNa(i).AnAllgda)) & _
   "," & datForm(rNa(i).AufnDat) & "," & CStr(CInt(rNa(i).Checkda)) & ",'" & rNa(i).DienstTel & "','" & rNa(i).DMTypaD & "','" & rNa(i).Email & _
   "'," & datForm(rNa(i).GebDat) & ",'" & rNa(i).Geschlecht & "','" & rNa(i).intZOGP & "','" & rNa(i).KVKStatus & "','" & rNa(i).KVNr & _
   "'," & rNa(i).lfdnr & ",'" & rNa(i).Nachname & "','" & rNa(i).Notiz & "','" & rNa(i).NVorsatz & "','" & rNa(i).Ort & "'," & rNa(i).Pat_ID & _
   ",'" & rNa(i).Plz & "','" & rNa(i).PrivatFax & "','" & rNa(i).PrivatMobil & "','" & rNa(i).PrivatTel & "','" & rNa(i).PrivatTel_2 & "'," & rNa(i).StByte & _
   ",'" & rNa(i).Straße & "','" & rNa(i).Titel & "','" & rNa(i).Versichertennummer & "','" & rNa(i).Verwandt & "','" & rNa(i).Vorname & _
   "','" & rNa(i).Weggeldzone & "','" & rNa(i).zubenach & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DienstTel":  If SpMod(rNa(i).DienstTel, rsc!character_maximum_length, "namen", "DienstTel") Then Exit Do
   Case "DMTypaD":  If SpMod(rNa(i).DMTypaD, rsc!character_maximum_length, "namen", "DMTypaD") Then Exit Do
   Case "Email":  If SpMod(rNa(i).Email, rsc!character_maximum_length, "namen", "Email") Then Exit Do
   Case "Geschlecht":  If SpMod(rNa(i).Geschlecht, rsc!character_maximum_length, "namen", "Geschlecht") Then Exit Do
   Case "intZOGP":  If SpMod(rNa(i).intZOGP, rsc!character_maximum_length, "namen", "intZOGP") Then Exit Do
   Case "KVKStatus":  If SpMod(rNa(i).KVKStatus, rsc!character_maximum_length, "namen", "KVKStatus") Then Exit Do
   Case "KVNr":  If SpMod(rNa(i).KVNr, rsc!character_maximum_length, "namen", "KVNr") Then Exit Do
   Case "Nachname":  If SpMod(rNa(i).Nachname, rsc!character_maximum_length, "namen", "Nachname") Then Exit Do
   Case "Notiz":  If SpMod(rNa(i).Notiz, rsc!character_maximum_length, "namen", "Notiz") Then Exit Do
   Case "NVorsatz":  If SpMod(rNa(i).NVorsatz, rsc!character_maximum_length, "namen", "NVorsatz") Then Exit Do
   Case "Ort":  If SpMod(rNa(i).Ort, rsc!character_maximum_length, "namen", "Ort") Then Exit Do
   Case "Plz":  If SpMod(rNa(i).Plz, rsc!character_maximum_length, "namen", "Plz") Then Exit Do
   Case "PrivatFax":  If SpMod(rNa(i).PrivatFax, rsc!character_maximum_length, "namen", "PrivatFax") Then Exit Do
   Case "PrivatMobil":  If SpMod(rNa(i).PrivatMobil, rsc!character_maximum_length, "namen", "PrivatMobil") Then Exit Do
   Case "PrivatTel":  If SpMod(rNa(i).PrivatTel, rsc!character_maximum_length, "namen", "PrivatTel") Then Exit Do
   Case "PrivatTel_2":  If SpMod(rNa(i).PrivatTel_2, rsc!character_maximum_length, "namen", "PrivatTel_2") Then Exit Do
   Case "Straße":  If SpMod(rNa(i).Straße, rsc!character_maximum_length, "namen", "Straße") Then Exit Do
   Case "Titel":  If SpMod(rNa(i).Titel, rsc!character_maximum_length, "namen", "Titel") Then Exit Do
   Case "Versichertennummer":  If SpMod(rNa(i).Versichertennummer, rsc!character_maximum_length, "namen", "Versichertennummer") Then Exit Do
   Case "Verwandt":  If SpMod(rNa(i).Verwandt, rsc!character_maximum_length, "namen", "Verwandt") Then Exit Do
   Case "Vorname":  If SpMod(rNa(i).Vorname, rsc!character_maximum_length, "namen", "Vorname") Then Exit Do
   Case "Weggeldzone":  If SpMod(rNa(i).Weggeldzone, rsc!character_maximum_length, "namen", "Weggeldzone") Then Exit Do
   Case "zubenach":  If SpMod(rNa(i).zubenach, rsc!character_maximum_length, "namen", "zubenach") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function faelleSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "faelle" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "faelle" & klz & " (AbrAr," & _
     "AbrGb,absPos,AdNam,AdOrt,AdPlz,AdStr,AktZeit,altQuart,AndÜw,ausgst," & _
     "BhFB,BhFE,BhFE1,BhFE2,f4202,f4206,f4210,f4237,Fanf,GebOr," & _
     "GOÄKatName,GOÄKatNr,IK,KtrAbrB,KVKs,KVKserg,lfdnr,lVorl,Mahnfrist,Nachname," & _
     "Pat_ID,PGeb,PGebErg,QAnf,QEnd,QS,QT,Quartal,s8000,s8100," & _
     "SchGr,statBehTage,statKlasse,statNuller,StByte,TMFNr,ÜbwV,ÜWNaN,ÜWNNr,ÜwText," & _
     "ÜWTit,ÜWVor,ÜWVsw,ÜWZiel,VKNr,Vorname,Weiterbeh) values"
 Else
  sql0 = "insert into " & kla & "faelle" & klz & " (AbrAr," & _
     "AbrGb,absPos,AdNam,AdOrt,AdPlz,AdStr,AktZeit,altQuart,AndÜw,ausgst," & _
     "BhFB,BhFE,BhFE1,BhFE2,f4202,f4206,f4210,f4237,Fanf,GebOr," & _
     "GOÄKatName,GOÄKatNr,IK,KtrAbrB,KVKs,KVKserg,lfdnr,lVorl,Mahnfrist,Nachname," & _
     "Pat_ID,PGeb,PGebErg,QAnf,QEnd,QS,QT,Quartal,s8000,s8100," & _
     "SchGr,statBehTage,statKlasse,statNuller,StByte,TMFNr,ÜbwV,ÜWNaN,ÜWNNr,ÜwText," & _
     "ÜWTit,ÜWVor,ÜWVsw,ÜWZiel,VKNr,Vorname,Weiterbeh) values"
 End If
 For i = 1 To UBound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rFa(i).AbrAr & "','" & rFa(i).AbrGb & "'," & rFa(i).absPos & ",'" & rFa(i).AdNam & "','" & rFa(i).AdOrt & "','" & rFa(i).AdPlz & _
   "','" & rFa(i).AdStr & "'," & datForm(rFa(i).AktZeit) & ",'" & rFa(i).altQuart & "','" & rFa(i).AndÜw & "'," & datForm(rFa(i).ausgst) & "," & datForm(rFa(i).BhFB) & _
   "," & datForm(rFa(i).BhFE) & "," & datForm(rFa(i).BhFE1) & "," & datForm(rFa(i).BhFE2) & ",'" & rFa(i).f4202 & "','" & rFa(i).f4206 & _
   "','" & rFa(i).f4210 & "','" & rFa(i).f4237 & "'," & datForm(rFa(i).Fanf) & ",'" & rFa(i).GebOr & "','" & rFa(i).GOÄKatName & "','" & rFa(i).GOÄKatNr & _
   "','" & rFa(i).IK & "','" & rFa(i).KtrAbrB & "','" & rFa(i).KVKs & "','" & rFa(i).KVKserg & "'," & rFa(i).lfdnr & "," & datForm(rFa(i).lVorl) & _
   ",'" & rFa(i).Mahnfrist & "','" & rFa(i).Nachname & "'," & rFa(i).Pat_ID & ",'" & rFa(i).PGeb & "','" & rFa(i).PGebErg & _
   "'," & datForm(rFa(i).QAnf) & "," & datForm(rFa(i).QEnd) & ",'" & rFa(i).QS & "','" & rFa(i).QT & "','" & rFa(i).Quartal & "','" & rFa(i).s8000 & _
   "','" & rFa(i).s8100 & "','" & rFa(i).SchGr & "'," & rFa(i).statBehTage & ",'" & rFa(i).statKlasse & "','" & rFa(i).statNuller & "'," & rFa(i).StByte & _
   ",'" & rFa(i).TMFNr & "','" & rFa(i).ÜbwV & "','" & rFa(i).ÜWNaN & "','" & rFa(i).ÜWNNr & "','" & rFa(i).ÜwText & "','" & rFa(i).ÜWTit & _
   "','" & rFa(i).ÜWVor & "','" & rFa(i).ÜWVsw & "','" & rFa(i).ÜWZiel & "','" & rFa(i).VKNr & "','" & rFa(i).Vorname & "','" & rFa(i).Weiterbeh & _
   "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AbrAr":  If SpMod(rFa(i).AbrAr, rsc!character_maximum_length, "faelle", "AbrAr") Then Exit Do
   Case "AbrGb":  If SpMod(rFa(i).AbrGb, rsc!character_maximum_length, "faelle", "AbrGb") Then Exit Do
   Case "AdNam":  If SpMod(rFa(i).AdNam, rsc!character_maximum_length, "faelle", "AdNam") Then Exit Do
   Case "AdOrt":  If SpMod(rFa(i).AdOrt, rsc!character_maximum_length, "faelle", "AdOrt") Then Exit Do
   Case "AdPlz":  If SpMod(rFa(i).AdPlz, rsc!character_maximum_length, "faelle", "AdPlz") Then Exit Do
   Case "AdStr":  If SpMod(rFa(i).AdStr, rsc!character_maximum_length, "faelle", "AdStr") Then Exit Do
   Case "altQuart":  If SpMod(rFa(i).altQuart, rsc!character_maximum_length, "faelle", "altQuart") Then Exit Do
   Case "AndÜw":  If SpMod(rFa(i).AndÜw, rsc!character_maximum_length, "faelle", "AndÜw") Then Exit Do
   Case "f4202":  If SpMod(rFa(i).f4202, rsc!character_maximum_length, "faelle", "f4202") Then Exit Do
   Case "f4206":  If SpMod(rFa(i).f4206, rsc!character_maximum_length, "faelle", "f4206") Then Exit Do
   Case "f4210":  If SpMod(rFa(i).f4210, rsc!character_maximum_length, "faelle", "f4210") Then Exit Do
   Case "f4237":  If SpMod(rFa(i).f4237, rsc!character_maximum_length, "faelle", "f4237") Then Exit Do
   Case "GebOr":  If SpMod(rFa(i).GebOr, rsc!character_maximum_length, "faelle", "GebOr") Then Exit Do
   Case "GOÄKatName":  If SpMod(rFa(i).GOÄKatName, rsc!character_maximum_length, "faelle", "GOÄKatName") Then Exit Do
   Case "GOÄKatNr":  If SpMod(rFa(i).GOÄKatNr, rsc!character_maximum_length, "faelle", "GOÄKatNr") Then Exit Do
   Case "IK":  If SpMod(rFa(i).IK, rsc!character_maximum_length, "faelle", "IK") Then Exit Do
   Case "KtrAbrB":  If SpMod(rFa(i).KtrAbrB, rsc!character_maximum_length, "faelle", "KtrAbrB") Then Exit Do
   Case "KVKs":  If SpMod(rFa(i).KVKs, rsc!character_maximum_length, "faelle", "KVKs") Then Exit Do
   Case "KVKserg":  If SpMod(rFa(i).KVKserg, rsc!character_maximum_length, "faelle", "KVKserg") Then Exit Do
   Case "Mahnfrist":  If SpMod(rFa(i).Mahnfrist, rsc!character_maximum_length, "faelle", "Mahnfrist") Then Exit Do
   Case "Nachname":  If SpMod(rFa(i).Nachname, rsc!character_maximum_length, "faelle", "Nachname") Then Exit Do
   Case "PGeb":  If SpMod(rFa(i).PGeb, rsc!character_maximum_length, "faelle", "PGeb") Then Exit Do
   Case "PGebErg":  If SpMod(rFa(i).PGebErg, rsc!character_maximum_length, "faelle", "PGebErg") Then Exit Do
   Case "QS":  If SpMod(rFa(i).QS, rsc!character_maximum_length, "faelle", "QS") Then Exit Do
   Case "QT":  If SpMod(rFa(i).QT, rsc!character_maximum_length, "faelle", "QT") Then Exit Do
   Case "Quartal":  If SpMod(rFa(i).Quartal, rsc!character_maximum_length, "faelle", "Quartal") Then Exit Do
   Case "s8000":  If SpMod(rFa(i).s8000, rsc!character_maximum_length, "faelle", "s8000") Then Exit Do
   Case "s8100":  If SpMod(rFa(i).s8100, rsc!character_maximum_length, "faelle", "s8100") Then Exit Do
   Case "SchGr":  If SpMod(rFa(i).SchGr, rsc!character_maximum_length, "faelle", "SchGr") Then Exit Do
   Case "statKlasse":  If SpMod(rFa(i).statKlasse, rsc!character_maximum_length, "faelle", "statKlasse") Then Exit Do
   Case "statNuller":  If SpMod(rFa(i).statNuller, rsc!character_maximum_length, "faelle", "statNuller") Then Exit Do
   Case "TMFNr":  If SpMod(rFa(i).TMFNr, rsc!character_maximum_length, "faelle", "TMFNr") Then Exit Do
   Case "ÜbwV":  If SpMod(rFa(i).ÜbwV, rsc!character_maximum_length, "faelle", "ÜbwV") Then Exit Do
   Case "ÜWNaN":  If SpMod(rFa(i).ÜWNaN, rsc!character_maximum_length, "faelle", "ÜWNaN") Then Exit Do
   Case "ÜWNNr":  If SpMod(rFa(i).ÜWNNr, rsc!character_maximum_length, "faelle", "ÜWNNr") Then Exit Do
   Case "ÜwText":  If SpMod(rFa(i).ÜwText, rsc!character_maximum_length, "faelle", "ÜwText") Then Exit Do
   Case "ÜWTit":  If SpMod(rFa(i).ÜWTit, rsc!character_maximum_length, "faelle", "ÜWTit") Then Exit Do
   Case "ÜWVor":  If SpMod(rFa(i).ÜWVor, rsc!character_maximum_length, "faelle", "ÜWVor") Then Exit Do
   Case "ÜWVsw":  If SpMod(rFa(i).ÜWVsw, rsc!character_maximum_length, "faelle", "ÜWVsw") Then Exit Do
   Case "ÜWZiel":  If SpMod(rFa(i).ÜWZiel, rsc!character_maximum_length, "faelle", "ÜWZiel") Then Exit Do
   Case "VKNr":  If SpMod(rFa(i).VKNr, rsc!character_maximum_length, "faelle", "VKNr") Then Exit Do
   Case "Vorname":  If SpMod(rFa(i).Vorname, rsc!character_maximum_length, "faelle", "Vorname") Then Exit Do
   Case "Weiterbeh":  If SpMod(rFa(i).Weiterbeh, rsc!character_maximum_length, "faelle", "Weiterbeh") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function auSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "au" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "au" & klz & " (absPos," & _
     "AktZeit,Beginn,Ende,FID,ICDs,Pat_ID,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "au" & klz & " (absPos," & _
     "AktZeit,Beginn,Ende,FID,ICDs,Pat_ID,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rAu(i).absPos & "," & datForm(rAu(i).AktZeit) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "'," & rAu(i).FID & ",'" & rAu(i).ICDs & _
   "'," & rAu(i).Pat_ID & "," & rAu(i).StByte & "," & datForm(rAu(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Beginn":  If SpMod(rAu(i).Beginn, rsc!character_maximum_length, "au", "Beginn") Then Exit Do
   Case "Ende":  If SpMod(rAu(i).Ende, rsc!character_maximum_length, "au", "Ende") Then Exit Do
   Case "ICDs":  If SpMod(rAu(i).ICDs, rsc!character_maximum_length, "au", "ICDs") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function briefeSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "briefe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "briefe" & klz & " (absPos," & _
     "AktZeit,Art,DokGroe,FID,Name,Pat_ID,Pfad,QS,QT,StByte," & _
     "Typ,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "briefe" & klz & " (absPos," & _
     "AktZeit,Art,DokGroe,FID,Name,Pat_ID,Pfad,QS,QT,StByte," & _
     "Typ,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rBr(i).absPos & "," & datForm(rBr(i).AktZeit) & ",'" & rBr(i).Art & "'," & rBr(i).DokGroe & "," & rBr(i).FID & ",'" & rBr(i).Name & _
   "'," & rBr(i).Pat_ID & ",'" & rBr(i).Pfad & "','" & rBr(i).QS & "','" & rBr(i).QT & "'," & rBr(i).StByte & ",'" & rBr(i).Typ & "'," & datForm(rBr(i).Zeitpunkt) & _
   ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rBr(i).Art, rsc!character_maximum_length, "briefe", "Art") Then Exit Do
   Case "Name":  If SpMod(rBr(i).Name, rsc!character_maximum_length, "briefe", "Name") Then Exit Do
   Case "Pfad":  If SpMod(rBr(i).Pfad, rsc!character_maximum_length, "briefe", "Pfad") Then Exit Do
   Case "QS":  If SpMod(rBr(i).QS, rsc!character_maximum_length, "briefe", "QS") Then Exit Do
   Case "QT":  If SpMod(rBr(i).QT, rsc!character_maximum_length, "briefe", "QT") Then Exit Do
   Case "Typ":  If SpMod(rBr(i).Typ, rsc!character_maximum_length, "briefe", "Typ") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function diagnosenSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "diagnosen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "diagnosen" & klz & " (absPos," & _
     "AktZeit,DiagAttr,DiagDatum,DiagSeite,DiagSicherheit,DiagText,FID,GesName,ICD,obDauer," & _
     "Pat_id,StByte) values"
 Else
  sql0 = "insert into " & kla & "diagnosen" & klz & " (absPos," & _
     "AktZeit,DiagAttr,DiagDatum,DiagSeite,DiagSicherheit,DiagText,FID,GesName,ICD,obDauer," & _
     "Pat_id,StByte) values"
 End If
 For i = 1 To UBound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDi(i).absPos & "," & datForm(rDi(i).AktZeit) & ",'" & rDi(i).DiagAttr & "'," & datForm(rDi(i).DiagDatum) & ",'" & rDi(i).DiagSeite & _
   "','" & rDi(i).DiagSicherheit & "','" & rDi(i).DiagText & "'," & rDi(i).FID & ",'" & rDi(i).GesName & "','" & rDi(i).ICD & "'," & CStr(CInt(rDi(i).obDauer)) & _
   "," & rDi(i).Pat_ID & "," & rDi(i).StByte & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DiagAttr":  If SpMod(rDi(i).DiagAttr, rsc!character_maximum_length, "diagnosen", "DiagAttr") Then Exit Do
   Case "DiagSeite":  If SpMod(rDi(i).DiagSeite, rsc!character_maximum_length, "diagnosen", "DiagSeite") Then Exit Do
   Case "DiagSicherheit":  If SpMod(rDi(i).DiagSicherheit, rsc!character_maximum_length, "diagnosen", "DiagSicherheit") Then Exit Do
   Case "DiagText":  If SpMod(rDi(i).DiagText, rsc!character_maximum_length, "diagnosen", "DiagText") Then Exit Do
   Case "GesName":  If SpMod(rDi(i).GesName, rsc!character_maximum_length, "diagnosen", "GesName") Then Exit Do
   Case "ICD":  If SpMod(rDi(i).ICD, rsc!character_maximum_length, "diagnosen", "ICD") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function dokumenteSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "dokumente" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "dokumente" & klz & " (absPos," & _
     "AktZeit,DokArt,DokGroe,DokName,DokPfad,FID,Pat_ID,QS,QT,Quelldatum," & _
     "StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "dokumente" & klz & " (absPos," & _
     "AktZeit,DokArt,DokGroe,DokName,DokPfad,FID,Pat_ID,QS,QT,Quelldatum," & _
     "StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDo(i).absPos & "," & datForm(rDo(i).AktZeit) & ",'" & rDo(i).DokArt & "'," & rDo(i).DokGroe & ",'" & rDo(i).DokName & "','" & rDo(i).DokPfad & _
   "'," & rDo(i).FID & "," & rDo(i).Pat_ID & ",'" & rDo(i).QS & "','" & rDo(i).QT & "'," & datForm(rDo(i).Quelldatum) & "," & rDo(i).StByte & _
   "," & datForm(rDo(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DokArt":  If SpMod(rDo(i).DokArt, rsc!character_maximum_length, "dokumente", "DokArt") Then Exit Do
   Case "DokName":  If SpMod(rDo(i).DokName, rsc!character_maximum_length, "dokumente", "DokName") Then Exit Do
   Case "DokPfad":  If SpMod(rDo(i).DokPfad, rsc!character_maximum_length, "dokumente", "DokPfad") Then Exit Do
   Case "QS":  If SpMod(rDo(i).QS, rsc!character_maximum_length, "dokumente", "QS") Then Exit Do
   Case "QT":  If SpMod(rDo(i).QT, rsc!character_maximum_length, "dokumente", "QT") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function eintraegeSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "eintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "eintraege" & klz & " (absPos," & _
     "AktZeit,Art,FID,Inhalt,Pat_ID,QS,QT,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "eintraege" & klz & " (absPos," & _
     "AktZeit,Art,FID,Inhalt,Pat_ID,QS,QT,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rEi(i).absPos & "," & datForm(rEi(i).AktZeit) & ",'" & rEi(i).Art & "'," & rEi(i).FID & ",'" & rEi(i).Inhalt & "'," & rEi(i).Pat_ID & _
   ",'" & rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & "," & datForm(rEi(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rEi(i).Art, rsc!character_maximum_length, "eintraege", "Art") Then Exit Do
   Case "Inhalt":  If SpMod(rEi(i).Inhalt, rsc!character_maximum_length, "eintraege", "Inhalt") Then Exit Do
   Case "QS":  If SpMod(rEi(i).QS, rsc!character_maximum_length, "eintraege", "QS") Then Exit Do
   Case "QT":  If SpMod(rEi(i).QT, rsc!character_maximum_length, "eintraege", "QT") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function forminhaltform_abkSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "forminhaltform_abk" & klz & " (Form_Abk," & _
     " values"
 Else
  sql0 = "insert into " & kla & "forminhaltform_abk" & klz & " (Form_Abk," & _
     " values"
 End If
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
  sql = sql0 & "('" & rFi(i).Form_Abk & "',"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFi1 = UBound(rFi)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFi(i).Form_Abk, rsc!character_maximum_length, "forminhaltform_abk", "Form_Abk") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function formulareSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "formulare" & klz & " (absPos," & _
     "AktZeit,Form_Abk,FormBez,FormID,FormVorl,StByte) values"
 Else
  sql0 = "insert into " & kla & "formulare" & klz & " (absPos," & _
     "AktZeit,Form_Abk,FormBez,FormID,FormVorl,StByte) values"
 End If
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFo(i).absPos & "," & datForm(rFo(i).AktZeit) & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "'," & rFo(i).FormID & ",'" & rFo(i).FormVorl & _
   "'," & rFo(i).StByte & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFo1 = UBound(rFo)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFo(i).Form_Abk, rsc!character_maximum_length, "formulare", "Form_Abk") Then Exit Do
   Case "FormBez":  If SpMod(rFo(i).FormBez, rsc!character_maximum_length, "formulare", "FormBez") Then Exit Do
   Case "FormVorl":  If SpMod(rFo(i).FormVorl, rsc!character_maximum_length, "formulare", "FormVorl") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function forminhkopfSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "forminhkopf" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "forminhkopf" & klz & " (AbsPos," & _
     "AktZeit,FID,FoID,Form_ID,Pat_ID,Satzart,Satzlänge,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "forminhkopf" & klz & " (AbsPos," & _
     "AktZeit,FID,FoID,Form_ID,Pat_ID,Satzart,Satzlänge,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFr(i).absPos & "," & datForm(rFr(i).AktZeit) & "," & rFr(i).FID & "," & rFr(i).FoID & "," & rFr(i).Form_ID & "," & rFr(i).Pat_ID & _
   ",'" & rFr(i).Satzart & "','" & rFr(i).Satzlänge & "'," & rFr(i).StByte & "," & datForm(rFr(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rFr(i).Satzart, rsc!character_maximum_length, "forminhkopf", "Satzart") Then Exit Do
   Case "Satzlänge":  If SpMod(rFr(i).Satzlänge, rsc!character_maximum_length, "forminhkopf", "Satzlänge") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function forminhfeldSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "forminhfeld" & klz & " (FeldInhVW," & _
     "FeldNr,FeldVW,FoID,Nr) values"
 Else
  sql0 = "insert into " & kla & "forminhfeld" & klz & " (FeldInhVW," & _
     "FeldNr,FeldVW,FoID,Nr) values"
 End If
 For i = 1 To UBound(rFm)
'  rFm(i).AktZeit = now()
  sql = sql0 & "(" & rFm(i).FeldInhVW & "," & CStr(CInt(rFm(i).FeldNr)) & "," & rFm(i).FeldVW & "," & rFm(i).FoID & "," & CStr(CInt(rFm(i).Nr)) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function kheinweisSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "kheinweis" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "kheinweis" & klz & " (absPos," & _
     "AktZeit,Diagnose,FID,Pat_ID,StByte,ZeitPunkt,Ziel) values"
 Else
  sql0 = "insert into " & kla & "kheinweis" & klz & " (absPos," & _
     "AktZeit,Diagnose,FID,Pat_ID,StByte,ZeitPunkt,Ziel) values"
 End If
 For i = 1 To UBound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKh(i).absPos & "," & datForm(rKh(i).AktZeit) & ",'" & rKh(i).Diagnose & "'," & rKh(i).FID & "," & rKh(i).Pat_ID & "," & rKh(i).StByte & _
   "," & datForm(rKh(i).Zeitpunkt) & ",'" & rKh(i).Ziel & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Diagnose":  If SpMod(rKh(i).Diagnose, rsc!character_maximum_length, "kheinweis", "Diagnose") Then Exit Do
   Case "Ziel":  If SpMod(rKh(i).Ziel, rsc!character_maximum_length, "kheinweis", "Ziel") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function lbanforderungenSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "lbanforderungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "lbanforderungen" & klz & " (absPos," & _
     "AktZeit,AnfText,FID,Pat_ID,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "lbanforderungen" & klz & " (absPos," & _
     "AktZeit,AnfText,FID,Pat_ID,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLb(i).absPos & "," & datForm(rLb(i).AktZeit) & ",'" & rLb(i).AnfText & "'," & rLb(i).FID & "," & rLb(i).Pat_ID & "," & rLb(i).StByte & _
   "," & datForm(rLb(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AnfText":  If SpMod(rLb(i).AnfText, rsc!character_maximum_length, "lbanforderungen", "AnfText") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function laborneuSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "laborneu" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborneu" & klz & " (Abkü," & _
     "AbsPos,AktZeit,Einheit,FertigStGrad,FID,KommentarVW,LangtextVW,Pat_ID,Refnr,StByte," & _
     "Wert,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "laborneu" & klz & " (Abkü," & _
     "AbsPos,AktZeit,Einheit,FertigStGrad,FID,KommentarVW,LangtextVW,Pat_ID,Refnr,StByte," & _
     "Wert,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rLa(i).Abkü & "'," & rLa(i).absPos & "," & datForm(rLa(i).AktZeit) & ",'" & rLa(i).Einheit & "','" & rLa(i).FertigStGrad & "'," & rLa(i).FID & _
   "," & rLa(i).KommentarVW & "," & rLa(i).LangtextVW & "," & rLa(i).Pat_ID & "," & rLa(i).RefNr & "," & rLa(i).StByte & ",'" & rLa(i).Wert & _
   "'," & datForm(rLa(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLa(i).Abkü, rsc!character_maximum_length, "laborneu", "Abkü") Then Exit Do
   Case "Einheit":  If SpMod(rLa(i).Einheit, rsc!character_maximum_length, "laborneu", "Einheit") Then Exit Do
   Case "FertigStGrad":  If SpMod(rLa(i).FertigStGrad, rsc!character_maximum_length, "laborneu", "FertigStGrad") Then Exit Do
   Case "Wert":  If SpMod(rLa(i).Wert, rsc!character_maximum_length, "laborneu", "Wert") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function leistungenSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "leistungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "leistungen" & klz & " (absPos," & _
     "AktZeit,f5002,f5005,f5006,f5009,f5015,f5016,f5021,f5026,FID," & _
     "Leistung,Med,Pat_ID,QS,QT,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "leistungen" & klz & " (absPos," & _
     "AktZeit,f5002,f5005,f5006,f5009,f5015,f5016,f5021,f5026,FID," & _
     "Leistung,Med,Pat_ID,QS,QT,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLe(i).absPos & "," & datForm(rLe(i).AktZeit) & ",'" & rLe(i).f5002 & "','" & rLe(i).f5005 & "','" & rLe(i).f5006 & "','" & rLe(i).f5009 & _
   "','" & rLe(i).f5015 & "','" & rLe(i).f5016 & "','" & rLe(i).f5021 & "','" & rLe(i).f5026 & "'," & rLe(i).FID & ",'" & rLe(i).Leistung & _
   "','" & rLe(i).Med & "'," & rLe(i).Pat_ID & ",'" & rLe(i).QS & "','" & rLe(i).QT & "'," & rLe(i).StByte & "," & datForm(rLe(i).Zeitpunkt) & _
   ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "f5002":  If SpMod(rLe(i).f5002, rsc!character_maximum_length, "leistungen", "f5002") Then Exit Do
   Case "f5005":  If SpMod(rLe(i).f5005, rsc!character_maximum_length, "leistungen", "f5005") Then Exit Do
   Case "f5006":  If SpMod(rLe(i).f5006, rsc!character_maximum_length, "leistungen", "f5006") Then Exit Do
   Case "f5009":  If SpMod(rLe(i).f5009, rsc!character_maximum_length, "leistungen", "f5009") Then Exit Do
   Case "f5015":  If SpMod(rLe(i).f5015, rsc!character_maximum_length, "leistungen", "f5015") Then Exit Do
   Case "f5016":  If SpMod(rLe(i).f5016, rsc!character_maximum_length, "leistungen", "f5016") Then Exit Do
   Case "f5021":  If SpMod(rLe(i).f5021, rsc!character_maximum_length, "leistungen", "f5021") Then Exit Do
   Case "f5026":  If SpMod(rLe(i).f5026, rsc!character_maximum_length, "leistungen", "f5026") Then Exit Do
   Case "Leistung":  If SpMod(rLe(i).Leistung, rsc!character_maximum_length, "leistungen", "Leistung") Then Exit Do
   Case "Med":  If SpMod(rLe(i).Med, rsc!character_maximum_length, "leistungen", "Med") Then Exit Do
   Case "QS":  If SpMod(rLe(i).QS, rsc!character_maximum_length, "leistungen", "QS") Then Exit Do
   Case "QT":  If SpMod(rLe(i).QT, rsc!character_maximum_length, "leistungen", "QT") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function medplanSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "medplan" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "medplan" & klz & " (ab," & _
     "AbsPos,AktZeit,bBed,Bemerkung,Datum,FeldNr,FID,MedAnfang,Medikament,mi," & _
     "mo,MPNr,nm,Pat_ID,StByte,ZeitPunkt,zn) values"
 Else
  sql0 = "insert into " & kla & "medplan" & klz & " (ab," & _
     "AbsPos,AktZeit,bBed,Bemerkung,Datum,FeldNr,FID,MedAnfang,Medikament,mi," & _
     "mo,MPNr,nm,Pat_ID,StByte,ZeitPunkt,zn) values"
 End If
 For i = 1 To UBound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rMe(i).ab & "'," & rMe(i).absPos & "," & datForm(rMe(i).AktZeit) & "," & CStr(CInt(rMe(i).bBed)) & ",'" & rMe(i).Bemerkung & "'," & datForm(rMe(i).Datum) & _
   "," & CStr(CInt(rMe(i).FeldNr)) & "," & rMe(i).FID & ",'" & rMe(i).MedAnfang & "','" & rMe(i).Medikament & "','" & rMe(i).mi & _
   "','" & rMe(i).mo & "'," & rMe(i).MPNr & ",'" & rMe(i).nm & "'," & rMe(i).Pat_ID & "," & rMe(i).StByte & "," & datForm(rMe(i).Zeitpunkt) & _
   ",'" & rMe(i).zn & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "ab":  If SpMod(rMe(i).ab, rsc!character_maximum_length, "medplan", "ab") Then Exit Do
   Case "Bemerkung":  If SpMod(rMe(i).Bemerkung, rsc!character_maximum_length, "medplan", "Bemerkung") Then Exit Do
   Case "MedAnfang":  If SpMod(rMe(i).MedAnfang, rsc!character_maximum_length, "medplan", "MedAnfang") Then Exit Do
   Case "Medikament":  If SpMod(rMe(i).Medikament, rsc!character_maximum_length, "medplan", "Medikament") Then Exit Do
   Case "mi":  If SpMod(rMe(i).mi, rsc!character_maximum_length, "medplan", "mi") Then Exit Do
   Case "mo":  If SpMod(rMe(i).mo, rsc!character_maximum_length, "medplan", "mo") Then Exit Do
   Case "nm":  If SpMod(rMe(i).nm, rsc!character_maximum_length, "medplan", "nm") Then Exit Do
   Case "zn":  If SpMod(rMe(i).zn, rsc!character_maximum_length, "medplan", "zn") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function rezepteintraegeSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "rezepteintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "rezepteintraege" & klz & " (absPos," & _
     "AktZeit,FID,Medikament,Pat_ID,PZN,QS,QT,Rezept,Rezeptklasse,StByte," & _
     "ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "rezepteintraege" & klz & " (absPos," & _
     "AktZeit,FID,Medikament,Pat_ID,PZN,QS,QT,Rezept,Rezeptklasse,StByte," & _
     "ZeitPunkt) values"
 End If
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRe(i).absPos & "," & datForm(rRe(i).AktZeit) & "," & rRe(i).FID & ",'" & rRe(i).Medikament & "'," & rRe(i).Pat_ID & ",'" & rRe(i).PZN & _
   "','" & rRe(i).QS & "','" & rRe(i).QT & "','" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "'," & rRe(i).StByte & "," & datForm(rRe(i).Zeitpunkt) & _
   ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Medikament":  If SpMod(rRe(i).Medikament, rsc!character_maximum_length, "rezepteintraege", "Medikament") Then Exit Do
   Case "PZN":  If SpMod(rRe(i).PZN, rsc!character_maximum_length, "rezepteintraege", "PZN") Then Exit Do
   Case "QS":  If SpMod(rRe(i).QS, rsc!character_maximum_length, "rezepteintraege", "QS") Then Exit Do
   Case "QT":  If SpMod(rRe(i).QT, rsc!character_maximum_length, "rezepteintraege", "QT") Then Exit Do
   Case "Rezept":  If SpMod(rRe(i).Rezept, rsc!character_maximum_length, "rezepteintraege", "Rezept") Then Exit Do
   Case "Rezeptklasse":  If SpMod(rRe(i).Rezeptklasse, rsc!character_maximum_length, "rezepteintraege", "Rezeptklasse") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function rrSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "rr" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "rr" & klz & " (absPos," & _
     "AktZeit,FID,Pat_ID,RR,StByte,ZeitPunkt) values"
 Else
  sql0 = "insert into " & kla & "rr" & klz & " (absPos," & _
     "AktZeit,FID,Pat_ID,RR,StByte,ZeitPunkt) values"
 End If
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRr(i).absPos & "," & datForm(rRr(i).AktZeit) & "," & rRr(i).FID & "," & rRr(i).Pat_ID & ",'" & rRr(i).RR & "'," & rRr(i).StByte & _
   "," & datForm(rRr(i).Zeitpunkt) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "RR":  If SpMod(rRr(i).RR, rsc!character_maximum_length, "rr", "RR") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function kvnrueSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "kvnrue" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "kvnrue" & klz & " (absPos," & _
     "AktZeit,KVNr,Pat_ID,StByte) values"
 Else
  sql0 = "insert into " & kla & "kvnrue" & klz & " (absPos," & _
     "AktZeit,KVNr,Pat_ID,StByte) values"
 End If
 For i = 1 To UBound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKv(i).absPos & "," & datForm(rKv(i).AktZeit) & ",'" & rKv(i).KVNr & "'," & rKv(i).Pat_ID & "," & rKv(i).StByte & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "KVNr":  If SpMod(rKv(i).KVNr, rsc!character_maximum_length, "kvnrue", "KVNr") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function unbekannte_kennungenSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "unbekannte kennungen" & klz & " (absPos," & _
     "Kennung,StByte) values"
 Else
  sql0 = "insert into " & kla & "unbekannte kennungen" & klz & " (absPos," & _
     "Kennung,StByte) values"
 End If
 For i = rUn1 + 1 To UBound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rUn(i).absPos & ",'" & rUn(i).Kennung & "'," & rUn(i).StByte & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rUn1 = UBound(rUn)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Kennung":  If SpMod(rUn(i).Kennung, rsc!character_maximum_length, "unbekannte kennungen", "Kennung") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 If obMySQL Then Resume Next Else Resume
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

Public Function laborxsaetzeSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxsaetze" & klz & " (Arztname," & _
     "Arztnr,DatID,Erstellungsdatum,Gesamtlänge,KBVPrüfnr,Kundenarztnr,Labor,OrtLabor,OrtPraxis,PLZLabor," & _
     "PLZPraxis,Satzart,Satzlänge,SatzlängeSchluss,StraßeLabor,StraßePraxis,VersionSatzb,Zeichensatz) values"
 Else
  sql0 = "insert into " & kla & "laborxsaetze" & klz & " (Arztname," & _
     "Arztnr,DatID,Erstellungsdatum,Gesamtlänge,KBVPrüfnr,Kundenarztnr,Labor,OrtLabor,OrtPraxis,PLZLabor," & _
     "PLZPraxis,Satzart,Satzlänge,SatzlängeSchluss,StraßeLabor,StraßePraxis,VersionSatzb,Zeichensatz) values"
 End If
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
  sql = sql0 & "('" & rLs(i).Arztname & "','" & rLs(i).Arztnr & "'," & rLs(i).DatID & ",'" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "','" & rLs(i).KBVPrüfnr & _
   "','" & rLs(i).Kundenarztnr & "','" & rLs(i).Labor & "','" & rLs(i).OrtLabor & "','" & rLs(i).OrtPraxis & "','" & rLs(i).PLZLabor & _
   "','" & rLs(i).PLZPraxis & "','" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).StraßeLabor & _
   "','" & rLs(i).StraßePraxis & "','" & rLs(i).VersionSatzb & "','" & rLs(i).Zeichensatz & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Arztname":  If SpMod(rLs(i).Arztname, rsc!character_maximum_length, "laborxsaetze", "Arztname") Then Exit Do
   Case "Arztnr":  If SpMod(rLs(i).Arztnr, rsc!character_maximum_length, "laborxsaetze", "Arztnr") Then Exit Do
   Case "Erstellungsdatum":  If SpMod(rLs(i).Erstellungsdatum, rsc!character_maximum_length, "laborxsaetze", "Erstellungsdatum") Then Exit Do
   Case "Gesamtlänge":  If SpMod(rLs(i).Gesamtlänge, rsc!character_maximum_length, "laborxsaetze", "Gesamtlänge") Then Exit Do
   Case "KBVPrüfnr":  If SpMod(rLs(i).KBVPrüfnr, rsc!character_maximum_length, "laborxsaetze", "KBVPrüfnr") Then Exit Do
   Case "Kundenarztnr":  If SpMod(rLs(i).Kundenarztnr, rsc!character_maximum_length, "laborxsaetze", "Kundenarztnr") Then Exit Do
   Case "Labor":  If SpMod(rLs(i).Labor, rsc!character_maximum_length, "laborxsaetze", "Labor") Then Exit Do
   Case "OrtLabor":  If SpMod(rLs(i).OrtLabor, rsc!character_maximum_length, "laborxsaetze", "OrtLabor") Then Exit Do
   Case "OrtPraxis":  If SpMod(rLs(i).OrtPraxis, rsc!character_maximum_length, "laborxsaetze", "OrtPraxis") Then Exit Do
   Case "PLZLabor":  If SpMod(rLs(i).PLZLabor, rsc!character_maximum_length, "laborxsaetze", "PLZLabor") Then Exit Do
   Case "PLZPraxis":  If SpMod(rLs(i).PLZPraxis, rsc!character_maximum_length, "laborxsaetze", "PLZPraxis") Then Exit Do
   Case "Satzart":  If SpMod(rLs(i).Satzart, rsc!character_maximum_length, "laborxsaetze", "Satzart") Then Exit Do
   Case "Satzlänge":  If SpMod(rLs(i).Satzlänge, rsc!character_maximum_length, "laborxsaetze", "Satzlänge") Then Exit Do
   Case "SatzlängeSchluss":  If SpMod(rLs(i).SatzlängeSchluss, rsc!character_maximum_length, "laborxsaetze", "SatzlängeSchluss") Then Exit Do
   Case "StraßeLabor":  If SpMod(rLs(i).StraßeLabor, rsc!character_maximum_length, "laborxsaetze", "StraßeLabor") Then Exit Do
   Case "StraßePraxis":  If SpMod(rLs(i).StraßePraxis, rsc!character_maximum_length, "laborxsaetze", "StraßePraxis") Then Exit Do
   Case "VersionSatzb":  If SpMod(rLs(i).VersionSatzb, rsc!character_maximum_length, "laborxsaetze", "VersionSatzb") Then Exit Do
   Case "Zeichensatz":  If SpMod(rLs(i).Zeichensatz, rsc!character_maximum_length, "laborxsaetze", "Zeichensatz") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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

Public Function laborxeingelSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxeingel" & klz & " (Name," & _
     "Pfad,Zp) values"
 Else
  sql0 = "insert into " & kla & "laborxeingel" & klz & " (Name," & _
     "Pfad,Zp) values"
 End If
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
  sql = sql0 & "('" & rLg(i).Name & "','" & rLg(i).Pfad & "'," & datForm(rLg(i).Zp) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Name":  If SpMod(rLg(i).Name, rsc!character_maximum_length, "laborxeingel", "Name") Then Exit Do
   Case "Pfad":  If SpMod(rLg(i).Pfad, rsc!character_maximum_length, "laborxeingel", "Pfad") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxus" & klz & " (Abrechnungstyp," & _
     "AfN,Auftragsnummer,Auftragsschlüssel,AuftrHinw,BefArt,Berichtsdatum,DatID,Eingang,GebDat,GebüOrd," & _
     "Geschlecht,LWerte,Nachname,Pat_id,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idErwVN,Pat_idErwVNG,Pat_idLaborNeu," & _
     "Pat_idUrsp,Patienteninformation,Satzart,SatzID,Satzlänge,Titel,verglichen,Vorname,ZdiP,ZdüP," & _
     "ZeitpunktLaborneu) values"
 Else
  sql0 = "insert into " & kla & "laborxus" & klz & " (Abrechnungstyp," & _
     "AfN,Auftragsnummer,Auftragsschlüssel,AuftrHinw,BefArt,Berichtsdatum,DatID,Eingang,GebDat,GebüOrd," & _
     "Geschlecht,LWerte,Nachname,Pat_id,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idErwVN,Pat_idErwVNG,Pat_idLaborNeu," & _
     "Pat_idUrsp,Patienteninformation,Satzart,SatzID,Satzlänge,Titel,verglichen,Vorname,ZdiP,ZdüP," & _
     "ZeitpunktLaborneu) values"
 End If
 For i = j To j
'  rLu(i).AktZeit = now()
  sql = sql0 & "('" & rLu(i).Abrechnungstyp & "'," & CStr(CInt(rLu(i).AfN)) & ",'" & rLu(i).Auftragsnummer & "','" & rLu(i).Auftragsschlüssel & "','" & rLu(i).AuftrHinw & _
   "','" & rLu(i).BefArt & "','" & rLu(i).Berichtsdatum & "'," & rLu(i).DatID & "," & datForm(rLu(i).Eingang) & ",'" & rLu(i).GebDat & _
   "','" & rLu(i).GebüOrd & "','" & rLu(i).Geschlecht & "','" & rLu(i).LWerte & "','" & rLu(i).Nachname & "'," & rLu(i).Pat_ID & ",'" & rLu(i).Pat_idErwG & _
   "','" & rLu(i).Pat_idErwGB & "','" & rLu(i).Pat_idErwGL & "','" & rLu(i).Pat_idErwVN & "','" & rLu(i).Pat_idErwVNG & "','" & rLu(i).Pat_idLaborNeu & _
   "','" & rLu(i).Pat_idUrsp & "','" & rLu(i).Patienteninformation & "','" & rLu(i).Satzart & "'," & rLu(i).SatzID & ",'" & rLu(i).Satzlänge & _
   "','" & rLu(i).Titel & "'," & datForm(rLu(i).verglichen) & ",'" & rLu(i).Vorname & "'," & rLu(i).ZdiP & "," & CStr(CInt(rLu(i).ZdüP)) & _
   "," & datForm(rLu(i).ZeitpunktLaborneu) & ")"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abrechnungstyp":  If SpMod(rLu(i).Abrechnungstyp, rsc!character_maximum_length, "laborxus", "Abrechnungstyp") Then Exit Do
   Case "Auftragsnummer":  If SpMod(rLu(i).Auftragsnummer, rsc!character_maximum_length, "laborxus", "Auftragsnummer") Then Exit Do
   Case "Auftragsschlüssel":  If SpMod(rLu(i).Auftragsschlüssel, rsc!character_maximum_length, "laborxus", "Auftragsschlüssel") Then Exit Do
   Case "AuftrHinw":  If SpMod(rLu(i).AuftrHinw, rsc!character_maximum_length, "laborxus", "AuftrHinw") Then Exit Do
   Case "BefArt":  If SpMod(rLu(i).BefArt, rsc!character_maximum_length, "laborxus", "BefArt") Then Exit Do
   Case "Berichtsdatum":  If SpMod(rLu(i).Berichtsdatum, rsc!character_maximum_length, "laborxus", "Berichtsdatum") Then Exit Do
   Case "GebDat":  If SpMod(rLu(i).GebDat, rsc!character_maximum_length, "laborxus", "GebDat") Then Exit Do
   Case "GebüOrd":  If SpMod(rLu(i).GebüOrd, rsc!character_maximum_length, "laborxus", "GebüOrd") Then Exit Do
   Case "Geschlecht":  If SpMod(rLu(i).Geschlecht, rsc!character_maximum_length, "laborxus", "Geschlecht") Then Exit Do
   Case "LWerte":  If SpMod(rLu(i).LWerte, rsc!character_maximum_length, "laborxus", "LWerte") Then Exit Do
   Case "Nachname":  If SpMod(rLu(i).Nachname, rsc!character_maximum_length, "laborxus", "Nachname") Then Exit Do
   Case "Pat_idErwG":  If SpMod(rLu(i).Pat_idErwG, rsc!character_maximum_length, "laborxus", "Pat_idErwG") Then Exit Do
   Case "Pat_idErwGB":  If SpMod(rLu(i).Pat_idErwGB, rsc!character_maximum_length, "laborxus", "Pat_idErwGB") Then Exit Do
   Case "Pat_idErwGL":  If SpMod(rLu(i).Pat_idErwGL, rsc!character_maximum_length, "laborxus", "Pat_idErwGL") Then Exit Do
   Case "Pat_idErwVN":  If SpMod(rLu(i).Pat_idErwVN, rsc!character_maximum_length, "laborxus", "Pat_idErwVN") Then Exit Do
   Case "Pat_idErwVNG":  If SpMod(rLu(i).Pat_idErwVNG, rsc!character_maximum_length, "laborxus", "Pat_idErwVNG") Then Exit Do
   Case "Pat_idLaborNeu":  If SpMod(rLu(i).Pat_idLaborNeu, rsc!character_maximum_length, "laborxus", "Pat_idLaborNeu") Then Exit Do
   Case "Pat_idUrsp":  If SpMod(rLu(i).Pat_idUrsp, rsc!character_maximum_length, "laborxus", "Pat_idUrsp") Then Exit Do
   Case "Patienteninformation":  If SpMod(rLu(i).Patienteninformation, rsc!character_maximum_length, "laborxus", "Patienteninformation") Then Exit Do
   Case "Satzart":  If SpMod(rLu(i).Satzart, rsc!character_maximum_length, "laborxus", "Satzart") Then Exit Do
   Case "Satzlänge":  If SpMod(rLu(i).Satzlänge, rsc!character_maximum_length, "laborxus", "Satzlänge") Then Exit Do
   Case "Titel":  If SpMod(rLu(i).Titel, rsc!character_maximum_length, "laborxus", "Titel") Then Exit Do
   Case "Vorname":  If SpMod(rLu(i).Vorname, rsc!character_maximum_length, "laborxus", "Vorname") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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

Public Function laborxbaktSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxbakt" & klz & " (AbnDat," & _
     "Erklärung,Keimzahl,Kommentar,KuQu,QSpez,Quelle,RefNr,Verf) values"
 Else
  sql0 = "insert into " & kla & "laborxbakt" & klz & " (AbnDat," & _
     "Erklärung,Keimzahl,Kommentar,KuQu,QSpez,Quelle,RefNr,Verf) values"
 End If
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
  sql = sql0 & "(" & datForm(rLo(i).AbnDat) & ",'" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "','" & rLo(i).Kommentar & "','" & rLo(i).KuQu & "','" & rLo(i).QSpez & _
   "','" & rLo(i).Quelle & "'," & rLo(i).RefNr & ",'" & rLo(i).Verf & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Erklärung":  If SpMod(rLo(i).Erklärung, rsc!character_maximum_length, "laborxbakt", "Erklärung") Then Exit Do
   Case "Keimzahl":  If SpMod(rLo(i).Keimzahl, rsc!character_maximum_length, "laborxbakt", "Keimzahl") Then Exit Do
   Case "Kommentar":  If SpMod(rLo(i).Kommentar, rsc!character_maximum_length, "laborxbakt", "Kommentar") Then Exit Do
   Case "KuQu":  If SpMod(rLo(i).KuQu, rsc!character_maximum_length, "laborxbakt", "KuQu") Then Exit Do
   Case "QSpez":  If SpMod(rLo(i).QSpez, rsc!character_maximum_length, "laborxbakt", "QSpez") Then Exit Do
   Case "Quelle":  If SpMod(rLo(i).Quelle, rsc!character_maximum_length, "laborxbakt", "Quelle") Then Exit Do
   Case "Verf":  If SpMod(rLo(i).Verf, rsc!character_maximum_length, "laborxbakt", "Verf") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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

Public Function laborxwertSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxwert" & klz & " (Abkü," & _
     "AbnDat,AuftrHinw,Einheit,Erklärung,Grenzwerti,Kommentar,Langname,Normbereich,NormO,NormU," & _
     "QSpez,Quelle,RefNr,Teststatus,Wert) values"
 Else
  sql0 = "insert into " & kla & "laborxwert" & klz & " (Abkü," & _
     "AbnDat,AuftrHinw,Einheit,Erklärung,Grenzwerti,Kommentar,Langname,Normbereich,NormO,NormU," & _
     "QSpez,Quelle,RefNr,Teststatus,Wert) values"
 End If
 For i = 1 To UBound(rLw)
'  rLw(i).AktZeit = now()
  sql = sql0 & "('" & rLw(i).Abkü & "'," & datForm(rLw(i).AbnDat) & ",'" & rLw(i).AuftrHinw & "','" & rLw(i).Einheit & "','" & rLw(i).Erklärung & "','" & rLw(i).Grenzwerti & _
   "','" & rLw(i).Kommentar & "','" & rLw(i).Langname & "','" & rLw(i).Normbereich & "','" & rLw(i).NormO & "','" & rLw(i).NormU & _
   "','" & rLw(i).QSpez & "','" & rLw(i).Quelle & "'," & rLw(i).RefNr & ",'" & rLw(i).Teststatus & "','" & rLw(i).Wert & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLw(i).Abkü, rsc!character_maximum_length, "laborxwert", "Abkü") Then Exit Do
   Case "AuftrHinw":  If SpMod(rLw(i).AuftrHinw, rsc!character_maximum_length, "laborxwert", "AuftrHinw") Then Exit Do
   Case "Einheit":  If SpMod(rLw(i).Einheit, rsc!character_maximum_length, "laborxwert", "Einheit") Then Exit Do
   Case "Erklärung":  If SpMod(rLw(i).Erklärung, rsc!character_maximum_length, "laborxwert", "Erklärung") Then Exit Do
   Case "Grenzwerti":  If SpMod(rLw(i).Grenzwerti, rsc!character_maximum_length, "laborxwert", "Grenzwerti") Then Exit Do
   Case "Kommentar":  If SpMod(rLw(i).Kommentar, rsc!character_maximum_length, "laborxwert", "Kommentar") Then Exit Do
   Case "Langname":  If SpMod(rLw(i).Langname, rsc!character_maximum_length, "laborxwert", "Langname") Then Exit Do
   Case "Normbereich":  If SpMod(rLw(i).Normbereich, rsc!character_maximum_length, "laborxwert", "Normbereich") Then Exit Do
   Case "NormO":  If SpMod(rLw(i).NormO, rsc!character_maximum_length, "laborxwert", "NormO") Then Exit Do
   Case "NormU":  If SpMod(rLw(i).NormU, rsc!character_maximum_length, "laborxwert", "NormU") Then Exit Do
   Case "QSpez":  If SpMod(rLw(i).QSpez, rsc!character_maximum_length, "laborxwert", "QSpez") Then Exit Do
   Case "Quelle":  If SpMod(rLw(i).Quelle, rsc!character_maximum_length, "laborxwert", "Quelle") Then Exit Do
   Case "Teststatus":  If SpMod(rLw(i).Teststatus, rsc!character_maximum_length, "laborxwert", "Teststatus") Then Exit Do
   Case "Wert":  If SpMod(rLw(i).Wert, rsc!character_maximum_length, "laborxwert", "Wert") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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

Public Function laborxleistSpeichern()
 Dim i%, sql0$, rAf&
 On Error GoTo fehler
 If obMySQL Then
  sql0 = "insert ignore into " & kla & "laborxleist" & klz & " (Abkü," & _
     "Anzahl,EBM,goä,RefNr,Verf) values"
 Else
  sql0 = "insert into " & kla & "laborxleist" & klz & " (Abkü," & _
     "Anzahl,EBM,goä,RefNr,Verf) values"
 End If
 For i = 1 To UBound(rLL)
'  rLL(i).AktZeit = now()
  sql = sql0 & "('" & rLL(i).Abkü & "','" & rLL(i).Anzahl & "','" & rLL(i).EBM & "','" & rLL(i).goä & "'," & rLL(i).RefNr & ",'" & rLL(i).Verf & "')"
  Call DBCn.Execute(sql, rAf) ', , adAsyncExecute)
  If obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLL(i).Abkü, rsc!character_maximum_length, "laborxleist", "Abkü") Then Exit Do
   Case "Anzahl":  If SpMod(rLL(i).Anzahl, rsc!character_maximum_length, "laborxleist", "Anzahl") Then Exit Do
   Case "EBM":  If SpMod(rLL(i).EBM, rsc!character_maximum_length, "laborxleist", "EBM") Then Exit Do
   Case "goä":  If SpMod(rLL(i).goä, rsc!character_maximum_length, "laborxleist", "goä") Then Exit Do
   Case "Verf":  If SpMod(rLL(i).Verf, rsc!character_maximum_length, "laborxleist", "Verf") Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If obMySQL Then Resume Next Else Resume
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

Public Function doSpeichern()
 On Error GoTo fehler
 Call namenSpeichern
 Call faelleSpeichern
   If Not obMySQL Then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   End If
 Call auSpeichern
 Call briefeSpeichern
 Call diagnosenSpeichern
 Call dokumenteSpeichern
 Call eintraegeSpeichern
 Call forminhaltform_abkSpeichern
   If Not obMySQL Then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   End If
 Call formulareSpeichern
 Call forminhkopfSpeichern
 Call forminhfeldSpeichern
 Call kheinweisSpeichern
 Call lbanforderungenSpeichern
 Call laborneuSpeichern
 Call leistungenSpeichern
 Call medplanSpeichern
 Call rezepteintraegeSpeichern
 Call rrSpeichern
 Call kvnrueSpeichern
 Call unbekannte_kennungenSpeichern
 Exit Function
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

Function SpMod%(SpVal$, BishL&, TName$, SpName$)
 If Len(SpVal) > BishL Then
  Call DBCn.Execute("Alter Table " & kla & TName & klz & " " & IIf(obMySQL, "MODIFY", "ALTER") & " Column " & SpName & " " & IIf(obMySQL, "VARCHAR", "TEXT") & "(" & Len(SpVal) & ")")
  SpMod = True
 End If
End Function ' SpMod
