# Statistik -> Anrufliste

Dokumentation des Menüpunkts „Anrufliste" (Stand: 26.08.2026, Commits `b37070d`/`9aaa515`).

## 1. Aufruf

Menü: **Statistik → Anrufliste** (direkt unter „suchTel").

- Menüeintrag: `Lese5.frm:773` (`Begin VB.Menu Anrufliste`)
- Klick-Handler: `Lese5.frm:4262` (`Private Sub Anrufliste_Click()`) → ruft `ProgStart` und
  `doAnrufliste(Me)`
- Eigentliche Logik: `Haupt.bas:3245`, Funktion `doAnrufliste`

Beim Aufruf fragt eine InputBox „Anzahl der letzten Anrufe:" (Default **30**). Leere Eingabe
oder Abbruch → ebenfalls 30. Die Eingabe wird über `CLng(Val(anz))` in eine Zahl gewandelt
(nicht-numerische Eingabe ergibt 0 → leere Liste, kein Absturz).

Die Liste wird per `TabAusgeb` angezeigt: einmal als Textdatei, einmal programmintern —
**absteigend** sortiert (neuester Anruf zuerst), im Gegensatz zum CLI-Befehl `anrliste -listt`
(Linux, `~root/anrliste`), der dieselbe Rohdatenbasis aufsteigend zeigt.

## 2. Datenquelle und Spalten

Basis ist `faxeinp.anrufe` (dieselbe Tabelle, die auch `anrliste -listt` verwendet), begrenzt
über `ORDER BY eind DESC LIMIT <Anzahl>` (Auswahl der letzten N Anrufe), danach zur Anzeige
nochmal nach `eind` sortiert (die einzige verlässliche chronologische Reihenfolge — `Datum` ist
nur minutengenau, mehrere Anrufe können denselben `Datum`-Wert haben).

Angezeigte Spalten:

| Spalte | Herkunft | Bedeutung |
|---|---|---|
| Datum | `anrufe.Datum` | Zeitpunkt (minutengenau) |
| Typ | `anrufe.Typ` (gemappt) | `1`→`->` (angenommen), `2`→`->\|` (Anrufbeantworter), `3`→`<-` (ausgehend), `9`→`->>` (verpasst), `11`→`fx<-` (Fax), sonst Rohwert |
| Name | `anrufe.Name` | **Fritzbox-eigene** Anrufer-Erkennung (eigenes/Online-Telefonbuch der Fritzbox) — unabhängig von der Patienten-/Praxis-Zuordnung unten. Meist leer; „Unbekannt" wenn die Fritzbox online nachgesehen aber nichts gefunden hat. Nur ca. 3,7 % aller Datensätze haben hier überhaupt einen Wert. |
| Rufnummer | `anrufe.Rufnummer` | Rohnummer des Gesprächspartners. Bei Typ `11` (Fax) steht hier ausnahmsweise `SIP: <eigene Nebenstelle>` statt einer echten Nummer (siehe Abschnitt 3). |
| Nebenstelle | `anrufe.Nebenstelle` | internes Fritzbox-Gerät/Port (z. B. „ISDN Gerät", „Pana-AMT1", „FonFax", „HylaFax") |
| Dauer | `anrufe.Dauer` | Gesprächsdauer in Minuten (Dezimal) |
| Pat-ID | siehe Abschnitt 3 | Patienten-ID, `"Praxis"` bei Arzt-/Praxistreffer, sonst leer |
| Klarname | siehe Abschnitt 3 | aufgelöster Name des Gesprächspartners, sonst leer |

**Bewusst weggelassen** gegenüber `anrliste -listt`: `EigeneNr` (immer `SIP: <Nebenstelle>`,
rein technisch), `angerNr`, `Nummerntyp` (immer `sip`), `Port`, `Id`, `eind` — diese liefern für
die Patientenzuordnung keinen Mehrwert.

## 3. Zuordnung zu Patient/Arzt/Praxis (`Pat-ID`, `Klarname`)

Für jede Zeile wird die `Rufnummer` gegen drei Quellen abgeglichen, in dieser Reihenfolge
(erster Treffer gewinnt, per `COALESCE`):

1. **Patient** — Tabelle `quelle.namen`, Felder `PrivatTel`, `PrivatTel_2`, `DienstTel`,
   `PrivatMobil`. `Pat-ID` = `Pat_id`, `Klarname` = `gesname(Pat_id)`.
2. **Überweiser/Hausarzt (Cache)** — Tabelle `quelle.aktlue`, Feld `telefon`. `Pat-ID` =
   `"Praxis"`, `Klarname` = `"Dr. <Nachname>, <Vorname>, <Ort>"`.
3. **Arzt/Praxis (Stammdaten)** — Tabelle `quelle.earzt` (Feld `FTelefon`), verknüpft über
   `FExtpraxisnr = epraxis.FSurogat` mit `quelle.epraxis` (Feld `FBezeichnung`, Praxisname).
   `Pat-ID` = `"Praxis"`, `Klarname` = `"Dr. <FNachname>, <FVorname>, <FBezeichnung>"`.

Kein Treffer in keiner der drei Quellen → beide Spalten leer (`''`, **nicht** `NULL` — siehe
Abschnitt 5).

„Dr." wird bei Quelle 2 und 3 **pauschal auf Verdacht** vorangestellt (nicht anhand eines
echten Titelfelds geprüft) — auf Nutzerwunsch, um Ärzte/Praxen optisch von Patienten
abzuheben.

### Nummernabgleich (robust gegen Formatierungs- und Mehrfachnummern-Chaos)

Telefonnummern sind in der Datenbank uneinheitlich gespeichert (Leerzeichen, `/`, `-`, teils
mehrere Nummern kommagetrennt in einem Feld, z. B. `earzt.FTelefon`). Der Abgleich:

1. Aus der eingehenden `Rufnummer` werden per `REGEXP_REPLACE(...,'[^0-9]','')` alle
   Nicht-Ziffern entfernt, dann die **letzten 8 Ziffern** genommen (`RN`). Das gleicht
   Präfix-Unterschiede (Vorwahl/Landeskennung) aus.
2. `Rufnummer LIKE 'SIP%'` (Fax-Sonderfall, s. o.) oder weniger als 6 Ziffern → `RN = ''`,
   kein Abgleichversuch (verhindert Fehltreffer bei zu kurzen/unechten „Nummern" — konkret
   beobachtet: ein Test-/Dummy-Patient mit dem gespeicherten „Telefon" `616380`, was zufällig
   der eigenen SIP-Nebenstelle entspricht).
3. Je Zielfeld ebenfalls `REGEXP_REPLACE(...,'[^0-9]','')`, und `RN` wird per `LOCATE()` als
   **Teilstring** darin gesucht (nicht als exakter Endvergleich) — das funktioniert auch, wenn
   im Zielfeld mehrere Nummern hintereinander stehen (`CONCAT_WS('|', ...)` bei `namen`
   fasst die vier Telefonfelder für die Suche zu einem String zusammen).

## 4. Performance

Gemessen (Live-DB, `namen` 33.145 / `aktlue` 4.116 / `earzt` 1.665 / `epraxis` 1.362 Zeilen):
komplette Abfrage bei `LIMIT 30` ca. 15–20 ms, bei `LIMIT 300` weiterhin unter 20 ms. Kein
Index nutzbar (Vergleich über `REGEXP_REPLACE`/`LOCATE`), aber durch `COALESCE`-Kurzschluss
(sobald eine Quelle greift, werden weitere Quellen für diese Zeile nicht mehr abgefragt) und
das vorgeschaltete `LIMIT` in der inneren Abfrage bleibt die Laufzeit unabhängig von der
Gesamtgröße von `anrufe` (aktuell ~838.000 Zeilen).

Bekannte, bewusst nicht behobene Ineffizienz: `Pat-ID` und `Klarname` werten für einen
Patiententreffer dieselbe `namen`-Bedingung zweimal aus (einmal für die ID, einmal für
`gesname()`). Bei den gemessenen Laufzeiten nicht relevant; könnte per `LATERAL`-Join auf
einen Durchlauf reduziert werden, falls die Liste künftig routinemäßig mit sehr großen
Stückzahlen (>1000) aufgerufen wird.

## 5. `NULL` vs. leerer Wert

`Pat-ID` und `Klarname` sind beide mit einem abschließenden `COALESCE(..., '')` abgesichert.
Grund: Ohne diesen Fallback lieferte SQL bei fehlendem Treffer `NULL`, was in der
programminternen `TabAusgeb`-Anzeige als störender Text „Null" auftauchte (in der
Textdatei-Ausgabe dagegen nicht) — behoben, indem die Abfrage selbst nie mehr `NULL`
zurückgibt.

## 6. Nebenbefund: VB6-255-Controls-Limit auf `Lese5.frm`

Beim ersten Einbau dieses Menüpunkts stürzte die VB6-IDE bei **jedem** F5/Kompilieren ab —
unabhängig von Name, Position oder Inhalt des neuen Menüeintrags. Ursache: Das MDI-Formular
`Lese` lag exakt bei **255 Controls** (Menüpunkte zählen dabei als Controls mit), der von VB6
pro Formular dokumentierten Obergrenze. Jeder zusätzliche Menüpunkt war der 256., was die IDE
zum Absturz brachte (kein regulärer Compile-Fehler-Dialog, sondern harter Crash).

Behoben durch Entfernen von 12 längst unsichtbaren (`Visible = 0`) Menüpunkten, deren
Klick-Handler entweder vollständig auskommentiert oder in `#If False Then...#End If`
eingeschlossen waren (u. a. mehrere `u:\Anamnese\Quelle.mdb`-Altlasten, zwei am 10.10.22 vom
Nutzer selbst auskommentierte Testfunktionen). Formular liegt seither bei 243 Controls
(12 Slots Puffer). Bei künftigen neuen Menüpunkten auf `Lese5.frm`: Control-Anzahl im Auge
behalten (`grep -c "Begin VB\." Lese5.frm` als grobe Näherung, siehe auch die genauere
Python-Zählung in der Commit-Historie).
