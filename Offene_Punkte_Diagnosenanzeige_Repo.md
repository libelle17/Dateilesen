# Diagnosenanzeige (Arztbrief, DMP-Benachrichtigung): Stand und offene Punkte

_Repository-Fassung von `Offene_Punkte_Diagnosenanzeige.md` (Original nur lokal): Diagnosen zu einzelnen Patientennummern entfernt bzw. die Nummer weggelassen._

Stand: 20.9.2026 (spätabends, nach Ihrer Tabellenänderung und J70-Korrektur). Programmcode geändert: bisher nur `MODiagnosen` in `vonMO.bas` (20.9., siehe Abschnitt 4). Alle Auswertungen liefen lesend gegen `medoff` und `quelle`.

## 1. Entschieden

| Thema | Entscheidung |
|---|---|
| Ausgangswünsche | gelten weiter: mehrere Z.n. PCI/Schlaganfall einzeln; Z.n. Nephropathie/Nikotinabusus nur einmal; Z.n. X nicht neben aktivem X (auch Diabetes); bei Diabetes mellitus nur Pathologische Glukosetoleranz anzeigen |
| Historische Diagnosen (`behgrund.FStatus=3`, 1.111 Zeilen) | beim Übertrag weglassen |
| Alles übrige | erst bei der Anzeige (`DiagString`/`MachDiagnosen`) |
| Abgeschlossene Diagnosen (Status 4) | mit aktivem Gegenstück ausnahmslos weglassen; ohne Gegenstück nach ICD-Entscheidungstabelle (n / j / je / j mit Z.n.) |
| Ausschluss-Diagnosen | abgeschlossene erscheinen nie; aktive erscheinen weiter als „Ausschluss …" |
| V.a. | zunächst dieselben Entscheidungen (j/je/n) wie für gesichert |
| Text-Präfix „gesichert" | entfällt in Kombination mit „Z.n." und/oder „V.a."; beim Zusammenstellen werden gesichert/Z.n./V.a. aus dem Text entfernt und das Label aus der Entscheidung gesetzt |
| Neue Spalte | `quelle.diagnosen.MOStatus` (= `behgrund.FStatus`): **angelegt** (alle Werte NULL), `typen.bas` **regeneriert** (Type `diagnosen` hat `MOStatus`); `diagnosenSpeichern` schreibt sie (Ihre IDE bzw. `typen.bas`), `MODiagnosen` befüllt sie seit 20.9. |
| Reihenfolge der Diagnosen | erst nach Umsetzung des Übrigen |
| Nicht gelistete ICD | `je` (damit nichts unter den Tisch fällt; im Brief löschbar) |
| T78 (Allergie) | unspezifiziert („Allergie“, „Allergie, nicht näher bezeichnet“) j; mit Spezifikation („Allergie, Furosemid“, „Allergie gegen …“) je |
| R29 | „Wadenschmerz“ im Text je, sonst j (z. B. Sturzneigung) |
| Z95 | „Stent“, „PTCA“, „PCI“ oder „Bypass“ im Text je (mehrere möglich), sonst j |
| J32 | „chron“ nicht im Text j, sonst je |
| M17 | „TEP“ im Text mit Z.n. (jz), sonst j |
| K52 | n (bestätigt: „K25" war ein Tippfehler für K52; K25 bleibt je) |
| Erläuterung (`FErlaeuterung`) | **zählt beim Abgleich mit**: die Erläuterung der abgeschlossenen Diagnose muss (wortweise) in der Erläuterung der aktiven vorkommen; leer ist immer erfüllt |
| Kernbegriff-Liste v2 | die Spalte „Vorschlag“ (54 ja, 6 nein) gilt als Entscheidung |
| J70 → I70 | in der Auswertung ersetzt und von Ihnen in `medoff` korrigiert (`behgrund`, `ltag`, `medfavorit`; geprüft: kein J70 mehr, I70 = 1.382 bzw. 462); in `quelle` stehen bis zur Neuübertragung noch 86 J70-Zeilen |

## 2. Fertige Dateien (Projektordner)

- `Diagnose_Entscheidungstabelle.csv` / `.sql` – Entscheidungstabelle (521 Regeln); das SQL ist ausgeführt, die Tabelle heißt `quelle.diagentscheid` (bei Änderungen der CSV: Tabelle leeren und Skript erneut ausführen, sonst doppelte Zeilen)
- `ICD_Liste_abgeschlossene_ohne_Gegenstueck_v4.csv` – ICD-3-Liste mit Ihren Entscheidungen (J70 in I70, K52 = n)
- `ICD_Liste_volle_ICD_v1 bearbeitet.csv` – Vollcode-Entscheidungen (ersetzt L89, I25, Z95, J32, I20)
- `Kernbegriff_Zuordnungen_abgeschlossen_zu_aktiv_v2.csv` – neue Regeln, mit Vorbelegung (ja/nein)
- `Aktive_Zn_gesichert_Paare_1_nach_ICD3.csv`, `_2_nach_voller_ICD.csv`, `_3_nach_ICD_und_Text.csv`
- Unterordner `Beispiel_Diagnosestrings/` – je Beispielpatient der Paare-Liste 1 eine Datei `<ICD-3>_<Pat_id>.txt` (313 Dateien, 278 Patienten) mit dem heutigen Diagnosestring (Python-Nachbau von `DiagString`/`MachDiagnosen`: dieselbe SQL-Abfrage gegen `quelle`, `obBrief=True`, `VorDat` = alle, `dmseit` aus `anamnesebogen`). Nicht mit dem laufenden Programm verglichen; ein Teil der Paare ist durch die bestehende SQL-Regel (gleiche ICD und gleicher Text nach Entfernen von Z.n./gesichert) schon heute ausgeblendet.
- `Abkuerzungen_und_Synonyme_Diagnosetexte.csv`
- Außerhalb des Repos: `/DATA/eigene Dateien/Diagnosen_Auswertung_vertraulich/` (die beiden Listen mit Patientennamen; nach Ihrer Prüfung löschbar)

## 3. Offen – Ihre Entscheidung oder Ihr Eingriff nötig

1. ~~Neue Spalte in `quelle.diagnosen`~~ – erledigt (von Ihnen ausgeführt), `typen.bas` regeneriert. Nächster Schritt: Code in `vonMO.bas`/`typen.bas` (siehe Abschnitt 4), sobald Sie es wünschen.
2. **Aktive Z.n./gesichert-Paare** (983 Patienten, 1.373 Paare in gleicher ICD-3-Stelle): 380 mit gleichem Text, 31 mit gleichem Kernbegriff, 403 gleiche ICD anderer Text, 559 nur gleiche ICD-3. Regel für „Z.n. weglassen, wenn gleiche Diagnose gesichert" ist Ihr Wunsch; für die drei Ungleich-Gruppen fehlt noch die Festlegung, was als „gleiche Diagnose" gilt (die Listen sind dafür gedacht).
3. **Z.n. V.a.** – umgesetzt als `DiagSicherheit = Z` + Textpräfix „V.a.“ + `MOStatus = 4` (siehe 6.); „V.a. Z.n." (32 Zeilen) bleibt Text.
4. ~~J70 in `medoff`~~ – erledigt; `quelle` bekommt die Korrektur mit der Neuübertragung.
5. ~~Offene Zeilen der Entscheidungstabelle~~ – erledigt mit Ihren Klarstellungen (siehe Abschnitt 1); `Diagnose_Entscheidungstabelle.csv/.sql` neu erzeugt (521 Regeln, keine offenen Fälle).
6. **Kernbegriff-Liste v2**: Sie haben in der Spalte „Vorschlag Claude“ 6 meiner „nein“ auf „ja“ gesetzt (jetzt 54 ja, 6 nein: T14.9 Wunde, T78.4 Allergie, -1 Allergie, H40.9 Glaukom, I80.3 Varikophlebitis, M54.5 Lumbago); die Spalte „Ihre Entscheidung“ ist leer. Ich werte die Spalte „Vorschlag“ als Ihre Entscheidung.
7. ~~Erläuterung als Spezifitätsmerkmal~~ – entschieden: zählt mit (siehe Abschnitt 1 und 7).
8. **DMP-Benachrichtigung an Hausärzte**: welcher Programmpfad erzeugt sie? Aufrufer von `DiagString`: Formular.bas:8001 (Brief, `obBrief`), Formular.bas:1636, Laufzettelneu.bas:1954, `dynDiag`; von `MachDiagnosen(...Sort:=True)`: Importiere.bas:6341. Noch nicht geprüft, welcher davon die DMP-Post erzeugt.

## 4. Offen – Umsetzung (danach, in dieser Reihenfolge)

1. ~~`MODiagnosen` (vonMO.bas)~~ – **erledigt am 20.9.** (noch nicht getestet, nicht committet): Status 3 wird nicht mehr übertragen; abgeschlossene V.a. werden mitgenommen (als Z mit Textpräfix „V.a.“); abgeschlossene Ausschluss-Diagnosen (Klasse 4, Status 4) werden nicht übertragen (sonst wären sie nach der Umsetzung auf Z nicht mehr erkennbar); `FStatus` wird in die Abfrage aufgenommen und als `MOStatus` gespeichert. Prüfen: mit einem Testpatienten übertragen und in `quelle.diagnosen` `MOStatus` ansehen.
2. ~~`diagnosenSpeichern`~~ – enthält `MOStatus` in Spalten- und Wertliste (typen.bas). Hinweis: Werte aus anderen Quellen (Turbomed-Import) bleiben 0, nicht NULL; beides bedeutet „kein MO-Status“. Achtung: `doConAnal` (Importiere.bas ~8345) legt Tabellen mit `CREATE TABLE IF NOT EXISTS … LIKE diagnosen` an; ältere Tabellen aus früheren Läufen (auch `_modia`, `_tmdia`) haben die Spalte nicht und würden beim INSERT einen Fehler „Unknown column MOStatus“ liefern → vor erneutem Gebrauch löschen oder um die Spalte ergänzen.
3. ~~Anzeige~~ – **umgesetzt am 21.9. (nicht kompiliert, nicht getestet)**, siehe Abschnitt 12. Ursprüngliche Aufgabe: Matching-Regeln (Präfixe, Datum einseitig, ICD „-1" als Platzhalter, Abkürzungen/Synonyme, Kernbegriff einseitig, Seite als Menge aus Text + Flag), Entscheidungstabelle, Label „Z.n."/„V.a."/„Z.n. V.a.", Präfix-Bereinigung, aktive Z.n./gesichert-Regeln. Die alte SQL-Regel in `DiagString` (Stand 22.2.26) wird dabei ersetzt.
4. Bestand nachziehen: `quelle` hinkt `medoff` hinterher; bis zur Neuübertragung (`turichtdiag`) ist `MOStatus = NULL` (= alter Stand) – die Anzeige muss NULL wie bisher behandeln.
5. Andere Verbraucher von `DiagSicherheit = 'Z'` prüfen (z. B. Haupt.bas:3575 DFS-Statistik, Formular.bas:9994, Views `diagview`, `diageview`, `diagmed`, `anakt` – diese Views verwenden keine `*`, die neue Spalte stört nicht).
6. Danach: Reihenfolge der Diagnosen (`diagreihe`/`diagg1`). Vorab erledigt (20.9.): `diagreihe` um I70.2 und I70.3 ergänzt (Gruppe 2, Rang 60, wie I70.20); `_modia` und `_tmdia` um `MOStatus` erweitert (weitere `<Tbnm>dia`-Tabellen gibt es in `quelle` derzeit nicht; `quelle4.diagnosen` ist leer und wurde nicht angefasst).

## 5. J70 → I70 und die neue Spalte – erledigt

**J70 in `medoff`** – von Ihnen ausgeführt und lesend nachgeprüft (Stand vorher): `behgrund.FIcdcode` 86 Zeilen (J70.2 ×78, J70.3 ×8), `ltag.FIcdcode` 86 Zeilen, `medfavorit.FIcdcode` 9 Zeilen (Favoritenliste – sonst wird der Fehler neu erzeugt). `vgs_diagnosen`, `vgs_formulare`, `vgs_briefe` sind **Views**. `ltag` und `behgrund` sind konsistent (alle 86 per Join gleich). Wichtig: `MODiagnosen` verknüpft `ltag.FIcdcode = behgrund.FIcdcode`; beide müssen **im selben Schritt** geändert werden, z. B. als eine Mehrtabellen-Anweisung:

```sql
UPDATE ltag l JOIN behgrund b ON l.FBehgrundnr = b.FSurogat
SET l.FIcdcode = CONCAT('I', SUBSTRING(l.FIcdcode, 2)), b.FIcdcode = CONCAT('I', SUBSTRING(b.FIcdcode, 2))
WHERE b.FIcdcode LIKE 'J70%';
UPDATE medfavorit SET FIcdcode = CONCAT('I', SUBSTRING(FIcdcode, 2)) WHERE FIcdcode LIKE 'J70%';
```
Vorher Sicherung; MO ist die Produktivdatenbank. Nachprüfung: `ltag` hat zwei Zeilen (Eintragsarten 13/14, ein Patient) mit ICD-Feld `I344959000` zu einem I70.3-Eintrag; das war schon vorher so und hat mit J70 nichts zu tun. `ltag` hat 13 Mio. Zeilen, die Join-Version nutzt die kleine Tabelle `behgrund` als Ausgangspunkt. In `quelle.diagnosen` habe ich die 86 J70-Zeilen am 20.9. direkt auf I70.2/I70.3 geändert (in einer Transaktion, danach 0 × J70). Offen: `diagreihe` kennt J70.2/J70.3 (Gruppe 17), aber nicht I70.2/I70.3 (nur I70.2-, I70.20 …, Gruppe 2/60): die 86 Zeilen sortieren jetzt ans Ende, bis Zeilen für I70.2 und I70.3 in `diagreihe` ergänzt werden.

**Neue Spalte** (Anweisung für Sie, z. B. im Direktfenster mit `! mariadb quelle -e "…"`):
```sql
ALTER TABLE diagnosen ADD COLUMN IF NOT EXISTS MOStatus TINYINT UNSIGNED NULL DEFAULT NULL
  COMMENT 'medoff behgrund.FStatus: 1 akut, 2 anamnestisch, 4 abgeschlossen, 5 Dauer (3 historisch wird nicht uebertragen); NULL = vor Einfuehrung uebertragen'
  AFTER KFdFA, ALGORITHM=INSTANT;
```
(rückgängig: `ALTER TABLE diagnosen DROP COLUMN MOStatus;`)

## 6. Z.n. V.a. – Darstellung in `quelle.diagnosen`

Umgesetzt (20.9.) in `MODiagnosen`: abgeschlossene V.a. → `DiagSicherheit = Z`, Text mit Präfix „V.a. “, `MOStatus = 4` (dieselbe Darstellung wie die migrierten Turbomed-Einträge; bestehende Verbraucher von `V` sehen sie dadurch nicht als aktiven Verdacht). Die Anzeige leitet die Gewissheit aus dem Präfix ab. Ist-Stand in `medoff`: Text „Z.n. V.a. …" kommt **nicht** vor. Abgeschlossene V.a. gibt es in zwei Formen: (a) Klasse V.a. + Status 4 (2.802 Zeilen; werden bisher nicht übertragen), (b) Klasse gesichert + Text-Präfix „V.a." (Status 4: 4.764, Status 5: 299, Status 1: 7) – (b) wird bisher als `DiagSicherheit = Z` übertragen und deshalb schon heute als „Z.n. V.a. …" angezeigt. „V.a. Z.n." kommt 32-mal vor (Status 4).

Vorschlag: `DiagSicherheit` = Gewissheit (G/V/A), `MOStatus` = Zeit. „Z.n. V.a. X" = `V` + `MOStatus 4`; „V.a. X" = `V` + Status 1/5; „Z.n. X" = `G` + `MOStatus 4`. Für (b) leitet die Anzeige die Gewissheit aus dem Text-Präfix „V.a." ab (Klasse gesichert ist dort ein Migrationsartefakt). „V.a. Z.n. X" bleibt Text; die Präfixbereinigung darf ein „Z.n." **hinter** einem „V.a." nicht entfernen.

## 7. Erläuterung (`FErlaeuterung`) – Beispiele

Das Feld ist die kurze Zusatzangabe zur Diagnose (in `DiagAttr`, im Brief in Klammern hinter dem Text). Es ist nur bei 5.369 von 147.325 Dauer-, 282 von 114.149 abgeschlossenen Zeilen gefüllt. Bisher zählt es beim Abgleich nicht mit. Wo abgeschlossene und aktive Diagnose (gleiche ICD, gleicher Text) sich nur in der Erläuterung unterscheiden:

- **395 Fälle**: abgeschlossen ohne, aktiv mit Erläuterung („Nikotinabusus" ↔ „Nikotinabusus (bis 2021)"). Aktiv ist spezifischer, der Treffer bleibt unverändert.
- **28 Fälle**: abgeschlossen **mit**, aktiv ohne Erläuterung („Arterielle Hypertonie (mit hypertensiver Entgleisung 5/23)" ↔ „Arterielle Hypertonie"; „Diabetes mellitus Typ 2 (Z.n. Entgleisung unter Cortisontherapie)"; „Nikotinabusus (2023: 1–2 Zig/d)"). Würde die Erläuterung zählen, gälten diese als **kein** Gegenstück, die abgeschlossene Diagnose wäre wieder sichtbar.
- **17 Fälle**: beide mit, aber verschieden („Bronchitis (rez., zul. 1/24)" ↔ „(6/23, mit Antibiose)"; „Aortenklappenstenose (…, li'ventr. Hypert…)" ↔ ohne diesen Zusatz). Ebenfalls kein Treffer.

Unterschied also nur 45 Fälle. **Entscheidung (20.9.): die Erläuterung zählt mit** (abgeschlossen ⊆ aktiv, wortweise; leer immer erfüllt). Die ICD-Listen wurden noch ohne diese Regel gerechnet; die Zahlen steigen dadurch geringfügig.

## 8. Zur Abkürzungs-/Synonymliste (früher Punkt 16)

Die Datei `Abkuerzungen_und_Synonyme_Diagnosetexte.csv` legt fest, welche Schreibweisen beim **Vergleich** als gleich gelten (z. B. „mult.Komp" = „multiplen Komplikationen"; „Hypertonus" = „arterielle Hypertonie" = „Bluthochdruck"). Die Spalte „Vorkommen" zeigt, wie oft die Schreibweise tatsächlich in den Diagnosetexten steht. Formen mit 0 stören nicht. Der Text auf dem Bildschirm oder im Brief bleibt unverändert. Ihre Aufgabe wäre nur, die Liste gelegentlich zu überfliegen und Fehleinträge oder fehlende Schreibweisen zu melden. Die Pipeline liest die Datei noch nicht ein; im späteren VB6-Code wäre es eine Tabelle.

## 9. Nebenbefunde (ergänzt)

1. **`MachDiagnosen` vergleicht `DSic(j) = "g"`** (Klein-g, Importiere.bas:6829 und 6838). In der Datenbank steht `G`, und `Option Compare` ist nicht gesetzt (Binärvergleich). Der Zweig „gesichert vor V/Z bevorzugen" greift damit für `G` nie. Bei gleichem Text und gleicher ICD wird das zweite Vorkommen unabhängig von der Klasse entfernt – ein Z.n. kann ein gesichert verdrängen. Mit ein Grund für die Z.n./gesichert-Probleme; einfache Korrektur möglich.
2. **ICD-„-1"-Einträge** (Turbomed ohne ICD): 1.194 abgeschlossene Gruppen, davon 842 über den Platzhalter einer aktiven Diagnose zugeordnet; für den Rest gilt die Regel n (Ihre Entscheidung).
3. **ICD-Text-Ungereimtheiten** außer J70 (z. B. D68 mit „Harnwegsinfektion", J44 mit „Polyglobulie"): systematische Liste auf Wunsch.
4. **Datumsangaben**: berücksichtigt sind `TT.MM.JJJJ`, `MM/JJ`, `JJJJ` und Wörter „seit/bis/ab/vor/am/ED" davor. Nicht erkannt: Quartale („Q3/2023"), Monatsnamen („Mai 2019"), relative Angaben („vor 3 Jahren"), Zeiträume in Worten. Die Formulierung „Vier Muster" in meiner früheren Zusammenstellung war ungenau.
5. **Regex-Falle bei `medoff`**: der Server läuft mit `NO_BACKSLASH_ESCAPES`. REGEXP-Muster in SQL, das der VB6-Code gegen `MOCon` schickt, dürfen keine Backslashes enthalten (stattdessen `[.]`, `[(]`).
6. **Alter Bericht** `Diagnosen_Kombis_Uebersicht.md` (Z.n./gesichert-Kombinationen nach Turbomed-Altdaten in `quelle`) entspricht dem Stand vor den Regeln oben und ist nur noch als Überblick brauchbar.
7. **Aktive Ausschluss-Zeilen** (Status 5, Klasse 4: 1.407) und **ICD-Häufungen bei Seiten-Flags** (L89 zählt jetzt je Seite) beeinflussen Zähl-Zahlen in den Listen.
8. `Creates.sql` im Projektordner enthält nur eine alte `quelle2`-Definition von `diagnosen`; sie ist nicht maßgeblich.

## 10. Neue Anzeige – Referenzlogik (Python, Stand 20.9.)

Dateien: `Referenz_Python/neuanzeige.py`, `normalisierung.py`, `seed3.py` (Abkürzungen/Synonyme); Ergebnis für die 313 Beispielpatienten in `Beispiel_Diagnosestrings_neu/` (gleiche Dateinamen wie in `Beispiel_Diagnosestrings/`, dort der alte Stand). Die Referenz arbeitet mit Zeilen, wie sie der geänderte `MODiagnosen` erzeugt (aus `medoff` simuliert), nicht mit den alten `quelle`-Zeilen; ein Teil der Unterschiede kommt daher vom Stand von `quelle`.

Regeln der Referenz:
1. **Aktive Zeilen** (MOStatus 1, 2, 5): identische Kopien zusammenfassen. Aktive Nicht-Z.n.: die bisherige Zusammenfassung (gleiche ICD-3-Stelle, gleiche Gewissheit, gleiche ersten 17 Zeichen) bleibt. Aktive Z.n.: weglassen, wenn dieselbe Diagnose aktiv gesichert vorliegt (Abgleichregeln aus Abschnitt 1); sonst nach Entscheidungstabelle: je = alle einzeln, sonst einmal je ICD-3-Stelle.
2. **Abgeschlossene Zeilen** (MOStatus 4): Kopien zusammenfassen; weglassen bei aktivem Gegenstück; weglassen, wenn eine spezifischere abgeschlossene Zeile derselben Diagnose vorliegt (z. B. „Wagner 1“ ohne Seite gegen „Wagner 1 bds.“); dann Entscheidungstabelle: n = weg; j/jz = einmal je ICD-3-Stelle und Gewissheit (jüngste Zeile); je = jede einzeln. Label: j „gesichert“ ohne Präfix (es sei denn, der Text ist selbst „Z.n.“), je/jz „Z.n.“, bei V.a. „V.a.“ bzw. „Z.n. V.a.“.
3. **Text**: führende gesichert/Z.n./V.a./ausgeschl. werden entfernt und das Label neu gesetzt („V.a. Z.n.“ bleibt). Die Umformungen aus `MachDiagnosen` (Diabetes, L89-Wagner, E66 …), die Reihenfolge (`diagreihe`) und „(q)“ bleiben; gleiche Anzeigezeile derselben ICD-3-Stelle erscheint nur einmal.
4. Geplanter **Fallback**: Patienten ohne `MOStatus` (NULL/0, noch nicht neu übertragen) behalten die alte Logik unverändert.

Ergebnis an den Beispielen: 9.226 → 8.839 Zeilen (152 Dateien kürzer, 127 gleich, 34 länger). Kürzer wird es vor allem durch Wegfall der „-1“-Zeilen und der Quartalskopien (z. B. Pat. 45: 49 → 23 Zeilen); länger, wo abgeschlossene V.a. nun mitkommen oder je-Diagnosen einzeln stehen.

Offen für die VB-Umsetzung: (A) Normalisierung in VB (die Abgleichregeln brauchen Lookbehind und Umlaute in Wortgrenzen, was `VBScript.RegExp` nicht kann) oder (B) als MariaDB-Funktionen in `quelle`, die `DiagString` abfragt (dann lassen sich die Funktionen hier gegen die Referenz prüfen).

## 11. Stand 20./21.9.: Übertrag geprüft, Funktionen in `quelle`, Paare-Liste v2

- **Testübertrag Pat. 155** (nach dem Neuladen der Quelldateien in der IDE): 27 Zeilen, davon 4 mit `MOStatus 4` (darunter ein V.a. als `Z` mit Präfix) und 23 mit `MOStatus 5`; keine Ausschluss- und keine historische Zeile. Der erste Versuch (29 Zeilen, `MOStatus 0`) lag an der IDE: sie hatte `vonMO.bas` schon geöffnet und meine Änderung auf der Platte nicht übernommen (Merkregel: nach Änderungen an Quelldateien außerhalb der IDE Projekt neu laden).
- **Variante B umgesetzt:** in `quelle` liegen die Tabelle `diagnormal` (55 Regeln = Abkürzungen/Synonyme/Rauschen, Spalte `Reihenfolge` = Anwendungsreihenfolge, `Muster` = PCRE-Ausdruck auf Kleinschreibung ohne Bindestriche) und die Funktionen `diag_clean`, `diag_suffix`, `diag_strip`, `diag_s`, `diag_nodate`, `diag_dates`, `diag_noside`, `diag_sides`, `diag_b0`, `diag_b1`, `diag_norm1`, `diag_core`, `diag_rest`, `diag_tokens`, `diag_info`. `diag_info(text, seite, erlaeuterung)` liefert in einem Aufruf, getrennt durch `Chr(31)`: s, dates, b0, b1, core, rest, sides (L/R/LR), erl. Geprüft gegen die Python-Referenz an 23.768 eindeutigen Diagnosetexten: 0 Abweichungen; Laufzeit ca. 1,3 ms je Text. Skripte: `Referenz_Python/diagnormal.sql`, `diag_funktionen.sql`. Änderungen an Abkürzungen/Synonymen: Zeile in `diagnormal` ändern (Reihenfolge beachten, PCRE-Syntax).
- **Aktive Z.n./gesichert-Paare mit ICD-Kürzeln:** `Aktive_Zn_gesichert_Paare_1_nach_ICD3_v2.csv` (154 ICD-3-Zeilen, 1.373 Paare; nur 88 Paare haben in beiden Texten ein Jahr, sonst gilt das Eintragsdatum). Kürzel (Ihre Definition): **g** gesichert bleibt, **z** Z.n. bleibt, **j** die jüngere bleibt, **b** beide bleiben, **a** die gesicherte und alle Z.n. bleiben, **Zahl** ist die gesicherte mindestens so viele Jahre jünger als die ältere (Z.n.), bleiben beide, sonst die jüngere. Arbeitsdefinition für mehrere Diagnosen je ICD-3 (bitte bestätigen): g = alle Z.n. weg; z = alle gesicherten weg; j = nur die jüngste Zeile; b = jüngste gesicherte und jüngstes Z.n.; a = jüngste gesicherte und alle Z.n.; Zahl = jüngste gesicherte gegen jüngstes Z.n.; Abstand ≥ Zahl → beide, sonst nur die jüngere. Datum je Zeile: höchstes Jahr im Text (auch „8/16“), sonst Eintragsdatum.

## 12. VB-Umsetzung der neuen Anzeige (21.9., nicht kompiliert und nicht getestet)

**Geänderte/neue Dateien**
- `DiagAnzeige.bas` (neu, reines ASCII, in `DateiLese.vbp` als `Module=DiagAnzeige` eingetragen): die neue Logik. Einstieg `DiagStringNeu`, Umschalter `DiagNeuAktiv`, Testroutine `DiagTest "155"`.
- `Importiere.bas` (Diff: 17 neue, 2 geänderte Zeilen): (a) `DiagString` ruft für Patienten mit `MOStatus > 0` zuerst `DiagStringNeu`; bei Fehler (`ok = False`) läuft die alte Logik weiter; (b) neue Hilfsroutine `DiagNrSetzen` (weil `DiagNr` und `gk` privat sind); (c) in `MachDiagnosen` der Fehler `DSic = "g"` → `"G"` (Zeilen 6829/6838; wirkt auch für die alte Logik).
- `Diagnose_Paarregeln.csv` / `.sql`: Tabelle `quelle.diagpaare` (212 Regeln aus Ihrer Liste; angelegt).
- Datenbank `quelle`: neue Funktion `diag_jahr`, `diag_info` liefert ein neuntes Feld (Vergleichstext). Skripte in `Referenz_Python/diag_funktionen.sql`.

**Test (Ihre IDE): Projekt neu von der Platte laden, kompilieren, dann im Direktfenster `DiagTest "155"`.** Ausgabe: erst der Diagnosestring der alten Logik, dann der neuen. Erwartet für Patient 155 sind 22 Zeilen, siehe `Referenz_Python/Pat155_erwartet.txt` (aus der Python-Referenz). Bei Abweichungen: `DiagNeuAbschalten = True` im Direktfenster schaltet auf die alte Logik um. Rückweg: `git checkout Importiere.bas` und die Modulzeile in der `.vbp` löschen (Backup der Vorversionen liegt nur im Session-Scratchpad, die Originale sind in Git).

**Was passiert wo**
- Patienten mit `MOStatus > 0` (neu aus `medoff` übertragen): neue Logik. Alle anderen (auch `MOStatus 0/NULL`): alte Logik unverändert (Fallback, bis alle neu übertragen sind).
- Pflegetabellen in `quelle` (Änderungen wirken nach Programmneustart oder nach `DiagRegelnNeuLaden` im Direktfenster):
  - `diagentscheid` – abgeschlossene Diagnosen ohne aktives Gegenstück und Z.n. ohne gesicherten Partner (n / j / je / jz, je ICD-Präfix, Textmuster, Altersgrenze).
  - `diagpaare` – aktive Z.n. gegen aktive gesichert je ICD-3: Code (g z j b a Zahl) und Bedingungen (GLEICH, GJUENGER, SEITE, SEITE_OHNE_AUSSENINNEN, GTEXT:…, ZTEXT:…), Auswahl: längster Präfix, dann erste passende Regel nach `Reihenfolge`. Kommentar-Kopf im SQL-Skript. Ein Paar gilt als „gleiche Diagnose“ (Bedingung GLEICH) nach den Abgleichregeln.
  - `diagnormal` – Abkürzungen/Synonyme für den Abgleich (nur Vergleich, nie Anzeige).
- Mehrere Paare in einer ICD-3-Stelle: eine Zeile entfällt, sobald irgendein Paar sie ausblendet; b fasst die nur mit b verbundenen Z.n. bzw. gesicherten Zeilen zur jüngsten zusammen, a und „Zahl“ (Abstand erreicht) nicht.

**Bekannte Unterschiede/Grenzen**
- Die Python-Referenz stellt Datums- und Erläuterungswörter als Menge dar, das VB-Modul vergleicht sie als Zeichenkette in der Reihenfolge des Textes (nur beim Zusammenfassen identischer Kopien relevant).
- `AnzeigeText` in `DiagAnzeige.bas` ist eine Kopie der Umformungen in `MachDiagnosen`; Änderungen dort bitte auch hier nachziehen.
- Laufzeit: eine zusätzliche kleine Abfrage je Aufruf plus etwa 1,3 ms je Diagnosezeile (Funktionen in der Datenbank).
- Aktive Ausschluss-Diagnosen erscheinen weiter als „Ausschluss …“; abgeschlossene V.a. erscheinen je nach Tabelle als „V.a.“ (j) bzw. „Z.n. V.a.“ (je/jz).

## 13. Änderungen nach den ersten Tests (21.9.) und Vorbereitung der Neuübertragung

**Test-Rückmeldungen und Regeländerungen** (Referenz und `DiagAnzeige.bas` angepasst; VB nicht kompiliert)
- Beispielpatient: E11 nur einmal ist gewollt (bisherige Zusammenfassung nach ICD-3 und Textanfang), R68 fehlt, weil `R68.8` auf der Ausschlussliste für Briefe steht (`obBrief`). Kein Fehler.
- Beispielpatient („Z.n. Cholezystolithiasis“ neben „… mit GB-Operation“): neue Abgleichstufe **Wortteilmenge** bei gleicher ICD (alle Wörter des weniger spezifischen Textes kommen im spezifischeren vor); außerdem verdrängt unter den **aktiven Z.n.** die spezifischere Zeile die weniger spezifische (mit Seiten-, Datums- und Erläuterungsregeln wie bisher).
- Beispielpatient (diabetisches Fußsyndrom): Die **Paarregeln (`diagpaare`) gelten jetzt für alle angezeigten Z.n.- und gesichert-Zeilen**, auch wenn sie aus abgeschlossenen Einträgen stammen. Ziffer nach dem Punkt = Wagner + 1: Regel `L89.1` (Wagner 0) → g ergänzt (Tabelle, CSV und SQL); `L89.0` → g und übrige L89 → a bleiben. Erwartung dort: drei L89-Zeilen (Wagner „ “ bds. [L89.08], Wagner 0 [L89.12], Z.n. Wagner 1 bds. [L89.28Z]).
- „seit“ bleibt für alle E1x-Codes (auch E15). Idee für später: „seit/ED …“ auch bei anderen Diagnosen anzeigen (Regeltabelle nötig).
- Neue Erwartungswerte ohne „seit“: `Referenz_Python/Erwartet_ohne_seit/<Pat_id>.txt` (155: 22, 45: 23, 196: 17, 193: 23, 282: 57, 21618: 28, 76: 19, 850: 63, 228: 43, 53142: 24 Zeilen).
- Die IDE hatte `DiagAnzeige.bas` mit vereinheitlichter Groß-/Kleinschreibung gespeichert (inhaltlich nur `Call Lese.ProgStart` in `DiagTest` neu, die ist übernommen). Die Datei auf der Platte wurde danach neu geschrieben: **Projekt neu laden**, sonst überschreibt die IDE die neue Fassung wieder.

**Neuübertragung aller Patienten** (`turichtdiag`, vonMO.bas): vorbereitet, nicht gestartet
- Umfang: 19.166 Patienten in `medoff` (bisher Grenze 20.000, jetzt 100.000). `quelle.diagnosen`: 18.923 Patienten, 262.701 Zeilen; bereits neu übertragen (Tests): 35 Patienten.
- Sicherung: `quelle.diagnosen_bak_20260921` (262.701 Zeilen, Stand vor dem Lauf). Rückweg bei Bedarf: Zeilen der betroffenen Patienten aus der Sicherung zurückkopieren.
- `turichtdiag` hat jetzt optionale Argumente: `turichtdiag , 200` (Testlauf mit den 200 höchsten Patientennummern), `turichtdiag 41234` (Wiederaufnahme mit Pat.-Nr. ≤ 41234), ohne Argumente alle. Start im Direktfenster: `Call turichtdiag`.
- Ablauf je Patient: `DELETE` der Patientenzeilen, `MODiagnosen` (MO-Abfrage ca. 15 ms), Einfügen, Fall-ID nachtragen. Gesamtdauer unbekannt (erst Testlauf mit 200 Patienten hochrechnen).
- Kontrolle danach: `mariadb quelle < Neuuebertragung_Kontrollen.sql` (Zeilen nach MOStatus, Patienten ohne MOStatus, Widerspruchsfreiheit, größte Änderungen gegenüber der Sicherung). Vor dem Lauf zeigt es 35 Patienten mit `MOStatus`.
- Patienten, die nur in `quelle` stehen (nicht in `medoff`), behalten ihre alten Zeilen ohne `MOStatus` und damit die alte Anzeige.

**Testlauf der Neuübertragung (22.9.):** 644 Patienten haben jetzt `MOStatus` (davon 35 aus den ersten Tests). Kontrollskript: alle Widerspruchsprüfungen 0, kein Patient ohne Diagnosen. Nachgerechnet gegen `medoff` (neue Regeln: ohne Status 3, ohne abgeschlossene Ausschluss-Diagnosen): bei allen 644 Patienten stimmt die Zeilenzahl in `quelle` **exakt** mit dem Soll überein. Abweichungen gegenüber der Sicherung (41 Patienten) erklären sich durch die schon vor der Sicherung neu übertragenen Testpatienten und durch Neueinträge in `medoff` seit der letzten Übertragung. Referenz über die 644 Patienten: 6.624 Anzeigezeilen (5.251 gesichert, 1.027 Z.n., 272 V.a., 21 Z.n. V.a., 53 Ausschluss); auffällig nur 21 Fälle mit gleichem Text mehrfach in einer ICD-3-Stelle und 32 Fälle mit weniger spezifischer Zeile neben spezifischerer (meist L89, G63, I83, E11, F17 mit unterschiedlicher ICD).

**Nachbesserungen nach dem Testlauf (22.9.):**
- L89.0x zeigt jetzt "Stadium Wagner 0" (wie L89.1x), in DiagAnzeige.bas und in der alten Logik (Importiere.bas), vorher stand nur "Wagner " ohne Ziffer. Erwartungsdateien 76 und 282 angepasst.
- Abgleich abgeschlossen gegen aktiv (DiagAnzeige.bas `Passt`, Referenz `match`): (a) die Erlaeuterung der aktiven Zeile zaehlt zu deren Woertern (Rest/Wortteilmenge), (b) eine Datumsangabe im Text der abgeschlossenen Zeile gilt als gedeckt, wenn alle ihre Zahlengruppen in der Erlaeuterung der aktiven Zeile stehen (typisch: aktiv "Pneumonie" + Erlaeuterung "8/23" gegen abgeschlossen "Z.n. Pneumonie (8/23)"), (c) ein abschliessender Strich in der ICD (J45.9-, K59.0-, I80.2-) zaehlt beim Abgleich nicht (`IcdK`; angezeigt wird die ICD unveraendert).
- `turichtdiag` schreibt je fertigem Patienten eine Zeile ins Direktfenster (`n/gesamt  Pat <nr> fertig  ...`); Wiederaufnahme mit `turichtdiag <nr>`.
- `Muster_Testlauf_644.csv`: Paare auffaelliger Zeilen (gleiche ICD-3 und Label) der 644 Testpatienten zur Beurteilung (ohne Namen).
- Entscheidung 22.9.: "spezifischere verdraengt weniger spezifische" gilt auch fuer aktive gesicherte Zeilen (bisher nur aktive Z.n.). Gleiche ICD-3, Woerter (inkl. Erlaeuterung), Datum und Seite der weniger spezifischen sind Teilmenge der spezifischeren; verschiedene ICD derselben 3-Stelle nur, wenn die Anzeige den Text nicht aus der ICD erzeugt (L89, E11, E66 ausgenommen, dort nur gleiche ICD). Umsetzung: DiagAnzeige.bas (Passt mit `locker`, Abschnitt 1c) und Referenz `visible`. Erwartungsdateien Erwartet_ohne_seit/ neu erzeugt (196, 228, 282, 850 geaendert).

**Vollauf der Neuübertragung (`Call turichtdiag`, abgeschlossen 22./23.9.):** Kontrollskript danach: MOStatus 1: 2.562 Zeilen/1.449 Pat.; 2: 500/335; 4: 112.423/10.515; 5: 147.351/16.758; NULL: 2/1 Pat. Alle Widerspruchsprüfungen 0. Abgeschlossene V.a.: 7.473 Zeilen/2.840 Patienten.
- **285 Patienten ohne Diagnosen nach dem Lauf** (vorher welche): gegen medoff geprüft — alle 308 `behgrund`-Zeilen dieser Patienten sind abgeschlossene Ausschluss-Diagnosen, 0 davon übertragbar. Erwartungsgemäß leer, kein Fehler.
- **1 Patient (60214) ohne MOStatus>0-Zeile:** hat keine einzige `behgrund`-Zeile mehr in medoff (nicht mehr dort vorhanden); `turichtdiag` waehlt seine Patientenliste aus `behgrund` und laesst ihn deshalb unangetastet. Seine 2 alten Zeilen (MOStatus NULL, Diagdatum 2014/2015) sind Restbestand von vor der Umstellung. Sonderfall, kein Fehler.
- **Grösste Zeilenzahl-Aenderungen gegenüber der Sicherung, nachgerechnet gegen medoff (Soll = FStatus<>3 und nicht abgeschlossene Ausschluss-Diagnose):** 63368 (36→24, Soll 24), 54127 (65→54, Soll 54), 51817 (61→50, Soll 50) — alle drei exakt. Die Rückgänge sind vollständig historische Diagnosen (FStatus=3): 13, 16 bzw. 11 Zeilen je Patient.

**Neue Regel L89 Wagner 0 (22./23.9., Ihr Vorschlag nach Test von 282):** Liegt eine aktive gesicherte Zeile "Diabetisches Fußsyndrom Stadium Wagner 0" (ICD-Ziffer nach dem Punkt 0 oder 1) vor, entfallen abgeschlossene/Z.n.-Wagner-0-Zeilen, deren Seite(n) bereits durch die Seite(n) der aktiven gesicherten Zeile(n) abgedeckt sind (keine Seitenangabe zaehlt als abgedeckt). Wagner 1-3 sind nicht betroffen. Umsetzung in DiagAnzeige.bas (Ende von Auswahl) und in der Referenz (visible()). Wirkung ueber die 644 Testpatienten: 2 betroffen (282: 5→3 Zeilen wie von Ihnen erwartet; 65582: 1 Zeile weniger). Erwartungsdatei 282 aktualisiert.
