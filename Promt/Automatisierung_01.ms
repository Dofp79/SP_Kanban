Fungiere als erfahrener Experte für die Automatisierung von Datenübertragungen in Microsoft SharePoint. In SharePoint existiert ein Kanban-Board, das mit HTML, CSS und JavaScript entwickelt wurde. Ziel ist es, alle abgeschlossenen User Stories aus der Spalte „Erledigt“ nach einem festgelegten Zeitraum (standardmäßig alle 30 Tage) automatisch auszulesen und an einem definierten Speicherort zu archivieren.

Erstelle einen vollständigen technischen Automatisierungsplan mit folgenden Aspekten:
1. Beschreibung der Datenquelle (HTML/DOM-Struktur des Boards)
2. Identifikation der Spalte „Erledigt“ und ihrer Inhalte
3. Datenextraktion (z. B. mit JavaScript oder REST-API)
4. Auswahl und Vergleich möglicher Automatisierungsansätze (z. B. Power Automate, Azure Function, Timer-Job)
5. Speicherformat (z. B. CSV, Excel, JSON oder SharePoint-Liste)
6. Zeitplanung und Automatisierungsintervall
7. Authentifizierung und Berechtigungskonzept
8. Fehlerbehandlung, Protokollierung, Wiederholungslogik
9. Benachrichtigungen bei Erfolg oder Fehler
10. Qualitätssicherung und Teststrategie
11. Wartbarkeit und Dokumentation der Lösung

Bitte liefere das Ergebnis als strukturiertes Konzeptdokument in Abschnitten und mit Best Practices aus der SharePoint-Automatisierung.


Zielbild (kurz)

Quelle: Kanban-Board (SharePoint-Liste mit Board-Ansicht oder Microsoft Planner / Azure DevOps Boards).

Regel: Alle User Stories im Status Erledigt werden nach 30 Tagen (oder konfigurierbar) in ein Archiv verschoben/kopiert.

Ablage: SharePoint-Archivliste (empfohlen), optional zusätzlich CSV/Excel in einer Doku-Bibliothek.

Reporting: Power BI (Direct/Import via SharePoint-Connector); fertige Measures für Durchsatz, Lead-/Cycle-Time, Velocity.

Power Automate ist die passende Automatisierungs-Engine; SharePoint-Listen eignen sich hervorragend für strukturierte Workflows & Reporting; Power BI bindet SharePoint-Listen nativ an.

Variante A — SharePoint-Liste mit Board-Ansicht (empfohlen)
1) Datenmodell (Quellliste)

Liste UserStories (Board-Ansicht)

Title (Einzelne Textzeile)

Status (Auswahl: Neu | In Arbeit | Blockiert | Erledigt)

DoneDate (Datum, wird beim Wechsel auf Erledigt gesetzt)

Assignee (Person)

StoryPoints (Zahl)

Priority, Labels/Tags, Sprint (Auswahl/Mehrfachauswahl)

Optional: CreatedDate, StartDate, Area/Team

Listen/Spalten, Inhalts­typen und JSON-Formatierung für saubere UX werden in den bereitgestellten Unterlagen systematisch behandelt.

2) Archivliste

Kopie der Spaltenstruktur als Liste UserStories_Archive (zusätzlich: ArchivedOn (Datum/Uhrzeit), SourceItemId (Zahl), SourceSiteUrl (Hyperlink)).

3) Power Automate — geplanter Flow (täglich)

Trigger: „Geplante Cloud-Durchführung“ (z. B. 02:00 Uhr, Europe/Berlin).
Schritte (Kernlogik):

„Get items“ der Liste UserStories mit OData-Filter:

Status eq 'Erledigt' and DoneDate le @{formatDateTime(addDays(utcNow(), -30), 'yyyy-MM-dd')}


„Apply to each“: Für jedes Element

„Create item“ in UserStories_Archive (Felder 1:1 mappen; ArchivedOn = utcNow(), SourceItemId = ID).

„Delete item“ aus UserStories oder (wenn du Wiederherstellbarkeit möchtest) „Update item“ mit Status = „Archiviert“ und optionalem Kennzeichen (IsArchived = true) statt Löschen.

Fehler-/Duplikat-Schutz (optional, aber empfohlen)

„Get items“ in Archiv mit Filter SourceItemId eq @{items('…')?['ID']}; nur kopieren, wenn leer.

Logging/Notify: Tägliche Zusammenfassung an Owner (Anzahl archiviert).

Archiving via Flow & Best Practices (Archivlisten, Wiederherstellbarkeit, Governance) sind im Material ausführlich beschrieben.

Hinweise

Wenn „DoneDate“ noch nicht gesetzt wird: Füge einen zweiten (ereignis­basierten) Flow hinzu („Wenn ein Element erstellt oder geändert wird“), der bei Status=Erledigt einmalig DoneDate = utcNow() schreibt (Guard: nur wenn leer).

Rechte: Archivliste nur für Admins/Report-Rolle änderbar; Leser für alle.

Wiederherstellung: Bei „Delete“ kannst du 93-Tage-Papierkorb nutzen; alternativ „Soft-Archive“ (kein Delete).

Variante B — Planner (oder Azure DevOps Boards)
Planner

Herausforderung: Planner hat limitierte Filter in Standard-Connectoren. Zwei robuste Wege:

Scheduled Flow + „List tasks“ je Board/Bucket, dann manuell filtern auf percentComplete = 100 und completedDateTime <= addDays(utcNow(),-30).

HTTP mit Graph (Application/Delegated) → GET /planner/plans/{id}/tasks inkl. completedDateTime; danach identische Archiv-Schritte wie oben (Ziel: SharePoint-Archivliste).

Datenanreicherung: Mappe Planner-Felder (title, bucket, assignedTo, labels, checklist count, completedDateTime) → Archivliste.

Azure DevOps Boards

Connector: „Azure DevOps – Abfragearbeitsitems“ mit Wiql (State = Done, ClosedDate <= Today-30).

Archivziel: ebenfalls SharePoint-Archivliste; ADO-IDs als SourceItemId.

Power Automate deckt beide Welten ab; für komplexe Filter sind HTTP/Graph bzw. Wiql üblich. Die Reporting-Anbindung an SharePoint/Power BI bleibt gleich.

Power BI — Reporting-Blueprint
1) Datenquelle

SharePoint-Listen-Connector (beide: UserStories & UserStories_Archive).

In Power Query:

Vereinheitliche Spalten, füge IsArchived (true/false) hinzu, Append zu Tabelle FactsUserStories.

Erzeuge DimDate, DimAssignee, DimSprint (aus eindeutigen Werten) für Sternschema.

2) Measures (DAX – Beispiele)
Throughput (30d) =
CALCULATE(
    COUNTROWS(FactsUserStories),
    FILTER(FactsUserStories, FactsUserStories[Status] = "Erledigt"),
    DATESINPERIOD(DimDate[Date], MAX(DimDate[Date]), -30, DAY)
)

Cycle Time (Ø Tage) =
AVERAGEX(
    FILTER(FactsUserStories, NOT ISBLANK([DoneDate]) && NOT ISBLANK([StartDate])),
    DATEDIFF(FactsUserStories[StartDate], FactsUserStories[DoneDate], DAY)
)

Velocity (Story Points / Sprint) =
SUMX(VALUES(DimSprint[Sprint]), CALCULATE(SUM(FactsUserStories[StoryPoints]), FactsUserStories[Status] = "Erledigt"))

3) Aktualisierung

Import mit geplanter Aktualisierung (8× täglich, je nach Lizenz) oder DirectQuery via Dataverse/SQL, falls später migriert.

Für große Datenmengen: Incremental Refresh anhand DoneDate/CreatedDate (Parameter RangeStart/RangeEnd).

4) Veröffentlichung & Einbettung

In Power BI Service veröffentlichen; Power BI Webpart in SharePoint einbinden, Berechtigungen via M365-Gruppen steuern.

Aufbau, Veröffentlichung und Einbettung von Power BI, inkl. Datenquellen-Übersicht, sind im bereitgestellten Material zusammengefasst.

Governance, Sicherheit & Betrieb

Retention/Compliance: Falls rechtliche Aufbewahrung nötig, nutze Labels/Retention Policies statt harter Löschung; Archivliste kann Record-fähig sein.

Versionierung: In Quell- und Archivlisten aktivieren; Audit-Spalten (Erstellt/Geändert von).

Fehlerhandling im Flow: „Konfiguration – Ausführen nach Fehlversuch“, Scope-Blöcke mit „Try/Catch“, E-Mail/Teams-Alert bei Fehlern.

Skalierung: Bei >5k Items OData-Filter + Pagination in „Get items“; ggf. monatliche Partitionslisten (Archiv_2025_10).

Umsetzungs-Checkliste (kurz & knackig)

Quellliste (oder Planner/ADO) prüfen: Status-Werte, DoneDate sicherstellen.

Archivliste anlegen (Spiegel der Felder + ArchivedOn, SourceItemId).

Flows:

A) On-Change: setzt DoneDate beim Wechsel auf Erledigt (nur einmal).

B) Scheduled: filtert Erledigt & DoneDate ≤ Heute-30, kopiert → Archiv, löscht/markiert Quelle.

Power BI: SharePoint-Connector, Append, Sternschema, Measures, Publish & Webpart.

Governance: Rechte, Monitoring, Backup/Recovery, Dokumentation.

Optional: Schnellstart-Snippets

Flow-Filter („Get items“)

Status eq 'Erledigt' and DoneDate le '@{formatDateTime(addDays(utcNow(), -30), 'yyyy-MM-dd')}'


Bedin­gung in Flow (Expression)

@and(
  equals(items('Apply_to_each')?['Status'], 'Erledigt'),
  lessOrEquals(items('Apply_to_each')?['DoneDate'], addDays(utcNow(), -30))
)
