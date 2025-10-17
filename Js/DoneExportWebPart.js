/*
  SPFx Web Part (React) – Done→CSV Exporter
  Zweck: Direkt in SharePoint modern pages nutzbar. Per Klick exportiert das Webpart
  alle erledigten Karten einer Liste (älter als X Tage) als CSV in eine Zielbibliothek.

  Hinweise & Grenzen
  - Läuft CLIENTSEITIG: kein echter Zeitplan ohne Benutzerinteraktion. Für Auto-Schedule → Power Automate.
  - Minimaler SPFx-Setup (SPFx 1.18+, Node 18). PnPjs für REST-Aufrufe.

  Installation (Kurz)
  1) yo @microsoft/sharepoint  → React Web Part
  2) npm i @pnp/sp @pnp/graph @pnp/logging @pnp/common
  3) Ersetze die WebPart-Komponente mit diesem Code (z.B. DoneExportWebPart.tsx)
  4) gulp serve / bundle / package-solution → App-Katalog → Webpart auf Seite hinzufügen
*/

import * as React from 'react';
import { useCallback, useMemo, useState } from 'react';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';

// ===== Helper: CSV =====
function toCsv(rows: Record<string, any>[], delimiter = ','): string {
  if (!rows || rows.length === 0) return '';
  const headers = Object.keys(rows[0]);
  const esc = (v: any) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    if (s.includes('"') || s.includes(delimiter) || /[\r\n]/.test(s)) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const lines = [headers.join(delimiter)];
  for (const r of rows) lines.push(headers.map(h => esc(r[h])).join(delimiter));
  return lines.join('\n');
}

// ===== Props aus dem Property Pane (in der echten SPFx-Klasse mappen) =====
interface IProps {
  spfxContext: any;
  listTitle: string;             // z.B. "UserStories"
  statusField: string;           // z.B. "Status"
  doneValue: string;             // z.B. "Erledigt"
  doneDateField: string;         // z.B. "DoneDate"
  selectFields: string;          // z.B. "ID,Title,Status,DoneDate,Assignee,StoryPoints,Sprint,Priority,Labels,Created,Modified"
  olderThanDays: number;         // z.B. 30
  targetLibraryServerRelUrl: string; // z.B. "/sites/Agile/Shared Documents/Exports"
  filePrefix?: string;           // z.B. "UserStories_Archive"
}

const DoneExportWebPart: React.FC<IProps> = (props) => {
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);
  const append = (m: string) => setLog(l => [...l, m]);

  const sp = useMemo(() => spfi().using(spSPFx(props.spfxContext)), [props.spfxContext]);

  const runExport = useCallback(async () => {
    try {
      setBusy(true); setLog([]);
      const {
        listTitle, statusField, doneValue, doneDateField,
        selectFields, olderThanDays, targetLibraryServerRelUrl, filePrefix
      } = props;

      append(`Lese aus Liste: ${listTitle} …`);
      const threshold = new Date();
      threshold.setUTCDate(threshold.getUTCDate() - (olderThanDays || 30));
      const thresholdIso = threshold.toISOString();

      // 1) Items holen – CAML Query via renderListDataAsStream (performanter & Filter serverseitig)
      // Alternativ: items.select(...).filter(...) – jedoch OData-Filter bei Datum teils knifflig.
      const viewXml = `
        <View>
          <Query>
            <Where>
              <And>
                <Eq>
                  <FieldRef Name='${statusField}' />
                  <Value Type='Text'>${doneValue}</Value>
                </Eq>
                <Leq>
                  <FieldRef Name='${doneDateField}' />
                  <Value IncludeTimeValue='TRUE' Type='DateTime'>${thresholdIso}</Value>
                </Leq>
              </And>
            </Where>
            <OrderBy><FieldRef Name='${doneDateField}' Ascending='TRUE' /></OrderBy>
          </Query>
        </View>`;

      const list = sp.web.lists.getByTitle(listTitle);
      // Wir nutzen die View-API, um passgenau die Felder zu laden
      const fields = selectFields.split(',').map(s => s.trim()).filter(Boolean);
      const r = await list.renderListDataAsStream({
        ViewXml: viewXml,
        ViewFields: fields,
        RenderOptions: 2 // ListData + weitere Metadaten
      });

      const rows = (r?.Row || []).map((row: any) => {
        const obj: any = {};
        for (const f of fields) {
          let v = row[f];
          // Person/Lookup-Felder kommen oft als "DisplayName" bereits flach an – ggf. Mapping anpassen
          if (v && typeof v === 'string' && /Z$/.test(v) && !isNaN(Date.parse(v))) {
            v = new Date(v).toISOString();
          }
          obj[f] = v ?? '';
        }
        obj['ExportedOnUtc'] = new Date().toISOString();
        return obj;
      });

      append(`Gefundene Elemente: ${rows.length}`);
      if (rows.length === 0) { append('Nichts zu exportieren.'); setBusy(false); return; }

      // 2) CSV bauen
      const csv = toCsv(rows);
      const ts = new Date().toISOString().slice(0,10);
      const fileName = `${filePrefix || 'Kanban_Archive'}_${ts}.csv`;

      // 3) Datei in Bibliothek speichern
      append(`Schreibe Datei in: ${targetLibraryServerRelUrl}/${fileName}`);
      const folder = sp.web.getFolderByServerRelativePath(targetLibraryServerRelUrl);
      await folder.files.addUsingPath(fileName, new Blob([csv], { type: 'text/csv;charset=utf-8' }), { Overwrite: true });

      append('Export abgeschlossen ✅');
    } catch (e: any) {
      console.error(e);
      append(`Fehler: ${e.message || e.toString()}`);
    } finally {
      setBusy(false);
    }
  }, [props, sp]);

  return (
    <div style={{fontFamily:'Segoe UI', padding:12}}>
      <h2 style={{margin:'0 0 8px'}}>Done → CSV Exporter</h2>
      <p style={{marginTop:0}}>Exportiert erledigte Items (älter als {props.olderThanDays || 30} Tage) aus <strong>{props.listTitle}</strong> in <code>{props.targetLibraryServerRelUrl}</code>.</p>
      <button disabled={busy} onClick={runExport} style={{
        padding:'8px 12px', borderRadius:8, border:'1px solid #ddd', cursor: busy? 'not-allowed':'pointer'
      }}>{busy ? 'Läuft…' : 'Jetzt exportieren'}</button>
      <div style={{marginTop:12, background:'#fafafa', border:'1px solid #eee', padding:8, borderRadius:8, maxHeight:200, overflow:'auto'}}>
        {log.map((l,i)=>(<div key={i} style={{fontSize:12}}>{l}</div>))}
      </div>
      <details style={{marginTop:12}}>
        <summary>Parameter (Property Pane)</summary>
        <ul>
          <li><strong>listTitle</strong>, <strong>statusField</strong>, <strong>doneValue</strong>, <strong>doneDateField</strong></li>
          <li><strong>selectFields</strong> (CSV-Spaltenauswahl)</li>
          <li><strong>olderThanDays</strong> (Standard 30)</li>
          <li><strong>targetLibraryServerRelUrl</strong> (z.B. "/sites/Agile/Shared Documents/Exports")</li>
          <li><strong>filePrefix</strong> (Dateinamenspräfix)</li>
        </ul>
      </details>
    </div>
  );
};

export default DoneExportWebPart;

/*
SPFx Verdrahtung (Kurz, pseudocode):

// DoneExportWebPartWebPart.ts
export interface IDoneExportWebPartProps {
  listTitle: string; statusField: string; doneValue: string; doneDateField: string;
  selectFields: string; olderThanDays: number; targetLibraryServerRelUrl: string; filePrefix?: string;
}

public render(): void {
  const element = React.createElement(DoneExportWebPart, {
    spfxContext: this.context,
    listTitle: this.properties.listTitle || 'UserStories',
    statusField: this.properties.statusField || 'Status',
    doneValue: this.properties.doneValue || 'Erledigt',
    doneDateField: this.properties.doneDateField || 'DoneDate',
    selectFields: this.properties.selectFields || 'ID,Title,Status,DoneDate,Assignee,StoryPoints,Sprint,Priority,Labels,Created,Modified',
    olderThanDays: this.properties.olderThanDays || 30,
    targetLibraryServerRelUrl: this.properties.targetLibraryServerRelUrl || '/Shared Documents/Exports',
    filePrefix: this.properties.filePrefix || 'UserStories_Archive'
  } as any);
  ReactDom.render(element, this.domElement);
}

// PropertyPane: Textfelder/Slider für die o.g. Eigenschaften erstellen

*/
