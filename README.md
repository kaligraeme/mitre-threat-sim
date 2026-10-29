# ThreatSim · MITRE ATT&CK data instance

A browser-based cyber threat simulator. It keeps a local copy of the MITRE ATT&CK **Enterprise** and
**ICS** datasets (Mobile is available but off by default) and maps Groups, Software and Campaigns to the
techniques they use, then to the relevant mitigations and detections.

## Pages

- **Data** (`index.html`): pull MITRE releases into the local database.
- **Simulator** (`simulator.html`): build a threat profile and map it to mitigations and detections.

## Simulator

1. **Profile:** pick Groups, Software and/or Campaigns, by search or with filters (description keywords,
   software type and platform, campaign activity window). **Select all shown** adds everything the
   current filter matches.
2. **Scope:** optionally pull in techniques from software that selected groups and campaigns use, and from
   campaigns attributed to selected groups. You can also limit results to chosen platforms.
3. **Results**, per domain (Enterprise / ICS):
   - **Matrix:** ATT&CK matrix of the profile's techniques. *Usage* shades each technique by how many
     selected adversaries use it; *Coverage* shows simulated posture.
   - **Techniques:** sortable table: who uses each technique, its mitigation and detection counts, and coverage.
   - **Mitigations / Detections:** ranked by priority. Tick *In place* / *Deployed*, or press
     **Mark top 5**, to see how much exposure is left.
   - **Telemetry:** the data components and log sources the detections rely on.
   - **ICS assets:** assets targeted by the profile's ICS techniques.
   - Click any row or matrix cell for details: procedure examples, mitigation guidance, and analytics
     with their log sources.
4. **Export:** an Excel workpaper (below), a CSV of the technique mapping, or ATT&CK Navigator layers
   (one JSON per domain).

Priority uses a weighted greedy method: each step picks the control that covers the most uncovered
techniques, with a technique counting once per selected adversary that uses it. Controls with equal
scores share a priority number and are ordered by reach, documented procedure examples, then ATT&CK ID,
so the result doesn't depend on the order you picked adversaries in. Your selections and
posture are saved in the browser's localStorage.

## Excel workpaper

**Export workpaper** on the Simulator tab writes the current profile and posture to an `.xlsx` workbook:

| Sheet | Contents |
|---|---|
| Executive Summary | Overall readiness with the **readiness bar**, readiness per domain, controls in place, the top 10 recommended mitigations and detection strategies per domain, coverage by tactic, the profile at a glance, the threat profile, and scope and method notes |
| Matrix (Enterprise) / Matrix (ICS) | ATT&CK matrix of the profile, shaded by how many selected adversaries use each technique, with a live coverage column (✓ mitigated + detected, M mitigated only, D detected only, ✕ no coverage) |
| Techniques | Every profile technique: tactics, platforms, who uses it, its mitigations and detection strategies, live coverage, procedure examples |
| Mitigations | **Implementation status** drop-down (*In Place* / *Not In Place*), owner, target date and notes; status at export and a live *Changed* flag; priority, adds and cumulative coverage at export; live count of unmitigated techniques it addresses |
| Detections | **Deployment status** drop-down (*Deployed* / *Not Deployed*), the same tracking and priority columns, analytics, data components and log sources |
| ICS Assets | Assets targeted by the profile's ICS techniques, with live counts of exposed and fully covered techniques |
| Telemetry | Data components and log sources the detection strategies rely on, with a live count of deployed strategies using each |
| Analytics | One row per analytic log source: what to collect and what to look for, with the strategy's live deployment status |
| Control Map | Each technique-to-control link, with the control's live status and how the mitigation applies |

The drop-downs start at the posture you set in ThreatSim. Everything that depends on them is a formula:
technique coverage, the readiness figures and bars, the matrix status column and the columns marked
*(live)*. Changing a status recalculates the workbook; no macros are involved. Columns marked
*(at export)*, including priorities, keep their values from export time; to re-rank, set the same
statuses in ThreatSim and export again. Controls with equal priority scores share a priority number and
are listed by reach, documented procedure examples, then ATT&CK ID.

The bars are conditional formatting on cells, so they work in Excel, LibreOffice Calc and Google Sheets.
Every formula also stores its value at export, so file previews that don't recalculate still show the
right figures. Sheets print landscape with page footers; the matrices print on A3.

The workbook is built in the browser with [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0 (MIT,
`js/vendor/exceljs.min.js`, licence in `js/vendor/exceljs.LICENSE.txt`). It loads only when you first
export.

## Run

```bash
cd mitre-threat-sim
python3 -m http.server 8000
```

Open http://localhost:8000 and press **Pull**. Opening `index.html` directly from disk also works,
but serving it keeps the database on one origin for every page you add later.

## How it works

- **Checking for updates:** reads MITRE's release index
  (`https://raw.githubusercontent.com/mitre-attack/attack-stix-data/master/index.json`) every time
  the page loads, and again when you press **Check for updates**. It compares the latest version of
  each domain with what's stored.
- **Pull:** downloads the STIX 2.1 bundle for each selected domain that's out of date. It checks that
  the bundle's version is the one requested, normalizes it, then replaces that domain's data in a
  single IndexedDB transaction. If a pull fails, the data already stored is left as it was.
- The **Pull** dropdown on each card can pin an older release. The **Re-pull** checkbox downloads
  again even when the stored data is current.

## Database (`mitre-attack-db`, IndexedDB)

| Store | Key | Contents |
|---|---|---|
| `meta` | domain | stored version, release date, source URL, pulled-at, counts |
| `objects` | `<domain>\|<stix id>` | slim normalized records: `kind`, `attackId`, `name`, `description`, `platforms`, `tactics`, `parent` (sub-techniques), `aliases`, `analyticRefs`, `logSources`, `active` … |
| `relationships` | `<domain>\|<stix id>` | `relType` (uses, mitigates, detects, attributed-to, subtechnique-of, targets, revoked-by), `source`, `target`, endpoint kinds, `active` |
| `raw` | `<domain>\|<stix id>` | original STIX object (citations, external references) |
| `history` | auto | log of pulls |

`active` is false for revoked or deprecated objects, and for relationships that touch one.

Mapping paths the simulator will use:

```
Group / Software / Campaign --uses--> Technique
Campaign --attributed-to--> Group,   Group --uses--> Software
Mitigation --mitigates--> Technique
Detection Strategy --detects--> Technique
Detection Strategy .analyticRefs -> Analytic .logSources[].dataComponent -> Data Component
Technique (ICS) --targets--> Asset
```

## API for other pages (`js/attack-db.js`, global `AttackDB`)

`getObjects(domain, kind)`, `getObject(domain, id)`, `getRaw(domain, id)`, `findByAttackId(id, domain?)`,
`relationshipsFrom(domain, id, relType?)`, `relationshipsTo(domain, id, relType?)`,
`relationshipsByType(domain, relType)`, `search(text, {domain, kinds, limit})`, `getAllMeta()`.
