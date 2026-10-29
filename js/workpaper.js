/*
 * workpaper.js — exports the simulator dashboard as an Excel workpaper.
 *
 * Global `Workpaper`. Builds an .xlsx with ExcelJS (vendored in js/vendor and
 * loaded on first use) with these sheets: Executive Summary, one Matrix sheet
 * per domain, Techniques, Mitigations, Detections, ICS Assets, Telemetry and
 * Control Map.
 *
 * The status cells on Mitigations and Detections are drop-down inputs. Every
 * figure that depends on them — technique coverage, the readiness bars, matrix
 * status, the "live" columns — is a formula, so the workbook recalculates when a
 * status changes. Each formula also carries its value for the posture at export,
 * so previews that never recalculate still show the right numbers.
 */
(function (global) {
  'use strict';

  const EXCELJS_SRC = 'js/vendor/exceljs.min.js';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const MIT = { on: 'In Place', off: 'Not In Place' };
  const DET = { on: 'Deployed', off: 'Not Deployed' };
  const COVERAGE = { covered: 'Mitigated + detected', mitigated: 'Mitigated only', detected: 'Detected only', gap: 'No coverage' };
  const GLYPH = { covered: '✓', mitigated: 'M', detected: 'D', gap: '✕' };
  const SHEET = {
    summary: 'Executive Summary', techniques: 'Techniques', mitigations: 'Mitigations', detections: 'Detections',
    assets: 'ICS Assets', telemetry: 'Telemetry', map: 'Control Map',
  };
  const matrixSheetName = label => 'Matrix (' + label + ')';

  // Readiness bars on the summary: BAR.cells narrow columns from column BAR.first, each worth 1/BAR.cells.
  const BAR = { first: 10, cells: 40 };
  const TOP_ACTIONS = 10;

  // Light palette only: spreadsheet cells have no dark mode.
  const C = {
    ink: 'FF111827', muted: 'FF6B7280', white: 'FFFFFFFF',
    header: 'FF1F2937', band: 'FF334155', subband: 'FFE2E8F0', border: 'FFD1D5DB',
    input: 'FFFFF4CC', inputBorder: 'FFD4A600', inputHeader: 'FF8A6D00', inputTab: 'FFFAB219',
    live: 'FFF1F5F9', link: 'FF1D4ED8', track: 'FFE5E7EB', idle: 'FFF3F4F6',
    good: 'FF0CA30C', warning: 'FFFAB219', critical: 'FFD03B3B',
    goodTint: 'FFD9F2D9', goodInk: 'FF0B5D0B', warnTint: 'FFFFF0C7', warnInk: 'FF7A5200', critTint: 'FFF9D6D6', critInk: 'FF8B1D1D',
  };
  // Sequential blue (usage), light steps 1-6 with the ink that stays legible on each.
  const SEQ = [
    ['FFCDE2FB', 'FF0B0B0B'], ['FF9EC5F4', 'FF0B0B0B'], ['FF6DA7EC', 'FF0B0B0B'],
    ['FF2A78D6', 'FFFFFFFF'], ['FF1C5CAB', 'FFFFFFFF'], ['FF104281', 'FFFFFFFF'],
  ];

  // ---------- column layouts (fixed, so formulas can address columns by key) ----------

  const TECH_COLS = [
    { key: 'domain', header: 'Domain', width: 11 },
    { key: 'id', header: 'Technique ID', width: 12, link: true },
    { key: 'name', header: 'Technique', width: 34 },
    { key: 'parent', header: 'Sub-technique of', width: 28 },
    { key: 'tactics', header: 'Tactics', width: 26 },
    { key: 'platforms', header: 'Platforms', width: 22 },
    { key: 'usedN', header: 'Used by (#)', width: 9, fmt: '0' },
    { key: 'usedBy', header: 'Used by', width: 32 },
    { key: 'mitN', header: 'Mitigations (#)', width: 11, fmt: '0' },
    { key: 'mits', header: 'Mitigations', width: 42 },
    { key: 'detN', header: 'Detection strategies (#)', width: 11, fmt: '0' },
    { key: 'dets', header: 'Detection strategies', width: 42 },
    { key: 'assets', header: 'ICS assets targeted', width: 30 },
    { key: 'mitigated', header: 'Mitigation in place? (live)', width: 13, live: true, center: true },
    { key: 'detected', header: 'Detection deployed? (live)', width: 13, live: true, center: true },
    { key: 'coverage', header: 'Coverage (live)', width: 21, live: true },
    { key: 'procedures', header: 'How the profile uses it (procedure examples)', width: 70 },
    { key: 'description', header: 'Description', width: 70 },
    { key: 'key', header: 'Key', helper: true },
    { key: 'tacticKeys', header: 'Tactic keys', helper: true },
    { key: 'mitKeys', header: 'Mitigation keys', helper: true },
    { key: 'detKeys', header: 'Detection keys', helper: true },
    { key: 'assetKeys', header: 'Asset keys', helper: true },
    { key: 'code', header: 'Coverage code (live)', helper: true },
  ];

  const MIT_COLS = [
    { key: 'priority', header: 'Priority (at export)', width: 10, fmt: '0', center: true },
    { key: 'domain', header: 'Domain', width: 11 },
    { key: 'id', header: 'Mitigation ID', width: 12, link: true },
    { key: 'name', header: 'Mitigation', width: 36 },
    { key: 'status', header: 'Implementation status ▼', width: 18, input: true },
    { key: 'techN', header: 'Profile techniques addressed (#)', width: 13, fmt: '0' },
    { key: 'share', header: 'Share of profile techniques', width: 12, fmt: '0%' },
    { key: 'open', header: 'Unmitigated techniques it addresses (live)', width: 15, fmt: '0', live: true },
    { key: 'techs', header: 'Profile techniques addressed', width: 50 },
    { key: 'description', header: 'Description', width: 70 },
    { key: 'owner', header: 'Owner (optional)', width: 18 },
    { key: 'notes', header: 'Notes / evidence (optional)', width: 36 },
    { key: 'key', header: 'Key', helper: true },
  ];

  const DET_COLS = [
    { key: 'priority', header: 'Priority (at export)', width: 10, fmt: '0', center: true },
    { key: 'domain', header: 'Domain', width: 11 },
    { key: 'id', header: 'Detection ID', width: 12, link: true },
    { key: 'name', header: 'Detection strategy', width: 44 },
    { key: 'status', header: 'Deployment status ▼', width: 18, input: true },
    { key: 'techN', header: 'Profile techniques covered (#)', width: 13, fmt: '0' },
    { key: 'share', header: 'Share of profile techniques', width: 12, fmt: '0%' },
    { key: 'open', header: 'Undetected techniques it covers (live)', width: 15, fmt: '0', live: true },
    { key: 'anN', header: 'Analytics (#)', width: 10, fmt: '0' },
    { key: 'platforms', header: 'Analytic platforms', width: 24 },
    { key: 'dcs', header: 'Data components', width: 36 },
    { key: 'logSources', header: 'Log sources', width: 36 },
    { key: 'techs', header: 'Profile techniques covered', width: 40 },
    { key: 'analytics', header: 'Analytics', width: 40 },
    { key: 'owner', header: 'Owner (optional)', width: 18 },
    { key: 'notes', header: 'Notes / evidence (optional)', width: 36 },
    { key: 'key', header: 'Key', helper: true },
    { key: 'dcKeys', header: 'Data component keys', helper: true },
  ];

  const ASSET_COLS = [
    { key: 'id', header: 'Asset ID', width: 10, link: true },
    { key: 'name', header: 'ICS asset', width: 36 },
    { key: 'techN', header: 'Profile techniques targeting it (#)', width: 14, fmt: '0' },
    { key: 'share', header: 'Share of ICS profile techniques', width: 13, fmt: '0%' },
    { key: 'exposed', header: 'Targeting techniques with no coverage (live)', width: 16, fmt: '0', live: true },
    { key: 'covered', header: 'Targeting techniques mitigated + detected (live)', width: 16, fmt: '0', live: true },
    { key: 'platforms', header: 'Platforms', width: 24 },
    { key: 'sectors', header: 'Sectors', width: 18 },
    { key: 'techs', header: 'Profile techniques targeting it', width: 50 },
    { key: 'description', header: 'Description', width: 70 },
    { key: 'key', header: 'Key', helper: true },
  ];

  const TEL_COLS = [
    { key: 'domain', header: 'Domain', width: 11 },
    { key: 'id', header: 'Data component ID', width: 13, link: true },
    { key: 'name', header: 'Data component', width: 32 },
    { key: 'techN', header: 'Profile techniques it helps detect (#)', width: 14, fmt: '0' },
    { key: 'share', header: 'Share of profile techniques', width: 12, fmt: '0%' },
    { key: 'dsN', header: 'Detection strategies relying on it (#)', width: 14, fmt: '0' },
    { key: 'deployed', header: 'Deployed strategies using it (live)', width: 14, fmt: '0', live: true },
    { key: 'logSources', header: 'Log sources', width: 44 },
    { key: 'details', header: 'What to collect (log source: what to look for)', width: 70 },
    { key: 'description', header: 'Description', width: 60 },
    { key: 'key', header: 'Key', helper: true },
  ];

  const MAP_COLS = [
    { key: 'domain', header: 'Domain', width: 11 },
    { key: 'techId', header: 'Technique ID', width: 12 },
    { key: 'tech', header: 'Technique', width: 34 },
    { key: 'type', header: 'Control type', width: 17 },
    { key: 'ctlId', header: 'Control ID', width: 11, link: true },
    { key: 'ctl', header: 'Control', width: 40 },
    { key: 'status', header: 'Control status (live)', width: 15, live: true },
    { key: 'guidance', header: 'How it applies to this technique', width: 80 },
    { key: 'techKey', header: 'Technique key', helper: true },
    { key: 'ctlKey', header: 'Control key', helper: true },
  ];
  const MAP_TYPE = { mit: 'Mitigation', det: 'Detection strategy' };

  // ---------- small helpers ----------

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  const quoteSheet = name => "'" + name.replace(/'/g, "''") + "'";
  const lit = s => '"' + String(s).replace(/"/g, '""') + '"';       // string literal inside a formula
  const F = (formula, result) => ({ formula, result });
  const uniq = a => [...new Set(a)];
  const byAttackId = (a, b) => String(a.attackId).localeCompare(String(b.attackId), undefined, { numeric: true });
  const byName = (a, b) => String(a.name || '').localeCompare(String(b.name || ''));

  const MAX_CELL = 32000;   // Excel's limit is 32,767 characters per cell

  /** Text that is valid XML 1.0 and fits in a cell. */
  function clean(s) {
    if (s == null) return '';
    let t = String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
    t = typeof t.toWellFormed === 'function'
      ? t.toWellFormed()
      : t.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDFFF]/g, m => (m.length === 2 ? m : '�'));
    if (t.length > MAX_CELL) {
      let cut = MAX_CELL - 1;
      if (/[\uD800-\uDBFF]/.test(t[cut - 1])) cut--;
      t = t.slice(0, cut) + '…';
    }
    return t;
  }

  /** ATT&CK description markdown flattened to plain text. */
  function plain(md) {
    if (!md) return '';
    return clean(String(md)
      .replace(/\s*\(Citation:[^)]*\)/g, '')
      .replace(/\[([^\]]+)\]\((?:https?:\/\/[^)\s]+)\)/g, '$1')
      .replace(/<\/?code>/g, '')
      .replace(/`([^`]+)`/g, '$1')
      .replace(/\*\*([^*]+)\*\*/g, '$1')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim());
  }

  const link = (text, url) => (url ? { text: clean(text), hyperlink: url } : clean(text));
  const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb }, bgColor: { argb } });
  const thin = argb => ({ style: 'thin', color: { argb } });
  const box = argb => ({ top: thin(argb), left: thin(argb), bottom: thin(argb), right: thin(argb) });
  const font = (o = {}) => Object.assign({ name: 'Calibri', size: 10, color: { argb: C.ink } }, o);

  function put(ws, row, col, value, style) {
    const cell = ws.getCell(row, col);
    cell.value = value;
    if (style) Object.assign(cell, style);
    return cell;
  }

  /** Column addressing for one table sheet (header in row 1, data from row 2). */
  function layout(sheet, cols, count) {
    const idx = new Map(cols.map((c, i) => [c.key, i + 1]));
    const last = Math.max(2, count + 1);   // never an inverted range, even for an empty table
    const letter = key => {
      if (!idx.has(key)) throw new Error(`Unknown column ${key} on ${sheet}`);
      return colLetter(idx.get(key));
    };
    return {
      sheet, cols, count, last, letter,
      range: key => `${quoteSheet(sheet)}!$${letter(key)}$2:$${letter(key)}$${last}`,
    };
  }

  /** Per-sheet conditional-format priority counter. */
  function priorities() { let p = 0; return () => ++p; }

  // ---------- data preparation (no ExcelJS) ----------

  function prepare(ctx) {
    const P = { domains: [], techniques: [], mitigations: [], detections: [], assets: [], telemetry: [], map: [], profile: [] };
    const generated = ctx.generatedAt || new Date();

    for (const r of ctx.results) {
      const m = r.model;
      const label = ctx.domainLabel(r.domain);
      const keyOf = o => label + ':' + (o.attackId || o.id);
      const tacticName = new Map(m.tactics.map(t => [t.shortname, t.name]));
      const mitOn = id => ctx.inPlace.mitigations.has(r.domain + '|' + id);
      const detOn = id => ctx.inPlace.detections.has(r.domain + '|' + id);
      const name = id => { const o = m.byId.get(id); return o ? o.name : id; };
      const idName = id => { const o = m.byId.get(id); return o ? (o.attackId ? o.attackId + ' ' : '') + o.name : id; };
      const total = r.techniques.size;
      const techList = [...r.techniques.values()];

      const D = {
        domain: r.domain, label, result: r, total,
        version: m.meta.version, released: m.meta.modified,
        counts: { covered: 0, mitigated: 0, detected: 0, gap: 0 },
        subs: 0, tacticsHit: 0, tacticsAll: m.tactics.length,
        tactics: [], mitTotal: r.mitigations.size, detTotal: r.detections.size, mitOn: 0, detOn: 0,
        dcTotal: r.dataComponents.size, assetTotal: r.assets.size, topMitigations: [], topDetections: [],
      };
      for (const t of techList) {
        D.counts[t.status]++;
        if (t.tech.isSubtechnique) D.subs++;
      }
      for (const tactic of m.tactics) {
        const ts = techList.filter(t => (t.tech.tactics || []).includes(tactic.shortname));
        const c = { covered: 0, mitigated: 0, detected: 0, gap: 0 };
        for (const t of ts) c[t.status]++;
        D.tactics.push({ tactic, total: ts.length, counts: c });
        if (ts.length) D.tacticsHit++;
      }
      for (const id of r.mitigations.keys()) if (mitOn(id)) D.mitOn++;
      for (const id of r.detections.keys()) if (detOn(id)) D.detOn++;
      P.domains.push(D);

      // Techniques: most-used first, then by ID.
      techList.sort((a, b) => b.roots.size - a.roots.size || byAttackId(a.tech, b.tech));
      for (const t of techList) {
        const mitIds = uniq(t.mitigations), detIds = uniq(t.detections), assetIds = uniq(t.assets);
        const parent = t.tech.parent ? m.byId.get(t.tech.parent) : null;
        const roots = [...t.roots].map(id => m.byId.get(id)).filter(Boolean).sort(byName);
        const seen = new Set();
        const procedures = [];
        for (const ev of t.evidence) {
          const k = ev.root + '>' + ev.via.join('>') + '|' + ev.description;
          if (seen.has(k)) continue;
          seen.add(k);
          const via = ev.via.length ? ' (via ' + ev.via.map(id => { const o = m.byId.get(id); return (o && o.kind === 'campaign' ? 'campaign ' : '') + name(id); }).join(' → ') + ')' : '';
          procedures.push(name(ev.root) + via + ': ' + plain(ev.description));
        }
        P.techniques.push({
          D, t,
          domain: label,
          id: link(t.tech.attackId, t.tech.url),
          name: clean(t.tech.name),
          parent: parent ? clean(parent.attackId + ' ' + parent.name) : '',
          tactics: clean((t.tech.tactics || []).map(s => tacticName.get(s) || s).join('; ')),
          platforms: clean((t.tech.platforms || []).join(', ')),
          usedN: t.roots.size,
          usedBy: clean(roots.map(o => o.name).join('; ')),
          mitN: mitIds.length,
          mits: clean(mitIds.map(idName).join('; ')),
          detN: detIds.length,
          dets: clean(detIds.map(idName).join('; ')),
          assets: clean(assetIds.map(name).sort().join('; ')),
          procedures: clean(procedures.join('\n')),
          description: plain(t.tech.description),
          key: keyOf(t.tech),
          tacticKeys: ';' + (t.tech.tactics || []).join(';') + ';',
          mitKeys: clean(';' + mitIds.map(id => keyOf(m.byId.get(id))).join(';') + ';'),
          detKeys: clean(';' + detIds.map(id => keyOf(m.byId.get(id))).join(';') + ';'),
          assetKeys: clean(';' + assetIds.map(id => keyOf(m.byId.get(id))).join(';') + ';'),
          isMitigated: t.isMitigated, isDetected: t.isDetected, status: t.status,
        });

        for (const id of mitIds) {
          const ctl = m.byId.get(id);
          P.map.push({
            domain: label, techId: t.tech.attackId, tech: clean(t.tech.name), type: MAP_TYPE.mit,
            ctlId: link(ctl.attackId, ctl.url), ctl: clean(ctl.name),
            guidance: plain(r.mitigations.get(id).guidance.get(t.tech.id)),
            techKey: keyOf(t.tech), ctlKey: keyOf(ctl), statusNow: mitOn(id) ? MIT.on : MIT.off, kind: 'mit',
            sort: [label, t.tech, 0, ctl],
          });
        }
        for (const id of detIds) {
          const ctl = m.byId.get(id);
          const e = r.detections.get(id);
          P.map.push({
            domain: label, techId: t.tech.attackId, tech: clean(t.tech.name), type: MAP_TYPE.det,
            ctlId: link(ctl.attackId, ctl.url), ctl: clean(ctl.name),
            guidance: clean('Analytics: ' + (e.analytics.map(a => a.attackId + ' (' + (a.platforms || []).join(', ') + ')').join('; ') || 'none for the selected platforms')),
            techKey: keyOf(t.tech), ctlKey: keyOf(ctl), statusNow: detOn(id) ? DET.on : DET.off, kind: 'det',
            sort: [label, t.tech, 1, ctl],
          });
        }
      }

      // Mitigations: in place at export first, then by priority, then by reach.
      const order = (pri, on) => (a, b) => {
        const pa = pri.get(a[0]), pb = pri.get(b[0]);
        const ka = on(a[0]) ? -1 : pa.rank == null ? 1e6 : pa.rank;
        const kb = on(b[0]) ? -1 : pb.rank == null ? 1e6 : pb.rank;
        return ka - kb || b[1].techniques.size - a[1].techniques.size || byAttackId(a[1].obj, b[1].obj);
      };
      for (const [id, e] of [...r.mitigations].sort(order(r.mitPriority, mitOn))) {
        const p = r.mitPriority.get(id);
        const techs = [...e.techniques].map(tid => m.byId.get(tid)).sort(byAttackId);
        const row = {
          D, e,
          priority: mitOn(id) || p.rank == null ? null : p.rank,
          domain: label,
          id: link(e.obj.attackId, e.obj.url),
          name: clean(e.obj.name),
          status: mitOn(id) ? MIT.on : MIT.off,
          techN: e.techniques.size,
          share: total ? e.techniques.size / total : 0,
          openNow: techs.filter(tech => !r.techniques.get(tech.id).isMitigated).length,
          techs: clean(techs.map(tech => tech.attackId).join(', ')),
          description: plain(e.obj.description),
          owner: '', notes: '',
          key: keyOf(e.obj),
        };
        P.mitigations.push(row);
        if (row.priority != null && row.priority <= TOP_ACTIONS) D.topMitigations.push(row);
      }

      for (const [id, e] of [...r.detections].sort(order(r.detPriority, detOn))) {
        const p = r.detPriority.get(id);
        const techs = [...e.techniques].map(tid => m.byId.get(tid)).sort(byAttackId);
        const dcs = [...e.dataComponents].map(dcId => m.byId.get(dcId)).filter(Boolean).sort(byName);
        const logSources = uniq(e.analytics.flatMap(a => (a.logSources || []).map(l => l.name).filter(Boolean))).sort();
        const row = {
          D, e,
          priority: detOn(id) || p.rank == null ? null : p.rank,
          domain: label,
          id: link(e.obj.attackId, e.obj.url),
          name: clean(e.obj.name),
          status: detOn(id) ? DET.on : DET.off,
          techN: e.techniques.size,
          share: total ? e.techniques.size / total : 0,
          openNow: techs.filter(tech => !r.techniques.get(tech.id).isDetected).length,
          anN: e.analytics.length,
          platforms: clean(uniq(e.analytics.flatMap(a => a.platforms || [])).sort().join(', ')),
          dcs: clean(dcs.map(dc => dc.name).join('; ')),
          logSources: clean(logSources.join('; ')),
          techs: clean(techs.map(tech => tech.attackId).join(', ')),
          analytics: clean(e.analytics.map(a => a.attackId + ' (' + (a.platforms || []).join(', ') + ')').join('; ')),
          owner: '', notes: '',
          key: keyOf(e.obj),
          dcKeys: clean(';' + [...e.dataComponents].map(dcId => m.byId.get(dcId)).filter(Boolean).map(keyOf).join(';') + ';'),
        };
        P.detections.push(row);
        if (row.priority != null && row.priority <= TOP_ACTIONS) D.topDetections.push(row);
      }

      for (const [, e] of [...r.assets].sort((a, b) => b[1].techniques.size - a[1].techniques.size || byAttackId(a[1].obj, b[1].obj))) {
        const techs = [...e.techniques].map(tid => r.techniques.get(tid));
        P.assets.push({
          D, e,
          id: link(e.obj.attackId, e.obj.url),
          name: clean(e.obj.name),
          techN: e.techniques.size,
          share: total ? e.techniques.size / total : 0,
          exposedNow: techs.filter(t => t.status === 'gap').length,
          coveredNow: techs.filter(t => t.status === 'covered').length,
          platforms: clean((e.obj.platforms || []).join(', ')),
          sectors: clean((e.obj.sectors || []).join(', ')),
          techs: clean(techs.map(t => t.tech).sort(byAttackId).map(tech => tech.attackId).join(', ')),
          description: plain(e.obj.description),
          key: keyOf(e.obj),
        });
      }

      for (const [, e] of [...r.dataComponents].sort((a, b) => b[1].techniques.size - a[1].techniques.size || byAttackId(a[1].obj, b[1].obj))) {
        const names = [...e.logSources.keys()].sort();
        const details = names.map(n => {
          const ch = [...e.logSources.get(n)];
          return n + (ch.length ? ': ' + ch.slice(0, 8).join(' | ') + (ch.length > 8 ? ' | +' + (ch.length - 8) + ' more' : '') : '');
        });
        P.telemetry.push({
          D, e,
          domain: label,
          id: link(e.obj.attackId, e.obj.url),
          name: clean(e.obj.name),
          techN: e.techniques.size,
          share: total ? e.techniques.size / total : 0,
          dsN: e.strategies.size,
          deployedNow: [...e.strategies].filter(detOn).length,
          logSources: clean(names.join('; ')),
          details: clean(details.join('\n')),
          description: plain(e.obj.description),
          key: keyOf(e.obj),
        });
      }
    }

    P.map.sort((a, b) => {
      const [da, ta, ka, ca] = a.sort, [db, tb, kb, cb] = b.sort;
      return P.domains.findIndex(d => d.label === da) - P.domains.findIndex(d => d.label === db)
        || byAttackId(ta, tb) || ka - kb || byAttackId(ca, cb);
    });

    // Threat profile: every selected adversary, with the techniques it contributes.
    for (const kind of ctx.kinds) {
      for (const id of ctx.selection[kind.key]) {
        const o = ctx.entity(id);
        if (!o) continue;
        const inDomains = P.domains.filter(D => D.result.roots.some(x => x.id === id));
        let n = 0;
        for (const D of inDomains) for (const t of D.result.techniques.values()) if (t.roots.has(id)) n++;
        P.profile.push({
          obj: o, kind: kind.one + (o.softwareType ? ' (' + o.softwareType + ')' : ''),
          domains: inDomains.map(D => D.label).join(', ') || 'Not in the selected domains',
          techniques: n,
          aliases: (o.aliases || []).filter(a => a !== o.name).join(', '),
        });
      }
    }

    P.generated = generated;
    P.profileNames = P.profile.map(p => p.obj.name);
    return P;
  }

  // ---------- generic table writer ----------

  /**
   * Header row + data rows. Row values may be functions of the row number
   * (used for formulas that reference their own row).
   */
  function writeTable(ws, L, rows, { freezeCols = 0, emptyNote } = {}) {
    L.cols.forEach((c, i) => {
      const col = ws.getColumn(i + 1);
      col.width = c.helper ? 14 : (c.width || 12);
      if (c.helper) col.hidden = true;
      put(ws, 1, i + 1, c.header, {
        font: font({ bold: true, color: { argb: C.white } }),
        fill: solid(c.input ? C.inputHeader : C.header),
        alignment: { vertical: 'middle', wrapText: true },
        border: box(C.border),
      });
    });
    ws.getRow(1).height = 42;

    rows.forEach((row, ri) => {
      const rn = ri + 2;
      L.cols.forEach((c, ci) => {
        let v = row[c.key];
        if (typeof v === 'function') v = v(rn);
        const cell = ws.getCell(rn, ci + 1);
        cell.value = v === undefined || v === '' ? null : v;
        const isLink = v && typeof v === 'object' && v.hyperlink;
        cell.font = font(isLink ? { color: { argb: C.link }, underline: true } : {});
        if (c.fmt) cell.numFmt = c.fmt;
        if (c.live) cell.fill = solid(C.live);
        if (c.input) {
          cell.fill = solid(C.input);
          cell.border = box(C.inputBorder);
          cell.font = font({ bold: true });
        }
        cell.alignment = { vertical: 'top', horizontal: c.center ? 'center' : undefined };
      });
    });

    if (!rows.length && emptyNote) {
      put(ws, 2, 1, emptyNote, { font: font({ italic: true, color: { argb: C.muted } }) });
    }
    // The filter range includes the hidden helper columns so sorting keeps each row's keys with it.
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: L.cols.length } };
    ws.views = [{ state: 'frozen', xSplit: freezeCols, ySplit: 1, activeCell: 'A2' }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    ws.pageSetup.printTitlesRow = '1:1';
  }

  function coverageRules(ws, ref, prio) {
    const rule = (text, fill, ink) => ({
      type: 'cellIs', operator: 'equal', formulae: [lit(text)], priority: prio(),
      style: { fill: solid(fill), font: { color: { argb: ink }, bold: true } },
    });
    ws.addConditionalFormatting({
      ref,
      rules: [
        rule(COVERAGE.covered, C.goodTint, C.goodInk),
        rule(COVERAGE.mitigated, C.warnTint, C.warnInk),
        rule(COVERAGE.detected, C.warnTint, C.warnInk),
        rule(COVERAGE.gap, C.critTint, C.critInk),
      ],
    });
  }

  function statusInputRules(ws, ref, onValue, prio) {
    ws.addConditionalFormatting({
      ref,
      rules: [{
        type: 'cellIs', operator: 'equal', formulae: [lit(onValue)], priority: prio(),
        style: { fill: solid(C.goodTint), font: { color: { argb: C.goodInk }, bold: true } },
      }],
    });
  }

  function dropdown(ws, L, onValue, offValue, what) {
    if (!L.count) return;
    const col = L.letter('status');
    ws.dataValidations.add(`${col}2:${col}${L.count + 1}`, {
      type: 'list',
      allowBlank: true,
      formulae: [lit(onValue + ',' + offValue)],
      showInputMessage: true,
      promptTitle: what,
      prompt: `Choose ${onValue} or ${offValue}. Readiness on the Executive Summary updates automatically.`,
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Choose from the list',
      error: `Use ${onValue} or ${offValue}.`,
    });
  }

  // ---------- sheets ----------

  function writeTechniques(ws, P, L, LM) {
    const n = rn => ({
      mitigated: `IF(COUNTIFS(${LM.range('techKey')},$${L.letter('key')}${rn},${LM.range('type')},${lit(MAP_TYPE.mit)},${LM.range('status')},${lit(MIT.on)})>0,"Yes","No")`,
      detected: `IF(COUNTIFS(${LM.range('techKey')},$${L.letter('key')}${rn},${LM.range('type')},${lit(MAP_TYPE.det)},${LM.range('status')},${lit(DET.on)})>0,"Yes","No")`,
    });
    const pick = (rn, out) => {
      const m = `$${L.letter('mitigated')}${rn}="Yes"`, d = `$${L.letter('detected')}${rn}="Yes"`;
      return `IF(AND(${m},${d}),${lit(out.covered)},IF(${m},${lit(out.mitigated)},IF(${d},${lit(out.detected)},${lit(out.gap)})))`;
    };
    const rows = P.techniques.map(row => Object.assign({}, row, {
      mitigated: rn => F(n(rn).mitigated, row.isMitigated ? 'Yes' : 'No'),
      detected: rn => F(n(rn).detected, row.isDetected ? 'Yes' : 'No'),
      coverage: rn => F(pick(rn, COVERAGE), COVERAGE[row.status]),
      code: rn => F(pick(rn, GLYPH), GLYPH[row.status]),
    }));
    writeTable(ws, L, rows, { freezeCols: 3, emptyNote: 'No techniques in this threat profile.' });
    const prio = priorities();
    const cov = L.letter('coverage');
    coverageRules(ws, `${cov}2:${cov}${L.last}`, prio);
    const yes = col => ws.addConditionalFormatting({
      ref: `${col}2:${col}${L.last}`,
      rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"Yes"'], priority: prio(),
        style: { font: { color: { argb: C.goodInk }, bold: true } } }],
    });
    yes(L.letter('mitigated'));
    yes(L.letter('detected'));
  }

  function writeControls(ws, rows, L, LT, kind) {
    const isMit = kind === 'mit';
    const V = isMit ? MIT : DET;
    const flag = isMit ? 'mitigated' : 'detected';
    const keysCol = isMit ? 'mitKeys' : 'detKeys';
    const out = rows.map(row => Object.assign({}, row, {
      open: rn => F(`SUMPRODUCT((${LT.range(flag)}="No")*ISNUMBER(SEARCH(";"&$${L.letter('key')}${rn}&";",${LT.range(keysCol)})))`, row.openNow),
    }));
    writeTable(ws, L, out, { freezeCols: 4, emptyNote: isMit ? 'No ATT&CK mitigations map to this threat profile.' : 'No detection strategies map to this threat profile.' });
    const prio = priorities();
    const st = L.letter('status');
    statusInputRules(ws, `${st}2:${st}${L.last}`, V.on, prio);
    dropdown(ws, L, V.on, V.off, isMit ? 'Implementation status' : 'Deployment status');
  }

  function writeAssets(ws, P, L, LT) {
    const rows = P.assets.map(row => Object.assign({}, row, {
      exposed: rn => F(`SUMPRODUCT((${LT.range('coverage')}=${lit(COVERAGE.gap)})*ISNUMBER(SEARCH(";"&$${L.letter('key')}${rn}&";",${LT.range('assetKeys')})))`, row.exposedNow),
      covered: rn => F(`SUMPRODUCT((${LT.range('coverage')}=${lit(COVERAGE.covered)})*ISNUMBER(SEARCH(";"&$${L.letter('key')}${rn}&";",${LT.range('assetKeys')})))`, row.coveredNow),
    }));
    writeTable(ws, L, rows, { freezeCols: 2, emptyNote: 'This threat profile has no ICS techniques, so no ICS assets are targeted. Add ICS groups, software or campaigns in ThreatSim to populate this sheet.' });
    const prio = priorities();
    const ex = L.letter('exposed');
    ws.addConditionalFormatting({
      ref: `${ex}2:${ex}${L.last}`,
      rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: prio(),
        style: { fill: solid(C.critTint), font: { color: { argb: C.critInk }, bold: true } } }],
    });
  }

  function writeTelemetry(ws, P, L, LD) {
    const rows = P.telemetry.map(row => Object.assign({}, row, {
      deployed: rn => F(`SUMPRODUCT((${LD.range('status')}=${lit(DET.on)})*ISNUMBER(SEARCH(";"&$${L.letter('key')}${rn}&";",${LD.range('dcKeys')})))`, row.deployedNow),
    }));
    writeTable(ws, L, rows, { freezeCols: 3, emptyNote: 'No detection strategies, and so no telemetry, map to this threat profile.' });
  }

  function writeMap(ws, P, L, LMit, LDet) {
    const rows = P.map.map(row => Object.assign({}, row, {
      status: rn => {
        const T = row.kind === 'mit' ? LMit : LDet;
        return F(`IFERROR(INDEX(${T.range('status')},MATCH($${L.letter('ctlKey')}${rn},${T.range('key')},0))&"","")`, row.statusNow);
      },
    }));
    writeTable(ws, L, rows, { freezeCols: 3, emptyNote: 'No mitigations or detection strategies map to this threat profile.' });
    const prio = priorities();
    const st = L.letter('status');
    ws.addConditionalFormatting({
      ref: `${st}2:${st}${L.last}`,
      rules: [MIT.on, DET.on].map(v => ({ type: 'cellIs', operator: 'equal', formulae: [lit(v)], priority: prio(),
        style: { fill: solid(C.goodTint), font: { color: { argb: C.goodInk }, bold: true } } })),
    });
  }

  // ---------- matrix ----------

  function usageStep(count, max) {
    if (max <= 1) return 4;
    return 1 + Math.round(((count - 1) / (max - 1)) * 5);
  }

  function writeMatrix(ws, D, ctx, LT) {
    const r = D.result;
    const cols = SimEngine.matrixColumns(r, { showAll: !!ctx.showAll });
    let max = 1;
    for (const t of r.techniques.values()) max = Math.max(max, t.roots.size);

    const TECH_W = 24, STATUS_W = 3.6;
    cols.forEach((c, i) => {
      ws.getColumn(2 * i + 1).width = STATUS_W;
      ws.getColumn(2 * i + 2).width = TECH_W;
    });

    // Title and legends
    put(ws, 1, 1, `${D.label} ATT&CK v${D.version} — threat profile matrix`, { font: font({ size: 14, bold: true }) });
    put(ws, 2, 1, `Profile: ${P_names(ctx)}. ${ctx.showAll ? 'All techniques are shown; grey ones are not used by the profile.' : 'Profile techniques only; grey parents give context for their sub-techniques.'}`,
      { font: font({ color: { argb: C.muted } }) });

    // Usage legend (technique columns of the first tactics)
    const groups = new Map();
    for (let c = 1; c <= max; c++) {
      const s = usageStep(c, max);
      if (!groups.has(s)) groups.set(s, []);
      groups.get(s).push(c);
    }
    put(ws, 3, 2, 'Shade = adversaries using it:', { font: font({ bold: true }) });
    let lc = 4;
    for (const [s, counts] of groups) {
      const lo = counts[0], hi = counts[counts.length - 1];
      put(ws, 3, lc, (lo === hi ? lo : lo + '–' + hi) + (hi === 1 ? ' adversary' : ' adversaries'), {
        font: font({ color: { argb: SEQ[s - 1][1] }, bold: true }), fill: solid(SEQ[s - 1][0]), alignment: { horizontal: 'center' },
        border: box(C.white),
      });
      lc += 2;
    }
    // Status legend
    const legend = [['covered', C.good, C.white], ['mitigated', C.warning, C.ink], ['detected', C.warning, C.ink], ['gap', C.critical, C.white]];
    legend.forEach(([k, fill, ink], i) => {
      put(ws, 4, 2 * i + 1, GLYPH[k], { font: font({ bold: true, color: { argb: ink } }), fill: solid(fill), alignment: { horizontal: 'center' } });
      put(ws, 4, 2 * i + 2, COVERAGE[k], { font: font() });
    });
    put(ws, 4, 2 * legend.length + 2, 'Status column is live: it follows the Mitigations and Detections sheets.', { font: font({ italic: true, color: { argb: C.muted } }) });

    const HEAD = 6, COUNTS = 7, FIRST = 8;
    const lists = cols.map(c => {
      const out = [];
      for (const row of c.rows) {
        out.push({ tech: row.tech, used: row.used, sub: false });
        for (const s of row.subs) out.push({ tech: s.tech, used: s.used, sub: true });
      }
      return out;
    });
    const height = Math.max(1, ...lists.map(l => l.length));
    const last = FIRST + height - 1;
    const lines = new Array(height).fill(1);

    cols.forEach((c, i) => {
      const sc = 2 * i + 1, tc = 2 * i + 2;
      const S = colLetter(sc);
      ws.mergeCells(HEAD, sc, HEAD, tc);
      put(ws, HEAD, sc, c.tactic.name, {
        font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.header),
        alignment: { vertical: 'middle', wrapText: true },
      });
      ws.mergeCells(COUNTS, sc, COUNTS, tc);
      const n = c.usedCount;
      const gaps = lists[i].filter(x => x.used && x.used.status === 'gap').length;
      const lead = n + ' technique' + (n === 1 ? '' : 's');
      put(ws, COUNTS, sc, n ? F(`${lit(lead + ' · ')}&COUNTIF(${S}${FIRST}:${S}${last},"✕")&" uncovered"`, `${lead} · ${gaps} uncovered`) : lead, {
        font: font({ color: { argb: C.muted } }), fill: solid(C.subband), alignment: { vertical: 'middle' },
      });

      lists[i].forEach((x, j) => {
        const rowN = FIRST + j;
        const text = `${x.tech.attackId} ${x.tech.name}`;
        const step = x.used ? usageStep(x.used.roots.size, max) : 0;
        const fill = x.used ? SEQ[step - 1][0] : C.idle;
        const ink = x.used ? SEQ[step - 1][1] : C.muted;
        put(ws, rowN, tc, clean(text), {
          font: font({ size: 9, color: { argb: ink }, italic: !x.used }),
          fill: solid(fill),
          alignment: x.sub ? { vertical: 'top', horizontal: 'left', wrapText: true, indent: 1 } : { vertical: 'top', wrapText: true },
          border: box(C.white),
        });
        if (x.used) {
          const key = `${D.label}:${x.tech.attackId}`;
          put(ws, rowN, sc, F(`IFERROR(INDEX(${LT.range('code')},MATCH(${lit(key)},${LT.range('key')},0)),"")`, GLYPH[x.used.status]), {
            font: font({ size: 9, bold: true }), alignment: { horizontal: 'center', vertical: 'top' }, border: box(C.white),
          });
        } else {
          put(ws, rowN, sc, null, { fill: solid(C.idle), border: box(C.white) });
        }
        const perLine = x.sub ? 27 : 30;   // ~characters per line at 9pt in a 24-wide column
        lines[j] = Math.max(lines[j], Math.ceil(text.length / perLine));
      });
    });

    // Live status colours (one rule set per status column)
    const prio = priorities();
    cols.forEach((c, i) => {
      const S = colLetter(2 * i + 1);
      const rule = (glyph, fill, ink) => ({ type: 'cellIs', operator: 'equal', formulae: [lit(glyph)], priority: prio(),
        style: { fill: solid(fill), font: { color: { argb: ink }, bold: true } } });
      ws.addConditionalFormatting({
        ref: `${S}${FIRST}:${S}${last}`,
        rules: [rule(GLYPH.covered, C.good, C.white), rule(GLYPH.mitigated, C.warning, C.ink),
          rule(GLYPH.detected, C.warning, C.ink), rule(GLYPH.gap, C.critical, C.white)],
      });
    });

    ws.getRow(HEAD).height = 30;
    ws.getRow(COUNTS).height = 16;
    lines.forEach((l, j) => { ws.getRow(FIRST + j).height = Math.max(15, 12 * l + 3); });
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: COUNTS, activeCell: 'A' + FIRST, showGridLines: false }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9, printTitlesRow: `${HEAD}:${COUNTS}` };
  }

  function P_names(ctx) {
    const names = [];
    for (const kind of ctx.kinds) for (const id of ctx.selection[kind.key]) { const o = ctx.entity(id); if (o) names.push(o.name); }
    return names.length > 8 ? names.slice(0, 8).join(', ') + ' +' + (names.length - 8) + ' more' : names.join(', ');
  }

  // ---------- executive summary ----------

  function writeSummary(ws, ctx, P, Ls) {
    const { T: LT, M: LM, D: LD } = Ls;
    const prio = priorities();
    const LAST_COL = BAR.first + BAR.cells - 1;
    const barA = colLetter(BAR.first), barZ = colLetter(LAST_COL);

    ws.getColumn(1).width = 34;
    for (let c = 2; c <= 8; c++) ws.getColumn(c).width = 13;
    ws.getColumn(9).width = 2;
    for (let c = BAR.first; c <= LAST_COL; c++) ws.getColumn(c).width = 1.4;

    let row = 1;
    const text = (r, c, v, f = {}) => put(ws, r, c, v, { font: font(f) });
    const band = (title) => {
      for (let c = 1; c <= LAST_COL; c++) put(ws, row, c, c === 1 ? title : null, { fill: solid(C.band), font: font({ bold: true, size: 12, color: { argb: C.white } }), alignment: { vertical: 'middle' } });
      ws.getRow(row).height = 22;
      row++;
    };
    const headerRow = (labels, barLabel) => {
      labels.forEach((l, i) => put(ws, row, i + 1, l, {
        font: font({ bold: true, color: { argb: C.ink } }), fill: solid(C.subband),
        alignment: { vertical: 'middle', wrapText: true, horizontal: i ? 'center' : undefined }, border: { bottom: thin(C.border) },
      }));
      if (barLabel) {
        ws.mergeCells(row, BAR.first, row, LAST_COL);
        put(ws, row, BAR.first, barLabel, { font: font({ bold: true }), fill: solid(C.subband), alignment: { vertical: 'middle' } });
      }
      ws.getRow(row).height = 30;
      row++;
    };
    const merged = (r, c1, c2, value, style, height) => {
      ws.mergeCells(r, c1, r, c2);
      put(ws, r, c1, value, style);
      if (height) ws.getRow(r).height = height;
    };
    const lines = (s, perLine) => Math.max(1, Math.ceil(String(s).length / perLine) + (String(s).match(/\n/g) || []).length);
    const barCells = r => {
      for (let k = 0; k < BAR.cells; k++) {
        put(ws, r, BAR.first + k, (k + 0.5) / BAR.cells, { numFmt: ';;;', fill: solid(C.track) });
      }
      ws.getRow(r).height = 18;
    };
    const stackedBarRules = (r1, r2) => {
      if (r2 < r1) return;
      const a = `${barA}${r1}`;
      ws.addConditionalFormatting({
        ref: `${barA}${r1}:${barZ}${r2}`,
        rules: [
          { type: 'expression', priority: prio(), formulae: [`AND($B${r1}>0,${a}<$C${r1}/$B${r1})`], style: { fill: solid(C.good) } },
          { type: 'expression', priority: prio(), formulae: [`AND($B${r1}>0,${a}<($C${r1}+$D${r1}+$E${r1})/$B${r1})`], style: { fill: solid(C.warning) } },
          { type: 'expression', priority: prio(), formulae: [`$B${r1}>0`], style: { fill: solid(C.critical) } },
        ],
      });
    };
    const progressBarRules = (r1, r2) => {
      if (r2 < r1) return;
      ws.addConditionalFormatting({
        ref: `${barA}${r1}:${barZ}${r2}`,
        rules: [{ type: 'expression', priority: prio(), formulae: [`AND($B${r1}>0,${barA}${r1}<$C${r1}/$B${r1})`], style: { fill: solid(C.good) } }],
      });
    };
    const pct = (num, den) => (den ? num / den : '');

    // Title block
    text(row, 1, 'ATT&CK threat readiness workpaper', { size: 18, bold: true });
    ws.getRow(row).height = 28;
    row++;
    text(row, 1, clean('Threat profile: ' + P_names(ctx)), { size: 12, bold: true });
    row++;
    const when = P.generated.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    text(row, 1, clean(`Generated ${when} with ThreatSim · ` + P.domains.map(D => `${D.label} ATT&CK v${D.version}`).join(' · ')), { color: { argb: C.muted } });
    row += 2;
    merged(row, 1, 8, 'How to use this workpaper: set each mitigation\'s Implementation status (In Place / Not In Place) on the Mitigations sheet, '
      + 'and each detection strategy\'s Deployment status (Deployed / Not Deployed) on the Detections sheet. They are the yellow drop-down cells. '
      + 'Every readiness figure and bar in this workbook recalculates from those choices.',
    { font: font({ bold: true, color: { argb: C.warnInk } }), fill: solid(C.input), alignment: { wrapText: true, vertical: 'middle' }, border: box(C.inputBorder) }, 48);
    row += 2;

    // Profile at a glance (static)
    band('Profile at a glance');
    const sum = k => P.domains.reduce((n, D) => n + D[k], 0);
    const totalTech = sum('total');
    const glance = [
      ['Adversaries in profile', P.profile.length, P_names(ctx)],
      ['Techniques in scope', totalTech, `${sum('subs')} sub-techniques · ${sum('tacticsHit')} of ${sum('tacticsAll')} tactics`],
      ['Mitigations mapped', sum('mitTotal'), 'ATT&CK mitigations that address at least one profile technique'],
      ['Detection strategies mapped', sum('detTotal'), 'ATT&CK detection strategies for the profile techniques'],
      ['Data components (telemetry)', sum('dcTotal'), 'Log data the detection strategies rely on'],
    ];
    if (P.domains.some(D => D.assetTotal)) glance.push(['ICS assets targeted', sum('assetTotal'), 'Assets targeted by the profile\'s ICS techniques']);
    for (const [label, value, note] of glance) {
      text(row, 1, label, { bold: true });
      put(ws, row, 2, value, { font: font({ size: 12, bold: true }), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      text(row, 3, clean(note), { color: { argb: C.muted } });
      row++;
    }
    row++;

    // Readiness (live)
    band('Readiness');
    const headline = row;
    const tableFirst = headline + 4;   // first data row of the readiness table (fixed below)
    put(ws, row, 1, 'Overall readiness', { font: font({ bold: true, size: 12 }), alignment: { vertical: 'middle' } });
    const all = { covered: 0, mitigated: 0, detected: 0, gap: 0 };
    for (const D of P.domains) for (const k of Object.keys(all)) all[k] += D.counts[k];
    const readyNow = pct(all.covered + all.mitigated + all.detected, totalTech);
    put(ws, row, 2, F(`$G$${tableFirst}`, readyNow), { font: font({ size: 22, bold: true }), numFmt: '0%', alignment: { horizontal: 'center', vertical: 'middle' } });
    merged(row, 3, 8, F(`$C$${tableFirst}&" of "&$B$${tableFirst}&" techniques are mitigated and detected, "&($D$${tableFirst}+$E$${tableFirst})&" are partly covered and "&$F$${tableFirst}&" have no coverage."`,
      `${all.covered} of ${totalTech} techniques are mitigated and detected, ${all.mitigated + all.detected} are partly covered and ${all.gap} have no coverage.`),
    { font: font({ size: 11 }), alignment: { wrapText: true, vertical: 'middle' } });
    ws.getRow(row).height = 36;
    row++;
    merged(row, 1, 8, 'Readiness = share of profile techniques with at least one mitigation In Place or detection strategy Deployed. Fully covered = both.',
      { font: font({ italic: true, color: { argb: C.muted } }) });
    row += 2;

    headerRow(['Scope', 'Techniques', COVERAGE.covered, COVERAGE.mitigated, COVERAGE.detected, COVERAGE.gap, 'Readiness', 'Fully covered'], 'Readiness bar');
    if (row !== tableFirst) throw new Error('Summary layout drifted');
    const T = key => LT.range(key);
    const scopeRows = [];
    if (P.domains.length > 1) scopeRows.push({ label: 'All domains', domain: null, counts: all, total: totalTech });
    for (const D of P.domains) scopeRows.push({ label: D.label + ' ATT&CK', domain: D.label, counts: D.counts, total: D.total });
    const firstScope = row;
    for (const s of scopeRows) {
      const crit = s.domain ? `${T('domain')},${lit(s.domain)},` : '';
      const count = status => s.domain
        ? `COUNTIFS(${crit}${T('coverage')},${lit(COVERAGE[status])})`
        : `COUNTIF(${T('coverage')},${lit(COVERAGE[status])})`;
      text(row, 1, s.label, { bold: true });
      put(ws, row, 2, F(s.domain ? `COUNTIF(${T('domain')},${lit(s.domain)})` : `COUNTA(${T('key')})`, s.total), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      ['covered', 'mitigated', 'detected', 'gap'].forEach((k, i) => {
        put(ws, row, 3 + i, F(count(k), s.counts[k]), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      });
      put(ws, row, 7, F(`IF($B${row}=0,"",($C${row}+$D${row}+$E${row})/$B${row})`, pct(s.counts.covered + s.counts.mitigated + s.counts.detected, s.total)),
        { font: font({ bold: true }), numFmt: '0%', alignment: { horizontal: 'center' } });
      put(ws, row, 8, F(`IF($B${row}=0,"",$C${row}/$B${row})`, pct(s.counts.covered, s.total)), { font: font(), numFmt: '0%', alignment: { horizontal: 'center' } });
      barCells(row);
      row++;
    }
    stackedBarRules(firstScope, row - 1);

    // Legend
    text(row, 1, 'Bar legend', { color: { argb: C.muted } });
    ws.mergeCells(row, 2, row, 3);
    put(ws, row, 2, COVERAGE.covered, { font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.good), alignment: { horizontal: 'center' } });
    put(ws, row, 4, 'Partial', { font: font({ bold: true, color: { argb: C.ink } }), fill: solid(C.warning), alignment: { horizontal: 'center' } });
    put(ws, row, 5, COVERAGE.gap, { font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.critical), alignment: { horizontal: 'center' } });
    row += 2;

    // Controls (live)
    headerRow(['Controls', 'Mapped', 'In place / deployed', 'Not yet', 'Share in place'], 'Progress');
    const ctlRows = [];
    const M = key => LM.range(key), Dt = key => LD.range(key);
    if (P.domains.length > 1) {
      ctlRows.push({ label: 'Mitigations in place', R: M, on: MIT.on, total: sum('mitTotal'), now: sum('mitOn') });
      for (const D of P.domains) ctlRows.push({ label: `  ${D.label} mitigations`, R: M, on: MIT.on, domain: D.label, total: D.mitTotal, now: D.mitOn });
      ctlRows.push({ label: 'Detection strategies deployed', R: Dt, on: DET.on, total: sum('detTotal'), now: sum('detOn') });
      for (const D of P.domains) ctlRows.push({ label: `  ${D.label} detection strategies`, R: Dt, on: DET.on, domain: D.label, total: D.detTotal, now: D.detOn });
    } else {
      ctlRows.push({ label: 'Mitigations in place', R: M, on: MIT.on, total: sum('mitTotal'), now: sum('mitOn') });
      ctlRows.push({ label: 'Detection strategies deployed', R: Dt, on: DET.on, total: sum('detTotal'), now: sum('detOn') });
    }
    const firstCtl = row;
    for (const c of ctlRows) {
      text(row, 1, c.label, { bold: !c.domain });
      put(ws, row, 2, F(c.domain ? `COUNTIF(${c.R('domain')},${lit(c.domain)})` : `COUNTA(${c.R('key')})`, c.total), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      put(ws, row, 3, F(c.domain ? `COUNTIFS(${c.R('domain')},${lit(c.domain)},${c.R('status')},${lit(c.on)})` : `COUNTIF(${c.R('status')},${lit(c.on)})`, c.now),
        { font: font({ bold: true }), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      put(ws, row, 4, F(`$B${row}-$C${row}`, c.total - c.now), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
      put(ws, row, 5, F(`IF($B${row}=0,"",$C${row}/$B${row})`, pct(c.now, c.total)), { font: font(), numFmt: '0%', alignment: { horizontal: 'center' } });
      barCells(row);
      row++;
    }
    progressBarRules(firstCtl, row - 1);
    row++;

    // Coverage by tactic (live)
    band('Coverage by tactic');
    headerRow(['Tactic', 'Techniques', COVERAGE.covered, COVERAGE.mitigated, COVERAGE.detected, COVERAGE.gap, 'Readiness', 'Fully covered'], 'Readiness bar');
    for (const D of P.domains) {
      if (P.domains.length > 1) {
        for (let c = 1; c <= 8; c++) put(ws, row, c, c === 1 ? `${D.label} ATT&CK` : null, { font: font({ bold: true }), fill: solid(C.idle) });
        row++;
      }
      const first = row;
      for (const tr of D.tactics) {
        const inTactic = `ISNUMBER(SEARCH(${lit(';' + tr.tactic.shortname + ';')},${T('tacticKeys')}))`;
        const inDomain = `(${T('domain')}=${lit(D.label)})`;
        const count = status => `SUMPRODUCT(${inDomain}*${inTactic}*(${T('coverage')}=${lit(COVERAGE[status])}))`;
        text(row, 1, clean(tr.tactic.name), { color: { argb: tr.total ? C.ink : C.muted } });
        put(ws, row, 2, F(`SUMPRODUCT(${inDomain}*${inTactic})`, tr.total), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
        ['covered', 'mitigated', 'detected', 'gap'].forEach((k, i) => {
          put(ws, row, 3 + i, F(count(k), tr.counts[k]), { font: font(), numFmt: '#,##0', alignment: { horizontal: 'center' } });
        });
        put(ws, row, 7, F(`IF($B${row}=0,"",($C${row}+$D${row}+$E${row})/$B${row})`, pct(tr.counts.covered + tr.counts.mitigated + tr.counts.detected, tr.total)),
          { font: font({ bold: true }), numFmt: '0%', alignment: { horizontal: 'center' } });
        put(ws, row, 8, F(`IF($B${row}=0,"",$C${row}/$B${row})`, pct(tr.counts.covered, tr.total)), { font: font(), numFmt: '0%', alignment: { horizontal: 'center' } });
        barCells(row);
        row++;
      }
      stackedBarRules(first, row - 1);
    }
    row++;

    // Recommended next actions (ranked at export, live status)
    band(`Recommended next actions (top ${TOP_ACTIONS} per domain, ranked at export)`);
    const actions = (title, list, L, V, openHeader) => {
      text(row, 1, title, { bold: true, size: 11 });
      row++;
      headerRow([title.startsWith('Mitigations') ? 'Mitigation' : 'Detection strategy', 'ID', 'Domain', 'Priority', 'Profile techniques', openHeader, 'Status (live)']);
      if (!list.length) {
        merged(row, 1, 8, 'Nothing to recommend: every mapped control was already in place when this workpaper was exported.', { font: font({ italic: true, color: { argb: C.muted } }) });
        row += 2;
        return;
      }
      const first = row;
      const top = { vertical: 'top' }, mid = { vertical: 'top', horizontal: 'center' };
      for (const a of list) {
        const key = a.key;
        put(ws, row, 1, clean(a.name), { font: font(), alignment: { wrapText: true, vertical: 'top' } });
        put(ws, row, 2, a.id, { font: font({ color: { argb: C.link }, underline: true }), alignment: top });
        put(ws, row, 3, a.domain, { font: font(), alignment: top });
        put(ws, row, 4, a.priority, { font: font({ bold: true }), alignment: mid });
        put(ws, row, 5, a.techN, { font: font(), alignment: mid });
        put(ws, row, 6, F(`IFERROR(INDEX(${L.range('open')},MATCH(${lit(key)},${L.range('key')},0)),"")`, a.openNow), { font: font(), alignment: mid, fill: solid(C.live) });
        put(ws, row, 7, F(`IFERROR(INDEX(${L.range('status')},MATCH(${lit(key)},${L.range('key')},0))&"","")`, a.status), { font: font({ bold: true }), alignment: mid });
        ws.getRow(row).height = 13 * lines(a.name, 36) + 3;   // column A holds ~36 characters per line
        row++;
      }
      ws.addConditionalFormatting({
        ref: `G${first}:G${row - 1}`,
        rules: [{ type: 'cellIs', operator: 'equal', formulae: [lit(V.on)], priority: prio(),
          style: { fill: solid(C.goodTint), font: { color: { argb: C.goodInk }, bold: true } } }],
      });
      row++;
    };
    actions('Mitigations to put in place', P.domains.flatMap(D => D.topMitigations), LM, MIT, 'Unmitigated techniques it addresses (live)');
    actions('Detection strategies to deploy', P.domains.flatMap(D => D.topDetections), LD, DET, 'Undetected techniques it covers (live)');

    // Threat profile
    band('Threat profile');
    headerRow(['Adversary', 'Type', 'ATT&CK ID', 'Domains', 'Techniques contributed', 'Aliases']);
    for (const p of P.profile) {
      put(ws, row, 1, link(p.obj.name, p.obj.url), { font: font({ color: { argb: C.link }, underline: true }) });
      text(row, 2, clean(p.kind));
      text(row, 3, p.obj.attackId || '');
      text(row, 4, p.domains);
      put(ws, row, 5, p.techniques, { font: font(), alignment: { horizontal: 'center' } });
      text(row, 6, clean(p.aliases), { color: { argb: C.muted } });
      row++;
    }
    row++;

    // Scope and method
    band('Scope and method');
    const opts = ctx.opts;
    const platforms = opts.platforms && opts.platforms.size ? [...opts.platforms].join(', ') : 'All platforms';
    const year = P.domains.reduce((y, D) => Math.max(y, D.released ? new Date(D.released).getFullYear() : 0), 0) || new Date().getFullYear();
    const method = [
      ['ATT&CK data', P.domains.map(D => `${D.label} ATT&CK v${D.version}` + (D.released ? ` (released ${new Date(D.released).toLocaleDateString(undefined, { dateStyle: 'medium' })})` : '')).join('; ')],
      ['Scope', `Techniques used by the selected adversaries. Software they use: ${opts.viaSoftware ? 'included' : 'not included'}. Campaigns attributed to selected groups: ${opts.viaCampaigns ? 'included' : 'not included'}. Platforms: ${platforms}.`],
      ['Readiness', 'Share of profile techniques with at least one mitigation In Place or detection strategy Deployed. Fully covered means both. Partial means one of the two.'],
      ['Coverage', 'A technique is mitigated when any mitigation mapped to it on the Control Map sheet is In Place, and detected when any detection strategy mapped to it is Deployed.'],
      ['Priority', 'Ranked by ThreatSim at export: each step picks the control covering the most still-uncovered techniques, weighting techniques used by more selected adversaries higher. Priorities do not change when you edit statuses here; the live columns do. Re-export from ThreatSim to re-rank.'],
      ['Sheets', 'Matrix: techniques by tactic, shaded by usage, with a live status column. Techniques: every profile technique with live coverage. '
        + 'Mitigations and Detections: set statuses here (yellow drop-downs). ICS Assets: assets targeted by ICS techniques and how exposed they are. '
        + 'Telemetry: data components and log sources the detection strategies rely on. Control Map: each technique-to-control link the coverage formulas use.'],
      ['Limits', 'ATT&CK does not enumerate every adversary behaviour, and some techniques have no ATT&CK mitigation. Readiness here measures coverage of ATT&CK mappings, not overall security.'],
      ['ATT&CK licence', `© ${year} The MITRE Corporation. This work is reproduced and distributed with the permission of The MITRE Corporation. `
        + 'License: The MITRE Corporation (MITRE) hereby grants you a non-exclusive, royalty-free license to use ATT&CK® for research, development, and commercial purposes. '
        + 'Any copy you make for such purposes is authorized provided that you reproduce MITRE\'s copyright designation and this license in any such copy. '
        + 'MITRE ATT&CK and ATT&CK are registered trademarks of The MITRE Corporation. Source: https://github.com/mitre-attack/attack-stix-data'],
    ];
    for (const [label, value] of method) {
      text(row, 1, label, { bold: true });
      merged(row, 2, 8, clean(value), { font: font(), alignment: { wrapText: true, vertical: 'top' } }, 15 * lines(value, 95) + 3);
      ws.getCell(row, 1).alignment = { vertical: 'top' };
      row++;
    }

    ws.views = [{ showGridLines: false, activeCell: 'A1' }];
    ws.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } };
  }

  // ---------- workbook ----------

  function build(ExcelJS, ctx) {
    const P = prepare(ctx);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ThreatSim';
    wb.lastModifiedBy = 'ThreatSim';
    wb.created = P.generated;
    wb.modified = P.generated;
    wb.title = 'ATT&CK threat readiness workpaper';
    wb.subject = clean('Threat profile: ' + P_names(ctx));
    wb.calcProperties.fullCalcOnLoad = true;

    // Sheets in tab order; filled once every table's size (and so every range) is known.
    const summary = wb.addWorksheet(SHEET.summary, { properties: { tabColor: { argb: C.header } } });
    const matrices = P.domains.map(D => wb.addWorksheet(matrixSheetName(D.label)));
    const techniques = wb.addWorksheet(SHEET.techniques);
    const mitigations = wb.addWorksheet(SHEET.mitigations, { properties: { tabColor: { argb: C.inputTab } } });
    const detections = wb.addWorksheet(SHEET.detections, { properties: { tabColor: { argb: C.inputTab } } });
    const assets = wb.addWorksheet(SHEET.assets);
    const telemetry = wb.addWorksheet(SHEET.telemetry);
    const map = wb.addWorksheet(SHEET.map);

    const LT = layout(SHEET.techniques, TECH_COLS, P.techniques.length);
    const LM = layout(SHEET.mitigations, MIT_COLS, P.mitigations.length);
    const LD = layout(SHEET.detections, DET_COLS, P.detections.length);
    const LA = layout(SHEET.assets, ASSET_COLS, P.assets.length);
    const LTel = layout(SHEET.telemetry, TEL_COLS, P.telemetry.length);
    const LMap = layout(SHEET.map, MAP_COLS, P.map.length);

    writeSummary(summary, ctx, P, { T: LT, M: LM, D: LD });
    P.domains.forEach((D, i) => writeMatrix(matrices[i], D, ctx, LT));
    writeTechniques(techniques, P, LT, LMap);
    writeControls(mitigations, P.mitigations, LM, LT, 'mit');
    writeControls(detections, P.detections, LD, LT, 'det');
    writeAssets(assets, P, LA, LT);
    writeTelemetry(telemetry, P, LTel, LD);
    writeMap(map, P, LMap, LM, LD);
    return wb;
  }

  // ---------- loading and download ----------

  let loading = null;
  function loadExcelJS() {
    if (global.ExcelJS) return Promise.resolve(global.ExcelJS);
    if (!loading) {
      loading = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = EXCELJS_SRC;
        s.onload = () => (global.ExcelJS ? resolve(global.ExcelJS) : reject(new Error('ExcelJS loaded but did not register')));
        s.onerror = () => reject(new Error('Could not load ' + EXCELJS_SRC));
        document.head.append(s);
      }).catch(e => { loading = null; throw e; });
    }
    return loading;
  }

  function fileName(ctx, when) {
    const slug = P_names(ctx).toLowerCase().replace(/\+\d+ more/, 'and-more').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 48).replace(/-$/, '');
    const pad = n => String(n).padStart(2, '0');
    const stamp = `${when.getFullYear()}${pad(when.getMonth() + 1)}${pad(when.getDate())}-${pad(when.getHours())}${pad(when.getMinutes())}`;
    return `threatsim-workpaper-${slug || 'profile'}-${stamp}.xlsx`;
  }

  /** Builds the workbook and returns { blob, name }. */
  async function create(ctx) {
    if (!ctx.results || !ctx.results.length) throw new Error('Build a threat profile first.');
    const ExcelJS = await loadExcelJS();
    const c = Object.assign({}, ctx, { generatedAt: ctx.generatedAt || new Date() });
    const wb = build(ExcelJS, c);
    const buf = await wb.xlsx.writeBuffer();
    return { blob: new Blob([buf], { type: XLSX_MIME }), name: fileName(c, c.generatedAt) };
  }

  async function download(ctx) {
    const { blob, name } = await create(ctx);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    return name;
  }

  global.Workpaper = { create, download, build, prepare, loadExcelJS, MIT, DET, COVERAGE, SHEET };
})(window);
