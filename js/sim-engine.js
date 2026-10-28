/*
 * sim-engine.js — in-memory analysis for the threat simulator.
 *
 * Global `SimEngine`. Loads a domain from AttackDB into lookup maps, then maps a
 * threat profile (selected groups / software / campaigns) to the techniques it
 * uses, and from there to mitigations, detection strategies, telemetry and
 * (ICS) targeted assets. No DOM access here.
 */
(function (global) {
  'use strict';

  /** Build lookup maps for one domain. */
  async function loadModel(domain) {
    const { meta, objects, relationships } = await AttackDB.loadDomain(domain);
    if (!meta) return null;

    const byId = new Map();
    const byKind = {};
    for (const o of objects) {
      byId.set(o.id, o);
      if (!o.active) continue;
      (byKind[o.kind] = byKind[o.kind] || []).push(o);
    }

    const relsFrom = new Map();
    const relsTo = new Map();
    for (const r of relationships) {
      if (!r.active) continue;
      if (!relsFrom.has(r.source)) relsFrom.set(r.source, []);
      if (!relsTo.has(r.target)) relsTo.set(r.target, []);
      relsFrom.get(r.source).push(r);
      relsTo.get(r.target).push(r);
    }

    // Tactic columns in matrix order.
    const matrix = (byKind.matrix || [])[0];
    const tactics = matrix
      ? matrix.tacticRefs.map(id => byId.get(id)).filter(t => t && t.active)
      : (byKind.tactic || []);

    const model = { domain, meta, byId, byKind, relsFrom, relsTo, tactics };
    model.from = (id, relType, kind) => (relsFrom.get(id) || [])
      .filter(r => r.relType === relType && (!kind || r.targetKind === kind));
    model.to = (id, relType, kind) => (relsTo.get(id) || [])
      .filter(r => r.relType === relType && (!kind || r.sourceKind === kind));
    return model;
  }

  /** Technique count per entity (direct "uses" only), for the picker lists. */
  function directTechniqueCount(model, id) {
    return model.from(id, 'uses', 'technique').length;
  }

  function platformMatch(platforms, filter) {
    if (!filter || !filter.size) return true;
    if (!platforms || !platforms.length || platforms.includes('None')) return true;
    return platforms.some(p => filter.has(p));
  }

  /**
   * Analyze one domain.
   *  selection: { groups:Set, software:Set, campaigns:Set } of STIX ids
   *  opts: { viaSoftware, viaCampaigns, platforms:Set }
   */
  function analyzeDomain(model, selection, opts) {
    const { byId } = model;
    const techniques = new Map();   // techId -> { tech, roots:Set, evidence:[] }
    const roots = [];               // selected entities found in this domain

    const addUse = (rootId, rel, via) => {
      const tech = byId.get(rel.target);
      if (!tech || !tech.active) return;
      let t = techniques.get(tech.id);
      if (!t) techniques.set(tech.id, t = { tech, roots: new Set(), evidence: [] });
      t.roots.add(rootId);
      t.evidence.push({ root: rootId, via, description: rel.description, sourceId: rel.source });
    };

    const softwareOf = id => model.from(id, 'uses', 'software').map(r => byId.get(r.target)).filter(s => s && s.active);

    const walk = (rootId, entityId, via) => {
      for (const r of model.from(entityId, 'uses', 'technique')) addUse(rootId, r, via);
      if (opts.viaSoftware) {
        for (const sw of softwareOf(entityId)) {
          for (const r of model.from(sw.id, 'uses', 'technique')) addUse(rootId, r, [...via, sw.id]);
        }
      }
    };

    for (const kind of ['groups', 'campaigns', 'software']) {
      for (const id of selection[kind]) {
        const ent = byId.get(id);
        if (!ent || !ent.active) continue;
        roots.push(ent);
        walk(id, id, []);
        if (kind === 'groups' && opts.viaCampaigns) {
          for (const r of model.to(id, 'attributed-to', 'campaign')) {
            const camp = byId.get(r.source);
            if (camp && camp.active) walk(id, camp.id, [camp.id]);
          }
        }
      }
    }

    // Platform filter
    for (const [id, t] of techniques) {
      if (!platformMatch(t.tech.platforms, opts.platforms)) techniques.delete(id);
    }

    // Mitigations, detection strategies, assets per technique
    const mitigations = new Map();  // id -> { obj, techniques:Set, guidance: Map techId->desc }
    const detections = new Map();   // id -> { obj, techniques:Set, analytics:[], dataComponents:Set }
    const dataComponents = new Map(); // id -> { obj, techniques:Set, strategies:Set, logSources: Map name->Set(channel) }
    const assets = new Map();       // id -> { obj, techniques:Set }

    for (const [techId, t] of techniques) {
      t.mitigations = [];
      t.detections = [];
      t.assets = [];

      for (const r of model.to(techId, 'mitigates', 'mitigation')) {
        const m = byId.get(r.source);
        let e = mitigations.get(m.id);
        if (!e) mitigations.set(m.id, e = { obj: m, techniques: new Set(), guidance: new Map() });
        e.techniques.add(techId);
        e.guidance.set(techId, r.description);
        t.mitigations.push(m.id);
      }

      for (const r of model.to(techId, 'detects', 'detection-strategy')) {
        const ds = byId.get(r.source);
        let e = detections.get(ds.id);
        if (!e) {
          const analytics = (ds.analyticRefs || []).map(id => byId.get(id))
            .filter(a => a && a.active && platformMatch(a.platforms, opts.platforms));
          const dcs = new Set();
          for (const a of analytics) for (const l of a.logSources || []) if (l.dataComponent) dcs.add(l.dataComponent);
          detections.set(ds.id, e = { obj: ds, techniques: new Set(), analytics, dataComponents: dcs });
        }
        e.techniques.add(techId);
        t.detections.push(ds.id);
      }

      for (const r of model.from(techId, 'targets', 'asset')) {
        const a = byId.get(r.target);
        let e = assets.get(a.id);
        if (!e) assets.set(a.id, e = { obj: a, techniques: new Set() });
        e.techniques.add(techId);
        t.assets.push(a.id);
      }
    }

    for (const [dsId, d] of detections) {
      for (const a of d.analytics) {
        for (const l of a.logSources || []) {
          const dc = byId.get(l.dataComponent);
          if (!dc) continue;
          let e = dataComponents.get(dc.id);
          if (!e) dataComponents.set(dc.id, e = { obj: dc, techniques: new Set(), strategies: new Set(), logSources: new Map() });
          for (const tid of d.techniques) e.techniques.add(tid);
          e.strategies.add(dsId);
          if (l.name) {
            if (!e.logSources.has(l.name)) e.logSources.set(l.name, new Set());
            if (l.channel && l.channel !== 'None') e.logSources.get(l.name).add(l.channel);
          }
        }
      }
    }

    return { domain: model.domain, model, roots, techniques, mitigations, detections, dataComponents, assets };
  }

  /**
   * Coverage under a defensive posture.
   *  inPlace: { mitigations:Set, detections:Set } of keys "<domain>|<id>"
   * Sets t.status on every technique: 'covered' | 'mitigated' | 'detected' | 'gap'.
   */
  function applyPosture(result, inPlace) {
    const d = result.domain;
    const counts = { covered: 0, mitigated: 0, detected: 0, gap: 0 };
    for (const t of result.techniques.values()) {
      const m = t.mitigations.some(id => inPlace.mitigations.has(d + '|' + id));
      const x = t.detections.some(id => inPlace.detections.has(d + '|' + id));
      t.isMitigated = m;
      t.isDetected = x;
      t.status = m && x ? 'covered' : m ? 'mitigated' : x ? 'detected' : 'gap';
      counts[t.status]++;
    }
    return counts;
  }

  /**
   * Greedy priority order: repeatedly pick the control whose still-uncovered
   * techniques carry the most weight, where a technique weighs as much as the
   * number of selected adversaries using it. Controls already in place count as
   * picked first. Returns Map id -> { rank, adds, cumulative } (rank null when a
   * control is in place or adds nothing).
   */
  function prioritize(entries, inPlaceKeys, domain, techniques) {
    const totalTechniques = techniques.size;
    const weight = tid => (techniques.get(tid) ? techniques.get(tid).roots.size : 1);
    const covered = new Set();
    const out = new Map();
    const remaining = [];
    for (const [id, e] of entries) {
      if (inPlaceKeys.has(domain + '|' + id)) {
        for (const t of e.techniques) covered.add(t);
      } else {
        remaining.push([id, e]);
      }
    }
    for (const [id, e] of entries) {
      if (inPlaceKeys.has(domain + '|' + id)) out.set(id, { rank: null, adds: 0, cumulative: covered.size / (totalTechniques || 1) });
    }
    let rank = 1;
    while (remaining.length) {
      let best = -1, bestScore = -1, bestAdds = 0;
      for (let i = 0; i < remaining.length; i++) {
        let score = 0, adds = 0;
        for (const t of remaining[i][1].techniques) if (!covered.has(t)) { score += weight(t); adds++; }
        if (score > bestScore || (score === bestScore && remaining[i][1].techniques.size > remaining[best][1].techniques.size)) {
          best = i; bestScore = score; bestAdds = adds;
        }
      }
      const [id, e] = remaining.splice(best, 1)[0];
      for (const t of e.techniques) covered.add(t);
      out.set(id, { rank: bestAdds > 0 ? rank++ : null, adds: bestAdds, cumulative: covered.size / (totalTechniques || 1) });
    }
    return out;
  }

  /** Techniques arranged for the matrix: per tactic, parents with nested sub-techniques. */
  function matrixColumns(result, { showAll }) {
    const { model, techniques } = result;
    const used = techniques;
    const allTech = model.byKind.technique || [];
    const subsByParent = new Map();
    for (const t of allTech) {
      if (t.isSubtechnique && t.parent) {
        if (!subsByParent.has(t.parent)) subsByParent.set(t.parent, []);
        subsByParent.get(t.parent).push(t);
      }
    }
    const byAttackId = (a, b) => String(a.attackId).localeCompare(String(b.attackId), undefined, { numeric: true });

    return model.tactics.map(tactic => {
      const parents = allTech.filter(t => !t.isSubtechnique && (t.tactics || []).includes(tactic.shortname));
      const rows = [];
      for (const p of parents.sort((a, b) => a.name.localeCompare(b.name))) {
        const subs = (subsByParent.get(p.id) || []).filter(s => (s.tactics || []).includes(tactic.shortname)).sort(byAttackId);
        const usedSubs = subs.filter(s => used.has(s.id));
        if (!showAll && !used.has(p.id) && !usedSubs.length) continue;
        rows.push({ tech: p, used: used.get(p.id) || null, subs: (showAll ? subs : usedSubs).map(s => ({ tech: s, used: used.get(s.id) || null })) });
      }
      const usedCount = rows.reduce((n, r) => n + (r.used ? 1 : 0) + r.subs.filter(s => s.used).length, 0);
      return { tactic, rows, usedCount };
    });
  }

  global.SimEngine = { loadModel, analyzeDomain, applyPosture, prioritize, matrixColumns, directTechniqueCount, platformMatch };
})(window);
