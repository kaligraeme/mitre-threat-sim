/*
 * attack-db.js — local MITRE ATT&CK data instance (IndexedDB).
 *
 * Shared by every page of the threat simulator. Exposes a global `AttackDB`
 * (classic script, so pages also work when opened straight from file://).
 *
 * Stores
 *   meta           one row per domain: version, source URL, pulled-at, counts
 *   objects        normalized STIX objects, key "<domain>|<stix id>"
 *   relationships  normalized STIX relationships, key "<domain>|<stix id>"
 *   raw            original STIX objects (incl. relationships), key "<domain>|<stix id>"
 *   history        append-only log of pulls
 */
(function (global) {
  'use strict';

  const DB_NAME = 'mitre-attack-db';
  const DB_VERSION = 1;

  const INDEX_URL = 'https://raw.githubusercontent.com/mitre-attack/attack-stix-data/master/index.json';

  // Domains the simulator works with. Mobile is available upstream but off by default.
  const DOMAINS = [
    { id: 'enterprise-attack', label: 'Enterprise', collection: 'Enterprise ATT&CK', enabled: true },
    { id: 'ics-attack',        label: 'ICS',        collection: 'ICS ATT&CK',        enabled: true },
    { id: 'mobile-attack',     label: 'Mobile',     collection: 'Mobile ATT&CK',     enabled: false },
  ];

  // STIX type -> simulator "kind"
  const KIND_BY_TYPE = {
    'attack-pattern': 'technique',
    'intrusion-set': 'group',
    'malware': 'software',
    'tool': 'software',
    'campaign': 'campaign',
    'course-of-action': 'mitigation',
    'x-mitre-detection-strategy': 'detection-strategy',
    'x-mitre-analytic': 'analytic',
    'x-mitre-data-component': 'data-component',
    'x-mitre-data-source': 'data-source',
    'x-mitre-tactic': 'tactic',
    'x-mitre-asset': 'asset',
    'x-mitre-matrix': 'matrix',
    'x-mitre-collection': 'collection',
  };

  const KIND_LABELS = {
    'technique': 'Techniques',
    'group': 'Groups',
    'software': 'Software',
    'campaign': 'Campaigns',
    'mitigation': 'Mitigations',
    'detection-strategy': 'Detection strategies',
    'analytic': 'Analytics',
    'data-component': 'Data components',
    'data-source': 'Data sources',
    'tactic': 'Tactics',
    'asset': 'Assets',
    'matrix': 'Matrices',
  };

  let dbPromise = null;

  function req(r) {
    return new Promise((resolve, reject) => {
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }

  function txDone(tx) {
    return new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));
    });
  }

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, DB_VERSION);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains('meta')) {
          db.createObjectStore('meta', { keyPath: 'domain' });
        }
        if (!db.objectStoreNames.contains('objects')) {
          const s = db.createObjectStore('objects', { keyPath: 'key' });
          s.createIndex('domain', 'domain');
          s.createIndex('domain_kind', ['domain', 'kind']);
          s.createIndex('attackId', 'attackId');
          s.createIndex('id', 'id');
        }
        if (!db.objectStoreNames.contains('relationships')) {
          const s = db.createObjectStore('relationships', { keyPath: 'key' });
          s.createIndex('domain', 'domain');
          s.createIndex('domain_source', ['domain', 'source']);
          s.createIndex('domain_target', ['domain', 'target']);
          s.createIndex('domain_relType', ['domain', 'relType']);
        }
        if (!db.objectStoreNames.contains('raw')) {
          db.createObjectStore('raw', { keyPath: 'key' });
        }
        if (!db.objectStoreNames.contains('history')) {
          db.createObjectStore('history', { keyPath: 'seq', autoIncrement: true });
        }
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.onblocked = () => reject(new Error('Database upgrade blocked — close other tabs of this app.'));
    });
    return dbPromise;
  }

  function domainRange(domain) {
    return IDBKeyRange.bound(domain + '|', domain + '|\uffff');
  }

  // ---------- normalization ----------

  function mitreRef(o) {
    return (o.external_references || []).find(r => r.source_name === 'mitre-attack' && r.external_id) || null;
  }

  function normalizeObject(domain, o) {
    const ref = mitreRef(o);
    const rec = {
      key: domain + '|' + o.id,
      domain,
      id: o.id,
      type: o.type,
      kind: KIND_BY_TYPE[o.type] || 'other',
      attackId: ref ? ref.external_id : null,
      url: ref ? ref.url || null : null,
      name: o.name || null,
      description: o.description || '',
      created: o.created || null,
      modified: o.modified || null,
      version: o.x_mitre_version || null,
      revoked: !!o.revoked,
      deprecated: !!o.x_mitre_deprecated,
      domains: o.x_mitre_domains || [domain],
      platforms: o.x_mitre_platforms || [],
    };
    rec.active = !rec.revoked && !rec.deprecated;

    switch (o.type) {
      case 'attack-pattern':
        rec.isSubtechnique = !!o.x_mitre_is_subtechnique;
        rec.tactics = (o.kill_chain_phases || []).map(p => p.phase_name);
        rec.parent = null; // filled from subtechnique-of relationships
        break;
      case 'intrusion-set':
      case 'campaign':
        rec.aliases = o.aliases || [];
        if (o.type === 'campaign') {
          rec.firstSeen = o.first_seen || null;
          rec.lastSeen = o.last_seen || null;
        }
        break;
      case 'malware':
      case 'tool':
        rec.softwareType = o.type;
        rec.aliases = o.x_mitre_aliases || [];
        break;
      case 'x-mitre-tactic':
        rec.shortname = o.x_mitre_shortname || null;
        break;
      case 'x-mitre-matrix':
        rec.tacticRefs = o.tactic_refs || [];
        break;
      case 'x-mitre-detection-strategy':
        rec.analyticRefs = o.x_mitre_analytic_refs || [];
        break;
      case 'x-mitre-analytic':
        rec.logSources = (o.x_mitre_log_source_references || []).map(l => ({
          dataComponent: l.x_mitre_data_component_ref || null,
          name: l.name || null,
          channel: l.channel || null,
        }));
        break;
      case 'x-mitre-data-component':
        rec.logSources = o.x_mitre_log_sources || [];
        rec.dataSource = o.x_mitre_data_source_ref || null;
        break;
      case 'x-mitre-asset':
        rec.sectors = o.x_mitre_sectors || [];
        break;
    }
    return rec;
  }

  function normalizeRelationship(domain, r, objIndex) {
    const s = objIndex.get(r.source_ref);
    const t = objIndex.get(r.target_ref);
    return {
      key: domain + '|' + r.id,
      domain,
      id: r.id,
      relType: r.relationship_type,
      source: r.source_ref,
      target: r.target_ref,
      sourceKind: s ? s.kind : null,
      targetKind: t ? t.kind : null,
      description: r.description || '',
      modified: r.modified || null,
      revoked: !!r.revoked,
      deprecated: !!r.x_mitre_deprecated,
      // usable for analysis: relationship and both endpoints are current
      active: !r.revoked && !r.x_mitre_deprecated && !!s && !!t && s.active && t.active,
    };
  }

  /** Turn a STIX bundle into {objects, relationships, counts}. */
  function normalizeBundle(domain, bundle) {
    const raw = bundle.objects || [];
    const objects = [];
    const rawRels = [];
    const objIndex = new Map();

    for (const o of raw) {
      if (o.type === 'relationship') { rawRels.push(o); continue; }
      const rec = normalizeObject(domain, o);
      objects.push(rec);
      objIndex.set(o.id, rec);
    }

    const relationships = rawRels.map(r => normalizeRelationship(domain, r, objIndex));

    for (const r of relationships) {
      if (r.relType === 'subtechnique-of' && !r.revoked && !r.deprecated) {
        const child = objIndex.get(r.source);
        if (child) child.parent = r.target;
      }
    }

    const counts = {};
    for (const o of objects) {
      if (!o.active) continue;
      counts[o.kind] = (counts[o.kind] || 0) + 1;
    }
    counts.relationships = relationships.filter(r => r.active).length;
    counts.totalObjects = objects.length;
    counts.totalRelationships = relationships.length;

    return { objects, relationships, raw, counts };
  }

  // ---------- writes ----------

  /** Atomically replace everything stored for `domain`. */
  async function replaceDomain(domain, normalized, metaRow) {
    const db = await open();
    const tx = db.transaction(['objects', 'relationships', 'raw', 'meta', 'history'], 'readwrite');
    const objects = tx.objectStore('objects');
    const rels = tx.objectStore('relationships');
    const raw = tx.objectStore('raw');
    objects.delete(domainRange(domain));
    rels.delete(domainRange(domain));
    raw.delete(domainRange(domain));
    for (const o of normalized.objects) objects.put(o);
    for (const r of normalized.relationships) rels.put(r);
    for (const o of normalized.raw) raw.put({ key: domain + '|' + o.id, stix: o });
    tx.objectStore('meta').put(metaRow);
    tx.objectStore('history').add({
      domain,
      version: metaRow.version,
      previousVersion: metaRow.previousVersion || null,
      pulledAt: metaRow.pulledAt,
      counts: metaRow.counts,
    });
    await txDone(tx);
  }

  async function clearAll() {
    const db = await open();
    const stores = ['objects', 'relationships', 'raw', 'meta', 'history'];
    const tx = db.transaction(stores, 'readwrite');
    for (const n of stores) tx.objectStore(n).clear();
    await txDone(tx);
  }

  // ---------- reads (used by the dashboard pages) ----------

  async function getMeta(domain) {
    const db = await open();
    return req(db.transaction('meta').objectStore('meta').get(domain));
  }

  async function getAllMeta() {
    const db = await open();
    return req(db.transaction('meta').objectStore('meta').getAll());
  }

  async function getHistory(limit = 50) {
    const db = await open();
    const all = await req(db.transaction('history').objectStore('history').getAll());
    return all.reverse().slice(0, limit);
  }

  /** Objects of one kind in a domain. Pass {includeInactive:true} to keep revoked/deprecated. */
  async function getObjects(domain, kind, opts = {}) {
    const db = await open();
    const idx = db.transaction('objects').objectStore('objects').index('domain_kind');
    const rows = await req(idx.getAll([domain, kind]));
    return opts.includeInactive ? rows : rows.filter(r => r.active);
  }

  async function getObject(domain, stixId) {
    const db = await open();
    return req(db.transaction('objects').objectStore('objects').get(domain + '|' + stixId));
  }

  /** Everything stored for one domain, for pages that work in memory. */
  async function loadDomain(domain) {
    const db = await open();
    const tx = db.transaction(['meta', 'objects', 'relationships']);
    const range = domainRange(domain);
    const [meta, objects, relationships] = await Promise.all([
      req(tx.objectStore('meta').get(domain)),
      req(tx.objectStore('objects').getAll(range)),
      req(tx.objectStore('relationships').getAll(range)),
    ]);
    return { meta, objects, relationships };
  }

  /** Original STIX object (external references, citations, etc.). */
  async function getRaw(domain, stixId) {
    const db = await open();
    const row = await req(db.transaction('raw').objectStore('raw').get(domain + '|' + stixId));
    return row ? row.stix : null;
  }

  /** Look up by ATT&CK ID (e.g. "T1059.001", "G0007") across all or one domain. */
  async function findByAttackId(attackId, domain) {
    const db = await open();
    const rows = await req(db.transaction('objects').objectStore('objects').index('attackId').getAll(attackId));
    return domain ? rows.filter(r => r.domain === domain) : rows;
  }

  async function relationshipsFrom(domain, stixId, relType) {
    const db = await open();
    const idx = db.transaction('relationships').objectStore('relationships').index('domain_source');
    const rows = await req(idx.getAll([domain, stixId]));
    return rows.filter(r => r.active && (!relType || r.relType === relType));
  }

  async function relationshipsTo(domain, stixId, relType) {
    const db = await open();
    const idx = db.transaction('relationships').objectStore('relationships').index('domain_target');
    const rows = await req(idx.getAll([domain, stixId]));
    return rows.filter(r => r.active && (!relType || r.relType === relType));
  }

  async function relationshipsByType(domain, relType) {
    const db = await open();
    const idx = db.transaction('relationships').objectStore('relationships').index('domain_relType');
    const rows = await req(idx.getAll([domain, relType]));
    return rows.filter(r => r.active);
  }

  /** Case-insensitive search over name / ATT&CK ID / aliases. */
  async function search(text, { domain, kinds, limit = 50 } = {}) {
    const q = text.trim().toLowerCase();
    if (!q) return [];
    const db = await open();
    const store = db.transaction('objects').objectStore('objects');
    const rows = await req(domain ? store.index('domain').getAll(domain) : store.getAll());
    const out = [];
    for (const o of rows) {
      if (!o.active || !KIND_LABELS[o.kind] || (kinds && !kinds.includes(o.kind))) continue;
      const hay = [o.name, o.attackId, ...(o.aliases || [])].filter(Boolean).join(' ').toLowerCase();
      if (hay.includes(q)) out.push(o);
    }
    // exact ID / name matches first, then by ATT&CK ID
    const rank = o => ((o.attackId || '').toLowerCase() === q || (o.name || '').toLowerCase() === q) ? 0 : 1;
    out.sort((x, y) => rank(x) - rank(y) || String(x.attackId).localeCompare(String(y.attackId)));
    return out.slice(0, limit);
  }

  // ---------- upstream ----------

  function compareVersions(a, b) {
    const pa = String(a).split('.').map(Number);
    const pb = String(b).split('.').map(Number);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
      const d = (pa[i] || 0) - (pb[i] || 0);
      if (d) return d;
    }
    return 0;
  }

  /** Fetch MITRE's index.json and return {domainId: {latest, versions[]}} */
  async function fetchUpstreamIndex() {
    const res = await fetch(INDEX_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('index.json: HTTP ' + res.status);
    const idx = await res.json();
    const out = {};
    for (const d of DOMAINS) {
      const col = (idx.collections || []).find(c => c.name === d.collection);
      if (!col) continue;
      const versions = [...col.versions].sort((a, b) => compareVersions(b.version, a.version));
      out[d.id] = { collectionId: col.id, latest: versions[0], versions };
    }
    return { modified: idx.modified, domains: out };
  }

  global.AttackDB = {
    DB_NAME, INDEX_URL, DOMAINS, KIND_BY_TYPE, KIND_LABELS,
    open, normalizeBundle, replaceDomain, clearAll,
    getMeta, getAllMeta, getHistory,
    getObjects, getObject, getRaw, findByAttackId, loadDomain,
    relationshipsFrom, relationshipsTo, relationshipsByType, search,
    fetchUpstreamIndex, compareVersions,
  };
})(window);
