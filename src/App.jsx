import React, { useState } from 'react';
import { supabase, cea } from './supabaseClient';
import * as XLSX from 'xlsx';
import { LayoutGrid, BookOpen, Building2, Calculator, GitCompare, Truck, Users, Search, LogOut, MapPin, Clock, Calendar as CalendarIcon } from 'lucide-react';

const FONT = "'Poppins', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif";

// Design token ufficiali Desearq (estratti da DSM — desearq-design-tokens.css)
const C = {
  maroon: '#801430',
  accentHover: '#650F26',
  accentSoft: '#F3DCE1',
  black: '#171717',
  sidebar: '#F6F4EF',
  sidebarHover: '#EDEAE1',
  darkGray: '#6B6B6B',
  midGray: '#6B6B6B',
  gray: '#A3A3A3',
  lightGray: '#C9C9C9',
  paleGray: '#E7E4DC',
  borderStrong: '#D8D4CB',
  bg: '#F6F4EF',
  surfaceSubtle: '#FAF9F6',
  white: '#FFFFFF',
  success: '#2E7D4F',
};
const PAGE_GRADIENT = 'radial-gradient(120% 90% at 85% 0%, #FBF9F4 0%, #F6F4EF 45%, #F3F0E9 100%)';

const NAV_ITEMS = [
  { key: 'dashboard', label: 'Dashboard', Icon: LayoutGrid },
  { key: 'listino', label: 'Listino prezzi', Icon: BookOpen },
  { key: 'progetti', label: 'Progetti', Icon: Building2 },
  { key: 'computi', label: 'Computi', Icon: Calculator },
  { key: 'confronto', label: 'Confronto revisioni', Icon: GitCompare },
  { key: 'fornitori', label: 'Fornitori', Icon: Truck },
  { key: 'team', label: 'Team', Icon: Users },
];

const STATUS_OPTIONS = ['In attesa di approvazione', 'Approvato', 'In fase di cantiere'];
const statusTone = {
  'In attesa di approvazione': 'orange',
  'Approvato': 'teal',
  'In fase di cantiere': 'gray',
};
const nextStatus = (s) => STATUS_OPTIONS[Math.min(STATUS_OPTIONS.indexOf(s) + 1, STATUS_OPTIONS.length - 1)];
const latestStatus = (project) => project.revisions[project.revisions.length - 1]?.status || STATUS_OPTIONS[0];

const parseEuro = (v) => parseFloat(String(v).replace(/\./g, '').replace(',', '.')) || 0;
const formatEuro = (n) => n.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';

// Calcola il prezzo cliente: un numero fisso, oppure una formula che inizia con "="
// e può usare la parola "impresa" per riferirsi al prezzo impresa (es. =impresa*1.3, =impresa+50).
function evalClientPrice(formula, impresaPrice) {
  const str = String(formula || '').trim();
  if (!str) return impresaPrice;
  if (!str.startsWith('=')) return parseEuro(str);
  let expr = str.slice(1).replace(/impresa/gi, String(impresaPrice)).replace(/,/g, '.');
  if (!/^[0-9+\-*/().\s]+$/.test(expr)) return impresaPrice;
  try {
    const result = Function(`"use strict"; return (${expr});`)();
    return typeof result === 'number' && isFinite(result) ? result : impresaPrice;
  } catch {
    return impresaPrice;
  }
}
const nowLabel = () => new Date().toLocaleString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const sumImpresa = (items) => (items || []).reduce((sum, it) => sum + parseEuro(it.unitPriceImpresa) * parseEuro(it.qty), 0);
const sumCliente = (items) => (items || []).reduce((sum, it) => sum + parseEuro(it.unitPriceCliente) * parseEuro(it.qty), 0);
const getVatInfo = (revision) => {
  const rate = revision?.vatRate !== undefined && revision?.vatRate !== null ? Number(revision.vatRate) : 22;
  const label = revision?.vatLabel && revision.vatLabel.trim() ? revision.vatLabel : `IVA ${rate}%`;
  return { rate, label };
};

// Aggiunge una voce (es. una fornitura) alla revisione scelta di un progetto.
// Se la revisione scelta è l'ultima, aggiorna sul posto; altrimenti crea automaticamente
// una nuova versione (fork) lasciando quella aperta intatta, come per ogni altra modifica.
function addItemToProjectRevision(project, revisionId, newItem) {
  const revisions = project.revisions;
  const latest = revisions[revisions.length - 1];
  const target = revisions.find((r) => r.id === revisionId) || latest;
  const isLatest = target.id === latest.id;
  const newItems = [...(target.items || []), newItem];
  const realItems = newItems.filter((it) => it.type !== 'subtotal');
  const total = formatEuro(sumImpresa(realItems));
  const totalCliente = formatEuro(sumCliente(realItems));
  if (isLatest) {
    return {
      ...project,
      revisions: revisions.map((r) => (r.id === target.id ? { ...r, items: newItems, dateModified: nowLabel(), total, totalCliente } : r)),
      value: total,
    };
  }
  const newRev = { ...target, id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: newItems, total, totalCliente };
  return { ...project, revisions: [...revisions, newRev], value: total };
}

function computeItemsDiff(itemsA, itemsB) {
  const mapA = Object.fromEntries((itemsA || []).map((it) => [it.code, it]));
  const mapB = Object.fromEntries((itemsB || []).map((it) => [it.code, it]));
  const codes = Array.from(new Set([...Object.keys(mapA), ...Object.keys(mapB)]));
  return codes.map((code) => {
    const a = mapA[code];
    const b = mapB[code];
    const ref = b || a;
    if (a && b) {
      const totalA = parseEuro(a.unitPriceImpresa) * parseEuro(a.qty);
      const totalB = parseEuro(b.unitPriceImpresa) * parseEuro(b.qty);
      const changed = a.qty !== b.qty || a.unitPriceImpresa !== b.unitPriceImpresa;
      const diff = totalB - totalA;
      return {
        code, desc: ref.desc, sezione: ref.macro || '—',
        esito: changed ? 'Modificata' : 'Invariata',
        qtyBefore: a.qty, qtyAfter: b.qty,
        priceBefore: `${a.unitPriceImpresa} €`, priceAfter: `${b.unitPriceImpresa} €`,
        variation: diff === 0 ? '—' : `${diff > 0 ? '+' : ''}${formatEuro(diff)}`,
        highlight: changed,
      };
    }
    if (b && !a) {
      const totalB = parseEuro(b.unitPriceImpresa) * parseEuro(b.qty);
      return {
        code, desc: ref.desc, sezione: ref.macro || '—', esito: 'Aggiunta',
        qtyBefore: '—', qtyAfter: b.qty, priceBefore: '—', priceAfter: `${b.unitPriceImpresa} €`,
        variation: `+${formatEuro(totalB)}`, highlight: true,
      };
    }
    const totalA = parseEuro(a.unitPriceImpresa) * parseEuro(a.qty);
    return {
      code, desc: ref.desc, sezione: ref.macro || '—', esito: 'Rimossa',
      qtyBefore: a.qty, qtyAfter: '—', priceBefore: `${a.unitPriceImpresa} €`, priceAfter: '—',
      variation: `-${formatEuro(totalA)}`, highlight: true,
    };
  });
}

function flattenListino(listino) {
  if (!listino) return [];
  const out = [];
  withCodes(listino.macros).forEach((m) => m.categorie.forEach((c) => c.sottocategorie.forEach((s) => {
    s.voci.forEach((v) => {
      const impresa = parseEuro(v.priceImpresa);
      const cliente = evalClientPrice(v.priceCliente, impresa);
      out.push({ ...v, macro: m.name, categoria: c.name, sottocategoria: s.name, impresaValue: impresa, clienteValue: cliente });
    });
  })));
  return out;
}

const PROJECTS = [
  {
    id: 1, name: 'Residenza Aurora', client: 'EdilNova S.r.l.', value: '23.422,85 €', items: 4,
    team: [{ name: 'Arch. Bianchi', role: 'Direttore lavori' }, { name: 'Ing. Rossi', role: 'Strutturista' }],
    header: { descrizione: 'Ristrutturazione integrale unità residenziale', ubicazione: 'Via degli Ontani 12, Mantova' },
    revisions: [
      {
        id: 1, label: 'Revisione 1', dateCreated: '08 lug 2026, 10:20', dateModified: '08 lug 2026, 10:20', total: '12.604,00 €', totalCliente: '16.385,20 €', status: 'Approvato',
        items: [
          { id: 101, code: 'ED.MUR.01.010', desc: 'Muratura perimetrale piano terra', unit: 'm²', unitPriceImpresa: '68,50', unitPriceCliente: '89,05', qty: '184', macro: 'Edilizia e strutture' },
        ],
      },
      {
        id: 2, label: 'Revisione 2', dateCreated: '22 lug 2026, 16:45', dateModified: '22 lug 2026, 16:45', total: '15.628,00 €', totalCliente: '20.165,20 €', status: 'In attesa di approvazione',
        items: [
          { id: 101, code: 'ED.MUR.01.010', desc: 'Muratura perimetrale piano terra', unit: 'm²', unitPriceImpresa: '68,50', unitPriceCliente: '89,05', qty: '184', macro: 'Edilizia e strutture' },
          { id: 102, code: 'IT.EL.01.005', desc: 'Punti luce appartamenti', unit: 'cadauna', unitPriceImpresa: '72,00', unitPriceCliente: '90,00', qty: '42', macro: 'Impianti tecnologici' },
        ],
      },
    ],
  },
  {
    id: 2, name: 'Villa Bellavista', client: 'Famiglia Rinaldi', value: '19.548,85 €', items: 2,
    team: [{ name: 'Arch. Bianchi', role: 'Direttore lavori' }],
    header: { descrizione: '', ubicazione: '' },
    revisions: [
      { id: 1, label: 'Revisione 1', dateCreated: '02 lug 2026, 09:00', dateModified: '02 lug 2026, 09:00', total: '19.548,85 €', status: 'In fase di cantiere', items: [] },
    ],
  },
  {
    id: 3, name: 'Riqualificazione Via Roma', client: 'Comune di Mantova', value: '4274,00 €', items: 2,
    team: [{ name: 'Geom. Verdi', role: 'Direttore lavori' }],
    header: { descrizione: '', ubicazione: '' },
    revisions: [
      { id: 1, label: 'Revisione 1', dateCreated: '15 giu 2026, 11:30', dateModified: '15 giu 2026, 11:30', total: '3.980,00 €', status: 'Approvato', items: [] },
      { id: 2, label: 'Revisione 2', dateCreated: '10 lug 2026, 14:10', dateModified: '10 lug 2026, 14:10', total: '4274,00 €', status: 'In attesa di approvazione', items: [] },
    ],
  },
];


const COMPUTO_SECTIONS = [
  {
    name: 'Edilizia e strutture', color: C.maroon,
    items: [{ code: 'ED.MUR.01.010', desc: 'Muratura perimetrale piano terra', qty: '184', unit: 'm²', price: '68,5', total: '12.604,00 €' }],
    subtotal: '12.604,00 €',
  },
  {
    name: 'Impianti tecnologici', color: C.maroon,
    items: [{ code: 'IT.EL.01.005', desc: 'Punti luce appartamenti', qty: '42', unit: 'cad', price: '72', total: '3024,00 €' }],
    subtotal: '3024,00 €',
  },
  {
    name: 'Finiture e opere esterne', color: C.darkGray,
    items: [{ code: 'FO.FIN.01.010', desc: 'Pavimento zona giorno', qty: '126,5', unit: 'm²', price: '54,9', total: '6944,85 €' }],
    subtotal: '6944,85 €',
  },
  {
    name: 'Extra', color: C.black, final: true,
    items: [{ code: 'EX.001', desc: 'Assistenza tecnica fuori standard', qty: '1', unit: 'corpo', price: '850', total: '850,00 €' }],
    subtotal: '850,00 €',
  },
];

const REVISIONS_TABLE = [
  { esito: 'Invariata', sezione: 'Edilizia e strutture', code: 'ED.MUR.01.010', desc: 'Muratura perimetrale piano terra', qtyBefore: '184', qtyAfter: '184', priceBefore: '68,50 €', priceAfter: '68,50 €', variation: '—', highlight: false },
  { esito: 'Invariata', sezione: 'Impianti tecnologici', code: 'IT.EL.01.005', desc: 'Punti luce appartamenti', qtyBefore: '42', qtyAfter: '42', priceBefore: '72,00 €', priceAfter: '72,00 €', variation: '—', highlight: false },
  { esito: 'Invariata', sezione: 'Finiture e opere esterne', code: 'FO.FIN.01.010', desc: 'Pavimento zona giorno', qtyBefore: '126,5', qtyAfter: '126,5', priceBefore: '54,90 €', priceAfter: '54,90 €', variation: '—', highlight: false },
  { esito: 'Aggiunta', sezione: 'Extra', code: 'EX.001', desc: 'Assistenza tecnica fuori standard', qtyBefore: '—', qtyAfter: '1', priceBefore: '—', priceAfter: '850,00 €', variation: '+850,00 €', highlight: true },
];

const badgeStyles = {
  orange: { background: 'rgba(107,107,107,0.14)', color: C.darkGray },
  teal: { background: C.accentSoft, color: C.maroon },
  gray: { background: '#EEECE6', color: C.darkGray },
};

const card = { background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 20, padding: 18, boxShadow: '0 1px 2px rgba(0,0,0,0.03), 0 4px 12px rgba(0,0,0,0.04)' };
const h1Style = { fontFamily: FONT, fontSize: 30, margin: 0, color: C.black, fontWeight: 700 };
const breadcrumb = { fontSize: 12, color: C.gray, margin: '0 0 4px' };
const freshBadge = { fontSize: 11, color: C.maroon, background: C.accentSoft, padding: '3px 9px', borderRadius: 999, fontWeight: 600 };
const iconBtn = { width: 26, height: 26, border: `1px solid ${C.paleGray}`, borderRadius: 6, background: C.white, cursor: 'pointer', fontSize: 11, color: C.gray };

function MetricCard({ label, value, note }) {
  return (
    <div style={card}>
      <p style={{ fontSize: 12, color: C.gray, margin: '0 0 8px' }}>{label}</p>
      <p style={{ fontFamily: FONT, fontSize: 24, fontWeight: 700, margin: '0 0 6px', color: C.black }}>{value}</p>
      {note && <p style={{ fontSize: 11, color: C.gray, margin: 0 }}>{note}</p>}
    </div>
  );
}

function Dashboard({ onNavigate, onOpenProject, projects }) {
  const topProject = projects[0];
  const progressByStatus = { 'In attesa di approvazione': 33, 'Approvato': 66, 'In fase di cantiere': 100 };
  const totalValoreListino = '23.422,85 €';
  const initials = (name) => name.split(' ').map((w) => w[0]).slice(0, 2).join('').toUpperCase();

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Dashboard</p>
      <p style={{ fontSize: 18, fontWeight: 600, color: C.black, margin: '2px 0 28px' }}>Dashboard</p>

      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', marginBottom: 32, flexWrap: 'wrap', gap: 16 }}>
        <div>
          <span style={{ display: 'block', fontSize: 12, fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase', color: C.maroon, marginBottom: 8 }}>Oggi</span>
          <h1 style={{ fontFamily: FONT, fontSize: 44, fontWeight: 700, lineHeight: 1.1, letterSpacing: '-0.02em', color: C.black, margin: 0 }}>Bentornato.</h1>
          <p style={{ fontSize: 15, color: C.darkGray, margin: '8px 0 0' }}>Progetti, computi e listini nello stesso contesto.</p>
        </div>
        <button
          onClick={() => onNavigate('progetti')}
          style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: C.sidebar, color: C.white, border: 'none', borderRadius: 999, padding: '12px 24px', fontFamily: FONT, fontWeight: 500, fontSize: 13, cursor: 'pointer' }}
        >
          <Building2 size={16} strokeWidth={1.5} />
          Apri i progetti
        </button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 24, marginBottom: 32 }}>
        <div style={card}>
          <p style={{ fontSize: 13, color: C.darkGray, margin: '0 0 16px' }}>Progetti attivi</p>
          <p style={{ fontFamily: FONT, fontSize: 44, fontWeight: 700, color: C.black, margin: 0, lineHeight: 1.1 }}>{projects.length}</p>
        </div>
        <div style={card}>
          <p style={{ fontSize: 13, color: C.darkGray, margin: '0 0 16px' }}>Voci di listino</p>
          <p style={{ fontFamily: FONT, fontSize: 44, fontWeight: 700, color: C.black, margin: 0, lineHeight: 1.1 }}>13</p>
        </div>
        <div style={card}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <p style={{ fontSize: 13, color: C.darkGray, margin: 0 }}>Macrosezioni</p>
            <Clock size={20} strokeWidth={1.5} color={C.maroon} />
          </div>
          <p style={{ fontFamily: FONT, fontSize: 44, fontWeight: 700, color: C.black, margin: 0, lineHeight: 1.1 }}>4</p>
        </div>
        <div style={card}>
          <p style={{ fontSize: 13, color: C.darkGray, margin: '0 0 16px' }}>Valore progetti</p>
          <p style={{ fontFamily: FONT, fontSize: 30, fontWeight: 700, color: C.black, margin: 0, lineHeight: 1.1 }}>{totalValoreListino}</p>
        </div>
      </div>

      <p style={{ fontSize: 18, fontWeight: 600, color: C.black, margin: '0 0 16px' }}>Progetti attivi</p>

      {topProject && (
        <div onClick={() => onOpenProject(topProject.id)} style={{ ...card, marginBottom: 20, cursor: 'pointer' }} className="hover-lift">
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 16 }}>
            <span style={{ fontSize: 13, fontWeight: 500, padding: '4px 12px', borderRadius: 999, ...badgeStyles[statusTone[latestStatus(topProject)]] }}>{latestStatus(topProject)}</span>
            <div style={{ width: 40, height: 40, borderRadius: 999, background: '#0D4D3C', color: C.white, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 13 }}>
              {initials(topProject.client)}
            </div>
          </div>
          <p style={{ fontFamily: FONT, fontSize: 22, fontWeight: 600, color: C.black, margin: '0 0 6px' }}>{topProject.name}</p>
          <div style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 13, color: C.darkGray, marginBottom: 16 }}>
            <MapPin size={14} strokeWidth={1.5} />
            {topProject.header?.ubicazione || topProject.client}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: 13, color: C.darkGray, marginBottom: 8 }}>
            <span>{topProject.revisions.length} revisioni</span>
            <span style={{ fontWeight: 600, color: C.black }}>{progressByStatus[latestStatus(topProject)]}%</span>
          </div>
          <div style={{ background: '#E9E6DE', borderRadius: 999, height: 6, overflow: 'hidden' }}>
            <div style={{ background: '#0D4D3C', height: '100%', width: `${progressByStatus[latestStatus(topProject)]}%`, borderRadius: 999 }} />
          </div>
        </div>
      )}

      <div style={{ ...card, marginBottom: 20 }}>
        <p style={{ fontSize: 18, fontWeight: 600, color: C.black, margin: '0 0 16px' }}>Altri progetti</p>
        {projects.slice(1).length === 0 && <p style={{ fontSize: 13, color: C.gray, margin: 0 }}>Nessun altro progetto.</p>}
        {projects.slice(1).map((p) => (
          <div
            key={p.id}
            onClick={() => onOpenProject(p.id)}
            style={{ display: 'flex', alignItems: 'center', gap: 16, padding: 16, borderRadius: 12, background: C.surfaceSubtle, cursor: 'pointer', marginBottom: 8 }}
          >
            <div style={{ width: 44, height: 44, borderRadius: 8, background: C.accentSoft, color: C.maroon, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 14, flexShrink: 0 }}>
              {initials(p.client)}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <p style={{ fontWeight: 500, fontSize: 13, margin: 0, color: C.black }}>{p.name}</p>
              <p style={{ fontSize: 12, color: C.darkGray, margin: '2px 0 0' }}>{p.client}</p>
            </div>
            <span style={{ fontSize: 13, fontWeight: 500, padding: '4px 12px', borderRadius: 999, ...badgeStyles[statusTone[latestStatus(p)]] }}>{latestStatus(p)}</span>
          </div>
        ))}
      </div>

      <div style={card}>
        <p style={{ fontSize: 18, fontWeight: 600, color: C.black, margin: '0 0 16px' }}>Azioni rapide</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 10 }}>
          <button onClick={() => onNavigate('computi')} style={{ background: C.surfaceSubtle, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '12px 14px', fontSize: 13, fontWeight: 500, textAlign: 'left', color: C.black, fontFamily: FONT, cursor: 'pointer' }}>
            Apri computo e listino
          </button>
          <button onClick={() => onNavigate('confronto')} style={{ background: C.surfaceSubtle, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '12px 14px', fontSize: 13, fontWeight: 500, textAlign: 'left', color: C.black, fontFamily: FONT, cursor: 'pointer' }}>
            Confronta revisioni
          </button>
          <button onClick={() => onNavigate('computi')} style={{ background: C.surfaceSubtle, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '12px 14px', fontSize: 13, fontWeight: 500, textAlign: 'left', color: C.black, fontFamily: FONT, cursor: 'pointer' }}>
            Personalizza documento
          </button>
        </div>
      </div>
    </div>
  );
}

function ListinoPage({ listini, setListini, activeId, setActiveId }) {
  const active = listini.find((l) => l.id === activeId);

  const setMacrosForActive = (macros) => {
    setListini(listini.map((l) => (l.id === activeId ? { ...l, macros } : l)));
  };

  const addListino = () => {
    const name = prompt('Nome del nuovo listino (es. "Listino ristrutturazioni 2026", "Listino nuove costruzioni"):');
    if (!name) return;
    const id = Math.max(...listini.map((l) => l.id)) + 1;
    setListini([...listini, { id, name, macros: [] }]);
    setActiveId(id);
  };

  const duplicateListino = () => {
    const name = prompt('Nome della copia:', `${active.name} (copia)`);
    if (!name) return;
    const id = Math.max(...listini.map((l) => l.id)) + 1;
    setListini([...listini, { id, name, macros: structuredClone(active.macros) }]);
    setActiveId(id);
  };

  const renameListino = () => {
    const name = prompt('Rinomina listino:', active.name);
    if (!name) return;
    setListini(listini.map((l) => (l.id === activeId ? { ...l, name } : l)));
  };

  const deleteListino = () => {
    if (listini.length === 1) { alert('Deve rimanere almeno un listino.'); return; }
    if (!confirm(`Eliminare "${active.name}"? Questa azione non si può annullare.`)) return;
    const remaining = listini.filter((l) => l.id !== activeId);
    setListini(remaining);
    setActiveId(remaining[0].id);
  };

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Listino prezzi</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 20 }}>
        <h1 style={h1Style}>Listino prezzi</h1>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>

      <div style={{ ...card, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Listino attivo</label>
        <select
          value={activeId}
          onChange={(e) => setActiveId(Number(e.target.value))}
          style={{ fontSize: 13, fontWeight: 600, padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, minWidth: 220 }}
        >
          {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
        </select>
        <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
          <button onClick={addListino} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Nuovo listino</button>
          <button onClick={duplicateListino} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⧉ Duplica</button>
          <button onClick={renameListino} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>✎ Rinomina</button>
          <button onClick={deleteListino} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>🗑 Elimina</button>
        </div>
      </div>

      <EditableCatalog macros={active.macros} setMacros={setMacrosForActive} />
    </div>
  );
}

const INITIAL_MACROS = [
  { name: 'Edilizia e strutture', categorie: [
    { name: 'Murature', sottocategorie: [
      { name: 'Murature portanti', voci: [{ code: 'ED.MUR.01.010', desc: 'Muratura perimetrale piano terra', unit: 'm²', priceImpresa: '68,50', priceCliente: '=impresa*1.3' }] },
      { name: 'Murature divisorie', voci: [{ code: 'ED.MUR.02.010', desc: 'Tramezzo in laterizio', unit: 'm²', priceImpresa: '32,00', priceCliente: '=impresa*1.3' }] },
    ]},
    { name: 'Strutture', sottocategorie: [
      { name: 'Strutture in c.a.', voci: [{ code: 'ED.STR.01.010', desc: 'Cordolo in c.a. armato', unit: 'm', priceImpresa: '62,00', priceCliente: '=impresa*1.3' }] },
    ]},
  ]},
  { name: 'Impianti tecnologici', categorie: [
    { name: 'Impianti elettrici', sottocategorie: [
      { name: 'Impianti civili', voci: [
        { code: 'IT.EL.01.005', desc: 'Punti luce appartamenti', unit: 'cad', priceImpresa: '72,00', priceCliente: '=impresa*1.25' },
        { code: 'IT.EL.01.020', desc: 'Contatore energia', unit: 'cad', priceImpresa: '210,00', priceCliente: '=impresa*1.25' },
      ]},
      { name: 'Impianti speciali', voci: [{ code: 'IT.EL.02.010', desc: 'Predisposizione domotica', unit: 'cad', priceImpresa: '180,00', priceCliente: '=impresa*1.25' }] },
    ]},
    { name: 'Impianti idrico-sanitari', sottocategorie: [
      { name: 'Impianti idrici', voci: [{ code: 'IT.ID.01.010', desc: 'Impianto idrico-sanitario completo', unit: 'cad', priceImpresa: '320,00', priceCliente: '=impresa*1.25' }] },
    ]},
  ]},
  { name: 'Finiture e opere esterne', categorie: [
    { name: 'Finiture', sottocategorie: [
      { name: 'Pavimenti', voci: [
        { code: 'FO.FIN.01.010', desc: 'Pavimento zona giorno', unit: 'm²', priceImpresa: '54,90', priceCliente: '=impresa+20' },
        { code: 'FO.FIN.01.020', desc: 'Massetto alleggerito', unit: 'm²', priceImpresa: '18,00', priceCliente: '=impresa+20' },
      ]},
      { name: 'Rivestimenti', voci: [
        { code: 'FO.FIN.02.010', desc: 'Rivestimento bagno ceramica', unit: 'm²', priceImpresa: '45,00', priceCliente: '=impresa+20' },
        { code: 'FO.FIN.02.020', desc: 'Rivestimento cucina', unit: 'm²', priceImpresa: '42,00', priceCliente: '=impresa+20' },
      ]},
    ]},
    { name: 'Serramenti', sottocategorie: [
      { name: 'Infissi esterni', voci: [
        { code: 'FO.SER.01.010', desc: 'Infisso PVC doppio vetro', unit: 'm²', priceImpresa: '320,00', priceCliente: '380,00' },
        { code: 'FO.SER.01.020', desc: 'Persiana in alluminio', unit: 'cad', priceImpresa: '140,00', priceCliente: '170,00' },
      ]},
    ]},
  ]},
];

function withCodes(macros) {
  return macros.map((m, i) => {
    const mCode = `M${String(i + 1).padStart(2, '0')}`;
    return {
      ...m, code: mCode,
      categorie: m.categorie.map((c, j) => {
        const cCode = `${mCode}-C${String(j + 1).padStart(2, '0')}`;
        return {
          ...c, code: cCode,
          sottocategorie: c.sottocategorie.map((s, k) => ({
            ...s, code: `${cCode}-S${String(k + 1).padStart(2, '0')}`,
          })),
        };
      }),
    };
  });
}

const UNIT_OPTIONS = ['m²', 'm³', 'ml', 'cadauna', 'a corpo', 'kg', 'ora', 'Altro'];

function VoceModal({ locations, initialLocationIdx = 0, initialVoce = null, onClose, onSave }) {
  const isEdit = !!initialVoce;
  const [locationIdx, setLocationIdx] = useState(initialLocationIdx);
  const [code, setCode] = useState(initialVoce?.code || '');
  const [desc, setDesc] = useState(initialVoce?.desc || '');
  const startsKnown = initialVoce && UNIT_OPTIONS.slice(0, -1).includes(initialVoce.unit);
  const [unit, setUnit] = useState(initialVoce ? (startsKnown ? initialVoce.unit : 'Altro') : 'm²');
  const [customUnit, setCustomUnit] = useState(initialVoce && !startsKnown ? initialVoce.unit : '');
  const [priceImpresa, setPriceImpresa] = useState(initialVoce?.priceImpresa || '');
  const [priceCliente, setPriceCliente] = useState(initialVoce?.priceCliente || '');
  const previewImpresa = parseEuro(priceImpresa);
  const previewCliente = evalClientPrice(priceCliente, previewImpresa);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 400 }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 16px', color: C.black }}>{isEdit ? 'Modifica voce di listino' : 'Nuova voce di listino'}</h2>

        {isEdit ? (
          <p style={{ fontSize: 11, color: C.gray, margin: '0 0 12px' }}>In: {locations[locationIdx]?.label}</p>
        ) : (
          <>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Sottocategoria</label>
            <select value={locationIdx} onChange={(e) => setLocationIdx(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
              {locations.map((l, i) => <option key={l.path} value={i}>{l.label}</option>)}
            </select>
          </>
        )}

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Codice voce</label>
        <input value={code} onChange={(e) => setCode(e.target.value)} placeholder="Es. EX.001" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }} />

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Descrizione tecnica</label>
        <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Descrivi lavorazione, materiali e condizioni…" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }} />

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Unità di misura</label>
        <select value={unit} onChange={(e) => setUnit(e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {UNIT_OPTIONS.map((u) => <option key={u} value={u}>{u}</option>)}
        </select>
        {unit === 'Altro' && (
          <input value={customUnit} onChange={(e) => setCustomUnit(e.target.value)} placeholder="Es. q.li, kWh…" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginBottom: 12 }} />
        )}

        <div style={{ display: 'flex', gap: 10 }}>
          <div style={{ flex: 1 }}>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Prezzo impresa (€)</label>
            <input value={priceImpresa} onChange={(e) => setPriceImpresa(e.target.value)} placeholder="0,00" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 6px' }} />
          </div>
          <div style={{ flex: 1 }}>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Prezzo cliente (€ o formula)</label>
            <input value={priceCliente} onChange={(e) => setPriceCliente(e.target.value)} placeholder="Es. =impresa*1.3" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 6px' }} />
          </div>
        </div>
        <p style={{ fontSize: 11, color: C.gray, margin: '0 0 12px' }}>
          Nel campo "prezzo cliente" puoi scrivere un numero fisso, oppure una formula che inizia con "=" e usa la parola <strong>impresa</strong>: es. <code>=impresa*1.3</code> (margine 30%), <code>=impresa+50</code> (fisso +50 €).
          {priceImpresa && <> Anteprima: impresa {formatEuro(previewImpresa)} → cliente {formatEuro(previewCliente)}.</>}
        </p>

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 6 }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              if (!desc.trim()) return;
              const finalUnit = unit === 'Altro' ? (customUnit || '—') : unit;
              onSave(locations[locationIdx].path, { code: code || '—', desc, unit: finalUnit, priceImpresa: priceImpresa || '0,00', priceCliente: priceCliente || '' });
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            {isEdit ? 'Salva modifiche' : 'Salva voce'}
          </button>
        </div>
      </div>
    </div>
  );
}

function EditableCatalog({ macros, setMacros }) {
  const [showVoceModal, setShowVoceModal] = useState(false);
  const [editingVoce, setEditingVoce] = useState(null); // { path: [mi,ci,si,vi] }
  const [expanded, setExpanded] = useState({});
  const coded = withCodes(macros);

  const isOpen = (key) => expanded[key] !== false; // default aperto
  const toggle = (key) => setExpanded({ ...expanded, [key]: !isOpen(key) });

  const totals = coded.reduce(
    (acc, m) => {
      acc.categorie += m.categorie.length;
      m.categorie.forEach((c) => {
        acc.sottocategorie += c.sottocategorie.length;
        c.sottocategorie.forEach((s) => { acc.voci += s.voci.length; });
      });
      return acc;
    },
    { voci: 0, categorie: 0, sottocategorie: 0 }
  );

  const locations = [];
  coded.forEach((m, mi) => m.categorie.forEach((c, ci) => c.sottocategorie.forEach((s, si) => {
    locations.push({ path: [mi, ci, si], label: `${m.name} › ${c.name} › ${s.name} (${s.code})` });
  })));

  const addMacro = () => {
    const name = prompt('Nome della nuova macrocategoria:');
    if (!name) return;
    setMacros([...macros, { name, categorie: [] }]);
  };
  const addCategoria = (mi) => {
    const name = prompt('Nome della nuova categoria:');
    if (!name) return;
    const next = structuredClone(macros);
    next[mi].categorie.push({ name, sottocategorie: [] });
    setMacros(next);
  };
  const addSotto = (mi, ci) => {
    const name = prompt('Nome della nuova sottocategoria:');
    if (!name) return;
    const next = structuredClone(macros);
    next[mi].categorie[ci].sottocategorie.push({ name, voci: [] });
    setMacros(next);
  };
  const addVoce = ([mi, ci, si], voce) => {
    const next = structuredClone(macros);
    next[mi].categorie[ci].sottocategorie[si].voci.push(voce);
    setMacros(next);
  };
  const saveVoce = (path, voce) => {
    const next = structuredClone(macros);
    if (editingVoce) {
      const [mi, ci, si, vi] = editingVoce.path;
      next[mi].categorie[ci].sottocategorie[si].voci[vi] = voce;
    } else {
      const [mi, ci, si] = path;
      next[mi].categorie[ci].sottocategorie[si].voci.push(voce);
    }
    setMacros(next);
  };
  const rename = (kind, path) => {
    const next = structuredClone(macros);
    if (kind === 'macro') {
      const name = prompt('Rinomina macrocategoria:', next[path[0]].name);
      if (name) next[path[0]].name = name;
    } else if (kind === 'categoria') {
      const cat = next[path[0]].categorie[path[1]];
      const name = prompt('Rinomina categoria:', cat.name);
      if (name) cat.name = name;
    } else {
      const sotto = next[path[0]].categorie[path[1]].sottocategorie[path[2]];
      const name = prompt('Rinomina sottocategoria:', sotto.name);
      if (name) sotto.name = name;
    }
    setMacros(next);
  };
  const remove = (kind, path) => {
    if (!confirm('Eliminare questo elemento e tutto ciò che contiene?')) return;
    const next = structuredClone(macros);
    if (kind === 'macro') next.splice(path[0], 1);
    else if (kind === 'categoria') next[path[0]].categorie.splice(path[1], 1);
    else if (kind === 'sotto') next[path[0]].categorie[path[1]].sottocategorie.splice(path[2], 1);
    else if (kind === 'voce') next[path[0]].categorie[path[1]].sottocategorie[path[2]].voci.splice(path[3], 1);
    setMacros(next);
  };

  const rowBtn = { border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 6, fontSize: 11, fontWeight: 600, padding: '3px 8px', cursor: 'pointer', color: C.midGray };
  const codeTag = { fontSize: 10, fontWeight: 700, color: C.maroon, background: 'rgba(128,20,48,0.08)', padding: '2px 6px', borderRadius: 5, marginLeft: 8 };

  return (
    <div>
      <div style={{ ...card, marginBottom: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
          <div>
            <p style={{ fontSize: 15, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Catalogo / Listino</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '4px 0 0' }}>{totals.voci} voci consultabili · clicca su una riga per aprirla o chiuderla</p>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => { setEditingVoce(null); setShowVoceModal(true); }} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Nuova voce</button>
            <button onClick={addMacro} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Macrocategoria</button>
          </div>
        </div>

        {coded.length === 0 && (
          <p style={{ fontSize: 12, color: C.gray, margin: '0 0 10px' }}>Questo listino è vuoto. Aggiungi una macrocategoria per iniziare.</p>
        )}

        {coded.map((m, mi) => {
          const mKey = `${mi}`;
          return (
            <div key={mi} style={{ marginBottom: 14, border: `1px solid ${C.paleGray}`, borderRadius: 10, overflow: 'hidden' }}>
              <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', background: C.bg, cursor: 'pointer' }}>
                <span style={{ fontSize: 11, color: C.gray }}>{isOpen(mKey) ? '⌄' : '›'}</span>
                <span style={{ width: 7, height: 7, borderRadius: '50%', background: C.maroon }} />
                <span style={{ fontWeight: 700, fontSize: 13, color: C.black }}>{m.name}</span>
                <span style={codeTag}>{m.code}</span>
                <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
                  <button onClick={() => rename('macro', [mi])} style={rowBtn}>✎ Rinomina</button>
                  <button onClick={() => addCategoria(mi)} style={rowBtn}>+ Categoria</button>
                  <button onClick={() => remove('macro', [mi])} style={rowBtn}>🗑</button>
                </div>
              </div>

              {isOpen(mKey) && m.categorie.map((c, ci) => {
                const cKey = `${mi}-${ci}`;
                return (
                  <div key={ci} style={{ paddingLeft: 20, borderTop: `1px solid ${C.paleGray}` }}>
                    <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '9px 12px', cursor: 'pointer' }}>
                      <span style={{ fontSize: 11, color: C.gray }}>{isOpen(cKey) ? '⌄' : '›'}</span>
                      <span style={{ fontWeight: 600, fontSize: 12, color: C.black }}>{c.name}</span>
                      <span style={codeTag}>{c.code}</span>
                      <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
                        <button onClick={() => rename('categoria', [mi, ci])} style={rowBtn}>✎</button>
                        <button onClick={() => addSotto(mi, ci)} style={rowBtn}>+ Sottocategoria</button>
                        <button onClick={() => remove('categoria', [mi, ci])} style={rowBtn}>🗑</button>
                      </div>
                    </div>

                    {isOpen(cKey) && c.sottocategorie.map((s, si) => {
                      const sKey = `${mi}-${ci}-${si}`;
                      return (
                        <div key={si} style={{ paddingLeft: 20, borderTop: `1px solid ${C.paleGray}` }}>
                          <div onClick={() => toggle(sKey)} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 12px', cursor: 'pointer' }}>
                            <span style={{ fontSize: 11, color: C.gray }}>{isOpen(sKey) ? '⌄' : '›'}</span>
                            <span style={{ fontSize: 12, color: C.midGray }}>{s.name}</span>
                            <span style={codeTag}>{s.code}</span>
                            <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
                              <button onClick={() => rename('sotto', [mi, ci, si])} style={rowBtn}>✎</button>
                              <button onClick={() => remove('sotto', [mi, ci, si])} style={rowBtn}>🗑</button>
                            </div>
                          </div>
                          {isOpen(sKey) && s.voci.length > 0 && (
                            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, marginBottom: 6 }}>
                              <thead>
                                <tr style={{ color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                                  <th style={{ padding: '4px 12px 4px 32px', textAlign: 'left' }}></th>
                                  <th style={{ padding: '4px 12px', textAlign: 'left' }}></th>
                                  <th style={{ padding: '4px 12px', textAlign: 'left' }}></th>
                                  <th style={{ padding: '4px 12px', textAlign: 'right' }}>Prezzo impresa</th>
                                  <th style={{ padding: '4px 12px', textAlign: 'right' }}>Prezzo cliente</th>
                                  <th style={{ padding: '4px 12px' }}></th>
                                </tr>
                              </thead>
                              <tbody>
                                {s.voci.map((v, vi) => {
                                  const impresaVal = parseEuro(v.priceImpresa);
                                  const clienteVal = evalClientPrice(v.priceCliente, impresaVal);
                                  return (
                                  <tr key={vi} style={{ borderTop: `1px solid ${C.paleGray}` }}>
                                    <td style={{ padding: '6px 12px 6px 32px', fontWeight: 700, color: C.black, width: 110 }}>{v.code}</td>
                                    <td style={{ padding: '6px 12px', color: C.midGray }}>{v.desc}</td>
                                    <td style={{ padding: '6px 12px', color: C.gray, width: 70 }}>{v.unit}</td>
                                    <td style={{ padding: '6px 12px', textAlign: 'right', fontWeight: 600, color: C.black, width: 90 }}>{formatEuro(impresaVal)}</td>
                                    <td style={{ padding: '6px 12px', textAlign: 'right', fontWeight: 600, color: C.maroon, width: 90 }}>{formatEuro(clienteVal)}</td>
                                    <td style={{ padding: '6px 12px', width: 60 }}>
                                      <div style={{ display: 'flex', gap: 4 }}>
                                        <button
                                          onClick={() => { setEditingVoce({ path: [mi, ci, si, vi] }); setShowVoceModal(true); }}
                                          style={{ ...rowBtn, padding: '2px 6px' }}
                                        >
                                          ✎
                                        </button>
                                        <button onClick={() => remove('voce', [mi, ci, si, vi])} style={{ ...rowBtn, padding: '2px 6px' }}>🗑</button>
                                      </div>
                                    </td>
                                  </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          )}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      <div style={card}>
        <p style={{ fontSize: 15, fontWeight: 700, margin: '0 0 14px', color: C.black, fontFamily: FONT }}>Listino prezzi aziendale</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 14, marginBottom: 14 }}>
          <MetricCard label="Articoli" value={String(totals.voci)} />
          <MetricCard label="Categorie" value={String(totals.categorie)} />
          <MetricCard label="Sottocategorie" value={String(totals.sottocategorie)} />
        </div>
        <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>
          I codici di macrocategoria, categoria e sottocategoria si aggiornano automaticamente in base all'ordine e alla struttura.
        </p>
      </div>

      {showVoceModal && (
        <VoceModal
          locations={locations}
          initialLocationIdx={editingVoce ? locations.findIndex((l) => l.path.join() === editingVoce.path.slice(0, 3).join()) : 0}
          initialVoce={editingVoce ? macros[editingVoce.path[0]].categorie[editingVoce.path[1]].sottocategorie[editingVoce.path[2]].voci[editingVoce.path[3]] : null}
          onClose={() => { setShowVoceModal(false); setEditingVoce(null); }}
          onSave={saveVoce}
        />
      )}
    </div>
  );
}

function NewProjectModal({ onClose, onCreate }) {
  const [name, setName] = useState('');
  const [client, setClient] = useState('');

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 380 }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 16px', color: C.black }}>Crea nuovo progetto</h2>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Nome progetto</label>
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Es. Ristrutturazione Palazzo Verdi"
          style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}
        />

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Cliente o committente</label>
        <input
          value={client}
          onChange={(e) => setClient(e.target.value)}
          placeholder="Es. Immobiliare Centro S.r.l."
          style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 18px' }}
        />

        <p style={{ fontSize: 11, color: C.gray, margin: '0 0 16px' }}>
          Il primo computo metrico si crea dopo, entrando nel progetto.
        </p>

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              onCreate({
                id: Date.now(),
                name: name || 'Nuovo progetto',
                client: client || 'Cliente da definire',
                items: 0,
                value: '—',
                team: [],
                revisions: [],
              });
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            Crea progetto
          </button>
        </div>
      </div>
    </div>
  );
}

function ProjectCard({ p, onOpen, onDelete }) {
  return (
    <div onClick={() => onOpen(p.id)} style={{ ...card, cursor: 'pointer', position: 'relative' }}>
      <button
        onClick={(e) => { e.stopPropagation(); onDelete(p.id); }}
        style={{ position: 'absolute', top: 14, right: 14, border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 6, width: 26, height: 26, fontSize: 11, cursor: 'pointer', color: C.gray }}
      >
        🗑
      </button>
      <div style={{ marginBottom: 14, paddingRight: 30 }}>
        <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, ...badgeStyles[statusTone[latestStatus(p)]] }}>{latestStatus(p)}</span>
      </div>
      <p style={{ fontFamily: FONT, fontSize: 20, fontWeight: 700, margin: '0 0 4px', color: C.black }}>{p.name}</p>
      <p style={{ fontSize: 12, color: C.gray, margin: '0 0 14px' }}>{p.client}</p>
      <p style={{ fontFamily: FONT, fontSize: 22, fontWeight: 700, margin: '0 0 6px', color: C.black }}>{p.value}</p>
      <p style={{ fontSize: 11, color: C.gray, margin: 0 }}>{p.items} voci · {p.revisions.length} revisioni salvate</p>
    </div>
  );
}

function ProgettiPage({ projects, setProjects, onOpenProject }) {
  const [showModal, setShowModal] = useState(false);

  const deleteProject = (id) => {
    const p = projects.find((x) => x.id === id);
    if (!confirm(`Eliminare il progetto "${p.name}" e tutti i suoi computi salvati?`)) return;
    setProjects(projects.filter((x) => x.id !== id));
  };

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Progetti</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 4 }}>
        <h1 style={h1Style}>Progetti</h1>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: 18 }}>
        <p style={{ fontSize: 13, color: C.gray, margin: 0 }}>Seleziona un progetto per aprirlo, oppure crea un nuovo progetto.</p>
        <button
          onClick={() => setShowModal(true)}
          style={{ marginLeft: 'auto', background: C.maroon, color: C.white, border: 'none', padding: '10px 16px', borderRadius: 8, fontSize: 13, fontWeight: 600 }}
        >
          + Nuovo progetto
        </button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16 }}>
        {projects.map((p) => <ProjectCard key={p.id} p={p} onOpen={onOpenProject} onDelete={deleteProject} />)}
      </div>

      {showModal && (
        <NewProjectModal
          onClose={() => setShowModal(false)}
          onCreate={(np) => setProjects([...projects, np])}
        />
      )}
    </div>
  );
}

function ComputoSectionsView({ sections }) {
  return (
    <>
      {sections.map((section) => (
        <div key={section.name} style={{ border: `1px solid ${C.paleGray}`, borderRadius: 12, overflow: 'hidden', marginBottom: 16 }}>
          <div style={{ background: section.color, color: C.white, padding: '10px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontWeight: 700, fontSize: 13, fontFamily: FONT }}>{section.name}</span>
            {section.final && <span style={{ fontSize: 10, fontWeight: 700, background: 'rgba(255,255,255,0.22)', padding: '3px 9px', borderRadius: 999 }}>Sezione finale</span>}
          </div>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, background: C.white }}>
            <thead>
              <tr style={{ textAlign: 'left', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                <th style={{ padding: '10px 16px' }}>Codice</th>
                <th style={{ padding: '10px 16px' }}>Descrizione</th>
                <th style={{ padding: '10px 16px', textAlign: 'right' }}>Q.tà</th>
                <th style={{ padding: '10px 16px' }}>U.M.</th>
                <th style={{ padding: '10px 16px', textAlign: 'right' }}>Prezzo</th>
                <th style={{ padding: '10px 16px', textAlign: 'right' }}>Totale</th>
                <th style={{ padding: '10px 16px' }}>Azioni</th>
              </tr>
            </thead>
            <tbody>
              {section.items.map((it) => (
                <tr key={it.code} style={{ borderTop: `1px solid ${C.paleGray}` }}>
                  <td style={{ padding: '10px 16px', fontWeight: 700, color: C.black }}>{it.code}</td>
                  <td style={{ padding: '10px 16px', color: C.midGray }}>{it.desc}</td>
                  <td style={{ padding: '10px 16px', textAlign: 'right' }}>{it.qty}</td>
                  <td style={{ padding: '10px 16px', color: C.gray }}>{it.unit}</td>
                  <td style={{ padding: '10px 16px', textAlign: 'right' }}>{it.price}</td>
                  <td style={{ padding: '10px 16px', textAlign: 'right', fontWeight: 700, color: C.black }}>{it.total}</td>
                  <td style={{ padding: '10px 16px' }}>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => alert(`Modifica voce ${it.code}`)} style={iconBtn}>✎</button>
                      <button onClick={() => alert(`Duplica voce ${it.code}`)} style={iconBtn}>⧉</button>
                      <button onClick={() => confirm(`Eliminare la voce ${it.code}?`)} style={iconBtn}>🗑</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '10px 16px', fontSize: 12, fontWeight: 700, color: C.black, borderTop: `1px solid ${C.paleGray}` }}>
            Subtotale {section.name}&nbsp;&nbsp;{section.subtotal}
          </div>
        </div>
      ))}
    </>
  );
}

const SECTION_COLORS = [C.maroon, C.darkGray, '#94706C', C.sidebar];

function DraggableCatalogTree({ listino }) {
  const [expanded, setExpanded] = useState({});
  const isOpen = (key) => expanded[key] !== false;
  const toggle = (key) => setExpanded({ ...expanded, [key]: !isOpen(key) });

  if (!listino) return null;
  const coded = withCodes(listino.macros);

  if (coded.length === 0) {
    return <p style={{ fontSize: 12, color: C.gray }}>Questo listino non ha ancora voci.</p>;
  }

  return (
    <div>
      {coded.map((m, mi) => {
        const mKey = `m${mi}`;
        return (
          <div key={mi} style={{ marginBottom: 8 }}>
            <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '6px 4px', borderRadius: 6, background: C.bg }}>
              <span style={{ fontSize: 10, color: C.gray }}>{isOpen(mKey) ? '⌄' : '›'}</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.black }}>{m.name}</span>
            </div>
            {isOpen(mKey) && m.categorie.map((c, ci) => {
              const cKey = `m${mi}c${ci}`;
              return (
                <div key={ci} style={{ paddingLeft: 12 }}>
                  <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '5px 4px' }}>
                    <span style={{ fontSize: 10, color: C.gray }}>{isOpen(cKey) ? '⌄' : '›'}</span>
                    <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{c.name}</span>
                  </div>
                  {isOpen(cKey) && c.sottocategorie.map((s, si) => {
                    const sKey = `m${mi}c${ci}s${si}`;
                    return (
                      <div key={si} style={{ paddingLeft: 12 }}>
                        <div onClick={() => toggle(sKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '4px 4px' }}>
                          <span style={{ fontSize: 10, color: C.gray }}>{isOpen(sKey) ? '⌄' : '›'}</span>
                          <span style={{ fontSize: 12, color: C.midGray }}>{s.name}</span>
                        </div>
                        {isOpen(sKey) && (
                          <div style={{ paddingLeft: 12 }}>
                            {s.voci.map((v, vi) => {
                              const impresaVal = parseEuro(v.priceImpresa);
                              const clienteVal = evalClientPrice(v.priceCliente, impresaVal);
                              return (
                              <div
                                key={vi}
                                draggable
                                onDragStart={(e) => e.dataTransfer.setData('application/json', JSON.stringify({ ...v, macro: m.name, impresaValue: impresaVal, clienteValue: clienteVal }))}
                                style={{ border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '7px 9px', marginBottom: 6, cursor: 'grab', background: C.white }}
                              >
                                <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: C.maroon }}>{v.code}</p>
                                <p style={{ margin: '2px 0 0', fontSize: 12, color: C.black }}>{v.desc}</p>
                                <p style={{ margin: '2px 0 0', fontSize: 11, color: C.gray }}>{v.unit} · impresa {formatEuro(impresaVal)} · cliente {formatEuro(clienteVal)}</p>
                              </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

function groupItemsForExport(items, extraSections) {
  const groups = [];
  const byName = (name) => {
    let g = groups.find((s) => s.name === name);
    if (!g) { g = { name, items: [] }; groups.push(g); }
    return g;
  };
  (extraSections || []).forEach((n) => byName(n));
  (items || []).forEach((it) => byName(it.section || it.macro || 'Voci varie').items.push(it));
  return groups;
}

function exportComputoExcel(project, revision, clientOnly) {
  const groups = groupItemsForExport(revision.items, revision.extraSections);
  const rows = [];
  rows.push([clientOnly ? 'Computo metrico (versione cliente)' : 'Computo metrico']);
  rows.push(['Progetto', project.name]);
  rows.push(['Cliente', project.client]);
  rows.push(['Versione', revision.customName || revision.label]);
  rows.push(['Data modifica', revision.dateModified]);
  rows.push([]);
  rows.push(clientOnly
    ? ['Sezione', 'Codice', 'Descrizione', 'Quantità', 'U.M.', 'Prezzo cliente', 'Totale cliente']
    : ['Sezione', 'Codice', 'Descrizione', 'Quantità', 'U.M.', 'Prezzo impresa', 'Totale impresa', 'Prezzo cliente', 'Totale cliente']);

  groups.forEach((g) => {
    let runImpresa = 0;
    let runCliente = 0;
    g.items.forEach((it) => {
      if (it.type === 'subtotal') {
        const hasVat = it.vatRate !== null && it.vatRate !== undefined;
        if (hasVat) {
          if (!clientOnly) {
            rows.push([g.name, '', `${it.title} — IVA esclusa`, '', '', '', formatEuro(runImpresa)]);
            rows.push([g.name, '', `${it.title} — ${it.vatLabel}`, '', '', '', formatEuro(runImpresa * (it.vatRate / 100))]);
            rows.push([g.name, '', `${it.title} — IVA inclusa`, '', '', '', formatEuro(runImpresa * (1 + it.vatRate / 100))]);
          } else {
            rows.push([g.name, '', `${it.title} — IVA esclusa`, '', '', formatEuro(runCliente)]);
            rows.push([g.name, '', `${it.title} — ${it.vatLabel}`, '', '', formatEuro(runCliente * (it.vatRate / 100))]);
            rows.push([g.name, '', `${it.title} — IVA inclusa`, '', '', formatEuro(runCliente * (1 + it.vatRate / 100))]);
          }
        } else {
          rows.push([g.name, '', `— ${it.title} —`]);
        }
        runImpresa = 0;
        runCliente = 0;
        return;
      }
      const qty = parseEuro(it.qty);
      const totImpresa = parseEuro(it.unitPriceImpresa) * qty;
      const totCliente = parseEuro(it.unitPriceCliente) * qty;
      runImpresa += totImpresa;
      runCliente += totCliente;
      rows.push(clientOnly
        ? [g.name, it.code, it.desc, it.qty, it.unit, it.unitPriceCliente, formatEuro(totCliente)]
        : [g.name, it.code, it.desc, it.qty, it.unit, it.unitPriceImpresa, formatEuro(totImpresa), it.unitPriceCliente, formatEuro(totCliente)]);
    });
  });

  const realItems = (revision.items || []).filter((it) => it.type !== 'subtotal');
  const impresaTot = sumImpresa(realItems);
  const clienteTot = sumCliente(realItems);
  const { rate: vatRate, label: vatLabel } = getVatInfo(revision);
  rows.push([]);
  if (!clientOnly) {
    rows.push(['Totale generale IVA esclusa (impresa)', '', '', '', '', '', formatEuro(impresaTot)]);
    rows.push([`${vatLabel} (impresa)`, '', '', '', '', '', formatEuro(impresaTot * (vatRate / 100))]);
    rows.push(['Totale generale IVA inclusa (impresa)', '', '', '', '', '', formatEuro(impresaTot * (1 + vatRate / 100))]);
    rows.push([]);
  }
  rows.push(['Totale generale IVA esclusa (cliente)', '', '', '', '', '', formatEuro(clienteTot)]);
  rows.push([`${vatLabel} (cliente)`, '', '', '', '', '', formatEuro(clienteTot * (vatRate / 100))]);
  rows.push(['Totale generale IVA inclusa (cliente)', '', '', '', '', '', formatEuro(clienteTot * (1 + vatRate / 100))]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Computo');
  const filename = `${project.name.replace(/[^a-z0-9]+/gi, '_')}_${(revision.customName || revision.label).replace(/[^a-z0-9]+/gi, '_')}${clientOnly ? '_cliente' : ''}.xlsx`;
  XLSX.writeFile(wb, filename);
}

function PrintableComputo({ project, revision, clientOnly }) {
  const groups = groupItemsForExport(revision.items, revision.extraSections);
  const realItems = (revision.items || []).filter((it) => it.type !== 'subtotal');
  const impresaTot = sumImpresa(realItems);
  const clienteTot = sumCliente(realItems);
  const header = project.header || {};
  const { rate: vatRate, label: vatLabel } = getVatInfo(revision);

  return (
    <div className="print-only" style={{ padding: 24, fontFamily: FONT, color: '#000' }}>
      <h1 style={{ fontSize: 20, marginBottom: 4 }}>{clientOnly ? 'Computo metrico — versione cliente' : 'Computo metrico'}</h1>
      <p style={{ fontSize: 12, margin: '2px 0' }}>Progetto: {project.name} — Cliente: {project.client}</p>
      <p style={{ fontSize: 12, margin: '2px 0' }}>Versione: {revision.customName || revision.label} — Modificata il {revision.dateModified}</p>
      {(header.descrizione || header.ubicazione) && (
        <p style={{ fontSize: 12, margin: '2px 0' }}>{header.descrizione} {header.ubicazione && `— ${header.ubicazione}`}</p>
      )}
      {groups.map((g, gi) => {
        let runI = 0;
        let runC = 0;
        return (
        <div key={gi} style={{ marginTop: 14 }}>
          <h3 style={{ fontSize: 13, margin: '0 0 6px', borderBottom: '1px solid #000', paddingBottom: 2 }}>{g.name}</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
            <thead>
              <tr>
                <th style={{ textAlign: 'left', borderBottom: '1px solid #999', padding: 3 }}>Codice</th>
                <th style={{ textAlign: 'left', borderBottom: '1px solid #999', padding: 3 }}>Descrizione</th>
                <th style={{ textAlign: 'right', borderBottom: '1px solid #999', padding: 3 }}>Qtà</th>
                <th style={{ textAlign: 'left', borderBottom: '1px solid #999', padding: 3 }}>U.M.</th>
                {!clientOnly && <th style={{ textAlign: 'right', borderBottom: '1px solid #999', padding: 3 }}>Prezzo impresa</th>}
                {!clientOnly && <th style={{ textAlign: 'right', borderBottom: '1px solid #999', padding: 3 }}>Totale impresa</th>}
                <th style={{ textAlign: 'right', borderBottom: '1px solid #999', padding: 3 }}>Prezzo cliente</th>
                <th style={{ textAlign: 'right', borderBottom: '1px solid #999', padding: 3 }}>Totale cliente</th>
              </tr>
            </thead>
            <tbody>
              {g.items.map((it, ii) => {
                if (it.type === 'subtotal') {
                  const hasVat = it.vatRate !== null && it.vatRate !== undefined;
                  const rowEl = (
                    <tr key={ii}>
                      <td colSpan={clientOnly ? 6 : 8} style={{ padding: '4px 3px', fontWeight: 700 }}>
                        {it.title}
                        {hasVat && (
                          <span style={{ fontWeight: 400 }}>
                            {' '}— IVA esclusa {formatEuro(clientOnly ? runC : runI)}, {it.vatLabel} {formatEuro((clientOnly ? runC : runI) * (it.vatRate / 100))}, IVA inclusa {formatEuro((clientOnly ? runC : runI) * (1 + it.vatRate / 100))}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                  runI = 0; runC = 0;
                  return rowEl;
                }
                runI += parseEuro(it.unitPriceImpresa) * parseEuro(it.qty);
                runC += parseEuro(it.unitPriceCliente) * parseEuro(it.qty);
                return (
                <tr key={ii}>
                  <td style={{ padding: '3px' }}>{it.code}</td>
                  <td style={{ padding: '3px' }}>{it.desc}</td>
                  <td style={{ padding: '3px', textAlign: 'right' }}>{it.qty}</td>
                  <td style={{ padding: '3px' }}>{it.unit}</td>
                  {!clientOnly && <td style={{ padding: '3px', textAlign: 'right' }}>{it.unitPriceImpresa} €</td>}
                  {!clientOnly && <td style={{ padding: '3px', textAlign: 'right' }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>}
                  <td style={{ padding: '3px', textAlign: 'right' }}>{it.unitPriceCliente} €</td>
                  <td style={{ padding: '3px', textAlign: 'right' }}>{formatEuro(parseEuro(it.unitPriceCliente) * parseEuro(it.qty))}</td>
                </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        );
      })}
      <div style={{ marginTop: 20, maxWidth: 300, marginLeft: 'auto' }}>
        {!clientOnly && (
          <>
            <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between' }}><span>Totale generale IVA esclusa (impresa)</span><strong>{formatEuro(impresaTot)}</strong></p>
            <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between' }}><span>{vatLabel}</span><span>{formatEuro(impresaTot * (vatRate / 100))}</span></p>
            <p style={{ fontSize: 13, display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}><span>Totale generale IVA inclusa (impresa)</span><span>{formatEuro(impresaTot * (1 + vatRate / 100))}</span></p>
            <hr />
          </>
        )}
        <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between' }}><span>Totale generale IVA esclusa (cliente)</span><strong>{formatEuro(clienteTot)}</strong></p>
        <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between' }}><span>{vatLabel}</span><span>{formatEuro(clienteTot * (vatRate / 100))}</span></p>
        <p style={{ fontSize: 14, display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}><span>Totale generale IVA inclusa (cliente)</span><span>{formatEuro(clienteTot * (1 + vatRate / 100))}</span></p>
      </div>
    </div>
  );
}

function ProjectDetailPage({ project, onBack, onUpdateProject, listini, initialRevisionId, requestPdf }) {
  const revisions = project.revisions;
  const latestRevision = revisions[revisions.length - 1];
  const [selectedRevisionId, setSelectedRevisionId] = useState(initialRevisionId || latestRevision?.id);
  const [selRevA, setSelRevA] = useState(revisions[0]?.id);
  const [selRevB, setSelRevB] = useState(revisions[revisions.length - 1]?.id);
  const [showCompare, setShowCompare] = useState(false);
  const [listinoId, setListinoId] = useState(listini[0]?.id);
  const [dragOver, setDragOver] = useState(false);

  const selectedRevision = revisions.find((r) => r.id === selectedRevisionId) || latestRevision;
  const isEditingLatest = selectedRevision && latestRevision && selectedRevision.id === latestRevision.id;
  const items = selectedRevision?.items || [];
  const extraSections = selectedRevision?.extraSections || [];
  const sectionDiscounts = selectedRevision?.sectionDiscounts || {};
  const header = project.header || {};
  const activeListino = listini.find((l) => l.id === listinoId) || listini[0];

  const realItems = items.filter((it) => it.type !== 'subtotal');
  const { rate: vatRate, label: vatLabel } = getVatInfo(selectedRevision);

  const groupedSections = [];
  const sectionByName = (name) => {
    let section = groupedSections.find((s) => s.name === name);
    if (!section) {
      section = { name, color: SECTION_COLORS[groupedSections.length % SECTION_COLORS.length], items: [] };
      groupedSections.push(section);
    }
    return section;
  };
  extraSections.forEach((name) => sectionByName(name));
  items.forEach((it) => {
    const sectionName = it.section || it.macro || 'Voci varie';
    sectionByName(sectionName).items.push(it);
  });
  const allSectionNames = groupedSections.map((s) => s.name);

  let importoLavori = 0;
  let importoLavoriCliente = 0;
  groupedSections.forEach((s) => {
    const realSectionItems = s.items.filter((it) => it.type !== 'subtotal');
    s.subtotalImpresa = sumImpresa(realSectionItems);
    s.subtotalCliente = sumCliente(realSectionItems);
    s.discountPct = parseFloat(sectionDiscounts[s.name]) || 0;
    s.netImpresa = s.subtotalImpresa * (1 - s.discountPct / 100);
    s.netCliente = s.subtotalCliente * (1 - s.discountPct / 100);
    importoLavori += s.netImpresa;
    importoLavoriCliente += s.netCliente;
  });

  const iva = importoLavori * (vatRate / 100);
  const totale = importoLavori + iva;
  const ivaCliente = importoLavoriCliente * (vatRate / 100);
  const totaleCliente = importoLavoriCliente + ivaCliente;

  const updateVat = (field, value) => {
    applyRevisionChange(() => ({ [field]: value }));
  };
  const applyVatPreset = (rate, label) => {
    applyRevisionChange(() => ({ vatRate: rate, vatLabel: label }));
  };

  const updateSectionDiscount = (sectionName, value) => {
    applyRevisionChange((rev) => ({ sectionDiscounts: { ...(rev.sectionDiscounts || {}), [sectionName]: value } }));
  };

  // Applica una modifica alla revisione (voci, sezioni extra...): se si sta lavorando sull'ultima
  // versione la aggiorna sul posto, altrimenti crea automaticamente una copia lasciando quella aperta intatta.
  const applyRevisionChange = (updater) => {
    const patch = updater(selectedRevision);
    const newItems = (patch.items || items).filter((it) => it.type !== 'subtotal');
    const total = formatEuro(sumImpresa(newItems));
    const totalCliente = formatEuro(sumCliente(newItems));
    if (isEditingLatest) {
      const updatedRevisions = revisions.map((r) => (r.id === selectedRevision.id ? { ...r, ...patch, dateModified: nowLabel(), total, totalCliente } : r));
      onUpdateProject({ ...project, revisions: updatedRevisions, value: total });
    } else {
      const newRev = { ...selectedRevision, ...patch, id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], total, totalCliente };
      onUpdateProject({ ...project, revisions: [...revisions, newRev], value: total });
      setSelectedRevisionId(newRev.id);
    }
  };

  const applyItemsChange = (updater) => applyRevisionChange((rev) => ({ items: updater(rev.items || []) }));

  const startComputo = () => {
    const rev = { id: Date.now(), label: 'Revisione 1', customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: [], extraSections: [], total: '0,00 €' };
    onUpdateProject({ ...project, revisions: [rev], header: header });
    setSelectedRevisionId(rev.id);
  };

  const addComputoItem = (voce, qty = '1') => {
    const impresaVal = voce.impresaValue !== undefined ? voce.impresaValue : parseEuro(voce.priceImpresa);
    const clienteVal = voce.clienteValue !== undefined ? voce.clienteValue : evalClientPrice(voce.priceCliente, impresaVal);
    const item = { id: Date.now() + Math.random(), code: voce.code, desc: voce.desc, unit: voce.unit, unitPriceImpresa: formatEuro(impresaVal).replace(' €', ''), unitPriceCliente: formatEuro(clienteVal).replace(' €', ''), qty: String(qty), macro: voce.macro, section: voce.macro };
    applyItemsChange((its) => [...its, item]);
  };

  // Plugin 1: importa un file Excel esterno e collega i codici al listino attivo,
  // creando in automatico le voci del computo con le quantità indicate nel file.
  const importFromExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const catalogItems = flattenListino(activeListino);
        const matched = [];
        const unmatched = [];
        rows.forEach((row) => {
          const keys = Object.keys(row);
          const codeKey = keys.find((k) => /cod/i.test(k));
          const qtyKey = keys.find((k) => /quant|qta|qty/i.test(k));
          const code = codeKey ? String(row[codeKey]).trim() : '';
          const qty = qtyKey ? row[qtyKey] : '';
          if (!code) return;
          const found = catalogItems.find((v) => v.code.toLowerCase() === code.toLowerCase());
          if (found) matched.push({ voce: found, qty: qty || 1 });
          else unmatched.push(code);
        });
        if (matched.length === 0) {
          alert('Nessuna voce del file corrisponde a un codice del listino attivo. Verifica che il file abbia una colonna "Codice" e una colonna "Quantità".');
          return;
        }
        applyItemsChange((its) => [
          ...its,
          ...matched.map((m) => {
            const impresaVal = m.voce.impresaValue;
            const clienteVal = m.voce.clienteValue;
            return {
              id: Date.now() + Math.random(), code: m.voce.code, desc: m.voce.desc, unit: m.voce.unit,
              unitPriceImpresa: formatEuro(impresaVal).replace(' €', ''), unitPriceCliente: formatEuro(clienteVal).replace(' €', ''),
              qty: String(m.qty), macro: m.voce.macro, section: m.voce.macro,
            };
          }),
        ]);
        alert(`Importate ${matched.length} voci dal file.${unmatched.length ? '\n\nCodici non trovati nel listino: ' + unmatched.join(', ') : ''}`);
      } catch (err) {
        alert('Non sono riuscito a leggere il file. Verifica che sia un .xlsx valido.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Plugin 2: planimetrie con punti cliccabili collegati al listino, che finiscono nel computo.
  const uploadPlanimetria = (file) => {
    const reader = new FileReader();
    reader.onload = () => {
      const planimetria = { id: Date.now(), name: file.name, image: reader.result, markers: [] };
      onUpdateProject({ ...project, planimetrie: [...(project.planimetrie || []), planimetria] });
    };
    reader.readAsDataURL(file);
  };

  const removePlanimetria = (id) => {
    if (!confirm('Eliminare questa planimetria e tutti i suoi punti?')) return;
    onUpdateProject({ ...project, planimetrie: (project.planimetrie || []).filter((p) => p.id !== id) });
  };

  const addPuntoOnPlanimetria = (planimetriaId, xPct, yPct) => {
    const catalogItems = flattenListino(activeListino);
    const code = prompt('Codice voce di listino da collegare a questo punto (es. IT.EL.01.005 per un punto luce):');
    if (!code) return;
    const voce = catalogItems.find((v) => v.code.toLowerCase() === code.trim().toLowerCase());
    if (!voce) { alert('Codice non trovato nel listino attivo.'); return; }
    const qty = prompt(`Quantità di "${voce.desc}" per questo punto:`, '1') || '1';
    addComputoItem(voce, qty);
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === planimetriaId
        ? { ...p, markers: [...p.markers, { id: Date.now() + Math.random(), x: xPct, y: yPct, code: voce.code, desc: voce.desc }] }
        : p)),
    });
  };

  const removePunto = (planimetriaId, markerId) => {
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === planimetriaId ? { ...p, markers: p.markers.filter((m) => m.id !== markerId) } : p)),
    });
  };

  const updateQty = (id, qty) => {
    applyItemsChange((its) => its.map((it) => (it.id === id ? { ...it, qty } : it)));
  };

  const removeItem = (id) => {
    applyItemsChange((its) => its.filter((it) => it.id !== id));
  };

  const moveItemToSection = (id, sectionName) => {
    applyItemsChange((its) => its.map((it) => (it.id === id ? { ...it, section: sectionName } : it)));
  };

  // Sposta una voce su/giù, scambiandola con la voce precedente/successiva della stessa sezione.
  const moveItemInSection = (id, direction) => {
    applyItemsChange((its) => {
      const item = its.find((it) => it.id === id);
      if (!item) return its;
      const sectionName = item.section || item.macro || 'Voci varie';
      const sameSectionIdx = its
        .map((it, idx) => ({ it, idx }))
        .filter((o) => (o.it.section || o.it.macro || 'Voci varie') === sectionName)
        .map((o) => o.idx);
      const idxA = its.indexOf(item);
      const posInSection = sameSectionIdx.indexOf(idxA);
      const swapPos = direction === 'up' ? posInSection - 1 : posInSection + 1;
      if (swapPos < 0 || swapPos >= sameSectionIdx.length) return its;
      const idxB = sameSectionIdx[swapPos];
      const next = [...its];
      [next[idxA], next[idxB]] = [next[idxB], next[idxA]];
      return next;
    });
  };

  const addCustomSection = () => {
    const name = prompt('Nome della nuova macrosezione:');
    if (!name) return;
    applyRevisionChange((rev) => ({ extraSections: [...(rev.extraSections || []), name] }));
  };

  const renameSection = (oldName) => {
    const newName = prompt('Rinomina macrosezione:', oldName);
    if (!newName || newName === oldName) return;
    applyRevisionChange((rev) => ({
      items: (rev.items || []).map((it) => ((it.section || it.macro) === oldName ? { ...it, section: newName } : it)),
      extraSections: (rev.extraSections || []).map((n) => (n === oldName ? newName : n)),
    }));
  };

  const addPartialSubtotal = (sectionName) => {
    const title = prompt('Titolo della sommatoria parziale:', 'Sommatoria parziale');
    if (!title) return;
    const applyVat = confirm('Applicare l\'IVA a questa sommatoria parziale? (Annulla = nessuna IVA qui)');
    let vatRate = null;
    let vatLabel = null;
    if (applyVat) {
      const defaultRate = selectedRevision?.vatRate !== undefined && selectedRevision?.vatRate !== null ? selectedRevision.vatRate : 22;
      const rateInput = prompt('Aliquota IVA per questa sommatoria (%):', String(defaultRate));
      if (rateInput !== null) {
        vatRate = Number(rateInput) || 0;
        vatLabel = prompt('Dicitura IVA (facoltativa):', `IVA ${vatRate}%`) || `IVA ${vatRate}%`;
      }
    }
    const marker = { id: Date.now() + Math.random(), type: 'subtotal', title, section: sectionName, vatRate, vatLabel };
    applyItemsChange((its) => [...its, marker]);
  };

  const editPartialSubtotal = (id) => {
    applyItemsChange((its) => its.map((it) => {
      if (it.id !== id) return it;
      const title = prompt('Titolo della sommatoria parziale:', it.title) || it.title;
      const applyVat = confirm('Applicare l\'IVA a questa sommatoria parziale? (Annulla = nessuna IVA qui)');
      let vatRate = null;
      let vatLabel = null;
      if (applyVat) {
        const rateInput = prompt('Aliquota IVA per questa sommatoria (%):', String(it.vatRate ?? 22));
        if (rateInput !== null) {
          vatRate = Number(rateInput) || 0;
          vatLabel = prompt('Dicitura IVA (facoltativa):', it.vatLabel || `IVA ${vatRate}%`) || `IVA ${vatRate}%`;
        }
      }
      return { ...it, title, vatRate, vatLabel };
    }));
  };

  const removeMarkerOrEmptySection = (sectionName) => {
    if (!confirm(`Rimuovere la macrosezione "${sectionName}"? (possibile solo se vuota)`)) return;
    applyRevisionChange((rev) => ({ extraSections: (rev.extraSections || []).filter((n) => n !== sectionName) }));
  };

  const updateHeader = (field, value) => {
    onUpdateProject({ ...project, header: { ...header, [field]: value } });
  };

  const saveNewVersion = () => {
    const total = formatEuro(sumImpresa(realItems));
    const totalCliente = formatEuro(sumCliente(realItems));
    const newRev = { id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: structuredClone(items), extraSections: structuredClone(extraSections), total, totalCliente };
    onUpdateProject({ ...project, revisions: [...revisions, newRev], value: total });
    setSelectedRevisionId(newRev.id);
  };

  const openRevision = (id) => setSelectedRevisionId(id);

  const deleteRevision = (id) => {
    if (!confirm('Eliminare definitivamente questa versione del computo?')) return;
    const remaining = revisions.filter((r) => r.id !== id);
    onUpdateProject({ ...project, revisions: remaining });
    if (selectedRevisionId === id) setSelectedRevisionId(remaining[remaining.length - 1]?.id);
  };

  const changeRevisionStatus = (id, status) => {
    onUpdateProject({ ...project, revisions: revisions.map((r) => (r.id === id ? { ...r, status } : r)) });
  };

  const renameRevision = (id) => {
    const rev = revisions.find((r) => r.id === id);
    const newName = prompt('Nuovo nome del computo (data di modifica e numero revisione restano invariati):', rev.customName || rev.label);
    if (!newName) return;
    onUpdateProject({ ...project, revisions: revisions.map((r) => (r.id === id ? { ...r, customName: newName } : r)) });
  };

  const uploadDocuments = (fileList) => {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;
    let remaining = files.length;
    const newDocs = [];
    files.forEach((file) => {
      const reader = new FileReader();
      reader.onload = () => {
        newDocs.push({ id: Date.now() + Math.random(), name: file.name, size: file.size, type: file.type, dataUrl: reader.result, uploadedAt: nowLabel() });
        remaining -= 1;
        if (remaining === 0) {
          onUpdateProject({ ...project, documents: [...(project.documents || []), ...newDocs] });
        }
      };
      reader.readAsDataURL(file);
    });
  };

  const removeDocument = (id) => {
    if (!confirm('Eliminare questo documento dal progetto?')) return;
    onUpdateProject({ ...project, documents: (project.documents || []).filter((d) => d.id !== id) });
  };

  const formatFileSize = (bytes) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const addTeamMember = () => {
    const name = prompt('Nome della persona da assegnare:');
    if (!name) return;
    const role = prompt('Ruolo:', 'Collaboratore') || 'Collaboratore';
    onUpdateProject({ ...project, team: [...project.team, { name, role }] });
  };

  const removeTeamMember = (idx) => {
    onUpdateProject({ ...project, team: project.team.filter((_, i) => i !== idx) });
  };

  const headerField = (label, field, placeholder) => (
    <div>
      <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>{label}</label>
      <input
        value={header[field] || ''}
        onChange={(e) => updateHeader(field, e.target.value)}
        placeholder={placeholder}
        style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }}
      />
    </div>
  );

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Progetti / {project.name}</p>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
        <button onClick={onBack} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>‹ Progetti</button>
        <h1 style={h1Style}>{project.name}</h1>
        <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, ...badgeStyles[statusTone[latestStatus(project)]] }}>{latestStatus(project)}</span>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>
      <p style={{ fontSize: 12, color: C.gray, margin: '0 0 18px' }}>{project.client}</p>

      {revisions.length === 0 ? (
        <div style={{ ...card, marginBottom: 24, textAlign: 'center', padding: 36 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 14px' }}>Questo progetto non ha ancora un computo metrico.</p>
          <button onClick={startComputo} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>+ Crea primo computo metrico</button>
        </div>
      ) : (
        <>
          <div style={{ ...card, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Versione aperta</label>
            <select value={selectedRevisionId} onChange={(e) => setSelectedRevisionId(Number(e.target.value))} style={{ fontSize: 13, fontWeight: 600, padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
              {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · modificata il {r.dateModified}</option>)}
            </select>
            {!isEditingLatest && (
              <span style={{ fontSize: 11, color: C.darkGray, background: 'rgba(67,67,67,0.1)', padding: '5px 10px', borderRadius: 8 }}>
                Stai visualizzando una versione precedente: la prima modifica creerà automaticamente una nuova versione, lasciando questa intatta.
              </span>
            )}
          </div>

          <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
            <button onClick={saveNewVersion} style={{ background: C.maroon, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.white, cursor: 'pointer' }}>+ Salva nuova versione</button>
            <button onClick={() => exportComputoExcel(project, selectedRevision, false)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬚ Scarica Excel completo</button>
            <button onClick={() => exportComputoExcel(project, selectedRevision, true)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬚ Scarica Excel solo cliente</button>
            <button onClick={() => requestPdf(project, selectedRevision, false)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬇ Scarica PDF completo (impresa+cliente)</button>
            <button onClick={() => requestPdf(project, selectedRevision, true)} style={{ background: C.maroon, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.white, cursor: 'pointer' }}>⬇ Scarica PDF solo cliente</button>
          </div>

          <div style={{ ...card, marginBottom: 18 }}>
            <p style={{ fontWeight: 700, fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Dati generali del computo</p>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
              {headerField("Descrizione generale dell'opera", 'descrizione', 'Es. Ristrutturazione integrale di unità residenziale')}
              {headerField('Ubicazione cantiere', 'ubicazione', 'Indirizzo del cantiere')}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12, marginBottom: 12 }}>
              {headerField('Progettista', 'progettista', 'Nome del progettista')}
              {headerField('Direttore lavori', 'direttoreLavori', 'Nome del direttore lavori')}
              {headerField('Impresa esecutrice', 'impresa', 'Ragione sociale impresa')}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              {headerField('Numero pratica / commessa', 'numeroPratica', 'Es. 24/2026')}
              {headerField('Data documento', 'dataDocumento', 'Es. 31/07/2026')}
            </div>
            <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0' }}>
              Committente: {project.client} · Progetto: {project.name}
            </p>
          </div>

          <div style={{ display: 'flex', gap: 18, alignItems: 'flex-start', flexWrap: 'wrap' }}>
            <div style={{ ...card, width: 300, flexShrink: 0, border: `2px solid ${C.maroon}` }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                <p style={{ fontSize: 14, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Listino</p>
              </div>
              <select value={listinoId} onChange={(e) => setListinoId(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '6px 0 10px' }}>
                {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
              </select>
              <p style={{ fontSize: 11, color: C.gray, margin: '0 0 10px' }}>Apri le categorie per trovare la voce giusta e trascinala nel computo a destra.</p>
              <label style={{ display: 'block', textAlign: 'center', background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '8px 0', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer', marginBottom: 12 }}>
                📥 Importa voci da Excel
                <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={(e) => { if (e.target.files[0]) importFromExcel(e.target.files[0]); e.target.value = ''; }} />
              </label>
              <p style={{ fontSize: 10, color: C.gray, margin: '-6px 0 10px' }}>Il file deve avere una colonna "Codice" e una "Quantità": le voci con codice corrispondente al listino attivo vengono aggiunte in automatico al computo.</p>
              <div style={{ maxHeight: 560, overflowY: 'auto' }}>
                <DraggableCatalogTree listino={activeListino} />
              </div>
            </div>

            <div style={{ flex: 1, minWidth: 320 }}>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                <button onClick={addCustomSection} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Nuova macrosezione</button>
              </div>
              <div
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragOver(false);
                  const data = e.dataTransfer.getData('application/json');
                  if (!data) return;
                  addComputoItem(JSON.parse(data));
                }}
                style={{
                  border: `2px dashed ${dragOver ? C.maroon : 'rgba(23,107,99,0.4)'}`,
                  borderRadius: 12,
                  background: dragOver ? 'rgba(128,20,48,0.04)' : 'rgba(23,107,99,0.03)',
                  padding: 16,
                  marginBottom: 18,
                  minHeight: 160,
                }}
              >
                {groupedSections.length === 0 ? (
                  <p style={{ fontSize: 13, color: C.gray, textAlign: 'center', margin: '40px 0' }}>Trascina qui le voci dal Listino per costruire il computo metrico.<br />Man mano che aggiungi voci, il computo si aggiorna qui.</p>
                ) : (
                  groupedSections.map((section) => {
                    let runningImpresa = 0;
                    let runningCliente = 0;
                    return (
                      <div key={section.name} style={{ border: `1px solid ${C.paleGray}`, borderRadius: 10, overflow: 'hidden', marginBottom: 14, background: C.white }}>
                        <div style={{ background: section.color, color: C.white, padding: '8px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                          <span style={{ fontWeight: 700, fontSize: 13, fontFamily: FONT }}>{section.name}</span>
                          <div style={{ display: 'flex', gap: 6 }}>
                            <button onClick={() => renameSection(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>✎ Rinomina</button>
                            <button onClick={() => addPartialSubtotal(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>+ Sommatoria parziale</button>
                            {section.items.length === 0 && (
                              <button onClick={() => removeMarkerOrEmptySection(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>🗑</button>
                            )}
                          </div>
                        </div>
                        {section.items.length === 0 ? (
                          <p style={{ fontSize: 12, color: C.gray, padding: '10px 14px' }}>Nessuna voce ancora in questa sezione.</p>
                        ) : (
                          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                            <thead>
                              <tr style={{ textAlign: 'left', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                                <th style={{ padding: '8px 6px' }}></th>
                                <th style={{ padding: '8px 6px' }}>Codice</th>
                                <th style={{ padding: '8px 6px' }}>Descrizione</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Quantità</th>
                                <th style={{ padding: '8px 6px' }}>U.M.</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Prezzo impresa</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Totale impresa</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Prezzo cliente</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Totale cliente</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Listino</th>
                                <th style={{ padding: '8px 6px', textAlign: 'right' }}>Sconto</th>
                                <th style={{ padding: '8px 6px' }}>Sezione</th>
                                <th style={{ padding: '8px 6px' }}></th>
                              </tr>
                            </thead>
                            <tbody>
                              {section.items.map((it) => {
                                if (it.type === 'subtotal') {
                                  const hasVat = it.vatRate !== null && it.vatRate !== undefined;
                                  const ivaImpresaPart = hasVat ? runningImpresa * (it.vatRate / 100) : 0;
                                  const ivaClientePart = hasVat ? runningCliente * (it.vatRate / 100) : 0;
                                  const row = (
                                    <tr key={it.id} style={{ borderTop: `2px solid ${C.paleGray}`, background: 'rgba(128,20,48,0.05)' }}>
                                      <td colSpan={6} style={{ padding: '8px 6px', fontWeight: 700, color: C.maroon, verticalAlign: 'top' }}>
                                        {it.title}
                                        <button onClick={() => editPartialSubtotal(it.id)} style={{ ...rowBtnStyle, marginLeft: 8, padding: '1px 6px', fontSize: 10 }}>✎</button>
                                      </td>
                                      <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.maroon, verticalAlign: 'top' }}>
                                        {hasVat ? (
                                          <>
                                            <div>IVA escl. {formatEuro(runningImpresa)}</div>
                                            <div style={{ fontWeight: 400, fontSize: 10 }}>{it.vatLabel} {formatEuro(ivaImpresaPart)}</div>
                                            <div>IVA incl. {formatEuro(runningImpresa + ivaImpresaPart)}</div>
                                          </>
                                        ) : formatEuro(runningImpresa)}
                                      </td>
                                      <td></td>
                                      <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.maroon, verticalAlign: 'top' }}>
                                        {hasVat ? (
                                          <>
                                            <div>IVA escl. {formatEuro(runningCliente)}</div>
                                            <div style={{ fontWeight: 400, fontSize: 10 }}>{it.vatLabel} {formatEuro(ivaClientePart)}</div>
                                            <div>IVA incl. {formatEuro(runningCliente + ivaClientePart)}</div>
                                          </>
                                        ) : formatEuro(runningCliente)}
                                      </td>
                                      <td style={{ padding: '8px 6px' }}></td>
                                      <td style={{ padding: '8px 6px' }}></td>
                                      <td style={{ padding: '8px 6px' }}></td>
                                      <td style={{ padding: '8px 6px' }}></td>
                                    </tr>
                                  );
                                  runningImpresa = 0;
                                  runningCliente = 0;
                                  return row;
                                }
                                runningImpresa += parseEuro(it.unitPriceImpresa) * parseEuro(it.qty);
                                runningCliente += parseEuro(it.unitPriceCliente) * parseEuro(it.qty);
                                return (
                                  <tr key={it.id} style={{ borderTop: `1px solid ${C.paleGray}` }}>
                                    <td style={{ padding: '8px 6px' }}>
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                        <button onClick={() => moveItemInSection(it.id, 'up')} style={{ ...iconBtn, height: 18, fontSize: 9, lineHeight: '16px' }}>▲</button>
                                        <button onClick={() => moveItemInSection(it.id, 'down')} style={{ ...iconBtn, height: 18, fontSize: 9, lineHeight: '16px' }}>▼</button>
                                      </div>
                                    </td>
                                    <td style={{ padding: '8px 6px', fontWeight: 700, color: C.black }}>{it.code}</td>
                                    <td style={{ padding: '8px 6px', color: C.midGray }}>{it.desc}</td>
                                    <td style={{ padding: '8px 6px', textAlign: 'right' }}>
                                      <input
                                        value={it.qty}
                                        onChange={(e) => updateQty(it.id, e.target.value)}
                                        style={{ width: 60, fontSize: 12, padding: '5px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}`, textAlign: 'right' }}
                                      />
                                    </td>
                                    <td style={{ padding: '8px 6px', color: C.gray }}>{it.unit}</td>
                                    <td style={{ padding: '8px 6px', textAlign: 'right' }}>{it.unitPriceImpresa} €</td>
                                    <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.black }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>
                                    <td style={{ padding: '8px 6px', textAlign: 'right', color: C.maroon }}>{it.unitPriceCliente} €</td>
                                    <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.maroon }}>{formatEuro(parseEuro(it.unitPriceCliente) * parseEuro(it.qty))}</td>
                                    <td style={{ padding: '8px 6px' }}>
                                      <select value={section.name} onChange={(e) => moveItemToSection(it.id, e.target.value)} style={{ fontSize: 11, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}>
                                        {allSectionNames.map((n) => <option key={n} value={n}>{n}</option>)}
                                      </select>
                                    </td>
                                    <td style={{ padding: '8px 6px' }}>
                                      <button onClick={() => removeItem(it.id)} style={iconBtn}>🗑</button>
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        )}
                        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.paleGray}` }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, fontWeight: 700, color: C.black, marginBottom: section.discountPct ? 6 : 0 }}>
                            <span>Subtotale {section.name} (impresa)&nbsp;&nbsp;{formatEuro(section.subtotalImpresa)}</span>
                            <span style={{ color: C.maroon }}>Subtotale {section.name} (cliente)&nbsp;&nbsp;{formatEuro(section.subtotalCliente)}</span>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 11 }}>
                            <label style={{ color: C.gray }}>Sconto su questa sezione:</label>
                            <input
                              type="number"
                              value={sectionDiscounts[section.name] || ''}
                              onChange={(e) => updateSectionDiscount(section.name, e.target.value)}
                              placeholder="0"
                              style={{ width: 60, fontSize: 11, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                            />
                            <span style={{ color: C.gray }}>%</span>
                            {section.discountPct > 0 && (
                              <span style={{ marginLeft: 'auto', fontWeight: 700, color: C.black }}>
                                Netto impresa {formatEuro(section.netImpresa)} · <span style={{ color: C.maroon }}>netto cliente {formatEuro(section.netCliente)}</span>
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              <div style={{ ...card, marginBottom: 18, maxWidth: 660 }}>
                <p style={{ fontWeight: 700, fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Configurazione IVA</p>
                <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 10 }}>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Aliquota IVA (%)</label>
                    <input
                      type="number"
                      value={selectedRevision.vatRate !== undefined && selectedRevision.vatRate !== null ? selectedRevision.vatRate : 22}
                      onChange={(e) => updateVat('vatRate', e.target.value === '' ? '' : Number(e.target.value))}
                      style={{ width: 90, fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4, display: 'block' }}
                    />
                  </div>
                  <div style={{ flex: 1, minWidth: 220 }}>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Dicitura personalizzata (facoltativa)</label>
                    <input
                      value={selectedRevision.vatLabel || ''}
                      onChange={(e) => updateVat('vatLabel', e.target.value)}
                      placeholder='Es. "Esente IVA", "IVA non dovuta", "IVA dovuta 50% al 22% e 50% al 10%"'
                      style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }}
                    />
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 8 }}>
                  <button onClick={() => applyVatPreset(22, null)} style={rowBtnStyle}>Standard 22%</button>
                  <button onClick={() => applyVatPreset(10, null)} style={rowBtnStyle}>Agevolata 10%</button>
                  <button onClick={() => applyVatPreset(4, null)} style={rowBtnStyle}>Agevolata 4%</button>
                  <button onClick={() => applyVatPreset(0, 'Esente IVA')} style={rowBtnStyle}>Esente IVA</button>
                  <button onClick={() => applyVatPreset(0, 'IVA non dovuta')} style={rowBtnStyle}>IVA non dovuta</button>
                  <button onClick={() => applyVatPreset(16, 'IVA dovuta 50% al 22% e 50% al 10%')} style={rowBtnStyle}>Mista 50%/50% (22%+10%)</button>
                </div>
                <p style={{ fontSize: 11, color: C.gray, margin: 0 }}>
                  Per aliquote miste (es. 50% delle opere al 22% e 50% al 10%), imposta l'aliquota media effettiva (in questo caso 16%) nel campo numerico, e descrivi la ripartizione nella dicitura personalizzata: verrà mostrata al posto di "IVA {vatRate}%" nei quadri economici e nei documenti esportati.
                </p>
              </div>

              <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap' }}>
                <div style={{ ...card, maxWidth: 320, flex: 1 }}>
                  <p style={{ fontWeight: 700, fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Quadro economico — Impresa</p>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginBottom: 8 }}>
                    <span style={{ color: C.gray }}>Totale IVA esclusa</span>
                    <span style={{ fontWeight: 700, color: C.black }}>{formatEuro(importoLavori)}</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginBottom: 10, paddingBottom: 10, borderBottom: `1px solid ${C.paleGray}` }}>
                    <span style={{ color: C.gray }}>{vatLabel}</span>
                    <span style={{ color: C.black }}>{formatEuro(iva)}</span>
                  </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 16, fontWeight: 700 }}>
                    <span style={{ color: C.black }}>Totale IVA inclusa</span>
                    <span style={{ color: C.black }}>{formatEuro(totale)}</span>
                  </div>
                </div>

                <div style={{ ...card, maxWidth: 320, flex: 1, border: `1px solid ${C.maroon}` }}>
                  <p style={{ fontWeight: 700, fontSize: 15, margin: '0 0 12px', color: C.maroon, fontFamily: FONT }}>Quadro economico — Cliente</p>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginBottom: 8 }}>
                    <span style={{ color: C.gray }}>Totale IVA esclusa</span>
                    <span style={{ fontWeight: 700, color: C.black }}>{formatEuro(importoLavoriCliente)}</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginBottom: 10, paddingBottom: 10, borderBottom: `1px solid ${C.paleGray}` }}>
                    <span style={{ color: C.gray }}>{vatLabel}</span>
                    <span style={{ color: C.black }}>{formatEuro(ivaCliente)}</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 16, fontWeight: 700 }}>
                    <span style={{ color: C.maroon }}>Totale IVA inclusa</span>
                    <span style={{ color: C.maroon }}>{formatEuro(totaleCliente)}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </>
      )}

      {project.revisions.length > 0 && (
        <div style={{ ...card, marginTop: 24, marginBottom: 24 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Revisioni salvate</p>
          </div>
          {revisions.map((r) => (
            <div key={r.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: `1px solid ${C.paleGray}`, gap: 10, flexWrap: 'wrap' }}>
              <div>
                <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{r.customName || r.label} {r.customName && <span style={{ fontSize: 10, color: C.gray, fontWeight: 400 }}>({r.label})</span>} {r.id === selectedRevisionId && <span style={{ fontSize: 10, color: C.maroon }}>(aperta)</span>}</p>
                <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>Creata il {r.dateCreated} · Modificata il {r.dateModified} · {r.total} impresa / {r.totalCliente || r.total} cliente</p>
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <select value={r.status} onChange={(e) => changeRevisionStatus(r.id, e.target.value)} style={{ fontSize: 11, fontWeight: 600, padding: '5px 8px', borderRadius: 6, border: `1px solid ${C.paleGray}`, color: C.midGray }}>
                  {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
                <button onClick={() => openRevision(r.id)} style={rowBtnStyle}>Apri</button>
                <button onClick={() => renameRevision(r.id)} style={rowBtnStyle}>✎ Rinomina</button>
                <button onClick={() => exportComputoExcel(project, r, false)} style={rowBtnStyle}>Excel</button>
                <button onClick={() => requestPdf(project, r, false)} style={rowBtnStyle}>PDF</button>
                <button onClick={() => deleteRevision(r.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
              </div>
            </div>
          ))}

          {revisions.length > 1 && (
            <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${C.paleGray}` }}>
              <p style={{ fontSize: 13, fontWeight: 700, margin: '0 0 10px', color: C.black }}>Confronta due revisioni</p>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <select value={selRevA} onChange={(e) => setSelRevA(Number(e.target.value))} style={{ fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
                  {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
                </select>
                <span style={{ color: C.gray }}>→</span>
                <select value={selRevB} onChange={(e) => setSelRevB(Number(e.target.value))} style={{ fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
                  {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
                </select>
                <button onClick={() => setShowCompare(true)} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Confronta</button>
              </div>

              {showCompare && (
                <div style={{ marginTop: 16 }}>
                  <DiffTable diff={computeItemsDiff(revisions.find((r) => r.id === selRevA)?.items, revisions.find((r) => r.id === selRevB)?.items)} />
                </div>
              )}
            </div>
          )}
        </div>
      )}

      <div style={{ ...card, marginBottom: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Documenti di progetto</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>Planimetrie, capitolati, foto, DWG, PDF e altri file relativi a questo computo.</p>
          </div>
          <label style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            + Carica file
            <input
              type="file"
              multiple
              onChange={(e) => { uploadDocuments(e.target.files); e.target.value = ''; }}
              style={{ display: 'none' }}
            />
          </label>
        </div>
        {(!project.documents || project.documents.length === 0) && (
          <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessun documento caricato ancora.</p>
        )}
        {(project.documents || []).map((d) => (
          <div key={d.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: `1px solid ${C.paleGray}` }}>
            <div>
              <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{d.name}</p>
              <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>{formatFileSize(d.size)} · caricato il {d.uploadedAt}</p>
            </div>
            <div style={{ display: 'flex', gap: 6 }}>
              <a href={d.dataUrl} download={d.name} style={{ ...rowBtnStyle, textDecoration: 'none', display: 'inline-block' }}>Scarica</a>
              <button onClick={() => removeDocument(d.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
            </div>
          </div>
        ))}
        <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0' }}>
          In questa anteprima i file restano in memoria per la sessione corrente; nella versione online andranno salvati su uno storage reale.
        </p>
      </div>

      <div style={{ ...card, marginBottom: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Planimetrie</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>Carica una planimetria e clicca sopra per aggiungere punti (es. punti luce) collegati a una voce di listino: finiscono in automatico nel computo.</p>
          </div>
          <label style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            + Carica planimetria
            <input type="file" accept="image/*" style={{ display: 'none' }} onChange={(e) => { if (e.target.files[0]) uploadPlanimetria(e.target.files[0]); e.target.value = ''; }} />
          </label>
        </div>
        {(!project.planimetrie || project.planimetrie.length === 0) && (
          <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna planimetria caricata ancora. Accetta immagini JPG/PNG (i PDF vanno caricati come documento nella sezione sopra).</p>
        )}
        {(project.planimetrie || []).map((pl) => (
          <div key={pl.id} style={{ marginBottom: 18, border: `1px solid ${C.paleGray}`, borderRadius: 12, overflow: 'hidden' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: C.surfaceSubtle }}>
              <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{pl.name}</span>
              <button onClick={() => removePlanimetria(pl.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
            </div>
            <div
              onClick={(e) => {
                const rect = e.currentTarget.getBoundingClientRect();
                const xPct = ((e.clientX - rect.left) / rect.width) * 100;
                const yPct = ((e.clientY - rect.top) / rect.height) * 100;
                addPuntoOnPlanimetria(pl.id, xPct, yPct);
              }}
              style={{ position: 'relative', cursor: 'crosshair', lineHeight: 0 }}
            >
              <img src={pl.image} alt={pl.name} style={{ width: '100%', display: 'block' }} />
              {pl.markers.map((m) => (
                <div
                  key={m.id}
                  onClick={(e) => { e.stopPropagation(); if (confirm(`Rimuovere il punto "${m.desc}"?`)) removePunto(pl.id, m.id); }}
                  title={`${m.code} — ${m.desc} (clicca per rimuovere)`}
                  style={{
                    position: 'absolute', left: `${m.x}%`, top: `${m.y}%`, transform: 'translate(-50%, -50%)',
                    width: 22, height: 22, borderRadius: 999, background: C.maroon, color: C.white,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700,
                    border: '2px solid white', boxShadow: '0 1px 3px rgba(0,0,0,0.3)', cursor: 'pointer',
                  }}
                >
                  {pl.markers.indexOf(m) + 1}
                </div>
              ))}
            </div>
            {pl.markers.length > 0 && (
              <div style={{ padding: '8px 12px', fontSize: 11, color: C.darkGray }}>
                {pl.markers.length} punti aggiunti al computo · clicca un punto per rimuoverlo
              </div>
            )}
          </div>
        ))}
      </div>

      <div style={card}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Team assegnato</p>
          <button onClick={addTeamMember} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Aggiungi persona</button>
        </div>
        {project.team.length === 0 && <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna persona assegnata a questo progetto.</p>}
        {project.team.map((m, i) => (
          <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: i < project.team.length - 1 ? `1px solid ${C.paleGray}` : 'none' }}>
            <div>
              <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{m.name}</p>
              <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>{m.role}</p>
            </div>
            <button onClick={() => removeTeamMember(i)} style={rowBtnStyle}>🗑</button>
          </div>
        ))}
      </div>
    </div>
  );
}

const rowBtnStyle = { border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 6, fontSize: 11, fontWeight: 600, padding: '5px 10px', cursor: 'pointer', color: C.midGray };

function DiffTable({ diff }) {
  const esitoColor = { Invariata: C.gray, Aggiunta: C.maroon, Modificata: C.darkGray, Rimossa: C.black };
  if (!diff || diff.length === 0) {
    return <p style={{ fontSize: 12, color: C.gray }}>Nessuna voce da confrontare tra queste due versioni.</p>;
  }
  return (
    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
      <thead>
        <tr style={{ textAlign: 'left', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
          <th style={{ padding: '8px 6px' }}>Esito</th>
          <th style={{ padding: '8px 6px' }}>Sezione</th>
          <th style={{ padding: '8px 6px' }}>Codice / Descrizione</th>
          <th style={{ padding: '8px 6px', textAlign: 'right' }}>Q.tà prima</th>
          <th style={{ padding: '8px 6px', textAlign: 'right' }}>Q.tà dopo</th>
          <th style={{ padding: '8px 6px', textAlign: 'right' }}>Prezzo prima</th>
          <th style={{ padding: '8px 6px', textAlign: 'right' }}>Prezzo dopo</th>
          <th style={{ padding: '8px 6px', textAlign: 'right' }}>Variazione</th>
        </tr>
      </thead>
      <tbody>
        {diff.map((r) => (
          <tr key={r.code} style={{ borderTop: `1px solid ${C.paleGray}`, background: r.highlight ? 'rgba(128,20,48,0.05)' : 'transparent' }}>
            <td style={{ padding: '10px 6px', fontWeight: 600, color: esitoColor[r.esito] }}>{r.esito}</td>
            <td style={{ padding: '10px 6px' }}>
              <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, background: 'rgba(129,129,129,0.14)', color: C.midGray }}>{r.sezione}</span>
            </td>
            <td style={{ padding: '10px 6px' }}>
              <p style={{ margin: 0, fontWeight: 700, color: C.black }}>{r.code}</p>
              <p style={{ margin: '2px 0 0', color: C.gray }}>{r.desc}</p>
            </td>
            <td style={{ padding: '10px 6px', textAlign: 'right' }}>{r.qtyBefore}</td>
            <td style={{ padding: '10px 6px', textAlign: 'right' }}>{r.qtyAfter}</td>
            <td style={{ padding: '10px 6px', textAlign: 'right' }}>{r.priceBefore}</td>
            <td style={{ padding: '10px 6px', textAlign: 'right' }}>{r.priceAfter}</td>
            <td style={{ padding: '10px 6px', textAlign: 'right', fontWeight: 700, color: r.variation.startsWith('+') ? C.maroon : C.black }}>{r.variation}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function ComputiPage({ projects, setProjects, onOpenProject, onOpenRevision, requestPdf }) {
  const [activeId, setActiveId] = useState(projects[0]?.id);
  const project = projects.find((p) => p.id === activeId) || projects[0];

  const updateProject = (updated) => setProjects(projects.map((p) => (p.id === updated.id ? updated : p)));
  const changeStatus = (revId, status) => {
    updateProject({ ...project, revisions: project.revisions.map((r) => (r.id === revId ? { ...r, status } : r)) });
  };
  const deleteRevision = (revId) => {
    if (!confirm('Eliminare definitivamente questa versione del computo?')) return;
    updateProject({ ...project, revisions: project.revisions.filter((r) => r.id !== revId) });
  };
  const renameRevision = (revId) => {
    const rev = project.revisions.find((r) => r.id === revId);
    const newName = prompt('Nuovo nome del computo (data di modifica e numero revisione restano invariati):', rev.customName || rev.label);
    if (!newName) return;
    updateProject({ ...project, revisions: project.revisions.map((r) => (r.id === revId ? { ...r, customName: newName } : r)) });
  };

  if (!project) return <p style={{ fontSize: 13, color: C.gray }}>Nessun progetto ancora creato.</p>;

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Computi</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 16 }}>
        <h1 style={h1Style}>Computi</h1>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>
      <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>
        Panoramica di tutti i computi creati, divisi per progetto. La creazione avviene dentro ogni singolo progetto; da qui puoi aprirli e modificarli — ogni modifica crea automaticamente una nuova versione.
      </p>

      <div style={{ display: 'flex', gap: 6, marginBottom: 18, flexWrap: 'wrap', borderBottom: `1px solid ${C.paleGray}`, paddingBottom: 10 }}>
        {projects.map((p) => (
          <button
            key={p.id}
            onClick={() => setActiveId(p.id)}
            style={{
              border: 'none', borderRadius: 8, padding: '9px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              background: p.id === project.id ? C.maroon : C.white,
              color: p.id === project.id ? C.white : C.black,
              border: `1px solid ${p.id === project.id ? C.maroon : C.paleGray}`,
            }}
          >
            {p.name}
          </button>
        ))}
      </div>

      {project.revisions.length === 0 ? (
        <div style={{ ...card, textAlign: 'center', padding: 30 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 14px' }}>"{project.name}" non ha ancora nessun computo creato.</p>
          <button onClick={() => onOpenProject(project.id)} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Apri il progetto per crearlo</button>
        </div>
      ) : (
        <div style={card}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
            <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>{project.name} · {project.client}</p>
            <button onClick={() => onOpenProject(project.id)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>Apri progetto ›</button>
          </div>
          {project.revisions.map((r) => (
            <div key={r.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '12px 0', borderBottom: `1px solid ${C.paleGray}`, gap: 10, flexWrap: 'wrap' }}>
              <div>
                <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{r.customName || r.label} {r.customName && <span style={{ fontSize: 10, color: C.gray, fontWeight: 400 }}>({r.label})</span>}</p>
                <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>Creata il {r.dateCreated} · Modificata il {r.dateModified} · {r.total} impresa / {r.totalCliente || r.total} cliente</p>
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <select value={r.status} onChange={(e) => changeStatus(r.id, e.target.value)} style={{ fontSize: 11, fontWeight: 600, padding: '5px 8px', borderRadius: 6, border: `1px solid ${C.paleGray}`, color: C.midGray }}>
                  {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
                <button onClick={() => onOpenRevision(project.id, r.id)} style={rowBtnStyle}>Apri e modifica</button>
                <button onClick={() => renameRevision(r.id)} style={rowBtnStyle}>✎ Rinomina</button>
                <button onClick={() => exportComputoExcel(project, r, false)} style={rowBtnStyle}>Excel</button>
                <button onClick={() => requestPdf(project, r, false)} style={rowBtnStyle}>PDF</button>
                <button onClick={() => deleteRevision(r.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ConfrontoPage({ projects }) {
  const [projectId, setProjectId] = useState(projects[0]?.id);
  const project = projects.find((p) => p.id === projectId) || projects[0];
  const revisions = project.revisions;
  const [from, setFrom] = useState(revisions[0]?.id);
  const [to, setTo] = useState(revisions[revisions.length - 1]?.id);

  const handleProjectChange = (id) => {
    setProjectId(id);
    const proj = projects.find((p) => p.id === id);
    setFrom(proj.revisions[0]?.id);
    setTo(proj.revisions[proj.revisions.length - 1]?.id);
  };

  const esitoColor = { Invariata: C.gray, Aggiunta: C.maroon };

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Confronto revisioni</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 20 }}>
        <h1 style={h1Style}>Confronto revisioni</h1>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>

      <div style={{ ...card, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 10 }}>
        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Progetto</label>
        <select value={projectId} onChange={(e) => handleProjectChange(Number(e.target.value))} style={{ fontSize: 13, fontWeight: 600, padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, minWidth: 220 }}>
          {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>

      {revisions.length < 2 ? (
        <div style={{ ...card, textAlign: 'center', padding: 30 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: 0 }}>"{project.name}" ha meno di due revisioni salvate: non c'è ancora nulla da confrontare.</p>
        </div>
      ) : (
        <div style={card}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 18 }}>
            <div>
              <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Confronto revisioni</p>
              <p style={{ fontSize: 12, color: C.gray, margin: '2px 0 0' }}>{project.name} · modifiche evidenziate secondo la palette</p>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, background: 'rgba(128,20,48,0.12)', color: C.maroon }}>Aggiunte</span>
              <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, background: 'rgba(67,67,67,0.14)', color: C.darkGray }}>Modificate</span>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 18 }}>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Revisione di partenza</label>
              <select value={from} onChange={(e) => setFrom(Number(e.target.value))} style={{ width: '100%', fontSize: 13, fontWeight: 600, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }}>
                {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
              </select>
            </div>
            <span style={{ fontSize: 16, color: C.gray, marginTop: 16 }}>→</span>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Revisione di arrivo</label>
              <select value={to} onChange={(e) => setTo(Number(e.target.value))} style={{ width: '100%', fontSize: 13, fontWeight: 600, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }}>
                {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
              </select>
            </div>
          </div>

          <DiffTable diff={computeItemsDiff(revisions.find((r) => r.id === from)?.items, revisions.find((r) => r.id === to)?.items)} />
        </div>
      )}

      <div style={{ ...card, marginTop: 18 }}>
        <h2 style={{ fontSize: 15, margin: '0 0 10px', color: C.black, fontFamily: FONT }}>Stampa ed esportazione</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 12px' }}>
          Intestazione, subtotali e sezione Extra sono sempre inclusi. Mantieni nella stampa le evidenziazioni di aggiunte, modifiche e rimozioni.
        </p>
        <div style={{ display: 'flex', gap: 10 }}>
          <button onClick={() => alert('Esportazione Excel/CSV: funzione da collegare al backend reale.')} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Excel / CSV</button>
          <button onClick={() => alert('Stampa/PDF: genera il documento con le evidenziazioni delle modifiche.')} style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Stampa / PDF</button>
        </div>
      </div>
    </div>
  );
}

const INITIAL_FORNITORI = [
  {
    name: 'Forniture sanitarie', categorie: [
      { name: 'WC', prodotti: [
        { id: 1, name: 'WC sospeso Serie Rovere', photo: null, listinoPrice: '180,00', fornitori: [
          { id: 1, name: 'Idroterm Forniture S.r.l.', prezzoListino: '180,00', prezzoScontato: '150,00', prezzoCliente: '165,00' },
          { id: 2, name: 'Ceramiche Rossi', prezzoListino: '190,00', prezzoScontato: '160,00', prezzoCliente: '172,00' },
        ]},
      ]},
      { name: 'Lavabi', prodotti: [] },
      { name: 'Rubinetteria', prodotti: [] },
    ],
  },
  {
    name: 'Forniture elettriche', categorie: [
      { name: 'Quadri elettrici', prodotti: [] },
      { name: 'Corpi illuminanti', prodotti: [] },
    ],
  },
  {
    name: 'Forniture infissi', categorie: [
      { name: 'Finestre', prodotti: [] },
      { name: 'Porte interne', prodotti: [] },
    ],
  },
];

function AddToComputoModal({ prodotto, projects, onClose, onAdd }) {
  const eligible = projects.filter((p) => p.revisions.length > 0);
  const [projectId, setProjectId] = useState(eligible[0]?.id);
  const project = projects.find((p) => p.id === projectId);
  const [revisionId, setRevisionId] = useState(project?.revisions[project.revisions.length - 1]?.id);
  const priceOptions = [
    { id: 'listino', name: 'Prezzo di listino prodotto (nessuno sconto)', prezzoListino: prodotto.listinoPrice, prezzoScontato: prodotto.listinoPrice, prezzoCliente: prodotto.listinoPrice },
    ...prodotto.fornitori,
  ];
  const [priceSourceId, setPriceSourceId] = useState(priceOptions[0]?.id ?? 'listino');
  const [qty, setQty] = useState('1');
  const selected = priceOptions.find((f) => f.id === priceSourceId) || priceOptions[0];
  const listinoNum = parseEuro(selected.prezzoListino);
  const clienteNum = parseEuro(selected.prezzoCliente);
  const scontoPct = listinoNum > 0 ? ((listinoNum - clienteNum) / listinoNum) * 100 : 0;

  const handleProjectChange = (id) => {
    setProjectId(id);
    const proj = projects.find((p) => p.id === id);
    setRevisionId(proj.revisions[proj.revisions.length - 1]?.id);
  };

  if (eligible.length === 0) {
    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
        <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 380 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 16px' }}>Nessun progetto ha ancora un computo creato. Apri un progetto e crea il primo computo prima di aggiungere forniture.</p>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Chiudi</button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 420 }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 6px', color: C.black }}>Aggiungi al computo</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>{prodotto.name}</p>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Progetto</label>
        <select value={projectId} onChange={(e) => handleProjectChange(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {eligible.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Computo (versione)</label>
        <select value={revisionId} onChange={(e) => setRevisionId(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {project?.revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
        </select>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Fornitore / fonte prezzo</label>
        <select value={priceSourceId} onChange={(e) => setPriceSourceId(e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {priceOptions.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select>

        <div style={{ display: 'flex', gap: 10, fontSize: 12, background: C.bg, borderRadius: 8, padding: '10px 12px', marginBottom: 12 }}>
          <div style={{ flex: 1 }}><span style={{ color: C.gray }}>Listino</span><br /><strong>{selected.prezzoListino} €</strong></div>
          <div style={{ flex: 1 }}><span style={{ color: C.gray }}>Scontato (costo impresa)</span><br /><strong>{selected.prezzoScontato} €</strong></div>
          <div style={{ flex: 1 }}><span style={{ color: C.gray }}>Cliente</span><br /><strong style={{ color: C.maroon }}>{selected.prezzoCliente} €</strong></div>
          <div style={{ flex: 1 }}><span style={{ color: C.gray }}>Sconto vs listino</span><br /><strong>{scontoPct.toFixed(1)}%</strong></div>
        </div>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Quantità</label>
        <input value={qty} onChange={(e) => setQty(e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 18px' }} />

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              onAdd(projectId, revisionId, selected, qty || '1');
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            Aggiungi
          </button>
        </div>
      </div>
    </div>
  );
}

function ProdottoCard({ prodotto, onUpdate, onRemove, onAddToComputo }) {
  const cheapestScontato = prodotto.fornitori.length
    ? Math.min(...prodotto.fornitori.map((f) => parseEuro(f.prezzoScontato)))
    : null;

  const uploadPhoto = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onUpdate({ ...prodotto, photo: reader.result });
    reader.readAsDataURL(file);
  };

  const renameProdotto = () => {
    const name = prompt('Nome prodotto:', prodotto.name);
    if (!name) return;
    onUpdate({ ...prodotto, name });
  };

  const updateListinoPrice = () => {
    const price = prompt('Prezzo di listino (€):', prodotto.listinoPrice);
    if (price === null) return;
    onUpdate({ ...prodotto, listinoPrice: price });
  };

  const addFornitore = () => {
    const name = prompt('Nome fornitore:');
    if (!name) return;
    const prezzoListino = prompt(`Prezzo di listino del fornitore "${name}" (€):`, prodotto.listinoPrice) || '0,00';
    const prezzoScontato = prompt('Prezzo scontato (quanto paga davvero l\'impresa) (€):', prezzoListino) || prezzoListino;
    const prezzoCliente = prompt('Prezzo che verrà mostrato al cliente (€):', prezzoListino) || prezzoListino;
    onUpdate({ ...prodotto, fornitori: [...prodotto.fornitori, { id: Date.now(), name, prezzoListino, prezzoScontato, prezzoCliente }] });
  };

  const editFornitore = (fid) => {
    const f = prodotto.fornitori.find((x) => x.id === fid);
    const name = prompt('Nome fornitore:', f.name);
    if (!name) return;
    const prezzoListino = prompt('Prezzo di listino del fornitore (€):', f.prezzoListino);
    if (prezzoListino === null) return;
    const prezzoScontato = prompt('Prezzo scontato (€):', f.prezzoScontato);
    if (prezzoScontato === null) return;
    const prezzoCliente = prompt('Prezzo cliente (€):', f.prezzoCliente);
    if (prezzoCliente === null) return;
    onUpdate({ ...prodotto, fornitori: prodotto.fornitori.map((x) => (x.id === fid ? { ...x, name, prezzoListino, prezzoScontato, prezzoCliente } : x)) });
  };

  const removeFornitore = (fid) => {
    onUpdate({ ...prodotto, fornitori: prodotto.fornitori.filter((x) => x.id !== fid) });
  };

  return (
    <div style={{ ...card, display: 'flex', gap: 14, marginBottom: 12 }}>
      <label style={{ width: 84, height: 84, borderRadius: 10, border: `1px dashed ${C.paleGray}`, flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', overflow: 'hidden', background: C.bg }}>
        {prodotto.photo ? (
          <img src={prodotto.photo} alt={prodotto.name} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        ) : (
          <span style={{ fontSize: 10, color: C.gray, textAlign: 'center', padding: 4 }}>+ Foto</span>
        )}
        <input type="file" accept="image/*" style={{ display: 'none' }} onChange={(e) => uploadPhoto(e.target.files[0])} />
      </label>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 6 }}>
          <div>
            <p style={{ fontWeight: 700, fontSize: 14, margin: 0, color: C.black, fontFamily: FONT }}>{prodotto.name}</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>Prezzo di listino: <strong style={{ color: C.black }}>{prodotto.listinoPrice} €</strong></p>
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={onAddToComputo} style={{ ...rowBtnStyle, background: C.maroon, color: C.white, border: 'none' }}>+ Computo</button>
            <button onClick={renameProdotto} style={rowBtnStyle}>✎ Nome</button>
            <button onClick={updateListinoPrice} style={rowBtnStyle}>✎ Listino</button>
            <button onClick={onRemove} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, marginBottom: 4 }}>
          <p style={{ fontSize: 11, fontWeight: 700, color: C.midGray, margin: 0 }}>Fornitori</p>
          <button onClick={addFornitore} style={rowBtnStyle}>+ Fornitore</button>
        </div>
        {prodotto.fornitori.length === 0 ? (
          <p style={{ fontSize: 11, color: C.gray, margin: 0 }}>Nessun fornitore inserito.</p>
        ) : (
          <>
            <div style={{ display: 'flex', fontSize: 10, color: C.gray, textTransform: 'uppercase', padding: '2px 0' }}>
              <span style={{ flex: 1 }}>Fornitore</span>
              <span style={{ width: 70, textAlign: 'right' }}>Listino</span>
              <span style={{ width: 70, textAlign: 'right' }}>Scontato</span>
              <span style={{ width: 70, textAlign: 'right' }}>Cliente</span>
              <span style={{ width: 56 }}></span>
            </div>
            {[...prodotto.fornitori].sort((a, b) => parseEuro(a.prezzoScontato) - parseEuro(b.prezzoScontato)).map((f) => (
              <div key={f.id} style={{ display: 'flex', alignItems: 'center', padding: '6px 0', borderTop: `1px solid ${C.paleGray}` }}>
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{ fontSize: 12, color: C.black }}>{f.name}</span>
                  {parseEuro(f.prezzoScontato) === cheapestScontato && (
                    <span style={{ fontSize: 10, fontWeight: 700, color: C.maroon, background: 'rgba(128,20,48,0.1)', padding: '2px 7px', borderRadius: 999 }}>Migliore offerta</span>
                  )}
                </div>
                <span style={{ width: 70, textAlign: 'right', fontSize: 12, color: C.gray }}>{f.prezzoListino} €</span>
                <span style={{ width: 70, textAlign: 'right', fontSize: 12, fontWeight: 700, color: C.black }}>{f.prezzoScontato} €</span>
                <span style={{ width: 70, textAlign: 'right', fontSize: 12, color: C.maroon }}>{f.prezzoCliente} €</span>
                <span style={{ width: 56, display: 'flex', gap: 4, justifyContent: 'flex-end' }}>
                  <button onClick={() => editFornitore(f.id)} style={{ ...rowBtnStyle, padding: '2px 6px' }}>✎</button>
                  <button onClick={() => removeFornitore(f.id)} style={{ ...rowBtnStyle, padding: '2px 6px' }}>🗑</button>
                </span>
              </div>
            ))}
          </>
        )}
      </div>
    </div>
  );
}

function FornitoriPage({ projects, setProjects, catalog, setCatalog }) {
  const [expanded, setExpanded] = useState({});
  const [addingTo, setAddingTo] = useState(null); // prodotto selezionato per l'aggiunta al computo
  const isOpen = (key) => expanded[key] !== false;
  const toggle = (key) => setExpanded({ ...expanded, [key]: !isOpen(key) });

  const addToComputo = (projectId, revisionId, priceSource, qty) => {
    const prodotto = addingTo;
    const newItem = {
      id: Date.now() + Math.random(),
      code: '',
      desc: prodotto.name,
      unit: 'cad',
      unitPriceImpresa: priceSource.prezzoScontato,
      unitPriceCliente: priceSource.prezzoCliente,
      listinoRef: priceSource.prezzoListino,
      qty,
      macro: 'FORNITURE',
      section: 'FORNITURE',
    };
    setProjects(projects.map((p) => (p.id === projectId ? addItemToProjectRevision(p, revisionId, newItem) : p)));
  };

  const addMacro = () => {
    const name = prompt('Nome della nuova macrosezione (es. Forniture sanitarie):');
    if (!name) return;
    setCatalog([...catalog, { name, categorie: [] }]);
  };
  const renameMacro = (mi) => {
    const name = prompt('Rinomina macrosezione:', catalog[mi].name);
    if (!name) return;
    const next = structuredClone(catalog);
    next[mi].name = name;
    setCatalog(next);
  };
  const removeMacro = (mi) => {
    if (!confirm('Eliminare questa macrosezione e tutto il suo contenuto?')) return;
    setCatalog(catalog.filter((_, i) => i !== mi));
  };
  const addCategoria = (mi) => {
    const name = prompt('Nome della nuova sottocategoria (es. WC):');
    if (!name) return;
    const next = structuredClone(catalog);
    next[mi].categorie.push({ name, prodotti: [] });
    setCatalog(next);
  };
  const renameCategoria = (mi, ci) => {
    const name = prompt('Rinomina sottocategoria:', catalog[mi].categorie[ci].name);
    if (!name) return;
    const next = structuredClone(catalog);
    next[mi].categorie[ci].name = name;
    setCatalog(next);
  };
  const removeCategoria = (mi, ci) => {
    if (!confirm('Eliminare questa sottocategoria e tutto il suo contenuto?')) return;
    const next = structuredClone(catalog);
    next[mi].categorie.splice(ci, 1);
    setCatalog(next);
  };
  const addProdotto = (mi, ci) => {
    const name = prompt('Nome del nuovo prodotto/sanitario:');
    if (!name) return;
    const listinoPrice = prompt('Prezzo di listino (€):', '0,00') || '0,00';
    const next = structuredClone(catalog);
    next[mi].categorie[ci].prodotti.push({ id: Date.now(), name, photo: null, listinoPrice, fornitori: [] });
    setCatalog(next);
  };
  const updateProdotto = (mi, ci, pi, updated) => {
    const next = structuredClone(catalog);
    next[mi].categorie[ci].prodotti[pi] = updated;
    setCatalog(next);
  };
  const removeProdotto = (mi, ci, pi) => {
    if (!confirm('Eliminare questo prodotto?')) return;
    const next = structuredClone(catalog);
    next[mi].categorie[ci].prodotti.splice(pi, 1);
    setCatalog(next);
  };

  return (
    <div>
      <p style={breadcrumb}>Gestionale / Fornitori</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 20 }}>
        <h1 style={h1Style}>Fornitori</h1>
        <span style={{ ...freshBadge, marginLeft: 'auto' }}>Dati aggiornati</span>
      </div>

      <div style={{ ...card, marginBottom: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
          <div>
            <p style={{ fontSize: 15, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Catalogo fornitori</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '4px 0 0' }}>Organizza per macrosezione e sottocategoria; per ogni prodotto confronta il prezzo di listino con quello dei diversi fornitori.</p>
          </div>
          <button onClick={addMacro} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Macrosezione</button>
        </div>

        {catalog.map((m, mi) => {
          const mKey = `m${mi}`;
          return (
            <div key={mi} style={{ marginBottom: 14, border: `1px solid ${C.paleGray}`, borderRadius: 10, overflow: 'hidden' }}>
              <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', background: C.bg, cursor: 'pointer' }}>
                <span style={{ fontSize: 11, color: C.gray }}>{isOpen(mKey) ? '⌄' : '›'}</span>
                <span style={{ fontWeight: 700, fontSize: 13, color: C.black }}>{m.name}</span>
                <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
                  <button onClick={() => renameMacro(mi)} style={rowBtnStyle}>✎ Rinomina</button>
                  <button onClick={() => addCategoria(mi)} style={rowBtnStyle}>+ Sottocategoria</button>
                  <button onClick={() => removeMacro(mi)} style={rowBtnStyle}>🗑</button>
                </div>
              </div>

              {isOpen(mKey) && m.categorie.map((c, ci) => {
                const cKey = `m${mi}c${ci}`;
                return (
                  <div key={ci} style={{ paddingLeft: 20, borderTop: `1px solid ${C.paleGray}`, paddingTop: 10, paddingBottom: 10 }}>
                    <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', marginBottom: 8 }}>
                      <span style={{ fontSize: 11, color: C.gray }}>{isOpen(cKey) ? '⌄' : '›'}</span>
                      <span style={{ fontWeight: 600, fontSize: 12, color: C.black }}>{c.name}</span>
                      <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
                        <button onClick={() => renameCategoria(mi, ci)} style={rowBtnStyle}>✎</button>
                        <button onClick={() => addProdotto(mi, ci)} style={rowBtnStyle}>+ Prodotto</button>
                        <button onClick={() => removeCategoria(mi, ci)} style={rowBtnStyle}>🗑</button>
                      </div>
                    </div>
                    {isOpen(cKey) && (
                      c.prodotti.length === 0 ? (
                        <p style={{ fontSize: 12, color: C.gray, paddingLeft: 20 }}>Nessun prodotto ancora in questa sottocategoria.</p>
                      ) : (
                        <div style={{ paddingLeft: 20 }}>
                          {c.prodotti.map((p, pi) => (
                            <ProdottoCard
                              key={p.id}
                              prodotto={p}
                              onUpdate={(updated) => updateProdotto(mi, ci, pi, updated)}
                              onRemove={() => removeProdotto(mi, ci, pi)}
                              onAddToComputo={() => setAddingTo(p)}
                            />
                          ))}
                        </div>
                      )
                    )}
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {addingTo && (
        <AddToComputoModal
          prodotto={addingTo}
          projects={projects}
          onClose={() => setAddingTo(null)}
          onAdd={addToComputo}
        />
      )}
    </div>
  );
}

function TeamPage({ profile }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState('');
  const [name, setName] = useState('');
  const [role, setRole] = useState('Membro');
  const [error, setError] = useState('');
  const isAdmin = profile?.role === 'Admin';

  const loadMembers = async () => {
    setLoading(true);
    const { data, error: err } = await cea.from('team_members').select('*').order('created_at');
    if (!err) setMembers(data || []);
    setLoading(false);
  };

  React.useEffect(() => { loadMembers(); }, []);

  const handleInvite = async () => {
    const trimmed = email.trim();
    if (!trimmed) return;
    setError('');
    const { error: err } = await cea.from('team_members').insert({ name: name || trimmed.split('@')[0], email: trimmed, role, status: 'Invito inviato', invited_by: profile.auth_user_id });
    if (err) { setError(err.message); return; }
    setEmail(''); setName('');
    loadMembers();
  };

  const removeMember = async (id) => {
    if (!confirm('Rimuovere questa persona dal team? Perderà l\'accesso al workspace.')) return;
    const { error: err } = await cea.from('team_members').delete().eq('id', id);
    if (err) { alert(err.message); return; }
    loadMembers();
  };

  const changeRole = async (id, newRole) => {
    const { error: err } = await cea.from('team_members').update({ role: newRole }).eq('id', id);
    if (err) { alert(err.message); return; }
    loadMembers();
  };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 20 }}>
        <h1 style={h1Style}>Team</h1>
        <span style={freshBadge}>Workspace condiviso</span>
      </div>

      {!isAdmin && (
        <div style={{ ...card, marginBottom: 18, background: '#FFF8E1', border: '1px solid #F0D98C' }}>
          <p style={{ fontSize: 12, color: C.black, margin: 0 }}>Sei collegato come <strong>membro</strong>: solo l'admin può invitare o rimuovere persone dal team.</p>
        </div>
      )}

      {isAdmin && (
        <div style={{ ...card, marginBottom: 18 }}>
          <h2 style={{ fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Invita via email</h2>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome" style={{ flex: 1, minWidth: 120, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, background: C.bg }} />
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="nome@studio.it"
              style={{ flex: 1, minWidth: 180, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, background: C.bg }}
            />
            <select value={role} onChange={(e) => setRole(e.target.value)} style={{ fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
              <option>Membro</option>
              <option>Admin</option>
            </select>
            <button
              onClick={handleInvite}
              style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 16px', borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
            >
              Invia invito
            </button>
          </div>
          {error && <p style={{ fontSize: 12, color: C.maroon, margin: '10px 0 0' }}>{error}</p>}
          <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0' }}>
            La persona invitata deve andare sul sito, cliccare "Crea un account" e registrarsi con <strong>questa stessa email</strong>: verrà collegata automaticamente al team con il ruolo scelto qui.
          </p>
        </div>
      )}

      <div style={{ ...card, marginBottom: 18 }}>
        <h2 style={{ fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Persone con accesso</h2>
        {loading && <p style={{ fontSize: 12, color: C.gray }}>Caricamento…</p>}
        {!loading && members.length === 0 && <p style={{ fontSize: 12, color: C.gray }}>Nessuna persona ancora.</p>}
        {members.map((m) => (
          <div key={m.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '12px 0', borderBottom: `1px solid ${C.paleGray}`, gap: 10, flexWrap: 'wrap' }}>
            <div>
              <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{m.name}</p>
              <p style={{ fontSize: 12, color: C.gray, margin: '2px 0 0' }}>{m.email}</p>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {isAdmin ? (
                <select value={m.role} onChange={(e) => changeRole(m.id, e.target.value)} style={{ fontSize: 11, fontWeight: 600, padding: '5px 8px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}>
                  <option>Admin</option>
                  <option>Membro</option>
                </select>
              ) : (
                <span style={{ fontSize: 11, color: C.gray }}>{m.role}</span>
              )}
              <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 9px', borderRadius: 999, ...badgeStyles[m.status === 'Attivo' ? 'teal' : 'gray'] }}>
                {m.status}
              </span>
              {isAdmin && m.id !== profile.id && (
                <button onClick={() => removeMember(m.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
              )}
            </div>
          </div>
        ))}
      </div>

      <div style={{ background: 'rgba(128,20,48,0.06)', border: '1px solid rgba(128,20,48,0.18)', borderRadius: 10, padding: 14, fontSize: 12, color: C.midGray }}>
        <strong style={{ color: C.black }}>Come funziona la condivisione.</strong> Tutti i membri invitati accedono allo stesso workspace e agli stessi progetti, computi e listino con un vero account: i dati restano salvati per sempre e sono visibili a tutto il team.
      </div>
    </div>
  );
}

function LoginScreen({ onSignedIn }) {
  const [mode, setMode] = useState('signin'); // 'signin' | 'signup'
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSignIn = async () => {
    setError(''); setInfo(''); setLoading(true);
    const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setLoading(false);
    if (err) { setError(err.message === 'Invalid login credentials' ? 'Email o password errata.' : err.message); return; }
    onSignedIn();
  };

  const handleSignUp = async () => {
    setError(''); setInfo(''); setLoading(true);
    const { data, error: err } = await supabase.auth.signUp({ email: email.trim(), password });
    setLoading(false);
    if (err) { setError(err.message); return; }

    if (!data.session) {
      // Supabase non dice esplicitamente se l'email esiste già (per sicurezza), ma se non crea
      // nuove "identities" per un utente che ha già un account confermato, questo è il segnale.
      if (data.user && data.user.identities && data.user.identities.length === 0) {
        setError('Questa email ha già un account. Usa "Accedi" con la password che avevi scelto.');
        return;
      }
      setInfo('Controlla la tua casella email per confermare la registrazione, poi torna qui e accedi.');
      return;
    }

    onSignedIn();
  };

  return (
    <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT }}>
      <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360 }}>
        <div style={{ width: 40, height: 40, borderRadius: 999, background: C.black, color: C.white, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 14, marginBottom: 16 }}>SCE</div>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: C.black, margin: '0 0 4px' }}>{mode === 'signin' ? 'Accedi' : 'Crea il tuo account'}</h1>
        <p style={{ fontSize: 13, color: C.gray, margin: '0 0 24px' }}>Software di Computazione Edile — Desearq Studio</p>

        {mode === 'signup' && (
          <>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Nome</label>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Il tuo nome" style={{ width: '100%', fontSize: 13, padding: '10px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px' }} />
          </>
        )}

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Email</label>
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="nome@studio.it"
          style={{ width: '100%', fontSize: 13, padding: '10px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px' }}
        />
        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Password</label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && (mode === 'signin' ? handleSignIn() : handleSignUp())}
          style={{ width: '100%', fontSize: 13, padding: '10px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 8px' }}
        />
        {error && <p style={{ fontSize: 12, color: C.maroon, margin: '4px 0 8px' }}>{error}</p>}
        {info && <p style={{ fontSize: 12, color: C.success, margin: '4px 0 8px' }}>{info}</p>}

        <button
          disabled={loading}
          onClick={mode === 'signin' ? handleSignIn : handleSignUp}
          style={{ width: '100%', background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '11px 0', fontSize: 13, fontWeight: 600, cursor: loading ? 'default' : 'pointer', marginTop: 12, opacity: loading ? 0.7 : 1 }}
        >
          {loading ? 'Un attimo…' : mode === 'signin' ? 'Accedi' : 'Crea account'}
        </button>

        <p style={{ fontSize: 12, color: C.gray, margin: '18px 0 0', textAlign: 'center' }}>
          {mode === 'signin' ? (
            <>Prima volta? <span onClick={() => { setMode('signup'); setError(''); setInfo(''); }} style={{ color: C.maroon, fontWeight: 600, cursor: 'pointer' }}>Crea un account</span></>
          ) : (
            <>Hai già un account? <span onClick={() => { setMode('signin'); setError(''); setInfo(''); }} style={{ color: C.maroon, fontWeight: 600, cursor: 'pointer' }}>Accedi</span></>
          )}
        </p>
        <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0', lineHeight: 1.5 }}>
          Se sei il primo ad accedere diventi automaticamente admin. Chi arriva dopo deve essere prima invitato dall'admin dalla sezione Team, con la stessa email.
        </p>
      </div>
    </div>
  );
}

export default function GestionaleEdilePreview() {
  const [page, setPage] = useState('dashboard');
  const [projects, setProjects] = useState(PROJECTS);
  const [selectedProjectId, setSelectedProjectId] = useState(null);
  const [openRevisionId, setOpenRevisionId] = useState(null);
  const [listini, setListini] = useState([{ id: 1, name: 'Listino standard 2026', macros: INITIAL_MACROS }]);
  const [activeListinoId, setActiveListinoId] = useState(1);
  const [fornitoriCatalog, setFornitoriCatalog] = useState(INITIAL_FORNITORI);

  // --- Autenticazione reale (Supabase Auth) ---
  const [authUser, setAuthUser] = useState(undefined); // undefined = ancora in caricamento, null = non collegato
  const [profile, setProfile] = useState(null); // riga di cea.team_members collegata a questo utente
  const [profileError, setProfileError] = useState('');

  React.useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setAuthUser(data.session?.user ?? null));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setAuthUser(session?.user ?? null);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  React.useEffect(() => {
    if (!authUser) { setProfile(null); return; }
    let cancelled = false;
    setProfileError('');
    cea.from('team_members').select('*').eq('auth_user_id', authUser.id).maybeSingle().then(async ({ data, error }) => {
      if (cancelled) return;
      if (error) { setProfileError(error.message); return; }
      if (data) { setProfile(data); return; }
      // Nessun profilo collegato: prova a reclamare un invito esistente o a fare da primo admin
      // (copre anche gli account creati prima che questo sistema di team esistesse).
      const { data: claimed, error: claimErr } = await cea.rpc('bootstrap_or_claim_invite', { p_name: authUser.email.split('@')[0] });
      if (cancelled) return;
      if (claimErr) { setProfileError(claimErr.message); return; }
      setProfile(claimed);
    });
    return () => { cancelled = true; };
  }, [authUser]);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setProfile(null);
  };

  // --- Dati condivisi del workspace (progetti, listini, fornitori): un unico documento salvato su Supabase ---
  const [dataLoaded, setDataLoaded] = useState(false);
  const saveTimer = React.useRef(null);

  React.useEffect(() => {
    if (!profile || profile.status !== 'Attivo') return;
    cea.from('app_state').select('data').eq('id', 1).maybeSingle().then(({ data, error }) => {
      if (error) { console.error(error); setDataLoaded(true); return; }
      const saved = data?.data || {};
      if (saved.projects) setProjects(saved.projects);
      if (saved.listini) setListini(saved.listini);
      if (saved.activeListinoId) setActiveListinoId(saved.activeListinoId);
      if (saved.fornitoriCatalog) setFornitoriCatalog(saved.fornitoriCatalog);
      setDataLoaded(true);
    });
  }, [profile]);

  React.useEffect(() => {
    if (!dataLoaded || !profile || profile.status !== 'Attivo') return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      cea.from('app_state').update({
        data: { projects, listini, activeListinoId, fornitoriCatalog },
        updated_at: new Date().toISOString(),
        updated_by: authUser?.id,
      }).eq('id', 1).then(({ error }) => { if (error) console.error('Salvataggio fallito:', error.message); });
    }, 900);
    return () => clearTimeout(saveTimer.current);
  }, [projects, listini, activeListinoId, fornitoriCatalog, dataLoaded]);

  const openProject = (id) => { setSelectedProjectId(id); setOpenRevisionId(null); setPage('progetto-dettaglio'); };
  const openRevisionInProject = (projectId, revisionId) => { setSelectedProjectId(projectId); setOpenRevisionId(revisionId); setPage('progetto-dettaglio'); };
  const selectedProject = projects.find((p) => p.id === selectedProjectId);

  const [printJob, setPrintJob] = useState(null); // { project, revision, clientOnly }
  const requestPdf = (project, revision, clientOnly) => setPrintJob({ project, revision, clientOnly });

  React.useEffect(() => {
    if (!printJob) return;
    const t = setTimeout(() => window.print(), 150);
    const handleAfterPrint = () => setPrintJob(null);
    window.addEventListener('afterprint', handleAfterPrint);
    return () => { clearTimeout(t); window.removeEventListener('afterprint', handleAfterPrint); };
  }, [printJob]);

  if (authUser === undefined) {
    return <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT, color: C.gray, fontSize: 13 }}>Caricamento…</div>;
  }

  if (!authUser) {
    return <LoginScreen onSignedIn={() => {}} />;
  }

  if (!profile) {
    return (
      <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT }}>
        <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360, textAlign: 'center' }}>
          <p style={{ fontSize: 13, color: C.black, margin: '0 0 12px' }}>{profileError || 'Nessun invito trovato per questa email. Chiedi a un admin di invitarti dalla sezione Team, poi ricarica la pagina.'}</p>
          <button onClick={handleLogout} style={{ background: C.darkGray, color: C.white, border: 'none', borderRadius: 999, padding: '9px 18px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Esci</button>
        </div>
      </div>
    );
  }

  if (profile.status !== 'Attivo') {
    return (
      <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT }}>
        <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360, textAlign: 'center' }}>
          <p style={{ fontSize: 13, color: C.black, margin: 0 }}>Il tuo invito non è stato ancora completato. Riprova tra poco o contatta l'admin.</p>
        </div>
      </div>
    );
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
        .print-only { display: none; }
        @media print {
          .no-print { display: none !important; }
          .print-only { display: block !important; }
        }
        .sidebar-item-btn:not(.active):hover { background: ${C.sidebarHover} !important; color: ${C.black} !important; }
        .btn-accent-pill:hover { background: #650F26 !important; }
        .input-focus:focus { outline: none; border-color: ${C.maroon} !important; box-shadow: 0 0 0 3px rgba(128,20,48,0.25); }
      `}</style>
      {printJob && (
        <div className="print-only">
          <PrintableComputo project={printJob.project} revision={printJob.revision} clientOnly={printJob.clientOnly} />
        </div>
      )}
      <div className="no-print" style={{ display: 'flex', minHeight: '100vh', fontFamily: FONT, background: PAGE_GRADIENT, backgroundAttachment: 'fixed' }}>
      <aside style={{ width: 280, flexShrink: 0, background: C.sidebar, color: C.darkGray, borderRight: `1px solid ${C.paleGray}`, display: 'flex', flexDirection: 'column', padding: '22px 14px' }}>
        <div style={{ padding: '0 6px 20px', borderBottom: `1px solid ${C.paleGray}`, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 36, height: 36, borderRadius: 999, background: C.black, color: C.white, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 14, flexShrink: 0 }}>SCE</div>
          <div>
            <p style={{ fontFamily: FONT, fontWeight: 700, fontSize: 15, color: C.black, margin: 0 }}>Software di Computazione Edile</p>
            <p style={{ fontSize: 11, margin: '2px 0 0', lineHeight: 1.4, color: C.gray }}>Desearq Studio</p>
          </div>
        </div>
        <p style={{ fontSize: 10, fontWeight: 700, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.gray, padding: '0 6px', marginBottom: 8 }}>Area di lavoro</p>
        <nav style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          {NAV_ITEMS.map((item) => {
            const active = page === item.key;
            const Icon = item.Icon;
            return (
              <button
                key={item.key}
                className={`sidebar-item-btn${active ? ' active' : ''}`}
                onClick={() => setPage(item.key)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 10,
                  padding: '11px 12px',
                  borderRadius: 999,
                  fontSize: 15,
                  fontWeight: 500,
                  textAlign: 'left',
                  border: 'none',
                  cursor: 'pointer',
                  fontFamily: FONT,
                  background: active ? C.black : 'transparent',
                  color: active ? C.white : C.black,
                }}
              >
                <Icon size={18} strokeWidth={1.5} />
                {item.label}
              </button>
            );
          })}
        </nav>
        <div style={{ marginTop: 'auto', paddingTop: 14, borderTop: `1px solid ${C.paleGray}`, display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 36, height: 36, borderRadius: 999, background: '#8B6F5C', color: C.white, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 13, flexShrink: 0 }}>{profile.email[0].toUpperCase()}</div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <p style={{ fontSize: 12, fontWeight: 600, color: C.black, margin: 0, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{profile.email}</p>
            <p style={{ fontSize: 10, color: C.gray, margin: '2px 0 0' }}>{profile.role === 'Admin' ? 'Admin' : 'Membro del team'}</p>
          </div>
          <button onClick={handleLogout} style={{ width: 32, height: 32, borderRadius: 999, border: `1px solid ${C.paleGray}`, background: 'transparent', color: C.darkGray, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', flexShrink: 0 }}>
            <LogOut size={16} strokeWidth={1.5} />
          </button>
        </div>
      </aside>

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
        <header style={{ background: 'transparent', padding: '18px 24px 0', display: 'flex', alignItems: 'center', justifyContent: 'flex-end' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 18px', boxShadow: '0 1px 2px rgba(0,0,0,0.03), 0 4px 12px rgba(0,0,0,0.04)' }}>
            <Search size={16} strokeWidth={1.5} color={C.darkGray} />
            <input
              type="text"
              placeholder="Cerca progetto, cliente, voce o codice…"
              style={{ border: 'none', outline: 'none', background: 'transparent', fontSize: 13, color: C.darkGray, width: 260 }}
            />
          </div>
        </header>

        <main style={{ padding: 24, flex: 1, overflowX: 'auto' }}>
          {page === 'dashboard' && <Dashboard onNavigate={setPage} onOpenProject={openProject} projects={projects} />}
          {page === 'listino' && <ListinoPage listini={listini} setListini={setListini} activeId={activeListinoId} setActiveId={setActiveListinoId} />}
          {page === 'progetti' && <ProgettiPage projects={projects} setProjects={setProjects} onOpenProject={openProject} />}
          {page === 'progetto-dettaglio' && selectedProject && (
            <ProjectDetailPage
              project={selectedProject}
              onBack={() => setPage('progetti')}
              onUpdateProject={(updated) => setProjects(projects.map((p) => (p.id === updated.id ? updated : p)))}
              listini={listini}
              initialRevisionId={openRevisionId}
              requestPdf={requestPdf}
            />
          )}
          {page === 'computi' && <ComputiPage projects={projects} setProjects={setProjects} onOpenProject={openProject} onOpenRevision={openRevisionInProject} requestPdf={requestPdf} />}
          {page === 'confronto' && <ConfrontoPage projects={projects} />}
          {page === 'fornitori' && <FornitoriPage projects={projects} setProjects={setProjects} catalog={fornitoriCatalog} setCatalog={setFornitoriCatalog} />}
          {page === 'team' && <TeamPage profile={profile} />}
        </main>
      </div>
      </div>
    </>
  );
}
