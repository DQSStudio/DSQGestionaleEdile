import React, { useState, useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';
import { createClient } from '@supabase/supabase-js';
import { supabase, cea, SUITE_URL, SUITE_KEY, OLD_URL, OLD_KEY } from './supabaseClient';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { LayoutGrid, BookOpen, Building2, Calculator, GitCompare, Truck, Users, Search, LogOut, MapPin, Clock, Calendar as CalendarIcon, Settings, UserSearch } from 'lucide-react';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

const FONT = "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif";

// Design token Desearq Flow (bordeaux #862D2D, cream #F4EFE8, font Inter)
const DEFAULT_ACCENT = '#862D2D';
const C = {
  // Colore di accento personalizzabile in Impostazioni studio (variabili CSS impostate in GestionaleEdilePreview);
  // il valore dopo la virgola è il bordeaux Desearq, usato finché le impostazioni non sono caricate.
  maroon: 'var(--accent, #862D2D)',
  accentHover: 'var(--accent-hover, #6E2424)',
  accentSoft: 'var(--accent-soft, #F1DCD3)',
  black: '#211A16',
  sidebar: '#F4EFE8',
  sidebarHover: '#EAE1D2',
  darkGray: '#6F655F',
  midGray: '#6F655F',
  gray: '#8B8178',
  lightGray: '#C9BEAF',
  paleGray: '#E9E0D4',
  borderStrong: '#E3D9CC',
  bg: '#F4EFE8',
  surfaceSubtle: '#FAF8F4',
  white: '#FFFDFA',
  success: '#2E7D4F',
};
const PAGE_GRADIENT = 'radial-gradient(120% 90% at 85% 0%, #FBF8F2 0%, #F4EFE8 45%, #F0E8D8 100%)';

const NAV_ITEMS = [
  { key: 'dashboard', label: 'Dashboard', Icon: LayoutGrid },
  { key: 'listino', label: 'Listino prezzi', Icon: BookOpen },
  { key: 'progetti', label: 'Progetti', Icon: Building2 },
  { key: 'computi', label: 'Computi', Icon: Calculator },
  { key: 'confronto', label: 'Confronto revisioni', Icon: GitCompare },
  { key: 'rilievo', label: 'Rilievo planimetrico', Icon: MapPin },
  { key: 'fornitori', label: 'Fornitori', Icon: Truck },
  { key: 'team', label: 'Team', Icon: Users },
  { key: 'impostazioni', label: 'Impostazioni studio', Icon: Settings },
];

// Icona di navigazione disegnata a mano (stessi tracciati dell'artefatto Flow), non un'icona lucide-react:
// qui serve riprodurre esattamente il disegno del mockup, non un'icona "equivalente".
function NavIcon({ children }) {
  return (
    <svg viewBox="0 0 24 24" width={20} height={20} fill="none" stroke="currentColor" strokeWidth={1.5} strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
      {children}
    </svg>
  );
}

// Struttura della sidebar raggruppata per area (Listino / Progetti / Computi / Forniture / Team), identica
// all'artefatto "Quant — Nuova grafica (Flow)": un gruppo senza etichetta per la Dashboard, poi un'etichetta
// di sezione prima di ciascun gruppo di voci.
const NAV_GROUPS = [
  { label: null, items: [
    { key: 'dashboard', label: 'Dashboard', icon: <NavIcon><rect x="3.5" y="3.5" width="7" height="9" rx="1.5" /><rect x="13.5" y="3.5" width="7" height="5" rx="1.5" /><rect x="13.5" y="11.5" width="7" height="9" rx="1.5" /><rect x="3.5" y="15.5" width="7" height="5" rx="1.5" /></NavIcon> },
  ]},
  { label: 'Listino', items: [
    { key: 'listino', label: 'Listino prezzi', icon: <NavIcon><path d="M4 4.5h7a2 2 0 0 1 2 2V20a2 2 0 0 0-2-1.5H4z" /><path d="M20 4.5h-7a2 2 0 0 0-2 2V20a2 2 0 0 1 2-1.5h7z" /></NavIcon> },
    { key: 'fornitori', label: 'Fornitori', icon: <NavIcon><path d="M12 2.6l8.5 4.8v9.2L12 21.4l-8.5-4.8V7.4z M3.5 7.4L12 12.2l8.5-4.8 M12 12.2v9.2" /></NavIcon> },
  ]},
  { label: 'Progetti', items: [
    { key: 'progetti', label: 'Progetti', icon: <NavIcon><rect x="2.5" y="7.5" width="19" height="12.5" rx="1.8" /><path d="M8.3 7.5V6a2 2 0 0 1 2-2h3.4a2 2 0 0 1 2 2v1.5" /></NavIcon>, countKey: 'progetti' },
    { key: 'rilievo', label: 'Rilievo planimetrico', icon: <NavIcon><path d="M3 6l6-2.3 6 2.3 6-2.3v15l-6 2.3-6-2.3-6 2.3z" /><path d="M9 3.7v15 M15 6v15" /></NavIcon> },
  ]},
  { label: 'Computi', items: [
    { key: 'computi', label: 'Computi', icon: <NavIcon><rect x="5" y="2.5" width="14" height="19" rx="2" /><path d="M8 7h8 M8 12h2 M8 16h2 M14 12h2 M14 16h2" /></NavIcon>, countKey: 'computi' },
    { key: 'confronto', label: 'Confronto revisioni', icon: <NavIcon><path d="M8 3v14 M3 12l5 5 5-5" /><path d="M16 21V7 M21 12l-5-5-5 5" /></NavIcon> },
  ]},
  { label: 'Team', items: [
    { key: 'team', label: 'Team', icon: <NavIcon><circle cx="8.5" cy="8.3" r="3.2" /><path d="M2.5 19.5c.4-3.8 3-6.5 6-6.5s5.6 2.7 6 6.5" /><circle cx="17" cy="9" r="2.3" /><path d="M15.6 13c2.2.4 4 2.6 4.4 6" /></NavIcon> },
    { key: 'impostazioni', label: 'Impostazioni studio', icon: <NavIcon><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.7 1.7 0 0 0 .34 1.87l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.7 1.7 0 0 0-1.87-.34 1.7 1.7 0 0 0-1.03 1.56V21a2 2 0 1 1-4 0v-.09A1.7 1.7 0 0 0 9 19.35a1.7 1.7 0 0 0-1.87.34l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.7 1.7 0 0 0 4.65 15a1.7 1.7 0 0 0-1.56-1.03H3a2 2 0 1 1 0-4h.09A1.7 1.7 0 0 0 4.65 9a1.7 1.7 0 0 0-.34-1.87l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.7 1.7 0 0 0 9 4.65a1.7 1.7 0 0 0 1.03-1.56V3a2 2 0 1 1 4 0v.09a1.7 1.7 0 0 0 1.03 1.56 1.7 1.7 0 0 0 1.87-.34l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.7 1.7 0 0 0-.34 1.87V9a1.7 1.7 0 0 0 1.56 1.03H21a2 2 0 1 1 0 4h-.09a1.7 1.7 0 0 0-1.56 1.03z" /></NavIcon> },
  ]},
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
// Numero in formato italiano con il punto delle migliaia SEMPRE presente: toLocaleString('it-IT') per
// convenzione non separa i numeri di 4 cifre (dava "1688,52" invece di "1.688,52").
const itNum = (n, decimals = 2) => {
  const v = Number(n) || 0;
  const [int, dec] = Math.abs(v).toFixed(decimals).split('.');
  return `${v < 0 && Number(Math.abs(v).toFixed(decimals)) !== 0 ? '-' : ''}${int.replace(/\B(?=(\d{3})+(?!\d))/g, '.')}${decimals ? `,${dec}` : ''}`;
};
const formatEuro = (n) => `${itNum(n, 2)} €`;

// Nome visualizzato di un prodotto fornitori: marca + modello (marca omessa se vuota).
const prodottoLabel = (p) => [p.marca, p.modello].filter(Boolean).join(' ') || '(prodotto senza nome)';

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
// Trova una voce in un albero macro→categorie→sottocategorie→voci cercando per "code" (l'unico identificativo
// stabile di una voce: la sua posizione può cambiare se qualcuno riordina o modifica il listino nel frattempo).
// Usata per riportare nel listino vero e proprio un costo impresa approvato da un fornitore.
function findVoceByCode(macros, code) {
  for (let mi = 0; mi < (macros || []).length; mi++) {
    const categorie = macros[mi].categorie || [];
    for (let ci = 0; ci < categorie.length; ci++) {
      const sottocategorie = categorie[ci].sottocategorie || [];
      for (let si = 0; si < sottocategorie.length; si++) {
        const voci = sottocategorie[si].voci || [];
        for (let vi = 0; vi < voci.length; vi++) {
          if (voci[vi].code === code) return [mi, ci, si, vi];
        }
      }
    }
  }
  return null;
}

const nowLabel = () => new Date().toLocaleString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const sumImpresa = (items) => (items || []).reduce((sum, it) => sum + parseEuro(it.unitPriceImpresa) * parseEuro(it.qty), 0);
const sumCliente = (items) => (items || []).reduce((sum, it) => sum + parseEuro(it.unitPriceCliente) * parseEuro(it.qty), 0);

// --- Date in formato ISO (yyyy-mm-dd) per il cronoprogramma: comode da salvare e da usare con <input type="date">.
const todayISO = () => new Date().toISOString().slice(0, 10);
const addDaysISO = (iso, days) => { const d = new Date(`${iso}T00:00:00`); d.setDate(d.getDate() + days); return d.toISOString().slice(0, 10); };
const formatDateIT = (iso) => { if (!iso) return ''; const d = new Date(`${iso}T00:00:00`); return d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short' }); };
const daysBetweenISO = (a, b) => Math.round((new Date(`${b}T00:00:00`) - new Date(`${a}T00:00:00`)) / 86400000);

// --- Voci del computo con misurazioni reali (par.ug × lung × larg × H/peso, come in un computo metrico
// estimativo tradizionale). Una voce può avere più "gruppi" di misurazione (uno per ogni misurazione
// separata, col segno per le detrazioni): la quantità della voce è la somma di tutti i gruppi, ed è quella
// che finisce nel prezzo — la si ricalcola e la si salva in item.qty ogni volta che le misurazioni cambiano,
// così tutto il resto dell'app (prezzi, stampe, sincronizzazione col Portale Clienti) continua a leggere
// semplicemente item.qty come sempre, senza bisogno di sapere nulla delle misurazioni sottostanti.
function computeMisurazioneRowValue(row, unitaCalcolo) {
  const parUg = row.parUg !== '' && row.parUg !== undefined && row.parUg !== null ? parseEuro(row.parUg) : 1;
  const lung = parseEuro(row.lung);
  const larg = parseEuro(row.larg);
  const hPeso = parseEuro(row.hPeso);
  let v;
  if (unitaCalcolo === 'ml') v = parUg * lung;
  else if (unitaCalcolo === 'mq') v = parUg * lung * larg;
  else if (unitaCalcolo === 'm3') v = parUg * lung * larg * hPeso;
  else return 0;
  return row.segno === '-' ? -v : v;
}
function computeGruppoTotal(gruppo, unitaCalcolo) {
  return (gruppo?.rows || []).reduce((sum, r) => sum + computeMisurazioneRowValue(r, unitaCalcolo), 0);
}
function computeVoceQtyTotal(misurazioni, unitaCalcolo) {
  return (misurazioni || []).reduce((sum, g) => sum + computeGruppoTotal(g, unitaCalcolo), 0);
}

// Abbreviazione automatica di un nome (macrocategoria/sottocategoria) per generare il codice di una voce:
// un'unica parola significativa -> le sue prime 3 lettere; più parole -> le iniziali delle prime 3.
const STOPWORDS_CODICE = new Set(['e', 'di', 'del', 'della', 'dei', 'delle', 'dello', 'per', 'con', 'il', 'la', 'lo', 'i', 'le', 'gli', 'un', 'una', 'ed']);
function abbreviaNome(nome) {
  const parole = (nome || '').split(/\s+/).map((w) => w.replace(/[^a-zA-ZÀ-ÿ0-9]/g, '')).filter((w) => w && !STOPWORDS_CODICE.has(w.toLowerCase()));
  if (parole.length === 0) return '---';
  if (parole.length === 1) return parole[0].slice(0, 3).toUpperCase();
  return parole.slice(0, 3).map((w) => w[0].toUpperCase()).join('');
}

// Rigenera il codice di ogni voce creata col nuovo sistema (item.autoCode) in base alla sua posizione:
// {ABBREV. MACROCATEGORIA}.{ABBREV. SOTTOCATEGORIA}.{progressivo × 10}. Il progressivo segue l'ordine con
// cui le voci compaiono nell'array (lo stesso che "sposta su/giù" modifica) all'interno della stessa coppia
// macrocategoria/sottocategoria. Le voci "storiche" (listino, importazioni) mantengono il loro codice originale.
function regenerateItemCodes(items) {
  const counters = {};
  return items.map((it) => {
    if (it.type === 'subtotal' || !it.autoCode) return it;
    const m = it.macro || it.section || 'Voci varie';
    const s = it.sottocategoria || 'Generale';
    const key = m + '␟' + s;
    counters[key] = (counters[key] || 0) + 1;
    return { ...it, code: `${abbreviaNome(m)}.${abbreviaNome(s)}.${String(counters[key] * 10).padStart(3, '0')}` };
  });
}

// Ordina un elenco di nomi (macrocategorie o sottocategorie) secondo un ordine salvato in precedenza,
// aggiungendo in fondo i nomi nuovi non ancora presenti in quell'ordine (es. appena creati).
function orderNames(existingNames, savedOrder) {
  const order = (savedOrder || []).filter((n) => existingNames.includes(n));
  existingNames.forEach((n) => { if (!order.includes(n)) order.push(n); });
  return order;
}

// Raggruppa le voci di una revisione per sezione/categoria (stesso criterio usato nel computo:
// it.section || it.macro) e calcola il totale cliente netto di ogni categoria (dopo lo sconto di sezione).
// Usata per gli Stati avanzamento pagamenti, che seguono l'ultima revisione approvata del computo.
function computeSectionTotals(revision) {
  const items = (revision?.items || []).filter((it) => it.type !== 'subtotal');
  const extraSections = revision?.extraSections || [];
  const sectionDiscounts = revision?.sectionDiscounts || {};
  const sections = [];
  const byName = (name) => {
    let s = sections.find((s) => s.name === name);
    if (!s) { s = { name, items: [] }; sections.push(s); }
    return s;
  };
  extraSections.forEach((name) => byName(name));
  items.forEach((it) => byName(it.section || it.macro || 'Voci varie').items.push(it));
  sections.forEach((s) => {
    const subtotalCliente = sumCliente(s.items);
    const discountPct = parseFloat(sectionDiscounts[s.name]) || 0;
    s.netCliente = subtotalCliente * (1 - discountPct / 100);
  });
  return sections.filter((s) => s.items.length > 0);
}

// L'ultima revisione del computo che è stata approvata (status diverso da "In attesa di approvazione"),
// scorrendo dalla più recente: è quella che alimenta gli Stati avanzamento pagamenti e il Portale Clienti.
function latestApprovedRevision(project) {
  const revs = project.revisions || [];
  for (let i = revs.length - 1; i >= 0; i--) {
    if (revs[i].status && revs[i].status !== STATUS_OPTIONS[0]) return revs[i];
  }
  return null;
}
const getVatInfo = (revision) => {
  const rate = revision?.vatRate !== undefined && revision?.vatRate !== null ? Number(revision.vatRate) : 22;
  const label = revision?.vatLabel && revision.vatLabel.trim() ? revision.vatLabel : `IVA ${rate}%`;
  return { rate, label };
};

// L'IVA si imposta sezione per sezione (revision.sectionVat[nomeSezione] = { rate, label }), non più sul
// totale generale: il totale generale somma l'IVA delle singole sezioni, divisa per aliquota/dicitura.
// Una sezione senza impostazione propria usa l'aliquota che il computo aveva prima (vecchia IVA unica,
// di default 22%), così i computi già esistenti non cambiano importi.
const VAT_PRESETS = [
  { rate: 22, label: null, text: 'Standard 22%' },
  { rate: 10, label: null, text: 'Agevolata 10%' },
  { rate: 4, label: null, text: 'Agevolata 4%' },
  { rate: 0, label: 'Esente IVA', text: 'Esente IVA' },
  { rate: 0, label: 'IVA non dovuta', text: 'IVA non dovuta' },
  { rate: 16, label: 'IVA dovuta 50% al 22% e 50% al 10%', text: 'Mista 50%/50% (22%+10%)' },
];
function getSectionVat(revision, sectionName) {
  const sv = revision?.sectionVat?.[sectionName];
  if (sv && sv.rate !== undefined && sv.rate !== null && sv.rate !== '') {
    const rate = Number(sv.rate) || 0;
    return { rate, label: sv.label && String(sv.label).trim() ? sv.label : `IVA ${rate}%` };
  }
  const base = getVatInfo(revision);
  return { rate: base.rate, label: sv?.label && String(sv.label).trim() ? sv.label : base.label };
}
function computeEconomics(revision) {
  const discounts = revision?.sectionDiscounts || {};
  const sections = buildComputoGroups(revision).map((g) => {
    const real = g.items.filter((it) => it.type !== 'subtotal');
    const grossImpresa = sumImpresa(real);
    const grossCliente = sumCliente(real);
    const discountPct = parseFloat(discounts[g.name]) || 0;
    const netImpresa = grossImpresa * (1 - discountPct / 100);
    const netCliente = grossCliente * (1 - discountPct / 100);
    const vat = getSectionVat(revision, g.name);
    const ivaImpresa = netImpresa * (vat.rate / 100);
    const ivaCliente = netCliente * (vat.rate / 100);
    return { name: g.name, grossImpresa, grossCliente, discountPct, netImpresa, netCliente, vat, ivaImpresa, ivaCliente, totImpresa: netImpresa + ivaImpresa, totCliente: netCliente + ivaCliente };
  });
  const breakdown = [];
  sections.forEach((s) => {
    if (!s.netImpresa && !s.netCliente) return;
    let b = breakdown.find((x) => x.label === s.vat.label && x.rate === s.vat.rate);
    if (!b) { b = { label: s.vat.label, rate: s.vat.rate, ivaImpresa: 0, ivaCliente: 0 }; breakdown.push(b); }
    b.ivaImpresa += s.ivaImpresa;
    b.ivaCliente += s.ivaCliente;
  });
  const sum = (k) => sections.reduce((n, s) => n + s[k], 0);
  return {
    sections, breakdown,
    netImpresa: sum('netImpresa'), netCliente: sum('netCliente'),
    ivaImpresa: sum('ivaImpresa'), ivaCliente: sum('ivaCliente'),
    totImpresa: sum('totImpresa'), totCliente: sum('totCliente'),
  };
}

// Aggiunge una voce (es. una fornitura) alla revisione scelta di un progetto.
// Se la revisione scelta è l'ultima, aggiorna sul posto; altrimenti crea automaticamente
// una nuova versione (fork) lasciando quella aperta intatta, come per ogni altra modifica.
function addItemToProjectRevision(project, revisionId, newItem) {
  const revisions = project.revisions;
  const latest = revisions[revisions.length - 1];
  // revisionId può essere direttamente l'oggetto revisione (scelto dall'elenco per posizione): così non si
  // dipende dagli id, che nei dati reali possono ripetersi.
  const target = (revisionId && typeof revisionId === 'object' && revisions.includes(revisionId) ? revisionId : revisions.find((r) => String(r.id) === String(revisionId))) || latest;
  const isLatest = target === latest;
  const newItems = [...(target.items || []), newItem];
  const realItems = newItems.filter((it) => it.type !== 'subtotal');
  const total = formatEuro(sumImpresa(realItems));
  const totalCliente = formatEuro(sumCliente(realItems));
  if (isLatest) {
    return {
      ...project,
      revisions: revisions.map((r) => (r === target ? { ...r, items: newItems, dateModified: nowLabel(), total, totalCliente } : r)),
      value: total,
    };
  }
  const newRev = { ...target, id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: newItems, total, totalCliente };
  return { ...project, revisions: [...revisions, newRev], value: total };
}

function computeItemsDiff(itemsA, itemsB) {
  // Le voci col codice auto-generato (item.autoCode) cambiano codice quando si spostano di posizione:
  // per loro il confronto usa l'id (stabile), non il codice, altrimenti ogni spostamento sembrerebbe
  // una rimozione+aggiunta invece di una modifica. Le voci storiche restano confrontate per codice.
  const keyOf = (it) => (it.autoCode ? 'id:' + it.id : 'code:' + it.code);
  const mapA = Object.fromEntries((itemsA || []).map((it) => [keyOf(it), it]));
  const mapB = Object.fromEntries((itemsB || []).map((it) => [keyOf(it), it]));
  const keys = Array.from(new Set([...Object.keys(mapA), ...Object.keys(mapB)]));
  return keys.map((key) => {
    const a = mapA[key];
    const b = mapB[key];
    const ref = b || a;
    const code = ref.code;
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
      out.push({ ...v, listinoId: listino.id, macro: m.name, categoria: c.name, sottocategoria: s.name, impresaValue: impresa, clienteValue: cliente });
    });
  })));
  return out;
}

// --- Importazione computo da PDF (solo macrocategorie + totale complessivo) ---
// Estrae il testo del PDF pagina per pagina (pdfjs ricostruisce le righe usando "hasEOL"),
// poi cerca su ogni riga l'ultimo importo in formato italiano (es. "12.604,00"): quello che precede
// diventa l'etichetta. Le righe la cui etichetta contiene "totale" insieme a "complessivo"/"generale"/
// "lavori"/"opera"/"computo" sono candidate come totale complessivo; le altre come macrocategorie.
// È un'estrazione "best effort": l'utente rivede e corregge le righe proposte prima di confermare.
const EURO_AMOUNT_RE = /(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)/g;

function parseItalianAmount(str) {
  const n = parseFloat(String(str).replace(/\./g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

async function extractTextLinesFromPdf(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const lines = [];
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    let current = '';
    content.items.forEach((item) => {
      current += item.str;
      if (item.hasEOL) {
        if (current.trim()) lines.push(current.trim());
        current = '';
      } else if (item.str) {
        current += ' ';
      }
    });
    if (current.trim()) lines.push(current.trim());
  }
  return lines;
}

function parseComputoPdfLines(lines) {
  const candidateRows = [];
  let totaleComplessivo = null;
  let bestTotaleAmount = -Infinity;

  lines.forEach((line) => {
    const matches = [...line.matchAll(EURO_AMOUNT_RE)];
    if (matches.length === 0) return;
    const last = matches[matches.length - 1];
    const amount = parseItalianAmount(last[1]);
    if (!(amount > 0)) return;
    let label = line.slice(0, last.index).replace(/[.\-–_\s€]+$/, '').trim();
    if (!label) return;
    const labelLower = label.toLowerCase();
    const looksLikeTotale = labelLower.includes('totale') &&
      /(complessivo|generale|lavori|opera|computo|importo)/.test(labelLower);
    if (looksLikeTotale) {
      if (amount > bestTotaleAmount) { bestTotaleAmount = amount; totaleComplessivo = amount; }
      return;
    }
    if (label.length > 90) return; // riga troppo lunga: probabilmente una voce di dettaglio, non un totale di categoria
    candidateRows.push({ id: Date.now() + Math.random(), name: label, totale: formatEuro(amount).replace(' €', '') });
  });

  return { rows: candidateRows, totaleComplessivo };
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
  // "In attesa di approvazione": bordeaux pieno (era giallo), distinto da Approvato (bordeaux chiaro) e In cantiere (grigio)
  orange: { background: C.maroon, color: '#FFFFFF' },
  teal: { background: `color-mix(in srgb, ${C.maroon} 18%, #fff)`, color: C.maroon },
  gray: { background: '#E8E0D6', color: '#5A4F4A' },
};

const card = { background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 20, padding: 18, boxShadow: '0 1px 2px rgba(0,0,0,0.03), 0 4px 12px rgba(0,0,0,0.04)' };
const h1Style = { fontFamily: FONT, fontSize: 30, margin: 0, color: C.black, fontWeight: 300, letterSpacing: '-0.01em' };
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

const IT_MONTHS = { gen: 0, feb: 1, mar: 2, apr: 3, mag: 4, giu: 5, lug: 6, ago: 7, set: 8, ott: 9, nov: 10, dic: 11 };
// Le date dei progetti/revisioni si trovano in due formati testuali diversi: quello dei progetti
// d'esempio ("22 lug 2026, 16:45", mese scritto per esteso) e quello che genera davvero "nowLabel()"
// ogni volta che si salva una modifica nell'app reale ("14/10/2026, 16:45", giorno/mese/anno numerico).
// Senza riconoscere anche il secondo formato, ogni data creata realmente dall'app risultava "non
// interpretabile" e spariva dai conteggi e dalla sezione "Modifiche recenti" della Dashboard.
function parseItDate(str) {
  if (!str) return null;
  const s = String(str);
  const named = s.match(/(\d{1,2})\s+([a-zà]{3,})\.?\s+(\d{4})(?:,\s*(\d{1,2}):(\d{2}))?/i);
  if (named) {
    const mon = IT_MONTHS[named[2].toLowerCase().slice(0, 3)];
    if (mon !== undefined) return new Date(Number(named[3]), mon, Number(named[1]), Number(named[4] || 0), Number(named[5] || 0));
  }
  const numeric = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,?\s*(\d{1,2}):(\d{2}))?/);
  if (numeric) {
    return new Date(Number(numeric[3]), Number(numeric[2]) - 1, Number(numeric[1]), Number(numeric[4] || 0), Number(numeric[5] || 0));
  }
  return null;
}
function daysBetween(a, b) { return Math.max(0, Math.round((b - a) / 86400000)); }

function KpiInfo() {
  return <span style={{ width: 15, height: 15, borderRadius: '50%', border: '1px solid #CFC4B6', fontSize: 9, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', color: '#8B8178', flexShrink: 0 }}>i</span>;
}

function DKpi({ label, value, note, accent }) {
  return (
    <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 6 }}>
      <div style={{ fontSize: 12.5, color: C.darkGray, display: 'flex', alignItems: 'center', gap: 8, letterSpacing: '.02em' }}>{label}<KpiInfo /></div>
      <div style={{ fontSize: 30, fontWeight: 500, letterSpacing: '-.03em', lineHeight: 1.1, color: accent ? C.maroon : C.black }}>{value}</div>
      {note && <div style={{ fontSize: 11.5, color: C.darkGray }}>{note}</div>}
    </div>
  );
}

// --- Dashboard: copia 1:1 della board "Main" dell'artefatto "Quant — Nuova grafica (Flow)" ---
// Stessa sequenza di blocchi, stessi stili (card #FFFFFF, bordo #E8E0D4, raggio 14, titoli 20/23px peso 300),
// riempita con i dati reali dell'app. Dove l'artefatto mostrava numeri d'esempio non tracciabili, il blocco
// resta identico ma mostra solo dati veri (es. cronoprogramma e scadenze: usano "Inizio lavori" / "Fine lavori"
// inseriti nei Dati generali del computo di ogni progetto).
const DASH = {
  card: { background: '#FFFFFF', border: '1px solid #E8E0D4', borderRadius: 14, boxSizing: 'border-box' },
  h3: { margin: 0, fontWeight: 300, letterSpacing: '-.01em', display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', color: '#2B2321' },
  tag2: { fontSize: 11.5, fontWeight: 400, color: '#8B8178', letterSpacing: 0 },
  th: { fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: 'left', padding: '8px 10px', borderBottom: '1px solid #E3D9CC' },
  td: { fontSize: 13.5, padding: '12px 10px', borderBottom: '1px solid #EFE8DD' },
  mono: { fontVariantNumeric: 'tabular-nums', color: '#6F655F', fontSize: 12.5 },
  pill: { fontSize: 11, padding: '3px 10px', borderRadius: 999, fontWeight: 500, whiteSpace: 'nowrap' },
  act: { display: 'flex', gap: 12, padding: '11px 0', borderBottom: '1px solid #EFE8DD' },
  actIc: { width: 28, height: 28, borderRadius: 8, flex: 'none', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 13 },
  li: { display: 'flex', alignItems: 'center', gap: 12, padding: '11px 0', borderBottom: '1px solid #EFE8DD', fontSize: 13.5 },
  quick: { display: 'flex', alignItems: 'center', gap: 10, border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '12px 18px', font: 'inherit', fontFamily: FONT, fontSize: 13, fontWeight: 300, color: '#2B2321', cursor: 'pointer', textAlign: 'left' },
};
const IT_MONTHS_SHORT = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
const fmtDayMonth = (d) => `${String(d.getDate()).padStart(2, '0')} ${IT_MONTHS_SHORT[d.getMonth()]}`;
const fmtDayMonthYear = (d) => `${fmtDayMonth(d)} ${d.getFullYear()}`;
const fmtTime = (d) => `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
// Le date "Inizio/Fine lavori" arrivano dal campo data del progetto (AAAA-MM-GG) o, se scritte a mano, in formato italiano.
function parseWorkDate(v) {
  if (!v) return null;
  const iso = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  return parseItDate(v);
}
function relDayLabel(d, now) {
  const a = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const b = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const diff = Math.round((b - a) / 86400000);
  if (diff === 0) return `oggi, ${fmtTime(d)}`;
  if (diff === 1) return `ieri, ${fmtTime(d)}`;
  return `${fmtDayMonth(d)}, ${fmtTime(d)}`;
}
function DashIcon({ children, stroke }) {
  return (
    <svg viewBox="0 0 24 24" style={{ width: 17, height: 17, flex: 'none', fill: 'none', stroke: stroke || C.maroon, strokeWidth: 1.6, strokeLinecap: 'round', strokeLinejoin: 'round' }}>{children}</svg>
  );
}
function DashKpi({ label, value, note, noteTone, valueColor, compactLabel = true, highlight = false }) {
  return (
    <div style={{ ...DASH.card, padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 6, borderColor: highlight ? `color-mix(in srgb, ${C.maroon} 35%, #E8E0D4)` : '#E8E0D4' }}>
      <div style={{ fontSize: compactLabel ? 10.5 : 12.5, letterSpacing: compactLabel ? '1.2px' : 0, color: '#6F655F', display: 'flex', alignItems: 'center', gap: 8 }}>{label}<KpiInfo /></div>
      <div style={{ fontSize: 30, fontWeight: 500, letterSpacing: '-.03em', lineHeight: 1.1, color: valueColor || '#2B2321' }}>{value}</div>
      {note && <div style={{ fontSize: 11.5, color: noteTone === 'up' ? '#4F7A4B' : noteTone === 'down' ? '#B3452B' : '#6F655F' }}>{note}</div>}
    </div>
  );
}

function Dashboard({ onNavigate, onOpenProject, projects, listini, activeListinoId, fornitoriCatalog, profile }) {
  const activeListino = listini.find((l) => l.id === activeListinoId) || listini[0];
  const listinoVoci = activeListino ? flattenListino(activeListino) : [];
  const macrosezioni = activeListino?.macros?.length || 0;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const sameMonth = (d, y, m) => d && d.getFullYear() === y && d.getMonth() === m;
  const prevMonthRef = new Date(now.getFullYear(), now.getMonth() - 1, 1);

  // Proposte di prezzo dei fornitori esterni ancora da approvare (stessa tabella del pannello "🔗 Fornitori" del Listino).
  const [proposteInAttesa, setProposteInAttesa] = useState(null);
  useEffect(() => {
    let cancelled = false;
    cea.from('fornitore_submissions').select('id', { count: 'exact', head: true }).eq('status', 'pending')
      .then(({ count, error }) => { if (!cancelled) setProposteInAttesa(error ? null : (count || 0)); });
    return () => { cancelled = true; };
  }, []);

  const allRevisions = projects.flatMap((p) => p.revisions.map((r) => ({
    project: p, revision: r, created: parseItDate(r.dateCreated), modified: parseItDate(r.dateModified),
  })));

  // Valore progetti attivi = somma del totale impresa del computo attivo (ultima revisione) di ogni progetto.
  // Lo stesso calcolo "a una certa data" (contando solo le revisioni create prima) dà il confronto col mese
  // precedente e i punti del grafico a linea, così pill e grafico sono dati veri e non decorazione.
  const valoreAl = (cutoff) => projects.reduce((sum, p) => {
    const revs = cutoff ? p.revisions.filter((r) => { const d = parseItDate(r.dateCreated); return d && d < cutoff; }) : p.revisions;
    const last = revs[revs.length - 1];
    return sum + (last ? parseEuro(last.total) : 0);
  }, 0);
  const valoreAttivo = valoreAl(null);
  const valoreMesePrec = valoreAl(new Date(now.getFullYear(), now.getMonth(), 1));
  const deltaValorePct = valoreMesePrec > 0 ? Math.round(((valoreAttivo - valoreMesePrec) / valoreMesePrec) * 100) : null;
  const sparkVals = [];
  for (let i = 5; i >= 0; i -= 1) sparkVals.push(valoreAl(new Date(now.getFullYear(), now.getMonth() - i + 1, 1)));
  sparkVals.push(valoreAttivo);
  const sMin = Math.min(...sparkVals);
  const sMax = Math.max(...sparkVals);
  const sparkPts = sparkVals.map((v, i) => [Math.round((i * 160) / (sparkVals.length - 1)), Math.round(44 - (sMax === sMin ? 0.5 : (v - sMin) / (sMax - sMin)) * 38)]);
  const sparkLine = sparkPts.map(([x, y], i) => `${i ? 'L' : 'M'}${x} ${y}`).join(' ');

  const firstRevDate = (p) => parseItDate(p.revisions[0]?.dateCreated);
  const progettiNuoviMese = projects.filter((p) => sameMonth(firstRevDate(p), now.getFullYear(), now.getMonth())).length;
  const revisioniQuestoMese = allRevisions.filter((x) => sameMonth(x.created, now.getFullYear(), now.getMonth())).length;
  const revisioniMesePrec = allRevisions.filter((x) => sameMonth(x.created, prevMonthRef.getFullYear(), prevMonthRef.getMonth())).length;
  const deltaRev = revisioniQuestoMese - revisioniMesePrec;
  const ultime48h = allRevisions.filter((x) => x.modified && (now - x.modified) / 36e5 <= 48);
  const progetti48h = new Set(ultime48h.map((x) => x.project.id)).size;

  const fornitoriNomi = new Set();
  let categorieForniture = 0;
  (fornitoriCatalog || []).forEach((macro) => (macro.categorie || []).forEach((cat) => {
    categorieForniture += 1;
    (cat.prodotti || []).forEach((prod) => (prod.varianti || []).forEach((v) => (v.fornitori || []).forEach((f) => fornitoriNomi.add(f.name))));
  }));

  // "Computi da approvare" = progetti il cui computo attivo (ultima revisione) è in attesa; "Revisioni in attesa"
  // = tutte le versioni in attesa di approvazione, anche quelle più vecchie dello stesso progetto.
  const computiDaApprovare = projects
    .map((p) => ({ project: p, revision: p.revisions[p.revisions.length - 1] }))
    .filter((x) => x.revision && x.revision.status === 'In attesa di approvazione')
    .map((x) => ({ ...x, modified: parseItDate(x.revision.dateModified) }))
    .sort((a, b) => (a.modified?.getTime() || 0) - (b.modified?.getTime() || 0));
  const revisioniInAttesa = allRevisions.filter((x) => x.revision.status === 'In attesa di approvazione');
  const oltre5gg = revisioniInAttesa.filter((x) => !x.modified || daysBetween(x.modified, now) > 5).length;

  const work = projects.map((p) => {
    const h = p.header || {};
    const start = parseWorkDate(h.dataInizioLavori);
    const endRaw = parseWorkDate(h.dataFineLavori);
    const end = endRaw && start && endRaw < start ? start : endRaw;
    return { project: p, start, end, status: latestStatus(p) };
  });
  const inPartenza30 = work.filter((w) => w.start && w.start >= today && daysBetween(today, w.start) <= 30).length;

  const recentiOrdinate = allRevisions.filter((x) => x.modified).sort((a, b) => b.modified - a.modified);
  const modificheRecenti = (ultime48h.length ? recentiOrdinate.filter((x) => (now - x.modified) / 36e5 <= 48) : recentiOrdinate).slice(0, 5);
  const computiRecenti = recentiOrdinate.slice(0, 6);

  // Cronoprogramma: 6 mesi a partire dal mese corrente.
  const ganttStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const ganttEnd = new Date(now.getFullYear(), now.getMonth() + 6, 1);
  const span = ganttEnd - ganttStart;
  const pct = (d) => Math.min(100, Math.max(0, ((d - ganttStart) / span) * 100));
  const todayPct = pct(now);
  const ganttMonths = Array.from({ length: 6 }, (_, i) => {
    const m = IT_MONTHS_SHORT[(now.getMonth() + i) % 12];
    return m.charAt(0).toUpperCase() + m.slice(1);
  });
  const ganttRows = work
    .filter((w) => w.start)
    .map((w) => {
      const end = w.end || new Date(w.start.getFullYear(), w.start.getMonth(), w.start.getDate() + 30);
      let kind; let label; let sub;
      if (w.status === 'In fase di cantiere') { kind = 'cantiere'; label = 'Esecuzione lavori'; sub = 'in cantiere'; }
      else if (w.start > today && daysBetween(today, w.start) <= 30) { kind = 'breve'; label = 'Avvio previsto'; sub = 'in partenza a breve'; }
      else if (w.start > today) { kind = 'futuro'; label = w.status === 'Approvato' ? 'Avvio programmato' : 'In attesa di conferma'; sub = 'non ancora iniziato'; }
      else { kind = 'attivo'; label = w.status === 'Approvato' ? 'Computo approvato' : 'Computo in revisione'; sub = 'attivo'; }
      return { ...w, end, kind, label, sub, left: pct(w.start), right: pct(end) };
    })
    .filter((r) => r.right > 0 && r.left < 100)
    .sort((a, b) => a.start - b.start);
  const barStyle = {
    attivo: { background: C.maroon, color: '#fff' },
    cantiere: { background: '#3A332F', color: '#fff' },
    breve: { background: '#C98A2E', color: '#fff' },
    futuro: { background: 'transparent', border: '1.6px dashed #B9AE9F', color: '#8B8178' },
  };

  // Scadenze imminenti: inizio e fine lavori nei prossimi 7 giorni.
  const scadenze = [];
  work.forEach((w) => {
    if (w.start && w.start >= today && daysBetween(today, w.start) <= 7) scadenze.push({ title: 'Inizio lavori', where: w.project.name, date: w.start, project: w.project });
    if (w.end && w.end >= today && daysBetween(today, w.end) <= 7) scadenze.push({ title: 'Fine lavori', where: w.project.name, date: w.end, project: w.project });
  });
  scadenze.sort((a, b) => a.date - b.date);

  // Situazione listino
  const vociModificateMese = listini.reduce((n, l) => n + flattenListino(l).filter((v) => {
    const d = v.updatedAt ? new Date(v.updatedAt) : null;
    return d && !isNaN(d) && sameMonth(d, now.getFullYear(), now.getMonth());
  }).length, 0);
  const rilieviInCorso = projects.filter((p) => (p.planimetrie || []).length > 0).length;

  const oggiLabel = now.toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).toUpperCase();
  const nomeBase = (profile?.name || (profile?.email ? profile.email.split('@')[0].replace(/[._]/g, ' ') : '')).trim();
  const nomeUtente = nomeBase ? nomeBase.split(/\s+/)[0].replace(/^\w/, (c) => c.toUpperCase()) : '';
  const revName = (r) => r.customName || r.label;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 22, color: '#2B2321' }}>
      {/* HERO */}
      <div>
        <div style={{ fontSize: 11, letterSpacing: '.2em', color: '#8B8178', fontWeight: 600, marginBottom: 8 }}>{oggiLabel}</div>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(44px, 7vw, 90px)', lineHeight: 1.05, fontWeight: 200, letterSpacing: '-.03em', color: '#2B2321' }}>
          Bentornato{nomeUtente ? `, ${nomeUtente}` : ''}.
        </h1>
        <p style={{ margin: '8px 0 0', fontSize: 16, fontWeight: 200, color: '#4A403C' }}>Ecco listino, progetti e computi metrici dello studio, in tempo reale.</p>
      </div>

      {/* KPI PRINCIPALE */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
        <div style={{ ...DASH.card, color: '#fff', padding: '20px 22px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', minHeight: 168, border: 0, width: 602, maxWidth: '100%', background: 'linear-gradient(145deg, #8A4443 0%, #15110F 100%)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10 }}>
            <div style={{ fontSize: 15, color: '#fff', fontWeight: 300 }}>Valore progetti attivi</div>
            {deltaValorePct !== null && (
              <span style={{ ...DASH.pill, background: 'rgba(255,255,255,.12)', color: '#fff', fontSize: 15 }}>{deltaValorePct >= 0 ? '+' : ''}{deltaValorePct}% vs mese prec.</span>
            )}
          </div>
          <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 16 }}>
            <div style={{ fontSize: 'clamp(38px, 5.5vw, 70px)', fontWeight: 200, letterSpacing: 0, lineHeight: 1, whiteSpace: 'nowrap' }}>€ {itNum(Math.round(valoreAttivo), 0)}</div>
            <svg viewBox="0 0 160 50" style={{ width: 150, height: 48, flex: 'none' }}>
              <defs><linearGradient id="dashSpk" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#fff" stopOpacity=".25" /><stop offset="1" stopColor="#fff" stopOpacity="0" /></linearGradient></defs>
              <path d={`${sparkLine} L160 50 L0 50Z`} fill="url(#dashSpk)" />
              <path d={sparkLine} fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
        </div>
      </div>

      {/* KPI */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))', gap: 12 }}>
        <DashKpi label="PROGETTI ATTIVI" value={projects.length} note={`+${progettiNuoviMese} questo mese`} noteTone={progettiNuoviMese > 0 ? 'up' : null} />
        <DashKpi label="VOCI DI LISTINO" value={listinoVoci.length} note={`su ${listini.length} listin${listini.length === 1 ? 'o' : 'i'} attiv${listini.length === 1 ? 'o' : 'i'}`} />
        <DashKpi label="MACROSEZIONI" value={macrosezioni} note={`catalogo ${(activeListino?.name || 'listino').toLowerCase()}`} />
        <DashKpi label="COMPUTI DA APPROVARE" value={computiDaApprovare.length} valueColor={C.maroon} note="in attesa di approvazione" />
        <DashKpi label="REVISIONI QUESTO MESE" value={revisioniQuestoMese} note={`${deltaRev >= 0 ? '+' : ''}${deltaRev} vs mese prec.`} noteTone={deltaRev > 0 ? 'up' : deltaRev < 0 ? 'down' : null} />
        <DashKpi label="FORNITORI COLLEGATI" value={fornitoriNomi.size} note={`su ${categorieForniture} categorie`} />
        <DashKpi label="REVISIONI IN ATTESA" value={revisioniInAttesa.length} valueColor={C.maroon} note={`${oltre5gg} oltre i 5 giorni`} compactLabel={false} highlight />
        <DashKpi label="MODIFICHE ULTIME 48H" value={ultime48h.length} note={`su ${progetti48h} progett${progetti48h === 1 ? 'o' : 'i'}`} noteTone={ultime48h.length > 0 ? 'up' : null} compactLabel={false} />
        <DashKpi label="IN PARTENZA ENTRO 30GG" value={inPartenza30} valueColor={C.maroon} note="da preparare listino/team" compactLabel={false} />
      </div>

      {/* COMPUTI DA APPROVARE */}
      <section style={{ ...DASH.card, padding: '18px 22px', overflowX: 'auto' }}>
        <h3 style={{ ...DASH.h3, fontSize: 20 }}>COMPUTI DA APPROVARE<span style={DASH.tag2}>ordinati per urgenza</span></h3>
        <table style={{ borderCollapse: 'collapse', width: '100%', marginTop: 10, minWidth: 560 }}>
          <thead><tr>{['Progetto', 'Revisione', 'Richiesta da', 'In attesa da', 'Totale impresa', 'Urgenza'].map((h) => <th key={h} style={DASH.th}>{h}</th>)}</tr></thead>
          <tbody>
            {computiDaApprovare.length === 0 && (
              <tr><td colSpan={6} style={{ ...DASH.td, color: '#8B8178', fontWeight: 300 }}>Nessun computo in attesa di approvazione.</td></tr>
            )}
            {computiDaApprovare.map(({ project, revision, modified }) => {
              const giorni = modified ? daysBetween(modified, now) : null;
              const critica = giorni === null || giorni >= 5;
              return (
                <tr key={project.id} onClick={() => onOpenProject(project.id)} style={{ cursor: 'pointer' }}>
                  <td style={{ ...DASH.td, fontWeight: 400 }}>{project.name}</td>
                  <td style={DASH.td}>{revName(revision)}</td>
                  <td style={DASH.td}>{revision.modifiedBy || revision.createdBy || '—'}</td>
                  <td style={{ ...DASH.td, ...DASH.mono }}>{giorni === null ? '—' : `${giorni} giorn${giorni === 1 ? 'o' : 'i'}`}</td>
                  <td style={{ ...DASH.td, ...DASH.mono }}>€ {itNum(Math.round(parseEuro(revision.total)), 0)}</td>
                  <td style={DASH.td}>
                    <span style={{ ...DASH.pill, background: critica ? 'color-mix(in srgb, #B3452B 16%, #fff)' : '#E8E0D6', color: critica ? '#8C3220' : '#5A4F4A' }}>{critica ? 'Critica' : "Da tenere d'occhio"}</span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </section>

      {/* MODIFICHE RECENTI */}
      <section style={{ ...DASH.card, padding: '18px 22px' }}>
        <h3 style={{ ...DASH.h3, fontSize: 20 }}>MODIFICHE RECENTI<span style={DASH.tag2}>{ultime48h.length ? 'ultime 48 ore' : 'le più recenti'}</span></h3>
        <div style={{ marginTop: 6 }}>
          {modificheRecenti.length === 0 && <div style={{ ...DASH.act, borderBottom: 0, fontSize: 13.5, fontWeight: 300, color: '#8B8178' }}>Nessuna modifica registrata.</div>}
          {modificheRecenti.map(({ project, revision, created, modified }, i) => {
            const nuova = created && modified && Math.abs(modified - created) < 60000;
            const chi = revision.modifiedBy || revision.createdBy;
            return (
              <div key={`${project.id}-${revision.id}`} onClick={() => onOpenProject(project.id)} style={{ ...DASH.act, borderBottom: i === modificheRecenti.length - 1 ? 0 : DASH.act.borderBottom, cursor: 'pointer' }}>
                <div style={{ ...DASH.actIc, background: nuova ? 'color-mix(in srgb, #4F7A4B 14%, #fff)' : `color-mix(in srgb, ${C.maroon} 14%, #fff)`, color: nuova ? '#2F5A2B' : C.maroon }}>{nuova ? '+' : '✎'}</div>
                <div>
                  <div style={{ fontSize: 13.5, fontWeight: 300 }}>
                    {chi ? <><b style={{ fontWeight: 500 }}>{chi}</b> {nuova ? 'ha creato' : 'ha modificato'} </> : (nuova ? 'Creato ' : 'Modificato ')}
                    il computo <b style={{ fontWeight: 500 }}>{revName(revision)}</b>
                  </div>
                  <div style={{ fontSize: 11.5, color: '#8B8178', marginTop: 2 }}>{project.name} · {relDayLabel(modified, now)}</div>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {/* AVANZAMENTO PROGETTI */}
      <section style={{ ...DASH.card, padding: '18px 22px', overflowX: 'auto' }}>
        <h3 style={{ ...DASH.h3, fontSize: 23 }}>AVANZAMENTO PROGETTI<span style={DASH.tag2}>cronoprogramma, 6 mesi</span></h3>
        <div style={{ minWidth: 620, marginTop: 16 }}>
          <div style={{ display: 'flex', paddingLeft: 168, marginBottom: 6 }}>
            {ganttMonths.map((m) => <span key={m} style={{ flex: 1, textAlign: 'center', fontSize: 11, color: '#8E847C', letterSpacing: '.04em' }}>{m}</span>)}
          </div>
          {ganttRows.length === 0 && (
            <p style={{ fontSize: 13, fontWeight: 300, color: '#8B8178', margin: '10px 0 0' }}>
              Inserisci «Inizio lavori» e «Fine lavori» nei Dati generali del computo di ogni progetto per vederlo nel cronoprogramma.
            </p>
          )}
          {ganttRows.map((r, i) => (
            <div key={r.project.id} onClick={() => onOpenProject(r.project.id)} style={{ display: 'flex', alignItems: 'center', padding: '9px 0', cursor: 'pointer' }}>
              <div style={{ width: 168, flex: 'none', paddingRight: 12, boxSizing: 'border-box' }}>
                <div style={{ fontSize: 13, fontWeight: 400, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.project.name}</div>
                <div style={{ fontSize: 11, color: '#8B8178' }}>{r.sub}</div>
              </div>
              <div style={{ position: 'relative', flex: 1, height: 22, background: 'repeating-linear-gradient(90deg,#EFE8DD,#EFE8DD 1px,transparent 1px,transparent calc(100%/6))', borderRadius: 6 }}>
                <div style={{ position: 'absolute', top: -4, bottom: -4, width: 0, left: `${todayPct}%`, borderLeft: `1.6px dashed ${C.maroon}` }}>
                  {i === 0 && <span style={{ position: 'absolute', top: -18, left: 4, fontSize: 9.5, color: C.maroon, fontWeight: 600, whiteSpace: 'nowrap' }}>Oggi</span>}
                </div>
                <div title={`${fmtDayMonthYear(r.start)} → ${fmtDayMonthYear(r.end)}`} style={{ position: 'absolute', top: 2, height: 18, borderRadius: 999, display: 'flex', alignItems: 'center', padding: '0 10px', fontSize: 10.5, fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', boxSizing: 'border-box', left: `${r.left}%`, width: `${Math.max(2, r.right - r.left)}%`, ...barStyle[r.kind] }}>{r.label}</div>
              </div>
            </div>
          ))}
          <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap', fontSize: 12, color: '#6F655F', marginTop: 14 }}>
            {[['Attivo', { background: C.maroon }], ['In cantiere', { background: '#3A332F' }], ['In partenza entro 30 gg', { background: '#C98A2E' }], ['Non ancora iniziato', { background: 'transparent', border: '1.6px dashed #B9AE9F', boxSizing: 'border-box' }]].map(([t, s]) => (
              <span key={t}><i style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 4, marginRight: 6, verticalAlign: -1, ...s }} />{t}</span>
            ))}
          </div>
        </div>
      </section>

      {/* SITUAZIONE LISTINO */}
      <section style={{ ...DASH.card, padding: '18px 22px' }}>
        <h3 style={{ ...DASH.h3, fontSize: 23, letterSpacing: 0 }}>SITUAZIONE LISTINO</h3>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '10px 28px', marginTop: 14 }}>
          {[
            ['Voci modificate questo mese', vociModificateMese],
            ['Listini attivi', listini.length],
            ['Proposte fornitori in attesa', proposteInAttesa === null ? '—' : proposteInAttesa],
            ['Rilievi planimetrici in corso', rilieviInCorso],
          ].map(([t, v]) => (
            <div key={t} style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ width: 10, height: 10, borderRadius: '50%', flex: 'none', background: C.maroon }} />
              <span style={{ fontSize: 13.5, flexGrow: 1, fontWeight: 300 }}>{t}</span>
              <b style={{ fontWeight: 600 }}>{v}</b>
            </div>
          ))}
        </div>
      </section>

      {/* COMPUTI MODIFICATI DI RECENTE */}
      <section style={{ ...DASH.card, padding: '18px 22px', overflowX: 'auto' }}>
        <h3 style={{ ...DASH.h3, fontSize: 23 }}>COMPUTI MODIFICATI DI RECENTE<span style={DASH.tag2}>{computiRecenti.length} totali</span></h3>
        <table style={{ borderCollapse: 'collapse', width: '100%', marginTop: 10, minWidth: 560 }}>
          <thead><tr>{['Progetto', 'Versione', 'Stato', 'Totale impresa', 'Modificato'].map((h) => <th key={h} style={{ ...DASH.th, fontWeight: 500, fontSize: 13 }}>{h}</th>)}</tr></thead>
          <tbody>
            {computiRecenti.length === 0 && <tr><td colSpan={5} style={{ ...DASH.td, color: '#8B8178', fontWeight: 300 }}>Nessun computo ancora creato.</td></tr>}
            {computiRecenti.map(({ project, revision, modified }) => {
              const pending = revision.status === 'In attesa di approvazione';
              return (
                <tr key={`${project.id}-${revision.id}`} onClick={() => onOpenProject(project.id)} style={{ cursor: 'pointer' }}>
                  <td style={{ ...DASH.td, fontWeight: 300 }}>{project.name}</td>
                  <td style={{ ...DASH.td, fontWeight: 300 }}>{revName(revision)}</td>
                  <td style={{ ...DASH.td, fontWeight: 300 }}>{revision.status}</td>
                  <td style={{ ...DASH.td, ...DASH.mono, ...(pending ? { color: C.maroon, fontWeight: 500 } : {}) }}>€ {itNum(Math.round(parseEuro(revision.total)), 0)}</td>
                  <td style={{ ...DASH.td, fontWeight: 300 }}>{fmtDayMonthYear(modified)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </section>

      {/* SCADENZE IMMINENTI */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'stretch' }}>
        <section style={{ ...DASH.card, flex: '2 1 320px', minWidth: 0, padding: '18px 22px' }}>
          <h3 style={{ ...DASH.h3, fontSize: 23 }}>SCADENZE IMMINENTI<span style={{ ...DASH.tag2, fontWeight: 300 }}>prossimi 7 giorni</span></h3>
          <div style={{ marginTop: 6 }}>
            {scadenze.length === 0 && <div style={{ ...DASH.li, borderBottom: 0, fontWeight: 300, color: '#8B8178' }}>Nessuna scadenza nei prossimi 7 giorni (inizio/fine lavori dai Dati generali del computo).</div>}
            {scadenze.map((s, i) => {
              const vicina = daysBetween(today, s.date) <= 1;
              return (
                <div key={i} onClick={() => onOpenProject(s.project.id)} style={{ ...DASH.li, borderBottom: i === scadenze.length - 1 ? 0 : DASH.li.borderBottom, fontWeight: 300, cursor: 'pointer' }}>
                  <div style={{ flexGrow: 1 }}>
                    <div style={{ fontWeight: 400 }}>{s.title}</div>
                    <div style={{ fontSize: 12, color: '#8B8178', fontWeight: 300 }}>{s.where}</div>
                  </div>
                  <span style={{ ...DASH.mono, ...(vicina ? { color: C.maroon, fontWeight: 500 } : { fontWeight: 300 }) }}>{fmtDayMonth(s.date)}</span>
                </div>
              );
            })}
          </div>
        </section>
      </div>

      {/* AZIONI RAPIDE */}
      <section style={{ ...DASH.card, padding: '18px 22px' }}>
        <h3 style={{ ...DASH.h3, fontSize: 23 }}>AZIONI RAPIDE</h3>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 10, marginTop: 12 }}>
          <button onClick={() => onNavigate('computi')} style={DASH.quick}><DashIcon><rect x="5" y="2.5" width="14" height="19" rx="2" /><path d="M8 7h8 M8 12h2 M8 16h2 M14 12h2 M14 16h2" /></DashIcon>Apri computo e listino</button>
          <button onClick={() => onNavigate('confronto')} style={DASH.quick}><DashIcon><path d="M8 3v14 M3 12l5 5 5-5" /><path d="M16 21V7 M21 12l-5-5-5 5" /></DashIcon>Confronta revisioni</button>
          <button onClick={() => onNavigate('rilievo')} style={DASH.quick}><DashIcon><path d="M3 6l6-2.3 6 2.3 6-2.3v15l-6 2.3-6-2.3-6 2.3z" /><path d="M9 3.7v15 M15 6v15" /></DashIcon>Apri rilievo planimetrico</button>
          <button onClick={() => onNavigate('listino')} style={{ ...DASH.quick, border: 0, background: C.maroon, color: '#fff', fontWeight: 400 }}><DashIcon stroke="#fff"><path d="M4 4.5h7a2 2 0 0 1 2 2V20a2 2 0 0 0-2-1.5H4z" /><path d="M20 4.5h-7a2 2 0 0 0-2 2V20a2 2 0 0 1 2-1.5h7z" /></DashIcon>Nuova voce di listino</button>
        </div>
      </section>
    </div>
  );
}

function ListinoPage({ listini, setListini, activeId, setActiveId, projects, setProjects }) {
  const active = listini.find((l) => l.id === activeId);
  const [showFornitoreSharing, setShowFornitoreSharing] = useState(false);
  const [pendingCount, setPendingCount] = useState(0);

  // Solo per mostrare un pallino con il numero di richieste in attesa sul bottone, senza dover aprire il
  // pannello: un piccolo conteggio caricato all'apertura della pagina e ogni volta che il pannello si chiude
  // (così il numero si aggiorna subito dopo aver approvato/rifiutato qualcosa).
  React.useEffect(() => {
    let cancelled = false;
    cea.from('fornitore_submissions').select('id', { count: 'exact', head: true }).eq('status', 'pending').then(({ count }) => {
      if (!cancelled) setPendingCount(count || 0);
    });
    return () => { cancelled = true; };
  }, [showFornitoreSharing]);

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

  const [showDuplica, setShowDuplica] = useState(false);
  const duplicateListino = () => setShowDuplica(true);
  // Crea la copia del listino attivo: con tutti i prezzi, oppure con la sola struttura (macro, categorie,
  // sottocategorie, voci con codice, descrizione, unità e note) e i prezzi da compilare. Senza prezzi, il prezzo
  // cliente scritto come formula (es. "=impresa*1.1") resta: appena si inserisce il costo impresa il prezzo
  // cliente si ricalcola da solo; i prezzi cliente fissi invece vengono svuotati.
  const confirmDuplicate = (name, withPrices) => {
    const id = Math.max(...listini.map((l) => l.id)) + 1;
    const macros = structuredClone(active.macros);
    if (!withPrices) {
      macros.forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((sc) => (sc.voci || []).forEach((v) => {
        v.priceImpresa = '';
        v.priceCliente = String(v.priceCliente || '').trim().startsWith('=') ? v.priceCliente : '';
      }))));
    }
    setListini([...listini, { id, name, macros }]);
    setActiveId(id);
    setShowDuplica(false);
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

  const pillBtn = { border: `1px solid ${C.paleGray}`, background: C.white, color: C.black, fontSize: 12.5, fontWeight: 400, padding: '7px 14px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT };
  const voci = flattenListino(active);
  const categorieCount = withCodes(active.macros).reduce((n, m) => n + m.categorie.length, 0);

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 20, flexWrap: 'wrap', marginBottom: 22 }}>
        <div>
          <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(30px, 4vw, 42px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Listino prezzi</h1>
          <p style={{ margin: '6px 0 0', fontSize: 14.5, fontWeight: 300, color: C.darkGray, maxWidth: 520 }}>Catalogo voci per macrocategoria, categoria e sottocategoria — prezzo impresa e prezzo cliente su ogni riga.</p>
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
          <select
            value={activeId}
            onChange={(e) => setActiveId(Number(e.target.value))}
            style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 13, fontWeight: 300, fontFamily: FONT, color: C.black, cursor: 'pointer' }}
          >
            {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
          <button onClick={addListino} style={{ ...pillBtn, border: 'none', background: C.maroon, color: C.white, fontWeight: 500 }}>+ Nuovo listino</button>
          <button onClick={duplicateListino} style={pillBtn}>⧉ Duplica</button>
          <button onClick={renameListino} style={pillBtn}>✎ Rinomina</button>
          <button onClick={deleteListino} style={pillBtn}>🗑 Elimina</button>
          <button onClick={() => setShowFornitoreSharing(true)} style={{ ...pillBtn, position: 'relative' }}>
            🔗 Fornitori
            {pendingCount > 0 && (
              <span style={{ position: 'absolute', top: -6, right: -6, background: C.maroon, color: C.white, borderRadius: 999, minWidth: 18, height: 18, fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px' }}>{pendingCount}</span>
            )}
          </button>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12, marginBottom: 22 }}>
        <DKpi label="VOCI TOTALI" value={voci.length} note={`su ${listini.length} listin${listini.length === 1 ? 'o' : 'i'} attiv${listini.length === 1 ? 'o' : 'i'}`} />
        <DKpi label="MACROSEZIONI" value={active.macros.length} note={active.macros.map((m) => m.name).join(', ') || 'nessuna macrosezione'} />
        <DKpi label="CATEGORIE" value={categorieCount} note="su tutte le macrosezioni" />
      </div>

      <EditableCatalog macros={active.macros} setMacros={setMacrosForActive} projects={projects} setProjects={setProjects} listinoId={active.id} />

      {showDuplica && (
        <DuplicaListinoModal sourceName={active.name} onClose={() => setShowDuplica(false)} onConfirm={confirmDuplicate} />
      )}
      {showFornitoreSharing && (
        <FornitoreSharingModal listini={listini} setListini={setListini} activeId={activeId} onClose={() => setShowFornitoreSharing(false)} />
      )}
    </div>
  );
}

function DuplicaListinoModal({ sourceName, onClose, onConfirm }) {
  const [name, setName] = useState(`${sourceName} (copia)`);
  const [withPrices, setWithPrices] = useState(true);
  const option = (value, title, desc) => (
    <label style={{ display: 'flex', gap: 10, alignItems: 'flex-start', padding: '11px 12px', border: `1px solid ${withPrices === value ? C.maroon : C.paleGray}`, borderRadius: 12, cursor: 'pointer', background: withPrices === value ? 'rgba(134,45,45,0.04)' : C.white }}>
      <input type="radio" checked={withPrices === value} onChange={() => setWithPrices(value)} style={{ marginTop: 3, accentColor: C.maroon }} />
      <span>
        <span style={{ display: 'block', fontSize: 13, fontWeight: 600, color: C.black }}>{title}</span>
        <span style={{ display: 'block', fontSize: 11.5, color: C.gray, marginTop: 2, lineHeight: 1.4 }}>{desc}</span>
      </span>
    </label>
  );
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 14, padding: 22, width: 440, maxWidth: '100%' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 6px', color: C.black }}>Duplica listino</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>Copia di "{sourceName}" con tutte le macrocategorie, categorie, sottocategorie e voci.</p>
        <label style={{ display: 'block', fontSize: 11, fontWeight: 600, color: C.black, marginBottom: 5 }}>Nome della copia</label>
        <input value={name} onChange={(e) => setName(e.target.value)} autoFocus style={{ width: '100%', boxSizing: 'border-box', border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '9px 11px', fontSize: 13, fontFamily: FONT, marginBottom: 14 }} />
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 18 }}>
          {option(true, 'Con i prezzi', 'Copia anche i prezzi impresa e cliente di ogni voce.')}
          {option(false, 'Senza prezzi', 'Copia solo voci, descrizioni, unità di misura e note: i prezzi restano da compilare. Le formule del prezzo cliente (es. =impresa*1.1) vengono mantenute.')}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button onClick={onClose} style={{ border: 'none', background: '#6B625C', color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>Annulla</button>
          <button onClick={() => name.trim() && onConfirm(name.trim(), withPrices)} disabled={!name.trim()} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: name.trim() ? 'pointer' : 'default', opacity: name.trim() ? 1 : 0.6 }}>Duplica</button>
        </div>
      </div>
    </div>,
    document.body
  );
}

// Pannello "Fornitori": genera link (+ PIN) con cui un fornitore esterno, senza account, può proporre il
// costo impresa per le voci di un listino scelto, e mostra le proposte in attesa perché qualcuno in studio
// le approvi (a quel punto il valore entra davvero nel listino) o le rifiuti. Tutto passa dalle tabelle
// cea.fornitore_links / cea.fornitore_submissions e dalle funzioni RPC pubbliche che le proteggono con PIN.
function FornitoreSharingModal({ listini, setListini, activeId, onClose }) {
  const [tab, setTab] = useState('richieste'); // 'richieste' | 'link'
  const [links, setLinks] = useState(null);
  const [submissions, setSubmissions] = useState(null);
  const [currentUserId, setCurrentUserId] = useState(null);
  const macrosOf = (listinoId) => (listini.find((l) => l.id === listinoId)?.macros || []).map((m) => m.name);
  const [newListinoId, setNewListinoId] = useState(activeId);
  const [newNome, setNewNome] = useState('');
  // Macrocategorie del listino visibili al fornitore per il link che si sta per creare: per default tutte
  // quelle del listino scelto (il fornitore vede tutto), l'admin toglie la spunta a quelle da nascondere.
  // Cambiare listino nel select qui sotto reimposta questa lista sulle macrocategorie del nuovo listino.
  const [newMacroNames, setNewMacroNames] = useState(() => macrosOf(activeId));
  const [justCreated, setJustCreated] = useState(null); // { url, pin }
  const [busyId, setBusyId] = useState(null);

  const changeNewListino = (id) => { setNewListinoId(id); setNewMacroNames(macrosOf(id)); };
  const toggleMacro = (name) => {
    setNewMacroNames((prev) => (prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name]));
  };

  const reload = () => {
    cea.from('fornitore_links').select('*').order('created_at', { ascending: false }).then(({ data }) => setLinks(data || []));
    cea.from('fornitore_submissions').select('*').order('submitted_at', { ascending: false }).then(({ data }) => setSubmissions(data || []));
  };

  React.useEffect(() => {
    reload();
    supabase.auth.getUser().then(({ data }) => setCurrentUserId(data.user?.id || null));
  }, []);

  const listinoName = (id) => listini.find((l) => l.id === id)?.name || 'Listino eliminato';
  const linkFor = (linkId) => (links || []).find((l) => l.id === linkId);

  const createLink = async () => {
    if (newMacroNames.length === 0) { alert('Seleziona almeno una macrocategoria da condividere.'); return; }
    const token = (crypto.randomUUID ? crypto.randomUUID().replace(/-/g, '') : `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`);
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    // Tutte le macrocategorie selezionate = nessuna restrizione (null), come per i link creati prima di questa
    // funzione: più semplice da leggere in "Link esistenti" che ripetere l'elenco completo.
    const allSelected = newMacroNames.length === macrosOf(newListinoId).length;
    const { error } = await cea.from('fornitore_links').insert({
      token, pin, listino_id: newListinoId, nome_fornitore: newNome.trim() || null,
      macro_names: allSelected ? null : newMacroNames,
    });
    if (error) { alert('Creazione del link non riuscita: ' + error.message); return; }
    const url = `${window.location.origin}${window.location.pathname}?fornitore=${token}`;
    setJustCreated({ url, pin });
    setNewNome('');
    reload();
  };

  const revokeLink = async (link) => {
    if (!confirm(`Disattivare il link per "${link.nome_fornitore || 'questo fornitore'}"? Non potrà più aprirlo.`)) return;
    await cea.from('fornitore_links').update({ active: !link.active }).eq('id', link.id);
    reload();
  };

  const copyLink = (token) => {
    const url = `${window.location.origin}${window.location.pathname}?fornitore=${token}`;
    navigator.clipboard?.writeText(url).then(() => alert('Link copiato.')).catch(() => alert(url));
  };

  const approve = async (sub) => {
    const link = linkFor(sub.link_id);
    if (!link) { alert('Il link collegato a questa richiesta non esiste più.'); return; }
    const listino = listini.find((l) => l.id === link.listino_id);
    if (!listino) { alert('Il listino collegato a questa richiesta è stato eliminato: non posso applicare il valore.'); return; }
    const path = findVoceByCode(listino.macros, sub.voce_code);
    if (!path) { alert(`La voce "${sub.voce_desc || sub.voce_code}" non esiste più in questo listino (forse eliminata o rinominata): approvala manualmente dopo averla ricreata, se serve.`); return; }
    setBusyId(sub.id);
    const [mi, ci, si, vi] = path;
    const nextListini = structuredClone(listini);
    const nextListino = nextListini.find((l) => l.id === link.listino_id);
    nextListino.macros[mi].categorie[ci].sottocategorie[si].voci[vi].priceImpresa = sub.costo_impresa_proposto;
    setListini(nextListini);
    const { error } = await cea.from('fornitore_submissions').update({ status: 'approved', reviewed_at: new Date().toISOString(), reviewed_by: currentUserId }).eq('id', sub.id);
    setBusyId(null);
    if (error) { alert('Aggiornamento dello stato non riuscito: ' + error.message); return; }
    reload();
  };

  const reject = async (sub) => {
    if (!confirm('Rifiutare questa proposta? Il costo indicato dal fornitore non entrerà nel listino.')) return;
    setBusyId(sub.id);
    const { error } = await cea.from('fornitore_submissions').update({ status: 'rejected', reviewed_at: new Date().toISOString(), reviewed_by: currentUserId }).eq('id', sub.id);
    setBusyId(null);
    if (error) { alert('Aggiornamento dello stato non riuscito: ' + error.message); return; }
    reload();
  };

  const pending = (submissions || []).filter((s) => s.status === 'pending');
  const decided = (submissions || []).filter((s) => s.status !== 'pending').slice(0, 20);

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(23,23,23,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 16, width: 680, maxWidth: '100%', maxHeight: '88vh', overflowY: 'auto', padding: 24 }}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: 16 }}>
          <h2 style={{ fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Fornitori</h2>
          <button onClick={onClose} style={{ marginLeft: 'auto', background: 'none', border: 'none', fontSize: 18, color: C.gray, cursor: 'pointer' }}>✕</button>
        </div>

        <div style={{ display: 'flex', gap: 8, marginBottom: 18 }}>
          <button onClick={() => setTab('richieste')} style={{ background: tab === 'richieste' ? C.black : C.white, color: tab === 'richieste' ? C.white : C.black, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            Richieste in attesa{pending.length > 0 ? ` (${pending.length})` : ''}
          </button>
          <button onClick={() => setTab('link')} style={{ background: tab === 'link' ? C.black : C.white, color: tab === 'link' ? C.white : C.black, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            Link di condivisione
          </button>
        </div>

        {tab === 'richieste' && (
          <div>
            {submissions === null && <p style={{ fontSize: 12, color: C.gray }}>Caricamento…</p>}
            {submissions !== null && pending.length === 0 && <p style={{ fontSize: 12, color: C.gray }}>Nessuna richiesta in attesa.</p>}
            {pending.map((sub) => {
              const link = linkFor(sub.link_id);
              return (
                <div key={sub.id} style={{ border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: 12, marginBottom: 10 }}>
                  <div style={{ fontSize: 11, color: C.gray, marginBottom: 4 }}>
                    {link ? listinoName(link.listino_id) : '—'}{link?.nome_fornitore ? ` · ${link.nome_fornitore}` : ''} · {sub.submitted_at ? new Date(sub.submitted_at).toLocaleString('it-IT', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : ''}
                  </div>
                  <div style={{ fontSize: 13, color: C.black, marginBottom: 2 }}>{sub.voce_desc || sub.voce_code} <span style={{ color: C.gray, fontSize: 11 }}>({sub.voce_code})</span></div>
                  <div style={{ fontSize: 13, color: C.maroon, fontWeight: 600, marginBottom: sub.note ? 4 : 8 }}>Costo impresa proposto: {sub.costo_impresa_proposto} €</div>
                  {sub.note && <div style={{ fontSize: 12, color: C.darkGray, marginBottom: 8, fontStyle: 'italic' }}>“{sub.note}”</div>}
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button disabled={busyId === sub.id} onClick={() => approve(sub)} style={{ background: C.success, color: C.white, border: 'none', borderRadius: 8, padding: '6px 12px', fontSize: 12, fontWeight: 600, cursor: busyId === sub.id ? 'default' : 'pointer', opacity: busyId === sub.id ? 0.6 : 1 }}>✓ Approva</button>
                    <button disabled={busyId === sub.id} onClick={() => reject(sub)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '6px 12px', fontSize: 12, fontWeight: 600, color: C.maroon, cursor: busyId === sub.id ? 'default' : 'pointer', opacity: busyId === sub.id ? 0.6 : 1 }}>✕ Rifiuta</button>
                  </div>
                </div>
              );
            })}
            {decided.length > 0 && (
              <>
                <p style={{ fontSize: 11, fontWeight: 700, color: C.midGray, margin: '18px 0 8px' }}>Decise di recente</p>
                {decided.map((sub) => (
                  <div key={sub.id} style={{ fontSize: 12, color: C.darkGray, padding: '6px 0', borderTop: `1px solid ${C.paleGray}`, display: 'flex', justifyContent: 'space-between' }}>
                    <span>{sub.voce_desc || sub.voce_code}</span>
                    <span style={{ color: sub.status === 'approved' ? C.success : C.maroon, fontWeight: 600 }}>{sub.status === 'approved' ? 'Approvata' : 'Rifiutata'}</span>
                  </div>
                ))}
              </>
            )}
          </div>
        )}

        {tab === 'link' && (
          <div>
            <div style={{ border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: 14, marginBottom: 18 }}>
              <p style={{ fontSize: 12, color: C.darkGray, margin: '0 0 12px' }}>Genera un nuovo link: chi lo apre non ha bisogno di un account, gli basta il PIN che gli comunichi separatamente.</p>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                <div style={{ flex: '1 1 200px' }}>
                  <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Listino da condividere</label>
                  <select value={newListinoId} onChange={(e) => changeNewListino(Number(e.target.value))} style={{ width: '100%', fontSize: 13, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }}>
                    {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
                  </select>
                </div>
                <div style={{ flex: '1 1 200px' }}>
                  <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Nome fornitore (facoltativo)</label>
                  <input value={newNome} onChange={(e) => setNewNome(e.target.value)} placeholder="Es. Impresa Rossi" style={{ width: '100%', fontSize: 13, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                </div>
              </div>

              <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray, display: 'block', marginBottom: 6 }}>Macrocategorie visibili a questo fornitore</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 12 }}>
                {macrosOf(newListinoId).map((name) => (
                  <label key={name} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: C.black, background: C.sidebar, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '5px 12px', cursor: 'pointer' }}>
                    <input type="checkbox" checked={newMacroNames.includes(name)} onChange={() => toggleMacro(name)} />
                    {name}
                  </label>
                ))}
                {macrosOf(newListinoId).length === 0 && <span style={{ fontSize: 12, color: C.gray }}>Questo listino non ha ancora macrocategorie.</span>}
              </div>

              <button onClick={createLink} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 16px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Genera link</button>

              {justCreated && (
                <div style={{ marginTop: 14, background: C.sidebar, borderRadius: 10, padding: 12 }}>
                  <div style={{ fontSize: 11, color: C.midGray, marginBottom: 4 }}>Link</div>
                  <div style={{ fontSize: 12, color: C.black, wordBreak: 'break-all', marginBottom: 8 }}>{justCreated.url}</div>
                  <div style={{ fontSize: 11, color: C.midGray, marginBottom: 4 }}>PIN</div>
                  <div style={{ fontSize: 16, fontWeight: 700, letterSpacing: '0.2em', color: C.black, marginBottom: 10 }}>{justCreated.pin}</div>
                  <button onClick={() => navigator.clipboard?.writeText(justCreated.url)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '6px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Copia link</button>
                </div>
              )}
            </div>

            <p style={{ fontSize: 11, fontWeight: 700, color: C.midGray, margin: '0 0 8px' }}>Link esistenti</p>
            {links === null && <p style={{ fontSize: 12, color: C.gray }}>Caricamento…</p>}
            {links !== null && links.length === 0 && <p style={{ fontSize: 12, color: C.gray }}>Nessun link creato finora.</p>}
            {(links || []).map((link) => (
              <div key={link.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderTop: `1px solid ${C.paleGray}` }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, color: C.black }}>{link.nome_fornitore || 'Senza nome'} <span style={{ color: C.gray, fontSize: 11 }}>· {listinoName(link.listino_id)}</span></div>
                  <div style={{ fontSize: 11, color: C.gray }}>
                    PIN {link.pin} · {link.active ? 'attivo' : 'disattivato'} · {link.macro_names && link.macro_names.length > 0 ? `${link.macro_names.length} macrocategori${link.macro_names.length === 1 ? 'a' : 'e'}: ${link.macro_names.join(', ')}` : 'tutte le macrocategorie'}
                  </div>
                </div>
                <button onClick={() => copyLink(link.token)} style={rowBtnStyle}>Copia link</button>
                <button onClick={() => revokeLink(link)} style={{ ...rowBtnStyle, color: link.active ? C.maroon : C.success }}>{link.active ? 'Disattiva' : 'Riattiva'}</button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  , document.body);
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

// Identificativo stabile di ogni voce di listino (non cambia se la voce viene spostata, rinominata o
// ricodificata): serve a collegare le voci dei computi alla voce di listino da cui sono nate, per
// aggiornarne i prezzi quando cambia il listino.
const newVoceUid = () => `v${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
function ensureVoceUids(listini) {
  let changed = false;
  const next = (listini || []).map((l) => ({
    ...l,
    macros: (l.macros || []).map((m) => ({
      ...m,
      categorie: (m.categorie || []).map((c) => ({
        ...c,
        sottocategorie: (c.sottocategorie || []).map((s) => ({
          ...s,
          voci: (s.voci || []).map((v) => (v.uid ? v : (changed = true, { ...v, uid: newVoceUid() }))),
        })),
      })),
    })),
  }));
  return changed ? next : listini;
}
const normText = (s) => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
// Ricerca "a filtro" usata in computo, listino e fornitori: il testo cercato viene diviso in parole e una
// riga compare solo se le contiene TUTTE (in qualsiasi ordine), confrontando i campi indicati.
const searchTerms = (q) => String(q || '').trim().toLowerCase().split(/\s+/).filter(Boolean);
const matchesTerms = (terms, ...fields) => {
  if (!terms.length) return true;
  const hay = fields.map((f) => String(f || '')).join(' ').toLowerCase();
  return terms.every((tm) => hay.includes(tm));
};
function SidebarSearch({ value, onChange, placeholder, count }) {
  return (
    <div style={{ margin: '0 0 10px', flexShrink: 0 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '6px 10px' }}>
        <svg viewBox="0 0 24 24" width={14} height={14} fill="none" stroke="#8B8178" strokeWidth={1.6} strokeLinecap="round"><circle cx="11" cy="11" r="7" /><path d="M20 20l-3.5-3.5" /></svg>
        <input value={value} onChange={(e) => onChange(e.target.value)} onKeyDown={(e) => { if (e.key === 'Escape') onChange(''); }} placeholder={placeholder}
          style={{ border: 'none', outline: 'none', background: 'transparent', flex: 1, minWidth: 0, fontSize: 12, fontFamily: FONT, color: '#2B2321' }} />
        {value && <button onClick={() => onChange('')} aria-label="Cancella ricerca" style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 12, padding: 0 }}>✕</button>}
      </div>
      {value.trim() && <div style={{ fontSize: 11, color: count ? C.maroon : '#8B8178', margin: '5px 2px 0' }}>{count ? `${count} vo${count === 1 ? 'ce trovata' : 'ci trovate'}` : 'Nessuna voce trovata'}</div>}
    </div>
  );
}

// Quando cambiano i prezzi o la descrizione di una voce di listino, li aggiorna nelle voci dei computi ANCORA IN ATTESA DI
// APPROVAZIONE nate da quella voce. Non tocca mai:
//  - i computi approvati / in cantiere (restano con i prezzi con cui sono stati approvati);
//  - nelle revisioni successive a un computo approvato, le voci che c'erano già nel computo approvato
//    (stesso id voce): mantengono i prezzi approvati. Solo le voci aggiunte dopo seguono il listino.
// Collegamento: voci create da ora in poi → listinoRef { listinoId, uid }; voci create prima (senza
// collegamento) → stessa descrizione della voce di listino prima della modifica.
function propagateListinoPrices(projects, changes) {
  let touchedItems = 0;
  const touchedProjects = new Set();
  const next = projects.map((p) => {
    const revisions = p.revisions || [];
    const lockedIds = new Set(revisions.filter((r) => r.status && r.status !== STATUS_OPTIONS[0]).flatMap((r) => (r.items || []).map((it) => String(it.id))));
    let pChanged = false;
    const newRevs = revisions.map((r) => {
      if ((r.status || STATUS_OPTIONS[0]) !== STATUS_OPTIONS[0]) return r;
      let rChanged = false;
      const items = (r.items || []).map((it) => {
        if (it.type === 'subtotal' || lockedIds.has(String(it.id))) return it;
        const ch = changes.find((c) => (it.listinoRef
          ? it.listinoRef.uid === c.uid && String(it.listinoRef.listinoId) === String(c.listinoId)
          : normText(it.desc) && normText(it.desc) === normText(c.oldDesc)));
        if (!ch) return it;
        const upd = {};
        if (ch.priceChanged) {
          const ui = formatEuro(ch.impresa).replace(' €', '');
          const uc = formatEuro(evalClientPrice(ch.clienteFormula, ch.impresa)).replace(' €', '');
          if (it.unitPriceImpresa !== ui) upd.unitPriceImpresa = ui;
          if (it.unitPriceCliente !== uc) upd.unitPriceCliente = uc;
        }
        if (ch.descChanged && it.desc !== ch.newDesc) upd.desc = ch.newDesc;
        // una voce creata prima del collegamento, riconosciuta dalla descrizione, da ora resta collegata
        if (!it.listinoRef) upd.listinoRef = { listinoId: ch.listinoId, uid: ch.uid };
        if (!upd.unitPriceImpresa && !upd.unitPriceCliente && upd.desc === undefined) return it;
        rChanged = true; touchedItems += 1;
        return { ...it, ...upd };
      });
      if (!rChanged) return r;
      pChanged = true;
      const real = items.filter((i) => i.type !== 'subtotal');
      return { ...r, items, total: formatEuro(sumImpresa(real)), totalCliente: formatEuro(sumCliente(real)), dateModified: nowLabel(), modifiedBy: 'Aggiornamento prezzi listino' };
    });
    if (!pChanged) return p;
    touchedProjects.add(p.id);
    return { ...p, revisions: newRevs, value: newRevs[newRevs.length - 1]?.total || p.value };
  });
  return { next, touchedItems, touchedProjects: touchedProjects.size };
}
// Confronta due versioni dei listini e restituisce le voci con prezzo impresa o cliente cambiato.
function diffListinoPrices(prev, next) {
  const index = new Map();
  (prev || []).forEach((l) => (l.macros || []).forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((s) => (s.voci || []).forEach((v) => {
    if (v.uid) index.set(`${l.id}|${v.uid}`, v);
  })))));
  const changes = [];
  (next || []).forEach((l) => (l.macros || []).forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((s) => (s.voci || []).forEach((v) => {
    const old = v.uid && index.get(`${l.id}|${v.uid}`);
    if (!old) return;
    const priceChanged = String(old.priceImpresa || '') !== String(v.priceImpresa || '') || String(old.priceCliente || '') !== String(v.priceCliente || '');
    const descChanged = normText(old.desc) !== normText(v.desc) && !!String(v.desc || '').trim();
    if (!priceChanged && !descChanged) return;
    changes.push({ listinoId: l.id, uid: v.uid, oldDesc: old.desc, newDesc: v.desc, priceChanged, descChanged, impresa: parseEuro(v.priceImpresa), clienteFormula: v.priceCliente });
  })))));
  return changes;
}

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

function VoceModal({ locations, initialLocationIdx = 0, initialVoce = null, onClose, onSave, offerAddToComputo = false }) {
  const isEdit = !!initialVoce;
  const [addToComputo, setAddToComputo] = useState(true);
  const [locationIdx, setLocationIdx] = useState(initialLocationIdx);
  const [code, setCode] = useState(initialVoce?.code || '');
  const [desc, setDesc] = useState(initialVoce?.desc || '');
  const startsKnown = initialVoce && UNIT_OPTIONS.slice(0, -1).includes(initialVoce.unit);
  const [unit, setUnit] = useState(initialVoce ? (startsKnown ? initialVoce.unit : 'Altro') : 'm²');
  const [customUnit, setCustomUnit] = useState(initialVoce && !startsKnown ? initialVoce.unit : '');
  const [priceImpresa, setPriceImpresa] = useState(initialVoce?.priceImpresa || '');
  const [priceCliente, setPriceCliente] = useState(initialVoce?.priceCliente || '');
  const [note, setNote] = useState(initialVoce?.note || '');
  const previewImpresa = parseEuro(priceImpresa);
  const previewCliente = evalClientPrice(priceCliente, previewImpresa);

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 400, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
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
        <textarea
          value={desc}
          onChange={(e) => setDesc(e.target.value)}
          placeholder="Descrivi lavorazione, materiali e condizioni…"
          rows={2}
          style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px', fontFamily: 'inherit', resize: 'vertical', minHeight: 36 }}
        />

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

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Note (facoltative, solo uso interno)</label>
        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="Es. condizioni particolari, fornitori consigliati, riferimenti capitolato…"
          rows={2}
          style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px', fontFamily: 'inherit', resize: 'vertical' }}
        />

        {offerAddToComputo && !isEdit && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5, color: C.black, margin: '0 0 12px', cursor: 'pointer' }}>
            <input type="checkbox" checked={addToComputo} onChange={(e) => setAddToComputo(e.target.checked)} />
            Aggiungila subito anche a questo computo
          </label>
        )}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 6 }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              if (!desc.trim()) return;
              const finalUnit = unit === 'Altro' ? (customUnit || '—') : unit;
              onSave(locations[locationIdx].path, { code: code || '—', desc, unit: finalUnit, priceImpresa: priceImpresa || '0,00', priceCliente: priceCliente || '', note: note || '', updatedAt: new Date().toISOString() }, offerAddToComputo && !isEdit && addToComputo);
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            {isEdit ? 'Salva modifiche' : 'Salva voce'}
          </button>
        </div>
      </div>
    </div>
  , document.body);
}

const CALC_UNIT_OPTIONS = [
  { value: '', label: 'Manuale (inserisci direttamente la quantità)' },
  { value: 'ml', label: 'Metri lineari (ml) — par.ug × lunghezza' },
  { value: 'mq', label: 'Metri quadri (m²) — par.ug × lunghezza × larghezza' },
  { value: 'm3', label: 'Metri cubi / peso (m³) — par.ug × lunghezza × larghezza × H/peso' },
];

function emptyMisurazioneRow() {
  return { parUg: '1', lung: '', larg: '', hPeso: '', segno: '+' };
}

// L'unità di misura di una voce di listino suggerisce automaticamente la modalità di calcolo da usare nel
// computo: m² -> formula a metri quadri, m³ -> formula a metri cubi, ml -> formula a metri lineari; le unità
// "a corpo", "cadauna", "kg", "ora" e "Altro" non hanno una formula geometrica e restano a quantità manuale.
function defaultCalcForUnit(unit) {
  if (unit === 'm²') return 'mq';
  if (unit === 'm³') return 'm3';
  if (unit === 'ml') return 'ml';
  return '';
}

// Modale per creare/modificare una voce del computo con misurazioni reali (par.ug, lunghezza, larghezza,
// H/peso), oppure per unire questa nuova misurazione a una voce già esistente della stessa sottocategoria
// (stessa riga di Costo: la quantità sommata si moltiplica per l'unico prezzo unitario della voce di destinazione).
// "prefill" arriva quando la voce si crea a partire da una voce di Listino (drag&drop o pulsante +): porta già
// descrizione, unità e prezzi del listino, così l'utente deve solo inserire le misurazioni reali di cantiere.
// Voci con quantità "collegata": prendono la quantità (le misurazioni) di un'altra voce dello stesso
// computo, eventualmente moltiplicata (es. ×2 per due mani). qtyLink = { sourceId, factor }. Ogni volta
// che cambiano le voci del computo la quantità collegata viene ricalcolata, anche a catena (A←B←C).
function resolveLinkedQuantities(items) {
  let current = items;
  let changed = false;
  for (let pass = 0; pass < 6; pass += 1) {
    const byId = new Map(current.map((it) => [String(it.id), it]));
    let passChanged = false;
    current = current.map((it) => {
      if (!it.qtyLink) return it;
      const src = byId.get(String(it.qtyLink.sourceId));
      if (!src || String(src.id) === String(it.id)) return it;
      const q = Math.round(parseEuro(src.qty) * (Number(it.qtyLink.factor) || 1) * 100) / 100;
      const qs = String(q).replace('.', ',');
      if (qs === it.qty) return it;
      passChanged = true;
      return { ...it, qty: qs };
    });
    if (!passChanged) break;
    changed = true;
  }
  return changed ? current : items;
}

function VoceComputoModal({ macroName, categoriaName, sottoName, initialItem, prefill, mergeCandidates, linkCandidates = [], onClose, onSave }) {
  const isEdit = !!initialItem;
  const [desc, setDesc] = useState(initialItem?.desc || prefill?.desc || '');
  const startsKnown = (initialItem || prefill) && UNIT_OPTIONS.slice(0, -1).includes((initialItem || prefill).unit);
  const [unit, setUnit] = useState((initialItem || prefill) ? (startsKnown ? (initialItem || prefill).unit : 'Altro') : 'm²');
  const [customUnit, setCustomUnit] = useState((initialItem || prefill) && !startsKnown ? (initialItem || prefill).unit : '');
  const [priceImpresa, setPriceImpresa] = useState(initialItem?.unitPriceImpresa || prefill?.priceImpresa || '');
  const [priceCliente, setPriceCliente] = useState(initialItem?.unitPriceCliente || prefill?.priceCliente || '');
  const [unitaCalcolo, setUnitaCalcolo] = useState(initialItem ? (initialItem.qtyLink ? 'link' : (initialItem.unitaCalcolo || '')) : (prefill ? defaultCalcForUnit(prefill.unit) : 'mq'));
  const [linkSourceId, setLinkSourceId] = useState(initialItem?.qtyLink ? String(initialItem.qtyLink.sourceId) : '');
  const [linkFactor, setLinkFactor] = useState(initialItem?.qtyLink ? String(initialItem.qtyLink.factor).replace('.', ',') : '1');
  const [linkQuery, setLinkQuery] = useState('');
  const isLinked = unitaCalcolo === 'link';
  const linkSource = linkCandidates.find((c) => String(c.id) === linkSourceId) || null;
  const linkFactorNum = parseEuro(linkFactor) || 0;
  const initialRows = initialItem ? (initialItem.misurazioni || []).flatMap((g) => g.rows || []) : [];
  const [rows, setRows] = useState(initialRows.length ? initialRows : [emptyMisurazioneRow()]);
  const [manualQty, setManualQty] = useState(initialItem && !initialItem.unitaCalcolo ? (initialItem.qty || '') : '');
  const [mergeIntoId, setMergeIntoId] = useState('');
  const [note, setNote] = useState(initialItem?.note || prefill?.note || '');
  const [noteCliente, setNoteCliente] = useState(initialItem?.noteCliente || '');

  const previewImpresa = parseEuro(priceImpresa);
  const previewCliente = evalClientPrice(priceCliente, previewImpresa);
  const computedQty = isLinked
    ? Math.round((linkSource ? parseEuro(linkSource.qty) : 0) * linkFactorNum * 100) / 100
    : (unitaCalcolo ? computeGruppoTotal({ rows }, unitaCalcolo) : parseEuro(manualQty));
  // Copia le righe di misurazione di un'altra voce (una tantum: poi restano modificabili qui).
  const copyRowsFrom = (src) => {
    const srcRows = (src?.misurazioni || []).flatMap((g) => g.rows || []);
    if (!srcRows.length) { alert('La voce scelta non ha righe di misurazione (la sua quantità è inserita a mano): usa piuttosto "Collegata a un\'altra voce".'); return; }
    if (src.unitaCalcolo) setUnitaCalcolo(src.unitaCalcolo); else if (!unitaCalcolo || isLinked) setUnitaCalcolo('mq');
    setRows((rs) => {
      const keep = rs.filter((r) => r.lung || r.larg || r.hPeso || (r.parUg && r.parUg !== '1'));
      return [...keep, ...srcRows.map((r) => ({ ...r }))];
    });
  };
  const mergeTarget = mergeIntoId ? mergeCandidates.find((c) => String(c.id) === mergeIntoId) : null;

  const updateRow = (idx, field, value) => {
    setRows((rs) => rs.map((r, i) => (i === idx ? { ...r, [field]: value } : r)));
  };
  const addRow = () => setRows((rs) => [...rs, emptyMisurazioneRow()]);
  const removeRow = (idx) => setRows((rs) => (rs.length > 1 ? rs.filter((_, i) => i !== idx) : rs));

  const colStyle = (active) => ({ width: '100%', fontSize: 12, padding: '5px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}`, textAlign: 'right', background: active ? C.white : '#F2EEE6', color: active ? C.black : C.gray });
  const labelStyle = { fontSize: 11, fontWeight: 700, color: C.midGray };
  const fieldStyle = { width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' };

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 640, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 4px', color: C.black }}>{isEdit ? 'Modifica voce' : 'Nuova voce'}</h2>
        <p style={{ fontSize: 11, color: C.gray, margin: '0 0 16px' }}>{macroName} › {categoriaName || 'Generale'}</p>

        <label style={labelStyle}>Descrizione tecnica</label>
        <textarea
          value={desc}
          onChange={(e) => setDesc(e.target.value)}
          placeholder="Descrivi lavorazione, materiali e condizioni…"
          rows={2}
          style={{ ...fieldStyle, fontFamily: 'inherit', resize: 'vertical', minHeight: 36 }}
        />

        {!isEdit && mergeCandidates.length > 0 && (
          <>
            <label style={labelStyle}>Unisci a una voce già esistente (facoltativo)</label>
            <select value={mergeIntoId} onChange={(e) => setMergeIntoId(e.target.value)} style={fieldStyle}>
              <option value="">— Crea come voce nuova, con Costo proprio —</option>
              {mergeCandidates.map((c) => <option key={c.id} value={c.id}>{c.code} — {c.desc}</option>)}
            </select>
            {mergeTarget && (
              <p style={{ fontSize: 11, color: C.gray, margin: '-8px 0 12px' }}>
                Le misurazioni inserite qui si sommeranno alla quantità di "{mergeTarget.desc}": un'unica riga di Costo, al prezzo unitario già impostato su quella voce.
              </p>
            )}
          </>
        )}

        {!mergeIntoId && (
          <>
            <label style={labelStyle}>Unità di misura</label>
            <select value={unit} onChange={(e) => setUnit(e.target.value)} style={fieldStyle}>
              {UNIT_OPTIONS.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
            {unit === 'Altro' && (
              <input value={customUnit} onChange={(e) => setCustomUnit(e.target.value)} placeholder="Es. q.li, kWh…" style={fieldStyle} />
            )}
          </>
        )}

        <label style={labelStyle}>Modalità di calcolo della quantità</label>
        <select value={unitaCalcolo} onChange={(e) => setUnitaCalcolo(e.target.value)} style={fieldStyle}>
          {CALC_UNIT_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
          <option value="link" disabled={linkCandidates.length === 0}>Collegata a un'altra voce del computo (stessa quantità/misure)</option>
        </select>

        {isLinked ? (
          <div style={{ marginBottom: 12, border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '12px 14px', background: '#FAF8F4' }}>
            <label style={labelStyle}>Voce da cui prendere la quantità</label>
            <input value={linkQuery} onChange={(e) => setLinkQuery(e.target.value)} placeholder="Filtra per codice o descrizione…" style={{ ...fieldStyle, margin: '4px 0 6px' }} />
            <select value={linkSourceId} onChange={(e) => setLinkSourceId(e.target.value)} size={Math.min(6, Math.max(2, linkCandidates.length))} style={{ ...fieldStyle, margin: '0 0 10px', height: 'auto' }}>
              {linkCandidates.filter((c) => matchesTerms(searchTerms(linkQuery), c.code, c.desc)).map((c) => (
                <option key={c.id} value={String(c.id)}>{c.code ? `${c.code} — ` : ''}{String(c.desc || '').slice(0, 70)} ({c.qty} {c.unit})</option>
              ))}
            </select>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              <label style={labelStyle}>Moltiplicatore</label>
              <input value={linkFactor} onChange={(e) => setLinkFactor(e.target.value)} style={{ width: 70, fontSize: 12, padding: '6px 8px', borderRadius: 6, border: `1px solid ${C.paleGray}`, textAlign: 'right' }} />
              <span style={{ fontSize: 11, color: C.gray }}>es. 2 per due mani, 0,5 per metà</span>
            </div>
            <p style={{ fontSize: 12, fontWeight: 700, color: C.black, margin: '10px 0 0' }}>
              {linkSource ? <>Quantità: {linkSource.qty} {linkSource.unit} × {linkFactor || '1'} = {itNum(computedQty, 2)}</> : 'Scegli la voce di origine.'}
            </p>
            <p style={{ fontSize: 11, color: C.gray, margin: '6px 0 0' }}>
              Resta collegata: se cambiano le misure della voce di origine, questa quantità si aggiorna da sola.
              {linkSource && (linkSource.misurazioni || []).length > 0 && (
                <> Preferisci partire dalle sue misure e modificarle? <button type="button" onClick={() => copyRowsFrom(linkSource)} style={{ ...rowBtnStyle, fontSize: 11, padding: '2px 8px' }}>Copia le misurazioni invece</button></>
              )}
            </p>
          </div>
        ) : unitaCalcolo ? (
          <div style={{ marginBottom: 12 }}>
            <div className="table-scroll">
              <table style={{ width: '100%', minWidth: 480, borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ textAlign: 'right', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                    <th style={{ padding: '4px 4px', textAlign: 'left' }}>Segno</th>
                    <th style={{ padding: '4px 4px' }}>Par.ug.</th>
                    <th style={{ padding: '4px 4px' }}>Lunghezza</th>
                    <th style={{ padding: '4px 4px' }}>Larghezza</th>
                    <th style={{ padding: '4px 4px' }}>H / Peso</th>
                    <th style={{ padding: '4px 4px', textAlign: 'right' }}>Valore</th>
                    <th style={{ padding: '4px 4px' }}></th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, idx) => (
                    <tr key={idx}>
                      <td style={{ padding: '3px 4px' }}>
                        <select value={r.segno} onChange={(e) => updateRow(idx, 'segno', e.target.value)} style={{ fontSize: 12, padding: '5px 4px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}>
                          <option value="+">+ somma</option>
                          <option value="-">− si detrae</option>
                        </select>
                      </td>
                      <td style={{ padding: '3px 4px' }}><input value={r.parUg} onChange={(e) => updateRow(idx, 'parUg', e.target.value)} style={colStyle(true)} /></td>
                      <td style={{ padding: '3px 4px' }}><input value={r.lung} onChange={(e) => updateRow(idx, 'lung', e.target.value)} style={colStyle(true)} /></td>
                      <td style={{ padding: '3px 4px' }}><input value={r.larg} onChange={(e) => updateRow(idx, 'larg', e.target.value)} disabled={unitaCalcolo === 'ml'} style={colStyle(unitaCalcolo !== 'ml')} /></td>
                      <td style={{ padding: '3px 4px' }}><input value={r.hPeso} onChange={(e) => updateRow(idx, 'hPeso', e.target.value)} disabled={unitaCalcolo !== 'm3'} style={colStyle(unitaCalcolo === 'm3')} /></td>
                      <td style={{ padding: '3px 4px', textAlign: 'right', fontWeight: 700 }}>{itNum(computeMisurazioneRowValue(r, unitaCalcolo), 2)}</td>
                      <td style={{ padding: '3px 4px' }}><button onClick={() => removeRow(idx)} style={iconBtn}>🗑</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', marginTop: 8 }}>
              <button onClick={addRow} style={rowBtnStyle}>+ Riga di misurazione</button>
              {linkCandidates.some((c) => (c.misurazioni || []).length > 0) && (
                <select value="" onChange={(e) => { const src = linkCandidates.find((c) => String(c.id) === e.target.value); if (src) copyRowsFrom(src); }} style={{ fontSize: 11.5, padding: '5px 8px', borderRadius: 999, border: `1px solid ${C.paleGray}`, background: C.white, maxWidth: 320 }}>
                  <option value="">⇩ Importa le misurazioni da un'altra voce…</option>
                  {linkCandidates.filter((c) => (c.misurazioni || []).length > 0).map((c) => (
                    <option key={c.id} value={String(c.id)}>{c.code ? `${c.code} — ` : ''}{String(c.desc || '').slice(0, 60)}</option>
                  ))}
                </select>
              )}
            </div>
            <p style={{ fontSize: 12, fontWeight: 700, color: C.black, margin: '10px 0 0' }}>
              Quantità {mergeTarget ? 'da aggiungere' : 'totale'}: {itNum(computedQty, 2)} {mergeIntoId ? '' : (unit === 'Altro' ? customUnit : unit)}
              {mergeTarget && <> — nuovo totale voce: {itNum((computedQty + parseEuro(mergeTarget.qty)), 2)}</>}
            </p>
          </div>
        ) : (
          <>
            <label style={labelStyle}>Quantità</label>
            <input value={manualQty} onChange={(e) => setManualQty(e.target.value)} placeholder="0" style={fieldStyle} />
          </>
        )}

        {!mergeIntoId && (
          <>
            <div style={{ display: 'flex', gap: 10 }}>
              <div style={{ flex: 1 }}>
                <label style={labelStyle}>Prezzo impresa (€)</label>
                <input value={priceImpresa} onChange={(e) => setPriceImpresa(e.target.value)} placeholder="0,00" style={{ ...fieldStyle, marginBottom: 6 }} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={labelStyle}>Prezzo cliente (€ o formula)</label>
                <input value={priceCliente} onChange={(e) => setPriceCliente(e.target.value)} placeholder="Es. =impresa*1.3" style={{ ...fieldStyle, marginBottom: 6 }} />
              </div>
            </div>
            <p style={{ fontSize: 11, color: C.gray, margin: '0 0 12px' }}>
              Questa è l'unica riga di Costo della voce: {priceImpresa && <>anteprima impresa {formatEuro(previewImpresa)} → cliente {formatEuro(previewCliente)}, </>}
              totale impresa {formatEuro(previewImpresa * computedQty)}.
            </p>
          </>
        )}

        <label style={labelStyle}>Nota per il cliente <span style={{ fontWeight: 400, color: C.gray }}>(facoltativa — compare nel PDF e nell'Excel, anche nella versione cliente)</span></label>
        <textarea
          value={noteCliente}
          onChange={(e) => setNoteCliente(e.target.value)}
          placeholder="Es. materiale fornito dal committente, finitura da confermare in cantiere…"
          rows={2}
          style={{ ...fieldStyle, fontFamily: 'inherit', resize: 'vertical' }}
        />

        <label style={labelStyle}>Note interne <span style={{ fontWeight: 400, color: C.gray }}>(facoltative, solo uso interno — non vanno mai al cliente)</span></label>
        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="Es. dettagli di cantiere, condizioni particolari, promemoria…"
          rows={2}
          style={{ ...fieldStyle, fontFamily: 'inherit', resize: 'vertical' }}
        />

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 6 }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              if (!desc.trim() && !mergeIntoId) return;
              if (isLinked && !linkSource) { alert('Scegli la voce da cui prendere la quantità.'); return; }
              const finalUnit = unit === 'Altro' ? (customUnit || '—') : unit;
              onSave({
                desc, unit: finalUnit, priceImpresa: priceImpresa || '0,00', priceCliente: priceCliente || '',
                unitaCalcolo: isLinked ? null : (unitaCalcolo || null),
                misurazioni: unitaCalcolo && !isLinked ? [{ rows }] : [],
                manualQty: isLinked ? String(computedQty).replace('.', ',') : (manualQty || '0'),
                qtyLink: isLinked && linkSource ? { sourceId: linkSource.id, factor: linkFactorNum || 1 } : null,
                note: note || '',
                noteCliente: noteCliente.trim(),
                listinoRef: prefill?.listinoRef || null,
                editId: initialItem?.id || null,
                mergeIntoItemId: mergeIntoId || null,
              });
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            {isEdit ? 'Salva modifiche' : 'Salva voce'}
          </button>
        </div>
      </div>
    </div>
  , document.body);
}

// Aggiunge una voce di listino a un computo scelto dalla pagina Listino prezzi. Se il computo scelto è
// già approvato (o non è l'ultima versione) non viene toccato: si crea una nuova revisione in attesa di
// approvazione che contiene anche la nuova voce — gli approvati restano intatti.
function addListinoVoceToProject(project, revisionId, item) {
  let revisions = project.revisions || [];
  if (revisions.length === 0) {
    revisions = [{ id: Date.now(), label: 'Revisione 1', customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: [], extraSections: [], total: '0,00 €', totalCliente: '0,00 €' }];
  }
  const latest = revisions[revisions.length - 1];
  const target = (revisionId && typeof revisionId === 'object' && revisions.includes(revisionId) ? revisionId : revisions.find((r) => String(r.id) === String(revisionId))) || latest;
  const editable = target === latest && (target.status || STATUS_OPTIONS[0]) === STATUS_OPTIONS[0];
  const nextItems = regenerateItemCodes(resolveLinkedQuantities([...(target.items || []), item]));
  const real = nextItems.filter((it) => it.type !== 'subtotal');
  const total = formatEuro(sumImpresa(real));
  const totalCliente = formatEuro(sumCliente(real));
  if (editable) {
    return { project: { ...project, revisions: revisions.map((r) => (r === target ? { ...r, items: nextItems, dateModified: nowLabel(), total, totalCliente } : r)), value: total }, label: target.customName || target.label, forked: false };
  }
  const label = `Revisione ${revisions.length + 1}`;
  const newRev = { ...target, id: Date.now() + 1, label, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: nextItems, total, totalCliente };
  return { project: { ...project, revisions: [...revisions, newRev], value: total }, label, forked: true };
}

function ListinoAddToComputoModal({ voce, projects, onClose, onAdd }) {
  // Progetto e revisione si scelgono per POSIZIONE nell'elenco, non per id (gli id possono ripetersi nei dati).
  const sorted = [...projects].sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'it'));
  const [projectIdx, setProjectIdx] = useState(0);
  const project = sorted[projectIdx] || sorted[0];
  const revs = project?.revisions || [];
  const [revIdx, setRevIdx] = useState(Math.max(0, revs.length - 1));
  const [qty, setQty] = useState('1');
  const rev = revs[revIdx] || revs[revs.length - 1];
  const willFork = rev && ((rev.status || STATUS_OPTIONS[0]) !== STATUS_OPTIONS[0] || rev !== revs[revs.length - 1]);
  const impresa = parseEuro(voce.priceImpresa);
  const cliente = evalClientPrice(voce.priceCliente, impresa);
  const field = { width: '100%', boxSizing: 'border-box', fontSize: 12.5, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px', fontFamily: FONT, background: '#fff' };
  const label = { fontSize: 11, fontWeight: 700, color: C.midGray };
  const changeProject = (idx) => {
    setProjectIdx(idx);
    const r = sorted[idx]?.revisions || [];
    setRevIdx(Math.max(0, r.length - 1));
  };
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 14, padding: 22, width: 460, maxWidth: '100%', maxHeight: '90vh', overflowY: 'auto' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, fontWeight: 400, margin: '0 0 4px', color: C.black }}>Aggiungi al computo</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 14px' }}>
          <strong style={{ color: C.black, fontWeight: 600 }}>{voce.code}</strong> — {voce.desc}<br />
          {voce.unit} · impresa {formatEuro(impresa)} · <span style={{ color: C.maroon }}>cliente {formatEuro(cliente)}</span>
        </p>
        {sorted.length === 0 ? (
          <p style={{ fontSize: 13, color: C.gray }}>Non ci sono ancora progetti: creane uno dalla sezione Progetti.</p>
        ) : (
          <>
            <label style={label}>Progetto</label>
            <select value={projectIdx} onChange={(e) => changeProject(Number(e.target.value))} style={field}>
              {sorted.map((p, i) => <option key={i} value={i}>{p.name}{p.client ? ` — ${p.client}` : ''}</option>)}
            </select>
            <label style={label}>Computo (versione)</label>
            {revs.length === 0 ? (
              <p style={{ fontSize: 12, color: C.gray, margin: '4px 0 12px' }}>Questo progetto non ha ancora un computo: verrà creata la "Revisione 1".</p>
            ) : (
              <select value={revIdx} onChange={(e) => setRevIdx(Number(e.target.value))} style={field}>
                {revs.map((r, i) => <option key={i} value={i}>{r.customName || r.label} · {r.status || STATUS_OPTIONS[0]} · {r.dateModified}</option>)}
              </select>
            )}
            {willFork && (
              <p style={{ fontSize: 11.5, color: C.maroon, margin: '-6px 0 12px' }}>
                Questa versione è {rev.status !== STATUS_OPTIONS[0] ? 'già approvata' : 'una versione precedente'}: resterà intatta e la voce verrà aggiunta in una nuova revisione in attesa di approvazione.
              </p>
            )}
            <label style={label}>Quantità ({voce.unit})</label>
            <input value={qty} onChange={(e) => setQty(e.target.value)} style={field} />
            <p style={{ fontSize: 11, color: C.gray, margin: '-6px 0 14px' }}>Le misurazioni dettagliate (lunghezze, larghezze…) puoi inserirle poi dal computo con ✎.</p>
          </>
        )}
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
          <button onClick={onClose} style={{ background: 'transparent', color: C.black, border: `1px solid ${C.paleGray}`, padding: '9px 16px', borderRadius: 999, fontSize: 12.5, cursor: 'pointer', fontFamily: FONT }}>Annulla</button>
          {sorted.length > 0 && (
            <button onClick={() => onAdd(project, rev || null, qty || '1')} style={{ background: C.maroon, color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 500, cursor: 'pointer', fontFamily: FONT }}>Aggiungi</button>
          )}
        </div>
      </div>
    </div>
  , document.body);
}

function EditableCatalog({ macros, setMacros, projects, setProjects, listinoId }) {
  const [addToComputoVoce, setAddToComputoVoce] = useState(null); // voce di listino da aggiungere a un computo
  const [addedNotice, setAddedNotice] = useState('');
  const [showVoceModal, setShowVoceModal] = useState(false);
  const [editingVoce, setEditingVoce] = useState(null); // { path: [mi,ci,si,vi] }
  const [newVoceAt, setNewVoceAt] = useState(null); // [mi,ci,si] quando si aggiunge dal pulsante "+ Voce" di una sottocategoria
  const [expanded, setExpanded] = useState({});
  const coded = withCodes(macros);

  const isOpen = (key) => expanded[key] !== false; // default aperto
  const toggle = (key) => setExpanded({ ...expanded, [key]: !isOpen(key) });

  // Ricerca nel listino per codice e/o parole della descrizione (anche più parole, in qualsiasi ordine:
  // devono comparire tutte). Durante la ricerca restano visibili solo i rami che contengono voci trovate,
  // aperti in automatico, con le lettere cercate evidenziate.
  const [query, setQuery] = useState('');
  const terms = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const searching = terms.length > 0;
  const voceMatches = (v) => {
    const hay = `${v.code || ''} ${v.desc || ''} ${v.note || ''}`.toLowerCase();
    return terms.every((tm) => hay.includes(tm));
  };
  const sottoHasMatch = (so) => so.voci.some(voceMatches);
  const catHasMatch = (ca) => ca.sottocategorie.some(sottoHasMatch);
  const macroHasMatch = (ma) => ma.categorie.some(catHasMatch);
  const showOpen = (key) => searching || isOpen(key);
  const matchCount = searching ? coded.reduce((n, ma) => n + ma.categorie.reduce((n2, ca) => n2 + ca.sottocategorie.reduce((n3, so) => n3 + so.voci.filter(voceMatches).length, 0), 0), 0) : 0;
  const escRe = (x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const hl = (text) => {
    const str = String(text || '');
    if (!searching) return str;
    const parts = str.split(new RegExp(`(${terms.map(escRe).join('|')})`, 'gi'));
    return parts.map((p, i) => (terms.includes(p.toLowerCase())
      ? <mark key={i} style={{ background: C.accentSoft, color: C.maroon, fontWeight: 600, padding: 0, borderRadius: 2 }}>{p}</mark>
      : <React.Fragment key={i}>{p}</React.Fragment>));
  };

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
      const prevVoce = next[mi].categorie[ci].sottocategorie[si].voci[vi];
      next[mi].categorie[ci].sottocategorie[si].voci[vi] = { uid: prevVoce?.uid || newVoceUid(), ...voce };
    } else {
      const [mi, ci, si] = path;
      next[mi].categorie[ci].sottocategorie[si].voci.push({ uid: newVoceUid(), ...voce });
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

  // Sposta su/giù macrocategorie, categorie, sottocategorie o singole voci: i codici di macro/categoria/
  // sottocategoria si aggiornano da soli perché withCodes li ricalcola in base all'ordine nell'array.
  const moveItem = (kind, path, direction) => {
    const next = structuredClone(macros);
    let arr;
    if (kind === 'macro') arr = next;
    else if (kind === 'categoria') arr = next[path[0]].categorie;
    else if (kind === 'sotto') arr = next[path[0]].categorie[path[1]].sottocategorie;
    else arr = next[path[0]].categorie[path[1]].sottocategorie[path[2]].voci;
    const idx = path[path.length - 1];
    const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= arr.length) return;
    [arr[idx], arr[swapIdx]] = [arr[swapIdx], arr[idx]];
    setMacros(next);
  };

  const rowBtn = { border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, fontSize: 10.5, fontWeight: 500, padding: '3px 9px', cursor: 'pointer', color: C.darkGray, fontFamily: FONT };
  const codeTag = { fontFamily: 'ui-monospace, monospace', fontSize: 11, fontWeight: 600, color: C.maroon, background: 'rgba(134,45,45,0.08)', padding: '2px 7px', borderRadius: 999, marginLeft: 8 };

  return (
    <section style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px', marginBottom: 18 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <p style={{ fontSize: 18, fontWeight: 300, margin: 0, color: C.black, fontFamily: FONT, display: 'flex', alignItems: 'baseline', gap: 10 }}>
          Catalogo / Listino<span style={{ fontSize: 11.5, fontWeight: 400, color: '#8B8178' }}>{totals.voci} voci consultabili · clicca su una riga per aprirla o chiuderla</span>
        </p>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={() => { setEditingVoce(null); setNewVoceAt(null); setShowVoceModal(true); }} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 500, cursor: 'pointer', fontFamily: FONT }}>+ Nuova voce</button>
          <button onClick={addMacro} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 400, color: C.black, cursor: 'pointer', fontFamily: FONT }}>+ Macrocategoria</button>
        </div>
      </div>

      {coded.length > 0 && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 14, flexWrap: 'wrap', position: 'sticky', top: 'var(--header-h, 68px)', zIndex: 15, background: C.white, padding: '10px 0', margin: '0 0 14px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flex: '1 1 320px', maxWidth: 560, border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '8px 14px' }}>
            <svg viewBox="0 0 24 24" width={16} height={16} fill="none" stroke="#8B8178" strokeWidth={1.6} strokeLinecap="round"><circle cx="11" cy="11" r="7" /><path d="M20 20l-3.5-3.5" /></svg>
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Escape') setQuery(''); }}
              placeholder="Cerca nel listino per codice o parole (es. OE.DMI, demolizione pavimenti)…"
              style={{ border: 'none', outline: 'none', background: 'transparent', flex: 1, minWidth: 0, fontSize: 13, fontWeight: 300, fontFamily: FONT, color: C.black }}
            />
            {query && <button onClick={() => setQuery('')} aria-label="Cancella ricerca" style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 14, padding: 0 }}>✕</button>}
          </div>
          {searching && (
            <span style={{ fontSize: 12.5, color: matchCount ? C.maroon : '#8B8178', fontWeight: 500 }}>
              {matchCount ? `${matchCount} vo${matchCount === 1 ? 'ce trovata' : 'ci trovate'}` : 'Nessuna voce trovata'}
            </span>
          )}
        </div>
      )}

      {coded.length === 0 && (
        <p style={{ fontSize: 13, color: C.gray, margin: '0 0 10px' }}>Questo listino è vuoto. Aggiungi una macrocategoria per iniziare.</p>
      )}

      {coded.map((m, mi) => {
        const mKey = `${mi}`;
        if (searching && !macroHasMatch(m)) return null;
        return (
          <div key={mi} style={{ marginBottom: 10, border: '1px solid #2B2321', borderRadius: 10, overflow: 'hidden' }}>
            <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '13px 16px', background: '#2B2321', cursor: 'pointer', flexWrap: 'wrap' }}>
              <span style={{ color: '#B8ADA0', transition: 'transform .15s', transform: showOpen(mKey) ? 'rotate(90deg)' : 'none', display: 'inline-block' }}>▸</span>
              <span style={{ fontWeight: 500, fontSize: 14.5, color: '#FFFFFF', textTransform: 'uppercase', letterSpacing: '.02em' }}>{m.name}</span>
              <span style={{ ...codeTag, background: 'rgba(255,255,255,.14)', color: '#FFFFFF', border: 'none' }}>{m.code}</span>
              <span style={{ fontSize: 11, color: '#B8ADA0', fontWeight: 400 }}>{m.categorie.length} categorie · {m.categorie.reduce((n, c) => n + c.sottocategorie.reduce((n2, s) => n2 + s.voci.length, 0), 0)} voci</span>
              <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                <button onClick={() => moveItem('macro', [mi], 'up')} disabled={mi === 0} style={{ ...rowBtn, opacity: mi === 0 ? 0.4 : 1 }}>▲</button>
                <button onClick={() => moveItem('macro', [mi], 'down')} disabled={mi === coded.length - 1} style={{ ...rowBtn, opacity: mi === coded.length - 1 ? 0.4 : 1 }}>▼</button>
                <button onClick={() => rename('macro', [mi])} style={rowBtn}>✎ Rinomina</button>
                <button onClick={() => addCategoria(mi)} style={rowBtn}>+ Categoria</button>
                <button onClick={() => remove('macro', [mi])} style={rowBtn}>🗑</button>
              </div>
            </div>

            {showOpen(mKey) && m.categorie.map((c, ci) => {
              const cKey = `${mi}-${ci}`;
              if (searching && !catHasMatch(c)) return null;
              return (
                <div key={ci} style={{ padding: '10px 16px 10px 16px', borderTop: `1px solid ${C.paleGray}` }}>
                  <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', cursor: 'pointer', flexWrap: 'wrap', background: '#EFE8DD', borderRadius: 8, borderLeft: `4px solid ${C.maroon}` }}>
                    <span style={{ fontSize: 12, color: '#2B2321', transform: showOpen(cKey) ? 'rotate(90deg)' : 'none', display: 'inline-block' }}>▸</span>
                    <span style={{ fontSize: 14, letterSpacing: '.04em', textTransform: 'uppercase', color: '#2B2321', fontWeight: 600 }}>{c.name}</span>
                    <span style={codeTag}>{c.code}</span>
                    <span style={{ fontSize: 11, color: '#6F655F' }}>{(() => { const n = c.sottocategorie.reduce((a, so) => a + so.voci.length, 0); return `${n} voc${n === 1 ? 'e' : 'i'}`; })()}</span>
                    <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                      <button onClick={() => moveItem('categoria', [mi, ci], 'up')} disabled={ci === 0} style={{ ...rowBtn, opacity: ci === 0 ? 0.4 : 1 }}>▲</button>
                      <button onClick={() => moveItem('categoria', [mi, ci], 'down')} disabled={ci === m.categorie.length - 1} style={{ ...rowBtn, opacity: ci === m.categorie.length - 1 ? 0.4 : 1 }}>▼</button>
                      <button onClick={() => rename('categoria', [mi, ci])} style={rowBtn}>✎</button>
                      <button onClick={() => addSotto(mi, ci)} style={rowBtn}>+ Sottocategoria</button>
                      <button onClick={() => remove('categoria', [mi, ci])} style={rowBtn}>🗑</button>
                    </div>
                  </div>

                  {showOpen(cKey) && c.sottocategorie.map((s, si) => {
                    const sKey = `${mi}-${ci}-${si}`;
                    if (searching && !sottoHasMatch(s)) return null;
                    return (
                      <div key={si} style={{ marginLeft: 16 }}>
                        <div onClick={() => toggle(sKey)} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '9px 4px 7px', cursor: 'pointer', flexWrap: 'wrap', borderBottom: '1px solid #E8E0D4' }}>
                          <span style={{ fontSize: 11, color: '#4A403C', transform: showOpen(sKey) ? 'rotate(90deg)' : 'none', display: 'inline-block' }}>▸</span>
                          <span style={{ fontSize: 13, color: '#2B2321', fontWeight: 500 }}>{s.name}</span>
                          <span style={codeTag}>{s.code}</span>
                          <div onClick={(e) => e.stopPropagation()} style={{ marginLeft: 'auto', display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                            <button onClick={() => moveItem('sotto', [mi, ci, si], 'up')} disabled={si === 0} style={{ ...rowBtn, opacity: si === 0 ? 0.4 : 1 }}>▲</button>
                            <button onClick={() => moveItem('sotto', [mi, ci, si], 'down')} disabled={si === c.sottocategorie.length - 1} style={{ ...rowBtn, opacity: si === c.sottocategorie.length - 1 ? 0.4 : 1 }}>▼</button>
                            <button onClick={() => rename('sotto', [mi, ci, si])} style={rowBtn}>✎</button>
                            <button
                              onClick={() => { setEditingVoce(null); setNewVoceAt([mi, ci, si]); setShowVoceModal(true); }}
                              style={{ ...rowBtn, background: C.maroon, color: C.white, border: 'none' }}
                            >
                              + Voce
                            </button>
                            <button onClick={() => remove('sotto', [mi, ci, si])} style={rowBtn}>🗑</button>
                          </div>
                        </div>
                        {showOpen(sKey) && (
                          s.voci.length === 0 ? (
                            <p style={{ fontSize: 12.5, color: '#8B8178', padding: '4px 0 10px 24px' }}>Nessuna voce ancora in questa sottocategoria. Usa "+ Voce" per aggiungerne una.</p>
                          ) : (
                            <div className="table-scroll">
                              {s.voci.map((v, vi) => {
                                if (searching && !voceMatches(v)) return null;
                                const impresaVal = parseEuro(v.priceImpresa);
                                const clienteVal = evalClientPrice(v.priceCliente, impresaVal);
                                return (
                                  <div key={vi} style={{ display: 'flex', alignItems: 'center', gap: 14, padding: '9px 0 9px 24px', borderBottom: '1px solid #F0EAE0', fontSize: 13.5, flexWrap: 'wrap' }}>
                                    <span style={{ fontFamily: 'ui-monospace, monospace', fontSize: 11.5, color: '#8B8178', width: 90, flexShrink: 0 }}>{hl(v.code)}</span>
                                    <span style={{ flex: '1 1 160px', minWidth: 0, fontWeight: 300, color: C.black }}>
                                      {hl(v.desc)}
                                      {v.note && <span style={{ display: 'block', margin: '3px 0 0', fontSize: 10.5, color: '#8B8178', fontStyle: 'italic' }}>📝 {v.note}</span>}
                                    </span>
                                    <span style={{ width: 54, flexShrink: 0, color: '#8B8178', fontSize: 12 }}>{v.unit}</span>
                                    <span style={{ width: 92, flexShrink: 0, textAlign: 'right', fontWeight: 500, color: C.black }}>{formatEuro(impresaVal)}</span>
                                    <span style={{ width: 92, flexShrink: 0, textAlign: 'right', color: C.maroon, fontWeight: 500 }}>{formatEuro(clienteVal)}</span>
                                    <div style={{ display: 'flex', gap: 4, flexShrink: 0 }}>
                                      {setProjects && (
                                        <button
                                          onClick={() => setAddToComputoVoce({ ...v, macro: m.name, categoria: c.name, sotto: s.name })}
                                          title="Aggiungi questa voce a un computo"
                                          style={{ ...rowBtn, padding: '2px 10px', background: C.maroon, color: '#fff', border: 'none' }}
                                        >
                                          + Computo
                                        </button>
                                      )}
                                      <button onClick={() => moveItem('voce', [mi, ci, si, vi], 'up')} disabled={vi === 0} style={{ ...rowBtn, padding: '2px 6px', opacity: vi === 0 ? 0.4 : 1 }}>▲</button>
                                      <button onClick={() => moveItem('voce', [mi, ci, si, vi], 'down')} disabled={vi === s.voci.length - 1} style={{ ...rowBtn, padding: '2px 6px', opacity: vi === s.voci.length - 1 ? 0.4 : 1 }}>▼</button>
                                      <button
                                        onClick={() => { setEditingVoce({ path: [mi, ci, si, vi] }); setNewVoceAt(null); setShowVoceModal(true); }}
                                        style={{ ...rowBtn, padding: '2px 6px' }}
                                      >
                                        ✎
                                      </button>
                                      <button onClick={() => remove('voce', [mi, ci, si, vi])} style={{ ...rowBtn, padding: '2px 6px' }}>🗑</button>
                                    </div>
                                  </div>
                                );
                              })}
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
        );
      })}

      {addToComputoVoce && (
        <ListinoAddToComputoModal
          voce={addToComputoVoce}
          projects={projects || []}
          onClose={() => setAddToComputoVoce(null)}
          onAdd={(project, revision, qty) => {
            const v = addToComputoVoce;
            if (!project) return;
            const revs = project.revisions || [];
            const target = revision || revs[revs.length - 1];
            const dupe = (target?.items || []).find((it) => it.type !== 'subtotal' && normText(it.desc) === normText(v.desc));
            if (dupe && !confirm(`La voce "${v.desc}" è già presente in questo computo${dupe.code ? ` (${dupe.code})` : ''}. Vuoi aggiungere comunque un duplicato?`)) return;
            const impresa = parseEuro(v.priceImpresa);
            const item = {
              id: Date.now() + Math.random(), code: '', autoCode: true,
              listinoRef: v.uid ? { listinoId, uid: v.uid } : null,
              desc: v.desc, unit: v.unit, note: v.note || '',
              unitPriceImpresa: formatEuro(impresa).replace(' €', ''),
              unitPriceCliente: formatEuro(evalClientPrice(v.priceCliente, impresa)).replace(' €', ''),
              qty: String(qty), macro: v.macro, section: v.macro, categoria: v.categoria || 'Generale', sottocategoria: v.sotto || 'Generale',
            };
            const res = addListinoVoceToProject(project, target, item);
            setProjects((projects || []).map((p) => (p === project ? res.project : p)));
            setAddToComputoVoce(null);
            setAddedNotice(`Aggiunta a "${project.name}" — ${res.label}${res.forked ? ' (nuova revisione: la versione approvata resta intatta)' : ''}.`);
            setTimeout(() => setAddedNotice(''), 6000);
          }}
        />
      )}
      {addedNotice && createPortal(
        <div role="status" onClick={() => setAddedNotice('')} style={{ position: 'fixed', right: 24, bottom: 24, zIndex: 300, maxWidth: 380, background: '#2B2321', color: '#fff', borderRadius: 12, padding: '12px 16px', fontSize: 13, fontWeight: 300, boxShadow: '0 10px 30px rgba(0,0,0,.18)', cursor: 'pointer' }}>{addedNotice}</div>
      , document.body)}

      <p style={{ fontSize: 11.5, color: '#8B8178', margin: '10px 0 0', fontWeight: 300 }}>Con le freccette ▲▼ puoi riordinare manualmente macrocategorie, categorie, sottocategorie e singole voci: i codici si aggiornano da soli in base al nuovo ordine.</p>

      {showVoceModal && (
        <VoceModal
          locations={locations}
          initialLocationIdx={
            editingVoce ? locations.findIndex((l) => l.path.join() === editingVoce.path.slice(0, 3).join())
            : newVoceAt ? locations.findIndex((l) => l.path.join() === newVoceAt.join())
            : 0
          }
          initialVoce={editingVoce ? macros[editingVoce.path[0]].categorie[editingVoce.path[1]].sottocategorie[editingVoce.path[2]].voci[editingVoce.path[3]] : null}
          onClose={() => { setShowVoceModal(false); setEditingVoce(null); setNewVoceAt(null); }}
          onSave={saveVoce}
        />
      )}
    </section>
  );
}

function NewProjectModal({ onClose, onCreate }) {
  const [name, setName] = useState('');
  const [client, setClient] = useState('');

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 380, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
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
  , document.body);
}

function ProjectCard({ p, onOpen, onDelete }) {
  return (
    <div onClick={() => onOpen(p.id)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: 16, cursor: 'pointer' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, ...badgeStyles[statusTone[latestStatus(p)]] }}>{latestStatus(p)}</span>
        <button
          onClick={(e) => { e.stopPropagation(); onDelete(p.id); }}
          style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 14, padding: 0 }}
          title="Elimina progetto"
        >
          🗑️
        </button>
      </div>
      <p style={{ margin: 0, fontSize: 15.5, fontWeight: 500, color: C.black, fontFamily: FONT }}>{p.name}</p>
      <div style={{ fontSize: 12, color: '#8B8178', marginTop: 2 }}>{p.client}</div>
      <div style={{ fontSize: 19, fontWeight: 200, marginTop: 6, color: C.black, fontFamily: FONT }}>{p.value}</div>
      <div style={{ fontVariantNumeric: 'tabular-nums', color: '#6F655F', fontSize: 12.5, marginTop: 8 }}>{p.items} voci · {p.revisions.length} revision{p.revisions.length === 1 ? 'e' : 'i'} salvat{p.revisions.length === 1 ? 'a' : 'e'}</div>
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
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 20, flexWrap: 'wrap' }}>
        <div>
          <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Progetti</p>
          <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Progetti</h1>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Dati aggiornati</span>
          <button
            onClick={() => setShowModal(true)}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 400, cursor: 'pointer', fontFamily: FONT }}
          >
            + Nuovo progetto
          </button>
        </div>
      </div>
      <p style={{ fontSize: 13.5, fontWeight: 300, color: '#8B8178', margin: '8px 0 20px' }}>Seleziona un progetto per aprirlo, oppure crea un nuovo progetto.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(230px, 1fr))', gap: 14 }}>
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
          <div className="table-scroll">
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
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '10px 16px', fontSize: 12, fontWeight: 700, color: C.black, borderTop: `1px solid ${C.paleGray}` }}>
            Subtotale {section.name}&nbsp;&nbsp;{section.subtotal}
          </div>
        </div>
      ))}
    </>
  );
}

// Colori di sfondo delle intestazioni di macrosezione (testo sempre bianco sopra): tutti scelti abbastanza
// scuri/saturi da garantire un contrasto leggibile col testo bianco. In precedenza l'ultimo colore del ciclo
// era C.sidebar (#F6F4EF, quasi bianco) — su testo bianco risultava illeggibile dalla quarta macrosezione in poi.
const SECTION_COLORS = [DEFAULT_ACCENT, C.darkGray, '#94706C', '#4A3F35'];

function DraggableCatalogTree({ listino, onAdd, onEdit }) {
  const [expanded, setExpanded] = useState({});
  const [q, setQ] = useState('');
  const terms = searchTerms(q);
  const searching = terms.length > 0;
  const isOpen = (key) => searching || expanded[key] !== false;
  const toggle = (key) => setExpanded({ ...expanded, [key]: !(expanded[key] !== false) });
  const voceMatch = (v) => matchesTerms(terms, v.code, v.desc, v.note);

  if (!listino) return null;
  const coded = withCodes(listino.macros);

  if (coded.length === 0) {
    return <p style={{ fontSize: 12, color: C.gray }}>Questo listino non ha ancora voci.</p>;
  }

  const found = searching ? coded.reduce((n, m) => n + m.categorie.reduce((n2, c) => n2 + c.sottocategorie.reduce((n3, s) => n3 + s.voci.filter(voceMatch).length, 0), 0), 0) : 0;
  return (
    <div>
      <div style={{ position: 'sticky', top: 0, zIndex: 2, background: '#FFFDFA', paddingBottom: 2 }}>
        <SidebarSearch value={q} onChange={setQ} placeholder="Cerca nel listino (codice o parole)…" count={found} />
      </div>
      {coded.map((m, mi) => {
        const mKey = `m${mi}`;
        if (searching && !m.categorie.some((c) => c.sottocategorie.some((s) => s.voci.some(voceMatch)))) return null;
        return (
          <div key={mi} style={{ marginBottom: 8 }}>
            <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '6px 4px', borderRadius: 6, background: C.bg }}>
              <span style={{ fontSize: 10, color: C.gray }}>{isOpen(mKey) ? '⌄' : '›'}</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.black }}>{m.name}</span>
            </div>
            {isOpen(mKey) && m.categorie.map((c, ci) => {
              const cKey = `m${mi}c${ci}`;
              if (searching && !c.sottocategorie.some((s) => s.voci.some(voceMatch))) return null;
              return (
                <div key={ci} style={{ paddingLeft: 12 }}>
                  <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '5px 4px' }}>
                    <span style={{ fontSize: 10, color: C.gray }}>{isOpen(cKey) ? '⌄' : '›'}</span>
                    <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{c.name}</span>
                  </div>
                  {isOpen(cKey) && c.sottocategorie.map((s, si) => {
                    const sKey = `m${mi}c${ci}s${si}`;
                    if (searching && !s.voci.some(voceMatch)) return null;
                    return (
                      <div key={si} style={{ paddingLeft: 12 }}>
                        <div onClick={() => toggle(sKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '4px 4px' }}>
                          <span style={{ fontSize: 10, color: C.gray }}>{isOpen(sKey) ? '⌄' : '›'}</span>
                          <span style={{ fontSize: 12, color: C.midGray }}>{s.name}</span>
                        </div>
                        {isOpen(sKey) && (
                          <div style={{ paddingLeft: 12 }}>
                            {s.voci.map((v, vi) => {
                              if (searching && !voceMatch(v)) return null;
                              const impresaVal = parseEuro(v.priceImpresa);
                              const clienteVal = evalClientPrice(v.priceCliente, impresaVal);
                              const voceData = { ...v, macro: m.name, categoria: c.name, sotto: s.name, impresaValue: impresaVal, clienteValue: clienteVal };
                              return (
                              <div
                                key={vi}
                                draggable
                                onDragStart={(e) => e.dataTransfer.setData('application/json', JSON.stringify(voceData))}
                                style={{ border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '7px 9px', marginBottom: 6, cursor: 'grab', background: C.white, display: 'flex', alignItems: 'center', gap: 8 }}
                              >
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: C.maroon }}>{v.code}</p>
                                  <p style={{ margin: '2px 0 0', fontSize: 12, color: C.black }}>{v.desc}</p>
                                  <p style={{ margin: '2px 0 0', fontSize: 11, color: C.gray }}>{v.unit} · impresa {formatEuro(impresaVal)} · cliente {formatEuro(clienteVal)}</p>
                                </div>
                                {onEdit && (
                                  <button
                                    onClick={(e) => { e.stopPropagation(); onEdit([mi, ci, si, vi]); }}
                                    title="Modifica la voce nel listino"
                                    aria-label="Modifica la voce nel listino"
                                    style={{ flexShrink: 0, width: 26, height: 26, borderRadius: 999, border: `1px solid ${C.paleGray}`, background: C.white, color: C.midGray, fontSize: 12, lineHeight: 1, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                                  >
                                    ✎
                                  </button>
                                )}
                                {onAdd && (
                                  <button
                                    onClick={(e) => { e.stopPropagation(); onAdd(voceData); }}
                                    title="Aggiungi al computo"
                                    aria-label="Aggiungi al computo"
                                    style={{ flexShrink: 0, width: 30, height: 30, borderRadius: 999, border: `1px solid ${C.paleGray}`, background: C.white, color: C.maroon, fontSize: 16, fontWeight: 700, lineHeight: 1, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                                  >
                                    +
                                  </button>
                                )}
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

// Versione compatta del catalogo Fornitori per il pannello laterale del computo (sotto il Listino): stessa
// struttura macro > categoria > prodotto > variante di FornitoriPage, ma ogni "fonte prezzo" (listino del
// prodotto oppure uno dei fornitori reali) è una riga con il suo + per aggiungerla subito al computo che si
// sta modificando, senza dover scegliere progetto/revisione (si è già dentro quella giusta).
function FornitoriSidebarTree({ catalog, onAdd }) {
  const [expanded, setExpanded] = useState({});
  const [q, setQ] = useState('');
  const terms = searchTerms(q);
  const searching = terms.length > 0;
  const isOpen = (key) => searching || expanded[key] !== false;
  const toggle = (key) => setExpanded({ ...expanded, [key]: !(expanded[key] !== false) });
  // Una variante corrisponde se marca/modello del prodotto, nome della variante o nomi dei fornitori contengono tutte le parole cercate.
  const varMatch = (p, v) => matchesTerms(terms, prodottoLabel(p), v.nome, ...(v.fornitori || []).map((f) => f.name));
  const prodMatch = (p) => (p.varianti || []).some((v) => varMatch(p, v));

  const macros = catalog || [];
  if (macros.length === 0) {
    return <p style={{ fontSize: 12, color: C.gray }}>Catalogo fornitori vuoto: aggiungi prodotti dalla pagina "Fornitori".</p>;
  }

  const found = searching ? macros.reduce((n, m) => n + (m.categorie || []).reduce((n2, c) => n2 + (c.prodotti || []).filter(prodMatch).length, 0), 0) : 0;
  return (
    <div>
      <div style={{ position: 'sticky', top: 0, zIndex: 2, background: '#FFFDFA', paddingBottom: 2 }}>
        <SidebarSearch value={q} onChange={setQ} placeholder="Cerca marca, modello, variante o fornitore…" count={found} />
      </div>
      {macros.map((m, mi) => {
        const mKey = `fm${mi}`;
        const categorie = m.categorie || [];
        if (searching && !categorie.some((c) => (c.prodotti || []).some(prodMatch))) return null;
        return (
          <div key={mi} style={{ marginBottom: 8 }}>
            <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '6px 4px', borderRadius: 6, background: C.bg }}>
              <span style={{ fontSize: 10, color: C.gray }}>{isOpen(mKey) ? '⌄' : '›'}</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.black }}>{m.name}</span>
            </div>
            {isOpen(mKey) && categorie.map((c, ci) => {
              const cKey = `${mKey}c${ci}`;
              const prodotti = (c.prodotti || []).filter((p) => !searching || prodMatch(p));
              if (searching && prodotti.length === 0) return null;
              return (
                <div key={ci} style={{ paddingLeft: 12 }}>
                  <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '5px 4px' }}>
                    <span style={{ fontSize: 10, color: C.gray }}>{isOpen(cKey) ? '⌄' : '›'}</span>
                    <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{c.name}</span>
                  </div>
                  {isOpen(cKey) && prodotti.length === 0 && (
                    <p style={{ fontSize: 11, color: '#9A8F86', fontStyle: 'italic', margin: '2px 0 6px 12px' }}>Nessun prodotto.</p>
                  )}
                  {isOpen(cKey) && prodotti.map((p, pi) => {
                    const pKey = `${cKey}p${pi}`;
                    return (
                      <div key={p.id ?? pi} style={{ paddingLeft: 12, marginBottom: 4 }}>
                        <div onClick={() => toggle(pKey)} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', padding: '4px 4px' }}>
                          <span style={{ fontSize: 10, color: C.gray }}>{isOpen(pKey) ? '⌄' : '›'}</span>
                          <span style={{ fontSize: 12, color: C.midGray }}>{prodottoLabel(p)}</span>
                        </div>
                        {isOpen(pKey) && (p.varianti || []).filter((v) => !searching || matchesTerms(terms, prodottoLabel(p)) || varMatch(p, v)).map((v) => {
                          const priceOptions = [
                            { id: 'listino', name: 'Prezzo di listino', prezzoListino: v.listinoPrice, prezzoScontato: v.listinoPrice, prezzoCliente: v.listinoPrice },
                            ...(v.fornitori || []),
                          ];
                          return (
                            <div key={v.id} style={{ paddingLeft: 12, marginBottom: 6 }}>
                              {v.nome && <p style={{ fontSize: 10.5, fontWeight: 600, color: C.midGray, margin: '2px 0 4px' }}>{v.nome}</p>}
                              {priceOptions.map((opt) => (
                                <div key={opt.id} style={{ display: 'flex', alignItems: 'center', gap: 8, border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '6px 8px', marginBottom: 4, background: C.white }}>
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <p style={{ margin: 0, fontSize: 11, fontWeight: 600, color: opt.id === 'listino' ? C.gray : C.black }}>{opt.name}</p>
                                    <p style={{ margin: '2px 0 0', fontSize: 10.5, color: C.gray }}>impresa {opt.prezzoScontato} € · cliente {opt.prezzoCliente} €</p>
                                  </div>
                                  <button
                                    onClick={() => onAdd(p, v, opt, m.name, c.name)}
                                    title="Aggiungi al computo"
                                    aria-label="Aggiungi al computo"
                                    style={{ flexShrink: 0, width: 26, height: 26, borderRadius: 999, border: `1px solid ${C.paleGray}`, background: C.white, color: C.maroon, fontSize: 14, fontWeight: 700, lineHeight: 1, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                                  >
                                    +
                                  </button>
                                </div>
                              ))}
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
      })}
    </div>
  );
}

// Raggruppa le voci di una revisione su due livelli (macrocategoria > categoria), con le voci elencate
// direttamente sotto la loro categoria (senza un'ulteriore intestazione di sottocategoria: nel listino la
// sottocategoria è spesso quasi identica alla singola voce, quindi mostrarla come riga a parte era solo
// rumore visivo — resta comunque salvata su ogni voce, non mostrata). Rispetta l'ordine personalizzato
// salvato (macroOrder / categorieOrder) e i gruppi creati ma ancora vuoti (extraSections / categorieDefinite).
// Usata sia per la visualizzazione a schermo del computo sia per la stampa PDF e l'export Excel, così
// restano sempre coerenti tra loro.
function buildComputoGroups(revision) {
  const items = revision?.items || [];
  const extraSections = revision?.extraSections || [];
  const categorieOrderSaved = revision?.categorieOrder || {};
  const categorieDefinedEmpty = revision?.categorieDefinite || {};

  const rawMacroNames = [];
  extraSections.forEach((name) => { if (!rawMacroNames.includes(name)) rawMacroNames.push(name); });
  items.forEach((it) => {
    const m = it.section || it.macro || 'Voci varie';
    if (!rawMacroNames.includes(m)) rawMacroNames.push(m);
  });
  const macroNames = orderNames(rawMacroNames, revision?.macroOrder);

  return macroNames.map((name, idx) => {
    const sectionItems = items.filter((it) => (it.section || it.macro || 'Voci varie') === name);
    const realSectionItems = sectionItems.filter((it) => it.type !== 'subtotal');
    const subtotalMarkers = sectionItems.filter((it) => it.type === 'subtotal');

    const rawCatNames = [];
    (categorieDefinedEmpty[name] || []).forEach((c) => { if (!rawCatNames.includes(c)) rawCatNames.push(c); });
    realSectionItems.forEach((it) => {
      const c = it.categoria || 'Generale';
      if (!rawCatNames.includes(c)) rawCatNames.push(c);
    });
    const catNames = orderNames(rawCatNames, categorieOrderSaved[name]);

    const categorie = catNames.map((catName) => ({
      name: catName,
      items: realSectionItems.filter((it) => (it.categoria || 'Generale') === catName),
    }));

    return { name, color: SECTION_COLORS[idx % SECTION_COLORS.length], items: sectionItems, categorie, subtotalMarkers };
  });
}

// Propone una lavorazione di cronoprogramma per ogni categoria del computo (stessa granularità e stessi
// colori di "Stati avanzamento pagamenti", che raggruppa invece per macrosezione): "Murature", "Tinteggiatura"
// ecc. invece di una riga per singola voce, troppo minuta per una schedulazione lavori. Ogni proposta porta con
// sé i codici delle voci di quella categoria, così la lavorazione generata resta collegata al computo.
function buildCronoSuggestions(revision) {
  const groups = buildComputoGroups(revision);
  const suggestions = [];
  groups.forEach((g) => {
    g.categorie.forEach((cat) => {
      if (cat.items.length === 0) return;
      suggestions.push({ name: cat.name, color: g.color, linkedCodes: cat.items.map((it) => it.code).filter(Boolean) });
    });
  });
  return suggestions;
}

// Esportazione Excel del computo, impaginata come un documento: intestazione, tabella con bordi, righe di
// sezione e di categoria evidenziate, totali per sezione e totale generale, note iniziale/finale, impostazioni
// di stampa A4 (adatta alla larghezza, intestazione di colonna ripetuta su ogni pagina, numero di pagina).
// mode: false = completo (impresa + cliente), true = solo cliente, 'muto' = senza prezzi (da far compilare
// a un'impresa: prezzo unitario vuoto e importi già come formule Quantità × Prezzo, con somme per sezione).
// ExcelJS si carica solo al momento dell'esportazione, così non appesantisce l'avvio dell'app.
async function exportComputoExcel(project, revision, mode) {
  if (mode === 'riepilogo') return exportComputoRiepilogoExcel(project, revision);
  const ExcelJS = (await import('exceljs')).default;
  const muto = mode === 'muto';
  const clientOnly = !!mode;
  const groups = buildComputoGroups(revision);
  const econ = computeEconomics(revision);
  const realItems = (revision.items || []).filter((it) => it.type !== 'subtotal');

  const INK = 'FF2B2321';
  const ACC = 'FF862D2D';
  const GRAY = 'FF8B8178';
  const LINE = 'FFDCD5C9';
  const EURO = '#,##0.00 "€"';
  const QTY = '#,##0.00';
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: LINE } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };

  // Colonne: le coppie prezzo/totale dipendono dalla versione
  const pairs = muto
    ? [{ key: 'm', price: 'Prezzo unitario', total: 'Importo', color: INK }]
    : clientOnly
      ? [{ key: 'c', price: 'Prezzo cliente', total: 'Totale cliente', color: ACC }]
      : [{ key: 'i', price: 'Prezzo impresa', total: 'Totale impresa', color: INK }, { key: 'c', price: 'Prezzo cliente', total: 'Totale cliente', color: ACC }];
  const columns = [
    { header: 'Codice', width: 13 },
    { header: 'Descrizione', width: 62 },
    { header: 'Q.tà', width: 10 },
    { header: 'U.M.', width: 10 },
  ];
  pairs.forEach((p) => { p.priceCol = columns.length + 1; columns.push({ header: p.price, width: 14 }); p.totCol = columns.length + 1; columns.push({ header: p.total, width: 16 }); });
  const noteCol = columns.length + 1;
  columns.push({ header: 'Note', width: 30 });
  const lastCol = columns.length;
  const colLetter = (n) => String.fromCharCode(64 + n);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Quant';
  const ws = wb.addWorksheet('Computo', {
    views: [{ showGridLines: false }],
    pageSetup: { paperSize: 9, orientation: clientOnly ? 'portrait' : 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.2, footer: 0.25 }, horizontalCentered: true },
  });
  ws.columns = columns.map((c) => ({ width: c.width }));
  ws.headerFooter.oddFooter = `&L&8${String(project.name || '').replace(/&/g, '&&')}&R&8Pagina &P di &N`;

  // Altezza riga stimata dal testo (Excel non ricalcola da solo l'altezza delle righe a capo all'apertura)
  const linesFor = (text, width) => String(text || '').split('\n').reduce((n, part) => n + Math.max(1, Math.ceil(part.length / (width * 1.2))), 0);
  const fitHeight = (row, cells) => {
    const lines = Math.max(1, ...cells.map(([text, width]) => linesFor(text, width)));
    row.height = Math.max(16, lines * 12.5 + 5);
  };
  const mergedText = (text, opts = {}) => {
    const row = ws.addRow([]);
    ws.mergeCells(row.number, 1, row.number, lastCol);
    const cell = row.getCell(1);
    cell.value = text;
    cell.alignment = { wrapText: true, vertical: 'top' };
    cell.font = { size: 10, color: { argb: INK }, ...(opts.font || {}) };
    if (opts.fill) cell.fill = fill(opts.fill);
    if (opts.border) cell.border = opts.border;
    const totalWidth = columns.reduce((n, c) => n + c.width, 0);
    row.height = opts.height || Math.max(16, linesFor(text, totalWidth) * 13 + 6);
    return row;
  };

  // --- Intestazione del documento ---
  const title = muto ? 'Computo metrico — senza prezzi' : clientOnly ? 'Computo metrico — versione cliente' : 'Computo metrico';
  mergedText(title, { font: { size: 16, bold: true, color: { argb: INK } }, height: 26 });
  const meta = [['Progetto', project.name]];
  if (!muto && project.client) meta.push(['Cliente', project.client]);
  if (project.header?.ubicazione) meta.push(['Ubicazione', project.header.ubicazione]);
  meta.push(['Versione', revision.customName || revision.label]);
  meta.push(['Data', revision.dateModified]);
  meta.forEach(([k, v]) => {
    const row = ws.addRow([k, v || '']);
    ws.mergeCells(row.number, 2, row.number, lastCol);
    row.getCell(1).font = { size: 9, color: { argb: GRAY } };
    row.getCell(2).font = { size: 10, color: { argb: INK } };
  });
  ws.addRow([]);
  if ((revision.noteInizio || '').trim()) {
    mergedText(revision.noteInizio.trim(), { fill: 'FFFAF8F4', border: { left: { style: 'thick', color: { argb: ACC } } } });
    ws.addRow([]);
  }

  // --- Intestazione di colonna (ripetuta in stampa su ogni pagina) ---
  const head = ws.addRow(columns.map((c) => c.header.toUpperCase()));
  head.height = 24;
  head.eachCell((cell, n) => {
    cell.font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    cell.fill = fill(INK);
    cell.border = box;
    cell.alignment = { vertical: 'middle', horizontal: n >= 3 && n !== 4 && n !== noteCol ? 'right' : 'left', wrapText: true };
  });
  ws.pageSetup.printTitlesRow = `${head.number}:${head.number}`;
  ws.views = [{ state: 'frozen', ySplit: head.number, showGridLines: false }];

  const totalsRow = (label, values, opts = {}) => {
    // values: { [colIndex]: numero | { formula, result } }
    const row = ws.addRow([]);
    ws.mergeCells(row.number, 1, row.number, 4);
    const lab = row.getCell(1);
    lab.value = label;
    lab.alignment = { horizontal: 'right', vertical: 'middle' };
    lab.font = { size: 10, bold: !!opts.bold, color: { argb: opts.color || INK } };
    for (let c = 1; c <= lastCol; c++) {
      const cell = row.getCell(c);
      if (opts.fill) cell.fill = fill(opts.fill);
      if (opts.topLine) cell.border = { top: { style: opts.topLine, color: { argb: opts.color || INK } } };
    }
    Object.entries(values).forEach(([col, v]) => {
      const cell = row.getCell(Number(col));
      cell.value = v;
      cell.numFmt = EURO;
      cell.font = { size: 10, bold: !!opts.bold, color: { argb: opts.color || INK } };
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    });
    row.height = opts.height || 17;
    return row;
  };

  const sectionTotalCells = []; // muto: celle con il totale di ciascuna sezione

  groups.forEach((g) => {
    const items = g.categorie.reduce((n, c) => n + c.items.length, 0);
    if (!items && !g.subtotalMarkers.length) return;
    ws.addRow([]).height = 8;
    const sec = mergedText(String(g.name || '').toUpperCase(), { font: { bold: true, size: 11, color: { argb: ACC } }, fill: 'FFF4EAEA', height: 20 });
    sec.getCell(1).alignment = { vertical: 'middle' };
    sec.getCell(1).border = { top: { style: 'medium', color: { argb: ACC } } };

    const itemRowNums = [];
    let runI = 0;
    let runC = 0;
    g.categorie.forEach((cat) => {
      if (!cat.items.length) return;
      const cr = mergedText(cat.name, { font: { bold: true, size: 10, color: { argb: INK } }, fill: 'FFF6F3EE', height: 18 });
      cr.getCell(1).alignment = { vertical: 'middle' };
      cat.items.forEach((it) => {
        const qty = parseEuro(it.qty);
        const pI = parseEuro(it.unitPriceImpresa);
        const pC = parseEuro(it.unitPriceCliente);
        runI += pI * qty;
        runC += pC * qty;
        const notes = [];
        if (it.qtyLink) {
          const src = realItems.find((x) => String(x.id) === String(it.qtyLink.sourceId));
          if (src) notes.push(`Quantità: vedi voce ${src.code}${Number(it.qtyLink.factor) !== 1 ? ` × ${String(it.qtyLink.factor).replace('.', ',')}` : ''}`);
        }
        if (it.noteCliente) notes.push(it.noteCliente);
        const row = ws.addRow([it.code || '', it.desc || '', qty, it.unit || '']);
        itemRowNums.push(row.number);
        pairs.forEach((p) => {
          const priceCell = row.getCell(p.priceCol);
          const totCell = row.getCell(p.totCol);
          if (p.key === 'm') {
            priceCell.value = null;
            priceCell.fill = fill('FFFFFBEA');
            const r = row.number;
            totCell.value = { formula: `IF(${colLetter(p.priceCol)}${r}="","",C${r}*${colLetter(p.priceCol)}${r})`, result: '' };
          } else {
            const price = p.key === 'i' ? pI : pC;
            priceCell.value = price;
            totCell.value = price * qty;
          }
          priceCell.numFmt = EURO;
          totCell.numFmt = EURO;
          priceCell.font = { size: 10, color: { argb: p.color } };
          totCell.font = { size: 10, bold: true, color: { argb: p.color } };
        });
        row.getCell(noteCol).value = notes.join('\n');
        row.eachCell({ includeEmpty: true }, (cell, n) => {
          if (n > lastCol) return;
          cell.border = box;
          const right = n === 3 || (n > 4 && n !== noteCol);
          cell.alignment = { vertical: 'top', wrapText: n === 2 || n === noteCol, horizontal: right ? 'right' : 'left' };
          if (!cell.font) cell.font = { size: 10, color: { argb: INK } };
        });
        row.getCell(1).font = { size: 9, color: { argb: GRAY } };
        row.getCell(2).font = { size: 10, color: { argb: INK } };
        row.getCell(3).numFmt = QTY;
        row.getCell(3).font = { size: 10, color: { argb: INK } };
        row.getCell(4).font = { size: 10, color: { argb: GRAY } };
        row.getCell(noteCol).font = { size: 9, italic: true, color: { argb: GRAY } };
        fitHeight(row, [[it.desc, columns[1].width], [notes.join('\n'), columns[noteCol - 1].width]]);
      });
    });

    // Sommatorie inserite a mano nel computo (con eventuale sconto/IVA propri)
    g.subtotalMarkers.forEach((it) => {
      if (muto) { totalsRow(it.title, {}, { bold: true, fill: 'FFFAF8F4' }); return; }
      const vals = {};
      const lines = [];
      pairs.forEach((p) => {
        const gross = p.key === 'i' ? runI : runC;
        const net = it.discountPct ? gross * (1 - it.discountPct / 100) : gross;
        vals[p.totCol] = net;
        lines.push([p, gross, net]);
      });
      totalsRow(it.discountPct ? `${it.title} (netto, sconto ${it.discountPct}%)` : it.title, vals, { bold: true, fill: 'FFFAF8F4' });
      if (it.vatRate !== null && it.vatRate !== undefined) {
        const iva = {};
        const tot = {};
        lines.forEach(([p, , net]) => { iva[p.totCol] = net * (it.vatRate / 100); tot[p.totCol] = net * (1 + it.vatRate / 100); });
        totalsRow(`${it.title} — ${it.vatLabel}`, iva, {});
        totalsRow(`${it.title} — IVA inclusa`, tot, { bold: true });
      }
      runI = 0;
      runC = 0;
    });

    // Totali di sezione
    if (muto) {
      if (itemRowNums.length) {
        const p = pairs[0];
        const L = colLetter(p.totCol);
        const r = totalsRow(`Totale ${g.name} IVA esclusa`, { [p.totCol]: { formula: `SUM(${L}${itemRowNums[0]}:${L}${itemRowNums[itemRowNums.length - 1]})`, result: 0 } }, { bold: true, topLine: 'thin', fill: 'FFFAF8F4' });
        sectionTotalCells.push(`${L}${r.number}`);
      }
      return;
    }
    const s = econ.sections.find((x) => x.name === g.name);
    if (!s || (!s.grossImpresa && !s.grossCliente)) return;
    const pick = (kI, kC) => Object.fromEntries(pairs.map((p) => [p.totCol, p.key === 'i' ? s[kI] : s[kC]]));
    if (s.discountPct > 0) {
      totalsRow(`Subtotale ${g.name}`, pick('grossImpresa', 'grossCliente'), { topLine: 'thin' });
      totalsRow(`Sconto ${s.discountPct}%`, Object.fromEntries(pairs.map((p) => [p.totCol, -(p.key === 'i' ? s.grossImpresa - s.netImpresa : s.grossCliente - s.netCliente)])), {});
    }
    totalsRow(`Totale ${g.name} IVA esclusa`, pick('netImpresa', 'netCliente'), { bold: true, topLine: s.discountPct > 0 ? undefined : 'thin', fill: 'FFFAF8F4' });
    totalsRow(s.vat.label, pick('ivaImpresa', 'ivaCliente'), {});
    totalsRow(`Totale ${g.name} IVA inclusa`, pick('totImpresa', 'totCliente'), { bold: true });
  });

  // --- Totale generale ---
  ws.addRow([]).height = 10;
  if (muto) {
    const p = pairs[0];
    totalsRow('Totale generale IVA esclusa', sectionTotalCells.length ? { [p.totCol]: { formula: sectionTotalCells.join('+'), result: 0 } } : {}, { bold: true, topLine: 'medium', fill: 'FFF4EAEA', height: 22 });
  } else {
    const pick = (kI, kC) => Object.fromEntries(pairs.map((p) => [p.totCol, p.key === 'i' ? econ[kI] : econ[kC]]));
    totalsRow('Totale generale IVA esclusa', pick('netImpresa', 'netCliente'), { bold: true, topLine: 'medium', fill: clientOnly ? 'FFF4EAEA' : 'FFFAF8F4', color: clientOnly ? ACC : INK, height: 22 });
    econ.breakdown.forEach((b) => totalsRow(b.label, Object.fromEntries(pairs.map((p) => [p.totCol, p.key === 'i' ? b.ivaImpresa : b.ivaCliente])), {}));
    totalsRow('Totale generale IVA inclusa', pick('totImpresa', 'totCliente'), { bold: !clientOnly, topLine: 'thin', height: 20 });
  }

  if ((revision.noteFine || '').trim()) {
    ws.addRow([]);
    mergedText('NOTE', { font: { bold: true, size: 9, color: { argb: GRAY } }, height: 16 });
    mergedText(revision.noteFine.trim(), {});
  }

  // Font uniforme (Arial) e un piccolo rientro nelle celle, così il testo non tocca i bordi
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { ...(cell.font || {}), name: 'Arial' };
      const al = cell.alignment || {};
      if (row.number >= head.number && al.horizontal !== 'center') cell.alignment = { ...al, indent: 1 };
    });
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const suffix = muto ? '_senza_prezzi' : clientOnly ? '_cliente' : '';
  const filename = `${String(project.name || 'computo').replace(/[^a-z0-9]+/gi, '_')}_${String(revision.customName || revision.label || '').replace(/[^a-z0-9]+/gi, '_')}${suffix}.xlsx`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

const exportComputoExcelSenzaPrezzi = (project, revision) => exportComputoExcel(project, revision, 'muto');

// Versione semplificata per il cliente: un foglio con sezioni, categorie (numero voci e importo), totali di
// sezione con IVA e totale generale. Nessuna singola voce.
async function exportComputoRiepilogoExcel(project, revision) {
  const ExcelJS = (await import('exceljs')).default;
  const groups = buildComputoGroups(revision);
  const econ = computeEconomics(revision);
  const INK = 'FF2B2321';
  const ACC = 'FF862D2D';
  const GRAY = 'FF8B8178';
  const EURO = '#,##0.00 "€"';
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const line = { style: 'thin', color: { argb: 'FFDCD5C9' } };
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Quant';
  const ws = wb.addWorksheet('Riepilogo', {
    views: [{ showGridLines: false }],
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.2, footer: 0.25 }, horizontalCentered: true },
  });
  ws.columns = [{ width: 58 }, { width: 10 }, { width: 18 }];
  ws.headerFooter.oddFooter = `&L&8${String(project.name || '').replace(/&/g, '&&')}&R&8Pagina &P di &N`;
  const merged = (text, font, height, fillArgb) => {
    const row = ws.addRow([text]);
    ws.mergeCells(row.number, 1, row.number, 3);
    row.getCell(1).font = { name: 'Arial', size: 10, color: { argb: INK }, ...font };
    row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
    if (fillArgb) row.getCell(1).fill = fill(fillArgb);
    row.height = height || Math.max(16, Math.ceil(String(text).length / 95) * 13 + 6);
    return row;
  };
  merged('Computo metrico — riepilogo', { size: 16, bold: true }, 26);
  [['Progetto', project.name], ['Cliente', project.client], ['Versione', revision.customName || revision.label], ['Data', revision.dateModified]].filter(([, v]) => v).forEach(([k, v]) => {
    const row = ws.addRow([`${k}: ${v}`]);
    ws.mergeCells(row.number, 1, row.number, 3);
    row.getCell(1).font = { name: 'Arial', size: 10, color: { argb: INK } };
  });
  ws.addRow([]);
  if ((revision.noteInizio || '').trim()) { merged(revision.noteInizio.trim(), {}, null, 'FFFAF8F4'); ws.addRow([]); }

  const head = ws.addRow(['SEZIONE / CATEGORIA', 'N. VOCI', 'IMPORTO']);
  head.height = 22;
  head.eachCell((c, n) => {
    c.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    c.fill = fill(INK);
    c.alignment = { vertical: 'middle', horizontal: n === 1 ? 'left' : 'right', indent: 1 };
  });
  ws.pageSetup.printTitlesRow = `${head.number}:${head.number}`;
  ws.views = [{ state: 'frozen', ySplit: head.number, showGridLines: false }];

  const totalRow = (label, value, opts = {}) => {
    const row = ws.addRow([label, '', value]);
    ws.mergeCells(row.number, 1, row.number, 2);
    const f = { name: 'Arial', size: opts.size || 10, bold: !!opts.bold, color: { argb: opts.color || INK } };
    row.getCell(1).font = f;
    row.getCell(1).alignment = { horizontal: 'right', indent: 1 };
    row.getCell(3).font = f;
    row.getCell(3).numFmt = EURO;
    row.getCell(3).alignment = { horizontal: 'right', indent: 1 };
    if (opts.fill) [1, 2, 3].forEach((n) => { row.getCell(n).fill = fill(opts.fill); });
    if (opts.top) [1, 2, 3].forEach((n) => { row.getCell(n).border = { top: { style: opts.top, color: { argb: opts.color || INK } } }; });
    row.height = opts.height || 17;
  };

  groups.forEach((g) => {
    const cats = g.categorie.filter((c) => c.items.length);
    if (!cats.length) return;
    ws.addRow([]).height = 8;
    const sec = merged(String(g.name).toUpperCase(), { bold: true, size: 11, color: { argb: ACC } }, 20, 'FFF4EAEA');
    sec.getCell(1).border = { top: { style: 'medium', color: { argb: ACC } } };
    sec.getCell(1).alignment = { vertical: 'middle', indent: 1 };
    cats.forEach((c) => {
      const row = ws.addRow([c.name, c.items.length, c.items.reduce((n, it) => n + parseEuro(it.unitPriceCliente) * parseEuro(it.qty), 0)]);
      row.getCell(1).font = { name: 'Arial', size: 10, color: { argb: INK } };
      row.getCell(1).alignment = { indent: 2, wrapText: true, vertical: 'top' };
      row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: GRAY } };
      row.getCell(2).alignment = { horizontal: 'right', indent: 1, vertical: 'top' };
      row.getCell(3).font = { name: 'Arial', size: 10, color: { argb: INK } };
      row.getCell(3).numFmt = EURO;
      row.getCell(3).alignment = { horizontal: 'right', indent: 1, vertical: 'top' };
      [1, 2, 3].forEach((n) => { row.getCell(n).border = { bottom: line }; });
      row.height = Math.max(17, Math.ceil(String(c.name).length / 60) * 13 + 4);
    });
    const s = econ.sections.find((x) => x.name === g.name);
    if (!s || (!s.grossCliente && !s.netCliente)) return;
    if (s.discountPct > 0) {
      totalRow(`Subtotale ${g.name}`, s.grossCliente, {});
      totalRow(`Sconto ${s.discountPct}%`, -(s.grossCliente - s.netCliente), {});
    }
    totalRow(`Totale ${g.name} IVA esclusa`, s.netCliente, { bold: true, fill: 'FFFAF8F4' });
    totalRow(s.vat.label, s.ivaCliente, { color: GRAY, size: 9 });
    totalRow(`Totale ${g.name} IVA inclusa`, s.totCliente, { color: GRAY, size: 9 });
  });

  ws.addRow([]).height = 10;
  totalRow('Totale generale IVA esclusa', econ.netCliente, { bold: true, size: 11, color: ACC, fill: 'FFF4EAEA', top: 'medium', height: 22 });
  econ.breakdown.forEach((b) => totalRow(b.label, b.ivaCliente, { color: GRAY }));
  totalRow('Totale generale IVA inclusa', econ.totCliente, { top: 'thin', height: 20 });
  if ((revision.noteFine || '').trim()) {
    ws.addRow([]);
    merged('NOTE', { bold: true, size: 9, color: { argb: GRAY } }, 16);
    merged(revision.noteFine.trim(), {});
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const filename = `${String(project.name || 'computo').replace(/[^a-z0-9]+/gi, '_')}_${String(revision.customName || revision.label || '').replace(/[^a-z0-9]+/gi, '_')}_riepilogo.xlsx`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function PrintableComputo({ project, revision, clientOnly: clientOnlyProp, studioSettings }) {
  // clientOnly può valere false (completo impresa+cliente), true (solo cliente) o 'muto' (senza prezzi: voci e
  // quantità con le colonne prezzo/importo vuote, da far compilare a un'impresa). Il muto usa l'impaginazione
  // della versione cliente, ma senza alcun importo.
  const noPrices = clientOnlyProp === 'muto';
  // 'riepilogo' = versione semplificata per il cliente: niente singole voci, solo sezioni, categorie e totali.
  const riepilogo = clientOnlyProp === 'riepilogo';
  const clientOnly = !!clientOnlyProp;
  const groups = buildComputoGroups(revision);
  const realItems = (revision.items || []).filter((it) => it.type !== 'subtotal');
  const header = project.header || {};
  const econ = computeEconomics(revision);
  const ss = studioSettings || DEFAULT_STUDIO_SETTINGS;
  const hasCustomHeader = ss.usaIntestazionePersonalizzata && ss.intestazioneImg;
  const hasCustomFooter = ss.usaPiePersonalizzato && ss.pieImg;
  const hasStudioInfo = ss.nome || ss.indirizzo || ss.piva || ss.cf || ss.telefono || ss.email || ss.sito || ss.logo;
  const headerActive = hasCustomHeader || hasStudioInfo;
  const footerActive = hasCustomFooter || !!ss.testoPiePagina;
  // Piccolo respiro interno per intestazione/piè/contenuto: il vero margine del foglio ora lo dà @page
  // (vedi sotto), quindi qui basta poco — non va sommato a un margine già presente, altrimenti il bordo
  // finale risulterebbe doppio.
  const PAGE_SIDE = 4;

  // Font personalizzato dello studio (caricato in Impostazioni), incorporato via @font-face e usato al posto
  // del font di sistema solo nel documento stampato. L'estensione del file scelto determina il formato dichiarato
  // a @font-face (solo un suggerimento per il browser, non blocca il caricamento se sbagliato).
  const hasCustomFont = ss.usaFontPersonalizzato && ss.fontPersonalizzato;
  const fontExt = (ss.fontPersonalizzatoNome || '').split('.').pop().toLowerCase();
  const fontFormat = { woff2: 'woff2', woff: 'woff', ttf: 'truetype', otf: 'opentype' }[fontExt] || 'woff2';
  const presetFont = (PDF_FONT_PRESETS[ss.pdfFontPreset] || PDF_FONT_PRESETS.inter).css;
  const effectiveFont = hasCustomFont ? `'StudioFontPersonalizzato', ${presetFont}` : presetFont;
  const fontFaceCss = hasCustomFont
    // font-display:swap (non "block"): mostra subito il testo con il font di riserva e passa al font
    // personalizzato appena pronto. Con "block" il testo resta INVISIBILE finché il font non è pronto (fino
    // a qualche secondo, per specifica) — troppo per una stampa che parte dopo poche centinaia di millisecondi,
    // col rischio concreto di un PDF con pagine bianche anche se il font è incorporato come data URL.
    ? `@font-face { font-family: 'StudioFontPersonalizzato'; src: url("${ss.fontPersonalizzato}") format("${fontFormat}"); font-display: swap; }`
    : '';

  // Intestazione e piè di pagina sono le <thead>/<tfoot> di un'unica tabella che avvolge tutto il documento:
  // è lo stesso meccanismo, nativo dei browser e già usato per le intestazioni di colonna di ogni tabella di
  // voci più sotto, che ripete <thead>/<tfoot> in cima e in fondo a OGNI pagina stampata quando una tabella
  // supera l'altezza di una pagina — a differenza di un'intestazione "position: fixed", che nei test si è
  // rivelata inaffidabile nella generazione del PDF (poteva comparire spostata, a cavallo tra una pagina e
  // l'altra). La regola @page (sotto, nel foglio di stile globale) imposta un vero margine di stampa su tutti
  // i lati: niente più tocca il bordo fisico del foglio, intestazione e piè inclusi — l'immagine personalizzata
  // di intestazione/piè riempie la larghezza disponibile DENTRO quel margine (non il foglio fisico), con una
  // scala regolabile dall'utente in Impostazioni. NOTA IMPORTANTE: la riga <tr> esterna che avvolge TUTTO il
  // contenuto (sotto) non deve MAI avere break-inside:avoid — essendo alta quanto l'intero documento, forzare
  // il browser a "non spezzarla" produce un'impaginazione rotta (righe tagliate, pagine confuse). L'unica
  // protezione da interruzione va messa sulle singole righe piccole (voci, intestazioni di categoria, totali).
  const headerScale = ss.intestazioneScala || 100;
  const footerScale = ss.pieScala || 100;
  return (
    <div className="print-only" style={{ fontFamily: effectiveFont, color: '#1A1A1A' }}>
      {fontFaceCss && <style>{fontFaceCss}</style>}
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        {headerActive && (
          <thead>
            <tr><td style={{ padding: 0 }}>
              {hasCustomHeader ? (
                <div style={{ display: 'flex', justifyContent: 'center', padding: `10px ${PAGE_SIDE}px`, borderBottom: '1px solid #E5E2DA', background: '#fff' }}>
                  <img src={ss.intestazioneImg} alt="" style={{ width: `${headerScale}%`, maxWidth: '100%', display: 'block' }} />
                </div>
              ) : (
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 14,
                  padding: `16px ${PAGE_SIDE}px`, borderBottom: '1px solid #E5E2DA', background: '#fff',
                }}>
                  {ss.logo && <img src={ss.logo} alt="" style={{ height: 42 }} />}
                  <div>
                    {ss.nome && <div style={{ fontSize: 15, fontWeight: 600, letterSpacing: 0.2, color: '#1A1A1A' }}>{ss.nome}</div>}
                    {(ss.indirizzo || ss.piva || ss.cf || ss.telefono || ss.email || ss.sito) && (
                      <div style={{ fontSize: 9, color: '#8A8A8A', marginTop: 2 }}>
                        {[ss.indirizzo, ss.piva && `P.IVA ${ss.piva}`, ss.telefono, ss.email].filter(Boolean).join('  ·  ')}
                      </div>
                    )}
                  </div>
                </div>
              )}
            </td></tr>
          </thead>
        )}
        {footerActive && (
          <tfoot>
            <tr><td style={{ padding: 0 }}>
              {hasCustomFooter ? (
                <div style={{ display: 'flex', justifyContent: 'center', padding: `8px ${PAGE_SIDE}px`, borderTop: '1px solid #E5E2DA', background: '#fff' }}>
                  <img src={ss.pieImg} alt="" style={{ width: `${footerScale}%`, maxWidth: '100%', display: 'block' }} />
                </div>
              ) : (
                <div style={{
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  padding: `8px ${PAGE_SIDE}px`, borderTop: '1px solid #E5E2DA', background: '#fff',
                }}>
                  <p style={{ fontSize: 8.5, color: '#8A8A8A', margin: 0, letterSpacing: 0.3, textAlign: 'center' }}>{ss.testoPiePagina}</p>
                </div>
              )}
            </td></tr>
          </tfoot>
        )}
        <tbody>
          <tr><td style={{ padding: `16px ${PAGE_SIDE}px` }}>

      <h1 style={{ fontSize: 18, fontWeight: 600, margin: '0 0 8px', color: '#1A1A1A' }}>{noPrices ? 'Computo metrico — senza prezzi' : riepilogo ? 'Computo metrico — riepilogo' : clientOnly ? 'Computo metrico — versione cliente' : 'Computo metrico'}</h1>
      <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>Progetto: {project.name}{!noPrices && <> — Cliente: {project.client}</>}</p>
      <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>Versione: {revision.customName || revision.label} — Modificata il {revision.dateModified}</p>
      {(header.descrizione || header.ubicazione) && (
        <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>{header.descrizione} {header.ubicazione && `— ${header.ubicazione}`}</p>
      )}
      {/* Nel computo senza prezzi (destinato a un'impresa) i recapiti del cliente non compaiono. */}
      {project.clientSheet && !noPrices && (
        <div style={{ marginTop: 14, padding: '12px 14px', border: '1px solid #E5E2DA', borderRadius: 8, breakInside: 'avoid', pageBreakInside: 'avoid' }}>
          <p style={{ fontSize: 10, fontWeight: 600, margin: '0 0 5px', color: '#1A1A1A', textTransform: 'uppercase', letterSpacing: 0.5 }}>Scheda cliente</p>
          <p style={{ fontSize: 10.5, margin: '2px 0' }}>{project.clientSheet.name}{project.clientSheet.type && ` — ${project.clientSheet.type}`}</p>
          {project.clientSheet.address && <p style={{ fontSize: 10.5, margin: '2px 0' }}>{project.clientSheet.address}</p>}
          <p style={{ fontSize: 10.5, margin: '2px 0' }}>
            {[project.clientSheet.phone, project.clientSheet.email].filter(Boolean).join(' — ')}
          </p>
          {(project.clientSheet.piva || project.clientSheet.cf) && (
            <p style={{ fontSize: 10.5, margin: '2px 0' }}>
              {[project.clientSheet.piva && `P.IVA ${project.clientSheet.piva}`, project.clientSheet.cf && `CF ${project.clientSheet.cf}`].filter(Boolean).join(' — ')}
            </p>
          )}
        </div>
      )}
      {(revision.noteInizio || '').trim() && (
        <div style={{ marginTop: 16, padding: '10px 14px', borderLeft: `3px solid ${C.maroon}`, background: '#FAF8F4', fontSize: 10.5, lineHeight: 1.5, color: '#3A3A3A', whiteSpace: 'pre-wrap', breakInside: 'avoid', pageBreakInside: 'avoid' }}>
          {revision.noteInizio.trim()}
        </div>
      )}
      {riepilogo && (() => {
        const cell = { padding: '6px 6px', fontSize: 10.5, borderBottom: '1px solid #F0EEE8' };
        const num = { ...cell, textAlign: 'right', whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums' };
        const th = { textAlign: 'left', borderBottom: '1.5px solid #1A1A1A', padding: '4px 6px', fontSize: 8, textTransform: 'uppercase', letterSpacing: 0.5, color: '#8A8A8A', fontWeight: 600 };
        return (
          <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 24, tableLayout: 'fixed' }}>
            <colgroup><col /><col style={{ width: 70 }} /><col style={{ width: 110 }} /></colgroup>
            <thead><tr><th style={th}>Sezione / categoria</th><th style={{ ...th, textAlign: 'right' }}>N. voci</th><th style={{ ...th, textAlign: 'right' }}>Importo</th></tr></thead>
            <tbody>
              {groups.map((g, gi) => {
                const s = econ.sections.find((x) => x.name === g.name);
                const cats = g.categorie.filter((c) => c.items.length);
                if (!cats.length) return null;
                return (
                  <React.Fragment key={gi}>
                    <tr style={{ breakInside: 'avoid', pageBreakInside: 'avoid', breakAfter: 'avoid', pageBreakAfter: 'avoid' }}>
                      <td colSpan={3} style={{ padding: '14px 6px 6px', fontSize: 11.5, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.6, color: '#1A1A1A', borderBottom: `1px solid ${C.maroon}` }}>
                        <span style={{ borderLeft: `3px solid ${C.maroon}`, paddingLeft: 8 }}>{g.name}</span>
                      </td>
                    </tr>
                    {cats.map((c, ci) => (
                      <tr key={ci} style={{ breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                        <td style={{ ...cell, paddingLeft: 18 }}>{c.name}</td>
                        <td style={{ ...num, color: '#8A8A8A' }}>{c.items.length}</td>
                        <td style={num}>{formatEuro(c.items.reduce((n, it) => n + parseEuro(it.unitPriceCliente) * parseEuro(it.qty), 0))}</td>
                      </tr>
                    ))}
                    {s && (s.grossCliente || s.netCliente) ? (
                      <>
                        {s.discountPct > 0 && (
                          <>
                            <tr><td colSpan={2} style={{ ...num, borderBottom: 'none', color: '#5A5A5A' }}>Subtotale {g.name}</td><td style={{ ...num, borderBottom: 'none', color: '#5A5A5A' }}>{formatEuro(s.grossCliente)}</td></tr>
                            <tr><td colSpan={2} style={{ ...num, borderBottom: 'none', color: '#5A5A5A' }}>Sconto {s.discountPct}%</td><td style={{ ...num, borderBottom: 'none', color: '#5A5A5A' }}>−{formatEuro(s.grossCliente - s.netCliente)}</td></tr>
                          </>
                        )}
                        <tr style={{ breakInside: 'avoid', pageBreakInside: 'avoid' }}><td colSpan={2} style={{ ...num, borderBottom: 'none', fontWeight: 700 }}>Totale {g.name} IVA esclusa</td><td style={{ ...num, borderBottom: 'none', fontWeight: 700 }}>{formatEuro(s.netCliente)}</td></tr>
                        <tr><td colSpan={2} style={{ ...num, borderBottom: 'none', color: '#5A5A5A', fontSize: 9.5 }}>{s.vat.label}</td><td style={{ ...num, borderBottom: 'none', color: '#5A5A5A', fontSize: 9.5 }}>{formatEuro(s.ivaCliente)}</td></tr>
                        <tr><td colSpan={2} style={{ ...num, borderBottom: 'none', color: '#5A5A5A', fontSize: 9.5 }}>Totale {g.name} IVA inclusa</td><td style={{ ...num, borderBottom: 'none', color: '#5A5A5A', fontSize: 9.5 }}>{formatEuro(s.totCliente)}</td></tr>
                      </>
                    ) : null}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        );
      })()}
      {!riepilogo && groups.map((g, gi) => {
        let runI = 0;
        let runC = 0;
        const thStyle = { textAlign: 'left', borderBottom: '1.5px solid #1A1A1A', padding: '4px 4px', fontSize: 8, textTransform: 'uppercase', letterSpacing: 0.5, color: '#8A8A8A', fontWeight: 600, verticalAlign: 'bottom' };
        // Celle numeriche: allineate a destra, mai a capo, cifre a larghezza fissa (tabular-nums) così i
        // valori si incolonnano esattamente uno sotto l'altro.
        const numCell = { padding: '5px 4px', textAlign: 'right', whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums' };
        return (
        <div key={gi} style={{ marginTop: 24 }}>
          <h3 style={{
            fontSize: 11.5, fontWeight: 700, margin: '0 0 9px', padding: '1px 0 1px 10px',
            borderLeft: `3px solid ${C.maroon}`, textTransform: 'uppercase', letterSpacing: 0.6,
            color: '#1A1A1A', pageBreakAfter: 'avoid', breakAfter: 'avoid',
          }}>{g.name}</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 10, tableLayout: 'fixed' }}>
            {/* Larghezze di colonna fisse e identiche in tutte le sezioni (table-layout: fixed): così prezzi e
                totali restano incolonnati da una tabella all'altra e non vanno a capo (es. "€" sulla riga sotto). */}
            <colgroup>
              <col style={{ width: 64 }} />
              <col />
              <col style={{ width: 46 }} />
              <col style={{ width: 50 }} />
              {!clientOnly && <col style={{ width: 70 }} />}
              {!clientOnly && <col style={{ width: 80 }} />}
              <col style={{ width: 70 }} />
              <col style={{ width: 80 }} />
            </colgroup>
            <thead>
              <tr>
                <th style={thStyle}>Codice</th>
                <th style={thStyle}>Descrizione</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Qtà</th>
                <th style={thStyle}>U.M.</th>
                {!clientOnly && <th style={{ ...thStyle, textAlign: 'right' }}>Prezzo impresa</th>}
                {!clientOnly && <th style={{ ...thStyle, textAlign: 'right' }}>Totale impresa</th>}
                <th style={{ ...thStyle, textAlign: 'right' }}>{noPrices ? 'Prezzo unitario' : 'Prezzo cliente'}</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>{noPrices ? 'Importo' : 'Totale cliente'}</th>
              </tr>
            </thead>
            <tbody>
              {g.categorie.map((cat, ci) => (
                <React.Fragment key={ci}>
                  <tr style={{ breakInside: 'avoid', pageBreakInside: 'avoid', breakAfter: 'avoid', pageBreakAfter: 'avoid' }}>
                    <td colSpan={clientOnly ? 6 : 8} style={{ padding: '9px 4px 4px', fontWeight: 600, fontSize: 10, color: '#3A3A3A' }}>{cat.name}</td>
                  </tr>
                  {cat.items.map((it, ii) => {
                    runI += parseEuro(it.unitPriceImpresa) * parseEuro(it.qty);
                    runC += parseEuro(it.unitPriceCliente) * parseEuro(it.qty);
                    return (
                    <tr key={ii} style={{ breakInside: 'avoid', pageBreakInside: 'avoid', borderBottom: '1px solid #F0EEE8' }}>
                      <td style={{ padding: '5px 4px', color: '#5A5A5A', fontSize: 8.5, overflowWrap: 'anywhere' }}>{it.code}</td>
                      <td style={{ padding: '5px 4px' }}>
                        {it.desc}
                        {it.qtyLink && (() => {
                          const src = realItems.find((x) => String(x.id) === String(it.qtyLink.sourceId));
                          return src ? <div style={{ fontSize: 8.5, color: '#8A8A8A', marginTop: 2 }}>Quantità: vedi voce {src.code}{Number(it.qtyLink.factor) !== 1 ? ` × ${String(it.qtyLink.factor).replace('.', ',')}` : ''}</div> : null;
                        })()}
                        {it.noteCliente && <div style={{ fontSize: 8.5, color: '#5A5A5A', fontStyle: 'italic', marginTop: 2 }}>Nota: {it.noteCliente}</div>}
                      </td>
                      <td style={{ ...numCell }}>{itNum(parseEuro(it.qty), parseEuro(it.qty) % 1 ? 2 : 0)}</td>
                      <td style={{ padding: '5px 4px', color: '#8A8A8A', whiteSpace: 'nowrap' }}>{it.unit}</td>
                      {!clientOnly && <td style={{ ...numCell, color: '#5A5A5A' }}>{formatEuro(parseEuro(it.unitPriceImpresa))}</td>}
                      {!clientOnly && <td style={{ ...numCell, fontWeight: 600 }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>}
                      {noPrices ? (
                        <>
                          <td style={{ ...numCell }}><div style={{ borderBottom: '1px solid #C9C4BA', height: 12, marginLeft: 8 }} /></td>
                          <td style={{ ...numCell }}><div style={{ borderBottom: '1px solid #C9C4BA', height: 12, marginLeft: 8 }} /></td>
                        </>
                      ) : (
                        <>
                          <td style={{ ...numCell, color: C.maroon }}>{formatEuro(parseEuro(it.unitPriceCliente))}</td>
                          <td style={{ ...numCell, fontWeight: 600, color: C.maroon }}>{formatEuro(parseEuro(it.unitPriceCliente) * parseEuro(it.qty))}</td>
                        </>
                      )}
                    </tr>
                    );
                  })}
                </React.Fragment>
              ))}
              {g.subtotalMarkers.map((it, ii) => {
                const hasVat = it.vatRate !== null && it.vatRate !== undefined;
                const hasDiscount = !!it.discountPct;
                // Lo sconto si applica sull'importo lordo di questa sommatoria, prima dell'IVA — e la riga
                // "Sconto" compare in stampa solo se è stato davvero impostato (nessuna riga a sconto 0/assente).
                const grossAmt = clientOnly ? runC : runI;
                const netAmt = hasDiscount ? grossAmt * (1 - it.discountPct / 100) : grossAmt;
                const rowEl = (
                  <tr key={ii} style={{ breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                    <td colSpan={clientOnly ? 6 : 8} style={{ padding: '8px 4px', fontWeight: 700, borderTop: '1px solid #E5E2DA' }}>
                      {it.title}
                      {hasDiscount && !noPrices && (
                        <span style={{ fontWeight: 400, color: '#8A8A8A' }}>
                          {' '}— lordo {formatEuro(grossAmt)}, Sconto {it.discountPct}% (−{formatEuro(grossAmt - netAmt)}), netto {formatEuro(netAmt)}
                        </span>
                      )}
                      {hasVat && !noPrices && (
                        <span style={{ fontWeight: 400, color: '#8A8A8A' }}>
                          {' '}— IVA esclusa {formatEuro(netAmt)}, {it.vatLabel} {formatEuro(netAmt * (it.vatRate / 100))}, IVA inclusa {formatEuro(netAmt * (1 + it.vatRate / 100))}
                        </span>
                      )}
                    </td>
                  </tr>
                );
                runI = 0; runC = 0;
                return rowEl;
              })}
            </tbody>
          </table>
          {(() => {
            if (noPrices) {
              return (
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 10, breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                  <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, width: 320, maxWidth: '100%', fontSize: 9.5, fontWeight: 700, color: '#1A1A1A' }}>
                    <span style={{ whiteSpace: 'nowrap' }}>Totale {g.name} IVA esclusa</span>
                    <div style={{ flex: 1, borderBottom: '1px solid #C9C4BA', height: 12 }} />
                  </div>
                </div>
              );
            }
            const s = econ.sections.find((x) => x.name === g.name);
            if (!s || (!s.grossImpresa && !s.grossCliente)) return null;
            const line = { fontSize: 9.5, display: 'flex', justifyContent: 'space-between', gap: 12, color: '#5A5A5A', margin: '2px 0' };
            const block = (who, gross, net, iva, tot, color) => (
              <div style={{ flex: 1, minWidth: 0 }}>
                {s.discountPct > 0 && (
                  <>
                    <p style={line}><span>Subtotale{who}</span><span>{formatEuro(gross)}</span></p>
                    <p style={line}><span>Sconto {s.discountPct}%</span><span>−{formatEuro(gross - net)}</span></p>
                  </>
                )}
                <p style={line}><span>Totale IVA esclusa{who}</span><span>{formatEuro(net)}</span></p>
                <p style={line}><span>{s.vat.label}</span><span>{formatEuro(iva)}</span></p>
                <p style={{ ...line, fontWeight: 700, color, borderTop: '1px solid #E5E2DA', paddingTop: 3, marginTop: 3 }}><span>Totale {g.name} IVA inclusa{who}</span><span>{formatEuro(tot)}</span></p>
              </div>
            );
            return (
              <div style={{ display: 'flex', gap: 24, justifyContent: 'flex-end', marginTop: 8, breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                <div style={{ display: 'flex', gap: 24, width: clientOnly ? 300 : 620, maxWidth: '100%' }}>
                  {!clientOnly && block(' (impresa)', s.grossImpresa, s.netImpresa, s.ivaImpresa, s.totImpresa, '#1A1A1A')}
                  {block(clientOnly ? '' : ' (cliente)', s.grossCliente, s.netCliente, s.ivaCliente, s.totCliente, C.maroon)}
                </div>
              </div>
            );
          })()}
        </div>
        );
      })}
      <div style={{ marginTop: 30, maxWidth: 320, marginLeft: 'auto', breakInside: 'avoid', pageBreakInside: 'avoid' }}>
        {noPrices && (
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, fontSize: 12, fontWeight: 700, color: '#1A1A1A', padding: '6px 0' }}>
            <span style={{ whiteSpace: 'nowrap' }}>Totale generale IVA esclusa</span>
            <div style={{ flex: 1, borderBottom: '1.5px solid #1A1A1A', height: 14 }} />
          </div>
        )}
        {!noPrices && !clientOnly && (
          <>
            <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>Totale generale IVA esclusa (impresa)</span><span>{formatEuro(econ.netImpresa)}</span></p>
            {econ.breakdown.map((b, i) => <p key={i} style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>{b.label}</span><span>{formatEuro(b.ivaImpresa)}</span></p>)}
            <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between', fontWeight: 700, margin: '6px 0 0', paddingTop: 6, borderTop: '1px solid #E5E2DA' }}><span>Totale generale IVA inclusa (impresa)</span><span>{formatEuro(econ.totImpresa)}</span></p>
            <div style={{ height: 16 }} />
          </>
        )}
        {noPrices ? null : clientOnly ? (
          <>
            {/* Versione cliente: in evidenza il totale generale IVA esclusa; IVA e totale IVA inclusa restano righe normali. */}
            <p style={{ fontSize: 13, display: 'flex', justifyContent: 'space-between', fontWeight: 700, color: C.maroon, margin: '0 0 8px', paddingBottom: 6, borderBottom: `1.5px solid ${C.maroon}` }}><span>Totale generale IVA esclusa</span><span>{formatEuro(econ.netCliente)}</span></p>
            {econ.breakdown.map((b, i) => <p key={i} style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>{b.label}</span><span>{formatEuro(b.ivaCliente)}</span></p>)}
            <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>Totale generale IVA inclusa</span><span>{formatEuro(econ.totCliente)}</span></p>
          </>
        ) : (
          <>
            <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>Totale generale IVA esclusa (cliente)</span><span>{formatEuro(econ.netCliente)}</span></p>
            {econ.breakdown.map((b, i) => <p key={i} style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>{b.label}</span><span>{formatEuro(b.ivaCliente)}</span></p>)}
            <p style={{ fontSize: 13, display: 'flex', justifyContent: 'space-between', fontWeight: 700, color: C.maroon, margin: '6px 0 0', paddingTop: 6, borderTop: `1.5px solid ${C.maroon}` }}><span>Totale generale IVA inclusa (cliente)</span><span>{formatEuro(econ.totCliente)}</span></p>
          </>
        )}
      </div>

      {(revision.noteFine || '').trim() && (
        <div style={{ marginTop: 26, padding: '10px 14px', borderTop: '1px solid #E5E2DA', fontSize: 10.5, lineHeight: 1.5, color: '#3A3A3A', whiteSpace: 'pre-wrap', breakInside: 'avoid', pageBreakInside: 'avoid' }}>
          <p style={{ fontSize: 9, fontWeight: 600, margin: '0 0 5px', color: '#8A8A8A', textTransform: 'uppercase', letterSpacing: 0.5 }}>Note</p>
          {revision.noteFine.trim()}
        </div>
      )}

          </td></tr>
        </tbody>
      </table>
    </div>
  );
}

// Nota libera del computo (in cima o in fondo): bozza locale salvata quando si esce dal campo, così non si
// salva (e non si crea una nuova revisione) a ogni lettera digitata. Il campo si può allargare in altezza.
function ComputoNoteField({ label, hint, value, onSave, placeholder }) {
  const [draft, setDraft] = useState(value || '');
  const [open, setOpen] = useState(!!(value || '').trim());
  useEffect(() => { setDraft(value || ''); if ((value || '').trim()) setOpen(true); }, [value]);
  const commit = () => { if ((draft || '') !== (value || '')) onSave(draft); };
  if (!open) {
    return (
      <button onClick={() => setOpen(true)} style={{ background: 'transparent', border: `1px dashed ${C.paleGray}`, borderRadius: 12, padding: '9px 14px', fontSize: 12.5, fontWeight: 400, color: C.gray, cursor: 'pointer', fontFamily: FONT, width: '100%', textAlign: 'left', margin: '0 0 12px' }}>
        + {label}
      </button>
    );
  }
  return (
    <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '12px 14px', margin: '0 0 12px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', gap: 10, marginBottom: 6 }}>
        <span style={{ fontSize: 12.5, fontWeight: 600, color: C.black, fontFamily: FONT }}>{label}</span>
        <span style={{ fontSize: 11, color: C.gray }}>{hint}</span>
      </div>
      <textarea
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        placeholder={placeholder}
        rows={3}
        style={{ width: '100%', boxSizing: 'border-box', border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '9px 11px', fontSize: 13, fontFamily: FONT, color: C.black, resize: 'vertical', minHeight: 64, lineHeight: 1.45, background: '#FDFCFA' }}
      />
    </div>
  );
}

function ProjectDetailPage({ project, onBack, onUpdateProject, listini, onUpdateListini, fornitoriCatalog, initialRevisionId, requestPdf, onOpenRilievo, profile }) {
  const revisions = project.revisions;
  // Identità di chi sta leggendo/modificando ora, per i segnalibri personali nel computo: una persona per
  // voce (email se disponibile, altrimenti il nome) così lo stesso nome digitato due volte non si confonde
  // con un'altra persona, e il nome mostrato è quello vero (nome e cognome) del profilo collegato.
  const myBookmarkKey = (profile?.email || profile?.name || 'io').trim().toLowerCase();
  const myBookmarkName = (profile?.name || profile?.email || 'Utente').trim();
  const latestRevision = revisions[revisions.length - 1];
  const [selectedRevisionId, setSelectedRevisionId] = useState(initialRevisionId || latestRevision?.id);
  const [selRevA, setSelRevA] = useState(revisions[0]?.id);
  const [selRevB, setSelRevB] = useState(revisions[revisions.length - 1]?.id);
  const [showCompare, setShowCompare] = useState(false);
  const [listinoId, setListinoId] = useState(listini[0]?.id);
  // Il pannello Listino/Fornitori a sinistra del computo: due sezioni contraibili indipendentemente, così
  // chi sta cercando nei fornitori può comprimere il listino (e viceversa) per avere più spazio su quella
  // che gli serve in quel momento, invece di scrollare dentro una colonna fissa divisa a metà.
  const [sidebarOpen, setSidebarOpen] = useState({ listino: true, fornitori: false });
  const toggleSidebarSection = (key) => setSidebarOpen((s) => ({ ...s, [key]: !s[key] }));
  const [dragOver, setDragOver] = useState(false);
  const [dragOverTarget, setDragOverTarget] = useState(null); // "macroName", "macroName|categoriaName" o "macroName|categoriaName|sottoName" evidenziato durante il drag
  const [showImportPdf, setShowImportPdf] = useState(false);
  const [voceComputoCtx, setVoceComputoCtx] = useState(null); // { macroName, categoriaName, sottoName, initialItem }
  const [expandedCronoTask, setExpandedCronoTask] = useState(null); // id della lavorazione di cronoprogramma con l'editor dei collegamenti aperto
  const [expandedItems, setExpandedItems] = useState({}); // { [itemId]: true } — dettaglio misurazioni/note aperto
  const [voceSearch, setVoceSearch] = useState(''); // ricerca testuale fra le voci del computo aperto
  const [voceSearchIdx, setVoceSearchIdx] = useState(0); // indice del risultato attualmente "a fuoco" fra quelli trovati

  // --- Scheda cliente importata da Desearq Studio Manager (stesso progetto Supabase, schema "public") ---
  const [showClientImport, setShowClientImport] = useState(false);
  const [dsmClients, setDsmClients] = useState(null); // null = non ancora caricati
  const [loadingDsmClients, setLoadingDsmClients] = useState(false);
  const [dsmClientsError, setDsmClientsError] = useState('');
  const [clientSearch, setClientSearch] = useState('');

  const selectedRevision = revisions.find((r) => r.id === selectedRevisionId) || latestRevision;
  const isEditingLatest = selectedRevision && latestRevision && selectedRevision.id === latestRevision.id;
  const items = selectedRevision?.items || [];
  const extraSections = selectedRevision?.extraSections || [];
  const sectionDiscounts = selectedRevision?.sectionDiscounts || {};
  const header = project.header || {};
  const activeListino = listini.find((l) => l.id === listinoId) || listini[0];

  const realItems = items.filter((it) => it.type !== 'subtotal');
  const sectionVatSaved = selectedRevision?.sectionVat || {};

  // Raggruppamento a tre livelli (macrocategoria > categoria > sottocategoria), calcolato da
  // buildComputoGroups (condiviso anche con la stampa PDF e l'export Excel, così restano sempre
  // coerenti con quanto mostrato qui a schermo).
  const groupedSections = buildComputoGroups(selectedRevision);
  const allSectionNames = groupedSections.map((s) => s.name);

  // Ricerca testuale fra le voci del computo aperto (per codice o descrizione): utile per ritrovare in
  // fretta una voce già inserita e da aggiornare/sistemare, senza dover scorrere a occhio tutto il computo.
  const voceSearchTerm = voceSearch.trim().toLowerCase();
  const voceTerms = searchTerms(voceSearch);
  const itemMatchesSearch = (it) => it.type !== 'subtotal' && matchesTerms(voceTerms, it.code, it.desc, it.note, it.noteCliente);
  const voceSearchMatches = voceSearchTerm
    ? realItems.filter(itemMatchesSearch)
    : [];
  const goToSearchMatch = (idx) => {
    if (voceSearchMatches.length === 0) return;
    const safeIdx = ((idx % voceSearchMatches.length) + voceSearchMatches.length) % voceSearchMatches.length;
    setVoceSearchIdx(safeIdx);
    const el = document.getElementById(`computo-item-${voceSearchMatches[safeIdx].id}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  };

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

  const economics = computeEconomics(selectedRevision);
  groupedSections.forEach((s) => {
    const e = economics.sections.find((x) => x.name === s.name);
    if (e) Object.assign(s, { vat: e.vat, ivaImpresa: e.ivaImpresa, ivaCliente: e.ivaCliente, totImpresa: e.totImpresa, totCliente: e.totCliente });
  });
  const totale = economics.totImpresa;
  const totaleCliente = economics.totCliente;

  // IVA di una singola sezione: stessa logica della vecchia "Configurazione IVA" generale (aliquota,
  // dicitura personalizzata, scelte rapide), ma salvata per sezione in revision.sectionVat.
  const updateSectionVat = (sectionName, patch) => {
    applyRevisionChange((rev) => {
      const current = (rev.sectionVat || {})[sectionName] || { rate: getSectionVat(rev, sectionName).rate, label: null };
      return { sectionVat: { ...(rev.sectionVat || {}), [sectionName]: { ...current, ...patch } } };
    });
  };

  const updateSectionDiscount = (sectionName, value) => {
    applyRevisionChange((rev) => ({ sectionDiscounts: { ...(rev.sectionDiscounts || {}), [sectionName]: value } }));
  };

  // Applica una modifica alla revisione (voci, sezioni extra...): se si sta lavorando sull'ultima
  // versione la aggiorna sul posto, altrimenti crea automaticamente una copia lasciando quella aperta intatta.
  // Ricalcola sempre i codici delle voci col codice automatico, in modo che qualunque spostamento
  // (voce, sottocategoria o macrocategoria) tenga i codici coerenti con la posizione attuale.
  const applyRevisionChange = (updater) => {
    const patch = updater(selectedRevision);
    const nextItems = regenerateItemCodes(resolveLinkedQuantities(patch.items || items));
    const patchWithCodes = patch.items ? { ...patch, items: nextItems } : patch;
    const realNextItems = nextItems.filter((it) => it.type !== 'subtotal');
    const total = formatEuro(sumImpresa(realNextItems));
    const totalCliente = formatEuro(sumCliente(realNextItems));
    if (isEditingLatest) {
      const updatedRevisions = revisions.map((r) => (r.id === selectedRevision.id ? { ...r, ...patchWithCodes, dateModified: nowLabel(), modifiedBy: myBookmarkName, total, totalCliente } : r));
      onUpdateProject({ ...project, revisions: updatedRevisions, value: total });
    } else {
      const newRev = { ...selectedRevision, ...patchWithCodes, id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), createdBy: myBookmarkName, modifiedBy: myBookmarkName, status: STATUS_OPTIONS[0], total, totalCliente };
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
    const item = { id: Date.now() + Math.random(), listinoRef: voce.uid ? { listinoId: voce.listinoId ?? activeListino?.id, uid: voce.uid } : null, code: voce.code, desc: voce.desc, unit: voce.unit, unitPriceImpresa: formatEuro(impresaVal).replace(' €', ''), unitPriceCliente: formatEuro(clienteVal).replace(' €', ''), qty: String(qty), macro: voce.macro, section: voce.macro };
    applyItemsChange((its) => [...its, item]);
  };

  // Aggiungere una voce dal Listino (trascinandola o toccando +) apre lo stesso pannello di misurazione
  // delle voci create da zero, precompilato con descrizione/unità/prezzi del listino: la modalità di calcolo
  // (ml/m²/m³, oppure quantità manuale per "a corpo"/"cadauna"/ecc.) si sceglie in base all'unità di misura.
  // La voce finisce come voce vera e propria nella macrocategoria/sottocategoria del listino di provenienza.
  // Voci di listino create o modificate direttamente da qui, senza lasciare il computo: si salvano nel
  // listino scelto nella colonna a sinistra (le modifiche di prezzo seguono le stesse regole di
  // aggiornamento dei computi in attesa). Una voce nuova, se richiesto, viene subito aggiunta al computo.
  const [listinoVoceModal, setListinoVoceModal] = useState(null); // { mode: 'new' } | { mode: 'edit', path }
  const listinoLocations = [];
  if (activeListino) {
    withCodes(activeListino.macros || []).forEach((m, mi) => m.categorie.forEach((c, ci) => c.sottocategorie.forEach((s, si) => {
      listinoLocations.push({ path: [mi, ci, si], label: `${m.name} › ${c.name} › ${s.name} (${s.code})` });
    })));
  }
  const saveListinoVoce = (path, voce, addNow) => {
    if (!activeListino || !onUpdateListini) return;
    const macrosNext = structuredClone(activeListino.macros || []);
    let saved;
    let at = path;
    if (listinoVoceModal?.mode === 'edit') {
      const [mi, ci, si, vi] = listinoVoceModal.path;
      const prev = macrosNext[mi].categorie[ci].sottocategorie[si].voci[vi];
      saved = { uid: prev?.uid || newVoceUid(), ...voce };
      macrosNext[mi].categorie[ci].sottocategorie[si].voci[vi] = saved;
      at = [mi, ci, si];
    } else {
      const [mi, ci, si] = path;
      saved = { uid: newVoceUid(), ...voce };
      macrosNext[mi].categorie[ci].sottocategorie[si].voci.push(saved);
    }
    onUpdateListini(listini.map((l) => (l.id === activeListino.id ? { ...l, macros: macrosNext } : l)));
    if (addNow) {
      const [mi, ci, si] = at;
      const m = macrosNext[mi]; const c = m.categorie[ci]; const s = c.sottocategorie[si];
      openVoceFromListino({ ...saved, listinoId: activeListino.id, macro: m.name, categoria: c.name, sotto: s.name });
    }
  };

  const openVoceFromListino = (voce) => {
    setVoceComputoCtx({
      macroName: voce.macro || 'Voci varie',
      categoriaName: voce.categoria || 'Generale',
      sottoName: voce.sotto || 'Generale',
      initialItem: null,
      prefill: { desc: voce.desc, unit: voce.unit, priceImpresa: voce.priceImpresa, priceCliente: voce.priceCliente, note: voce.note, listinoRef: voce.uid ? { listinoId: voce.listinoId ?? activeListino?.id, uid: voce.uid } : null },
    });
  };

  // Aggiungere un prodotto dal catalogo Fornitori (pannello sotto il Listino): a differenza delle voci di
  // listino non serve misurare nulla, basta la quantità. Stessa segnalazione di duplicato già vista altrove
  // (Listino->computo, Fornitori->computo): se una voce con la stessa descrizione esiste già, si chiede
  // conferma prima di crearne un'altra uguale.
  // Tutte le forniture finiscono nella sezione "FORNITURE", divise in sottosezioni per tipo (la macrosezione
  // del catalogo Fornitori: Forniture sanitarie, elettriche, infissi…), senza il nome del fornitore nella voce:
  // il fornitore scelto resta salvato a parte (campo "fornitore") ma non compare nel computo né nel PDF.
  const addFornitoreToComputo = (prodotto, variante, priceSource, tipoFornitura, gruppoFornitura) => {
    const desc = prodottoLabel(prodotto) + (variante.nome ? ` — ${variante.nome}` : '');
    const qtyStr = prompt(`Quantità per "${desc}":`, '1');
    if (qtyStr === null) return;
    const sectionName = 'FORNITURE';
    applyItemsChange((its) => {
      const dupe = its.find((it) => it.type !== 'subtotal' && String(it.desc || '').trim().toLowerCase() === desc.trim().toLowerCase());
      if (dupe && !confirm(`La voce "${desc}" è già presente nel computo${dupe.code ? ` (${dupe.code})` : ''}. Vuoi aggiungere comunque un duplicato?`)) {
        return its;
      }
      const newItem = {
        id: Date.now() + Math.random(), code: '', desc, unit: 'cad',
        unitPriceImpresa: priceSource.prezzoScontato, unitPriceCliente: priceSource.prezzoCliente,
        listinoRef: priceSource.prezzoListino, qty: qtyStr || '1', macro: sectionName, section: sectionName,
        categoria: tipoFornitura || 'Generale', sottocategoria: gruppoFornitura || 'Generale',
        fornitore: priceSource.id !== 'listino' ? priceSource.name : null,
      };
      return [...its, newItem];
    });
  };

  // Come openVoceFromListino, ma la macrosezione (e facoltativamente categoria/sottocategoria) di destinazione
  // sono quelle su cui l'utente ha trascinato fisicamente la voce nel computo — non quelle del listino d'origine.
  // Così si possono organizzare le voci secondo le macrocategorie/categorie/sottocategorie create nel computo,
  // anche quando non coincidono con quelle del listino.
  const addVoceToTarget = (voce, macroName, categoriaName, sottoName) => {
    setVoceComputoCtx({
      macroName,
      categoriaName: categoriaName || voce.categoria || 'Generale',
      sottoName: sottoName || voce.sotto || 'Generale',
      initialItem: null,
      prefill: { desc: voce.desc, unit: voce.unit, priceImpresa: voce.priceImpresa, priceCliente: voce.priceCliente, note: voce.note, listinoRef: voce.uid ? { listinoId: voce.listinoId ?? activeListino?.id, uid: voce.uid } : null },
    });
  };

  // Il rilievo da planimetria (punti, aree, lunghezze, volumi) ha una sua sezione dedicata nel menu laterale
  // ("Rilievo planimetrico", componente RilievoPage più sotto in questo file) — qui nella scheda progetto
  // resta solo un rimando, per non duplicare logica e stato.

  const updateQty = (id, qty) => {
    applyItemsChange((its) => its.map((it) => (it.id === id ? { ...it, qty } : it)));
  };

  const removeItem = (id) => {
    if (!confirm('Eliminare questa voce?')) return;
    applyItemsChange((its) => its.filter((it) => it.id !== id));
  };

  // Segnalibro personale sul computo: segna "fin dove sono arrivato a leggere" su una voce precisa,
  // uno per persona (chi lo sposta sposta solo il PROPRIO, non quello dei colleghi). Aggiorna la revisione
  // aperta sul posto, senza creare una nuova versione: è solo un promemoria di lettura, non una modifica
  // al computo, quindi non deve "sporcare" lo storico delle revisioni né quelle precedenti/chiuse.
  const toggleBookmark = (itemId) => {
    const existing = selectedRevision.bookmarks || [];
    const mine = existing.find((b) => b.key === myBookmarkKey);
    const nextBookmarks = mine && mine.itemId === itemId
      ? existing.filter((b) => b.key !== myBookmarkKey) // si clicca di nuovo sulla stessa voce: rimuove il proprio segnalibro
      : [...existing.filter((b) => b.key !== myBookmarkKey), { key: myBookmarkKey, name: myBookmarkName, date: nowLabel(), itemId }];
    const updatedRevisions = revisions.map((r) => (r.id === selectedRevision.id ? { ...r, bookmarks: nextBookmarks } : r));
    onUpdateProject({ ...project, revisions: updatedRevisions });
  };

  const scrollToItem = (itemId) => {
    const el = document.getElementById(`computo-item-${itemId}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  };

  const moveItemToSection = (id, sectionName) => {
    applyItemsChange((its) => its.map((it) => (it.id === id ? { ...it, section: sectionName, categoria: 'Generale', sottocategoria: 'Generale' } : it)));
  };

  // Sposta una voce in un'altra categoria della STESSA macrosezione (senza toccare macro o sottocategoria).
  const moveItemToCategoria = (id, categoriaName) => {
    applyItemsChange((its) => its.map((it) => (it.id === id ? { ...it, categoria: categoriaName } : it)));
  };

  // Sposta una voce su/giù, scambiandola con la voce precedente/successiva della stessa categoria
  // (all'interno della stessa macrocategoria): è quest'ordine, insieme a quello di macro e categorie,
  // a determinare il codice automatico di ogni voce.
  const moveItemInSection = (id, direction) => {
    applyItemsChange((its) => {
      const item = its.find((it) => it.id === id);
      if (!item) return its;
      const sectionName = item.section || item.macro || 'Voci varie';
      const catName = item.categoria || 'Generale';
      const sameGroupIdx = its
        .map((it, idx) => ({ it, idx }))
        .filter((o) => (o.it.section || o.it.macro || 'Voci varie') === sectionName && (o.it.categoria || 'Generale') === catName)
        .map((o) => o.idx);
      const idxA = its.indexOf(item);
      const posInSection = sameGroupIdx.indexOf(idxA);
      const swapPos = direction === 'up' ? posInSection - 1 : posInSection + 1;
      if (swapPos < 0 || swapPos >= sameGroupIdx.length) return its;
      const idxB = sameGroupIdx[swapPos];
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
    const moveKey = (obj) => {
      if (!obj || !(oldName in obj)) return obj;
      const { [oldName]: v, ...rest } = obj;
      return { ...rest, [newName]: v };
    };
    applyRevisionChange((rev) => ({
      items: (rev.items || []).map((it) => ((it.section || it.macro) === oldName ? { ...it, section: newName } : it)),
      extraSections: (rev.extraSections || []).map((n) => (n === oldName ? newName : n)),
      sectionDiscounts: moveKey(rev.sectionDiscounts),
      sectionVat: moveKey(rev.sectionVat),
      macroOrder: (rev.macroOrder || []).map((n) => (n === oldName ? newName : n)),
    }));
  };

  // Sposta una macrocategoria su/giù, scambiandola con la precedente/successiva nell'ordine mostrato.
  const moveMacroSection = (name, direction) => {
    const idx = allSectionNames.indexOf(name);
    const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= allSectionNames.length) return;
    const next = [...allSectionNames];
    [next[idx], next[swapIdx]] = [next[swapIdx], next[idx]];
    applyRevisionChange(() => ({ macroOrder: next }));
  };

  // Sposta una categoria su/giù all'interno della sua macrocategoria.
  const moveCategoria = (macroName, catName, direction) => {
    const section = groupedSections.find((s) => s.name === macroName);
    const current = section ? section.categorie.map((c) => c.name) : [];
    const idx = current.indexOf(catName);
    const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= current.length) return;
    const next = [...current];
    [next[idx], next[swapIdx]] = [next[swapIdx], next[idx]];
    applyRevisionChange((rev) => ({ categorieOrder: { ...(rev.categorieOrder || {}), [macroName]: next } }));
  };

  const addCategoria = (macroName) => {
    const name = prompt('Nome della nuova categoria (es. Opere edili, Opere da piastrellista...):');
    if (!name) return;
    applyRevisionChange((rev) => ({
      categorieDefinite: { ...(rev.categorieDefinite || {}), [macroName]: [...((rev.categorieDefinite || {})[macroName] || []), name] },
    }));
  };

  const renameCategoria = (macroName, oldName) => {
    const newName = prompt('Rinomina categoria:', oldName);
    if (!newName || newName === oldName) return;
    applyRevisionChange((rev) => ({
      items: (rev.items || []).map((it) => ((it.section || it.macro) === macroName && (it.categoria || 'Generale') === oldName ? { ...it, categoria: newName } : it)),
      categorieDefinite: { ...(rev.categorieDefinite || {}), [macroName]: ((rev.categorieDefinite || {})[macroName] || []).map((n) => (n === oldName ? newName : n)) },
      categorieOrder: { ...(rev.categorieOrder || {}), [macroName]: ((rev.categorieOrder || {})[macroName] || []).map((n) => (n === oldName ? newName : n)) },
    }));
  };

  const removeCategoria = (macroName, catName) => {
    if (!confirm(`Rimuovere la categoria "${catName}"? (possibile solo se vuota)`)) return;
    applyRevisionChange((rev) => ({
      categorieDefinite: { ...(rev.categorieDefinite || {}), [macroName]: ((rev.categorieDefinite || {})[macroName] || []).filter((n) => n !== catName) },
      categorieOrder: { ...(rev.categorieOrder || {}), [macroName]: ((rev.categorieOrder || {})[macroName] || []).filter((n) => n !== catName) },
    }));
  };

  // Crea o modifica una voce con misurazioni reali, oppure — se mergeIntoItemId è indicato — aggiunge i
  // nuovi gruppi di misurazione a una voce già esistente della stessa sottocategoria, sommandone le
  // quantità sotto un'unica riga di costo (un solo prezzo unitario per il totale sommato).
  const saveVoceComputo = (macroName, categoriaName, sottoName, voceData) => {
    const impresaVal = parseEuro(voceData.priceImpresa);
    const clienteVal = evalClientPrice(voceData.priceCliente, impresaVal);
    const priceFields = {
      desc: voceData.desc, unit: voceData.unit, note: voceData.note || '', noteCliente: voceData.noteCliente || '',
      unitPriceImpresa: formatEuro(impresaVal).replace(' €', ''),
      unitPriceCliente: formatEuro(clienteVal).replace(' €', ''),
    };
    applyItemsChange((its) => {
      if (voceData.mergeIntoItemId) {
        return its.map((it) => {
          if (String(it.id) !== String(voceData.mergeIntoItemId)) return it;
          const mergedMisurazioni = [...(it.misurazioni || []), ...voceData.misurazioni];
          const qty = voceData.unitaCalcolo
            ? String(Math.round(computeVoceQtyTotal(mergedMisurazioni, it.unitaCalcolo) * 100) / 100).replace('.', ',')
            : it.qty;
          return { ...it, misurazioni: mergedMisurazioni, qty };
        });
      }
      const qty = voceData.unitaCalcolo
        ? String(Math.round(computeVoceQtyTotal(voceData.misurazioni, voceData.unitaCalcolo) * 100) / 100).replace('.', ',')
        : voceData.manualQty || '0';
      if (voceData.editId) {
        return its.map((it) => (it.id === voceData.editId ? {
          ...it, ...priceFields, unitaCalcolo: voceData.unitaCalcolo, misurazioni: voceData.misurazioni, qty, qtyLink: voceData.qtyLink || null,
        } : it));
      }
      // Prima di creare davvero una nuova voce (non in modifica, non in unione con una già esistente):
      // se nel computo c'è già una voce con la stessa descrizione, lo si segnala e si chiede conferma,
      // così nessuno aggiunge per sbaglio un doppione senza accorgersene — se si risponde di no non
      // succede nulla, la voce resta quella già presente.
      const dupe = its.find((it) => it.type !== 'subtotal' && String(it.desc || '').trim().toLowerCase() === String(priceFields.desc || '').trim().toLowerCase());
      if (dupe && !confirm(`La voce "${priceFields.desc}" è già presente nel computo${dupe.code ? ` (${dupe.code})` : ''}. Vuoi aggiungere comunque un duplicato?`)) {
        return its;
      }
      const newItem = {
        id: Date.now() + Math.random(), code: '', autoCode: true, ...priceFields, listinoRef: voceData.listinoRef || null, qtyLink: voceData.qtyLink || null,
        qty, macro: macroName, section: macroName, categoria: categoriaName || 'Generale', sottocategoria: sottoName,
        unitaCalcolo: voceData.unitaCalcolo, misurazioni: voceData.misurazioni,
      };
      return [...its, newItem];
    });
  };

  const addPartialSubtotal = (sectionName) => {
    const title = prompt('Titolo della sommatoria parziale:', 'Sommatoria parziale');
    if (!title) return;
    const applyVat = confirm('Applicare l\'IVA a questa sommatoria parziale? (Annulla = nessuna IVA qui)');
    let vatRate = null;
    let vatLabel = null;
    if (applyVat) {
      const defaultRate = getSectionVat(selectedRevision, sectionName).rate;
      const rateInput = prompt('Aliquota IVA per questa sommatoria (%):', String(defaultRate));
      if (rateInput !== null) {
        vatRate = Number(rateInput) || 0;
        vatLabel = prompt('Dicitura IVA (facoltativa):', `IVA ${vatRate}%`) || `IVA ${vatRate}%`;
      }
    }
    // Sconto percentuale facoltativo su questa sommatoria: se lasciato vuoto/zero non viene salvato,
    // così in stampa PDF la riga "Sconto" non compare affatto (uno sconto assente non è un errore).
    const discountInput = prompt('Sconto percentuale su questa sommatoria (facoltativo, lascia vuoto se nessuno):', '');
    const discountPct = discountInput && Number(discountInput) ? Number(discountInput) : null;
    const marker = { id: Date.now() + Math.random(), type: 'subtotal', title, section: sectionName, vatRate, vatLabel, discountPct };
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
      const discountInput = prompt('Sconto percentuale su questa sommatoria (facoltativo, lascia vuoto se nessuno):', it.discountPct ? String(it.discountPct) : '');
      const discountPct = discountInput && Number(discountInput) ? Number(discountInput) : null;
      return { ...it, title, vatRate, vatLabel, discountPct };
    }));
  };

  // Rimuove una macrosezione senza voci vere: ripulisce anche eventuali "Sommatoria parziale" rimaste
  // agganciate a quella sezione (altrimenti la macrosezione ricomparirebbe da sola perché quelle righe
  // esistono ancora), oltre a categorie/sottocategorie vuote definite lì, ai loro ordini salvati e al
  // suo ordine di macrosezione.
  const removeMarkerOrEmptySection = (sectionName) => {
    if (!confirm(`Rimuovere la macrosezione "${sectionName}"?`)) return;
    const belongsToSection = (k) => k === sectionName || k.startsWith(`${sectionName}|`);
    applyRevisionChange((rev) => ({
      extraSections: (rev.extraSections || []).filter((n) => n !== sectionName),
      items: (rev.items || []).filter((it) => (it.section || it.macro || 'Voci varie') !== sectionName),
      categorieDefinite: Object.fromEntries(Object.entries(rev.categorieDefinite || {}).filter(([k]) => k !== sectionName)),
      categorieOrder: Object.fromEntries(Object.entries(rev.categorieOrder || {}).filter(([k]) => k !== sectionName)),
      sottocategorie: Object.fromEntries(Object.entries(rev.sottocategorie || {}).filter(([k]) => !belongsToSection(k))),
      sottocategorieOrder: Object.fromEntries(Object.entries(rev.sottocategorieOrder || {}).filter(([k]) => !belongsToSection(k))),
      macroOrder: (rev.macroOrder || []).filter((n) => n !== sectionName),
    }));
  };

  const updateHeader = (field, value) => {
    onUpdateProject({ ...project, header: { ...header, [field]: value } });
  };

  // --- Scheda cliente importata da Desearq Studio Manager ---
  // Desearq Studio Manager e Gestionale Edile condividono lo stesso progetto Supabase: l'anagrafica clienti
  // vive in public.app_state.data.clients (schema "public", diverso da "cea" usato per i dati di questa app).
  // La lettura è di sola consultazione: si importa lo snapshot dei dati del cliente scelto dentro al progetto,
  // che poi restano modificabili qui senza toccare in alcun modo i dati in Desearq Studio Manager.
  const openClientImport = () => {
    setShowClientImport(true);
    if (dsmClients !== null || loadingDsmClients) return;
    setLoadingDsmClients(true);
    setDsmClientsError('');
    supabase.rpc('quant_clienti').then(({ data, error }) => {
      setLoadingDsmClients(false);
      if (error) { setDsmClientsError('Impossibile leggere l\'anagrafica clienti da Desearq Studio Manager: ' + error.message); return; }
      setDsmClients(Array.isArray(data) ? data : []);
    });
  };

  const importClientSheet = (c) => {
    onUpdateProject({
      ...project,
      clientSheet: {
        source: 'desearq-studio-manager', importedId: c.id, importedAt: nowLabel(),
        name: c.name || '', type: c.type || '', cf: c.cf || '', piva: c.piva || '',
        email: c.email || '', phone: c.phone || '', address: c.address || '',
        properties: c.properties || [],
      },
    });
    setShowClientImport(false);
    setClientSearch('');
  };

  const updateClientSheetField = (field, value) => {
    onUpdateProject({ ...project, clientSheet: { ...(project.clientSheet || {}), [field]: value } });
  };

  const removeClientSheet = () => {
    if (!confirm('Rimuovere la scheda cliente da questo progetto?')) return;
    onUpdateProject({ ...project, clientSheet: null });
  };

  // --- Offerte imprese: link per far quotare il computo a un'impresa e offerte ricevute ---
  const [showOfferte, setShowOfferte] = useState(false);
  const [pendingOfferte, setPendingOfferte] = useState(0);
  const refreshPendingOfferte = React.useCallback(async () => {
    const { data: l } = await cea.from('computo_offerta_links').select('id').eq('project_id', String(project.id));
    const ids = (l || []).map((x) => x.id);
    if (!ids.length) { setPendingOfferte(0); return; }
    const { count } = await cea.from('computo_offerte').select('id', { count: 'exact', head: true }).in('link_id', ids).eq('status', 'pending');
    setPendingOfferte(count || 0);
  }, [project.id]);
  React.useEffect(() => { refreshPendingOfferte(); }, [refreshPendingOfferte]);

  // Accettazione di un'offerta: nuova revisione copiata da quella inviata all'impresa, intitolata con il nome
  // dell'impresa, con i prezzi offerti come prezzi impresa. Prezzi cliente: invariati oppure prezzo offerto +
  // ricarico %. Le voci che l'impresa non ha quotato restano come sono.
  const acceptOffer = (offerta, link, clienteMode, ricarico) => {
    const source = revisions.find((r) => String(r.id) === String(link?.revision_id)) || latestRevision;
    if (!source) { alert('Questo progetto non ha più revisioni a cui applicare l\'offerta.'); return null; }
    const prezzi = new Map((offerta.prezzi || []).map((p) => [String(p.id), p]));
    const newItems = structuredClone(source.items || []).map((it) => {
      if (it.type === 'subtotal') return it;
      const p = prezzi.get(String(it.id));
      if (!p) return it;
      const imp = Number(p.price) || 0;
      const next = { ...it, unitPriceImpresa: formatEuro(imp).replace(' €', '') };
      if (clienteMode === 'ricarico') next.unitPriceCliente = formatEuro(imp * (1 + (Number(ricarico) || 0) / 100)).replace(' €', '');
      return next;
    });
    const real = newItems.filter((it) => it.type !== 'subtotal');
    const total = formatEuro(sumImpresa(real));
    const totalCliente = formatEuro(sumCliente(real));
    const newRev = {
      ...structuredClone(source), id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: offerta.impresa_nome || link?.impresa_nome || 'Offerta impresa',
      dateCreated: nowLabel(), dateModified: nowLabel(), createdBy: myBookmarkName, modifiedBy: myBookmarkName, status: STATUS_OPTIONS[0],
      items: newItems, total, totalCliente, offerta: { id: offerta.id, impresa: offerta.impresa_nome, totale: Number(offerta.totale) || 0 },
    };
    onUpdateProject({ ...project, revisions: [...revisions, newRev], value: total });
    setSelectedRevisionId(newRev.id);
    return newRev.id;
  };

  const saveNewVersion = () => {
    const total = formatEuro(sumImpresa(realItems));
    const totalCliente = formatEuro(sumCliente(realItems));
    const newRev = { id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: structuredClone(items), extraSections: structuredClone(extraSections), total, totalCliente };
    onUpdateProject({ ...project, revisions: [...revisions, newRev], value: total });
    setSelectedRevisionId(newRev.id);
  };

  // Crea una nuova revisione (o la prima, se il progetto non ne ha ancora) a partire dai totali per
  // macrocategoria importati da un PDF: un solo elemento per categoria, con il totale cliente letto dal
  // PDF e la quota di totale impresa ripartita in proporzione. Da qui si potrà "Salvare una nuova versione"
  // come sempre, aggiungendo o correggendo le singole voci una alla volta.
  const createRevisionFromPdfTotals = (rowsIn, totaleImpresaManuale) => {
    const sommaCategorie = rowsIn.reduce((sum, r) => sum + r.totale, 0);
    const newItems = rowsIn.map((r) => {
      const quotaImpresa = sommaCategorie > 0 ? totaleImpresaManuale * (r.totale / sommaCategorie) : 0;
      return {
        id: Date.now() + Math.random(),
        code: '',
        desc: 'Totale importato da PDF — da dettagliare con le singole voci',
        unit: '',
        unitPriceImpresa: formatEuro(quotaImpresa).replace(' €', ''),
        unitPriceCliente: formatEuro(r.totale).replace(' €', ''),
        qty: '1',
        macro: r.name,
        section: r.name,
        imported: true,
      };
    });
    const total = formatEuro(sumImpresa(newItems));
    const totalCliente = formatEuro(sumCliente(newItems));
    const newRev = {
      id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null,
      dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0],
      items: newItems, extraSections: [], total, totalCliente,
    };
    onUpdateProject({ ...project, revisions: [...revisions, newRev], header, value: total });
    setSelectedRevisionId(newRev.id);
    setShowImportPdf(false);
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

  // Team assegnato: si sceglie fra le persone della sezione Team (cea.team_members), non più a testo libero.
  // Nel progetto si salva il riferimento (memberId) più nome/email di riserva; nome, foto e ruolo mostrati
  // vengono sempre dal profilo attuale della persona, così restano aggiornati se cambiano nel Team.
  const [teamMembers, setTeamMembers] = useState([]);
  const [showTeamPicker, setShowTeamPicker] = useState(false);
  React.useEffect(() => {
    cea.from('team_members').select('id, name, email, role, job_title, photo, status').order('name').then(({ data }) => setTeamMembers(data || []));
  }, []);
  // Collegamento alla persona del Team: per id; per le voci vecchie (solo nome) per nome esatto, oppure per
  // nome di battesimo se una sola persona del Team inizia con quel nome (es. "Angelica" → "Angelica Rossi").
  const memberFor = (m) => {
    if (m.memberId) return teamMembers.find((t) => t.id === m.memberId);
    const n = (m.name || '').trim().toLowerCase();
    if (!n) return undefined;
    const exact = teamMembers.find((t) => (t.name || '').trim().toLowerCase() === n);
    if (exact) return exact;
    const byFirst = teamMembers.filter((t) => (t.name || '').trim().toLowerCase().startsWith(n + ' '));
    return byFirst.length === 1 ? byFirst[0] : undefined;
  };
  const addTeamMember = () => setShowTeamPicker(true);
  const assignMembers = (picked) => {
    const current = project.team || [];
    const additions = picked
      .filter((t) => !current.some((m) => memberFor(m)?.id === t.id))
      .map((t) => ({ memberId: t.id, name: t.name || t.email, email: t.email || '', role: t.job_title || t.role || 'Collaboratore' }));
    // Collega anche le voci create a mano in passato, se il nome corrisponde a una persona del Team
    const relinked = current.map((m) => { const t = memberFor(m); return t && !m.memberId ? { ...m, memberId: t.id, email: t.email || '' } : m; });
    onUpdateProject({ ...project, team: [...relinked, ...additions] });
    setShowTeamPicker(false);
  };

  const removeTeamMember = (idx) => {
    onUpdateProject({ ...project, team: (project.team || []).filter((_, i) => i !== idx) });
  };

  // --- Stati avanzamento pagamenti (SAL cantiere per categoria) ---
  // project.cantiereSal è un oggetto { [nomeCategoria]: { impresa, fatture: [{id, label, importo, stato}] } }.
  // Le categorie e il loro totale arrivano in automatico dall'ultima revisione approvata del computo;
  // impresa e fatture restano quelle inserite qui finché non le cambi tu, anche se il computo si aggiorna.
  const cantiereSal = project.cantiereSal || {};

  const setImpresaCategoria = (categoria) => {
    const current = cantiereSal[categoria] || { impresa: '', fatture: [] };
    const impresa = prompt('Impresa assegnata a questa categoria:', current.impresa || '');
    if (impresa === null) return;
    onUpdateProject({ ...project, cantiereSal: { ...cantiereSal, [categoria]: { ...current, impresa } } });
  };

  // Fattura di un SAL: si apre una finestra dove si inserisce la percentuale (l'importo si calcola da solo sul
  // totale della categoria) oppure direttamente l'importo. { categoria, base, fattura? } — fattura = modifica.
  const [fatturaCtx, setFatturaCtx] = useState(null);
  const addFatturaCategoria = (categoria, base) => setFatturaCtx({ categoria, base, fattura: null });
  const saveFatturaCategoria = (categoria, data) => {
    const current = cantiereSal[categoria] || { impresa: '', fatture: [] };
    const list = current.fatture || [];
    const fatture = data.id && list.some((f) => f.id === data.id)
      ? list.map((f) => (f.id === data.id ? { ...f, ...data } : f))
      : [...list, { ...data, id: Date.now(), stato: data.stato || 'Emessa' }];
    onUpdateProject({ ...project, cantiereSal: { ...cantiereSal, [categoria]: { ...current, fatture } } });
    setFatturaCtx(null);
  };

  const updateFatturaStato = (categoria, fatturaId, stato) => {
    const current = cantiereSal[categoria] || { impresa: '', fatture: [] };
    const fatture = (current.fatture || []).map((f) => (f.id === fatturaId ? { ...f, stato } : f));
    onUpdateProject({ ...project, cantiereSal: { ...cantiereSal, [categoria]: { ...current, fatture } } });
  };

  const removeFatturaCategoria = (categoria, fatturaId) => {
    const current = cantiereSal[categoria] || { impresa: '', fatture: [] };
    const fatture = (current.fatture || []).filter((f) => f.id !== fatturaId);
    onUpdateProject({ ...project, cantiereSal: { ...cantiereSal, [categoria]: { ...current, fatture } } });
  };

  // --- Cronoprogramma: lavorazioni con data di inizio e durata, salvate su project.cronoprogramma =
  // { startDate, tasks: [{id, name, color, startDay, durationDays, linkedCodes, manual}] }. startDay è
  // l'offset in giorni da startDate (più semplice da ricalcolare che una data assoluta per ogni lavorazione).
  // "Genera dal computo" propone una lavorazione per ogni categoria del computo APPROVATO (stesso riferimento
  // di "Stati avanzamento pagamenti" qui sotto) che non ha già una lavorazione con lo stesso nome: non tocca
  // mai quelle esistenti, per non perdere date spostate a mano quando il computo cambia.
  const crono = project.cronoprogramma || { startDate: todayISO(), tasks: [] };
  const updateCrono = (patch) => onUpdateProject({ ...project, cronoprogramma: { ...crono, ...patch } });
  const nextCronoCursor = (tasks) => tasks.reduce((max, t) => Math.max(max, t.startDay + t.durationDays), 0);

  const generateCronoFromComputo = () => {
    const approved = latestApprovedRevision(project);
    if (!approved) { alert('Nessuna revisione approvata: approva prima un computo (dalle Revisioni salvate qui sotto) per poter generare le lavorazioni.'); return; }
    const suggestions = buildCronoSuggestions(approved);
    const existingNames = new Set(crono.tasks.map((t) => t.name));
    const toAdd = suggestions.filter((s) => !existingNames.has(s.name));
    if (toAdd.length === 0) { alert('Tutte le categorie del computo approvato hanno già una lavorazione corrispondente nel cronoprogramma.'); return; }
    let cursor = nextCronoCursor(crono.tasks);
    const newTasks = toAdd.map((s) => {
      const task = { id: Date.now() + Math.random(), name: s.name, color: s.color, startDay: cursor, durationDays: 5, linkedCodes: s.linkedCodes, manual: false };
      cursor += 5;
      return task;
    });
    updateCrono({ tasks: [...crono.tasks, ...newTasks] });
  };

  const addManualCronoTask = () => {
    const name = prompt('Nome della lavorazione:');
    if (!name) return;
    const cursor = nextCronoCursor(crono.tasks);
    const task = { id: Date.now() + Math.random(), name, color: C.midGray, startDay: cursor, durationDays: 5, linkedCodes: [], manual: true };
    updateCrono({ tasks: [...crono.tasks, task] });
  };

  const updateCronoTask = (id, patch) => updateCrono({ tasks: crono.tasks.map((t) => (t.id === id ? { ...t, ...patch } : t)) });
  const renameCronoTask = (id, currentName) => { const name = prompt('Nome lavorazione:', currentName); if (!name) return; updateCronoTask(id, { name }); };
  const removeCronoTask = (id) => { if (confirm('Rimuovere questa lavorazione dal cronoprogramma?')) updateCrono({ tasks: crono.tasks.filter((t) => t.id !== id) }); };

  const cronoLinkableCategories = (() => {
    const ref = latestApprovedRevision(project) || revisions[revisions.length - 1];
    return ref ? buildCronoSuggestions(ref) : [];
  })();
  const toggleCronoTaskCategoria = (id, categoria) => {
    const task = crono.tasks.find((t) => t.id === id);
    if (!task) return;
    const linked = task.linkedCodes || [];
    const allLinked = categoria.linkedCodes.every((c) => linked.includes(c));
    const next = allLinked
      ? linked.filter((c) => !categoria.linkedCodes.includes(c))
      : [...new Set([...linked, ...categoria.linkedCodes])];
    updateCronoTask(id, { linkedCodes: next });
  };

  const headerField = (label, field, placeholder, type = 'text') => (
    <div>
      <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>{label}</label>
      <input
        type={type}
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

      {showImportPdf && <ImportPdfComputoModal onClose={() => setShowImportPdf(false)} onConfirm={createRevisionFromPdfTotals} />}

      {revisions.length === 0 ? (
        <div style={{ ...card, marginBottom: 24, textAlign: 'center', padding: 36 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 14px' }}>Questo progetto non ha ancora un computo metrico.</p>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button onClick={startComputo} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>+ Crea primo computo metrico</button>
            <button onClick={() => setShowImportPdf(true)} style={{ background: C.white, color: C.black, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>📄 Importa da PDF (macrocategorie)</button>
          </div>
        </div>
      ) : (
        <>
          <div style={{ ...card, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Versione aperta</label>
            <select value={selectedRevisionId} onChange={(e) => setSelectedRevisionId(revisions.find((r) => String(r.id) === e.target.value)?.id ?? selectedRevisionId)} style={{ fontSize: 13, fontWeight: 600, padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
              {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · modificata il {r.dateModified}</option>)}
            </select>
            {!isEditingLatest && (
              <span style={{ fontSize: 11, color: C.darkGray, background: 'rgba(67,67,67,0.1)', padding: '5px 10px', borderRadius: 8 }}>
                Stai visualizzando una versione precedente: la prima modifica creerà automaticamente una nuova versione, lasciando questa intatta.
              </span>
            )}
          </div>

          {/* Segnalibri di lettura: uno per persona, su questa revisione. Ogni membro del team vede dove sono
              arrivati tutti (nome e cognome + data), e può spostare solo il PROPRIO cliccando "🔖" sulla riga
              di una voce (vedi tabella sotto) — non crea una nuova versione, è solo un promemoria di lettura. */}
          {(selectedRevision.bookmarks || []).length > 0 && (
            <div style={{ ...card, marginBottom: 18 }}>
              <p style={{ fontWeight: 700, fontSize: 13, margin: '0 0 8px', color: C.black, fontFamily: FONT }}>🔖 Segnalibri di lettura su questa revisione</p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {(selectedRevision.bookmarks || []).map((b) => {
                  const atItem = realItems.find((it) => it.id === b.itemId);
                  const isMine = b.key === myBookmarkKey;
                  return (
                    <div key={b.key} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, flexWrap: 'wrap' }}>
                      <span style={{ fontWeight: 700, color: isMine ? C.maroon : C.black }}>{b.name}{isMine ? ' (tu)' : ''}</span>
                      <span style={{ color: C.gray }}>· {b.date}</span>
                      <span style={{ color: C.midGray, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 360 }}>
                        {atItem ? `— ${atItem.code} ${atItem.desc}` : '— voce non più presente in questa revisione'}
                      </span>
                      {atItem && <button onClick={() => scrollToItem(atItem.id)} style={rowBtnStyle}>📍 Vai</button>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Barra azioni del computo, divisa per categorie: versioni, esportazioni (menu PDF / Excel con tutte le
              varianti) e richieste di offerta alle imprese. */}
          <div className="computo-toolbar" style={{ display: 'flex', alignItems: 'stretch', gap: 0, marginBottom: 18, flexWrap: 'wrap', background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 16, padding: '10px 6px' }}>
            <ToolbarGroup label="Versioni">
              <button onClick={saveNewVersion} style={tbBtn(true)}>+ Salva nuova versione</button>
              <button onClick={() => setShowImportPdf(true)} style={tbBtn()}>Nuova revisione da PDF</button>
            </ToolbarGroup>
            <ToolbarGroup label="Esporta">
              <ExportMenu
                label="PDF"
                items={[
                  { label: 'Versione cliente', desc: 'Tutte le voci con i prezzi cliente', onClick: () => requestPdf(project, selectedRevision, true) },
                  { label: 'Riepilogo cliente', desc: 'Versione semplificata: solo categorie e totali', onClick: () => requestPdf(project, selectedRevision, 'riepilogo') },
                  { label: 'Completo impresa + cliente', desc: 'Uso interno, con i costi impresa', onClick: () => requestPdf(project, selectedRevision, false) },
                  { label: 'Senza prezzi', desc: "Voci e quantità da far quotare a un'impresa", onClick: () => requestPdf(project, selectedRevision, 'muto') },
                ]}
              />
              <ExportMenu
                label="Excel"
                items={[
                  { label: 'Versione cliente', desc: 'Tutte le voci con i prezzi cliente', onClick: () => exportComputoExcel(project, selectedRevision, true) },
                  { label: 'Riepilogo cliente', desc: 'Versione semplificata: solo categorie e totali', onClick: () => exportComputoExcel(project, selectedRevision, 'riepilogo') },
                  { label: 'Completo impresa + cliente', desc: 'Uso interno, con i costi impresa', onClick: () => exportComputoExcel(project, selectedRevision, false) },
                  { label: 'Senza prezzi', desc: 'Con le formule pronte per i prezzi', onClick: () => exportComputoExcelSenzaPrezzi(project, selectedRevision) },
                ]}
              />
            </ToolbarGroup>
            <ToolbarGroup label="Imprese" last>
              <button onClick={() => setShowOfferte(true)} title="Manda il computo a un'impresa per farlo quotare online" style={{ ...tbBtn(), position: 'relative' }}>
                Offerte imprese
                {pendingOfferte > 0 && <span style={{ position: 'absolute', top: -7, right: -7, background: C.maroon, color: C.white, borderRadius: 999, minWidth: 18, height: 18, fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px', border: '2px solid #fff' }}>{pendingOfferte}</span>}
              </button>
            </ToolbarGroup>
            {showOfferte && (
              <OfferteImpreseModal project={project} revision={selectedRevision} onClose={() => { setShowOfferte(false); refreshPendingOfferte(); }} onAcceptOffer={acceptOffer} />
            )}
          </div>

          <div style={{ ...card, marginBottom: 18 }}>
            <p style={{ fontWeight: 700, fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Dati generali del computo</p>
            <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
              {headerField("Descrizione generale dell'opera", 'descrizione', 'Es. Ristrutturazione integrale di unità residenziale')}
              {headerField('Ubicazione cantiere', 'ubicazione', 'Indirizzo del cantiere')}
            </div>
            <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12, marginBottom: 12 }}>
              {headerField('Progettista', 'progettista', 'Nome del progettista')}
              {headerField('Direttore lavori', 'direttoreLavori', 'Nome del direttore lavori')}
              {headerField('Impresa esecutrice', 'impresa', 'Ragione sociale impresa')}
            </div>
            <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              {headerField('Numero pratica / commessa', 'numeroPratica', 'Es. 24/2026')}
              {headerField('Data documento', 'dataDocumento', 'Es. 31/07/2026')}
            </div>
            <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginTop: 12 }}>
              {headerField('Inizio lavori', 'dataInizioLavori', '', 'date')}
              {headerField('Fine lavori', 'dataFineLavori', '', 'date')}
            </div>
            <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0' }}>
              Committente: {project.client} · Progetto: {project.name}
            </p>
          </div>

          <div style={{ ...card, marginBottom: 18 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8, marginBottom: 12 }}>
              <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Scheda cliente</p>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={openClientImport} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>
                  ⇩ {project.clientSheet ? 'Reimporta' : 'Importa'} da Desearq Studio Manager
                </button>
                {project.clientSheet && (
                  <button onClick={removeClientSheet} style={rowBtnStyle}>🗑 Rimuovi</button>
                )}
              </div>
            </div>

            {showClientImport && (
              <div style={{ border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: 12, marginBottom: 14, background: '#f7f5f0' }}>
                <input
                  value={clientSearch}
                  onChange={(e) => setClientSearch(e.target.value)}
                  placeholder="Cerca cliente per nome…"
                  autoFocus
                  style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginBottom: 8 }}
                />
                {loadingDsmClients && <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Caricamento anagrafica da Desearq Studio Manager…</p>}
                {dsmClientsError && <p style={{ fontSize: 12, color: C.maroon, margin: 0 }}>{dsmClientsError}</p>}
                {!loadingDsmClients && !dsmClientsError && dsmClients && (
                  <div style={{ maxHeight: 220, overflowY: 'auto' }}>
                    {dsmClients
                      .filter((c) => !clientSearch.trim() || (c.name || '').toLowerCase().includes(clientSearch.trim().toLowerCase()))
                      .map((c) => (
                        <div
                          key={c.id}
                          onClick={() => importClientSheet(c)}
                          style={{ padding: '8px 10px', borderRadius: 8, cursor: 'pointer', display: 'flex', justifyContent: 'space-between', gap: 8, background: C.white, marginBottom: 4 }}
                          onMouseEnter={(e) => { e.currentTarget.style.background = '#efe8db'; }}
                          onMouseLeave={(e) => { e.currentTarget.style.background = C.white; }}
                        >
                          <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{c.name}</span>
                          <span style={{ fontSize: 11, color: C.gray }}>{c.email || c.phone || ''}</span>
                        </div>
                      ))}
                    {dsmClients.filter((c) => !clientSearch.trim() || (c.name || '').toLowerCase().includes(clientSearch.trim().toLowerCase())).length === 0 && (
                      <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessun cliente trovato.</p>
                    )}
                  </div>
                )}
                <button onClick={() => setShowClientImport(false)} style={{ ...rowBtnStyle, marginTop: 8 }}>Chiudi</button>
              </div>
            )}

            {!project.clientSheet ? (
              <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna scheda cliente ancora collegata a questo progetto. Importala da Desearq Studio Manager oppure lascia questa sezione vuota.</p>
            ) : (
              <>
                <p style={{ fontSize: 11, color: C.gray, margin: '0 0 10px' }}>
                  Importata da Desearq Studio Manager il {project.clientSheet.importedAt}. I campi restano modificabili qui senza alcun effetto su Desearq Studio Manager.
                </p>
                <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Nome / Ragione sociale</label>
                    <input value={project.clientSheet.name || ''} onChange={(e) => updateClientSheetField('name', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Indirizzo</label>
                    <input value={project.clientSheet.address || ''} onChange={(e) => updateClientSheetField('address', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                </div>
                <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr', gap: 12 }}>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Telefono</label>
                    <input value={project.clientSheet.phone || ''} onChange={(e) => updateClientSheetField('phone', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Email</label>
                    <input value={project.clientSheet.email || ''} onChange={(e) => updateClientSheetField('email', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>P.IVA</label>
                    <input value={project.clientSheet.piva || ''} onChange={(e) => updateClientSheetField('piva', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Codice fiscale</label>
                    <input value={project.clientSheet.cf || ''} onChange={(e) => updateClientSheetField('cf', e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, marginTop: 4 }} />
                  </div>
                </div>
              </>
            )}
          </div>

          <div style={{ display: 'flex', gap: 18, alignItems: 'flex-start', flexWrap: 'wrap' }}>
            <div className="listino-catalog-card" style={{ ...card, width: 300, maxWidth: '100%', flexShrink: 0, border: `2px solid ${C.maroon}`, position: 'sticky', top: 'calc(var(--header-h, 68px) + 12px)', alignSelf: 'flex-start', maxHeight: 'calc(100vh - var(--header-h, 68px) - 24px)', display: 'flex', flexDirection: 'column', gap: 10 }}>
              <div style={{ display: 'flex', flexDirection: 'column', flex: sidebarOpen.listino ? 1 : '0 0 auto', minHeight: sidebarOpen.listino ? 60 : 'auto', overflow: 'hidden' }}>
                <div onClick={() => toggleSidebarSection('listino')} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', flexShrink: 0 }}>
                  <p style={{ fontSize: 14, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Listino</p>
                  <span style={{ fontSize: 11, color: C.gray }}>{sidebarOpen.listino ? '▾ comprimi' : '▸ espandi'}</span>
                </div>
                {sidebarOpen.listino && (
                  <>
                    <select value={listinoId} onChange={(e) => setListinoId(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '6px 0 10px', flexShrink: 0 }}>
                      {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
                    </select>
                    <p style={{ fontSize: 11, color: C.gray, margin: '0 0 10px', flexShrink: 0 }}>Apri le categorie per trovare la voce giusta: trascinala nel computo a destra, oppure tocca + per aggiungerla subito (utile su tablet e smartphone).</p>
                    {onUpdateListini && (
                      <button onClick={() => setListinoVoceModal({ mode: 'new' })} disabled={listinoLocations.length === 0} title={listinoLocations.length === 0 ? 'Crea prima almeno una sottocategoria nel listino' : ''} style={{ alignSelf: 'flex-start', border: `1px solid ${C.maroon}`, background: C.white, color: C.maroon, borderRadius: 999, padding: '7px 14px', fontSize: 12, fontWeight: 500, cursor: listinoLocations.length === 0 ? 'default' : 'pointer', fontFamily: FONT, margin: '0 0 10px', flexShrink: 0, opacity: listinoLocations.length === 0 ? 0.5 : 1 }}>+ Nuova voce di listino</button>
                    )}
                    <div className="listino-catalog-scroll" style={{ flex: 1, minHeight: 0, overflowY: 'auto' }}>
                      <DraggableCatalogTree listino={activeListino} onAdd={openVoceFromListino} onEdit={onUpdateListini ? (vpath) => setListinoVoceModal({ mode: 'edit', path: vpath }) : null} />
                    </div>
                  </>
                )}
              </div>

              <div style={{ borderTop: `1px solid ${C.paleGray}`, flexShrink: 0 }} />

              <div style={{ display: 'flex', flexDirection: 'column', flex: sidebarOpen.fornitori ? 1 : '0 0 auto', minHeight: sidebarOpen.fornitori ? 60 : 'auto', overflow: 'hidden' }}>
                <div onClick={() => toggleSidebarSection('fornitori')} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', flexShrink: 0 }}>
                  <p style={{ fontSize: 14, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Fornitori</p>
                  <span style={{ fontSize: 11, color: C.gray }}>{sidebarOpen.fornitori ? '▾ comprimi' : '▸ espandi'}</span>
                </div>
                {sidebarOpen.fornitori && (
                  <>
                    <p style={{ fontSize: 11, color: C.gray, margin: '6px 0 10px', flexShrink: 0 }}>Catalogo prodotti/fornitori: tocca + per aggiungere al computo la voce con il prezzo scelto.</p>
                    <div className="listino-catalog-scroll" style={{ flex: 1, minHeight: 0, overflowY: 'auto' }}>
                      <FornitoriSidebarTree catalog={fornitoriCatalog} onAdd={addFornitoreToComputo} />
                    </div>
                  </>
                )}
              </div>
            </div>

            <div style={{ flex: 1, minWidth: 320 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, marginBottom: 8, flexWrap: 'wrap', position: 'sticky', top: 'var(--header-h, 68px)', zIndex: 15, background: '#FAF9F6', padding: '8px 0' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, flex: '1 1 220px', maxWidth: 320 }}>
                  <input
                    value={voceSearch}
                    onChange={(e) => { setVoceSearch(e.target.value); setVoceSearchIdx(0); }}
                    onKeyDown={(e) => { if (e.key === 'Escape') setVoceSearch(''); }}
                    placeholder="🔍 Cerca nel computo per codice o parole…"
                    style={{ width: '100%', fontSize: 12, padding: '7px 10px', borderRadius: 999, border: `1px solid ${C.paleGray}`, background: C.white }}
                  />
                  {voceSearchTerm && (
                    <>
                      <span style={{ fontSize: 11, color: voceSearchMatches.length ? C.maroon : '#8B8178', whiteSpace: 'nowrap', fontWeight: 500 }}>
                        {voceSearchMatches.length ? `${voceSearchMatches.length} vo${voceSearchMatches.length === 1 ? 'ce' : 'ci'}` : 'nessuna voce'}
                      </span>
                      <button onClick={() => setVoceSearch('')} title="Cancella ricerca" style={{ ...iconBtn, height: 24 }}>✕</button>
                    </>
                  )}
                </div>
                <button onClick={addCustomSection} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer', whiteSpace: 'nowrap' }}>+ Nuova macrosezione</button>
              </div>
              <ComputoNoteField
                label="Nota iniziale del computo"
                hint="compare in cima al PDF e all'Excel, anche nella versione cliente"
                value={selectedRevision.noteInizio}
                onSave={(v) => applyRevisionChange(() => ({ noteInizio: v }))}
                placeholder="Es. premesse, oggetto dei lavori, condizioni generali…"
              />
              <div
                data-general-dropzone="true"
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragOver(false);
                  const data = e.dataTransfer.getData('application/json');
                  if (!data) return;
                  openVoceFromListino(JSON.parse(data));
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
                  groupedSections.map((section, sIdx) => {
                    // Righe "Sommatoria parziale" della sezione: seguono l'ordine originale delle voci
                    // (indipendente dalla sottocategoria), azzerando il totale corrente ogni volta che
                    // se ne incontra una — stessa logica IVA-aware di prima, mostrata ora in fondo alla
                    // macrocategoria invece che intercalata fra le sottocategorie.
                    const markerRows = [];
                    { let runImp = 0, runCli = 0;
                      section.items.forEach((it) => {
                        if (it.type === 'subtotal') { markerRows.push({ marker: it, runImp, runCli }); runImp = 0; runCli = 0; }
                        else { runImp += parseEuro(it.unitPriceImpresa) * parseEuro(it.qty); runCli += parseEuro(it.unitPriceCliente) * parseEuro(it.qty); }
                      });
                    }
                    // La macrosezione è eliminabile finché non contiene voci vere (le eventuali "Sommatoria
                    // parziale" rimaste vengono ripulite da removeMarkerOrEmptySection insieme alla sezione).
                    const hasRealItems = section.items.some((it) => it.type !== 'subtotal');
                    // Durante la ricerca restano visibili solo le voci che la contengono (e le loro sezioni/categorie).
                    if (voceSearchTerm && !section.items.some(itemMatchesSearch)) return null;
                    return (
                      <div
                        key={section.name}
                        style={{
                          border: `1px solid ${dragOverTarget === section.name ? C.maroon : C.paleGray}`,
                          borderRadius: 10, overflow: 'hidden', marginBottom: 14,
                          background: dragOverTarget === section.name ? 'rgba(128,20,48,0.05)' : C.white,
                        }}
                        data-macro-card={section.name}
                      >
                        {/* Il "trascina qui" per forzare questa specifica macrocategoria è solo sulla fascia
                            colorata dell'intestazione, non su tutta la card: le card riempiono quasi tutto lo
                            schermo una volta popolate, quindi se l'intero corpo fosse un bersaglio valido
                            qualunque voce trascinata (anche di un'altra macrocategoria del listino) finirebbe
                            "per sbaglio" in quella visualizzata al momento. Lasciando libero il corpo, un
                            rilascio lì risale fino alla zona tratteggiata generale, che usa la macrocategoria
                            corretta della voce nel listino. */}
                        <div
                          onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOverTarget(section.name); }}
                          onDragLeave={() => setDragOverTarget((t) => (t === section.name ? null : t))}
                          onDrop={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            setDragOverTarget(null);
                            const data = e.dataTransfer.getData('application/json');
                            if (!data) return;
                            addVoceToTarget(JSON.parse(data), section.name);
                          }}
                          style={{ background: section.color, color: C.white, padding: '8px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 6 }}
                        >
                          <span style={{ fontWeight: 700, fontSize: 13, fontFamily: FONT }}>{section.name}</span>
                          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                            <button onClick={() => moveMacroSection(section.name, 'up')} disabled={sIdx === 0} style={{ ...iconBtn, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none', opacity: sIdx === 0 ? 0.4 : 1 }}>▲</button>
                            <button onClick={() => moveMacroSection(section.name, 'down')} disabled={sIdx === groupedSections.length - 1} style={{ ...iconBtn, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none', opacity: sIdx === groupedSections.length - 1 ? 0.4 : 1 }}>▼</button>
                            <button onClick={() => renameSection(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>✎ Rinomina</button>
                            <button onClick={() => addCategoria(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>+ Categoria</button>
                            <button onClick={() => addPartialSubtotal(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>+ Sommatoria parziale</button>
                            {!hasRealItems && (
                              <button onClick={() => removeMarkerOrEmptySection(section.name)} style={{ ...rowBtnStyle, background: 'rgba(255,255,255,0.15)', color: C.white, border: 'none' }}>🗑</button>
                            )}
                          </div>
                        </div>

                        {section.categorie.length === 0 ? (
                          <p style={{ fontSize: 12, color: C.gray, padding: '10px 14px' }}>Nessuna categoria ancora in questa macrocategoria.</p>
                        ) : section.categorie.map((cat, catIdx) => {
                          const catKey = `${section.name}|${cat.name}`;
                          if (voceSearchTerm && !cat.items.some(itemMatchesSearch)) return null;
                          return (
                            <div
                              key={cat.name}
                              style={{
                                borderTop: `1px solid ${C.paleGray}`,
                                background: dragOverTarget === catKey ? 'rgba(128,20,48,0.06)' : 'transparent',
                              }}
                              data-categoria-row={catKey}
                            >
                              {/* Come per la macrocategoria: solo la fascia di intestazione della categoria è
                                  un bersaglio di trascinamento "forzato", non l'intero blocco sotto. */}
                              <div
                                onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOverTarget(catKey); }}
                                onDragLeave={() => setDragOverTarget((t) => (t === catKey ? null : t))}
                                onDrop={(e) => {
                                  e.preventDefault();
                                  e.stopPropagation();
                                  setDragOverTarget(null);
                                  const data = e.dataTransfer.getData('application/json');
                                  if (!data) return;
                                  addVoceToTarget(JSON.parse(data), section.name, cat.name);
                                }}
                                style={{ background: '#efe8db', padding: '7px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 6, borderLeft: `4px solid ${section.color}` }}
                              >
                                <span style={{ fontWeight: 700, fontSize: 12.5, color: C.black }}>{cat.name}</span>
                                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                                  <button onClick={() => moveCategoria(section.name, cat.name, 'up')} disabled={catIdx === 0} style={{ ...iconBtn, height: 22, opacity: catIdx === 0 ? 0.4 : 1 }}>▲</button>
                                  <button onClick={() => moveCategoria(section.name, cat.name, 'down')} disabled={catIdx === section.categorie.length - 1} style={{ ...iconBtn, height: 22, opacity: catIdx === section.categorie.length - 1 ? 0.4 : 1 }}>▼</button>
                                  <button onClick={() => renameCategoria(section.name, cat.name)} style={rowBtnStyle}>✎ Rinomina</button>
                                  <button onClick={() => setVoceComputoCtx({ macroName: section.name, categoriaName: cat.name, sottoName: 'Generale', initialItem: null })} style={{ ...rowBtnStyle, background: C.maroon, color: C.white, border: 'none' }}>+ Voce</button>
                                  {cat.items.length === 0 && (
                                    <button onClick={() => removeCategoria(section.name, cat.name)} style={rowBtnStyle}>🗑</button>
                                  )}
                                </div>
                              </div>

                              {cat.items.length === 0 ? (
                                <p style={{ fontSize: 12, color: C.gray, padding: '8px 14px' }}>Nessuna voce ancora in questa categoria.</p>
                              ) : (
                                <div className="table-scroll">
                                <table style={{ width: '100%', minWidth: 960, borderCollapse: 'collapse', fontSize: 12 }}>
                                  <thead>
                                    <tr style={{ textAlign: 'left', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                                      <th style={{ padding: '8px 6px' }}></th>
                                      <th style={{ padding: '8px 6px' }}>Codice</th>
                                      <th style={{ padding: '8px 6px' }}>Descrizione</th>
                                      <th style={{ padding: '8px 6px', textAlign: 'right' }}>Quantità</th>
                                      <th style={{ padding: '8px 6px' }}>U.M.</th>
                                      <th style={{ padding: '8px 6px', textAlign: 'right' }}>Costo unitario impresa</th>
                                      <th style={{ padding: '8px 6px', textAlign: 'right' }}>Totale impresa</th>
                                      <th style={{ padding: '8px 6px', textAlign: 'right' }}>Costo unitario cliente</th>
                                      <th style={{ padding: '8px 6px', textAlign: 'right' }}>Totale cliente</th>
                                      <th style={{ padding: '8px 6px' }}>Categoria</th>
                                      <th style={{ padding: '8px 6px' }}>Macrosezione</th>
                                      <th style={{ padding: '8px 6px' }}></th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {cat.items.map((it) => {
                                      if (voceSearchTerm && !itemMatchesSearch(it)) return null;
                                      const detailRows = (it.misurazioni || []).flatMap((g) => g.rows || []);
                                      const hasDetail = detailRows.length > 0 || !!it.note;
                                      const isExpanded = !!expandedItems[it.id];
                                      const itemBookmarks = (selectedRevision.bookmarks || []).filter((b) => b.itemId === it.id);
                                      const myBookmarkHere = itemBookmarks.some((b) => b.key === myBookmarkKey);
                                      const searchMatchIdx = -1; // con la ricerca "a filtro" le voci non trovate sono nascoste: niente più evidenziazione
                                      const isCurrentSearchMatch = false;
                                      const qtyIsZero = parseEuro(it.qty) === 0;
                                      const impresaIsZero = parseEuro(it.unitPriceImpresa) === 0;
                                      const clienteIsZero = parseEuro(it.unitPriceCliente) === 0;
                                      const hasZeroValue = qtyIsZero || impresaIsZero || clienteIsZero;
                                      const rowBg = isCurrentSearchMatch ? 'rgba(201,138,46,0.30)' : searchMatchIdx >= 0 ? 'rgba(201,138,46,0.14)' : hasZeroValue ? 'rgba(179,69,43,0.10)' : myBookmarkHere ? 'rgba(128,20,48,0.05)' : 'transparent';
                                      return (
                                      <React.Fragment key={it.id}>
                                      <tr id={`computo-item-${it.id}`} className="no-print-zero-flag" title={hasZeroValue ? 'Un valore di questa voce è a zero (quantità o prezzo) — solo a schermo, per segnalartelo: non compare né nel PDF né nell\'export Excel.' : undefined} style={{ borderTop: `1px solid ${C.paleGray}`, background: rowBg }}>
                                        <td style={{ padding: '8px 6px' }}>
                                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                            {hasDetail && (
                                              <button
                                                onClick={() => setExpandedItems((ex) => ({ ...ex, [it.id]: !ex[it.id] }))}
                                                title="Mostra/nascondi dettaglio misurazioni e note"
                                                style={{ ...iconBtn, height: 18, fontSize: 9, lineHeight: '16px' }}
                                              >
                                                {isExpanded ? '▾' : '▸'}
                                              </button>
                                            )}
                                            <button onClick={() => moveItemInSection(it.id, 'up')} style={{ ...iconBtn, height: 18, fontSize: 9, lineHeight: '16px' }}>▲</button>
                                            <button onClick={() => moveItemInSection(it.id, 'down')} style={{ ...iconBtn, height: 18, fontSize: 9, lineHeight: '16px' }}>▼</button>
                                            <button
                                              onClick={() => toggleBookmark(it.id)}
                                              title={myBookmarkHere ? 'Rimuovi il tuo segnalibro da questa voce' : 'Segna qui il tuo punto di lettura (sposta il tuo segnalibro)'}
                                              style={{ ...iconBtn, height: 18, fontSize: 10, lineHeight: '16px', color: myBookmarkHere ? C.maroon : C.midGray }}
                                            >
                                              {myBookmarkHere ? '🔖' : '📑'}
                                            </button>
                                          </div>
                                        </td>
                                        <td style={{ padding: '8px 6px', fontWeight: 700, color: C.black }}>{it.code}</td>
                                        <td style={{ padding: '8px 6px', color: C.midGray }}>
                                          {it.desc}
                                          {it.qtyLink && (() => {
                                            const src = realItems.find((x) => String(x.id) === String(it.qtyLink.sourceId));
                                            return (
                                              <p style={{ margin: '3px 0 0', fontSize: 10.5, color: src ? C.midGray : '#B3452B' }}>
                                                {src ? <>🔗 Quantità da {src.code || 'voce'} — {String(src.desc || '').slice(0, 40)}{Number(it.qtyLink.factor) !== 1 ? ` × ${String(it.qtyLink.factor).replace('.', ',')}` : ''}</> : '⚠️ La voce da cui prendeva la quantità è stata eliminata: quantità ferma all\'ultimo valore'}
                                              </p>
                                            );
                                          })()}
                                          {it.noteCliente && <p style={{ margin: '3px 0 0', fontSize: 10.5, color: C.maroon }}>💬 Per il cliente: {it.noteCliente}</p>}
                                          {it.note && <p style={{ margin: '3px 0 0', fontSize: 10.5, color: C.gray, fontStyle: 'italic' }}>📝 {it.note}</p>}
                                          {itemBookmarks.length > 0 && (
                                            <p style={{ margin: '3px 0 0', fontSize: 10, color: C.maroon }}>
                                              🔖 {itemBookmarks.map((b) => `${b.name} (${b.date})`).join(' · ')}
                                            </p>
                                          )}
                                        </td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right' }}>
                                          {it.autoCode || it.qtyLink ? (
                                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                              <span style={{ fontWeight: 700, color: qtyIsZero ? '#B3452B' : 'inherit' }}>{it.qtyLink ? '🔗 ' : ''}{it.qty}{qtyIsZero && ' ⚠️'}</span>
                                              <button onClick={() => setVoceComputoCtx({ macroName: section.name, categoriaName: cat.name, sottoName: it.sottocategoria || 'Generale', initialItem: it })} style={{ ...iconBtn, width: 20, height: 20, fontSize: 10 }}>✎</button>
                                            </span>
                                          ) : (
                                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                              <input
                                                value={it.qty}
                                                onChange={(e) => updateQty(it.id, e.target.value)}
                                                style={{ width: 60, fontSize: 12, padding: '5px 6px', borderRadius: 6, border: `1px solid ${qtyIsZero ? '#B3452B' : C.paleGray}`, textAlign: 'right', color: qtyIsZero ? '#B3452B' : 'inherit' }}
                                              />
                                              <button
                                                onClick={() => setVoceComputoCtx({ macroName: section.name, categoriaName: cat.name, sottoName: it.sottocategoria || 'Generale', initialItem: it })}
                                                title="Apri il pannello misurazioni per questa voce (i dati del listino restano invariati finché non salvi)"
                                                style={{ ...iconBtn, width: 20, height: 20, fontSize: 10 }}
                                              >
                                                ✎
                                              </button>
                                            </span>
                                          )}
                                        </td>
                                        <td style={{ padding: '8px 6px', color: C.gray }}>{it.unit}</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', color: impresaIsZero ? '#B3452B' : 'inherit', fontWeight: impresaIsZero ? 700 : 400 }}>{it.unitPriceImpresa} €{impresaIsZero && ' ⚠️'}</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.black }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', color: clienteIsZero ? '#B3452B' : C.maroon, fontWeight: clienteIsZero ? 700 : 400 }}>{it.unitPriceCliente} €{clienteIsZero && ' ⚠️'}</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.maroon }}>{formatEuro(parseEuro(it.unitPriceCliente) * parseEuro(it.qty))}</td>
                                        <td style={{ padding: '8px 6px' }}>
                                          <select
                                            value={cat.name}
                                            onChange={(e) => {
                                              const val = e.target.value;
                                              if (val === '__new__') {
                                                const name = prompt('Nome della nuova categoria (es. Opere di demolizione, Opere di costruzione...):');
                                                if (name) moveItemToCategoria(it.id, name);
                                                return;
                                              }
                                              moveItemToCategoria(it.id, val);
                                            }}
                                            style={{ fontSize: 11, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                                          >
                                            {section.categorie.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                                            <option value="__new__">+ Nuova categoria…</option>
                                          </select>
                                        </td>
                                        <td style={{ padding: '8px 6px' }}>
                                          <select value={section.name} onChange={(e) => moveItemToSection(it.id, e.target.value)} style={{ fontSize: 11, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}>
                                            {allSectionNames.map((n) => <option key={n} value={n}>{n}</option>)}
                                          </select>
                                        </td>
                                        <td style={{ padding: '8px 6px' }}>
                                          <button onClick={() => removeItem(it.id)} style={iconBtn}>🗑</button>
                                        </td>
                                      </tr>
                                      {isExpanded && hasDetail && (
                                        <tr style={{ background: '#f7f5f0' }}>
                                          <td></td>
                                          <td colSpan={10} style={{ padding: '8px 6px 12px' }}>
                                            {detailRows.length > 0 && (
                                              <div className="table-scroll">
                                                <table style={{ width: '100%', minWidth: 420, borderCollapse: 'collapse', fontSize: 11 }}>
                                                  <thead>
                                                    <tr style={{ textAlign: 'right', color: C.gray, fontSize: 10, textTransform: 'uppercase' }}>
                                                      <th style={{ padding: '3px 6px', textAlign: 'left' }}>Segno</th>
                                                      <th style={{ padding: '3px 6px' }}>Par.ug.</th>
                                                      <th style={{ padding: '3px 6px' }}>Lunghezza</th>
                                                      <th style={{ padding: '3px 6px' }}>Larghezza</th>
                                                      <th style={{ padding: '3px 6px' }}>H / Peso</th>
                                                      <th style={{ padding: '3px 6px' }}>Valore</th>
                                                    </tr>
                                                  </thead>
                                                  <tbody>
                                                    {detailRows.map((r, ri) => (
                                                      <tr key={ri} style={{ borderTop: `1px solid ${C.paleGray}` }}>
                                                        <td style={{ padding: '3px 6px', textAlign: 'left' }}>{r.segno === '-' ? '− si detrae' : '+ somma'}</td>
                                                        <td style={{ padding: '3px 6px', textAlign: 'right' }}>{r.parUg || '—'}</td>
                                                        <td style={{ padding: '3px 6px', textAlign: 'right' }}>{r.lung || '—'}</td>
                                                        <td style={{ padding: '3px 6px', textAlign: 'right' }}>{r.larg || '—'}</td>
                                                        <td style={{ padding: '3px 6px', textAlign: 'right' }}>{r.hPeso || '—'}</td>
                                                        <td style={{ padding: '3px 6px', textAlign: 'right', fontWeight: 700 }}>{itNum(computeMisurazioneRowValue(r, it.unitaCalcolo), 2)}</td>
                                                      </tr>
                                                    ))}
                                                  </tbody>
                                                </table>
                                              </div>
                                            )}
                                            {it.note && (
                                              <p style={{ fontSize: 11, color: C.midGray, margin: detailRows.length > 0 ? '8px 0 0' : 0 }}><strong>Nota:</strong> {it.note}</p>
                                            )}
                                          </td>
                                        </tr>
                                      )}
                                      </React.Fragment>
                                      );
                                    })}
                                  </tbody>
                                </table>
                                </div>
                              )}
                            </div>
                          );
                        })}

                        {!voceSearchTerm && markerRows.length > 0 && (
                          <div style={{ borderTop: `1px solid ${C.paleGray}` }}>
                            {markerRows.map(({ marker, runImp, runCli }) => {
                              const hasVat = marker.vatRate !== null && marker.vatRate !== undefined;
                              const hasDiscount = !!marker.discountPct;
                              // Lo sconto di questa sommatoria si applica sul lordo prima dell'eventuale IVA.
                              const netImp = hasDiscount ? runImp * (1 - marker.discountPct / 100) : runImp;
                              const netCli = hasDiscount ? runCli * (1 - marker.discountPct / 100) : runCli;
                              const ivaImpresaPart = hasVat ? netImp * (marker.vatRate / 100) : 0;
                              const ivaClientePart = hasVat ? netCli * (marker.vatRate / 100) : 0;
                              return (
                                <div key={marker.id} style={{ padding: '8px 14px', background: 'rgba(128,20,48,0.05)', borderTop: `2px solid ${C.paleGray}`, display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
                                  <div style={{ fontWeight: 700, color: C.maroon, fontSize: 12 }}>
                                    {marker.title}
                                    <button onClick={() => editPartialSubtotal(marker.id)} style={{ ...rowBtnStyle, marginLeft: 8, padding: '1px 6px', fontSize: 10 }}>✎</button>
                                  </div>
                                  <div style={{ fontSize: 12, textAlign: 'right' }}>
                                    {hasDiscount && (
                                      <div style={{ color: C.gray, fontSize: 11 }}>
                                        Lordo: impresa {formatEuro(runImp)} · cliente {formatEuro(runCli)} — Sconto {marker.discountPct}%
                                      </div>
                                    )}
                                    {hasVat ? (
                                      <>
                                        <div>Impresa{hasDiscount ? ' netto' : ''} — IVA escl. {formatEuro(netImp)}, {marker.vatLabel} {formatEuro(ivaImpresaPart)}, IVA incl. {formatEuro(netImp + ivaImpresaPart)}</div>
                                        <div style={{ color: C.maroon }}>Cliente{hasDiscount ? ' netto' : ''} — IVA escl. {formatEuro(netCli)}, {marker.vatLabel} {formatEuro(ivaClientePart)}, IVA incl. {formatEuro(netCli + ivaClientePart)}</div>
                                      </>
                                    ) : (
                                      <div>Impresa{hasDiscount ? ' netto' : ''} {formatEuro(netImp)} · <span style={{ color: C.maroon }}>Cliente{hasDiscount ? ' netto' : ''} {formatEuro(netCli)}</span></div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {!voceSearchTerm && <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.paleGray}` }}>
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
                          {(() => {
                            const saved = sectionVatSaved[section.name];
                            const rateValue = saved && saved.rate !== undefined && saved.rate !== null ? saved.rate : (section.vat?.rate ?? 22);
                            const labelValue = saved?.label || '';
                            const isPreset = (p) => Number(section.vat?.rate) === p.rate && (p.label ? labelValue === p.label : !labelValue);
                            return (
                              <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px dashed ${C.paleGray}` }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 11, flexWrap: 'wrap' }}>
                                  <label style={{ color: C.gray }}>IVA su questa sezione:</label>
                                  <input
                                    type="number"
                                    value={rateValue}
                                    onChange={(e) => updateSectionVat(section.name, { rate: e.target.value === '' ? '' : Number(e.target.value) })}
                                    style={{ width: 60, fontSize: 11, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                                  />
                                  <span style={{ color: C.gray }}>%</span>
                                  <input
                                    value={labelValue}
                                    onChange={(e) => updateSectionVat(section.name, { label: e.target.value })}
                                    placeholder='Dicitura facoltativa, es. "Esente IVA"'
                                    style={{ flex: '1 1 200px', minWidth: 160, fontSize: 11, padding: '4px 8px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                                  />
                                </div>
                                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 6 }}>
                                  {VAT_PRESETS.map((p) => (
                                    <button
                                      key={p.text}
                                      onClick={() => updateSectionVat(section.name, { rate: p.rate, label: p.label })}
                                      style={{ ...rowBtnStyle, fontSize: 10.5, padding: '3px 9px', ...(isPreset(p) ? { background: '#2B2321', color: '#fff', borderColor: '#2B2321' } : {}) }}
                                    >
                                      {p.text}
                                    </button>
                                  ))}
                                </div>
                                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', fontSize: 11.5, marginTop: 8 }}>
                                  <span style={{ color: C.black }}>
                                    Impresa — IVA escl. {formatEuro(section.netImpresa)} · {section.vat?.label} {formatEuro(section.ivaImpresa || 0)} · <strong>IVA incl. {formatEuro(section.totImpresa || 0)}</strong>
                                  </span>
                                  <span style={{ color: C.maroon }}>
                                    Cliente — IVA escl. {formatEuro(section.netCliente)} · {section.vat?.label} {formatEuro(section.ivaCliente || 0)} · <strong>IVA incl. {formatEuro(section.totCliente || 0)}</strong>
                                  </span>
                                </div>
                              </div>
                            );
                          })()}
                        </div>}
                      </div>
                    );
                  })
                )}
                {listinoVoceModal && activeListino && (
                  <VoceModal
                    locations={listinoLocations}
                    initialLocationIdx={listinoVoceModal.mode === 'edit' ? Math.max(0, listinoLocations.findIndex((l) => l.path.join() === listinoVoceModal.path.slice(0, 3).join())) : 0}
                    initialVoce={listinoVoceModal.mode === 'edit' ? activeListino.macros[listinoVoceModal.path[0]].categorie[listinoVoceModal.path[1]].sottocategorie[listinoVoceModal.path[2]].voci[listinoVoceModal.path[3]] : null}
                    offerAddToComputo
                    onClose={() => setListinoVoceModal(null)}
                    onSave={saveListinoVoce}
                  />
                )}
                {voceComputoCtx && (
                  <VoceComputoModal
                    macroName={voceComputoCtx.macroName}
                    categoriaName={voceComputoCtx.categoriaName}
                    sottoName={voceComputoCtx.sottoName}
                    initialItem={voceComputoCtx.initialItem}
                    prefill={voceComputoCtx.prefill}
                    linkCandidates={realItems.filter((x) => String(x.id) !== String(voceComputoCtx.initialItem?.id))}
                    mergeCandidates={(groupedSections.find((s) => s.name === voceComputoCtx.macroName)?.categorie.find((c) => c.name === (voceComputoCtx.categoriaName || 'Generale'))?.items || []).filter((it) => it.autoCode && it.id !== voceComputoCtx.initialItem?.id)}
                    onClose={() => setVoceComputoCtx(null)}
                    onSave={(voceData) => saveVoceComputo(voceComputoCtx.macroName, voceComputoCtx.categoriaName || 'Generale', voceComputoCtx.sottoName, voceData)}
                  />
                )}
              </div>

              <ComputoNoteField
                label="Nota finale del computo"
                hint="compare in fondo al PDF e all'Excel, dopo i totali"
                value={selectedRevision.noteFine}
                onSave={(v) => applyRevisionChange(() => ({ noteFine: v }))}
                placeholder="Es. esclusioni, modalità di pagamento, validità dell'offerta…"
              />
              <p style={{ fontSize: 11.5, color: C.gray, margin: '0 0 10px', maxWidth: 660 }}>
                L'IVA si imposta in fondo a ogni sezione: il totale generale qui sotto somma l'IVA delle singole sezioni, divisa per aliquota.
              </p>
              <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap' }}>
                <div style={{ ...card, maxWidth: 320, flex: 1 }}>
                  <p style={{ fontWeight: 700, fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Quadro economico — Impresa</p>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginBottom: 8 }}>
                    <span style={{ color: C.gray }}>Totale IVA esclusa</span>
                    <span style={{ fontWeight: 700, color: C.black }}>{formatEuro(importoLavori)}</span>
                  </div>
                  <div style={{ marginBottom: 10, paddingBottom: 10, borderBottom: `1px solid ${C.paleGray}` }}>
                    {economics.breakdown.length === 0 && <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13 }}><span style={{ color: C.gray }}>IVA</span><span style={{ color: C.black }}>{formatEuro(0)}</span></div>}
                    {economics.breakdown.map((b, i) => (
                      <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginTop: i ? 6 : 0 }}>
                        <span style={{ color: C.gray }}>{b.label}</span>
                        <span style={{ color: C.black }}>{formatEuro(b.ivaImpresa)}</span>
                      </div>
                    ))}
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
                  <div style={{ marginBottom: 10, paddingBottom: 10, borderBottom: `1px solid ${C.paleGray}` }}>
                    {economics.breakdown.length === 0 && <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13 }}><span style={{ color: C.gray }}>IVA</span><span style={{ color: C.black }}>{formatEuro(0)}</span></div>}
                    {economics.breakdown.map((b, i) => (
                      <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, marginTop: i ? 6 : 0 }}>
                        <span style={{ color: C.gray }}>{b.label}</span>
                        <span style={{ color: C.black }}>{formatEuro(b.ivaCliente)}</span>
                      </div>
                    ))}
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
                <select value={selRevA} onChange={(e) => setSelRevA(revisions.find((r) => String(r.id) === e.target.value)?.id ?? selRevA)} style={{ fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
                  {revisions.map((r) => <option key={r.id} value={r.id}>{r.customName || r.label} · {r.dateModified}</option>)}
                </select>
                <span style={{ color: C.gray }}>→</span>
                <select value={selRevB} onChange={(e) => setSelRevB(revisions.find((r) => String(r.id) === e.target.value)?.id ?? selRevB)} style={{ fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
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

      {(() => {
        const approvedRevision = latestApprovedRevision(project);
        if (!approvedRevision) return null;
        const sections = computeSectionTotals(approvedRevision);
        return (
          <div style={{ ...card, marginBottom: 18 }}>
            <div style={{ marginBottom: 12 }}>
              <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Stati avanzamento pagamenti</p>
              <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>
                Per categoria del computo approvato ({approvedRevision.customName || approvedRevision.label}): impresa assegnata, fatture e residuo. Si aggiorna da solo quando approvi una nuova revisione, e finisce in automatico nel Portale Clienti.
              </p>
            </div>
            {sections.length === 0 && <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Il computo approvato non ha ancora voci.</p>}
            {fatturaCtx && (() => {
              const entry = cantiereSal[fatturaCtx.categoria] || { fatture: [] };
              const giaFatturato = (entry.fatture || []).filter((f) => f.id !== fatturaCtx.fattura?.id).reduce((n, f) => n + (f.importo || 0), 0);
              return (
                <FatturaSalModal
                  categoria={fatturaCtx.categoria}
                  base={fatturaCtx.base}
                  giaFatturato={giaFatturato}
                  numero={(entry.fatture || []).length + 1}
                  fattura={fatturaCtx.fattura}
                  onClose={() => setFatturaCtx(null)}
                  onSave={(data) => saveFatturaCategoria(fatturaCtx.categoria, data)}
                />
              );
            })()}
            {sections.map((s, si) => {
              const entry = cantiereSal[s.name] || { impresa: '', fatture: [] };
              const fatture = entry.fatture || [];
              const fatturato = fatture.reduce((sum, f) => sum + (f.importo || 0), 0);
              const pagato = fatture.filter((f) => f.stato === 'Pagata').reduce((sum, f) => sum + (f.importo || 0), 0);
              const residuo = Math.max(s.netCliente - fatturato, 0);
              return (
                <div key={s.name} style={{ padding: '14px 0', borderTop: si === 0 ? 'none' : `1px solid ${C.paleGray}` }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: 8 }}>
                    <div>
                      <p style={{ fontWeight: 700, fontSize: 14, margin: 0, color: C.black }}>{s.name}</p>
                      <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>
                        Totale categoria {formatEuro(s.netCliente)}{entry.impresa ? ` · Impresa: ${entry.impresa}` : ' · nessuna impresa assegnata'}
                      </p>
                    </div>
                    <div style={{ display: 'flex', gap: 6 }}>
                      <button onClick={() => setImpresaCategoria(s.name)} style={rowBtnStyle}>{entry.impresa ? '✎ Impresa' : '+ Impresa'}</button>
                      <button onClick={() => addFatturaCategoria(s.name, s.netCliente)} style={rowBtnStyle}>+ Fattura</button>
                    </div>
                  </div>

                  {fatture.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      {fatture.map((f) => (
                        <div key={f.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0', gap: 8, flexWrap: 'wrap' }}>
                          <span style={{ fontSize: 12, color: C.black }}>
                            {f.label}{f.data ? <span style={{ color: C.gray }}> · {new Date(f.data).toLocaleDateString('it-IT')}</span> : null} — <strong>{formatEuro(f.importo)}</strong>
                            <span style={{ color: C.gray }}> ({itNum(f.percentuale != null ? f.percentuale : (s.netCliente ? (f.importo / s.netCliente) * 100 : 0), 2)}%)</span>
                          </span>
                          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                            <button onClick={() => setFatturaCtx({ categoria: s.name, base: s.netCliente, fattura: f })} style={rowBtnStyle} title="Modifica fattura">✎</button>
                            <select value={f.stato} onChange={(e) => updateFatturaStato(s.name, f.id, e.target.value)} style={{ fontSize: 11, fontWeight: 600, padding: '4px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}`, color: C.midGray }}>
                              {['Emessa', 'Pagata', 'Da pagare', 'Scaduta'].map((opt) => <option key={opt} value={opt}>{opt}</option>)}
                            </select>
                            <button onClick={() => removeFatturaCategoria(s.name, f.id)} style={rowBtnStyle}>🗑</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  <div style={{ display: 'flex', gap: 18, marginTop: 10, fontSize: 12, flexWrap: 'wrap' }}>
                    <span style={{ color: C.gray }}>Fatturato: <strong style={{ color: C.black }}>{formatEuro(fatturato)}</strong>{s.netCliente ? ` (${itNum((fatturato / s.netCliente) * 100, 1)}%)` : ''}</span>
                    <span style={{ color: C.gray }}>Pagato: <strong style={{ color: C.black }}>{formatEuro(pagato)}</strong></span>
                    <span style={{ color: C.gray }}>Residuo da fatturare: <strong style={{ color: residuo > 0.005 ? C.maroon : C.black }}>{formatEuro(residuo)}</strong></span>
                  </div>
                </div>
              );
            })}
          </div>
        );
      })()}

      <div style={{ ...card, marginBottom: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
          <div>
            <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Cronoprogramma</p>
            <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>
              Lavorazioni di cantiere con data di inizio e durata. "Genera dal computo" propone una lavorazione per ogni categoria del computo approvato, già collegata alle sue voci — modificabile e integrabile a mano.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <label style={{ fontSize: 10, color: C.gray, display: 'block', marginBottom: 2 }}>Inizio lavori</label>
              <input type="date" value={crono.startDate} onChange={(e) => updateCrono({ startDate: e.target.value })} style={{ fontSize: 12, padding: '7px 8px', borderRadius: 8, border: `1px solid ${C.paleGray}` }} />
            </div>
            <button onClick={generateCronoFromComputo} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Genera dal computo</button>
            <button onClick={addManualCronoTask} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Lavorazione manuale</button>
          </div>
        </div>

        {crono.tasks.length === 0 ? (
          <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna lavorazione ancora. Genera il cronoprogramma dal computo approvato, oppure aggiungine una a mano.</p>
        ) : (() => {
          const maxDay = Math.max(14, ...crono.tasks.map((t) => t.startDay + t.durationDays)) + 2;
          const todayOffset = daysBetweenISO(crono.startDate, todayISO());
          const weekMarks = [];
          for (let d = 0; d <= maxDay; d += 7) weekMarks.push(d);
          return (
            <div>
              <div style={{ position: 'relative', height: 18, marginBottom: 4, marginLeft: 180 }}>
                {weekMarks.map((d) => (
                  <span key={d} style={{ position: 'absolute', left: `${(d / maxDay) * 100}%`, fontSize: 10, color: C.gray, transform: 'translateX(-50%)' }}>
                    {formatDateIT(addDaysISO(crono.startDate, d))}
                  </span>
                ))}
              </div>
              {crono.tasks.map((t) => {
                const linkedCount = (t.linkedCodes || []).length;
                const isExpanded = expandedCronoTask === t.id;
                return (
                  <div key={t.id} style={{ padding: '8px 0', borderTop: `1px solid ${C.paleGray}` }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                      <div style={{ width: 170, flexShrink: 0, display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                        <span style={{ width: 9, height: 9, borderRadius: 999, background: t.color, flexShrink: 0 }} />
                        <button onClick={() => renameCronoTask(t.id, t.name)} title="Rinomina" style={{ background: 'none', border: 'none', padding: 0, textAlign: 'left', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.name}</button>
                      </div>
                      <div style={{ flex: 1, minWidth: 120, position: 'relative', height: 22, background: C.surfaceSubtle, borderRadius: 6 }}>
                        {todayOffset >= 0 && todayOffset <= maxDay && (
                          <div style={{ position: 'absolute', left: `${(todayOffset / maxDay) * 100}%`, top: -2, bottom: -2, width: 1, background: C.maroon, opacity: 0.5 }} />
                        )}
                        <div
                          title={`${formatDateIT(addDaysISO(crono.startDate, t.startDay))} → ${formatDateIT(addDaysISO(crono.startDate, t.startDay + t.durationDays))}`}
                          style={{ position: 'absolute', left: `${(t.startDay / maxDay) * 100}%`, width: `${Math.max((t.durationDays / maxDay) * 100, 2)}%`, top: 2, bottom: 2, background: t.color, borderRadius: 4 }}
                        />
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 4, flexShrink: 0 }}>
                        <input type="number" min={0} value={t.startDay} onChange={(e) => updateCronoTask(t.id, { startDay: Math.max(0, Number(e.target.value) || 0) })} style={{ width: 46, fontSize: 11, padding: '4px 5px', borderRadius: 6, border: `1px solid ${C.paleGray}` }} title="Giorni dall'inizio lavori" />
                        <span style={{ fontSize: 10, color: C.gray }}>×</span>
                        <input type="number" min={1} value={t.durationDays} onChange={(e) => updateCronoTask(t.id, { durationDays: Math.max(1, Number(e.target.value) || 1) })} style={{ width: 46, fontSize: 11, padding: '4px 5px', borderRadius: 6, border: `1px solid ${C.paleGray}` }} title="Durata in giorni" />
                        <span style={{ fontSize: 10, color: C.gray }}>gg</span>
                      </div>
                      <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                        <button onClick={() => setExpandedCronoTask(isExpanded ? null : t.id)} style={rowBtnStyle}>{linkedCount > 0 ? `🔗 ${linkedCount}` : '🔗 Collega'}</button>
                        <button onClick={() => removeCronoTask(t.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
                      </div>
                    </div>
                    {isExpanded && (
                      <div style={{ marginTop: 8, marginLeft: 180, padding: 10, background: C.surfaceSubtle, borderRadius: 8 }}>
                        <p style={{ fontSize: 11, color: C.gray, margin: '0 0 6px' }}>Voci del computo collegate a questa lavorazione (per categoria):</p>
                        {cronoLinkableCategories.length === 0 ? (
                          <p style={{ fontSize: 11, color: C.gray, margin: 0 }}>Il computo non ha ancora voci da collegare.</p>
                        ) : (
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                            {cronoLinkableCategories.map((cat) => {
                              const allLinked = cat.linkedCodes.every((c) => (t.linkedCodes || []).includes(c));
                              return (
                                <button
                                  key={cat.name}
                                  onClick={() => toggleCronoTaskCategoria(t.id, cat)}
                                  style={{
                                    fontSize: 11, padding: '5px 10px', borderRadius: 999, cursor: 'pointer',
                                    border: `1px solid ${allLinked ? cat.color : C.paleGray}`,
                                    background: allLinked ? cat.color : C.white,
                                    color: allLinked ? C.white : C.black, fontWeight: 600,
                                  }}
                                >
                                  {allLinked ? '✓ ' : ''}{cat.name}
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}
      </div>

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

      <div style={{ ...card, marginBottom: 18, display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
        <div>
          <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Rilievo planimetrico</p>
          <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>
            {(project.planimetrie || []).length > 0
              ? `${project.planimetrie.length} planimetria/e caricata/e per questo progetto. Carica, calibra e disegna aree/lunghezze/volumi collegate al listino dalla sezione dedicata nel menu laterale.`
              : 'Carica una planimetria, calibra la scala e disegna aree (m²/m³) e lunghezze (ml) collegate al listino, dalla sezione dedicata nel menu laterale — finiscono in automatico nel computo.'}
          </p>
        </div>
        {onOpenRilievo && (
          <button onClick={onOpenRilievo} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer', flexShrink: 0 }}>
            <MapPin size={13} strokeWidth={2} style={{ verticalAlign: -2, marginRight: 4 }} />
            Vai al Rilievo planimetrico
          </button>
        )}
      </div>

      <div style={card}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <p style={{ fontWeight: 700, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Team assegnato</p>
          <button onClick={addTeamMember} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Aggiungi persona</button>
        </div>
        {(project.team || []).length === 0 && <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna persona assegnata a questo progetto.</p>}
        {(project.team || []).map((m, i) => {
          const t = memberFor(m);
          return (
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, padding: '10px 0', borderBottom: i < (project.team || []).length - 1 ? `1px solid ${C.paleGray}` : 'none' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, minWidth: 0 }}>
                <MemberAvatar member={t || { name: m.name }} size={36} />
                <div style={{ minWidth: 0 }}>
                  <p style={{ fontWeight: 600, fontSize: 13, margin: 0, color: C.black }}>{t?.name || m.name}</p>
                  <p style={{ fontSize: 11, color: C.gray, margin: '2px 0 0' }}>
                    {m.role || t?.job_title || t?.role}{(t?.email || m.email) ? ` · ${t?.email || m.email}` : ''}
                    {!t && <span style={{ color: C.maroon }}> · non presente nel Team</span>}
                  </p>
                </div>
              </div>
              <button onClick={() => removeTeamMember(i)} style={rowBtnStyle}>🗑</button>
            </div>
          );
        })}
        {showTeamPicker && (
          <TeamPickerModal members={teamMembers} assigned={(project.team || []).map((m) => memberFor(m)?.id).filter(Boolean)} onClose={() => setShowTeamPicker(false)} onConfirm={assignMembers} />
        )}
      </div>
    </div>
  );
}

const rowBtnStyle = { border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 6, fontSize: 11, fontWeight: 600, padding: '5px 10px', cursor: 'pointer', color: C.midGray };

// Aggiunge una voce di listino alla revisione più recente di un progetto, fuori dal contesto della scheda
// di dettaglio (usata dalla sezione "Rilievo planimetrico", che lavora su un progetto scelto dall'utente
// senza aprire la relativa pagina di dettaglio). Stessa logica di "addComputoItem" lì dentro, ma ritorna
// il progetto aggiornato invece di affidarsi allo stato locale di quella pagina (selectedRevision, ecc.).
// Se il progetto non ha ancora nessuna revisione, ne crea una prima automaticamente ("Revisione 1", stessa
// forma usata da "startComputo" nella scheda di dettaglio) invece di bloccare l'utente: il computo nasce
// così da solo, misura dopo misura, man mano che si rileva dalle planimetrie.
function addComputoItemToProject(project, voce, qty = '1') {
  let revisions = project.revisions || [];
  if (revisions.length === 0) {
    revisions = [{ id: Date.now(), label: 'Revisione 1', customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], items: [], extraSections: [], total: '0,00 €', totalCliente: '0,00 €' }];
  }
  const latestRevision = revisions[revisions.length - 1];
  const impresaVal = voce.impresaValue !== undefined ? voce.impresaValue : parseEuro(voce.priceImpresa);
  const clienteVal = voce.clienteValue !== undefined ? voce.clienteValue : evalClientPrice(voce.priceCliente, impresaVal);
  const item = { id: Date.now() + Math.random(), listinoRef: voce.uid ? { listinoId: voce.listinoId, uid: voce.uid } : null, code: voce.code, desc: voce.desc, unit: voce.unit, unitPriceImpresa: formatEuro(impresaVal).replace(' €', ''), unitPriceCliente: formatEuro(clienteVal).replace(' €', ''), qty: String(qty), macro: voce.macro, section: voce.macro };
  const nextItems = regenerateItemCodes(resolveLinkedQuantities([...(latestRevision.items || []), item]));
  const realNextItems = nextItems.filter((it) => it.type !== 'subtotal');
  const total = formatEuro(sumImpresa(realNextItems));
  const totalCliente = formatEuro(sumCliente(realNextItems));
  const updatedRevisions = revisions.map((r) => (r.id === latestRevision.id ? { ...r, items: nextItems, dateModified: nowLabel(), total, totalCliente } : r));
  return { project: { ...project, revisions: updatedRevisions, value: total }, ok: true };
}

// Incrementa la quantità di una voce GIÀ presente nella revisione più recente del computo (invece di
// aggiungerne una nuova): usata quando una misura del rilievo viene attribuita a una voce già esistente
// nel computo del progetto, così la quantità rilevata si somma a quella già inserita senza duplicare la riga.
function incrementComputoItemQty(project, itemId, deltaNum) {
  const revisions = project.revisions || [];
  const latestRevision = revisions[revisions.length - 1];
  if (!latestRevision || !deltaNum) return { project, ok: false };
  let found = false;
  const nextItems = (latestRevision.items || []).map((it) => {
    if (it.id !== itemId) return it;
    found = true;
    const newQty = parseEuro(it.qty) + deltaNum;
    return { ...it, qty: newQty.toLocaleString('it-IT', { maximumFractionDigits: 2 }) };
  });
  if (!found) return { project, ok: false };
  const realItems = nextItems.filter((it) => it.type !== 'subtotal');
  const total = formatEuro(sumImpresa(realItems));
  const totalCliente = formatEuro(sumCliente(realItems));
  const updatedRevisions = revisions.map((r) => (r.id === latestRevision.id ? { ...r, items: nextItems, dateModified: nowLabel(), total, totalCliente } : r));
  return { project: { ...project, revisions: updatedRevisions, value: total }, ok: true };
}

// --- Matematica del rilievo (scala/area/lunghezza/volume), a livello di modulo e non più chiusa dentro
// RilievoPage: così la stessa logica si può richiamare sia alla creazione di una forma sia al salvataggio
// di una modifica (trascinamento punti) o di un semplice ricollegamento della voce di listino, senza doverla
// duplicare in più punti.
function clamp(n, min, max) { return Math.min(max, Math.max(min, n)); }
function metersPerPixel(pl) {
  if (!pl.scale || !pl.naturalWidth) return null;
  const { p1, p2, realDistance } = pl.scale;
  const dx = ((p2.x - p1.x) / 100) * pl.naturalWidth;
  const dy = ((p2.y - p1.y) / 100) * pl.naturalHeight;
  const pxDist = Math.sqrt(dx * dx + dy * dy);
  return pxDist > 0 ? realDistance / pxDist : null;
}
function pointToMeters(pl, mpp, pt) { return { x: (pt.x / 100) * pl.naturalWidth * mpp, y: (pt.y / 100) * pl.naturalHeight * mpp }; }
function polygonAreaM2(pl, mpp, points) {
  const mp = points.map((p) => pointToMeters(pl, mpp, p));
  let sum = 0;
  for (let i = 0; i < mp.length; i++) { const a = mp[i]; const b = mp[(i + 1) % mp.length]; sum += a.x * b.y - b.x * a.y; }
  return Math.abs(sum) / 2;
}
function polylineLengthM(pl, mpp, points) {
  const mp = points.map((p) => pointToMeters(pl, mpp, p));
  let sum = 0;
  for (let i = 0; i < mp.length - 1; i++) sum += Math.hypot(mp[i + 1].x - mp[i].x, mp[i + 1].y - mp[i].y);
  return sum;
}
// Calcola valore/unità di una forma (area o lunghezza), con eventuale altezza minima/massima (falda):
// se l'altezza massima non è data o è minore della minima, si comporta come altezza costante.
function computeShapeValue(pl, mpp, type, points, heightMin, heightMax) {
  if (type === 'area') {
    const area = polygonAreaM2(pl, mpp, points);
    if (heightMin > 0) {
      const hMax = heightMax != null && heightMax > heightMin ? heightMax : heightMin;
      return { value: area * (heightMin + hMax) / 2, unit: 'm³', height: heightMin, heightMinOut: heightMin, heightMaxOut: hMax, baseLength: null };
    }
    return { value: area, unit: 'm²', height: null, heightMinOut: null, heightMaxOut: null, baseLength: null };
  }
  const lengthM = polylineLengthM(pl, mpp, points);
  if (heightMin > 0) {
    const hMax = heightMax != null && heightMax > heightMin ? heightMax : heightMin;
    return { value: lengthM * (heightMin + hMax) / 2, unit: 'm²', height: heightMin, heightMinOut: heightMin, heightMaxOut: hMax, baseLength: lengthM };
  }
  return { value: lengthM, unit: 'ml', height: null, heightMinOut: null, heightMaxOut: null, baseLength: null };
}

// --- Modale "collega al listino": sostituisce i vecchi prompt() per codice/quantità/altezze, sia alla
// creazione di un punto/forma sia per ricollegare (o collegare per la prima volta) una forma già disegnata.
// Menu a tendina con ricerca libera su codice o descrizione, invece dell'albero completo macro/categoria/sottocategoria.
// Se il progetto ha già almeno una voce nel computo corrente (prop "computoItems"), l'utente può scegliere
// da dove prendere la voce: dal listino (comportamento originale, crea una nuova riga nel computo) oppure
// da una voce già presente in quel computo (la quantità rilevata si somma a quella già inserita, invece
// di creare una riga duplicata) — la scelta si fa qui, nel momento stesso in cui si attribuisce la misura.
function PlanLinkModal({ mode, listino, computoItems, initialCode, initialComputoItemId, initialQty, initialHeightMin, initialHeightMax, onConfirm, onCancel }) {
  const listinoItems = flattenListino(listino);
  const existingItems = computoItems || [];
  const hasExistingItems = existingItems.length > 0;
  const [source, setSource] = useState(initialComputoItemId ? 'computo' : 'listino');
  const initialVoce = initialCode ? listinoItems.find((v) => v.code === initialCode) : null;
  const initialItem = initialComputoItemId ? existingItems.find((it) => it.id === initialComputoItemId) : null;
  const [code, setCode] = useState(initialVoce?.code || null);
  const [computoItemId, setComputoItemId] = useState(initialItem?.id || null);
  const [query, setQuery] = useState(
    initialItem ? `${initialItem.code} — ${initialItem.desc}` : initialVoce ? `${initialVoce.code} — ${initialVoce.desc}` : ''
  );
  const [showList, setShowList] = useState(false);
  const [qty, setQty] = useState(initialQty ?? '1');
  const [heightMin, setHeightMin] = useState(initialHeightMin != null ? String(initialHeightMin).replace('.', ',') : '');
  const [heightMax, setHeightMax] = useState(initialHeightMax != null ? String(initialHeightMax).replace('.', ',') : '');

  const activeItems = source === 'computo' ? existingItems : listinoItems;
  const filtered = (query.trim()
    ? activeItems.filter((v) => (v.code || '').toLowerCase().includes(query.toLowerCase()) || (v.desc || '').toLowerCase().includes(query.toLowerCase()))
    : activeItems
  ).slice(0, 40);
  const selectedVoce = source === 'listino' ? listinoItems.find((v) => v.code === code) : null;
  const selectedItem = source === 'computo' ? existingItems.find((it) => it.id === computoItemId) : null;
  const selected = source === 'listino' ? selectedVoce : selectedItem;
  const hMinNum = parseEuro(heightMin || '');
  const canConfirm = mode !== 'punto' || !!selected;

  const pick = (v) => {
    if (source === 'listino') setCode(v.code); else setComputoItemId(v.id);
    setQuery(`${v.code} — ${v.desc}`);
    setShowList(false);
  };
  const clear = () => { setCode(null); setComputoItemId(null); setQuery(''); };
  const switchSource = (next) => { setSource(next); setCode(null); setComputoItemId(null); setQuery(''); };

  const confirm = () => {
    if (!canConfirm) return;
    const hMaxNum = hMinNum > 0 ? Math.max(parseEuro(heightMax || '') || hMinNum, hMinNum) : null;
    onConfirm({
      source,
      code: source === 'listino' ? (selectedVoce?.code || null) : null,
      computoItemId: source === 'computo' ? (selectedItem?.id || null) : null,
      desc: selected?.desc || null,
      qty: qty || '1',
      heightMin: hMinNum > 0 ? hMinNum : null,
      heightMax: hMaxNum,
    });
  };

  const title = mode === 'punto' ? 'Collega un punto al listino' : mode === 'area' ? 'Collega l\'area al listino' : 'Collega la lunghezza al listino';

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }} onClick={onCancel}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 420, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }} onClick={(e) => e.stopPropagation()}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 16px', color: C.black }}>{title}</h2>

        {hasExistingItems && (
          <div style={{ display: 'flex', gap: 6, margin: '0 0 14px' }}>
            <button
              onClick={() => switchSource('listino')}
              style={{ flex: 1, fontSize: 11, fontWeight: 600, padding: '7px 10px', borderRadius: 8, cursor: 'pointer', border: `1px solid ${source === 'listino' ? C.maroon : C.paleGray}`, background: source === 'listino' ? C.maroon : C.white, color: source === 'listino' ? C.white : C.black }}
            >
              Da listino
            </button>
            <button
              onClick={() => switchSource('computo')}
              style={{ flex: 1, fontSize: 11, fontWeight: 600, padding: '7px 10px', borderRadius: 8, cursor: 'pointer', border: `1px solid ${source === 'computo' ? C.maroon : C.paleGray}`, background: source === 'computo' ? C.maroon : C.white, color: source === 'computo' ? C.white : C.black }}
            >
              Dal computo di questo progetto
            </button>
          </div>
        )}

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>
          {source === 'computo' ? 'Voce già presente nel computo' : `Voce di listino${mode !== 'punto' ? ' (opzionale)' : ''}`}
        </label>
        <div style={{ position: 'relative', margin: '4px 0 14px' }}>
          <input
            value={query}
            onChange={(e) => { setQuery(e.target.value); setShowList(true); if (!e.target.value.trim()) { setCode(null); setComputoItemId(null); } }}
            onFocus={() => setShowList(true)}
            placeholder={source === 'computo' ? 'Cerca tra le voci già nel computo…' : (mode === 'punto' ? 'Cerca per codice o descrizione…' : 'Cerca per codice o descrizione… (lascia vuoto per non collegare)')}
            style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, boxSizing: 'border-box' }}
          />
          {query && <button onClick={clear} style={{ position: 'absolute', right: 8, top: 7, border: 'none', background: 'none', color: C.gray, cursor: 'pointer', fontSize: 13 }}>✕</button>}
          {showList && (
            <div style={{ position: 'absolute', zIndex: 5, top: '100%', left: 0, right: 0, background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 8, marginTop: 4, maxHeight: 220, overflowY: 'auto', boxShadow: '0 4px 16px rgba(0,0,0,0.12)' }}>
              {filtered.length === 0 && <p style={{ fontSize: 12, color: C.gray, margin: 0, padding: 10 }}>{source === 'computo' ? 'Nessuna voce ancora nel computo di questo progetto.' : 'Nessuna voce trovata.'}</p>}
              {filtered.map((v) => (
                <div key={source === 'listino' ? v.code : v.id} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(v)} style={{ padding: '8px 10px', fontSize: 12, cursor: 'pointer', borderBottom: `1px solid ${C.paleGray}` }}>
                  <strong>{v.code}</strong> — {v.desc} <span style={{ color: C.gray }}>({v.unit}{source === 'computo' ? ` · attuale ${v.qty}` : ''})</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {mode === 'punto' && (
          <>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>{source === 'computo' ? 'Quantità da aggiungere alla voce scelta' : 'Quantità'}</label>
            <input value={qty} onChange={(e) => setQty(e.target.value)} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px', boxSizing: 'border-box' }} />
          </>
        )}

        {(mode === 'area' || mode === 'lunghezza') && (
          <>
            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>
              {mode === 'area' ? 'Altezza minima in metri, per calcolare anche il volume — lascia vuoto per restare in m²' : 'Altezza minima della parete in metri, per calcolare anche i m² — lascia vuoto per restare in ml'}
            </label>
            <input value={heightMin} onChange={(e) => setHeightMin(e.target.value)} placeholder="es. 2,70" style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 10px', boxSizing: 'border-box' }} />
            {hMinNum > 0 && (
              <>
                <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Altezza massima in metri — uguale alla minima se costante, maggiore per una falda/pendenza</label>
                <input value={heightMax} onChange={(e) => setHeightMax(e.target.value)} placeholder={heightMin} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px', boxSizing: 'border-box' }} />
              </>
            )}
          </>
        )}

        {mode === 'punto' && !selected && <p style={{ fontSize: 11, color: C.maroon, margin: '0 0 12px' }}>Scegli una voce {source === 'computo' ? 'dal computo' : 'dall\'elenco'} per poter aggiungere il punto.</p>}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 8 }}>
          <button onClick={onCancel} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Annulla</button>
          <button onClick={confirm} disabled={!canConfirm} style={{ background: canConfirm ? C.maroon : C.paleGray, color: canConfirm ? C.white : C.gray, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: canConfirm ? 'pointer' : 'default' }}>
            {mode === 'punto' ? 'Aggiungi punto' : 'Salva'}
          </button>
        </div>
      </div>
    </div>
  , document.body);
}

// --- Rilievo planimetrico: sezione a sé nel menu laterale. Prima si scegli il progetto di riferimento
// (e il listino da collegare), poi si carica la planimetria e si disegna sopra — stessa logica di calibrazione
// scala / punti / aree / lunghezze / volumi vista nella scheda progetto, ma qui il progetto non è un prop
// fisso: si scegli da un menu a tendina. Le aree e le lunghezze accettano anche due altezze (minima e
// massima) per le forme a falda (es. tetti a pendenza, pareti sotto falda): il volume o i m² di parete
// usano l'altezza media tra le due come approssimazione (calcolo trapezoidale) — se l'altezza è costante
// basta lasciare la massima uguale alla minima. La planimetria si può ingrandire/rimpicciolire e spostare
// (zoom/pan) per lavorare sui dettagli; le voci di listino si scelgono da un menu a tendina con ricerca
// invece di scrivere il codice a mano, e le aree/lunghezze già disegnate si possono ricollegare a un'altra
// voce o modificare trascinando i punti già tracciati.
function RilievoPage({ projects, setProjects, listini, initialProjectId }) {
  // Il progetto scelto si tiene per INDICE nell'array (non per "id"): così la selezione funziona anche se,
  // nei dati reali, due progetti finissero per condividere lo stesso id (es. importati/creati in blocco) —
  // un confronto per id in quel caso troverebbe sempre il primo progetto con quell'id, indipendentemente
  // da quale sia stato scelto nel menu a tendina.
  const [projectIdx, setProjectIdxRaw] = useState(() => {
    if (initialProjectId != null) {
      const idx = projects.findIndex((p) => p.id === initialProjectId);
      if (idx >= 0) return idx;
    }
    return 0;
  });
  const project = projects[projectIdx] || projects[0];
  const [listinoId, setListinoId] = useState(listini[0]?.id);
  const activeListino = listini.find((l) => l.id === listinoId) || listini[0];
  // Voci già presenti nell'ultima revisione del computo di questo progetto (se esiste): servono alla
  // modale "collega al listino" per offrire, oltre alle voci di listino, anche la scelta di una voce
  // già esistente nel computo a cui sommare la misura appena presa, invece di crearne una nuova.
  const latestComputoItems = (project.revisions && project.revisions.length > 0)
    ? (project.revisions[project.revisions.length - 1].items || []).filter((it) => it.type !== 'subtotal')
    : [];

  // planMode/planDraft sono stato locale di sola interazione (modalità attiva e punti del disegno in corso):
  // non vengono salvati, contano solo mentre questa pagina è aperta — chiave per planimetria, come nel resto del rilievo.
  const [planMode, setPlanModeRaw] = useState({});
  const [planDraft, setPlanDraft] = useState({});
  const setPlanMode = (plId, mode) => { setPlanModeRaw((m) => ({ ...m, [plId]: mode })); setPlanDraft((d) => ({ ...d, [plId]: [] })); setMagnifier(null); };

  // Zoom/pan per planimetria: { scale, tx, ty } in pixel di "viewport" — il contenuto (immagine + svg +
  // marker) resta sempre percentuale (0-100) come prima, solo la conversione clic↔percentuale tiene conto
  // della trasformazione. viewportElsRef tiene i nodi DOM dei contenitori per leggerne le dimensioni reali.
  const [planView, setPlanView] = useState({});
  const viewportElsRef = useRef({});
  const [panDragState, setPanDragState] = useState(null);
  const [draggingHandle, setDraggingHandle] = useState(null);
  // Modifica forma: copia di lavoro dei punti di una forma già salvata, mentre si trascinano i vertici
  // (non si scrive nulla fino a "Salva forma").
  const [editingShape, setEditingShape] = useState(null);
  // Modale "collega al listino", usata sia per nuovi punti/forme sia per ricollegare una forma esistente.
  const [planModal, setPlanModal] = useState(null);
  // Anteprima al volo della voce collegata a una forma già disegnata, cliccandola direttamente sul disegno
  // (senza dover scorrere l'elenco sotto, che resta comunque intatto e completo).
  const [shapePreview, setShapePreview] = useState(null); // { plId, shapeId, xPct, yPct }
  // Lente di ingrandimento: mostra un ritaglio ingrandito della planimetria sotto il cursore mentre si
  // posizionano i punti di una nuova area/lunghezza o si trascina un punto già tracciato — utile per
  // ancorarsi con precisione a un dettaglio del disegno senza dover zoomare tutta la vista (difficile da
  // fare proprio mentre si sta tracciando).
  const [magnifier, setMagnifier] = useState(null); // { plId, loupeLeft, loupeTop, bgWidth, bgHeight, bgPosX, bgPosY }
  const LOUPE_SIZE = 140; // diametro in px della lente
  const LOUPE_ZOOM = 3; // quanto la lente ingrandisce rispetto alla scala "naturale" (non rispetto allo zoom corrente della vista)
  const computeMagnifier = (pl, clientX, clientY) => {
    const el = viewportElsRef.current[pl.id];
    if (!el || !pl.naturalWidth) return null;
    const rect = el.getBoundingClientRect();
    const view = getView(pl.id);
    const baseWidth = rect.width;
    const baseHeight = baseWidth * (pl.naturalHeight / pl.naturalWidth);
    const contentX = (clientX - rect.left - view.tx) / view.scale;
    const contentY = (clientY - rect.top - view.ty) / view.scale;
    const cursorLeft = clientX - rect.left;
    const cursorTop = clientY - rect.top;
    // Di default la lente sta sopra e a destra del cursore (così non ci copre sotto il dito/puntatore);
    // si ribalta sul lato opposto quando non c'entra nel riquadro.
    let loupeLeft = cursorLeft + 16;
    let loupeTop = cursorTop - LOUPE_SIZE - 16;
    if (loupeLeft + LOUPE_SIZE > rect.width) loupeLeft = cursorLeft - LOUPE_SIZE - 16;
    if (loupeTop < 0) loupeTop = cursorTop + 16;
    return {
      plId: pl.id,
      loupeLeft, loupeTop,
      bgWidth: baseWidth * LOUPE_ZOOM,
      bgHeight: baseHeight * LOUPE_ZOOM,
      bgPosX: -(contentX * LOUPE_ZOOM - LOUPE_SIZE / 2),
      bgPosY: -(contentY * LOUPE_ZOOM - LOUPE_SIZE / 2),
    };
  };

  const setProjectIdx = (idx) => {
    setProjectIdxRaw(idx);
    setPlanModal(null);
    setEditingShape(null);
    setDraggingHandle(null);
    setPanDragState(null);
    setShapePreview(null);
    setMagnifier(null);
  };

  if (!projects || projects.length === 0 || !project) {
    return (
      <div>
        <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Rilievo planimetrico</p>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Rilievo planimetrico</h1>
        <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: 18, marginTop: 18 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: 0 }}>Crea prima un progetto per poter caricare una planimetria e fare il rilievo.</p>
        </div>
      </div>
    );
  }

  // Aggiorna per indice (non per id) per lo stesso motivo della selezione sopra: se due progetti condividessero
  // lo stesso id, un confronto per id aggiornerebbe entrambi insieme invece che solo quello scelto qui.
  const onUpdateProject = (updated) => setProjects(projects.map((p, i) => (i === projectIdx ? updated : p)));

  const uploadPlanimetria = (file) => {
    const reader = new FileReader();
    reader.onload = () => {
      const planimetria = { id: Date.now(), name: file.name, image: reader.result, markers: [], shapes: [], scale: null, naturalWidth: null, naturalHeight: null };
      onUpdateProject({ ...project, planimetrie: [...(project.planimetrie || []), planimetria] });
    };
    reader.readAsDataURL(file);
  };

  const removePlanimetria = (id) => {
    if (!confirm('Eliminare questa planimetria e tutti i suoi punti?')) return;
    onUpdateProject({ ...project, planimetrie: (project.planimetrie || []).filter((p) => p.id !== id) });
  };

  // Apre la modale per scegliere la voce di listino e la quantità: l'aggiunta effettiva (voce in computo
  // + marker sulla planimetria) avviene in handleMarkerConfirm con UNA sola onUpdateProject, per non perdere
  // la modifica al computo come accadeva prima con due chiamate separate.
  const addPuntoOnPlanimetria = (planimetriaId, xPct, yPct) => setPlanModal({ kind: 'marker', plId: planimetriaId, xPct, yPct });

  const removePunto = (planimetriaId, markerId) => {
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === planimetriaId ? { ...p, markers: p.markers.filter((m) => m.id !== markerId) } : p)),
    });
  };

  const backfillNaturalSize = (plId, w, h) => {
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === plId && !p.naturalWidth ? { ...p, naturalWidth: w, naturalHeight: h } : p)),
    });
  };

  // --- Zoom / pan ---
  // Scala 1 = esattamente come prima (immagine a piena larghezza del riquadro, altezza naturale, nessun
  // ritaglio): è sia il valore di partenza sia il minimo consentito, così la planimetria non viene mai
  // rimpicciolita per "entrare" in un riquadro fisso — si vede sempre tutta, grande come prima, e lo zoom
  // serve solo per ingrandire oltre quello quando serve vedere un dettaglio (col pan per spostarsi).
  const getView = (plId) => planView[plId] || { scale: 1, tx: 0, ty: 0 };
  const resetView = (plId) => setPlanView((v) => ({ ...v, [plId]: { scale: 1, tx: 0, ty: 0 } }));
  const zoomBy = (plId, pl, factor, center) => {
    const view = getView(plId);
    const el = viewportElsRef.current[plId];
    const rect = el ? el.getBoundingClientRect() : null;
    const cx = center ? center.x : (rect ? rect.width / 2 : 0);
    const cy = center ? center.y : (rect ? rect.height / 2 : 0);
    const newScale = clamp(view.scale * factor, 1, 8);
    if (newScale <= 1) { setPlanView((v) => ({ ...v, [plId]: { scale: 1, tx: 0, ty: 0 } })); return; }
    const contentX = (cx - view.tx) / view.scale;
    const contentY = (cy - view.ty) / view.scale;
    setPlanView((v) => ({ ...v, [plId]: { scale: newScale, tx: cx - contentX * newScale, ty: cy - contentY * newScale } }));
  };
  const startPan = (e, plId) => {
    const view = getView(plId);
    if (view.scale <= 1) return; // niente da spostare finché non si è ingrandito
    e.preventDefault();
    setPanDragState({ plId, startX: e.clientX, startY: e.clientY, startTx: view.tx, startTy: view.ty });
  };
  useEffect(() => {
    if (!panDragState) return;
    const onMove = (e) => {
      const dx = e.clientX - panDragState.startX;
      const dy = e.clientY - panDragState.startY;
      setPlanView((v) => ({ ...v, [panDragState.plId]: { ...(v[panDragState.plId] || { scale: 1, tx: 0, ty: 0 }), tx: panDragState.startTx + dx, ty: panDragState.startTy + dy } }));
    };
    const onUp = () => setPanDragState(null);
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
  }, [panDragState]);

  // Converte una posizione del mouse (coordinate schermo) nella percentuale 0-100 sul contenuto "base"
  // (non trasformato) della planimetria — tiene conto sia dello zoom/pan correnti sia delle dimensioni
  // reali dell'immagine, usata sia per il clic che aggiunge punti/forme sia per il trascinamento dei vertici.
  const toContentPercent = (plId, clientX, clientY) => {
    const el = viewportElsRef.current[plId];
    const pl = (project.planimetrie || []).find((p) => p.id === plId);
    if (!el || !pl || !pl.naturalWidth) return null;
    const rect = el.getBoundingClientRect();
    const view = getView(plId);
    const baseWidth = rect.width;
    const baseHeight = baseWidth * (pl.naturalHeight / pl.naturalWidth);
    const contentX = (clientX - rect.left - view.tx) / view.scale;
    const contentY = (clientY - rect.top - view.ty) / view.scale;
    return { xPct: clamp((contentX / baseWidth) * 100, 0, 100), yPct: clamp((contentY / baseHeight) * 100, 0, 100) };
  };

  const handlePlanClick = (pl, xPct, yPct) => {
    const mode = planMode[pl.id] || 'punto';
    if (mode === 'punto') { addPuntoOnPlanimetria(pl.id, xPct, yPct); return; }

    const nextDraft = [...(planDraft[pl.id] || []), { x: xPct, y: yPct }];

    if (mode === 'calibra') {
      if (nextDraft.length < 2) { setPlanDraft((d) => ({ ...d, [pl.id]: nextDraft })); return; }
      const realDistance = parseEuro(prompt('Distanza reale tra i due punti appena cliccati, in metri (es. "3,20"):', '') || '');
      if (!realDistance) { setPlanDraft((d) => ({ ...d, [pl.id]: [] })); return; }
      onUpdateProject({
        ...project,
        planimetrie: (project.planimetrie || []).map((p) => (p.id === pl.id ? { ...p, scale: { p1: nextDraft[0], p2: nextDraft[1], realDistance } } : p)),
      });
      setPlanMode(pl.id, 'punto');
      return;
    }

    setPlanDraft((d) => ({ ...d, [pl.id]: nextDraft })); // area / lunghezza: accumula punti finché non si preme "Fine"
  };

  const finishPlanShape = (pl) => {
    const mode = planMode[pl.id];
    const draft = planDraft[pl.id] || [];
    const minPoints = mode === 'area' ? 3 : 2;
    if (draft.length < minPoints) { alert(`Servono almeno ${minPoints} punti prima di premere "Fine".`); return; }
    const mpp = metersPerPixel(pl);
    if (!mpp) { alert('Calibra prima la scala di questa planimetria (bottone "Calibra scala").'); return; }
    setPlanModal({ kind: 'shape-new', plId: pl.id, shapeType: mode, points: draft });
  };

  const removePlanShape = (plId, shapeId) => {
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === plId ? { ...p, shapes: (p.shapes || []).filter((s) => s.id !== shapeId) } : p)),
    });
  };

  // Duplica una misura già tracciata: stessa geometria/valore della forma originale (che resta intatta),
  // così la si può ricollegare a un'altra voce di listino con "✏️ Voce" senza ridisegnarla — utile quando
  // la stessa area/lunghezza misurata deve finire nel computo sotto più voci diverse (es. intonaco e pittura
  // sulla stessa parete).
  const duplicatePlanShape = (plId, shape) => {
    const copy = { ...shape, id: Date.now() + Math.random(), points: shape.points.map((p) => ({ ...p })) };
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === plId ? { ...p, shapes: [...(p.shapes || []), copy] } : p)),
    });
  };

  // Forma collegata a una voce di LISTINO: aggiunge una nuova riga al computo (creando la prima
  // revisione in automatico se il progetto non ne ha ancora una).
  const addShapeToComputo = (shape) => {
    const catalogItems = flattenListino(activeListino);
    const voce = catalogItems.find((v) => v.code === shape.code);
    if (!voce) { alert('Voce di listino non trovata nel listino attivo.'); return; }
    const { project: withItem, ok } = addComputoItemToProject(project, voce, shape.value.toFixed(2).replace('.', ','));
    if (!ok) { alert('Non sono riuscito ad aggiungere la misura al computo.'); return; }
    onUpdateProject(withItem);
  };

  // Forma collegata a una voce GIÀ ESISTENTE nel computo: somma il valore della misura alla quantità
  // già presente in quella riga, invece di creare una riga duplicata.
  const updateShapeInComputo = (shape) => {
    const { project: withItem, ok } = incrementComputoItemQty(project, shape.computoItemId, shape.value);
    if (!ok) { alert('Non sono riuscito ad aggiornare la voce scelta nel computo: potrebbe essere stata rimossa o rinominata.'); return; }
    onUpdateProject(withItem);
  };

  const shapeHeightLabel = (s) => {
    if (s.heightMin == null) return '';
    const fmt = (n) => n.toLocaleString('it-IT', { maximumFractionDigits: 2 });
    let label = s.heightMax > s.heightMin ? ` · h ${fmt(s.heightMin)}–${fmt(s.heightMax)} m (falda)` : ` · h ${fmt(s.heightMin)} m`;
    if (s.baseLength != null) label += ` · base ${fmt(s.baseLength)} ml`;
    return label;
  };

  // --- Modale "collega al listino": dispatch verso il caso giusto (nuovo punto, nuova forma, o ricollega
  // una forma esistente) ---
  const closeModal = () => setPlanModal(null);

  const handleMarkerConfirm = ({ source, code, computoItemId, qty }) => {
    if (!planModal) return;
    const { plId, xPct, yPct } = planModal;
    let withItem, ok, linkCode = null, linkComputoItemId = null, linkDesc = null;
    if (source === 'computo') {
      const existingItem = latestComputoItems.find((it) => it.id === computoItemId);
      if (!existingItem) { closeModal(); return; }
      const result = incrementComputoItemQty(project, computoItemId, parseEuro(qty || '0'));
      withItem = result.project; ok = result.ok;
      linkComputoItemId = computoItemId; linkDesc = existingItem.desc;
    } else {
      const catalogItems = flattenListino(activeListino);
      const voce = catalogItems.find((v) => v.code === code);
      if (!voce) { closeModal(); return; }
      const result = addComputoItemToProject(project, voce, qty || '1');
      withItem = result.project; ok = result.ok;
      linkCode = voce.code; linkDesc = voce.desc;
    }
    if (!ok) { alert(source === 'computo' ? 'Non sono riuscito ad aggiornare la voce scelta nel computo.' : 'Non sono riuscito ad aggiungere il punto al computo.'); closeModal(); return; }
    onUpdateProject({
      ...withItem,
      planimetrie: (withItem.planimetrie || []).map((p) => (p.id === plId
        ? { ...p, markers: [...p.markers, { id: Date.now() + Math.random(), x: xPct, y: yPct, code: linkCode, computoItemId: linkComputoItemId, desc: linkDesc }] }
        : p)),
    });
    closeModal();
  };

  const handleShapeNewConfirm = ({ source, code, computoItemId, heightMin, heightMax }) => {
    if (!planModal) return;
    const { plId, shapeType, points } = planModal;
    const pl = (project.planimetrie || []).find((p) => p.id === plId);
    if (!pl) { closeModal(); return; }
    const mpp = metersPerPixel(pl);
    let linkCode = null, linkComputoItemId = null, linkDesc = null;
    if (source === 'computo' && computoItemId) {
      const existingItem = latestComputoItems.find((it) => it.id === computoItemId);
      linkComputoItemId = computoItemId; linkDesc = existingItem?.desc || null;
    } else if (code) {
      const catalogItems = flattenListino(activeListino);
      const voce = catalogItems.find((v) => v.code === code);
      linkCode = voce?.code || null; linkDesc = voce?.desc || null;
    }
    const { value, unit, height, heightMinOut, heightMaxOut, baseLength } = computeShapeValue(pl, mpp, shapeType, points, heightMin, heightMax);
    const shape = {
      id: Date.now() + Math.random(), type: shapeType, points, height, heightMin: heightMinOut, heightMax: heightMaxOut, baseLength,
      code: linkCode, computoItemId: linkComputoItemId, desc: linkDesc || (shapeType === 'area' ? 'Area misurata' : 'Lunghezza misurata'),
      value, unit, color: SECTION_COLORS[(pl.shapes || []).length % SECTION_COLORS.length],
    };
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => (p.id === plId ? { ...p, shapes: [...(p.shapes || []), shape] } : p)),
    });
    setPlanMode(plId, 'punto');
    closeModal();
  };

  const handleShapeEditConfirm = ({ source, code, computoItemId, heightMin, heightMax }) => {
    if (!planModal) return;
    const { plId, shapeId } = planModal;
    let linkCode = null, linkComputoItemId = null, linkDesc = null;
    if (source === 'computo' && computoItemId) {
      const existingItem = latestComputoItems.find((it) => it.id === computoItemId);
      linkComputoItemId = computoItemId; linkDesc = existingItem?.desc || null;
    } else if (code) {
      const catalogItems = flattenListino(activeListino);
      const voce = catalogItems.find((v) => v.code === code);
      linkCode = voce?.code || null; linkDesc = voce?.desc || null;
    }
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => {
        if (p.id !== plId) return p;
        return {
          ...p,
          shapes: p.shapes.map((s) => {
            if (s.id !== shapeId) return s;
            const mpp = metersPerPixel(p);
            const { value, unit, height, heightMinOut, heightMaxOut, baseLength } = computeShapeValue(p, mpp, s.type, s.points, heightMin, heightMax);
            return { ...s, code: linkCode, computoItemId: linkComputoItemId, desc: linkDesc || s.desc, value, unit, height, heightMin: heightMinOut, heightMax: heightMaxOut, baseLength };
          }),
        };
      }),
    });
    closeModal();
  };

  const handlePlanModalConfirm = (result) => {
    if (!planModal) return;
    if (planModal.kind === 'marker') handleMarkerConfirm(result);
    else if (planModal.kind === 'shape-new') handleShapeNewConfirm(result);
    else if (planModal.kind === 'shape-edit') handleShapeEditConfirm(result);
  };

  const openShapeLinkModal = (plId, shape) => setPlanModal({ kind: 'shape-edit', plId, shapeId: shape.id });

  // --- Modifica forma: trascinamento dei vertici già tracciati ---
  const startEditShape = (pl, shape) => { setPlanModal(null); setShapePreview(null); setEditingShape({ plId: pl.id, shapeId: shape.id, points: shape.points.map((p) => ({ ...p })) }); };
  const cancelEditShape = () => setEditingShape(null);
  const saveEditShape = () => {
    if (!editingShape) return;
    const { plId, shapeId, points } = editingShape;
    onUpdateProject({
      ...project,
      planimetrie: (project.planimetrie || []).map((p) => {
        if (p.id !== plId) return p;
        return {
          ...p,
          shapes: p.shapes.map((s) => {
            if (s.id !== shapeId) return s;
            const mpp = metersPerPixel(p);
            const { value, unit, baseLength } = computeShapeValue(p, mpp, s.type, points, s.heightMin, s.heightMax);
            return { ...s, points, value, unit, baseLength };
          }),
        };
      }),
    });
    setEditingShape(null);
  };
  const startDragHandle = (e, plId, idx) => { e.stopPropagation(); e.preventDefault(); setDraggingHandle({ plId, idx }); };
  useEffect(() => {
    if (!draggingHandle) return;
    const onMove = (e) => {
      const pct = toContentPercent(draggingHandle.plId, e.clientX, e.clientY);
      if (pct) {
        setEditingShape((es) => {
          if (!es || es.plId !== draggingHandle.plId) return es;
          const pts = es.points.slice();
          pts[draggingHandle.idx] = { x: pct.xPct, y: pct.yPct };
          return { ...es, points: pts };
        });
      }
      // Lente attiva anche mentre si trascina un punto già tracciato: è lì che serve più precisione,
      // ed è esattamente il momento in cui è più scomodo fermarsi a zoomare tutta la vista.
      const draggedPl = (project.planimetrie || []).find((p) => p.id === draggingHandle.plId);
      if (draggedPl) setMagnifier(computeMagnifier(draggedPl, e.clientX, e.clientY));
    };
    const onUp = () => { setDraggingHandle(null); setMagnifier(null); };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
  }, [draggingHandle]);

  // Props per la modale "collega al listino", risolte in base al tipo di richiesta in corso.
  let modalProps = null;
  if (planModal) {
    if (planModal.kind === 'marker') {
      modalProps = { mode: 'punto', initialCode: null, initialComputoItemId: null, initialQty: '1', initialHeightMin: null, initialHeightMax: null };
    } else if (planModal.kind === 'shape-new') {
      modalProps = { mode: planModal.shapeType, initialCode: null, initialComputoItemId: null, initialQty: '1', initialHeightMin: null, initialHeightMax: null };
    } else if (planModal.kind === 'shape-edit') {
      const pl = (project.planimetrie || []).find((p) => p.id === planModal.plId);
      const shape = pl?.shapes.find((s) => s.id === planModal.shapeId);
      if (shape) modalProps = { mode: shape.type, initialCode: shape.code, initialComputoItemId: shape.computoItemId || null, initialQty: '1', initialHeightMin: shape.heightMin, initialHeightMax: shape.heightMax };
    }
  }

  return (
    <div>
      <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Rilievo planimetrico</p>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 6, flexWrap: 'wrap', justifyContent: 'space-between' }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Rilievo planimetrico</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Dati aggiornati</span>
      </div>
      <p style={{ fontSize: 13, fontWeight: 300, color: '#8B8178', margin: '0 0 18px', maxWidth: 680 }}>
        Scegli il progetto, carica la planimetria, calibra la scala cliccando due punti di cui conosci la distanza reale, poi disegna aree, lunghezze e volumi direttamente sopra il disegno: ogni misura collegata a una voce di listino finisce con un click nel computo, senza scriverla a mano. Usa i controlli in alto a destra sopra l'immagine (🖐 ／ － ＋) per ingrandire e spostare la vista.
      </p>

      <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px', marginBottom: 16, display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 11.5, fontWeight: 600, color: '#6F655F' }}>PROGETTO</span>
          <select value={projectIdx} onChange={(e) => setProjectIdx(Number(e.target.value))} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 10, padding: '9px 14px', fontSize: 13, fontWeight: 400, fontFamily: FONT, color: '#2B2321', cursor: 'pointer', minWidth: 200 }}>
            {projects.map((p, i) => <option key={p.id ?? i} value={i}>{p.name}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 11.5, fontWeight: 600, color: '#6F655F' }}>LISTINO DA COLLEGARE</span>
          <select value={listinoId} onChange={(e) => setListinoId(Number(e.target.value))} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 10, padding: '9px 14px', fontSize: 13, fontWeight: 400, fontFamily: FONT, color: '#2B2321', cursor: 'pointer', minWidth: 170 }}>
            {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
        </div>
        {(!project.revisions || project.revisions.length === 0) && (
          <span style={{ fontSize: 11, color: C.midGray, fontWeight: 600 }}>Questo progetto non ha ancora un computo: verrà creata automaticamente la "Revisione 1" non appena colleghi la prima misura al listino.</span>
        )}
      </div>

      <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px', marginBottom: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
          <div>
            <p style={{ fontWeight: 300, fontSize: 18, margin: 0, color: C.black, fontFamily: FONT }}>Planimetrie di "{project.name}"</p>
            <p style={{ fontSize: 11.5, color: '#8B8178', margin: '2px 0 0' }}>Carica una planimetria: aggiungi punti (es. punti luce), oppure calibra la scala e disegna aree (m²/m³) e lunghezze (ml) collegate al listino — finiscono in automatico nel computo.</p>
          </div>
          <label style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 400, fontFamily: FONT, cursor: 'pointer' }}>
            + Carica planimetria
            <input type="file" accept="image/*" style={{ display: 'none' }} onChange={(e) => { if (e.target.files[0]) uploadPlanimetria(e.target.files[0]); e.target.value = ''; }} />
          </label>
        </div>
        {(!project.planimetrie || project.planimetrie.length === 0) && (
          <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna planimetria caricata ancora per questo progetto. Accetta immagini JPG/PNG.</p>
        )}
        {(project.planimetrie || []).map((pl) => {
          const mode = planMode[pl.id] || 'punto';
          const draft = planDraft[pl.id] || [];
          const isDrawing = mode === 'area' || mode === 'lunghezza';
          const mpp = metersPerPixel(pl);
          const view = getView(pl.id);
          const isEditingThisPl = editingShape && editingShape.plId === pl.id;
          const modeBtn = (key, label, title) => (
            <button
              onClick={() => setPlanMode(pl.id, key)}
              title={title}
              style={{
                fontSize: 11, fontWeight: 600, padding: '6px 10px', borderRadius: 999, cursor: 'pointer',
                border: `1px solid ${mode === key ? C.maroon : C.paleGray}`,
                background: mode === key ? C.maroon : C.white,
                color: mode === key ? C.white : C.black,
              }}
            >
              {label}
            </button>
          );
          return (
            <div key={pl.id} style={{ marginBottom: 18, border: `1px solid ${C.paleGray}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: C.surfaceSubtle }}>
                <span style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{pl.name}</span>
                <button onClick={() => removePlanimetria(pl.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', padding: '8px 12px', borderBottom: `1px solid ${C.paleGray}` }}>
                {modeBtn('punto', '📍 Punto', 'Aggiungi un punto singolo (es. punto luce)')}
                {modeBtn('calibra', mpp ? '📏 Ricalibra' : '📏 Calibra scala', 'Clicca due punti di cui conosci la distanza reale, per convertire i disegni in metri')}
                {modeBtn('area', '▭ Area', 'Disegna un poligono: calcola m², o m³ se indichi un\'altezza (anche a falda, con minima e massima)')}
                {modeBtn('lunghezza', '／ Lunghezza', 'Disegna una linea: calcola i metri lineari, o i m² di una parete se indichi un\'altezza')}
                {isDrawing && (
                  <>
                    <button onClick={() => finishPlanShape(pl)} style={{ ...rowBtnStyle, background: C.maroon, color: C.white, borderColor: C.maroon }}>✓ Fine ({draft.length} punti)</button>
                    <button onClick={() => setPlanMode(pl.id, 'punto')} style={rowBtnStyle}>✕ Annulla</button>
                  </>
                )}
                {mode === 'calibra' && draft.length === 1 && <span style={{ fontSize: 11, color: C.gray }}>Clicca il secondo punto di riferimento…</span>}
                <span style={{ fontSize: 11, color: C.gray, marginLeft: 'auto' }}>
                  {mpp ? `Scala: ${pl.scale.realDistance.toLocaleString('it-IT')} m calibrati` : 'Scala non calibrata'}
                </span>
              </div>

              {isEditingThisPl && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '7px 12px', background: '#FBEFE9', borderBottom: `1px solid ${C.paleGray}` }}>
                  <span style={{ fontSize: 11, color: C.darkGray, fontWeight: 600 }}>✏️ Modifica forma in corso: trascina i punti, poi salva o annulla.</span>
                  <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
                    <button onClick={cancelEditShape} style={rowBtnStyle}>✕ Annulla</button>
                    <button onClick={saveEditShape} style={{ ...rowBtnStyle, background: C.maroon, color: C.white, borderColor: C.maroon }}>✓ Salva forma</button>
                  </div>
                </div>
              )}

              <div
                ref={(el) => { viewportElsRef.current[pl.id] = el; }}
                onMouseDown={(e) => { if (mode === 'sposta') startPan(e, pl.id); }}
                onClick={(e) => {
                  setShapePreview(null);
                  if (mode === 'sposta' || isEditingThisPl) return;
                  const pct = toContentPercent(pl.id, e.clientX, e.clientY);
                  if (!pct) return;
                  handlePlanClick(pl, pct.xPct, pct.yPct);
                }}
                onMouseMove={(e) => {
                  // La lente segue il cursore solo mentre si sta effettivamente tracciando una nuova area/
                  // lunghezza (i click sono precisi solo se si vede bene dove si sta per cliccare), non durante
                  // lo spostamento della vista né quando si è semplicemente "a riposo".
                  if (isDrawing && !panDragState) setMagnifier(computeMagnifier(pl, e.clientX, e.clientY));
                }}
                onMouseLeave={() => setMagnifier((m) => (m && m.plId === pl.id ? null : m))}
                style={{
                  position: 'relative',
                  // A riposo (scala 1) il riquadro non ritaglia nulla: si adatta all'altezza naturale
                  // dell'immagine esattamente come prima, niente bande vuote né rimpicciolimenti forzati.
                  // Solo quando si è effettivamente ingranditi il riquadro si limita a un'altezza ragionevole
                  // con overflow nascosto, per poter spostare la vista (pan) dentro uno spazio contenuto.
                  overflow: view.scale > 1 ? 'hidden' : 'visible',
                  height: view.scale > 1 ? 'min(72vh, 900px)' : 'auto',
                  background: view.scale > 1 ? '#EDEAE5' : 'transparent',
                  cursor: mode === 'sposta' ? (view.scale <= 1 ? 'default' : (panDragState ? 'grabbing' : 'grab')) : (isEditingThisPl ? 'default' : 'crosshair'),
                  userSelect: panDragState || draggingHandle ? 'none' : 'auto',
                }}
              >
                {/* Controlli di zoom/pan: fissi in alto a destra sopra l'immagine (non nella toolbar), così
                    restano sempre a portata di click e non c'entrano nulla con lo scroll della pagina — la
                    rotella del mouse qui sopra scorre la pagina normalmente, lo zoom si fa solo con questi
                    bottoni. */}
                <div
                  onClick={(e) => e.stopPropagation()}
                  onMouseDown={(e) => e.stopPropagation()}
                  style={{
                    position: 'absolute', top: 8, right: 8, zIndex: 5, display: 'flex', alignItems: 'center', gap: 4,
                    background: 'rgba(255,255,255,0.94)', borderRadius: 999, padding: '4px 5px',
                    boxShadow: '0 2px 10px rgba(0,0,0,0.18)', border: `1px solid ${C.paleGray}`,
                  }}
                >
                  <button
                    onClick={() => setPlanMode(pl.id, mode === 'sposta' ? 'punto' : 'sposta')}
                    title="Sposta la vista (pan), utile quando sei ingrandito"
                    style={{
                      fontSize: 13, padding: '5px 8px', borderRadius: 999, cursor: 'pointer', border: 'none',
                      background: mode === 'sposta' ? C.maroon : 'transparent', color: mode === 'sposta' ? C.white : C.black,
                    }}
                  >
                    🖐
                  </button>
                  <button onClick={() => zoomBy(pl.id, pl, 1 / 1.25)} title="Rimpicciolisci" style={{ fontSize: 13, padding: '5px 8px', borderRadius: 999, cursor: 'pointer', border: 'none', background: 'transparent', color: C.black }}>－</button>
                  <span style={{ fontSize: 11, color: C.gray, minWidth: 34, textAlign: 'center' }}>{Math.round(view.scale * 100)}%</span>
                  <button onClick={() => zoomBy(pl.id, pl, 1.25)} title="Ingrandisci" style={{ fontSize: 13, padding: '5px 8px', borderRadius: 999, cursor: 'pointer', border: 'none', background: 'transparent', color: C.black }}>＋</button>
                  {view.scale > 1 && (
                    <button onClick={() => resetView(pl.id)} title="Torna alla visualizzazione intera" style={{ fontSize: 11, fontWeight: 600, padding: '5px 8px', borderRadius: 999, cursor: 'pointer', border: 'none', background: 'transparent', color: C.black }}>⤢</button>
                  )}
                </div>

                <div style={{ position: 'relative', width: '100%', transform: `translate(${view.tx}px, ${view.ty}px) scale(${view.scale})`, transformOrigin: '0 0' }}>
                  <img
                    src={pl.image} alt={pl.name} draggable={false}
                    style={{ width: '100%', display: 'block', pointerEvents: 'none', userSelect: 'none' }}
                    onLoad={(e) => { if (!pl.naturalWidth) backfillNaturalSize(pl.id, e.target.naturalWidth, e.target.naturalHeight); }}
                  />

                  <svg viewBox="0 0 100 100" preserveAspectRatio="none" style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', pointerEvents: 'none' }}>
                    {(pl.shapes || []).map((s) => {
                      const isEditingThis = editingShape && editingShape.plId === pl.id && editingShape.shapeId === s.id;
                      const pts = isEditingThis ? editingShape.points : s.points;
                      const ptsAttr = pts.map((p) => `${p.x},${p.y}`).join(' ');
                      // Si può cliccare una forma già disegnata per vederne l'anteprima (voce collegata, valore)
                      // solo "a riposo" (modalità punto, non mentre si sta disegnando/calibrando/modificando):
                      // altrimenti il click deve continuare a passare al contenitore sotto, come già succedeva.
                      const previewClickable = mode === 'punto' && !isEditingThisPl;
                      const onShapeClick = (e) => {
                        if (!previewClickable) return;
                        e.stopPropagation();
                        const pct = toContentPercent(pl.id, e.clientX, e.clientY);
                        setShapePreview({ plId: pl.id, shapeId: s.id, xPct: pct ? pct.xPct : pts[0].x, yPct: pct ? pct.yPct : pts[0].y });
                      };
                      return (
                        <React.Fragment key={s.id}>
                          {s.type === 'area' ? (
                            <polygon
                              points={ptsAttr} fill={s.color} fillOpacity={0.28} stroke={s.color} strokeWidth={isEditingThis ? 0.6 : 0.4}
                              vectorEffect="non-scaling-stroke"
                              style={{ pointerEvents: previewClickable ? 'auto' : 'none', cursor: previewClickable ? 'pointer' : 'inherit' }}
                              onClick={onShapeClick}
                            />
                          ) : (
                            <>
                              {/* Bersaglio di click invisibile e più largo della linea vera, altrimenti una
                                  lunghezza sottile sarebbe troppo difficile da colpire con un click. */}
                              <polyline points={ptsAttr} fill="none" stroke="transparent" strokeWidth={3} vectorEffect="non-scaling-stroke" style={{ pointerEvents: previewClickable ? 'auto' : 'none', cursor: previewClickable ? 'pointer' : 'inherit' }} onClick={onShapeClick} />
                              <polyline points={ptsAttr} fill="none" stroke={s.color} strokeWidth={isEditingThis ? 0.9 : 0.6} vectorEffect="non-scaling-stroke" style={{ pointerEvents: 'none' }} />
                            </>
                          )}
                          {isEditingThis && pts.map((p, i) => (
                            <React.Fragment key={i}>
                              {/* Bersaglio di trascinamento più largo (invisibile) del pallino visibile sotto:
                                  il pallino vero si è rimpicciolito per ancorarsi con più precisione (non
                                  nasconde più il punto esatto sotto il cursore), ma resta facile da afferrare. */}
                              <circle
                                cx={p.x} cy={p.y} r={3} fill="transparent"
                                vectorEffect="non-scaling-stroke" style={{ cursor: 'move', pointerEvents: 'auto' }}
                                onMouseDown={(e) => startDragHandle(e, pl.id, i)}
                                onClick={(e) => e.stopPropagation()}
                              />
                              <circle cx={p.x} cy={p.y} r={1} fill={C.white} stroke={s.color} strokeWidth={0.4} vectorEffect="non-scaling-stroke" style={{ pointerEvents: 'none' }} />
                            </React.Fragment>
                          ))}
                        </React.Fragment>
                      );
                    })}
                    {draft.length > 0 && (
                      <polyline points={draft.map((p) => `${p.x},${p.y}`).join(' ')} fill="none" style={{ stroke: C.maroon }} strokeWidth={0.5} strokeDasharray="1.5,1" vectorEffect="non-scaling-stroke" />
                    )}
                    {draft.map((p, i) => (
                      <circle key={i} cx={p.x} cy={p.y} r={0.5} style={{ fill: C.maroon }} stroke="white" strokeWidth={0.25} vectorEffect="non-scaling-stroke" />
                    ))}
                  </svg>

                  {shapePreview && shapePreview.plId === pl.id && (() => {
                    const previewShape = (pl.shapes || []).find((s) => s.id === shapePreview.shapeId);
                    if (!previewShape) return null;
                    return (
                      <div
                        onClick={(e) => e.stopPropagation()}
                        style={{
                          position: 'absolute', left: `${shapePreview.xPct}%`, top: `${shapePreview.yPct}%`,
                          transform: `scale(${1 / view.scale}) translate(-50%, calc(-100% - 14px))`, transformOrigin: 'bottom center',
                          background: C.white, borderRadius: 10, padding: '8px 10px', minWidth: 180, maxWidth: 240,
                          boxShadow: '0 4px 16px rgba(0,0,0,0.25)', border: `1px solid ${C.paleGray}`, zIndex: 7, pointerEvents: 'auto',
                        }}
                      >
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 6 }}>
                          <span style={{ width: 9, height: 9, borderRadius: 3, background: previewShape.color, flexShrink: 0, marginTop: 3 }} />
                          <button onClick={() => setShapePreview(null)} style={{ border: 'none', background: 'none', color: C.gray, cursor: 'pointer', fontSize: 12, padding: 0, marginLeft: 'auto' }}>✕</button>
                        </div>
                        <p style={{ fontSize: 12, fontWeight: 700, color: C.black, margin: '2px 0 1px' }}>
                          {previewShape.desc}{previewShape.code ? ` (${previewShape.code})` : ''}
                        </p>
                        <p style={{ fontSize: 12, color: C.maroon, fontWeight: 700, margin: 0 }}>
                          {previewShape.value.toLocaleString('it-IT', { maximumFractionDigits: 2 })} {previewShape.unit}{shapeHeightLabel(previewShape)}
                        </p>
                        {!previewShape.code && !previewShape.computoItemId && (
                          <p style={{ fontSize: 10.5, color: C.gray, margin: '2px 0 0' }}>Non ancora collegata a una voce.</p>
                        )}
                      </div>
                    );
                  })()}

                  {pl.markers.map((m) => (
                    <div
                      key={m.id}
                      onClick={(e) => { e.stopPropagation(); if (confirm(`Rimuovere il punto "${m.desc}"?`)) removePunto(pl.id, m.id); }}
                      title={`${m.code ? m.code + ' — ' : ''}${m.desc} (clicca per rimuovere)`}
                      style={{
                        position: 'absolute', left: `${m.x}%`, top: `${m.y}%`, transform: 'translate(-50%, -50%)',
                        width: 22, height: 22, borderRadius: 999, background: C.maroon, color: C.white,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700,
                        border: '2px solid white', boxShadow: '0 1px 3px rgba(0,0,0,0.3)', cursor: 'pointer', pointerEvents: 'auto',
                      }}
                    >
                      {pl.markers.indexOf(m) + 1}
                    </div>
                  ))}
                </div>

                {/* Lente di ingrandimento: FUORI dal div trasformato (scale/pan), così resta sempre della
                    stessa dimensione a schermo e posizionata in pixel reali rispetto al cursore, indipendente
                    dallo zoom corrente della vista. */}
                {magnifier && magnifier.plId === pl.id && (
                  <div
                    style={{
                      position: 'absolute', left: magnifier.loupeLeft, top: magnifier.loupeTop,
                      width: LOUPE_SIZE, height: LOUPE_SIZE, borderRadius: '50%', overflow: 'hidden',
                      border: `3px solid ${C.maroon}`, boxShadow: '0 4px 14px rgba(0,0,0,0.35)',
                      backgroundImage: `url(${pl.image})`, backgroundRepeat: 'no-repeat',
                      backgroundSize: `${magnifier.bgWidth}px ${magnifier.bgHeight}px`,
                      backgroundPosition: `${magnifier.bgPosX}px ${magnifier.bgPosY}px`,
                      pointerEvents: 'none', zIndex: 8,
                    }}
                  >
                    <div style={{ position: 'absolute', left: '50%', top: '50%', width: 1, height: 16, background: C.maroon, transform: 'translate(-50%, -50%)' }} />
                    <div style={{ position: 'absolute', left: '50%', top: '50%', width: 16, height: 1, background: C.maroon, transform: 'translate(-50%, -50%)' }} />
                  </div>
                )}
              </div>

              {pl.markers.length > 0 && (
                <div style={{ padding: '8px 12px', fontSize: 11, color: C.darkGray, borderTop: `1px solid ${C.paleGray}` }}>
                  {pl.markers.length} punti aggiunti al computo · clicca un punto per rimuoverlo
                </div>
              )}

              {(pl.shapes || []).length > 0 && (
                <div style={{ padding: '10px 12px', borderTop: `1px solid ${C.paleGray}` }}>
                  {pl.shapes.map((s) => (
                    <div key={s.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap', padding: '6px 0' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
                        <span style={{ width: 10, height: 10, borderRadius: 3, background: s.color, flexShrink: 0 }} />
                        <span style={{ fontSize: 12, color: C.black, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {s.desc}{s.code ? ` (${s.code})` : ''} — <strong>{s.value.toLocaleString('it-IT', { maximumFractionDigits: 2 })} {s.unit}</strong>{shapeHeightLabel(s)}
                        </span>
                      </div>
                      <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                        <button onClick={() => openShapeLinkModal(pl.id, s)} style={rowBtnStyle}>✏️ Voce</button>
                        <button onClick={() => startEditShape(pl, s)} style={rowBtnStyle}>✏️ Forma</button>
                        <button onClick={() => duplicatePlanShape(pl.id, s)} title="Duplica questa misura per collegarla anche a un'altra voce di listino" style={rowBtnStyle}>⧉ Duplica</button>
                        {s.computoItemId ? (
                          <button onClick={() => updateShapeInComputo(s)} style={rowBtnStyle} title="Aggiunge il valore di questa misura alla quantità già presente nel computo">✎ Aggiorna computo</button>
                        ) : s.code && (
                          <button onClick={() => addShapeToComputo(s)} style={rowBtnStyle} title="Crea una nuova riga nel computo con il valore di questa misura">+ Al computo</button>
                        )}
                        <button onClick={() => removePlanShape(pl.id, s.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {planModal && modalProps && (
        <PlanLinkModal
          {...modalProps}
          listino={activeListino}
          computoItems={latestComputoItems}
          onCancel={closeModal}
          onConfirm={handlePlanModalConfirm}
        />
      )}
    </div>
  );
}

function DiffTable({ diff }) {
  const esitoColor = { Invariata: '#8B8178', Aggiunta: '#2F5A2B', Modificata: C.maroon, Rimossa: '#8C3220' };
  const esitoBg = { Invariata: '#EFE8DD', Aggiunta: 'color-mix(in srgb, #4F7A4B 16%, #fff)', Modificata: `color-mix(in srgb, ${C.maroon} 16%, #fff)`, Rimossa: 'color-mix(in srgb, #B3452B 16%, #fff)' };
  const rowBg = { Aggiunta: 'color-mix(in srgb, #4F7A4B 10%, #fff)', Rimossa: 'color-mix(in srgb, #B3452B 10%, #fff)', Modificata: `color-mix(in srgb, ${C.maroon} 8%, #fff)` };
  if (!diff || diff.length === 0) {
    return <p style={{ fontSize: 13, color: '#8B8178' }}>Nessuna voce da confrontare tra queste due versioni.</p>;
  }
  return (
    <div className="table-scroll">
    <table style={{ width: '100%', minWidth: 760, borderCollapse: 'collapse' }}>
      <thead>
        <tr>
          {['Voce', 'Sezione', 'Q.tà prima', 'Q.tà dopo', 'Prezzo prima', 'Prezzo dopo', 'Variazione'].map((h) => (
            <th key={h} style={{ fontSize: 11, letterSpacing: '.1em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: h.startsWith('Q.') || h.startsWith('Prezzo') || h === 'Variazione' ? 'right' : 'left', padding: '8px 10px', borderBottom: `1px solid ${C.borderStrong}` }}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {diff.map((r) => (
          <tr key={r.code} style={{ background: rowBg[r.esito] || 'transparent' }}>
            <td style={{ fontSize: 13.5, fontWeight: 300, padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textDecoration: r.esito === 'Rimossa' ? 'line-through' : 'none', color: r.esito === 'Rimossa' ? '#8B8178' : C.black }}>
              <div style={{ fontWeight: 400 }}>{r.code}</div>
              <div style={{ fontSize: 11.5, color: '#8B8178' }}>{r.desc}</div>
            </td>
            <td style={{ padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}` }}>
              <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 9px', borderRadius: 999, background: '#EFE8DD', color: '#6F655F' }}>{r.sezione}</span>
            </td>
            <td style={{ fontSize: 12.5, fontVariantNumeric: 'tabular-nums', color: '#6F655F', padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textAlign: 'right' }}>{r.qtyBefore}</td>
            <td style={{ fontSize: 12.5, fontVariantNumeric: 'tabular-nums', color: '#6F655F', padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textAlign: 'right' }}>{r.qtyAfter}</td>
            <td style={{ fontSize: 12.5, fontVariantNumeric: 'tabular-nums', color: '#6F655F', padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textAlign: 'right' }}>{r.priceBefore}</td>
            <td style={{ fontSize: 12.5, fontVariantNumeric: 'tabular-nums', color: '#6F655F', padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textAlign: 'right' }}>{r.priceAfter}</td>
            <td style={{ padding: '10px 10px', borderBottom: `1px solid ${C.paleGray}`, textAlign: 'right' }}>
              <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: esitoBg[r.esito], color: esitoColor[r.esito] }}>{r.esito}</span>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
    </div>
  );
}

function ComputiPage({ projects, setProjects, onOpenProject, onOpenRevision, requestPdf }) {
  const [filterId, setFilterId] = useState('all'); // 'all' oppure l'id di un progetto

  const updateProject = (updated) => setProjects(projects.map((p) => (p.id === updated.id ? updated : p)));
  const changeStatus = (proj, revId, status) => {
    updateProject({ ...proj, revisions: proj.revisions.map((r) => (r.id === revId ? { ...r, status } : r)) });
  };
  const deleteRevision = (proj, revId) => {
    if (!confirm('Eliminare definitivamente questa versione del computo?')) return;
    updateProject({ ...proj, revisions: proj.revisions.filter((r) => r.id !== revId) });
  };
  const renameRevision = (proj, revId) => {
    const rev = proj.revisions.find((r) => r.id === revId);
    const newName = prompt('Nuovo nome del computo (data di modifica e numero revisione restano invariati):', rev.customName || rev.label);
    if (!newName) return;
    updateProject({ ...proj, revisions: proj.revisions.map((r) => (r.id === revId ? { ...r, customName: newName } : r)) });
  };

  const ghostPillBtn = { border: `1px solid ${C.paleGray}`, background: C.white, color: '#2B2321', fontSize: 11.5, fontWeight: 400, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT, whiteSpace: 'nowrap' };

  const projectsWithRevisions = projects.filter((p) => p.revisions && p.revisions.length > 0);
  const activeProject = filterId !== 'all' ? projects.find((p) => p.id === filterId) : null;

  // Una riga-revisione: usata sia nella vista "Tutti i computi" (solo il computo attivo, una riga)
  // sia nella vista di un singolo progetto (tutta la cronologia delle revisioni, una riga per versione).
  const revisionRow = (proj, r) => (
    <div key={r.id} style={{ padding: '14px 0', borderTop: `1px solid ${C.paleGray}` }}>
      <div style={{ fontWeight: 500, fontSize: 14, color: C.black }}>
        {r.customName || r.label} {r.customName && <span style={{ fontSize: 10, color: '#8B8178', fontWeight: 400 }}>({r.label})</span>}
      </div>
      <div style={{ fontVariantNumeric: 'tabular-nums', color: '#8B8178', fontSize: 12, marginTop: 2 }}>
        Creata il {r.dateCreated} · Modificata il {r.dateModified} · {r.total} impresa / {r.totalCliente || r.total} cliente
      </div>
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 12, flexWrap: 'wrap' }}>
        <select
          value={r.status}
          onChange={(e) => changeStatus(proj, r.id, e.target.value)}
          style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '8px 14px', fontSize: 12.5, fontWeight: 400, color: '#2B2321', fontFamily: FONT }}
        >
          {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <button onClick={() => onOpenRevision(proj.id, r.id)} style={ghostPillBtn}>Apri e modifica</button>
        <button onClick={() => renameRevision(proj, r.id)} style={ghostPillBtn}>✎ Rinomina</button>
        <button onClick={() => exportComputoExcel(proj, r, false)} style={ghostPillBtn}>Excel</button>
        <button onClick={() => requestPdf(proj, r, false)} style={ghostPillBtn}>PDF</button>
        <button onClick={() => deleteRevision(proj, r.id)} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 14, marginLeft: 4 }}>🗑️</button>
      </div>
    </div>
  );

  return (
    <div>
      <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Computi</p>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 6, flexWrap: 'wrap', justifyContent: 'space-between' }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Computi</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Dati aggiornati</span>
      </div>
      <p style={{ fontSize: 13.5, fontWeight: 300, color: '#8B8178', margin: '0 0 18px', maxWidth: 640 }}>
        "Tutti i computi" mostra il computo attivo di ogni progetto, uno sotto l'altro. Seleziona un progetto per vedere tutta la sua cronologia di revisioni.
      </p>

      <div style={{ display: 'flex', gap: 8, marginBottom: 18, flexWrap: 'wrap' }}>
        <span
          onClick={() => setFilterId('all')}
          style={{
            fontSize: 12.5, fontWeight: filterId === 'all' ? 500 : 400, padding: '8px 14px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT,
            background: filterId === 'all' ? '#1E1A18' : C.white,
            color: filterId === 'all' ? '#fff' : '#4A403C',
            border: `1px solid ${filterId === 'all' ? '#1E1A18' : C.paleGray}`,
          }}
        >
          Tutti i computi
        </span>
        {projects.map((p) => (
          <span
            key={p.id}
            onClick={() => setFilterId(p.id)}
            style={{
              fontSize: 12.5, fontWeight: p.id === filterId ? 500 : 400, padding: '8px 14px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT,
              background: p.id === filterId ? '#1E1A18' : C.white,
              color: p.id === filterId ? '#fff' : '#4A403C',
              border: `1px solid ${p.id === filterId ? '#1E1A18' : C.paleGray}`,
            }}
          >
            {p.name}
          </span>
        ))}
      </div>

      {filterId === 'all' ? (
        projectsWithRevisions.length === 0 ? (
          <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 12, textAlign: 'center', padding: 30 }}>
            <p style={{ fontSize: 13, color: C.gray, margin: 0 }}>Nessun progetto ha ancora un computo creato.</p>
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            {projectsWithRevisions.map((p) => {
              const latest = p.revisions[p.revisions.length - 1];
              return (
                <div key={p.id} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '18px 20px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, flexWrap: 'wrap', marginBottom: 6 }}>
                    <div onClick={() => setFilterId(p.id)} style={{ cursor: 'pointer', fontSize: 16, fontWeight: 500, color: C.black }}>
                      {p.name} · {p.client} {p.revisions.length > 1 && <span style={{ fontSize: 11, fontWeight: 400, color: '#8B8178' }}>({p.revisions.length} revisioni — clicca per vederle tutte)</span>}
                    </div>
                    <a onClick={() => onOpenProject(p.id)} style={{ cursor: 'pointer', fontSize: 12.5, color: '#4A403C', border: `1px solid ${C.paleGray}`, padding: '7px 14px', borderRadius: 999 }}>Apri progetto ›</a>
                  </div>
                  {revisionRow(p, latest)}
                </div>
              );
            })}
          </div>
        )
      ) : !activeProject ? (
        <p style={{ fontSize: 13, color: C.gray }}>Progetto non trovato.</p>
      ) : activeProject.revisions.length === 0 ? (
        <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 12, textAlign: 'center', padding: 30 }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 14px' }}>"{activeProject.name}" non ha ancora nessun computo creato.</p>
          <button onClick={() => onOpenProject(activeProject.id)} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 16px', fontSize: 13, fontWeight: 500, cursor: 'pointer', fontFamily: FONT }}>Apri il progetto per crearlo</button>
        </div>
      ) : (
        <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '18px 20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 16, flexWrap: 'wrap', marginBottom: 6 }}>
            <div style={{ fontSize: 16, fontWeight: 500, color: C.black }}>{activeProject.name} · {activeProject.client}</div>
            <a onClick={() => onOpenProject(activeProject.id)} style={{ cursor: 'pointer', fontSize: 12.5, color: '#4A403C', border: `1px solid ${C.paleGray}`, padding: '7px 14px', borderRadius: 999 }}>Apri progetto ›</a>
          </div>
          {activeProject.revisions.map((r) => revisionRow(activeProject, r))}
        </div>
      )}
    </div>
  );
}

function ConfrontoPage({ projects }) {
  // Progetto e revisioni si scelgono per POSIZIONE nell'elenco (non per "id"), come nel Rilievo: nei dati
  // reali alcuni progetti possono avere lo stesso id (es. importati o creati in blocco) o id non numerici,
  // e una ricerca per id finiva sempre sul primo progetto, così la tendina sembrava non cambiare.
  const firstComparable = Math.max(0, projects.findIndex((p) => (p.revisions || []).length >= 2));
  const [projectIdx, setProjectIdx] = useState(firstComparable);
  const project = projects[projectIdx] || projects[0];
  const revisions = project?.revisions || [];
  const [fromIdx, setFromIdx] = useState(0);
  const [toIdx, setToIdx] = useState(null); // null = ultima revisione
  const fi = Math.min(fromIdx, Math.max(revisions.length - 1, 0));
  const ti = toIdx == null || toIdx >= revisions.length ? revisions.length - 1 : toIdx;

  const handleProjectChange = (idx) => {
    setProjectIdx(idx);
    setFromIdx(0);
    setToIdx(null);
  };

  if (!project) {
    return (
      <div>
        <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Confronto revisioni</p>
        <h1 style={{ margin: '0 0 18px', fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Confronto revisioni</h1>
        <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, textAlign: 'center', padding: 30 }}>
          <p style={{ fontSize: 14, color: '#8B8178', margin: 0 }}>Non ci sono ancora progetti: crea un progetto e salva almeno due revisioni del computo per confrontarle.</p>
        </div>
      </div>
    );
  }

  const selStyle = { border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 10, padding: '9px 14px', fontSize: 13, fontWeight: 400, color: '#2B2321', fontFamily: FONT, cursor: 'pointer' };

  return (
    <div>
      <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Confronto revisioni</p>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 18, flexWrap: 'wrap', justifyContent: 'space-between' }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Confronto revisioni</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Dati aggiornati</span>
      </div>

      <section style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px', marginBottom: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
          <span style={{ fontSize: 11.5, fontWeight: 600, color: '#6F655F' }}>PROGETTO</span>
          <select value={projectIdx} onChange={(e) => handleProjectChange(Number(e.target.value))} style={{ ...selStyle, fontSize: 14, minWidth: 260 }}>
            {projects.map((p, i) => <option key={i} value={i}>{p.name}{(p.revisions || []).length < 2 ? ' (meno di 2 revisioni)' : ''}</option>)}
          </select>
        </div>
      </section>

      {revisions.length < 2 ? (
        <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, textAlign: 'center', padding: 30 }}>
          <p style={{ fontSize: 13, color: '#8B8178', margin: 0 }}>"{project.name}" ha meno di due revisioni salvate: non c'è ancora nulla da confrontare.</p>
        </div>
      ) : (
        <section style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px', marginBottom: 16, overflowX: 'auto' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 18, flexWrap: 'wrap' }}>
            <select value={fi} onChange={(e) => setFromIdx(Number(e.target.value))} style={{ ...selStyle, fontSize: 14 }}>
              {revisions.map((r, i) => <option key={i} value={i}>{r.customName || r.label} · {r.dateModified}</option>)}
            </select>
            <span style={{ fontSize: 15, color: '#8B8178' }}>→</span>
            <select value={ti} onChange={(e) => setToIdx(Number(e.target.value))} style={{ ...selStyle, fontSize: 14 }}>
              {revisions.map((r, i) => <option key={i} value={i}>{r.customName || r.label} · {r.dateModified}</option>)}
            </select>
          </div>

          <DiffTable diff={computeItemsDiff(revisions[fi]?.items, revisions[ti]?.items)} />
        </section>
      )}

      <section style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px' }}>
        <h3 style={{ margin: '0 0 2px', fontWeight: 300, fontSize: 18, color: C.black }}>Stampa ed esportazione</h3>
        <p style={{ fontSize: 11.5, fontWeight: 400, color: '#8B8178', margin: 0, maxWidth: 560 }}>
          Intestazione, subtotali e sezione Extra sono sempre inclusi. Mantieni nella stampa le evidenziazioni di aggiunte, modifiche e rimozioni.
        </p>
        <div style={{ display: 'flex', gap: 10, marginTop: 12 }}>
          <button onClick={() => alert('Esportazione Excel/CSV: funzione da collegare al backend reale.')} style={{ border: `1px solid ${C.paleGray}`, background: C.white, color: '#2B2321', padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 400, cursor: 'pointer', fontFamily: FONT }}>Excel / CSV</button>
          <button onClick={() => alert('Stampa/PDF: genera il documento con le evidenziazioni delle modifiche.')} style={{ border: 'none', background: C.maroon, color: C.white, padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 400, cursor: 'pointer', fontFamily: FONT }}>Stampa / PDF</button>
        </div>
      </section>
    </div>
  );
}

// Impostazioni studio: dati anagrafici mostrati in automatico su intestazione/piè di pagina del PDF del
// computo, più l'eventuale intestazione/piè di pagina "personalizzata" (immagini caricate dallo studio,
// che sostituiscono del tutto l'intestazione/piè di pagina generata dai campi). Le immagini sono salvate
// come data URL inline, con lo stesso meccanismo già usato per le planimetrie dei progetti.
const DEFAULT_STUDIO_SETTINGS = {
  nome: '', indirizzo: '', piva: '', cf: '', telefono: '', email: '', sito: '',
  logo: null, // data URL
  usaIntestazionePersonalizzata: false,
  intestazioneImg: null, // data URL, mostrata al posto dei dati anagrafici in cima a ogni pagina stampata
  intestazioneScala: 100, // % della larghezza disponibile (dentro il margine di pagina) occupata dall'immagine
  usaPiePersonalizzato: false,
  pieImg: null, // data URL, mostrata al posto del testo di piè di pagina predefinito
  pieScala: 100, // % della larghezza disponibile occupata dall'immagine del piè di pagina
  testoPiePagina: '',
  usaFontPersonalizzato: false,
  fontPersonalizzato: null, // data URL del file font (woff2/woff/ttf/otf), usato nel PDF al posto del font di sistema
  fontPersonalizzatoNome: '', // nome del file caricato, solo per mostrarlo in Impostazioni
  coloreAccento: DEFAULT_ACCENT, // colore di accento di tutta l'app e dei PDF (pulsanti, evidenze, totali)
  pdfFontPreset: 'inter', // font del PDF quando non si usa un font personalizzato: 'inter' | 'serif' | 'helvetica'
};

const PDF_FONT_PRESETS = {
  inter: { label: 'Inter — moderno, pulito (predefinito)', css: FONT },
  serif: { label: 'Serif classico — Georgia, stile carta intestata', css: "Georgia, 'Times New Roman', Times, serif" },
  helvetica: { label: 'Sans neutro — Helvetica / Arial', css: "'Helvetica Neue', Helvetica, Arial, sans-serif" },
};
const ACCENT_PRESETS = [
  { hex: '#862D2D', name: 'Bordeaux Desearq' },
  { hex: '#2B2321', name: 'Grafite' },
  { hex: '#1F3A5F', name: 'Blu notte' },
  { hex: '#1E5A5A', name: 'Ottanio' },
  { hex: '#4F6B57', name: 'Verde salvia' },
  { hex: '#A4502F', name: 'Terracotta' },
  { hex: '#6B4E8A', name: 'Prugna' },
];
const isHexColor = (v) => /^#[0-9a-fA-F]{6}$/.test(String(v || ''));
// Luminanza relativa (WCAG): sopra ~0.45 il testo bianco sul colore di accento diventa poco leggibile.
const hexLuminance = (hex) => {
  const ch = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
};
function applyAccentColor(hex) {
  const accent = isHexColor(hex) ? hex : DEFAULT_ACCENT;
  const root = document.documentElement.style;
  root.setProperty('--accent', accent);
  root.setProperty('--accent-hover', `color-mix(in srgb, ${accent} 82%, #000)`);
  root.setProperty('--accent-soft', `color-mix(in srgb, ${accent} 16%, #fff)`);
}

const INITIAL_FORNITORI = [
  {
    name: 'Forniture sanitarie', categorie: [
      { name: 'WC', prodotti: [
        { id: 1, marca: '', modello: 'WC sospeso', photo: null, varianti: [
          { id: 1, nome: 'Serie Rovere', listinoPrice: '180,00', fornitori: [
            { id: 1, name: 'Idroterm Forniture S.r.l.', prezzoListino: '180,00', prezzoScontato: '150,00', prezzoCliente: '165,00' },
            { id: 2, name: 'Ceramiche Rossi', prezzoListino: '190,00', prezzoScontato: '160,00', prezzoCliente: '172,00' },
          ]},
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

// Un prodotto salvato nel vecchio formato (name/listinoPrice/fornitori diretti sul prodotto,
// senza varianti) viene convertito al volo in una singola variante senza nome: i dati già
// salvati online restano intatti quando arriva questo aggiornamento.
function migrateProdotto(p) {
  if (p.varianti) return p;
  return {
    id: p.id,
    marca: '',
    modello: p.name || '',
    photo: p.photo ?? null,
    varianti: [{ id: p.id, nome: '', listinoPrice: p.listinoPrice || '0,00', fornitori: p.fornitori || [] }],
  };
}
function migrateFornitoriCatalog(catalog) {
  return (catalog || []).map((m) => ({
    ...m,
    categorie: (m.categorie || []).map((c) => ({ ...c, prodotti: (c.prodotti || []).map(migrateProdotto) })),
  }));
}

// Modale per creare una revisione del computo a partire da un PDF: ne estrae solo le macrocategorie
// e i loro totali (mai le singole voci), lascia rivedere/correggere tutto a mano prima di confermare,
// e permette di inserire manualmente il totale complessivo dell'impresa (non deducibile dal PDF cliente).
function ImportPdfComputoModal({ onClose, onConfirm }) {
  const [stage, setStage] = useState('upload'); // upload | loading | review | error
  const [rows, setRows] = useState([]);
  const [totaleRilevato, setTotaleRilevato] = useState(null);
  const [totaleImpresa, setTotaleImpresa] = useState('');
  const [errorMsg, setErrorMsg] = useState('');

  const handleFile = async (file) => {
    if (!file) return;
    setStage('loading');
    try {
      const lines = await extractTextLinesFromPdf(file);
      const { rows: found, totaleComplessivo } = parseComputoPdfLines(lines);
      if (found.length === 0) {
        setErrorMsg('Non ho trovato righe con un importo riconoscibile in questo PDF. Puoi comunque inserire le macrocategorie a mano qui sotto.');
      }
      setRows(found);
      setTotaleRilevato(totaleComplessivo);
      setStage('review');
    } catch (err) {
      setErrorMsg('Non sono riuscito a leggere questo PDF (' + String(err?.message || err) + '). Verifica che sia un PDF testuale valido, oppure inserisci le macrocategorie a mano.');
      setRows([]);
      setTotaleRilevato(null);
      setStage('review');
    }
  };

  const updateRow = (id, field, value) => setRows((rs) => rs.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  const removeRow = (id) => setRows((rs) => rs.filter((r) => r.id !== id));
  const addRow = () => setRows((rs) => [...rs, { id: Date.now() + Math.random(), name: '', totale: '0,00' }]);

  const sommaCategorie = rows.reduce((sum, r) => sum + parseEuro(r.totale), 0);
  const scostamento = totaleRilevato !== null ? Math.abs(sommaCategorie - totaleRilevato) : 0;
  const mismatchTotale = totaleRilevato !== null && scostamento > 0.02;

  const canConfirm = rows.length > 0 && rows.every((r) => r.name.trim()) && sommaCategorie > 0;

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 560, maxWidth: '100%', maxHeight: '90vh', overflowY: 'auto' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 6px', color: C.black }}>Importa computo da PDF</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>
          Dal PDF vengono lette solo le macrocategorie e i loro totali (mai le singole voci): creerai una nuova revisione da questi totali, per poi dettagliarla con le voci una alla volta.
        </p>

        {stage === 'upload' && (
          <label style={{ display: 'block', textAlign: 'center', background: C.bg, border: `1px dashed ${C.paleGray}`, borderRadius: 10, padding: '30px 14px', fontSize: 13, fontWeight: 600, color: C.black, cursor: 'pointer' }}>
            📄 Carica il PDF del computo
            <input type="file" accept="application/pdf" style={{ display: 'none' }} onChange={(e) => handleFile(e.target.files[0])} />
          </label>
        )}

        {stage === 'loading' && (
          <p style={{ fontSize: 13, color: C.gray, textAlign: 'center', padding: 30 }}>Lettura del PDF in corso…</p>
        )}

        {stage === 'review' && (
          <>
            {errorMsg && <p style={{ fontSize: 12, color: C.maroon, background: 'rgba(128,20,48,0.08)', borderRadius: 8, padding: '8px 10px', margin: '0 0 12px' }}>{errorMsg}</p>}

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
              <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Macrocategorie e relativi totali (cliente)</label>
              <button onClick={addRow} style={rowBtnStyle}>+ Categoria</button>
            </div>
            {rows.length === 0 ? (
              <p style={{ fontSize: 12, color: C.gray, margin: '0 0 10px' }}>Nessuna macrocategoria ancora. Aggiungine una con "+ Categoria".</p>
            ) : rows.map((r) => (
              <div key={r.id} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                <input value={r.name} onChange={(e) => updateRow(r.id, 'name', e.target.value)} placeholder="Nome macrocategoria"
                  style={{ flex: 1, fontSize: 12, padding: '7px 9px', borderRadius: 8, border: `1px solid ${C.paleGray}` }} />
                <input value={r.totale} onChange={(e) => updateRow(r.id, 'totale', e.target.value)} placeholder="0,00"
                  style={{ width: 100, fontSize: 12, padding: '7px 9px', borderRadius: 8, border: `1px solid ${C.paleGray}`, textAlign: 'right' }} />
                <span style={{ fontSize: 11, color: C.gray }}>€</span>
                <button onClick={() => removeRow(r.id)} style={{ ...rowBtnStyle, padding: '5px 8px' }}>🗑</button>
              </div>
            ))}

            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, margin: '12px 0', padding: '10px 12px', background: C.bg, borderRadius: 8 }}>
              <span style={{ color: C.gray }}>Somma categorie</span>
              <strong style={{ color: C.black }}>{formatEuro(sommaCategorie)}</strong>
            </div>
            {totaleRilevato !== null && (
              <p style={{ fontSize: 11, margin: '-6px 0 12px', color: mismatchTotale ? C.maroon : C.success }}>
                Totale complessivo letto dal PDF: {formatEuro(totaleRilevato)}
                {mismatchTotale ? ` — non coincide con la somma delle categorie (scarto ${formatEuro(scostamento)}): correggi le righe sopra se serve.` : ' — corrisponde alla somma delle categorie.'}
              </p>
            )}

            <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Totale complessivo impresa (inserito a mano — il PDF cliente non lo contiene)</label>
            <input value={totaleImpresa} onChange={(e) => setTotaleImpresa(e.target.value)} placeholder="0,00"
              style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 6px' }} />
            <p style={{ fontSize: 10, color: C.gray, margin: '0 0 16px' }}>
              Viene ripartito automaticamente tra le categorie in proporzione al loro totale cliente: potrai comunque correggere ogni voce in seguito.
            </p>
          </>
        )}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 6 }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          {stage === 'review' && (
            <button
              disabled={!canConfirm}
              onClick={() => onConfirm(rows.map((r) => ({ name: r.name.trim(), totale: parseEuro(r.totale) })), parseEuro(totaleImpresa))}
              style={{ background: canConfirm ? C.maroon : C.lightGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: canConfirm ? 'pointer' : 'default' }}
            >
              Crea revisione da questi totali
            </button>
          )}
        </div>
      </div>
    </div>
  , document.body);
}

function AddToComputoModal({ prodotto, variante, projects, onClose, onAdd }) {
  // Progetto e revisione si scelgono per POSIZIONE nell'elenco, non per id (gli id possono ripetersi nei dati).
  const eligible = projects.filter((p) => (p.revisions || []).length > 0);
  const [projectIdx, setProjectIdx] = useState(0);
  const project = eligible[projectIdx] || eligible[0];
  const [revIdx, setRevIdx] = useState(Math.max(0, (project?.revisions.length || 1) - 1));
  const revision = project?.revisions[revIdx] || project?.revisions[project.revisions.length - 1];
  const priceOptions = [
    { id: 'listino', name: 'Prezzo di listino (nessuno sconto)', prezzoListino: variante.listinoPrice, prezzoScontato: variante.listinoPrice, prezzoCliente: variante.listinoPrice },
    ...variante.fornitori,
  ];
  const [priceSourceId, setPriceSourceId] = useState(priceOptions[0]?.id ?? 'listino');
  const [qty, setQty] = useState('1');
  const selected = priceOptions.find((f) => f.id === priceSourceId) || priceOptions[0];
  const listinoNum = parseEuro(selected.prezzoListino);
  const clienteNum = parseEuro(selected.prezzoCliente);
  const scontoPct = listinoNum > 0 ? ((listinoNum - clienteNum) / listinoNum) * 100 : 0;

  const handleProjectChange = (idx) => {
    setProjectIdx(idx);
    setRevIdx(Math.max(0, (eligible[idx]?.revisions.length || 1) - 1));
  };

  if (eligible.length === 0) {
    return createPortal(
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200 }}>
        <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 380, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 16px' }}>Nessun progetto ha ancora un computo creato. Apri un progetto e crea il primo computo prima di aggiungere forniture.</p>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Chiudi</button>
        </div>
      </div>
    , document.body);
  }

  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 420, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 6px', color: C.black }}>Aggiungi al computo</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>{prodottoLabel(prodotto)}{variante.nome ? ` — ${variante.nome}` : ''}</p>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Progetto</label>
        <select value={projectIdx} onChange={(e) => handleProjectChange(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {eligible.map((p, i) => <option key={i} value={i}>{p.name}</option>)}
        </select>

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Computo (versione)</label>
        <select value={revIdx} onChange={(e) => setRevIdx(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' }}>
          {project?.revisions.map((r, i) => <option key={i} value={i}>{r.customName || r.label} · {r.dateModified}</option>)}
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
              onAdd(project, revision, selected, qty || '1');
              onClose();
            }}
            style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600 }}
          >
            Aggiungi
          </button>
        </div>
      </div>
    </div>
  , document.body);
}

// Campo di testo con suggerimenti "mentre scrivi": nessun elenco a tendina all'apertura, i valori già
// usati nel catalogo compaiono solo quando si digitano lettere che contengono (prima quelli che iniziano
// con quelle lettere). Si può sempre scrivere un valore nuovo: i suggerimenti sono solo un aiuto.
function SuggestInput({ value, onChange, options, placeholder, autoFocus, style }) {
  const [open, setOpen] = useState(false);
  const [hi, setHi] = useState(0);
  const q = String(value || '').trim().toLowerCase();
  const uniq = [...new Set((options || []).map((o) => String(o || '').trim()).filter(Boolean))];
  const matches = q
    ? uniq
      .filter((o) => o.toLowerCase().includes(q) && o.toLowerCase() !== q)
      .sort((a, b) => (a.toLowerCase().startsWith(q) ? 0 : 1) - (b.toLowerCase().startsWith(q) ? 0 : 1) || a.localeCompare(b, 'it'))
      .slice(0, 8)
    : [];
  const show = open && matches.length > 0;
  const pick = (o) => { onChange(o); setOpen(false); };
  const hl = (o) => {
    const i = o.toLowerCase().indexOf(q);
    if (i < 0) return o;
    return <>{o.slice(0, i)}<b style={{ fontWeight: 600, color: C.maroon }}>{o.slice(i, i + q.length)}</b>{o.slice(i + q.length)}</>;
  };
  return (
    <div style={{ position: 'relative' }}>
      <input
        value={value}
        autoFocus={autoFocus}
        placeholder={placeholder}
        autoComplete="off"
        onChange={(e) => { onChange(e.target.value); setOpen(true); setHi(0); }}
        onBlur={() => setTimeout(() => setOpen(false), 120)}
        onKeyDown={(e) => {
          if (!show) return;
          if (e.key === 'ArrowDown') { e.preventDefault(); setHi((hi + 1) % matches.length); }
          else if (e.key === 'ArrowUp') { e.preventDefault(); setHi((hi - 1 + matches.length) % matches.length); }
          else if (e.key === 'Enter') { e.preventDefault(); pick(matches[hi]); }
          else if (e.key === 'Escape') { setOpen(false); }
        }}
        style={style}
      />
      {show && (
        <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 50, background: '#fff', border: '1px solid #E3D9CC', borderRadius: 10, boxShadow: '0 8px 24px rgba(33,26,22,.10)', marginTop: 4, overflow: 'hidden' }}>
          {matches.map((o, i) => (
            <div
              key={o}
              onMouseDown={(e) => { e.preventDefault(); pick(o); }}
              onMouseEnter={() => setHi(i)}
              style={{ padding: '8px 12px', fontSize: 13, cursor: 'pointer', color: '#2B2321', background: i === hi ? '#F4EFE8' : '#fff' }}
            >
              {hl(o)}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Modale unica per prodotti e varianti del catalogo Fornitori (sostituisce le finestre "prompt" del browser,
// che non permettevano i suggerimenti). mode: 'nuovo' (marca, modello, variante, prezzo di listino),
// 'modifica' (marca, modello), 'variante-nuova' (variante, prezzo), 'variante-rinomina' (variante).
function ProdottoFornitoreModal({ mode, initial = {}, suggestions, onClose, onSave }) {
  const [marca, setMarca] = useState(initial.marca || '');
  const [modello, setModello] = useState(initial.modello || '');
  const [variante, setVariante] = useState(initial.variante || '');
  const [listinoPrice, setListinoPrice] = useState(initial.listinoPrice || '');
  const showProdotto = mode === 'nuovo' || mode === 'modifica';
  const showVariante = mode !== 'modifica';
  const showPrezzo = mode === 'nuovo' || mode === 'variante-nuova';
  const title = { nuovo: 'Nuovo prodotto', modifica: 'Marca e modello', 'variante-nuova': 'Nuova variante', 'variante-rinomina': 'Rinomina variante' }[mode];
  // I modelli suggeriti sono quelli della marca scritta, se esiste già; altrimenti tutti.
  const marcaKey = marca.trim().toLowerCase();
  const modelliDellaMarca = (suggestions.modelli || []).filter((x) => String(x.marca || '').trim().toLowerCase() === marcaKey).map((x) => x.modello);
  const modelliOpts = marcaKey && modelliDellaMarca.length ? modelliDellaMarca : (suggestions.modelli || []).map((x) => x.modello);
  const field = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, fontFamily: FONT, color: C.black, background: '#fff' };
  const label = { fontSize: 11, fontWeight: 700, color: C.midGray, display: 'block', margin: '0 0 4px' };
  const canSave = mode === 'modifica' ? !!modello.trim() : mode === 'nuovo' ? !!modello.trim() : mode === 'variante-rinomina' ? true : true;
  const save = () => {
    if (!canSave) return;
    onSave({ marca: marca.trim(), modello: modello.trim(), variante: variante.trim(), listinoPrice: listinoPrice.trim() || '0,00' });
    onClose();
  };
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && !e.defaultPrevented) save(); }} style={{ background: C.white, borderRadius: 14, padding: 22, width: 420, maxWidth: '100%' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, fontWeight: 400, margin: '0 0 4px', color: C.black }}>{title}</h2>
        <p style={{ fontSize: 11.5, color: C.gray, margin: '0 0 16px' }}>Mentre scrivi compaiono i valori già inseriti nel catalogo: sceglili oppure continua a scrivere per crearne uno nuovo.</p>
        {showProdotto && (
          <>
            <div style={{ marginBottom: 12 }}>
              <span style={label}>Marca</span>
              <SuggestInput value={marca} onChange={setMarca} options={suggestions.marche} placeholder="Es. Geberit" autoFocus style={field} />
            </div>
            <div style={{ marginBottom: 12 }}>
              <span style={label}>Modello</span>
              <SuggestInput value={modello} onChange={setModello} options={modelliOpts} placeholder="Es. Sigma 70" style={field} />
            </div>
          </>
        )}
        {showVariante && (
          <div style={{ marginBottom: 12 }}>
            <span style={label}>Variante / finitura {mode === 'nuovo' && <span style={{ fontWeight: 400, color: C.gray }}>(facoltativa)</span>}</span>
            <SuggestInput value={variante} onChange={setVariante} options={suggestions.varianti} placeholder="Es. bianco opaco, cromo…" autoFocus={!showProdotto} style={field} />
          </div>
        )}
        {showPrezzo && (
          <div style={{ marginBottom: 12 }}>
            <span style={label}>Prezzo di listino (€)</span>
            <input value={listinoPrice} onChange={(e) => setListinoPrice(e.target.value)} placeholder="0,00" style={field} />
          </div>
        )}
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 8 }}>
          <button onClick={onClose} style={{ background: 'transparent', color: C.black, border: `1px solid ${C.paleGray}`, padding: '9px 16px', borderRadius: 999, fontSize: 12.5, cursor: 'pointer', fontFamily: FONT }}>Annulla</button>
          <button onClick={save} disabled={!canSave} style={{ background: C.maroon, color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 500, cursor: canSave ? 'pointer' : 'default', opacity: canSave ? 1 : 0.5, fontFamily: FONT }}>Salva</button>
        </div>
      </div>
    </div>
  , document.body);
}

// Modale per aggiungere/modificare un fornitore di una variante: il nome ha i suggerimenti "mentre scrivi"
// tra i fornitori già presenti nel catalogo (così non nascono doppioni con nomi scritti in modo diverso),
// e i prezzi impresa/cliente si inseriscono come prezzo finale oppure come sconto % dal listino.
function FornitorePrezziModal({ initial = {}, isNew, suggestions, onClose, onSave }) {
  const fmt = (n) => n.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const [name, setName] = useState(initial.name || '');
  const [listino, setListino] = useState(initial.prezzoListino || '');
  const [impresa, setImpresa] = useState(initial.prezzoScontato || '');
  const [cliente, setCliente] = useState(initial.prezzoCliente || '');
  const [scImpresa, setScImpresa] = useState('');
  const [scCliente, setScCliente] = useState('');
  const fromSconto = (pctStr, listinoStr) => {
    const pct = parseFloat(String(pctStr).replace(',', '.'));
    const base = parseEuro(listinoStr);
    return !isNaN(pct) && base > 0 ? fmt(base * (1 - pct / 100)) : null;
  };
  const changeListino = (v) => {
    setListino(v);
    if (scImpresa) { const p = fromSconto(scImpresa, v); if (p) setImpresa(p); }
    if (scCliente) { const p = fromSconto(scCliente, v); if (p) setCliente(p); }
  };
  const field = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, fontFamily: FONT, color: C.black, background: '#fff' };
  const label = { fontSize: 11, fontWeight: 700, color: C.midGray, display: 'block', margin: '0 0 4px' };
  const priceRow = (title, sc, setSc, val, setVal) => (
    <div style={{ marginBottom: 12 }}>
      <span style={label}>{title}</span>
      <div style={{ display: 'flex', gap: 8 }}>
        <div style={{ position: 'relative', width: 110, flex: 'none' }}>
          <input value={sc} onChange={(e) => { setSc(e.target.value); const p = fromSconto(e.target.value, listino); if (p) setVal(p); }} placeholder="Sconto" style={{ ...field, paddingRight: 26 }} />
          <span style={{ position: 'absolute', right: 10, top: 9, fontSize: 13, color: C.gray }}>%</span>
        </div>
        <div style={{ position: 'relative', flex: 1 }}>
          <input value={val} onChange={(e) => { setVal(e.target.value); setSc(''); }} placeholder="0,00" style={{ ...field, paddingRight: 26 }} />
          <span style={{ position: 'absolute', right: 10, top: 9, fontSize: 13, color: C.gray }}>€</span>
        </div>
      </div>
    </div>
  );
  const canSave = !!name.trim();
  const save = () => {
    if (!canSave) return;
    const l = listino.trim() || '0,00';
    onSave({ name: name.trim(), prezzoListino: l, prezzoScontato: impresa.trim() || l, prezzoCliente: cliente.trim() || l });
    onClose();
  };
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && !e.defaultPrevented) save(); }} style={{ background: C.white, borderRadius: 14, padding: 22, width: 440, maxWidth: '100%' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, fontWeight: 400, margin: '0 0 4px', color: C.black }}>{isNew ? 'Nuovo fornitore' : 'Modifica fornitore'}</h2>
        <p style={{ fontSize: 11.5, color: C.gray, margin: '0 0 16px' }}>Mentre scrivi il nome compaiono i fornitori già inseriti nel catalogo: sceglilo oppure scrivine uno nuovo. Per i prezzi puoi scrivere lo sconto % sul listino oppure direttamente il prezzo.</p>
        <div style={{ marginBottom: 12 }}>
          <span style={label}>Fornitore</span>
          <SuggestInput value={name} onChange={setName} options={suggestions.fornitori} placeholder="Es. Edilceramica Rusconi" autoFocus style={field} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <span style={label}>Prezzo di listino del fornitore</span>
          <div style={{ position: 'relative' }}>
            <input value={listino} onChange={(e) => changeListino(e.target.value)} placeholder="0,00" style={{ ...field, paddingRight: 26 }} />
            <span style={{ position: 'absolute', right: 10, top: 9, fontSize: 13, color: C.gray }}>€</span>
          </div>
        </div>
        {priceRow('Prezzo impresa (scontato)', scImpresa, setScImpresa, impresa, setImpresa)}
        {priceRow('Prezzo cliente', scCliente, setScCliente, cliente, setCliente)}
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 8 }}>
          <button onClick={onClose} style={{ background: 'transparent', color: C.black, border: `1px solid ${C.paleGray}`, padding: '9px 16px', borderRadius: 999, fontSize: 12.5, cursor: 'pointer', fontFamily: FONT }}>Annulla</button>
          <button onClick={save} disabled={!canSave} style={{ background: C.maroon, color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 999, fontSize: 12.5, fontWeight: 500, cursor: canSave ? 'pointer' : 'default', opacity: canSave ? 1 : 0.5, fontFamily: FONT }}>Salva</button>
        </div>
      </div>
    </div>
  , document.body);
}

function ProdottoCard({ prodotto, onUpdate, onRemove, onAddToComputo, suggestions }) {
  const [modal, setModal] = useState(null); // { mode, vi?, initial }
  const uploadPhoto = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onUpdate({ ...prodotto, photo: reader.result });
    reader.readAsDataURL(file);
  };

  const renameProdotto = () => setModal({ mode: 'modifica', initial: { marca: prodotto.marca || '', modello: prodotto.modello || '' } });

  const addVariante = () => setModal({ mode: 'variante-nuova', initial: { listinoPrice: '' } });
  const renameVariante = (vi) => setModal({ mode: 'variante-rinomina', vi, initial: { variante: prodotto.varianti[vi].nome || '' } });
  const saveModal = (vals) => {
    if (modal.mode === 'modifica') {
      onUpdate({ ...prodotto, marca: vals.marca, modello: vals.modello });
    } else if (modal.mode === 'variante-nuova') {
      onUpdate({ ...prodotto, varianti: [...prodotto.varianti, { id: Date.now(), nome: vals.variante, listinoPrice: vals.listinoPrice, fornitori: [] }] });
    } else if (modal.mode === 'variante-rinomina') {
      onUpdate({ ...prodotto, varianti: prodotto.varianti.map((x, i) => (i === modal.vi ? { ...x, nome: vals.variante } : x)) });
    }
  };
  const updateVarianteListino = (vi) => {
    const v = prodotto.varianti[vi];
    const listinoPrice = prompt('Prezzo di listino (€):', v.listinoPrice);
    if (listinoPrice === null) return;
    onUpdate({ ...prodotto, varianti: prodotto.varianti.map((x, i) => (i === vi ? { ...x, listinoPrice } : x)) });
  };
  const removeVariante = (vi) => {
    if (prodotto.varianti.length <= 1) { alert('Il prodotto deve avere almeno una variante.'); return; }
    if (!confirm('Eliminare questa variante e i suoi fornitori?')) return;
    onUpdate({ ...prodotto, varianti: prodotto.varianti.filter((_, i) => i !== vi) });
  };

  const addFornitore = (vi) => {
    const v = prodotto.varianti[vi];
    setModal({ mode: 'fornitore', vi, fid: null, initial: { prezzoListino: v.listinoPrice || '', prezzoScontato: v.listinoPrice || '', prezzoCliente: v.listinoPrice || '' } });
  };
  const editFornitore = (vi, fid) => {
    const f = prodotto.varianti[vi].fornitori.find((x) => x.id === fid);
    setModal({ mode: 'fornitore', vi, fid, initial: { ...f } });
  };
  const saveFornitore = (vals) => {
    const { vi, fid } = modal;
    const v = prodotto.varianti[vi];
    const doppio = v.fornitori.find((y) => y.id !== fid && (y.name || '').trim().toLowerCase() === vals.name.toLowerCase());
    if (doppio && !confirm(`"${vals.name}" è già tra i fornitori di questa variante. Vuoi aggiungerlo comunque?`)) return;
    onUpdate({
      ...prodotto,
      varianti: prodotto.varianti.map((x, i) => (i !== vi ? x : {
        ...x,
        fornitori: fid ? x.fornitori.map((y) => (y.id === fid ? { ...y, ...vals } : y)) : [...x.fornitori, { id: Date.now(), ...vals }],
      })),
    });
  };

  const removeFornitore = (vi, fid) => {
    onUpdate({ ...prodotto, varianti: prodotto.varianti.map((x, i) => (i === vi ? { ...x, fornitori: x.fornitori.filter((y) => y.id !== fid) } : x)) });
  };

  const fBtn = { border: `1px solid ${C.paleGray}`, background: C.white, color: '#2B2321', fontSize: 11, fontWeight: 400, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT, whiteSpace: 'nowrap' };

  return (
    <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '14px 16px', marginBottom: 12, display: 'flex', gap: 14 }}>
      {modal && modal.mode !== 'fornitore' && <ProdottoFornitoreModal mode={modal.mode} initial={modal.initial} suggestions={suggestions || {}} onClose={() => setModal(null)} onSave={saveModal} />}
      {modal && modal.mode === 'fornitore' && <FornitorePrezziModal isNew={!modal.fid} initial={modal.initial} suggestions={suggestions || {}} onClose={() => setModal(null)} onSave={saveFornitore} />}
      <label style={{ width: 64, height: 64, borderRadius: 8, background: '#F0EAE0', border: '1px dashed #CFC4B6', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', overflow: 'hidden', textAlign: 'center' }}>
        {prodotto.photo ? (
          <img src={prodotto.photo} alt={prodottoLabel(prodotto)} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        ) : (
          <span style={{ fontSize: 11, color: '#9A8F86' }}>+ Foto</span>
        )}
        <input type="file" accept="image/*" style={{ display: 'none' }} onChange={(e) => uploadPhoto(e.target.files[0])} />
      </label>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 6, gap: 10, flexWrap: 'wrap' }}>
          <div>
            <p style={{ fontWeight: 500, fontSize: 14.5, margin: 0, color: C.black, fontFamily: FONT }}>{prodottoLabel(prodotto)}</p>
            <p style={{ fontSize: 12.5, color: '#8B8178', margin: '2px 0 0' }}>{prodotto.varianti.length} {prodotto.varianti.length === 1 ? 'variante' : 'varianti'}</p>
          </div>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            <button onClick={renameProdotto} style={fBtn}>✎ Marca/Modello</button>
            <button onClick={addVariante} style={fBtn}>+ Variante</button>
            <button onClick={onRemove} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 14 }}>🗑️</button>
          </div>
        </div>

        {prodotto.varianti.map((v, vi) => {
          const cheapestScontato = v.fornitori.length ? Math.min(...v.fornitori.map((f) => parseEuro(f.prezzoScontato))) : null;
          return (
            <div key={v.id} style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${C.paleGray}` }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 6 }}>
                <div>
                  <span style={{ fontSize: 13.5, fontWeight: 500, color: C.black }}>{v.nome || 'Variante senza nome'}</span>
                  <span style={{ fontSize: 12, color: '#8B8178', marginLeft: 8 }}>Listino: <strong style={{ color: C.black, fontWeight: 500 }}>{v.listinoPrice} €</strong></span>
                </div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  <button onClick={() => onAddToComputo(v)} style={{ ...fBtn, background: C.maroon, color: C.white, border: 'none' }}>+ Computo</button>
                  <button onClick={() => renameVariante(vi)} style={fBtn}>✎ Nome</button>
                  <button onClick={() => updateVarianteListino(vi)} style={fBtn}>✎ Listino</button>
                  <button onClick={() => addFornitore(vi)} style={fBtn}>+ Fornitore</button>
                  <button onClick={() => removeVariante(vi)} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 13 }}>🗑️</button>
                </div>
              </div>

              {v.fornitori.length === 0 ? (
                <p style={{ fontSize: 12.5, color: '#9A8F86', fontStyle: 'italic', margin: '6px 0 0' }}>Nessun fornitore inserito per questa variante.</p>
              ) : (
                <div className="table-scroll" style={{ marginTop: 8 }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 420 }}>
                    <thead>
                      <tr>
                        <th style={{ fontSize: 10, letterSpacing: '.08em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: 'left', padding: '5px 8px', borderBottom: `1px solid ${C.borderStrong}` }}>Fornitore</th>
                        <th style={{ fontSize: 10, letterSpacing: '.08em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: 'right', padding: '5px 8px', borderBottom: `1px solid ${C.borderStrong}` }}>Listino</th>
                        <th style={{ fontSize: 10, letterSpacing: '.08em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: 'right', padding: '5px 8px', borderBottom: `1px solid ${C.borderStrong}` }}>Scontato</th>
                        <th style={{ fontSize: 10, letterSpacing: '.08em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, textAlign: 'right', padding: '5px 8px', borderBottom: `1px solid ${C.borderStrong}` }}>Cliente</th>
                        <th style={{ borderBottom: `1px solid ${C.borderStrong}` }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...v.fornitori].sort((a, b) => parseEuro(a.prezzoScontato) - parseEuro(b.prezzoScontato)).map((f) => (
                        <tr key={f.id}>
                          <td style={{ fontSize: 12.5, padding: 8, borderBottom: '1px solid #F0EAE0' }}>
                            {f.name}
                            {parseEuro(f.prezzoScontato) === cheapestScontato && (
                              <span style={{ fontSize: 10, fontWeight: 500, color: C.maroon, background: `color-mix(in srgb, ${C.maroon} 16%, #fff)`, padding: '2px 8px', borderRadius: 999, marginLeft: 6, whiteSpace: 'nowrap' }}>Migliore offerta</span>
                            )}
                          </td>
                          <td style={{ fontSize: 12.5, padding: 8, borderBottom: '1px solid #F0EAE0', textAlign: 'right', color: '#6F655F' }}>{f.prezzoListino} €</td>
                          <td style={{ fontSize: 12.5, padding: 8, borderBottom: '1px solid #F0EAE0', textAlign: 'right', fontWeight: 600, color: C.black }}>{f.prezzoScontato} €</td>
                          <td style={{ fontSize: 12.5, padding: 8, borderBottom: '1px solid #F0EAE0', textAlign: 'right', color: C.maroon }}>{f.prezzoCliente} €</td>
                          <td style={{ padding: 8, borderBottom: '1px solid #F0EAE0', textAlign: 'right', whiteSpace: 'nowrap' }}>
                            <button onClick={() => editFornitore(vi, f.id)} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 12, marginRight: 6 }}>✎</button>
                            <button onClick={() => removeFornitore(vi, f.id)} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 12 }}>🗑️</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function FornitoriPage({ projects, setProjects, catalog, setCatalog }) {
  const [expanded, setExpanded] = useState({});
  const [addingTo, setAddingTo] = useState(null); // prodotto selezionato per l'aggiunta al computo
  const isOpen = (key) => expanded[key] !== false;
  const toggle = (key) => setExpanded({ ...expanded, [key]: !isOpen(key) });

  const addToComputo = (targetProject, targetRevision, priceSource, qty) => {
    const { prodotto, variante, tipoFornitura, gruppoFornitura } = addingTo;
    // Sezione unica "FORNITURE", sottosezione = tipo di fornitura (macrosezione del catalogo), nessun nome
    // del fornitore nella voce (resta solo nel campo nascosto "fornitore").
    const sectionName = 'FORNITURE';
    const desc = prodottoLabel(prodotto) + (variante.nome ? ` — ${variante.nome}` : '');

    // Se questo prodotto/variante è già stato aggiunto a questo computo, lo si segnala e si chiede
    // conferma prima di creare un doppione — rispondendo di no non cambia nulla.
    if (!targetProject) return;
    const dupe = (targetRevision?.items || []).find((it) => it.type !== 'subtotal' && String(it.desc || '').trim().toLowerCase() === desc.trim().toLowerCase());
    if (dupe && !confirm(`La voce "${desc}" è già presente nel computo${dupe.code ? ` (${dupe.code})` : ''}. Vuoi aggiungere comunque un duplicato?`)) {
      return;
    }

    const newItem = {
      id: Date.now() + Math.random(),
      code: '',
      desc,
      unit: 'cad',
      unitPriceImpresa: priceSource.prezzoScontato,
      unitPriceCliente: priceSource.prezzoCliente,
      listinoRef: priceSource.prezzoListino,
      qty,
      macro: sectionName,
      section: sectionName,
      categoria: tipoFornitura || 'Generale',
      sottocategoria: gruppoFornitura || 'Generale',
      fornitore: priceSource.id !== 'listino' ? priceSource.name : null,
    };
    setProjects(projects.map((p) => (p === targetProject ? addItemToProjectRevision(p, targetRevision, newItem) : p)));
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
  const [nuovoProdottoAt, setNuovoProdottoAt] = useState(null); // [mi, ci] della sottocategoria in cui aggiungere
  const allProdotti = catalog.flatMap((m) => (m.categorie || []).flatMap((c) => c.prodotti || []));
  const suggestions = {
    marche: allProdotti.map((p) => p.marca),
    modelli: allProdotti.map((p) => ({ marca: p.marca, modello: p.modello })),
    varianti: allProdotti.flatMap((p) => (p.varianti || []).map((v) => v.nome)),
    fornitori: allProdotti.flatMap((p) => (p.varianti || []).flatMap((v) => (v.fornitori || []).map((f) => f.name))),
  };
  const addProdotto = (mi, ci) => setNuovoProdottoAt([mi, ci]);
  const saveNuovoProdotto = ({ marca, modello, variante, listinoPrice }) => {
    const [mi, ci] = nuovoProdottoAt;
    const next = structuredClone(catalog);
    const prodotti = next[mi].categorie[ci].prodotti;
    // Stessa marca e stesso modello già presenti in questa sottocategoria: invece di un doppione si propone
    // di aggiungere la variante al prodotto esistente.
    const same = prodotti.find((p) => (p.marca || '').trim().toLowerCase() === marca.toLowerCase() && (p.modello || '').trim().toLowerCase() === modello.toLowerCase());
    if (same && confirm(`"${[marca, modello].filter(Boolean).join(' ')}" esiste già in questa sottocategoria. Vuoi aggiungere la variante a quel prodotto invece di crearne un altro?`)) {
      if (!same.varianti.some((v) => (v.nome || '').trim().toLowerCase() === variante.toLowerCase())) {
        same.varianti.push({ id: Date.now(), nome: variante, listinoPrice, fornitori: [] });
      }
    } else {
      prodotti.push({
        id: Date.now(), marca, modello, photo: null,
        varianti: [{ id: Date.now() + 1, nome: variante, listinoPrice, fornitori: [] }],
      });
    }
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

  const fBtn = { border: `1px solid ${C.paleGray}`, background: C.white, color: '#2B2321', fontSize: 11, fontWeight: 400, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT, whiteSpace: 'nowrap' };

  return (
    <div>
      <p style={{ fontSize: 12, color: '#9A8F86', margin: '0 0 6px' }}>Gestionale / Fornitori</p>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 18, flexWrap: 'wrap', justifyContent: 'space-between' }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Fornitori</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Dati aggiornati</span>
      </div>

      <div style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '18px 22px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 16, gap: 10, flexWrap: 'wrap' }}>
          <div>
            <p style={{ fontSize: 18, fontWeight: 300, margin: 0, color: C.black, fontFamily: FONT }}>Catalogo fornitori</p>
            <p style={{ fontSize: 11.5, color: '#8B8178', margin: '2px 0 0', maxWidth: 560 }}>Organizza per macrosezione e sottocategoria; per ogni prodotto confronta il prezzo di listino con quello dei diversi fornitori.</p>
          </div>
          <button onClick={addMacro} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 400, cursor: 'pointer', fontFamily: FONT }}>+ Macrosezione</button>
        </div>

        {catalog.map((m, mi) => {
          const mKey = `m${mi}`;
          return (
            <div key={mi} style={{ marginBottom: 12, border: '1px solid #2B2321', borderRadius: 12, overflow: 'hidden' }}>
              <div onClick={() => toggle(mKey)} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '12px 14px', background: '#2B2321', cursor: 'pointer', flexWrap: 'wrap' }}>
                <span style={{ color: '#B8ADA0', transform: isOpen(mKey) ? 'rotate(0deg)' : 'rotate(-90deg)', display: 'inline-block', fontSize: 11 }}>▾</span>
                <span style={{ fontWeight: 600, fontSize: 14.5, color: '#FFFFFF', flexGrow: 1 }}>{m.name}</span>
                <div onClick={(e) => e.stopPropagation()} style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  <button onClick={() => renameMacro(mi)} style={fBtn}>✎ Rinomina</button>
                  <button onClick={() => addCategoria(mi)} style={fBtn}>+ Sottocategoria</button>
                  <button onClick={() => removeMacro(mi)} style={{ border: 'none', background: 'transparent', color: '#B8ADA0', cursor: 'pointer', fontSize: 13 }}>🗑️</button>
                </div>
              </div>

              {isOpen(mKey) && m.categorie.map((c, ci) => {
                const cKey = `m${mi}c${ci}`;
                return (
                  <div key={ci} style={{ padding: '10px 14px 14px 30px' }}>
                    <div onClick={() => toggle(cKey)} style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', marginBottom: 10, flexWrap: 'wrap' }}>
                      <span style={{ color: '#8B8178', transform: isOpen(cKey) ? 'rotate(0deg)' : 'rotate(-90deg)', display: 'inline-block', fontSize: 11 }}>▾</span>
                      <span style={{ fontWeight: 500, fontSize: 13.5, color: C.black, flexGrow: 1 }}>{c.name}</span>
                      <div onClick={(e) => e.stopPropagation()} style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                        <button onClick={() => renameCategoria(mi, ci)} style={fBtn}>✎</button>
                        <button onClick={() => addProdotto(mi, ci)} style={fBtn}>+ Prodotto</button>
                        <button onClick={() => removeCategoria(mi, ci)} style={{ border: 'none', background: 'transparent', color: '#8B8178', cursor: 'pointer', fontSize: 13 }}>🗑️</button>
                      </div>
                    </div>
                    {isOpen(cKey) && (
                      c.prodotti.length === 0 ? (
                        <p style={{ fontSize: 12.5, color: '#9A8F86', fontStyle: 'italic' }}>Nessun prodotto ancora in questa sottocategoria.</p>
                      ) : (
                        c.prodotti.map((p, pi) => (
                          <ProdottoCard
                            key={p.id}
                            prodotto={p}
                            onUpdate={(updated) => updateProdotto(mi, ci, pi, updated)}
                            onRemove={() => removeProdotto(mi, ci, pi)}
                            onAddToComputo={(variante) => setAddingTo({ prodotto: p, variante, tipoFornitura: m.name, gruppoFornitura: c.name })}
                            suggestions={suggestions}
                          />
                        ))
                      )
                    )}
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {nuovoProdottoAt && (
        <ProdottoFornitoreModal mode="nuovo" suggestions={suggestions} onClose={() => setNuovoProdottoAt(null)} onSave={saveNuovoProdotto} />
      )}

      {addingTo && (
        <AddToComputoModal
          prodotto={addingTo.prodotto}
          variante={addingTo.variante}
          projects={projects}
          onClose={() => setAddingTo(null)}
          onAdd={addToComputo}
        />
      )}
    </div>
  );
}

// Gestione accessi della suite (funzione "admin-manage-team"): stessa forma di risposta di functions.invoke.
async function callAdminTeam(body) {
  const { data: { session } } = await supabase.auth.getSession();
  if (!session) return { data: null, error: { message: 'Sessione scaduta, ricarica la pagina e riaccedi.' } };
  const res = await fetch(SUITE_URL + '/functions/v1/admin-manage-team', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + session.access_token, apikey: SUITE_KEY },
    body: JSON.stringify(body),
  });
  let out = {};
  try { out = await res.json(); } catch (_e) { /* risposta non JSON */ }
  if (!res.ok || out.error) return { data: null, error: { message: out.error || ('Errore ' + res.status) } };
  return { data: out, error: null };
}
async function teamMemberAction(action, member, extra) {
  if (!member.auth_user_id) return { data: null, error: { message: 'Questa persona non ha un accesso collegato.' } };
  return callAdminTeam({ action, user_id: member.auth_user_id, ...extra });
}
async function teamCreate({ email, password, name, role }) {
  const c = await callAdminTeam({ action: 'create_team_login', email, password, full_name: name });
  if (c.error) return c;
  if (role === 'Admin') {
    const r = await callAdminTeam({ action: 'set_role', user_id: c.data.user_id, ruolo: 'amministratore' });
    if (r.error) return r;
  }
  await supabase.rpc('quant_io'); // allinea l'elenco persone
  return c;
}

// Il client supabase-js, quando una Edge Function risponde con uno stato non-2xx, mette in `error`
// un FunctionsHttpError generico ("Edge Function returned a non-2xx status code") e nasconde il vero
// messaggio (es. "Solo un amministratore può...") dentro `error.context`, che è la Response originale
// e va letta a parte. Questa funzione recupera il messaggio reale quando possibile.
async function extractFunctionErrorMessage(err, fallback) {
  try {
    if (err?.context && typeof err.context.json === 'function') {
      const body = await err.context.clone().json();
      if (body?.error) return body.error;
    }
  } catch (_e) { /* risposta non JSON: usa il messaggio generico sotto */ }
  return err?.message || fallback;
}

// Ridimensiona una foto scelta dall'utente a un quadrato (taglio centrato) in JPEG, così la scheda del team
// resta leggera da salvare nel database e da caricare (circa 20–40 KB invece di diversi MB).
function resizeImageToDataUrl(file, size = 256) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Lettura del file non riuscita.'));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error('Il file scelto non è un\'immagine valida.'));
      img.onload = () => {
        const side = Math.min(img.width, img.height);
        const canvas = document.createElement('canvas');
        canvas.width = size; canvas.height = size;
        canvas.getContext('2d').drawImage(img, (img.width - side) / 2, (img.height - side) / 2, side, side, 0, 0, size, size);
        resolve(canvas.toDataURL('image/jpeg', 0.85));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

const memberInitials = (nm) => (nm || '?').trim().split(/\s+/).slice(0, 2).map((p) => p[0]).join('').toUpperCase() || '?';

// Scelta delle persone del Team da assegnare a un progetto (selezione multipla, con ricerca).
function TeamPickerModal({ members, assigned, onClose, onConfirm }) {
  const [q, setQ] = useState('');
  const [sel, setSel] = useState([]);
  const terms = searchTerms(q);
  const list = members.filter((t) => t.status !== 'Disattivato' && t.status !== 'Rimosso').filter((t) => matchesTerms(terms, t.name, t.email, t.job_title, t.role));
  const toggle = (id) => setSel((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 16, padding: 22, width: 460, maxWidth: '100%', maxHeight: '86vh', display: 'flex', flexDirection: 'column' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, fontWeight: 600, margin: '0 0 4px', color: C.black }}>Assegna persone del team</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 12px' }}>Scegli una o più persone dalla sezione Team.</p>
        <input value={q} onChange={(e) => setQ(e.target.value)} autoFocus placeholder="Cerca per nome, email o ruolo…" style={{ border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 13, fontFamily: FONT, marginBottom: 10 }} />
        <div style={{ overflowY: 'auto', flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column', gap: 4 }}>
          {members.length === 0 && <p style={{ fontSize: 12.5, color: C.gray }}>Nessuna persona nel Team: aggiungile prima dalla sezione Team.</p>}
          {list.map((t) => {
            const isAssigned = assigned.includes(t.id);
            const isSel = sel.includes(t.id);
            return (
              <div key={t.id} onClick={() => !isAssigned && toggle(t.id)} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '9px 10px', borderRadius: 12, cursor: isAssigned ? 'default' : 'pointer', border: `1px solid ${isSel ? C.maroon : 'transparent'}`, background: isSel ? 'rgba(134,45,45,0.05)' : 'transparent', opacity: isAssigned ? 0.5 : 1 }}>
                <MemberAvatar member={t} size={36} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 500, color: C.black }}>{t.name || t.email}</div>
                  <div style={{ fontSize: 11.5, color: C.gray, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{[t.job_title || t.role, t.email].filter(Boolean).join(' · ')}</div>
                </div>
                {isAssigned ? <span style={{ fontSize: 11, color: C.gray }}>già assegnato</span> : (
                  <span style={{ width: 20, height: 20, borderRadius: 6, border: `1.5px solid ${isSel ? C.maroon : '#CFC4B6'}`, background: isSel ? C.maroon : C.white, color: C.white, fontSize: 13, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{isSel ? '✓' : ''}</span>
                )}
              </div>
            );
          })}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 14 }}>
          <button onClick={onClose} style={{ border: 'none', background: '#6B625C', color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>Annulla</button>
          <button disabled={!sel.length} onClick={() => onConfirm(members.filter((t) => sel.includes(t.id)))} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: sel.length ? 'pointer' : 'default', opacity: sel.length ? 1 : 0.6 }}>Assegna{sel.length ? ` (${sel.length})` : ''}</button>
        </div>
      </div>
    </div>,
    document.body
  );
}

// Nuova fattura / modifica fattura di un SAL: modalità "percentuale" (l'importo = % × totale categoria,
// calcolato e mostrato subito) oppure "importo" (si vede la % corrispondente). Avvisa se si supera il totale.
function FatturaSalModal({ categoria, base, giaFatturato, numero, fattura, onClose, onSave }) {
  const [label, setLabel] = useState(fattura?.label || `SAL ${numero}`);
  const [data, setData] = useState(fattura?.data || new Date().toISOString().slice(0, 10));
  const [mode, setMode] = useState(fattura ? (fattura.percentuale != null ? 'percentuale' : 'importo') : 'percentuale');
  const [pct, setPct] = useState(fattura?.percentuale != null ? itNum(fattura.percentuale, 2) : '');
  const [imp, setImp] = useState(fattura && fattura.percentuale == null ? itNum(fattura.importo || 0, 2) : '');
  const pctNum = parseEuro(pct);
  const importo = mode === 'percentuale' ? Math.round(base * pctNum) / 100 : parseEuro(imp);
  const pctEff = mode === 'percentuale' ? pctNum : (base ? (importo / base) * 100 : 0);
  const residuoDopo = base - giaFatturato - importo;
  const residuoPct = base ? Math.max(0, ((base - giaFatturato) / base) * 100) : 0;
  const switchMode = (m) => {
    if (m === mode) return;
    if (m === 'importo') setImp(importo ? itNum(importo, 2) : '');
    else setPct(pctEff ? itNum(pctEff, 2) : '');
    setMode(m);
  };
  const save = () => {
    if (!label.trim()) { alert('Scrivi una descrizione per la fattura.'); return; }
    if (!(importo > 0)) { alert(mode === 'percentuale' ? 'Inserisci una percentuale maggiore di zero.' : 'Inserisci un importo maggiore di zero.'); return; }
    if (residuoDopo < -0.005 && !confirm(`Con questa fattura il fatturato supera il totale della categoria di ${formatEuro(-residuoDopo)}. Salvare comunque?`)) return;
    onSave({ ...(fattura || {}), label: label.trim(), data, importo: Math.round(importo * 100) / 100, percentuale: mode === 'percentuale' ? pctNum : null });
  };
  const inputSt = { width: '100%', boxSizing: 'border-box', border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '9px 11px', fontSize: 13.5, fontFamily: FONT, color: C.black };
  const lab = { display: 'block', fontSize: 11, fontWeight: 600, color: C.black, marginBottom: 5 };
  const seg = (key, text) => (
    <button onClick={() => switchMode(key)} style={{ flex: 1, border: 'none', borderRadius: 999, padding: '7px 10px', fontSize: 12.5, fontWeight: 500, cursor: 'pointer', fontFamily: FONT, background: mode === key ? C.black : 'transparent', color: mode === key ? C.white : C.darkGray }}>{text}</button>
  );
  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 16, padding: 22, width: 440, maxWidth: '100%' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, fontWeight: 600, margin: '0 0 4px', color: C.black }}>{fattura ? 'Modifica fattura' : 'Nuova fattura'}</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>{categoria} · totale categoria {formatEuro(base)} IVA esclusa · già fatturato {formatEuro(giaFatturato)}</p>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 150px', gap: 10, marginBottom: 14 }}>
          <div><label style={lab}>Descrizione</label><input value={label} onChange={(e) => setLabel(e.target.value)} style={inputSt} placeholder="Es. SAL 1, Fattura n.3" /></div>
          <div><label style={lab}>Data</label><input type="date" value={data} onChange={(e) => setData(e.target.value)} style={inputSt} /></div>
        </div>
        <div style={{ display: 'flex', gap: 4, background: '#F4F0EA', borderRadius: 999, padding: 4, marginBottom: 12 }}>
          {seg('percentuale', 'In percentuale')}
          {seg('importo', 'Importo')}
        </div>
        {mode === 'percentuale' ? (
          <div style={{ marginBottom: 12 }}>
            <label style={lab}>Percentuale del totale categoria</label>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <input value={pct} onChange={(e) => setPct(e.target.value)} autoFocus inputMode="decimal" placeholder="Es. 30" style={{ ...inputSt, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }} />
              <span style={{ fontSize: 15, color: C.gray }}>%</span>
            </div>
            <div style={{ display: 'flex', gap: 6, marginTop: 8, flexWrap: 'wrap' }}>
              {[10, 20, 25, 30, 50].map((v) => (
                <button key={v} onClick={() => setPct(String(v))} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '4px 10px', fontSize: 11.5, cursor: 'pointer' }}>{v}%</button>
              ))}
              {residuoPct > 0.005 && <button onClick={() => setPct(itNum(residuoPct, 2))} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '4px 10px', fontSize: 11.5, cursor: 'pointer' }}>Saldo ({itNum(residuoPct, 2)}%)</button>}
            </div>
          </div>
        ) : (
          <div style={{ marginBottom: 12 }}>
            <label style={lab}>Importo (€, IVA esclusa)</label>
            <input value={imp} onChange={(e) => setImp(e.target.value)} autoFocus inputMode="decimal" placeholder="0,00" style={{ ...inputSt, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }} />
          </div>
        )}
        <div style={{ background: '#FAF8F4', border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: '12px 14px', marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
            <span style={{ fontSize: 12, color: C.gray }}>Importo fattura</span>
            <span style={{ fontSize: 20, fontWeight: 600, color: C.maroon, fontVariantNumeric: 'tabular-nums' }}>{formatEuro(importo)}</span>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: C.gray, marginTop: 4 }}>
            <span>pari al {itNum(pctEff, 2)}% del totale categoria</span>
            <span style={{ color: residuoDopo < -0.005 ? C.maroon : C.gray }}>residuo dopo: {formatEuro(residuoDopo)}</span>
          </div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button onClick={onClose} style={{ border: 'none', background: '#6B625C', color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>Annulla</button>
          <button onClick={save} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>{fattura ? 'Salva modifiche' : 'Aggiungi fattura'}</button>
        </div>
      </div>
    </div>,
    document.body
  );
}

// --- Barra azioni del computo ---
const tbBtn = (primary = false) => ({
  background: primary ? C.maroon : C.white, color: primary ? C.white : C.black,
  border: primary ? 'none' : `1px solid ${C.paleGray}`, borderRadius: 999, padding: '8px 14px',
  fontSize: 12.5, fontWeight: 500, fontFamily: FONT, cursor: 'pointer', whiteSpace: 'nowrap',
});

function ToolbarGroup({ label, last = false, children }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, padding: '0 14px', borderRight: last ? 'none' : `1px solid ${C.paleGray}`, margin: '2px 0' }}>
      <span style={{ fontSize: 9.5, fontWeight: 600, letterSpacing: '0.12em', textTransform: 'uppercase', color: '#8B8178' }}>{label}</span>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>{children}</div>
    </div>
  );
}

// Pulsante con menu a tendina (es. "PDF ▾" con le varianti di esportazione). Si chiude cliccando fuori o con Esc.
function ExportMenu({ label, items }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  React.useEffect(() => {
    if (!open) return;
    const onDown = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    const onKey = (e) => { if (e.key === 'Escape') setOpen(false); };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => { document.removeEventListener('mousedown', onDown); document.removeEventListener('keydown', onKey); };
  }, [open]);
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button onClick={() => setOpen((o) => !o)} style={{ ...tbBtn(), display: 'inline-flex', alignItems: 'center', gap: 7, background: open ? '#F4F0EA' : C.white }}>
        <svg viewBox="0 0 24 24" width={14} height={14} fill="none" stroke="currentColor" strokeWidth={1.6} strokeLinecap="round" strokeLinejoin="round"><path d="M12 4v11" /><path d="M7.5 10.5 12 15l4.5-4.5" /><path d="M5 19.5h14" /></svg>
        {label}
        <span style={{ fontSize: 10, color: '#8B8178', transform: open ? 'rotate(180deg)' : 'none', transition: 'transform .15s' }}>▾</span>
      </button>
      {open && (
        <div style={{ position: 'absolute', top: 'calc(100% + 6px)', left: 0, zIndex: 40, background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, boxShadow: '0 12px 32px rgba(33,26,22,0.14)', padding: 6, minWidth: 310 }}>
          {items.map((it, i) => (
            <button
              key={i}
              onClick={() => { setOpen(false); it.onClick(); }}
              className="export-menu-item"
              style={{ display: 'block', width: '100%', textAlign: 'left', border: 'none', background: 'transparent', borderRadius: 10, padding: '9px 11px', cursor: 'pointer', fontFamily: FONT }}
            >
              <span style={{ display: 'block', fontSize: 13, fontWeight: 500, color: C.black }}>{it.label}</span>
              <span style={{ display: 'block', fontSize: 11.5, color: '#8B8178', marginTop: 1 }}>{it.desc}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// Primo avvio nella suite: porta qui i dati del Gestionale attuale (copia, l'originale non viene toccato).
function ImportaDaVecchioGestionale({ isAdmin, onDone, onSkip }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [msg, setMsg] = useState('');
  const [busy, setBusy] = useState(false);
  const box = { minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT };
  const card = { background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 400, maxWidth: 'calc(100vw - 32px)' };
  const inp = { width: '100%', boxSizing: 'border-box', border: '1px solid #E3D9CC', borderRadius: 8, padding: '10px 12px', fontSize: 13, fontFamily: FONT, marginBottom: 10 };
  const run = async () => {
    setMsg(''); setBusy(true);
    try {
      const old = createClient(OLD_URL, OLD_KEY, { auth: { persistSession: false, autoRefreshToken: false } });
      const li = await old.auth.signInWithPassword({ email: email.trim(), password });
      if (li.error) { setMsg('Accesso al Gestionale attuale non riuscito.'); setBusy(false); return; }
      const st = await old.schema('cea').from('app_state').select('data').eq('id', 1).maybeSingle();
      if (st.error || !st.data || !st.data.data) { setMsg('Non riesco a leggere i dati: ' + (st.error?.message || 'vuoti')); setBusy(false); return; }
      const tm = await old.schema('cea').from('team_members').select('*');
      const lk = await old.schema('cea').from('fornitore_links').select('*');
      const sb = await old.schema('cea').from('fornitore_submissions').select('*');
      const ol = await old.schema('cea').from('computo_offerta_links').select('*');
      const of = await old.schema('cea').from('computo_offerte').select('*');
      const r = await supabase.rpc('quant_salva', { p_dati: st.data.data });
      if (r.error) { setMsg('Salvataggio non riuscito: ' + r.error.message); setBusy(false); return; }
      const t = await supabase.rpc('quant_importa_extra', {
        p_team: tm.data || [], p_fornitore_links: lk.data || [], p_fornitore_submissions: sb.data || [],
        p_offerta_links: ol.data || [], p_offerte: of.data || [],
      });
      await old.auth.signOut();
      if (t.error) { setMsg('Dati copiati, ma team e link non importati: ' + t.error.message); setBusy(false); return; }
      if (r.data && r.data.proiezione_errore) alert('Dati copiati. Il riallineamento delle tabelle ha segnalato: ' + r.data.proiezione_errore);
      onDone();
    } catch (e) { setMsg('Errore: ' + (e.message || e)); setBusy(false); }
  };
  return (
    <div style={box}><div style={card}>
      <h2 style={{ margin: '0 0 8px', fontWeight: 300, fontSize: 22 }}>Primo avvio nella suite</h2>
      {isAdmin ? (<>
        <p style={{ fontSize: 13, color: C.black, margin: '0 0 14px' }}>Inserisci le credenziali con cui entri oggi nel Gestionale: i dati vengono copiati senza toccare l'originale e sostituiscono quelli presenti qui.</p>
        <input style={inp} type="email" placeholder="Email del Gestionale attuale" value={email} onChange={(e) => setEmail(e.target.value)} />
        <input style={inp} type="password" placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} />
        {msg && <p style={{ fontSize: 12.5, color: C.maroon }}>{msg}</p>}
        <button disabled={busy || !email || !password} onClick={run} style={{ background: C.maroon, color: '#fff', border: 'none', borderRadius: 999, padding: '10px 20px', fontSize: 13, cursor: 'pointer' }}>{busy ? 'Copia in corso…' : 'Copia i dati'}</button>
        <button onClick={onSkip} style={{ background: 'none', border: 'none', color: '#8B8178', fontSize: 12, marginLeft: 12, cursor: 'pointer' }}>Annulla</button>
      </>) : <p style={{ fontSize: 13 }}>Un amministratore deve ancora copiare i dati dal Gestionale attuale.</p>}
    </div></div>
  );
}

function MemberAvatar({ member, size = 38 }) {
  const base = { width: size, height: size, borderRadius: '50%', flex: 'none', border: '1px solid #E3D9CC', boxSizing: 'border-box' };
  if (member?.photo) return <img src={member.photo} alt="" style={{ ...base, objectFit: 'cover', display: 'block' }} />;
  return (
    <span style={{ ...base, background: '#F1EBE2', color: '#5A4F4A', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontWeight: 500, letterSpacing: '.02em', fontSize: size >= 34 ? Math.round(size * 0.33) : 11 }}>
      {memberInitials(member?.name || member?.email)}
    </span>
  );
}

function generatePassword(len = 10) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const arr = new Uint32Array(len);
  (window.crypto || window.msCrypto).getRandomValues(arr);
  return Array.from(arr, (n) => chars[n % chars.length]).join('');
}

// Scheda di una persona del team: tutti la possono aprire, solo un admin può modificarla
// (lo impone anche il database: la tabella accetta modifiche solo da un admin attivo).
function TeamMemberModal({ member, isAdmin, isSelf, onClose, onSaved, onRemoved }) {
  const [name, setName] = useState(member.name || '');
  const [jobTitle, setJobTitle] = useState(member.job_title || '');
  const [email, setEmail] = useState(member.email || '');
  const [phone, setPhone] = useState(member.phone || '');
  const [notes, setNotes] = useState(member.notes || '');
  const [role, setRole] = useState(member.role || 'Membro');
  const [photo, setPhoto] = useState(member.photo || null);
  const [newPassword, setNewPassword] = useState('');
  const [showPw, setShowPw] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');
  const ro = !isAdmin;

  const field = { width: '100%', boxSizing: 'border-box', border: '1px solid #E3D9CC', background: ro ? '#FAF8F4' : '#fff', borderRadius: 8, padding: '9px 12px', fontSize: 13, fontFamily: FONT, color: '#2B2321' };
  const label = { fontSize: 11, fontWeight: 600, color: '#6F655F', display: 'block', margin: '0 0 4px', letterSpacing: '.02em' };
  const pillBtn = { border: '1px solid #E3D9CC', background: '#fff', color: '#2B2321', borderRadius: 999, padding: '7px 14px', fontSize: 12, fontWeight: 300, fontFamily: FONT, cursor: 'pointer' };

  const pickPhoto = async (file) => {
    if (!file) return;
    try { setPhoto(await resizeImageToDataUrl(file)); setError(''); } catch (e) { setError(e.message); }
  };

  const save = async () => {
    if (!isAdmin) return;
    if (!name.trim()) { setError('Il nome è obbligatorio.'); return; }
    if (newPassword && newPassword.length < 8) { setError('La nuova password deve avere almeno 8 caratteri.'); return; }
    if (isSelf && role !== 'Admin' && member.role === 'Admin' && !confirm('Stai togliendo a te stesso il ruolo di Admin: non potrai più modificare il team. Continuare?')) return;
    setSaving(true); setError(''); setInfo('');
    if (role !== member.role && member.auth_user_id) {
      const { error: roleErr } = await teamMemberAction('set_role', member, { ruolo: role === 'Admin' ? 'amministratore' : 'socio' });
      if (roleErr) { setSaving(false); setError('Ruolo non cambiato: ' + roleErr.message); return; }
    }
    const { data: updated, error: upErr } = await cea.from('team_members')
      .update({ name: name.trim(), job_title: jobTitle.trim() || null, phone: phone.trim() || null, notes: notes.trim() || null, role, photo })
      .eq('id', member.id).select().maybeSingle();
    if (upErr || !updated) { setSaving(false); setError(upErr?.message || 'Salvataggio non riuscito (servono i permessi di admin).'); return; }
    let finalRow = updated;
    const newEmail = email.trim().toLowerCase();
    if (newEmail && newEmail !== (member.email || '').toLowerCase()) {
      const { data, error: err } = await teamMemberAction('update_email', member, { email: newEmail });
      const msg = err ? await extractFunctionErrorMessage(err, 'Cambio email non riuscito.') : data?.error;
      if (msg) { setSaving(false); onSaved(finalRow); setError(`Dati salvati, ma l'email non è stata cambiata: ${msg}`); return; }
      finalRow = { ...finalRow, email: newEmail };
    }
    if (newPassword) {
      const { data, error: err } = await teamMemberAction('reset_password', member, { password: newPassword });
      const msg = err ? await extractFunctionErrorMessage(err, 'Cambio password non riuscito.') : data?.error;
      if (msg) { setSaving(false); onSaved(finalRow); setError(`Dati salvati, ma la password non è stata cambiata: ${msg}`); return; }
    }
    setSaving(false);
    onSaved(finalRow);
    setNewPassword('');
    setInfo(newPassword ? 'Salvato. Nuova password attiva: comunicala alla persona di persona o in chat privata.' : 'Salvato.');
  };

  const remove = async () => {
    if (!confirm(`Rimuovere ${member.name} dal team? Perderà l'accesso al workspace.`)) return;
    if (member.auth_user_id) {
      const { error: susErr } = await teamMemberAction('set_suspended', member, { suspended: true });
      if (susErr) { setError('Accesso non disattivato: ' + susErr.message); return; }
    }
    const { error: err } = await cea.from('team_members').delete().eq('id', member.id);
    if (err) { setError(err.message); return; }
    onRemoved(member.id);
  };

  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: '#FFFDFA', borderRadius: 16, width: 560, maxWidth: '100%', maxHeight: '92vh', overflowY: 'auto', padding: '22px 24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 20 }}>
          <MemberAvatar member={{ ...member, name, photo }} size={72} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 20, fontWeight: 300, color: '#2B2321', letterSpacing: '-.01em' }}>{name || member.email}</div>
            <div style={{ fontSize: 12.5, color: '#8B8178', fontWeight: 300 }}>{jobTitle || (role === 'Admin' ? 'Amministratore' : 'Membro del team')}</div>
            {isAdmin && (
              <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
                <label style={{ ...pillBtn, display: 'inline-block' }}>
                  {photo ? 'Cambia foto' : 'Carica foto'}
                  <input type="file" accept="image/*" style={{ display: 'none' }} onChange={(e) => pickPhoto(e.target.files[0])} />
                </label>
                {photo && <button onClick={() => setPhoto(null)} style={pillBtn}>Rimuovi foto</button>}
              </div>
            )}
          </div>
          <button onClick={onClose} aria-label="Chiudi" style={{ alignSelf: 'flex-start', border: 'none', background: 'transparent', fontSize: 18, color: '#8B8178', cursor: 'pointer' }}>✕</button>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }} className="form-grid">
          <div><span style={label}>Nome e cognome</span><input value={name} onChange={(e) => setName(e.target.value)} readOnly={ro} style={field} /></div>
          <div><span style={label}>Qualifica</span><input value={jobTitle} onChange={(e) => setJobTitle(e.target.value)} readOnly={ro} placeholder="Es. Architetto, CEO…" style={field} /></div>
          <div><span style={label}>Email di accesso</span><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} readOnly={ro} style={field} /></div>
          <div><span style={label}>Telefono</span><input value={phone} onChange={(e) => setPhone(e.target.value)} readOnly={ro} style={field} /></div>
          <div>
            <span style={label}>Ruolo nell'app</span>
            {isAdmin ? (
              <select value={role} onChange={(e) => setRole(e.target.value)} style={field}><option value="Admin">Admin</option><option value="Membro">Membro</option></select>
            ) : <input value={role} readOnly style={field} />}
          </div>
          <div><span style={label}>Stato</span><input value={member.status || ''} readOnly style={{ ...field, background: '#FAF8F4' }} /></div>
        </div>
        <div style={{ marginTop: 12 }}>
          <span style={label}>Note</span>
          <textarea value={notes} onChange={(e) => setNotes(e.target.value)} readOnly={ro} rows={2} style={{ ...field, resize: 'vertical' }} />
        </div>

        {isAdmin && (
          <div style={{ marginTop: 16, border: '1px solid #E8E0D4', borderRadius: 12, padding: '14px 16px', background: '#fff' }}>
            <div style={{ fontSize: 14, fontWeight: 400, color: '#2B2321', marginBottom: 4 }}>Password di accesso</div>
            <p style={{ fontSize: 11.5, color: '#8B8178', margin: '0 0 10px', fontWeight: 300 }}>
              Per sicurezza le password sono salvate criptate: nessuno, nemmeno l'admin, può leggere quella attuale. Puoi impostarne una nuova qui — la vedi mentre la scrivi — e comunicarla alla persona.
            </p>
            {member.auth_user_id ? (
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                <input type={showPw ? 'text' : 'password'} value={newPassword} onChange={(e) => setNewPassword(e.target.value)} placeholder="Nuova password (min. 6 caratteri)" autoComplete="new-password" style={{ ...field, flex: '1 1 200px', width: 'auto', fontVariantNumeric: 'tabular-nums' }} />
                <button onClick={() => setShowPw(!showPw)} style={pillBtn}>{showPw ? 'Nascondi' : 'Mostra'}</button>
                <button onClick={() => { setNewPassword(generatePassword()); setShowPw(true); }} style={pillBtn}>Genera</button>
                {newPassword && <button onClick={() => navigator.clipboard?.writeText(newPassword)} style={pillBtn}>Copia</button>}
              </div>
            ) : (
              <p style={{ fontSize: 12, color: '#6F655F', margin: 0 }}>Questa persona non ha ancora un account collegato.</p>
            )}
          </div>
        )}

        {error && <p style={{ fontSize: 12.5, color: C.maroon, margin: '12px 0 0' }}>{error}</p>}
        {info && <p style={{ fontSize: 12.5, color: '#2F5A2B', margin: '12px 0 0' }}>{info}</p>}

        <div style={{ display: 'flex', gap: 10, justifyContent: 'space-between', alignItems: 'center', marginTop: 18, flexWrap: 'wrap' }}>
          <div>{isAdmin && !isSelf && <button onClick={remove} style={{ ...pillBtn, color: C.maroon }}>Rimuovi dal team</button>}</div>
          <div style={{ display: 'flex', gap: 10 }}>
            <button onClick={onClose} style={pillBtn}>{isAdmin ? 'Chiudi' : 'Chiudi'}</button>
            {isAdmin && (
              <button onClick={save} disabled={saving} style={{ border: 'none', background: C.maroon, color: '#fff', borderRadius: 999, padding: '9px 18px', fontSize: 12.5, fontWeight: 400, fontFamily: FONT, cursor: saving ? 'default' : 'pointer', opacity: saving ? 0.7 : 1 }}>
                {saving ? 'Salvataggio…' : 'Salva modifiche'}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  , document.body);
}

function TeamPage({ profile, onProfileChange }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState('');
  const [name, setName] = useState('');
  const [password, setPassword] = useState('');
  const [role, setRole] = useState('Membro');
  const [error, setError] = useState('');
  const [creating, setCreating] = useState(false);
  const [openId, setOpenId] = useState(null);
  const isAdmin = profile?.role === 'Admin';

  const loadMembers = async () => {
    setLoading(true);
    const { data, error: err } = await cea.from('team_members').select('*').order('created_at');
    if (!err) setMembers(data || []);
    setLoading(false);
  };

  React.useEffect(() => { loadMembers(); }, []);

  // Crea (o completa, se c'era già una vecchia riga "Invito inviato" con la stessa email) l'accesso
  // di una persona: l'admin sceglie qui email e password, l'account viene attivato subito lato server
  // (funzione Edge "cea-admin-create-user", con service role) senza inviare alcuna email di conferma.
  const handleCreate = async () => {
    const trimmedEmail = email.trim();
    if (!name.trim() || !trimmedEmail || !password) { setError('Nome, email e password sono obbligatori.'); return; }
    if (password.length < 8) { setError('La password deve avere almeno 8 caratteri.'); return; }
    setError(''); setCreating(true);
    const { data, error: err } = await teamCreate({ email: trimmedEmail, password, name: name.trim(), role });
    setCreating(false);
    if (err) { setError(await extractFunctionErrorMessage(err, 'Creazione non riuscita.')); return; }
    if (data?.error) { setError(data.error); return; }
    setEmail(''); setName(''); setPassword(''); setRole('Membro');
    loadMembers();
  };

  const onSaved = (row) => {
    setMembers((ms) => ms.map((m) => (m.id === row.id ? row : m)));
    if (row.id === profile?.id && onProfileChange) onProfileChange(row);
  };
  const onRemoved = (id) => { setMembers((ms) => ms.filter((m) => m.id !== id)); setOpenId(null); };

  const inputStyle = { border: '1px solid #E3D9CC', background: '#fff', borderRadius: 8, padding: '9px 12px', fontSize: 13, fontFamily: FONT, boxSizing: 'border-box' };
  const openMember = members.find((m) => m.id === openId);

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 18, flexWrap: 'wrap', justifyContent: 'space-between' }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Team</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Workspace condiviso</span>
      </div>

      {!isAdmin && (
        <div style={{ background: '#FFFFFF', border: '1px solid #E8E0D4', borderRadius: 14, padding: '14px 18px', marginBottom: 16 }}>
          <p style={{ fontSize: 12.5, color: '#4A403C', margin: 0, fontWeight: 300 }}>Sei collegato come <strong style={{ fontWeight: 500 }}>membro</strong>: puoi consultare le schede del team, ma solo un admin può modificarle o creare nuovi accessi.</p>
        </div>
      )}

      <section style={{ background: '#FFFFFF', border: '1px solid #E8E0D4', borderRadius: 14, padding: '18px 22px', marginBottom: 16 }}>
        <h3 style={{ margin: '0 0 14px', fontWeight: 300, fontSize: 20, color: '#2B2321', display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap' }}>
          PERSONE CON ACCESSO<span style={{ fontSize: 11.5, fontWeight: 400, color: '#8B8178' }}>clicca una scheda per aprirla</span>
        </h3>
        {loading && <p style={{ fontSize: 12.5, color: '#8B8178' }}>Caricamento…</p>}
        {!loading && members.length === 0 && <p style={{ fontSize: 12.5, color: '#8B8178' }}>Nessuna persona ancora.</p>}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))', gap: 12 }}>
          {members.map((m) => (
            <div
              key={m.id}
              onClick={() => setOpenId(m.id)}
              className="team-card"
              style={{ border: '1px solid #E8E0D4', borderRadius: 14, padding: 16, background: '#fff', cursor: 'pointer', display: 'flex', flexDirection: 'column', gap: 12 }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <MemberAvatar member={m} size={52} />
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontSize: 15, fontWeight: 400, color: '#2B2321', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.name}</div>
                  <div style={{ fontSize: 12, color: '#8B8178', fontWeight: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.job_title || (m.role === 'Admin' ? 'Amministratore' : 'Membro del team')}</div>
                </div>
              </div>
              <div style={{ fontSize: 12.5, color: '#6F655F', fontWeight: 300, lineHeight: 1.5, minHeight: 38 }}>
                <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{m.email}</div>
                {m.phone && <div>{m.phone}</div>}
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, ...(m.role === 'Admin' ? { background: '#2B2321', color: '#fff' } : { background: '#EFE7DA', color: '#5A4F4A' }) }}>{m.role}</span>
                <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, ...(m.status === 'Attivo' ? { background: 'color-mix(in srgb, #4F7A4B 14%, #fff)', color: '#2F5A2B' } : { background: '#EFE7DA', color: '#6F655F' }) }}>{m.status}</span>
                {m.id === profile?.id && <span style={{ fontSize: 11, color: '#8B8178' }}>· tu</span>}
                <span style={{ marginLeft: 'auto', fontSize: 12, color: '#8B8178' }}>Apri ›</span>
              </div>
            </div>
          ))}
        </div>
      </section>

      {isAdmin && (
        <section style={{ background: '#FFFFFF', border: '1px solid #E8E0D4', borderRadius: 14, padding: '18px 22px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 12px', fontWeight: 300, fontSize: 20, color: '#2B2321' }}>CREA UN NUOVO ACCESSO</h3>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'stretch' }}>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome e cognome" style={{ ...inputStyle, flex: '1 1 160px', maxWidth: 200 }} />
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="nome@studio.it" style={{ ...inputStyle, flex: '1 1 180px', maxWidth: 230 }} />
            <input type="text" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Password (min. 6 caratt.)" style={{ ...inputStyle, flex: '1 1 160px', maxWidth: 200 }} />
            <button onClick={() => setPassword(generatePassword())} style={{ border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '9px 14px', fontSize: 12.5, fontWeight: 300, fontFamily: FONT, cursor: 'pointer', color: '#2B2321' }}>Genera</button>
            <select value={role} onChange={(e) => setRole(e.target.value)} style={{ border: '1px solid #E3D9CC', background: '#fff', borderRadius: 10, padding: '9px 14px', fontSize: 13, fontFamily: FONT, color: '#2B2321' }}>
              <option value="Membro">Membro</option>
              <option value="Admin">Admin</option>
            </select>
            <button
              onClick={handleCreate}
              disabled={creating}
              style={{ border: 'none', background: C.maroon, color: '#fff', padding: '10px 20px', borderRadius: 999, fontSize: 13, fontWeight: 400, fontFamily: FONT, cursor: creating ? 'default' : 'pointer', opacity: creating ? 0.7 : 1 }}
            >
              {creating ? 'Un attimo…' : 'Crea accesso'}
            </button>
          </div>
          {error && <p style={{ fontSize: 12, color: C.maroon, margin: '10px 0 0' }}>{error}</p>}
          <p style={{ fontSize: 11.5, color: '#8B8178', margin: '10px 0 0', maxWidth: 640, fontWeight: 300 }}>
            L'account viene attivato subito: non parte nessuna email di conferma. Comunica tu email e password alla persona (di persona, in chat privata…). Foto, qualifica, telefono e note si completano poi dalla sua scheda.
          </p>
        </section>
      )}

      {openMember && (
        <TeamMemberModal
          key={openMember.id}
          member={openMember}
          isAdmin={isAdmin}
          isSelf={openMember.id === profile?.id}
          onClose={() => setOpenId(null)}
          onSaved={onSaved}
          onRemoved={onRemoved}
        />
      )}
    </div>
  );
}

// Colore di accento dello studio: vale per tutta l'app (per tutto il team, è salvato nel workspace) e per i PDF.
function AccentColorCard({ settings, onUpdate, sCard, sH3, sTag, sFlabel, sInput, sBtn }) {
  const current = isHexColor(settings.coloreAccento) ? settings.coloreAccento : DEFAULT_ACCENT;
  const [hexDraft, setHexDraft] = useState(current);
  React.useEffect(() => { setHexDraft(current); }, [current]);
  const tooLight = hexLuminance(current) > 0.45;
  return (
    <div style={sCard}>
      <h3 style={sH3}>Colore di accento</h3>
      <p style={sTag}>Il colore dei pulsanti principali, delle evidenze e dei totali, in tutta l'app e nei PDF dei computi. Vale per tutto il team.</p>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 14, flexWrap: 'wrap' }}>
        <input type="color" value={current} onChange={(e) => onUpdate({ coloreAccento: e.target.value.toUpperCase() })} style={{ width: 44, height: 36, padding: 0, border: `1px solid ${C.paleGray}`, borderRadius: 8, background: 'none', cursor: 'pointer' }} />
        <input
          value={hexDraft}
          onChange={(e) => { const v = e.target.value.trim(); setHexDraft(v); if (isHexColor(v)) onUpdate({ coloreAccento: v.toUpperCase() }); }}
          placeholder="#862D2D"
          style={{ ...sInput, width: 110, fontVariantNumeric: 'tabular-nums', textTransform: 'uppercase' }}
        />
        {current.toUpperCase() !== DEFAULT_ACCENT && <button onClick={() => onUpdate({ coloreAccento: DEFAULT_ACCENT })} style={sBtn}>Ripristina bordeaux</button>}
      </div>
      <span style={{ ...sFlabel, marginTop: 14 }}>Scelte rapide</span>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        {ACCENT_PRESETS.map((p) => {
          const on = p.hex.toUpperCase() === current.toUpperCase();
          return (
            <button key={p.hex} onClick={() => onUpdate({ coloreAccento: p.hex })} title={p.name} style={{ ...sBtn, display: 'inline-flex', alignItems: 'center', gap: 8, borderColor: on ? '#2B2321' : C.paleGray, fontWeight: on ? 500 : 400 }}>
              <span style={{ width: 14, height: 14, borderRadius: '50%', background: p.hex, display: 'inline-block' }} />{p.name}
            </button>
          );
        })}
      </div>
      <span style={{ ...sFlabel, marginTop: 14 }}>Anteprima</span>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <span style={{ background: current, color: '#fff', borderRadius: 999, padding: '8px 16px', fontSize: 12.5 }}>+ Nuova voce</span>
        <span style={{ background: current, color: '#fff', borderRadius: 999, padding: '3px 10px', fontSize: 11, fontWeight: 500 }}>In attesa di approvazione</span>
        <span style={{ background: `color-mix(in srgb, ${current} 16%, #fff)`, color: current, borderRadius: 999, padding: '3px 10px', fontSize: 11, fontWeight: 500 }}>Approvato</span>
        <span style={{ color: current, fontWeight: 700, fontSize: 14 }}>Totale generale IVA esclusa 113.443,49 €</span>
      </div>
      {tooLight && <p style={{ fontSize: 12, color: '#8C3220', margin: '10px 0 0' }}>Questo colore è molto chiaro: il testo bianco sui pulsanti potrebbe leggersi male. Meglio un tono più scuro.</p>}
    </div>
  );
}

// Pagina "Impostazioni studio": dati anagrafici usati per l'intestazione/piè di pagina automatici del PDF
// del computo, più la possibilità di caricare un'intestazione e un piè di pagina "personalizzati" (immagini
// pronte, es. carta intestata già impaginata) che sostituiscono del tutto quelli generati dai campi. Tutto
// è salvato in cea.app_state insieme al resto del workspace, quindi condiviso da tutto il team.
function ImpostazioniPage({ settings, onUpdate }) {
  const set = (field) => (e) => onUpdate({ [field]: e.target.value });

  const uploadImage = (field, file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onUpdate({ [field]: reader.result });
    reader.readAsDataURL(file);
  };

  const uploadFont = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onUpdate({ fontPersonalizzato: reader.result, fontPersonalizzatoNome: file.name, usaFontPersonalizzato: true });
    reader.readAsDataURL(file);
  };

  const sCard = { background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 14, padding: '20px 22px', marginBottom: 18 };
  const sH3 = { margin: 0, fontWeight: 300, fontSize: 18, color: C.black, fontFamily: FONT };
  const sTag = { fontSize: 12.5, color: C.gray, margin: '4px 0 0', maxWidth: 640 };
  const sFlabel = { fontSize: 11.5, fontWeight: 600, color: C.midGray, marginBottom: 5, display: 'block' };
  const sInput = { border: `1px solid ${C.paleGray}`, background: C.surfaceSubtle, borderRadius: 8, padding: '9px 12px', fontSize: 13, width: '100%', boxSizing: 'border-box', fontFamily: FONT, color: C.black };
  const sBtn = { border: `1px solid ${C.paleGray}`, background: C.white, color: C.black, fontSize: 11.5, fontWeight: 400, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', fontFamily: FONT };
  const sChk = { display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, fontWeight: 300, color: C.black, marginTop: 10, cursor: 'pointer' };
  const sPreview = { border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '10px 14px', marginTop: 10, background: C.surfaceSubtle, fontSize: 11, color: C.gray };

  const initials = (settings.nome || 'DSQ').split(' ').filter(Boolean).slice(0, 2).map((w) => w[0]).join('').toUpperCase() || 'DSQ';

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 20, flexWrap: 'wrap', marginBottom: 22 }}>
        <h1 style={{ margin: 0, fontFamily: FONT, fontSize: 'clamp(28px, 4vw, 38px)', fontWeight: 200, letterSpacing: '-.03em', color: C.black }}>Impostazioni studio</h1>
        <span style={{ fontSize: 11, fontWeight: 500, padding: '3px 10px', borderRadius: 999, background: C.accentSoft, color: C.maroon }}>Workspace condiviso</span>
      </div>

      <div style={sCard}>
        <h3 style={sH3}>Dati dello studio</h3>
        <p style={sTag}>Questi dati compaiono automaticamente in intestazione e piè di pagina di ogni computo metrico stampato in PDF, a meno di caricare un'intestazione/piè di pagina personalizzata qui sotto.</p>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginTop: 14 }}>
          <div>
            <span style={sFlabel}>Nome studio</span>
            <input value={settings.nome} onChange={set('nome')} placeholder="Desearq Studio" style={sInput} />
          </div>
          <div>
            <span style={sFlabel}>Indirizzo</span>
            <input value={settings.indirizzo} onChange={set('indirizzo')} placeholder="Via, numero civico, città" style={sInput} />
          </div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 14, marginTop: 14 }}>
          <div>
            <span style={sFlabel}>P.IVA</span>
            <input value={settings.piva} onChange={set('piva')} style={sInput} />
          </div>
          <div>
            <span style={sFlabel}>Codice fiscale</span>
            <input value={settings.cf} onChange={set('cf')} placeholder="Codice fiscale" style={sInput} />
          </div>
          <div>
            <span style={sFlabel}>Telefono</span>
            <input value={settings.telefono} onChange={set('telefono')} style={sInput} />
          </div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginTop: 14 }}>
          <div>
            <span style={sFlabel}>Email</span>
            <input value={settings.email} onChange={set('email')} style={sInput} />
          </div>
          <div>
            <span style={sFlabel}>Sito web</span>
            <input value={settings.sito} onChange={set('sito')} style={sInput} />
          </div>
        </div>

        <span style={{ ...sFlabel, marginTop: 14 }}>Logo studio</span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 6 }}>
          {settings.logo ? (
            <img src={settings.logo} alt="Logo studio" style={{ height: 40, borderRadius: 8, border: `1px solid ${C.paleGray}` }} />
          ) : (
            <div style={{ width: 40, height: 40, borderRadius: 8, background: C.maroon, display: 'flex', alignItems: 'center', justifyContent: 'center', color: C.white, fontWeight: 600, fontSize: 11 }}>{initials}</div>
          )}
          <label style={{ ...sBtn, cursor: 'pointer' }}>
            {settings.logo ? 'Sostituisci logo' : 'Carica logo'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('logo', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.logo && <button onClick={() => onUpdate({ logo: null })} style={sBtn}>🗑️ Rimuovi</button>}
        </div>
      </div>

      <AccentColorCard settings={settings} onUpdate={onUpdate} sCard={sCard} sH3={sH3} sTag={sTag} sFlabel={sFlabel} sInput={sInput} sBtn={sBtn} />

      <div style={sCard}>
        <h3 style={sH3}>Intestazione personalizzata</h3>
        <p style={sTag}>Carica un'immagine (es. la tua carta intestata già impaginata) da usare al posto dei dati anagrafici in cima a ogni pagina del PDF, dentro il margine della pagina.</p>
        <label style={sChk}>
          <input type="checkbox" checked={settings.usaIntestazionePersonalizzata} onChange={(e) => onUpdate({ usaIntestazionePersonalizzata: e.target.checked })} />
          Usa l'immagine personalizzata invece dei dati anagrafici
        </label>
        {settings.intestazioneImg && (
          <div style={sPreview}>
            <img src={settings.intestazioneImg} alt="Intestazione personalizzata" style={{ maxHeight: 50, maxWidth: 260, display: 'block', borderRadius: 4 }} />
          </div>
        )}
        <div style={{ display: 'flex', gap: 10, marginTop: 10 }}>
          <label style={{ ...sBtn, cursor: 'pointer' }}>
            {settings.intestazioneImg ? 'Sostituisci immagine' : 'Carica immagine'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('intestazioneImg', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.intestazioneImg && <button onClick={() => onUpdate({ intestazioneImg: null })} style={sBtn}>🗑️ Rimuovi</button>}
        </div>
        {settings.intestazioneImg && (
          <div>
            <span style={{ ...sFlabel, marginTop: 14 }}>Dimensione immagine ({settings.intestazioneScala || 100}% della larghezza disponibile)</span>
            <input type="range" min="25" max="200" step="5" value={settings.intestazioneScala || 100}
              onChange={(e) => onUpdate({ intestazioneScala: Number(e.target.value) })} style={{ width: '100%', maxWidth: 320, display: 'block', margin: '2px 0 0' }} />
          </div>
        )}
      </div>

      <div style={sCard}>
        <h3 style={sH3}>Piè di pagina</h3>
        <p style={sTag}>Un testo breve (es. dati di contatto o un promemoria legale) ripetuto in fondo a ogni pagina, oppure un'immagine personalizzata al posto del testo, anche questa dentro il margine della pagina.</p>

        <span style={{ ...sFlabel, marginTop: 12 }}>Testo piè di pagina</span>
        <input value={settings.testoPiePagina} onChange={set('testoPiePagina')} placeholder="es. Desearq Studio — Via Roma 1, Milano — info@desearq.com" style={sInput} />

        <label style={sChk}>
          <input type="checkbox" checked={settings.usaPiePersonalizzato} onChange={(e) => onUpdate({ usaPiePersonalizzato: e.target.checked })} />
          Usa un'immagine personalizzata invece del testo
        </label>
        {settings.pieImg && (
          <div style={sPreview}>
            <img src={settings.pieImg} alt="Piè di pagina personalizzato" style={{ maxHeight: 40, maxWidth: 260, display: 'block', borderRadius: 4 }} />
          </div>
        )}
        <div style={{ display: 'flex', gap: 10, marginTop: 10 }}>
          <label style={{ ...sBtn, cursor: 'pointer' }}>
            {settings.pieImg ? 'Sostituisci immagine' : 'Carica immagine'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('pieImg', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.pieImg && <button onClick={() => onUpdate({ pieImg: null })} style={sBtn}>🗑️ Rimuovi</button>}
        </div>
        {settings.pieImg && (
          <div>
            <span style={{ ...sFlabel, marginTop: 14 }}>Dimensione immagine ({settings.pieScala || 100}% della larghezza disponibile)</span>
            <input type="range" min="25" max="200" step="5" value={settings.pieScala || 100}
              onChange={(e) => onUpdate({ pieScala: Number(e.target.value) })} style={{ width: '100%', maxWidth: 320, display: 'block', margin: '2px 0 0' }} />
          </div>
        )}
      </div>

      <div style={{ ...sCard, marginBottom: 0 }}>
        <h3 style={sH3}>Font del PDF</h3>
        <p style={sTag}>Scegli il carattere dei computi metrici stampati in PDF, oppure carica il font del tuo studio (.woff2, .woff, .ttf o .otf): se attivo, ha la precedenza sulla scelta qui sotto.</p>
        <span style={{ ...sFlabel, marginTop: 14 }}>Carattere</span>
        <select value={settings.pdfFontPreset || 'inter'} onChange={(e) => onUpdate({ pdfFontPreset: e.target.value })} style={{ ...sInput, maxWidth: 420 }} disabled={!!(settings.usaFontPersonalizzato && settings.fontPersonalizzato)}>
          {Object.entries(PDF_FONT_PRESETS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <div style={{ ...sPreview, fontFamily: settings.usaFontPersonalizzato && settings.fontPersonalizzato ? `'StudioFontPersonalizzato', ${FONT}` : (PDF_FONT_PRESETS[settings.pdfFontPreset] || PDF_FONT_PRESETS.inter).css, fontSize: 13, color: C.black }}>
          {settings.fontPersonalizzato && settings.usaFontPersonalizzato && <style>{`@font-face { font-family: 'StudioFontPersonalizzato'; src: url("${settings.fontPersonalizzato}"); }`}</style>}
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 4 }}>Computo metrico — versione cliente</div>
          Demolizione di pavimenti interni, compreso il trasporto a discarica · 136,84 m² · <span style={{ color: C.maroon, fontWeight: 600 }}>11.289,30 €</span>
        </div>
        <span style={{ ...sFlabel, marginTop: 16 }}>Font personalizzato</span>
        <label style={sChk}>
          <input type="checkbox" checked={settings.usaFontPersonalizzato} disabled={!settings.fontPersonalizzato} onChange={(e) => onUpdate({ usaFontPersonalizzato: e.target.checked })} />
          Usa il font personalizzato nel PDF invece del font di sistema
        </label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 10, flexWrap: 'wrap' }}>
          {settings.fontPersonalizzatoNome && <span style={{ fontSize: 12.5, color: C.darkGray, fontVariantNumeric: 'tabular-nums' }}>{settings.fontPersonalizzatoNome}</span>}
          <label style={{ ...sBtn, cursor: 'pointer' }}>
            {settings.fontPersonalizzato ? 'Sostituisci font' : 'Carica font'}
            <input type="file" accept=".woff,.woff2,.ttf,.otf" onChange={(e) => uploadFont(e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.fontPersonalizzato && <button onClick={() => onUpdate({ fontPersonalizzato: null, fontPersonalizzatoNome: '', usaFontPersonalizzato: false })} style={sBtn}>🗑️ Rimuovi</button>}
        </div>
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

  const handleForgotPassword = async () => {
    if (!email.trim()) { setError('Inserisci prima la tua email, poi clicca su "Password dimenticata?".'); return; }
    setError(''); setInfo(''); setLoading(true);
    const { error: err } = await supabase.auth.resetPasswordForEmail(email.trim());
    setLoading(false);
    if (err) { setError(err.message); return; }
    setInfo('Ti abbiamo inviato un\'email con le istruzioni per reimpostare la password.');
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

  // Questa pagina è portata testualmente dal file CSS/HTML del design (Quant-Login.dc.html): stesse regole
  // CSS (nomi classe compresi), scoperte sotto ".dc-login" per non toccare il resto dell'app. Unica modifica
  // strutturale voluta: ".login" usa min-height:100vh invece di 900px fisso (altrimenti sotto schermi più alti
  // di 900px restava una fascia bianca in fondo). Il testo aggiuntivo su chi diventa admin (".helper") è
  // posizionato in basso in modo assoluto apposta per non alterare l'altezza della card e quindi il suo
  // centraggio verticale — con quel testo dentro al flusso normale la card risultava più alta e quindi
  // visibilmente spostata rispetto al file originale.
  return (
    <div className="dc-login">
      <style>{`
        .dc-login { --paper:#F4EFE8; --ink:#090909; --muted:#77756f; --line:#DED3C4; --bordeaux:#862D2D; --dark:#171411; }
        .dc-login, .dc-login * { box-sizing: border-box; }
        .dc-login { margin:0; background:var(--paper); color:var(--ink); font-family:'Inter',Arial,Helvetica,sans-serif; -webkit-font-smoothing:antialiased; }
        .dc-login a { color:inherit; text-decoration:none; }
        .dc-login .login { min-height:100vh; display:grid; grid-template-columns:1fr 1fr; }
        .dc-login .side { padding:56px; display:flex; flex-direction:column; justify-content:space-between; }
        .dc-login .side.dark { background:var(--dark); color:#f4eee5; }
        .dc-login .side.light { background:var(--paper); align-items:center; justify-content:center; display:flex; position:relative; }
        .dc-login .back { font-size:11px; text-transform:uppercase; letter-spacing:.1em; color:#bdb6ad; display:inline-flex; align-items:center; gap:8px; }
        .dc-login .back:hover { color:#f4eee5; }
        .dc-login .eyebrow { font-size:11px; text-transform:uppercase; letter-spacing:.18em; color:var(--muted); }
        .dc-login .side.dark .eyebrow { color:#aaa39a; }
        .dc-login .copy h1 { font-size:52px; font-weight:400; letter-spacing:-.045em; line-height:.98; margin:22px 0 20px; max-width:480px; }
        .dc-login .copy p { font-size:16px; line-height:1.55; color:#c9c2b8; max-width:420px; margin:0; }
        .dc-login .logo-mark { width:70px; aspect-ratio:1; object-fit:cover; display:block; }
        .dc-login .card { width:100%; max-width:380px; }
        .dc-login .card .eyebrow { margin-bottom:18px; }
        .dc-login .card h2 { font-size:34px; font-weight:400; letter-spacing:-.03em; margin:0 0 14px; color:var(--bordeaux); }
        .dc-login .card>p { font-size:15px; line-height:1.55; color:#5d5952; margin:0 0 38px; max-width:360px; }
        .dc-login form { display:flex; flex-direction:column; gap:24px; }
        .dc-login .field label { display:block; font-size:11px; text-transform:uppercase; letter-spacing:.1em; color:var(--muted); margin-bottom:10px; }
        .dc-login .field input { width:100%; border:0; border-bottom:1px solid var(--line); background:transparent; padding:10px 0; font-size:16px; font-family:inherit; color:var(--ink); }
        .dc-login .field input:focus { outline:0; border-bottom-color:var(--ink); }
        .dc-login .row { display:flex; justify-content:flex-end; margin-top:-12px; }
        .dc-login .row a, .dc-login .row span { font-size:12px; color:var(--muted); cursor:pointer; }
        .dc-login .row a:hover, .dc-login .row span:hover { color:var(--bordeaux); }
        .dc-login .submit { display:flex; align-items:center; justify-content:center; gap:16px; padding:16px; border:1px solid var(--ink); background:var(--ink); color:#fff; font-size:12px; text-transform:uppercase; letter-spacing:.09em; width:100%; cursor:pointer; margin-top:4px; }
        .dc-login .submit:hover { transform:translateY(-1px); }
        .dc-login .submit:disabled { opacity:.7; cursor:default; transform:none; }
        .dc-login .arrow { display:inline-grid; place-items:center; width:20px; height:20px; border:1px solid currentColor; border-radius:50%; font-size:12px; }
        .dc-login .altline { display:flex; align-items:center; gap:14px; margin:32px 0; color:var(--muted); font-size:11px; letter-spacing:.08em; text-transform:uppercase; }
        .dc-login .altline:before, .dc-login .altline:after { content:""; flex:1; height:1px; background:var(--line); }
        .dc-login .request { font-size:14px; color:#5d5952; text-align:center; margin:0; }
        .dc-login .request span { color:var(--bordeaux); border-bottom:1px solid currentColor; cursor:pointer; }
        .dc-login .foot { margin-top:40px; font-size:11px; color:var(--muted); text-align:center; text-transform:uppercase; letter-spacing:.08em; }
        .dc-login .helper { position:absolute; left:32px; right:32px; bottom:20px; font-size:11px; color:var(--muted); text-align:center; line-height:1.5; }
        .dc-login .msg-error { font-size:12px; color:var(--bordeaux); margin:0 0 -8px; }
        .dc-login .msg-info { font-size:12px; color:#2e7d4f; margin:0 0 -8px; }
        @media (max-width: 860px) {
          .dc-login .login { grid-template-columns: 1fr; }
          .dc-login .side.dark { display: none !important; }
        }
      `}</style>
      <div className="login">
        <div className="side dark">
          <a className="back" href="https://desearq.com" target="_blank" rel="noopener noreferrer">← Desearq Lab</a>
          <div className="copy">
            <div className="eyebrow">DESEARQ LAB&nbsp; — QUANT</div>
            <h1>Benvenuto su Quant.</h1>
            <p>Computi metrici più ordinati, il tuo listino sempre sotto controllo.</p>
          </div>
          <img className="logo-mark" src="/quant-logo.png" alt="Quant" />
        </div>

        <div className="side light">
          <div className="card">
            <div className="eyebrow">{mode === 'signin' ? 'Accesso' : 'Registrazione'}</div>
            <h2>{mode === 'signin' ? 'Accedi al tuo account' : 'Crea il tuo account'}</h2>
            <p>{mode === 'signin' ? 'Inserisci le tue credenziali per entrare in Quant.' : 'Inserisci i tuoi dati per creare il tuo account.'}</p>
            <form onSubmit={(e) => e.preventDefault()}>
              {mode === 'signup' && (
                <div className="field">
                  <label htmlFor="name">Nome</label>
                  <input id="name" value={name} onChange={(e) => setName(e.target.value)} placeholder="Il tuo nome" autoComplete="name" />
                </div>
              )}
              <div className="field">
                <label htmlFor="email">Email</label>
                <input id="email" type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="nome@studio.it" autoComplete="email" />
              </div>
              <div className="field">
                <label htmlFor="password">Password</label>
                <input
                  id="password"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && (mode === 'signin' ? handleSignIn() : handleSignUp())}
                  placeholder="••••••••"
                  autoComplete="current-password"
                />
              </div>
              {mode === 'signin' && (
                <div className="row"><span onClick={handleForgotPassword}>Password dimenticata?</span></div>
              )}
              {error && <p className="msg-error">{error}</p>}
              {info && <p className="msg-info">{info}</p>}
              <button type="button" className="submit" disabled={loading} onClick={mode === 'signin' ? handleSignIn : handleSignUp}>
                {loading ? 'Un attimo…' : mode === 'signin' ? 'Accedi' : 'Crea account'}
                {!loading && <span className="arrow">→</span>}
              </button>
            </form>
            <div className="altline">oppure</div>
            <p className="request">
              {mode === 'signin' ? (
                <>Non hai ancora un account? <span onClick={() => { setMode('signup'); setError(''); setInfo(''); }}>Richiedi l'accesso</span></>
              ) : (
                <>Hai già un account? <span onClick={() => { setMode('signin'); setError(''); setInfo(''); }}>Accedi</span></>
              )}
            </p>
            <div className="foot">© 2026 Desearq Lab</div>
          </div>
          <p className="helper">
            Se sei il primo ad accedere diventi automaticamente admin. Chi arriva dopo riceve l'accesso già pronto (email e password) dall'admin nella sezione Team: userà direttamente "Accedi" qui sopra, senza bisogno di registrarsi.
          </p>
        </div>
      </div>
    </div>
  );
}

// ===================== Offerte imprese sul computo =====================
// Lo studio genera, per una revisione del computo, un link (+ PIN) da mandare a un'impresa. L'impresa apre il
// computo senza prezzi, scrive il proprio prezzo unitario voce per voce e lo invia allo studio. L'offerta resta
// "in attesa" (tabella cea.computo_offerte) finché lo studio non la apre: accettandola si crea una nuova
// revisione del computo, intitolata con il nome dell'impresa, con i suoi prezzi come prezzi impresa.

const offertaUrl = (token) => `${window.location.origin}${window.location.pathname}?offerta=${token}`;

// Fotografia del computo inviata all'impresa: solo struttura, voci, quantità e note visibili al cliente —
// nessun prezzo dello studio.
function buildOffertaSnapshot(project, revision) {
  const groups = buildComputoGroups(revision);
  return {
    noteInizio: revision.noteInizio || '',
    noteFine: revision.noteFine || '',
    ubicazione: project.header?.ubicazione || '',
    sections: groups.map((g) => ({
      name: g.name,
      categorie: g.categorie.filter((c) => c.items.length).map((c) => ({
        name: c.name,
        items: c.items.map((it) => ({ id: String(it.id), code: it.code || '', desc: it.desc || '', qty: it.qty || '0', unit: it.unit || '', note: it.noteCliente || '' })),
      })),
    })).filter((s) => s.categorie.length),
  };
}

function OfferteImpreseModal({ project, revision, onClose, onAcceptOffer }) {
  const [tab, setTab] = useState('offerte'); // 'offerte' | 'link'
  const [links, setLinks] = useState(null);
  const [offerte, setOfferte] = useState(null);
  const [userId, setUserId] = useState(null);
  const [newImpresa, setNewImpresa] = useState('');
  const [justCreated, setJustCreated] = useState(null); // { url, pin, impresa }
  const [opened, setOpened] = useState(null); // offerta aperta nel dettaglio
  const [busy, setBusy] = useState(false);
  const projectId = String(project.id);

  const reload = async () => {
    const { data: l } = await cea.from('computo_offerta_links').select('*').eq('project_id', projectId).order('created_at', { ascending: false });
    setLinks(l || []);
    const ids = (l || []).map((x) => x.id);
    if (!ids.length) { setOfferte([]); return; }
    const { data: o } = await cea.from('computo_offerte').select('*').in('link_id', ids).order('submitted_at', { ascending: false });
    setOfferte(o || []);
  };
  React.useEffect(() => {
    reload();
    supabase.auth.getUser().then(({ data }) => setUserId(data.user?.id || null));
  }, []);

  const linkOf = (id) => (links || []).find((l) => l.id === id);
  const shareText = (impresa, url, pin) => `Buongiorno,\nle inviamo il computo metrico del progetto "${project.name}" da quotare.\nApra il link, inserisca il PIN e compili i prezzi unitari, poi clicchi "Invia allo studio".\n\nLink: ${url}\nPIN: ${pin}\n\nGrazie`;
  const copy = (text) => navigator.clipboard?.writeText(text).then(() => alert('Copiato negli appunti.')).catch(() => window.prompt('Copia il testo:', text));

  const createLink = async () => {
    const impresa = newImpresa.trim();
    if (!impresa) { alert("Scrivi il nome dell'impresa."); return; }
    if (!revision) { alert('Apri prima una revisione del computo.'); return; }
    setBusy(true);
    const token = crypto.randomUUID ? crypto.randomUUID().replace(/-/g, '') : `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`;
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    const { error } = await cea.from('computo_offerta_links').insert({
      token, pin, project_id: projectId, project_name: project.name, revision_id: String(revision.id),
      revision_label: revision.customName || revision.label, impresa_nome: impresa, snapshot: buildOffertaSnapshot(project, revision),
    });
    setBusy(false);
    if (error) { alert('Creazione del link non riuscita: ' + error.message); return; }
    setJustCreated({ url: offertaUrl(token), pin, impresa });
    setNewImpresa('');
    reload();
  };

  const toggleLink = async (link) => {
    if (link.active && !confirm(`Disattivare il link di "${link.impresa_nome}"? L'impresa non potrà più aprirlo né inviare offerte.`)) return;
    await cea.from('computo_offerta_links').update({ active: !link.active }).eq('id', link.id);
    reload();
  };

  const reject = async (o) => {
    if (!confirm(`Rifiutare l'offerta di "${o.impresa_nome}"?`)) return;
    setBusy(true);
    const { error } = await cea.from('computo_offerte').update({ status: 'rejected', reviewed_at: new Date().toISOString(), reviewed_by: userId }).eq('id', o.id);
    setBusy(false);
    if (error) { alert('Operazione non riuscita: ' + error.message); return; }
    setOpened(null);
    reload();
  };

  const accept = async (o, clienteMode, ricarico) => {
    const link = linkOf(o.link_id);
    setBusy(true);
    const revId = onAcceptOffer(o, link, clienteMode, ricarico);
    if (!revId) { setBusy(false); return; }
    const { error } = await cea.from('computo_offerte').update({ status: 'accepted', reviewed_at: new Date().toISOString(), reviewed_by: userId, revision_created: String(revId) }).eq('id', o.id);
    setBusy(false);
    if (error) alert("La revisione è stata creata, ma non sono riuscito ad aggiornare lo stato dell'offerta: " + error.message);
    setOpened(null);
    reload();
  };

  const pendingCount = (offerte || []).filter((o) => o.status === 'pending').length;
  const statusChip = (st) => {
    const map = { pending: ['In attesa', '#5A4E47', '#F1ECE6'], accepted: ['Accettata', '#2F5D45', '#E7F0EA'], rejected: ['Rifiutata', '#862D2D', '#F4EAEA'] };
    const [label, color, bg] = map[st] || [st, C.gray, '#F1ECE6'];
    return <span style={{ fontSize: 11, fontWeight: 600, color, background: bg, borderRadius: 999, padding: '3px 9px', whiteSpace: 'nowrap' }}>{label}</span>;
  };
  const tabBtn = (key, label) => (
    <button onClick={() => setTab(key)} style={{ border: 'none', background: tab === key ? C.black : 'transparent', color: tab === key ? C.white : C.darkGray, borderRadius: 999, padding: '7px 14px', fontSize: 12.5, fontWeight: 500, cursor: 'pointer', fontFamily: FONT }}>{label}</button>
  );
  const input = { width: '100%', boxSizing: 'border-box', border: `1px solid ${C.paleGray}`, borderRadius: 10, padding: '9px 11px', fontSize: 13, fontFamily: FONT };

  if (opened) {
    return <OffertaDettaglioModal offerta={opened} link={linkOf(opened.link_id)} project={project} busy={busy} onBack={() => setOpened(null)} onReject={() => reject(opened)} onAccept={(mode, ric) => accept(opened, mode, ric)} />;
  }

  return createPortal(
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 16, padding: 22, width: 720, maxWidth: '100%', maxHeight: '90vh', overflowY: 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12, marginBottom: 14 }}>
          <div>
            <h2 style={{ fontFamily: FONT, fontSize: 19, fontWeight: 600, margin: 0, color: C.black }}>Offerte imprese</h2>
            <p style={{ fontSize: 12, color: C.gray, margin: '4px 0 0' }}>Manda il computo a un'impresa: compila i prezzi online e li invia allo studio.</p>
          </div>
          <button onClick={onClose} style={{ border: 'none', background: 'transparent', fontSize: 20, color: C.gray, cursor: 'pointer', lineHeight: 1 }}>×</button>
        </div>
        <div style={{ display: 'inline-flex', gap: 4, background: '#F4F0EA', borderRadius: 999, padding: 4, marginBottom: 16 }}>
          {tabBtn('offerte', `Offerte ricevute${pendingCount ? ` (${pendingCount})` : ''}`)}
          {tabBtn('link', 'Invia a un\'impresa')}
        </div>

        {tab === 'offerte' && (
          offerte === null ? <p style={{ fontSize: 13, color: C.gray }}>Caricamento…</p>
          : offerte.length === 0 ? (
            <div style={{ border: `1px dashed ${C.paleGray}`, borderRadius: 12, padding: 18, fontSize: 13, color: C.gray, textAlign: 'center' }}>
              Nessuna offerta ricevuta per questo progetto. Vai su <b style={{ color: C.black, cursor: 'pointer' }} onClick={() => setTab('link')}>Invia a un'impresa</b> per creare il link.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {offerte.map((o) => {
                const link = linkOf(o.link_id);
                return (
                  <div key={o.id} style={{ display: 'flex', alignItems: 'center', gap: 12, border: `1px solid ${o.status === 'pending' ? C.maroon : C.paleGray}`, borderRadius: 12, padding: '11px 14px', flexWrap: 'wrap' }}>
                    <div style={{ flex: '1 1 220px', minWidth: 0 }}>
                      <div style={{ fontSize: 14, fontWeight: 600, color: C.black }}>{o.impresa_nome || link?.impresa_nome}</div>
                      <div style={{ fontSize: 11.5, color: C.gray }}>
                        Inviata il {new Date(o.submitted_at).toLocaleString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                        {link?.revision_label ? ` · su ${link.revision_label}` : ''}{o.referente ? ` · ${o.referente}` : ''}
                      </div>
                    </div>
                    <div style={{ fontSize: 14, fontWeight: 600, color: C.black, whiteSpace: 'nowrap' }}>{formatEuro(Number(o.totale) || 0)}</div>
                    {statusChip(o.status)}
                    <button onClick={() => setOpened(o)} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '6px 13px', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}>{o.status === 'pending' ? 'Apri e valuta' : 'Vedi'}</button>
                  </div>
                );
              })}
            </div>
          )
        )}

        {tab === 'link' && (
          <>
            <div style={{ background: '#FAF8F4', border: `1px solid ${C.paleGray}`, borderRadius: 12, padding: 14, marginBottom: 16 }}>
              <label style={{ display: 'block', fontSize: 11, fontWeight: 600, color: C.black, marginBottom: 5 }}>Nome dell'impresa</label>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <input value={newImpresa} onChange={(e) => setNewImpresa(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && createLink()} placeholder="Es. Edil Rossi S.r.l." style={{ ...input, flex: '1 1 240px', width: 'auto' }} />
                <button disabled={busy} onClick={createLink} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '9px 16px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>Crea link</button>
              </div>
              <p style={{ fontSize: 11.5, color: C.gray, margin: '8px 0 0' }}>
                L'impresa vedrà la revisione aperta ora (<b>{revision?.customName || revision?.label}</b>) con voci, quantità e note, ma nessun prezzo dello studio e nessun dato del cliente.
              </p>
            </div>
            {justCreated && (
              <div style={{ border: `1px solid ${C.maroon}`, borderRadius: 12, padding: 14, marginBottom: 16 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.black, marginBottom: 6 }}>Link per {justCreated.impresa} pronto</div>
                <div style={{ fontSize: 12, color: C.darkGray, wordBreak: 'break-all', marginBottom: 4 }}>{justCreated.url}</div>
                <div style={{ fontSize: 13, color: C.black, marginBottom: 10 }}>PIN: <b style={{ letterSpacing: '0.15em' }}>{justCreated.pin}</b></div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <button onClick={() => copy(shareText(justCreated.impresa, justCreated.url, justCreated.pin))} style={{ border: 'none', background: C.black, color: C.white, borderRadius: 999, padding: '7px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Copia messaggio con link e PIN</button>
                  <button onClick={() => copy(justCreated.url)} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '7px 14px', fontSize: 12, cursor: 'pointer' }}>Copia solo link</button>
                </div>
              </div>
            )}
            <p style={{ fontSize: 12, fontWeight: 600, color: C.black, margin: '0 0 8px' }}>Link creati per questo progetto</p>
            {links === null ? <p style={{ fontSize: 13, color: C.gray }}>Caricamento…</p> : links.length === 0 ? <p style={{ fontSize: 12.5, color: C.gray }}>Nessun link ancora.</p> : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {links.map((l) => (
                  <div key={l.id} style={{ display: 'flex', alignItems: 'center', gap: 10, borderBottom: `1px solid ${C.paleGray}`, padding: '8px 0', flexWrap: 'wrap', opacity: l.active ? 1 : 0.55 }}>
                    <div style={{ flex: '1 1 200px', minWidth: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.black }}>{l.impresa_nome}</div>
                      <div style={{ fontSize: 11.5, color: C.gray }}>{l.revision_label} · PIN {l.pin} · {l.active ? 'attivo' : 'disattivato'}</div>
                    </div>
                    <button onClick={() => copy(shareText(l.impresa_nome, offertaUrl(l.token), l.pin))} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '5px 11px', fontSize: 11.5, cursor: 'pointer' }}>Copia link e PIN</button>
                    <button onClick={() => toggleLink(l)} style={{ border: `1px solid ${C.paleGray}`, background: C.white, borderRadius: 999, padding: '5px 11px', fontSize: 11.5, cursor: 'pointer', color: l.active ? C.maroon : C.black }}>{l.active ? 'Disattiva' : 'Riattiva'}</button>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </div>,
    document.body
  );
}

// Dettaglio di un'offerta: confronto voce per voce con il prezzo impresa della revisione di partenza, poi
// accetta (crea la revisione con il nome dell'impresa) o rifiuta.
function OffertaDettaglioModal({ offerta, link, project, busy, onBack, onAccept, onReject }) {
  const [clienteMode, setClienteMode] = useState('mantieni'); // 'mantieni' | 'ricarico'
  const [ricarico, setRicarico] = useState('10');
  const snapshot = link?.snapshot || { sections: [] };
  const prezzi = new Map((offerta.prezzi || []).map((p) => [String(p.id), p]));
  const sourceRev = (project.revisions || []).find((r) => String(r.id) === String(link?.revision_id));
  const sourceItems = new Map((sourceRev?.items || []).map((it) => [String(it.id), it]));
  const missing = snapshot.sections.reduce((n, s) => n + s.categorie.reduce((m, c) => m + c.items.filter((it) => !prezzi.has(it.id)).length, 0), 0);
  let totOfferta = 0;
  let totStudio = 0;
  const th = { textAlign: 'left', fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.4, color: C.gray, fontWeight: 600, padding: '6px 6px', borderBottom: `1.5px solid ${C.black}` };
  const td = { padding: '7px 6px', fontSize: 12, borderBottom: `1px solid ${C.paleGray}`, verticalAlign: 'top' };
  const num = { ...td, textAlign: 'right', whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums' };
  return createPortal(
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 16, padding: 22, width: 980, maxWidth: '100%', maxHeight: '92vh', display: 'flex', flexDirection: 'column' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12, marginBottom: 12 }}>
          <div>
            <button onClick={onBack} style={{ border: 'none', background: 'transparent', color: C.gray, fontSize: 12, cursor: 'pointer', padding: 0, marginBottom: 4 }}>← Offerte</button>
            <h2 style={{ fontFamily: FONT, fontSize: 19, fontWeight: 600, margin: 0, color: C.black }}>Offerta di {offerta.impresa_nome}</h2>
            <p style={{ fontSize: 12, color: C.gray, margin: '4px 0 0' }}>
              Su {link?.revision_label || 'revisione eliminata'} · {[offerta.referente, offerta.email, offerta.telefono].filter(Boolean).join(' · ') || 'nessun referente indicato'}
            </p>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 10.5, color: C.gray, textTransform: 'uppercase', letterSpacing: 0.5 }}>Totale offerta IVA esclusa</div>
            <div style={{ fontSize: 22, fontWeight: 600, color: C.maroon }}>{formatEuro(Number(offerta.totale) || 0)}</div>
          </div>
        </div>
        {offerta.note && <div style={{ background: '#FAF8F4', borderLeft: `3px solid ${C.maroon}`, padding: '8px 12px', fontSize: 12.5, color: C.darkGray, whiteSpace: 'pre-wrap', marginBottom: 10 }}>{offerta.note}</div>}
        {missing > 0 && <p style={{ fontSize: 12, color: C.maroon, margin: '0 0 8px' }}>{missing} voc{missing === 1 ? 'e' : 'i'} senza prezzo: nella nuova revisione manterranno il prezzo impresa attuale.</p>}
        <div style={{ overflowY: 'auto', flex: 1, minHeight: 0, border: `1px solid ${C.paleGray}`, borderRadius: 12 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed' }}>
            <colgroup><col style={{ width: 90 }} /><col /><col style={{ width: 80 }} /><col style={{ width: 105 }} /><col style={{ width: 110 }} /><col style={{ width: 110 }} /><col style={{ width: 70 }} /></colgroup>
            <thead style={{ position: 'sticky', top: 0, background: C.white }}>
              <tr><th style={th}>Codice</th><th style={th}>Descrizione</th><th style={{ ...th, textAlign: 'right' }}>Q.tà</th><th style={{ ...th, textAlign: 'right' }}>Prezzo offerto</th><th style={{ ...th, textAlign: 'right' }}>Importo offerto</th><th style={{ ...th, textAlign: 'right' }}>Importo attuale</th><th style={{ ...th, textAlign: 'right' }}>Diff.</th></tr>
            </thead>
            <tbody>
              {snapshot.sections.map((s, si) => (
                <React.Fragment key={si}>
                  <tr><td colSpan={7} style={{ ...td, background: '#F4EAEA', color: C.maroon, fontWeight: 700, textTransform: 'uppercase', fontSize: 11.5 }}>{s.name}</td></tr>
                  {s.categorie.map((c, ci) => (
                    <React.Fragment key={ci}>
                      <tr><td colSpan={7} style={{ ...td, background: '#F6F3EE', fontWeight: 600 }}>{c.name}</td></tr>
                      {c.items.map((it) => {
                        const p = prezzi.get(it.id);
                        const qty = parseEuro(it.qty);
                        const offered = p ? Number(p.price) || 0 : null;
                        const src = sourceItems.get(it.id);
                        const studioImp = src ? parseEuro(src.unitPriceImpresa) * parseEuro(src.qty) : null;
                        const offImp = offered !== null ? offered * qty : null;
                        if (offImp !== null) totOfferta += offImp;
                        if (studioImp !== null && offImp !== null) totStudio += studioImp; // confronto solo sulle voci quotate
                        const diff = offImp !== null && studioImp ? ((offImp - studioImp) / studioImp) * 100 : null;
                        return (
                          <tr key={it.id}>
                            <td style={{ ...td, color: C.gray, fontSize: 11 }}>{it.code}</td>
                            <td style={td}>
                              {it.desc}
                              {p?.note && <div style={{ fontSize: 11, color: C.maroon, fontStyle: 'italic', marginTop: 3 }}>Nota impresa: {p.note}</div>}
                            </td>
                            <td style={num}>{it.qty} {it.unit}</td>
                            <td style={{ ...num, fontWeight: 600 }}>{offered !== null ? formatEuro(offered) : '—'}</td>
                            <td style={{ ...num, fontWeight: 600, color: C.maroon }}>{offImp !== null ? formatEuro(offImp) : '—'}</td>
                            <td style={{ ...num, color: C.gray }}>{studioImp !== null ? formatEuro(studioImp) : '—'}</td>
                            <td style={{ ...num, color: diff === null ? C.gray : diff > 0 ? C.maroon : '#2F5D45' }}>{diff === null ? '—' : `${diff > 0 ? '+' : ''}${itNum(diff, 1)}%`}</td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  ))}
                </React.Fragment>
              ))}
              <tr>
                <td colSpan={4} style={{ ...td, textAlign: 'right', fontWeight: 700 }}>Totale IVA esclusa{missing > 0 ? ' (solo voci quotate)' : ''}</td>
                <td style={{ ...num, fontWeight: 700, color: C.maroon }}>{formatEuro(totOfferta)}</td>
                <td style={{ ...num, fontWeight: 700, color: C.gray }}>{formatEuro(totStudio)}</td>
                <td style={num}>{totStudio ? `${totOfferta > totStudio ? '+' : ''}${itNum(((totOfferta - totStudio) / totStudio) * 100, 1)}%` : ''}</td>
              </tr>
            </tbody>
          </table>
        </div>

        {offerta.status === 'pending' ? (
          <div style={{ marginTop: 14, display: 'flex', gap: 14, alignItems: 'flex-end', justifyContent: 'space-between', flexWrap: 'wrap' }}>
            <div style={{ fontSize: 12.5, color: C.darkGray }}>
              <div style={{ fontWeight: 600, color: C.black, marginBottom: 6 }}>Prezzi cliente nella nuova revisione</div>
              <label style={{ display: 'flex', gap: 6, alignItems: 'center', marginBottom: 4, cursor: 'pointer' }}>
                <input type="radio" checked={clienteMode === 'mantieni'} onChange={() => setClienteMode('mantieni')} style={{ accentColor: C.maroon }} /> Mantieni i prezzi cliente attuali
              </label>
              <label style={{ display: 'flex', gap: 6, alignItems: 'center', cursor: 'pointer' }}>
                <input type="radio" checked={clienteMode === 'ricarico'} onChange={() => setClienteMode('ricarico')} style={{ accentColor: C.maroon }} /> Prezzo offerto +
                <input value={ricarico} onChange={(e) => { setRicarico(e.target.value); setClienteMode('ricarico'); }} style={{ width: 52, border: `1px solid ${C.paleGray}`, borderRadius: 8, padding: '4px 6px', fontSize: 12.5, textAlign: 'right' }} /> % di ricarico
              </label>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button disabled={busy} onClick={onReject} style={{ border: `1px solid ${C.paleGray}`, background: C.white, color: C.maroon, borderRadius: 999, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Rifiuta</button>
              <button disabled={busy} onClick={() => onAccept(clienteMode, parseEuro(ricarico))} style={{ border: 'none', background: C.maroon, color: C.white, borderRadius: 999, padding: '10px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>{busy ? 'Un attimo…' : `Accetta e crea revisione "${offerta.impresa_nome}"`}</button>
            </div>
          </div>
        ) : (
          <p style={{ fontSize: 12.5, color: C.gray, margin: '12px 0 0' }}>
            Offerta {offerta.status === 'accepted' ? 'accettata: è stata creata la revisione con il nome dell\'impresa.' : 'rifiutata.'}
          </p>
        )}
      </div>
    </div>,
    document.body
  );
}

// Vista pubblica per l'impresa (link + PIN, nessun account): computo senza prezzi da compilare e inviare.
function OffertaImpresaView({ token }) {
  const [pin, setPin] = useState('');
  const [checking, setChecking] = useState(false);
  const [error, setError] = useState('');
  const [data, setData] = useState(null);
  const [prices, setPrices] = useState({}); // id -> stringa digitata
  const [notes, setNotes] = useState({}); // id -> nota
  const [openNote, setOpenNote] = useState({});
  const [referente, setReferente] = useState('');
  const [email, setEmail] = useState('');
  const [telefono, setTelefono] = useState('');
  const [notaGenerale, setNotaGenerale] = useState('');
  const [sending, setSending] = useState(false);
  const [sentAt, setSentAt] = useState(null);

  const unlock = async () => {
    if (!pin.trim() || checking) return;
    setChecking(true); setError('');
    const { data: result, error: err } = await cea.rpc('offerta_get', { p_token: token, p_pin: pin.trim() });
    setChecking(false);
    if (err || (result && result.errore)) { setError(result && result.errore ? result.errore : 'Link o PIN non validi. Controlla con lo studio.'); return; }
    setData(result);
    if (result.pending) {
      const pr = {}; const nt = {};
      (result.pending.prezzi || []).forEach((p) => { pr[String(p.id)] = itNum(Number(p.price) || 0, 2); if (p.note) nt[String(p.id)] = p.note; });
      setPrices(pr); setNotes(nt);
      setReferente(result.pending.referente || ''); setEmail(result.pending.email || ''); setTelefono(result.pending.telefono || '');
      setNotaGenerale(result.pending.note || '');
    }
  };

  const sections = data?.snapshot?.sections || [];
  const allItems = sections.flatMap((s) => s.categorie.flatMap((c) => c.items));
  const priceOf = (id) => (String(prices[id] ?? '').trim() === '' ? null : parseEuro(prices[id]));
  const sectionTotal = (s) => s.categorie.reduce((n, c) => n + c.items.reduce((m, it) => m + (priceOf(it.id) || 0) * parseEuro(it.qty), 0), 0);
  const total = sections.reduce((n, s) => n + sectionTotal(s), 0);
  const filled = allItems.filter((it) => priceOf(it.id) !== null).length;

  const submit = async () => {
    if (!filled) { alert('Inserisci almeno un prezzo prima di inviare.'); return; }
    if (filled < allItems.length && !confirm(`Hai compilato ${filled} voci su ${allItems.length}. Inviare comunque l'offerta?`)) return;
    if (!referente.trim() && !confirm('Non hai indicato un referente. Inviare comunque?')) return;
    const payload = allItems.filter((it) => priceOf(it.id) !== null).map((it) => ({ id: it.id, price: priceOf(it.id), note: (notes[it.id] || '').trim() }));
    setSending(true);
    const { data: esito, error: err } = await cea.rpc('offerta_submit', {
      p_token: token, p_pin: pin.trim(), p_prezzi: payload, p_note: notaGenerale.trim(), p_referente: referente.trim(),
      p_email: email.trim(), p_telefono: telefono.trim(), p_totale: Math.round(total * 100) / 100,
    });
    setSending(false);
    if (err || (esito && esito.errore)) { alert('Invio non riuscito: ' + (err ? err.message : esito.errore)); return; }
    setSentAt(nowLabel());
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const field = { width: '100%', boxSizing: 'border-box', border: '1px solid #E3D9CC', borderRadius: 10, padding: '9px 11px', fontSize: 13.5, fontFamily: FONT, background: '#fff', color: '#2B2321' };

  if (!data) {
    return (
      <div style={{ minHeight: '100vh', background: '#F4EFE8', display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT, padding: 16 }}>
        <div style={{ background: '#fff', borderRadius: 20, boxShadow: '0 8px 24px rgba(33,26,22,0.08)', padding: 32, width: 380, maxWidth: '100%' }}>
          <img src="/quant-logo-full.png" alt="Quant" style={{ height: 54, marginBottom: 18 }} />
          <h1 style={{ fontSize: 21, fontWeight: 500, color: '#2B2321', margin: '0 0 6px' }}>Computo da quotare</h1>
          <p style={{ fontSize: 13, color: '#8B8178', margin: '0 0 22px', lineHeight: 1.5 }}>Inserisci il PIN che ti ha mandato lo studio per aprire il computo metrico e indicare i tuoi prezzi.</p>
          <label style={{ fontSize: 11, fontWeight: 600, color: '#4A403C' }}>PIN</label>
          <input value={pin} onChange={(e) => setPin(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && unlock()} inputMode="numeric" placeholder="••••••" style={{ ...field, fontSize: 17, letterSpacing: '0.25em', margin: '5px 0 14px' }} />
          {error && <p style={{ fontSize: 12.5, color: '#862D2D', margin: '0 0 10px' }}>{error}</p>}
          <button disabled={checking} onClick={unlock} style={{ width: '100%', background: '#862D2D', color: '#fff', border: 'none', borderRadius: 999, padding: '12px 0', fontSize: 14, fontWeight: 500, cursor: 'pointer', fontFamily: FONT }}>{checking ? 'Un attimo…' : 'Apri il computo'}</button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', background: '#F4EFE8', fontFamily: FONT, color: '#2B2321' }}>
      <header style={{ background: '#2B2321', color: '#fff', padding: '14px 20px' }}>
        <div style={{ maxWidth: 1060, margin: '0 auto', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 500 }}>{data.projectName}</div>
            <div style={{ fontSize: 12, color: '#C9BFB4' }}>{data.studioNome || 'Studio'} · richiesta di offerta per <b style={{ color: '#fff' }}>{data.impresaNome}</b></div>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 10.5, color: '#C9BFB4', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Totale offerta IVA esclusa</div>
            <div style={{ fontSize: 20, fontWeight: 600 }}>{formatEuro(total)}</div>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1060, margin: '0 auto', padding: '20px 16px 120px' }}>
        {sentAt && (
          <div style={{ background: '#E7F0EA', border: '1px solid #B9D3C2', color: '#2F5D45', borderRadius: 14, padding: '14px 16px', marginBottom: 16, fontSize: 13.5 }}>
            <b>Offerta inviata allo studio</b> ({sentAt}). Puoi ancora correggerla e inviarla di nuovo finché lo studio non l'ha valutata.
          </div>
        )}
        {!sentAt && data.pending && (
          <div style={{ background: '#fff', border: '1px solid #E3D9CC', borderRadius: 14, padding: '12px 16px', marginBottom: 16, fontSize: 13, color: '#4A403C' }}>
            Hai già inviato un'offerta: trovi i tuoi prezzi qui sotto. Se la modifichi e la invii di nuovo, sostituisce quella precedente.
          </div>
        )}
        <div style={{ background: '#fff', border: '1px solid #E3D9CC', borderRadius: 14, padding: '14px 16px', marginBottom: 16, fontSize: 13, color: '#4A403C', lineHeight: 1.55 }}>
          Per ogni voce indica il tuo <b>prezzo unitario</b> (IVA esclusa): l'importo si calcola da solo. Puoi aggiungere una nota alla singola voce. Quando hai finito clicca <b>Invia allo studio</b> in basso.
          {data.snapshot?.ubicazione && <div style={{ marginTop: 6, color: '#8B8178' }}>Ubicazione: {data.snapshot.ubicazione}</div>}
        </div>
        {(data.snapshot?.noteInizio || '').trim() && (
          <div style={{ background: '#FAF8F4', borderLeft: '3px solid #862D2D', borderRadius: 6, padding: '10px 14px', marginBottom: 16, fontSize: 13, whiteSpace: 'pre-wrap', lineHeight: 1.5 }}>{data.snapshot.noteInizio}</div>
        )}

        {sections.map((s, si) => (
          <section key={si} style={{ background: '#fff', border: '1px solid #E3D9CC', borderRadius: 16, marginBottom: 16, overflow: 'hidden' }}>
            <div style={{ background: '#2B2321', color: '#fff', padding: '11px 16px', display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
              <span style={{ fontSize: 13, letterSpacing: '0.06em', textTransform: 'uppercase' }}>{s.name}</span>
              <span style={{ fontSize: 13, fontWeight: 600, whiteSpace: 'nowrap' }}>{formatEuro(sectionTotal(s))}</span>
            </div>
            {s.categorie.map((c, ci) => (
              <div key={ci}>
                <div style={{ background: '#F6F3EE', padding: '8px 16px', fontSize: 12.5, fontWeight: 600, borderBottom: '1px solid #EDE6DB' }}>{c.name}</div>
                {c.items.map((it) => {
                  const p = priceOf(it.id);
                  const importo = p !== null ? p * parseEuro(it.qty) : null;
                  return (
                    <div key={it.id} style={{ display: 'flex', flexWrap: 'wrap', gap: '8px 14px', alignItems: 'flex-start', padding: '11px 16px', borderBottom: '1px solid #F0EAE1' }}>
                      <div style={{ flex: '1 1 340px', minWidth: 220 }}>
                        <div style={{ fontSize: 11, color: '#8B8178', fontFamily: 'ui-monospace, monospace' }}>{it.code}</div>
                        <div style={{ fontSize: 13, lineHeight: 1.45 }}>{it.desc}</div>
                        {it.note && <div style={{ fontSize: 11.5, color: '#8B8178', fontStyle: 'italic', marginTop: 3 }}>{it.note}</div>}
                        {openNote[it.id] || notes[it.id] ? (
                          <input value={notes[it.id] || ''} onChange={(e) => setNotes((n) => ({ ...n, [it.id]: e.target.value }))} placeholder="Nota per lo studio su questa voce" style={{ ...field, fontSize: 12, padding: '6px 9px', marginTop: 6 }} />
                        ) : (
                          <button onClick={() => setOpenNote((o) => ({ ...o, [it.id]: true }))} style={{ border: 'none', background: 'transparent', color: '#862D2D', fontSize: 11.5, padding: 0, marginTop: 5, cursor: 'pointer' }}>+ aggiungi nota</button>
                        )}
                      </div>
                      <div style={{ width: 96, textAlign: 'right' }}>
                        <div style={{ fontSize: 10, color: '#8B8178', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Quantità</div>
                        <div style={{ fontSize: 13.5, fontVariantNumeric: 'tabular-nums', whiteSpace: 'nowrap' }}>{it.qty} {it.unit}</div>
                      </div>
                      <div style={{ width: 130 }}>
                        <div style={{ fontSize: 10, color: '#8B8178', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 3 }}>Prezzo unitario €</div>
                        <input value={prices[it.id] ?? ''} onChange={(e) => setPrices((pr) => ({ ...pr, [it.id]: e.target.value }))} inputMode="decimal" placeholder="0,00" style={{ ...field, textAlign: 'right', fontVariantNumeric: 'tabular-nums', background: p === null ? '#FFFBEA' : '#fff' }} />
                      </div>
                      <div style={{ width: 120, textAlign: 'right' }}>
                        <div style={{ fontSize: 10, color: '#8B8178', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Importo</div>
                        <div style={{ fontSize: 14, fontWeight: 600, color: '#862D2D', fontVariantNumeric: 'tabular-nums', whiteSpace: 'nowrap', marginTop: 8 }}>{importo !== null ? formatEuro(importo) : '—'}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ))}
          </section>
        ))}

        {(data.snapshot?.noteFine || '').trim() && (
          <div style={{ background: '#fff', border: '1px solid #E3D9CC', borderRadius: 14, padding: '12px 16px', marginBottom: 16, fontSize: 13, whiteSpace: 'pre-wrap', lineHeight: 1.5 }}>
            <div style={{ fontSize: 10.5, color: '#8B8178', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 4 }}>Note</div>
            {data.snapshot.noteFine}
          </div>
        )}

        <div style={{ background: '#fff', border: '1px solid #E3D9CC', borderRadius: 16, padding: 16 }}>
          <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 10 }}>I tuoi riferimenti</div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 10, marginBottom: 10 }}>
            <input value={referente} onChange={(e) => setReferente(e.target.value)} placeholder="Referente" style={field} />
            <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="Email" type="email" style={field} />
            <input value={telefono} onChange={(e) => setTelefono(e.target.value)} placeholder="Telefono" style={field} />
          </div>
          <textarea value={notaGenerale} onChange={(e) => setNotaGenerale(e.target.value)} placeholder="Note generali sull'offerta (tempi, condizioni, esclusioni…)" rows={3} style={{ ...field, resize: 'vertical' }} />
        </div>
      </main>

      <div style={{ position: 'fixed', left: 0, right: 0, bottom: 0, background: '#fff', borderTop: '1px solid #E3D9CC', padding: '12px 16px', zIndex: 20 }}>
        <div style={{ maxWidth: 1060, margin: '0 auto', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 14, flexWrap: 'wrap' }}>
          <div style={{ fontSize: 12.5, color: '#8B8178' }}>{filled} di {allItems.length} voci compilate · <b style={{ color: '#2B2321' }}>Totale {formatEuro(total)}</b> IVA esclusa</div>
          <button disabled={sending} onClick={submit} style={{ background: '#862D2D', color: '#fff', border: 'none', borderRadius: 999, padding: '12px 26px', fontSize: 14, fontWeight: 500, cursor: sending ? 'default' : 'pointer', fontFamily: FONT }}>
            {sending ? 'Invio…' : sentAt ? 'Invia di nuovo' : 'Invia allo studio'}
          </button>
        </div>
      </div>
    </div>
  );
}
// =================== fine offerte imprese ===================

// Vista pubblica per un fornitore esterno: niente account, si accede con un link + PIN che lo studio genera
// dalla pagina "Listino prezzi" (bottone "Condividi con fornitore"). Il fornitore vede le voci del listino
// scelto — con il costo cliente, come riferimento — e può proporre il proprio costo impresa voce per voce.
// L'invio NON tocca subito il listino dello studio: resta "in attesa" nella tabella cea.fornitore_submissions
// finché qualcuno in studio non lo approva dalla stessa pagina "Listino prezzi".
function FornitoreShareView({ token }) {
  const [pin, setPin] = useState('');
  const [checking, setChecking] = useState(false);
  const [error, setError] = useState('');
  const [data, setData] = useState(null); // { nomeFornitore, listinoNome, macros }
  const [values, setValues] = useState({}); // code -> { impresa, note }
  const [expanded, setExpanded] = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [sentAt, setSentAt] = useState(null);

  const unlock = async () => {
    if (!pin.trim() || checking) return;
    setChecking(true); setError('');
    const { data: result, error: err } = await cea.rpc('fornitore_get_listino', { p_token: token, p_pin: pin.trim() });
    setChecking(false);
    if (err || (result && result.errore)) { setError(result && result.errore ? result.errore : 'Link o PIN non validi. Controlla con lo studio.'); return; }
    setData(result);
    const initial = {};
    (result.macros || []).forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((s) => (s.voci || []).forEach((v) => {
      initial[v.code] = { impresa: v.priceImpresa || '', note: '' };
    }))));
    setValues(initial);
  };

  const setImpresa = (code, val) => setValues((prev) => ({ ...prev, [code]: { ...prev[code], impresa: val } }));
  const setNote = (code, val) => setValues((prev) => ({ ...prev, [code]: { ...prev[code], note: val } }));

  const submit = async () => {
    // Invia solo le voci che il fornitore ha davvero cambiato rispetto al costo impresa già presente nel
    // listino (o a cui ha aggiunto una nota) — non tutte le voci con un campo compilato: ogni voce arriva
    // precompilata con il costo attuale, quindi mandarle tutte riempirebbe "Richieste in attesa" di proposte
    // identiche a quello che c'è già, anche se il fornitore non ha toccato nulla.
    const items = [];
    (data.macros || []).forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((s) => (s.voci || []).forEach((v) => {
      const val = values[v.code];
      if (!val) return;
      const proposedImpresa = String(val.impresa).trim();
      if (proposedImpresa === '') return;
      const originalImpresa = String(v.priceImpresa ?? '').trim();
      const note = (val.note || '').trim();
      if (proposedImpresa === originalImpresa && note === '') return;
      items.push({ code: v.code, desc: v.desc, costoImpresa: proposedImpresa, note });
    }))));
    if (items.length === 0) { alert('Non hai modificato nessuna voce: cambia almeno un costo impresa (o aggiungi una nota) prima di inviare.'); return; }
    setSubmitting(true);
    const { data: esito, error: err } = await cea.rpc('fornitore_submit', { p_token: token, p_pin: pin.trim(), p_items: items });
    setSubmitting(false);
    if (err || (esito && esito.errore)) { alert('Invio non riuscito: ' + (err ? err.message : esito.errore)); return; }
    setSentAt(nowLabel());
  };

  if (!data) {
    return (
      <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT, padding: 16 }}>
        <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360, maxWidth: '100%' }}>
          <img src="/quant-logo.png" alt="Quant" style={{ width: 40, height: 40, borderRadius: 11, marginBottom: 16, objectFit: 'cover' }} />
          <h1 style={{ fontSize: 20, fontWeight: 700, color: C.black, margin: '0 0 4px' }}>Listino fornitore</h1>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 24px' }}>Inserisci il PIN che ti ha dato lo studio per vedere e compilare le voci.</p>
          <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>PIN</label>
          <input
            value={pin}
            onChange={(e) => setPin(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && unlock()}
            placeholder="••••"
            style={{ width: '100%', fontSize: 15, letterSpacing: '0.2em', padding: '10px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px' }}
          />
          {error && <p style={{ fontSize: 12, color: C.maroon, margin: '0 0 10px' }}>{error}</p>}
          <button disabled={checking} onClick={unlock} style={{ width: '100%', background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '11px 0', fontSize: 13, fontWeight: 600, cursor: checking ? 'default' : 'pointer', opacity: checking ? 0.7 : 1 }}>
            {checking ? 'Un attimo…' : 'Continua'}
          </button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, fontFamily: FONT }}>
      <div style={{ background: C.black, color: C.white, padding: '16px 24px', display: 'flex', alignItems: 'center', gap: 12 }}>
        <img src="/quant-logo.png" alt="Quant" style={{ width: 32, height: 32, borderRadius: 9, flexShrink: 0, objectFit: 'cover' }} />
        <div>
          <div style={{ fontSize: 14, fontWeight: 600 }}>{data.listinoNome}</div>
          <div style={{ fontSize: 11, color: '#AAA39A' }}>Desearq Studio{data.nomeFornitore ? ` — proposta per ${data.nomeFornitore}` : ''}</div>
        </div>
      </div>

      <div style={{ maxWidth: 880, margin: '0 auto', padding: '24px 16px 110px' }}>
        <div style={{ ...card, marginBottom: 18, fontSize: 12.5, color: C.darkGray, lineHeight: 1.6 }}>
          Per ogni voce trovi il <strong>costo cliente</strong> (come riferimento) e un campo dove indicare il tuo <strong>costo impresa</strong>. I valori inviati restano in attesa di approvazione dello studio prima di entrare nel listino.
        </div>

        {(data.macros || []).map((m, mi) => (
          <div key={mi} style={{ ...card, marginBottom: 14 }}>
            <div onClick={() => setExpanded((e) => ({ ...e, [mi]: e[mi] === false ? true : false }))} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer' }}>
              <h2 style={{ fontSize: 16, margin: 0, color: C.black, fontFamily: FONT }}>{m.name}</h2>
              <span style={{ fontSize: 12, color: C.gray }}>{expanded[mi] === false ? 'Mostra ▾' : 'Nascondi ▴'}</span>
            </div>
            {expanded[mi] !== false && (m.categorie || []).map((c, ci) => (
              <div key={ci} style={{ marginTop: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: C.midGray, marginBottom: 6 }}>{c.name}</div>
                {(c.sottocategorie || []).map((s, si) => (
                  <div key={si} style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 11, color: C.gray, marginBottom: 6 }}>{s.name}</div>
                    {(s.voci || []).map((v) => {
                      const impresaForCliente = parseEuro(values[v.code]?.impresa || v.priceImpresa);
                      const cliente = evalClientPrice(v.priceCliente, impresaForCliente);
                      return (
                        <div key={v.code} style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 10, padding: '10px 0', borderTop: `1px solid ${C.paleGray}` }}>
                          <div style={{ flex: '1 1 240px', minWidth: 200 }}>
                            <div style={{ fontSize: 13, color: C.black }}>{v.desc}</div>
                            <div style={{ fontSize: 11, color: C.gray }}>{v.code} · {v.unit}</div>
                          </div>
                          <div style={{ width: 110, textAlign: 'right' }}>
                            <div style={{ fontSize: 10, color: C.gray, textTransform: 'uppercase', letterSpacing: '0.06em' }}>Costo cliente</div>
                            <div style={{ fontSize: 13, color: C.midGray }}>{formatEuro(cliente)}</div>
                          </div>
                          <div style={{ width: 140 }}>
                            <div style={{ fontSize: 10, color: C.gray, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 4 }}>Costo impresa</div>
                            <input
                              value={values[v.code]?.impresa || ''}
                              onChange={(e) => setImpresa(v.code, e.target.value)}
                              placeholder="0,00"
                              style={{ width: '100%', fontSize: 13, padding: '7px 9px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                            />
                          </div>
                          <div style={{ width: 170 }}>
                            <div style={{ fontSize: 10, color: C.gray, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 4 }}>Nota (opzionale)</div>
                            <input
                              value={values[v.code]?.note || ''}
                              onChange={(e) => setNote(v.code, e.target.value)}
                              placeholder="Es. non disponibile"
                              style={{ width: '100%', fontSize: 12, padding: '7px 9px', borderRadius: 6, border: `1px solid ${C.paleGray}` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            ))}
          </div>
        ))}
      </div>

      <div style={{ position: 'fixed', left: 0, right: 0, bottom: 0, background: C.white, borderTop: `1px solid ${C.paleGray}`, padding: '14px 16px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 16 }}>
        {sentAt && <span style={{ fontSize: 12, color: C.success }}>Inviato alle {sentAt.split(', ')[1]} ✓</span>}
        <button disabled={submitting} onClick={submit} style={{ background: C.black, color: C.white, border: 'none', borderRadius: 999, padding: '12px 28px', fontSize: 13, fontWeight: 600, cursor: submitting ? 'default' : 'pointer', opacity: submitting ? 0.7 : 1 }}>
          {submitting ? 'Invio…' : sentAt ? 'Invia di nuovo' : 'Invia allo studio'}
        </button>
      </div>
    </div>
  );
}

export default function GestionaleEdilePreview() {
  // Vista pubblica per il fornitore (link + PIN, nessun account): va controllata primissima cosa, prima di
  // qualsiasi hook legato all'autenticazione dello studio, così chi apre questo link non passa mai dalla
  // schermata di accesso né tocca i dati del team.
  const fornitoreToken = new URLSearchParams(window.location.search).get('fornitore');
  if (fornitoreToken) {
    return <FornitoreShareView token={fornitoreToken} />;
  }
  // Vista pubblica per un'impresa che compila i prezzi di un computo (link + PIN, nessun account)
  const offertaToken = new URLSearchParams(window.location.search).get('offerta');
  if (offertaToken) {
    return <OffertaImpresaView token={offertaToken} />;
  }

  const [page, setPage] = useState('dashboard');
  const [projects, setProjects] = useState(PROJECTS);
  const [priceNotice, setPriceNotice] = useState(null); // avviso temporaneo dopo l'aggiornamento prezzi nei computi
  const [selectedProjectId, setSelectedProjectId] = useState(null);
  const [openRevisionId, setOpenRevisionId] = useState(null);
  const [listini, setListini] = useState([{ id: 1, name: 'Listino standard 2026', macros: INITIAL_MACROS }]);
  const [activeListinoId, setActiveListinoId] = useState(1);
  const [fornitoriCatalog, setFornitoriCatalog] = useState(INITIAL_FORNITORI);
  const [studioSettings, setStudioSettings] = useState(DEFAULT_STUDIO_SETTINGS);
  React.useEffect(() => { applyAccentColor(studioSettings.coloreAccento); }, [studioSettings.coloreAccento]);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const topHeaderRef = useRef(null);
  React.useEffect(() => {
    const el = topHeaderRef.current;
    if (!el) return undefined;
    const setH = () => document.documentElement.style.setProperty('--header-h', `${Math.ceil(el.getBoundingClientRect().height)}px`);
    setH();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(setH) : null;
    if (ro) ro.observe(el); else window.addEventListener('resize', setH);
    return () => { if (ro) ro.disconnect(); else window.removeEventListener('resize', setH); };
  });

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
    supabase.rpc('quant_io').then(({ data, error }) => {
      if (cancelled) return;
      if (error) { setProfileError(error.message); return; }
      if (!data || !data.profilo) { setProfileError('Il Gestionale non è attivo per questo accesso. Chiedi a un amministratore dello studio.'); return; }
      setProfile({ ...data.profilo, org_id: data.org_id || data.profilo.org_id });
    });
    return () => { cancelled = true; };
  }, [authUser]);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setProfile(null);
  };

  // --- Dati condivisi del workspace (progetti, listini, fornitori): un unico documento salvato su Supabase ---
  const [dataLoaded, setDataLoaded] = useState(false);
  const [needsImport, setNeedsImport] = useState(false);
  const [saveStatus, setSaveStatus] = useState('ok'); // 'ok' | 'saving' | 'error'
  const saveTimer = React.useRef(null);

  React.useEffect(() => {
    if (!profile || profile.status !== 'Attivo') return;
    supabase.rpc('quant_carica').then(({ data: r, error }) => {
      if (error) { console.error(error); setDataLoaded(true); return; }
      const saved = (r && r.dati) || {};
      const ORG_STORICA = 'd0000000-0000-4000-8000-0000000000d5'; // lo studio che arriva dal vecchio Gestionale
      const nuovoStudio = (!saved.projects || !saved.projects.length) && profile.org_id && profile.org_id !== ORG_STORICA;
      if (nuovoStudio) {
        // uno studio nuovo parte da zero: niente dati di esempio e niente importazione
        setProjects([]);
        setListini([{ id: 1, name: 'Listino standard', macros: [] }]);
        setActiveListinoId(1);
        setFornitoriCatalog([]);
      } else if (!saved.projects || !saved.projects.length || /[?&]importa=1/.test(window.location.search)) setNeedsImport(true);
      if (saved.projects) setProjects(saved.projects);
      if (saved.listini) setListini(ensureVoceUids(saved.listini));
      if (saved.activeListinoId) setActiveListinoId(saved.activeListinoId);
      if (saved.fornitoriCatalog) setFornitoriCatalog(migrateFornitoriCatalog(saved.fornitoriCatalog));
      if (saved.studioSettings) setStudioSettings({ ...DEFAULT_STUDIO_SETTINGS, ...saved.studioSettings });
      setDataLoaded(true);
    });
  }, [profile?.id, profile?.status]);

  React.useEffect(() => {
    if (!dataLoaded || needsImport || !profile || profile.status !== 'Attivo') return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    setSaveStatus('saving');
    saveTimer.current = setTimeout(() => {
      supabase.rpc('quant_salva', {
        p_dati: { projects, listini, activeListinoId, fornitoriCatalog, studioSettings },
      }).then(({ data: r, error }) => {
        if (error) { console.error('Salvataggio fallito:', error.message); setSaveStatus('error'); } else { if (r && r.proiezione_errore) console.warn('Riallineamento tabelle suite:', r.proiezione_errore); setSaveStatus('ok'); }
      });
    }, 900);
    return () => clearTimeout(saveTimer.current);
  }, [projects, listini, activeListinoId, fornitoriCatalog, studioSettings, dataLoaded, needsImport]);

  // Ogni modifica ai listini passa da qui: se cambiano i prezzi di qualche voce, li si riporta nei computi
  // in attesa di approvazione (vedi propagateListinoPrices) e si avvisa quante voci sono state aggiornate.
  const updateListini = (nextListini) => {
    const resolved = ensureVoceUids(typeof nextListini === 'function' ? nextListini(listini) : nextListini);
    const changes = diffListinoPrices(listini, resolved);
    setListini(resolved);
    if (changes.length === 0) return;
    const { next, touchedItems, touchedProjects } = propagateListinoPrices(projects, changes);
    if (touchedItems > 0) {
      setProjects(next);
      const what = changes.some((c) => c.priceChanged) && changes.some((c) => c.descChanged) ? 'Prezzo e descrizione aggiornati' : changes.some((c) => c.descChanged) ? 'Descrizione aggiornata' : 'Prezzo aggiornato';
      setPriceNotice(`${what} anche in ${touchedItems} vo${touchedItems === 1 ? 'ce' : 'ci'} di ${touchedProjects} computo${touchedProjects === 1 ? '' : 'i'} in attesa di approvazione. I computi approvati restano invariati.`);
      setTimeout(() => setPriceNotice(null), 7000);
    }
  };

  const openProject = (id) => { setSelectedProjectId(id); setOpenRevisionId(null); setPage('progetto-dettaglio'); };
  const openRevisionInProject = (projectId, revisionId) => { setSelectedProjectId(projectId); setOpenRevisionId(revisionId); setPage('progetto-dettaglio'); };
  const selectedProject = projects.find((p) => p.id === selectedProjectId);
  const [rilievoProjectId, setRilievoProjectId] = useState(null); // progetto preselezionato quando si entra nel Rilievo planimetrico da una scheda progetto

  const [printJob, setPrintJob] = useState(null); // { project, revision, clientOnly }
  const requestPdf = (project, revision, clientOnly) => setPrintJob({ project, revision, clientOnly });

  React.useEffect(() => {
    if (!printJob) return;
    let cancelled = false;
    // Prima di stampare: un minimo di tempo perché il DOM di stampa si monti, poi — se è attivo un font
    // personalizzato — lo si carica esplicitamente con la FontFace API invece di limitarsi a dichiararlo in
    // CSS (@font-face) e sperare che il browser lo richieda in tempo: legato al solo CSS, un font incorporato
    // come data URL può restare "in coda" e non essere pronto quando parte la stampa, col risultato che il PDF
    // esce col font di riserva senza errori né avvisi. Un tetto massimo di attesa evita che un font rotto o
    // troppo lento blocchi la stampa indefinitamente (si stampa comunque, col font di riserva).
    const ssNow = studioSettings || DEFAULT_STUDIO_SETTINGS;
    const hasCustomFontNow = ssNow.usaFontPersonalizzato && ssNow.fontPersonalizzato;
    const minWait = new Promise((resolve) => setTimeout(resolve, 150));
    const timeoutGuard = new Promise((resolve) => setTimeout(resolve, 1500));
    const fontLoad = hasCustomFontNow
      ? (async () => {
          try {
            const face = new FontFace('StudioFontPersonalizzato', `url(${JSON.stringify(ssNow.fontPersonalizzato)})`);
            const loaded = await face.load();
            document.fonts.add(loaded);
          } catch (e) {
            console.error('Caricamento del font personalizzato fallito, stampa con il font di riserva:', e);
          }
        })()
      : Promise.resolve();
    Promise.all([minWait, Promise.race([fontLoad, timeoutGuard])]).then(() => { if (!cancelled) window.print(); });
    const handleAfterPrint = () => setPrintJob(null);
    window.addEventListener('afterprint', handleAfterPrint);
    return () => { cancelled = true; window.removeEventListener('afterprint', handleAfterPrint); };
  }, [printJob, studioSettings]);

  if (authUser === undefined) {
    return <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT, color: C.gray, fontSize: 13 }}>Caricamento…</div>;
  }

  if (!authUser) {
    return <LoginScreen onSignedIn={() => {}} />;
  }

  if (!profile) {
    return (
      <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT }}>
        <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto', textAlign: 'center' }}>
          <p style={{ fontSize: 13, color: C.black, margin: '0 0 12px' }}>{profileError || 'Nessun invito trovato per questa email. Chiedi a un admin di invitarti dalla sezione Team, poi ricarica la pagina.'}</p>
          <button onClick={handleLogout} style={{ background: C.darkGray, color: C.white, border: 'none', borderRadius: 999, padding: '9px 18px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Esci</button>
        </div>
      </div>
    );
  }

  if (needsImport && dataLoaded) {
    return <ImportaDaVecchioGestionale isAdmin={profile.role === 'Admin'} onDone={() => window.location.reload()} onSkip={() => setNeedsImport(false)} />;
  }

  if (profile.status !== 'Attivo') {
    return (
      <div style={{ minHeight: '100vh', background: PAGE_GRADIENT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: FONT }}>
        <div style={{ background: C.white, borderRadius: 20, boxShadow: '0 8px 24px rgba(0,0,0,0.10)', padding: 32, width: 360, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto', textAlign: 'center' }}>
          <p style={{ fontSize: 13, color: C.black, margin: 0 }}>Il tuo invito non è stato ancora completato. Riprova tra poco o contatta l'admin.</p>
        </div>
      </div>
    );
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        html, body, #root { max-width: 100%; overflow-x: clip; }
        body { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }
        .team-card:hover { border-color: #CFC4B6 !important; }
        .export-menu-item:hover { background: #F6F3EE !important; }
        @media (max-width: 700px) { .computo-toolbar > div { border-right: none !important; padding: 6px 10px !important; } }
        .print-only { display: none; }
        /* Margine vero della pagina stampata: niente (testo, tabelle, intestazione, piè) tocca più il bordo
           fisico del foglio. Le immagini personalizzate di intestazione/piè riempiono la larghezza DISPONIBILE
           dentro questo margine, non il foglio intero — la loro dimensione si regola con la "scala" in
           Impostazioni studio, non allargando il margine della pagina. */
        @page { margin: 16mm 14mm; }
        @media print {
          html, body { margin: 0 !important; padding: 0 !important; }
          .no-print { display: none !important; }
          .print-only { display: block !important; }
          .print-only table { border-collapse: collapse; }
          .print-only thead { display: table-header-group; break-inside: avoid; page-break-inside: avoid; }
          .print-only tfoot { display: table-footer-group; }
        }
        .sidebar-item-btn:not(.active):hover { background: #E9E0D6 !important; color: #4A403C !important; }
        .btn-accent-pill:hover { background: #6E2424 !important; }
        .input-focus:focus { outline: none; border-color: ${C.maroon} !important; box-shadow: 0 0 0 3px rgba(134,45,45,0.25); }

        /* --- Rendering mobile --- */
        input, select, textarea, button { max-width: 100%; }
        .table-scroll { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }
        .table-scroll > table { min-width: 620px; }
        .mobile-menu-btn { display: none; }
        .sidebar-backdrop { display: none; }
        .form-grid { display: grid; }

        /* Su schermi stretti, qualunque riga flessibile può andare a capo invece di traboccare.
           Esclusi i contenitori "flex-direction:column" (:not([style*="column"])): per una colonna, andare
           a capo significa aprire una SECONDA COLONNA affiancata invece di impilare — è esattamente il bug
           della card "Listino" nel computo (sticky, con scroll interno): il suo contenuto veniva spinto fuori
           a destra, oltre il bordo della card, su schermo stretto. Le colonne non hanno comunque bisogno di
           andare a capo: i loro figli sono già impilati verticalmente dalla flex-direction stessa. */
        @media (max-width: 760px) {
          [style*="display:flex"]:not([style*="column"]), [style*="display: flex"]:not([style*="column"]) { flex-wrap: wrap; min-width: 0; }

          /* La card "Listino" nel computo è pensata per stare FISSA ("sticky") accanto al computo quando le
             due colonne sono affiancate su schermi larghi: resta visibile mentre si scorre il computo per
             trascinare le voci. Sotto i 760px le due colonne vanno già a capo (impilate, per via della regola
             sopra) e quel comportamento "sticky" smette di avere senso: la card, restando agganciata in alto,
             finirebbe per sovrapporsi al computo che ora le scorre sotto nello stesso spazio orizzontale,
             coprendone parte del contenuto. Su mobile diventa quindi un blocco normale a piena larghezza. */
          .listino-catalog-card {
            position: static !important;
            width: 100% !important;
            max-height: none !important;
          }
          .listino-catalog-scroll {
            flex: none !important;
            min-height: 0 !important;
            max-height: 420px;
          }
        }

        @media (max-width: 880px) {
          .app-shell { position: relative; }
          .sidebar-aside {
            position: fixed !important;
            top: 0; left: 0;
            height: 100vh !important;
            max-height: 100dvh;
            transform: translateX(-100%);
            transition: transform 0.22s ease;
            z-index: 60;
            box-shadow: 6px 0 24px rgba(0,0,0,0.18);
          }
          .sidebar-aside.open { transform: translateX(0); }
          .sidebar-backdrop.open {
            display: block;
            position: fixed; inset: 0;
            background: rgba(5,5,5,0.45);
            z-index: 55;
          }
          .mobile-menu-btn {
            display: inline-flex !important;
            align-items: center; justify-content: center;
            width: 38px; height: 38px;
            border-radius: 10px;
            border: 1px solid ${C.paleGray};
            background: ${C.white};
            cursor: pointer;
            flex-shrink: 0;
          }
          .top-header { padding: 14px 16px 0 !important; justify-content: space-between !important; }
          .search-input { width: 100% !important; }
          .search-input-wrap { flex: 1; min-width: 0; }
          main.app-main { padding: 16px !important; }
          h1 { font-size: 24px !important; }
        }

        @media (max-width: 640px) {
          .form-grid { grid-template-columns: 1fr !important; }
        }
      `}</style>
      {printJob && (
        <div className="print-only">
          <PrintableComputo project={printJob.project} revision={printJob.revision} clientOnly={printJob.clientOnly} studioSettings={studioSettings} />
        </div>
      )}
      <div className="no-print app-shell" style={{ display: 'flex', minHeight: '100vh', fontFamily: FONT, background: '#F4EFE8' }}>
      <div className={`sidebar-backdrop${mobileNavOpen ? ' open' : ''}`} onClick={() => setMobileNavOpen(false)} />
      {priceNotice && (
        <div role="status" onClick={() => setPriceNotice(null)} style={{ position: 'fixed', right: 24, bottom: 24, zIndex: 300, maxWidth: 380, background: '#2B2321', color: '#fff', borderRadius: 12, padding: '12px 16px', fontSize: 13, fontWeight: 300, lineHeight: 1.45, boxShadow: '0 10px 30px rgba(0,0,0,.18)', cursor: 'pointer' }}>
          {priceNotice}
        </div>
      )}
      <aside className={`sidebar-aside${mobileNavOpen ? ' open' : ''}`} style={{ width: 270, flexShrink: 0, background: '#FBFAF7', color: C.darkGray, borderRight: `1px solid ${C.paleGray}`, display: 'flex', flexDirection: 'column', padding: '28px 16px 22px', boxSizing: 'border-box', gap: 2 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '0 12px 18px' }}>
          <img src="/quant-logo-full.png" alt="Quant" width={169} height={228} style={{ maxWidth: '100%', height: 63, width: 'auto', display: 'block', objectFit: 'contain', aspectRatio: '169 / 228' }} />
        </div>
        <div style={{ fontSize: 9.5, letterSpacing: 2, color: '#8C8776', textTransform: 'uppercase', marginTop: 1, lineHeight: 2, marginBottom: 10 }}>Desearq Lab</div>

        {NAV_GROUPS.map((group, gi) => (
          <React.Fragment key={gi}>
            {group.label && (
              <div style={{ fontSize: 10, letterSpacing: '.16em', textTransform: 'uppercase', color: '#9A8F86', fontWeight: 600, padding: '14px 14px 4px' }}>{group.label}</div>
            )}
            {group.items.map((item) => {
              const active = page === item.key;
              const count = item.key === 'progetti'
                ? projects.length
                : item.key === 'computi'
                  ? projects.filter((p) => latestStatus(p) === 'In attesa di approvazione').length
                  : null;
              return (
                <a
                  key={item.key}
                  className={`sidebar-item-btn${active ? ' active' : ''}`}
                  onClick={() => { setPage(item.key); setMobileNavOpen(false); }}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 13, padding: '8px 14px', borderRadius: 10, lineHeight: 1.25,
                    fontFamily: FONT, fontSize: 15, fontWeight: 400, cursor: 'pointer', textDecoration: 'none',
                    color: active ? C.white : '#4A403C',
                    background: active ? C.black : 'transparent',
                  }}
                >
                  {item.icon}
                  <span style={{ flex: count ? '0 1 auto' : 1 }}>{item.label}</span>
                  {!!count && (
                    <span style={{ marginLeft: 'auto', fontSize: 11, fontWeight: 600, minWidth: 20, height: 20, borderRadius: 10, background: active ? '#fff' : C.maroon, color: active ? '#2B2321' : '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 5px' }}>{count}</span>
                  )}
                </a>
              );
            })}
          </React.Fragment>
        ))}

        <div style={{ flexGrow: 1, minHeight: 24 }} />
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 14px', borderRadius: 12, background: '#2B2321' }}>
          <svg viewBox="0 0 24 24" style={{ width: 20, height: 20, fill: 'none', stroke: '#fff', strokeWidth: 1.5 }}><path d="M6 3h9l4 4v14H6z" /></svg>
          <div style={{ flexGrow: 1, lineHeight: 1.25, minWidth: 0 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: '#fff' }}>Studio</div>
            <div style={{ fontSize: 12, color: '#fff', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{studioSettings.nome || 'Desearq Studio'}</div>
          </div>
          <span onClick={() => { setPage('impostazioni'); setMobileNavOpen(false); }} style={{ color: '#B8ADA0', cursor: 'pointer' }}>&rsaquo;</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '10px 0 2px', padding: 12, border: '1px solid #E8E0D4', borderRadius: 14, background: '#fff' }}>
          <span onClick={() => { setPage('team'); setMobileNavOpen(false); }} style={{ cursor: 'pointer', display: 'flex' }} title="Apri il team">
            <MemberAvatar member={profile} size={38} />
          </span>
          <div style={{ flex: 1, minWidth: 0, lineHeight: 1.25 }}>
            <b style={{ display: 'block', fontWeight: 400, fontSize: 13.5, color: '#2B2321', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{profile.name || profile.email}</b>
            <span style={{ display: 'block', fontSize: 11.5, color: '#8B8178', fontWeight: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{profile.job_title || (profile.role === 'Admin' ? 'Amministratore' : 'Membro del team')}</span>
          </div>
          <button onClick={handleLogout} title="Esci" style={{ border: '1px solid #E3D9CC', background: '#fff', color: '#2B2321', borderRadius: 999, padding: '6px 12px', fontFamily: FONT, fontSize: 12, fontWeight: 300, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, flexShrink: 0 }}>
            <svg viewBox="0 0 24 24" width={14} height={14} fill="none" stroke="currentColor" strokeWidth={1.5} strokeLinecap="round" strokeLinejoin="round"><path d="M9 3.5H5.5a1 1 0 00-1 1v15a1 1 0 001 1H9" /><path d="M15 8l5 4-5 4" /><line x1="20" y1="12" x2="9.5" y2="12" /></svg>
            Esci
          </button>
        </div>
        <div style={{ padding: '10px 14px 0', fontSize: 10.5, fontWeight: 300, color: '#8B8178' }}>
          {saveStatus === 'error'
            ? <span style={{ color: C.maroon }}>Salvataggio non riuscito · riprova tra poco</span>
            : saveStatus === 'saving'
              ? <>Salvataggio in corso…</>
              : <>Dati sincronizzati &middot; <span style={{ color: '#3E6B52' }}>&#9679;</span> live</>}
        </div>
      </aside>

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, width: '100%' }}>
        <header ref={topHeaderRef} className="top-header" style={{ background: '#FAF9F6', padding: '14px 32px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 20, flexWrap: 'wrap', position: 'sticky', top: 0, zIndex: 20, borderBottom: '1px solid #E8E0D4' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <button
              className="mobile-menu-btn"
              onClick={() => setMobileNavOpen(true)}
              aria-label="Apri il menu"
              style={{ fontSize: 16 }}
            >
              ☰
            </button>
            <select
              value={activeListinoId}
              onChange={(e) => setActiveListinoId(Number(e.target.value))}
              style={{ border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '9px 16px', fontSize: 13, fontWeight: 300, fontFamily: FONT, color: '#2B2321', cursor: 'pointer', minWidth: 180 }}
            >
              {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            </select>
            <select
              value={selectedProjectId || ''}
              onChange={(e) => { if (e.target.value) openProject(Number(e.target.value)); }}
              style={{ border: '1px solid #E3D9CC', background: '#fff', borderRadius: 999, padding: '9px 16px', fontSize: 13, fontWeight: 300, fontFamily: FONT, color: '#2B2321', cursor: 'pointer', minWidth: 170 }}
            >
              <option value="">Scegli progetto</option>
              {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 26, fontSize: 13, fontWeight: 300, color: '#5E5550', flexWrap: 'wrap' }}>
            <div className="search-input-wrap" style={{ display: 'inline-flex', alignItems: 'center' }}>
              <input
                type="text"
                placeholder="Cerca progetto, cliente, voce o codice…"
                className="search-input"
                style={{ border: 'none', outline: 'none', background: 'transparent', fontFamily: FONT, fontSize: 13, fontWeight: 300, color: '#5E5550', width: 'clamp(160px, 22vw, 250px)', padding: 0 }}
              />
            </div>
            <button
              onClick={() => setPage('progetti')}
              style={{ border: 0, background: '#1E1A18', color: '#F7F3EC', fontFamily: FONT, fontSize: 13, fontWeight: 400, padding: '11px 22px', borderRadius: 999, cursor: 'pointer' }}
            >
              Nuovo progetto +
            </button>
          </div>
        </header>

        {/* overflowX qui era 'auto': su qualunque browser questo trasforma <main> in un "contenitore di scroll",
            che diventa il riferimento per il posizionamento "sticky" al suo interno — rompendo lo sticky della
            colonna Listino nel computo metrico (restava ancorato al box di <main>, che non scorre mai da solo,
            invece che alla pagina). 'clip' taglia comunque l'overflow orizzontale ma senza creare quel contenitore,
            quindi lo sticky funziona di nuovo; le singole tabelle larghe restano scorrevili con la classe .table-scroll. */}
        <main className="app-main" style={{ padding: '22px 32px 48px', flex: 1, overflowX: 'clip', minWidth: 0, width: '100%', background: '#FAF9F6' }}>
          {page === 'dashboard' && <Dashboard onNavigate={setPage} onOpenProject={openProject} projects={projects} listini={listini} activeListinoId={activeListinoId} fornitoriCatalog={fornitoriCatalog} profile={profile} />}
          {page === 'listino' && <ListinoPage listini={listini} setListini={updateListini} activeId={activeListinoId} setActiveId={setActiveListinoId} projects={projects} setProjects={setProjects} />}
          {page === 'progetti' && <ProgettiPage projects={projects} setProjects={setProjects} onOpenProject={openProject} />}
          {page === 'progetto-dettaglio' && selectedProject && (
            <ProjectDetailPage
              project={selectedProject}
              onBack={() => setPage('progetti')}
              onUpdateProject={(updated) => setProjects(projects.map((p) => (p.id === updated.id ? updated : p)))}
              listini={listini}
              onUpdateListini={updateListini}
              fornitoriCatalog={fornitoriCatalog}
              initialRevisionId={openRevisionId}
              requestPdf={requestPdf}
              onOpenRilievo={() => { setRilievoProjectId(selectedProject.id); setPage('rilievo'); }}
              profile={profile}
            />
          )}
          {page === 'computi' && <ComputiPage projects={projects} setProjects={setProjects} onOpenProject={openProject} onOpenRevision={openRevisionInProject} requestPdf={requestPdf} />}
          {page === 'confronto' && <ConfrontoPage projects={projects} />}
          {page === 'rilievo' && <RilievoPage projects={projects} setProjects={setProjects} listini={listini} initialProjectId={rilievoProjectId} />}
          {page === 'fornitori' &&<FornitoriPage projects={projects} setProjects={setProjects} catalog={fornitoriCatalog} setCatalog={setFornitoriCatalog} />}
          {page === 'team' && <TeamPage profile={profile} onProfileChange={setProfile} />}
          {page === 'impostazioni' && <ImpostazioniPage settings={studioSettings} onUpdate={(patch) => setStudioSettings((s) => ({ ...s, ...patch }))} />}
        </main>
      </div>
      </div>
    </>
  );
}
