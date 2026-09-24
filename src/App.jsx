import React, { useState } from 'react';
import { supabase, cea } from './supabaseClient';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { LayoutGrid, BookOpen, Building2, Calculator, GitCompare, Truck, Users, Search, LogOut, MapPin, Clock, Calendar as CalendarIcon, Settings, UserSearch } from 'lucide-react';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

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
  { key: 'impostazioni', label: 'Impostazioni studio', Icon: Settings },
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
      out.push({ ...v, macro: m.name, categoria: c.name, sottocategoria: s.name, impresaValue: impresa, clienteValue: cliente });
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
            style={{ display: 'flex', flexWrap: 'wrap', rowGap: 8, alignItems: 'center', gap: 16, padding: 16, borderRadius: 12, background: C.surfaceSubtle, cursor: 'pointer', marginBottom: 8 }}
          >
            <div style={{ width: 44, height: 44, borderRadius: 8, background: C.accentSoft, color: C.maroon, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 14, flexShrink: 0 }}>
              {initials(p.client)}
            </div>
            <div style={{ flex: '1 1 160px', minWidth: 0 }}>
              <p style={{ fontWeight: 500, fontSize: 13, margin: 0, color: C.black }}>{p.name}</p>
              <p style={{ fontSize: 12, color: C.darkGray, margin: '2px 0 0' }}>{p.client}</p>
            </div>
            <span style={{ fontSize: 13, fontWeight: 500, padding: '4px 12px', borderRadius: 999, flexShrink: 0, ...badgeStyles[statusTone[latestStatus(p)]] }}>{latestStatus(p)}</span>
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
          <button onClick={() => setShowFornitoreSharing(true)} style={{ position: 'relative', background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>
            🔗 Fornitori
            {pendingCount > 0 && (
              <span style={{ position: 'absolute', top: -6, right: -6, background: C.maroon, color: C.white, borderRadius: 999, minWidth: 18, height: 18, fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px' }}>{pendingCount}</span>
            )}
          </button>
        </div>
      </div>

      <EditableCatalog macros={active.macros} setMacros={setMacrosForActive} />

      {showFornitoreSharing && (
        <FornitoreSharingModal listini={listini} setListini={setListini} activeId={activeId} onClose={() => setShowFornitoreSharing(false)} />
      )}
    </div>
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

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(23,23,23,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, padding: 16 }} onClick={onClose}>
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
  const [note, setNote] = useState(initialVoce?.note || '');
  const previewImpresa = parseEuro(priceImpresa);
  const previewCliente = evalClientPrice(priceCliente, previewImpresa);

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
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

        <label style={{ fontSize: 11, fontWeight: 700, color: C.midGray }}>Note (facoltative, solo uso interno)</label>
        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="Es. condizioni particolari, fornitori consigliati, riferimenti capitolato…"
          rows={2}
          style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px', fontFamily: 'inherit', resize: 'vertical' }}
        />

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 6 }}>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Annulla</button>
          <button
            onClick={() => {
              if (!desc.trim()) return;
              const finalUnit = unit === 'Altro' ? (customUnit || '—') : unit;
              onSave(locations[locationIdx].path, { code: code || '—', desc, unit: finalUnit, priceImpresa: priceImpresa || '0,00', priceCliente: priceCliente || '', note: note || '' });
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
function VoceComputoModal({ macroName, categoriaName, sottoName, initialItem, prefill, mergeCandidates, onClose, onSave }) {
  const isEdit = !!initialItem;
  const [desc, setDesc] = useState(initialItem?.desc || prefill?.desc || '');
  const startsKnown = (initialItem || prefill) && UNIT_OPTIONS.slice(0, -1).includes((initialItem || prefill).unit);
  const [unit, setUnit] = useState((initialItem || prefill) ? (startsKnown ? (initialItem || prefill).unit : 'Altro') : 'm²');
  const [customUnit, setCustomUnit] = useState((initialItem || prefill) && !startsKnown ? (initialItem || prefill).unit : '');
  const [priceImpresa, setPriceImpresa] = useState(initialItem?.unitPriceImpresa || prefill?.priceImpresa || '');
  const [priceCliente, setPriceCliente] = useState(initialItem?.unitPriceCliente || prefill?.priceCliente || '');
  const [unitaCalcolo, setUnitaCalcolo] = useState(initialItem ? (initialItem.unitaCalcolo || '') : (prefill ? defaultCalcForUnit(prefill.unit) : 'mq'));
  const initialRows = initialItem ? (initialItem.misurazioni || []).flatMap((g) => g.rows || []) : [];
  const [rows, setRows] = useState(initialRows.length ? initialRows : [emptyMisurazioneRow()]);
  const [manualQty, setManualQty] = useState(initialItem && !initialItem.unitaCalcolo ? (initialItem.qty || '') : '');
  const [mergeIntoId, setMergeIntoId] = useState('');
  const [note, setNote] = useState(initialItem?.note || prefill?.note || '');

  const previewImpresa = parseEuro(priceImpresa);
  const previewCliente = evalClientPrice(priceCliente, previewImpresa);
  const computedQty = unitaCalcolo ? computeGruppoTotal({ rows }, unitaCalcolo) : parseEuro(manualQty);
  const mergeTarget = mergeIntoId ? mergeCandidates.find((c) => String(c.id) === mergeIntoId) : null;

  const updateRow = (idx, field, value) => {
    setRows((rs) => rs.map((r, i) => (i === idx ? { ...r, [field]: value } : r)));
  };
  const addRow = () => setRows((rs) => [...rs, emptyMisurazioneRow()]);
  const removeRow = (idx) => setRows((rs) => (rs.length > 1 ? rs.filter((_, i) => i !== idx) : rs));

  const colStyle = (active) => ({ width: '100%', fontSize: 12, padding: '5px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}`, textAlign: 'right', background: active ? C.white : '#f2f2f2', color: active ? C.black : C.gray });
  const labelStyle = { fontSize: 11, fontWeight: 700, color: C.midGray };
  const fieldStyle = { width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 12px' };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10, padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 640, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
        <h2 style={{ fontFamily: FONT, fontSize: 18, margin: '0 0 4px', color: C.black }}>{isEdit ? 'Modifica voce' : 'Nuova voce'}</h2>
        <p style={{ fontSize: 11, color: C.gray, margin: '0 0 16px' }}>{macroName} › {categoriaName || 'Generale'}</p>

        <label style={labelStyle}>Descrizione tecnica</label>
        <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Descrivi lavorazione, materiali e condizioni…" style={fieldStyle} />

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
        </select>

        {unitaCalcolo ? (
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
                      <td style={{ padding: '3px 4px', textAlign: 'right', fontWeight: 700 }}>{computeMisurazioneRowValue(r, unitaCalcolo).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                      <td style={{ padding: '3px 4px' }}><button onClick={() => removeRow(idx)} style={iconBtn}>🗑</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button onClick={addRow} style={{ ...rowBtnStyle, marginTop: 8 }}>+ Riga di misurazione</button>
            <p style={{ fontSize: 12, fontWeight: 700, color: C.black, margin: '10px 0 0' }}>
              Quantità {mergeTarget ? 'da aggiungere' : 'totale'}: {computedQty.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} {mergeIntoId ? '' : (unit === 'Altro' ? customUnit : unit)}
              {mergeTarget && <> — nuovo totale voce: {(computedQty + parseEuro(mergeTarget.qty)).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</>}
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

        <label style={labelStyle}>Note (facoltative, solo uso interno)</label>
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
              const finalUnit = unit === 'Altro' ? (customUnit || '—') : unit;
              onSave({
                desc, unit: finalUnit, priceImpresa: priceImpresa || '0,00', priceCliente: priceCliente || '',
                unitaCalcolo: unitaCalcolo || null,
                misurazioni: unitaCalcolo ? [{ rows }] : [],
                manualQty: manualQty || '0',
                note: note || '',
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
  );
}

function EditableCatalog({ macros, setMacros }) {
  const [showVoceModal, setShowVoceModal] = useState(false);
  const [editingVoce, setEditingVoce] = useState(null); // { path: [mi,ci,si,vi] }
  const [newVoceAt, setNewVoceAt] = useState(null); // [mi,ci,si] quando si aggiunge dal pulsante "+ Voce" di una sottocategoria
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
            <button onClick={() => { setEditingVoce(null); setNewVoceAt(null); setShowVoceModal(true); }} style={{ background: C.maroon, color: C.white, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Nuova voce</button>
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
                  <button onClick={() => moveItem('macro', [mi], 'up')} disabled={mi === 0} style={{ ...rowBtn, opacity: mi === 0 ? 0.4 : 1 }}>▲</button>
                  <button onClick={() => moveItem('macro', [mi], 'down')} disabled={mi === coded.length - 1} style={{ ...rowBtn, opacity: mi === coded.length - 1 ? 0.4 : 1 }}>▼</button>
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
                        <button onClick={() => moveItem('categoria', [mi, ci], 'up')} disabled={ci === 0} style={{ ...rowBtn, opacity: ci === 0 ? 0.4 : 1 }}>▲</button>
                        <button onClick={() => moveItem('categoria', [mi, ci], 'down')} disabled={ci === m.categorie.length - 1} style={{ ...rowBtn, opacity: ci === m.categorie.length - 1 ? 0.4 : 1 }}>▼</button>
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
                          {isOpen(sKey) && (
                            s.voci.length === 0 ? (
                              <p style={{ fontSize: 12, color: C.gray, padding: '6px 12px 10px 32px' }}>Nessuna voce ancora in questa sottocategoria. Usa "+ Voce" per aggiungerne una.</p>
                            ) : (
                            <div className="table-scroll">
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
                                    <td style={{ padding: '6px 12px 6px 32px', fontWeight: 700, color: C.black, width: 110, verticalAlign: 'top' }}>{v.code}</td>
                                    <td style={{ padding: '6px 12px', color: C.midGray, verticalAlign: 'top' }}>
                                      {v.desc}
                                      {v.note && <p style={{ margin: '3px 0 0', fontSize: 10.5, color: C.gray, fontStyle: 'italic' }}>📝 {v.note}</p>}
                                    </td>
                                    <td style={{ padding: '6px 12px', color: C.gray, width: 70, verticalAlign: 'top' }}>{v.unit}</td>
                                    <td style={{ padding: '6px 12px', textAlign: 'right', fontWeight: 600, color: C.black, width: 90, verticalAlign: 'top' }}>{formatEuro(impresaVal)}</td>
                                    <td style={{ padding: '6px 12px', textAlign: 'right', fontWeight: 600, color: C.maroon, width: 90, verticalAlign: 'top' }}>{formatEuro(clienteVal)}</td>
                                    <td style={{ padding: '6px 12px', width: 100, verticalAlign: 'top' }}>
                                      <div style={{ display: 'flex', gap: 4 }}>
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
                                    </td>
                                  </tr>
                                  );
                                })}
                              </tbody>
                            </table>
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
      </div>

      <div style={card}>
        <p style={{ fontSize: 15, fontWeight: 700, margin: '0 0 14px', color: C.black, fontFamily: FONT }}>Listino prezzi aziendale</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 14, marginBottom: 14 }}>
          <MetricCard label="Articoli" value={String(totals.voci)} />
          <MetricCard label="Categorie" value={String(totals.categorie)} />
          <MetricCard label="Sottocategorie" value={String(totals.sottocategorie)} />
        </div>
        <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>
          Con le frecce ▲▼ puoi riordinare manualmente macrocategorie, categorie, sottocategorie e singole voci: i codici di macrocategoria, categoria e sottocategoria si aggiornano da soli in base al nuovo ordine.
        </p>
      </div>

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
    </div>
  );
}

function NewProjectModal({ onClose, onCreate }) {
  const [name, setName] = useState('');
  const [client, setClient] = useState('');

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
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
const SECTION_COLORS = [C.maroon, C.darkGray, '#94706C', '#4A3F35'];

function DraggableCatalogTree({ listino, onAdd }) {
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

function exportComputoExcel(project, revision, clientOnly) {
  const groups = buildComputoGroups(revision);
  const rows = [];
  rows.push([clientOnly ? 'Computo metrico (versione cliente)' : 'Computo metrico']);
  rows.push(['Progetto', project.name]);
  rows.push(['Cliente', project.client]);
  rows.push(['Versione', revision.customName || revision.label]);
  rows.push(['Data modifica', revision.dateModified]);
  rows.push([]);
  rows.push(clientOnly
    ? ['Sezione', 'Categoria', 'Codice', 'Descrizione', 'Quantità', 'U.M.', 'Prezzo cliente', 'Totale cliente']
    : ['Sezione', 'Categoria', 'Codice', 'Descrizione', 'Quantità', 'U.M.', 'Prezzo impresa', 'Totale impresa', 'Prezzo cliente', 'Totale cliente']);

  groups.forEach((g) => {
    let runImpresa = 0;
    let runCliente = 0;
    g.categorie.forEach((cat) => {
      cat.items.forEach((it) => {
        const qty = parseEuro(it.qty);
        const totImpresa = parseEuro(it.unitPriceImpresa) * qty;
        const totCliente = parseEuro(it.unitPriceCliente) * qty;
        runImpresa += totImpresa;
        runCliente += totCliente;
        rows.push(clientOnly
          ? [g.name, cat.name, it.code, it.desc, it.qty, it.unit, it.unitPriceCliente, formatEuro(totCliente)]
          : [g.name, cat.name, it.code, it.desc, it.qty, it.unit, it.unitPriceImpresa, formatEuro(totImpresa), it.unitPriceCliente, formatEuro(totCliente)]);
      });
    });
    g.subtotalMarkers.forEach((it) => {
      const hasVat = it.vatRate !== null && it.vatRate !== undefined;
      if (hasVat) {
        if (!clientOnly) {
          rows.push([g.name, '', '', `${it.title} — IVA esclusa`, '', '', '', formatEuro(runImpresa)]);
          rows.push([g.name, '', '', `${it.title} — ${it.vatLabel}`, '', '', '', formatEuro(runImpresa * (it.vatRate / 100))]);
          rows.push([g.name, '', '', `${it.title} — IVA inclusa`, '', '', '', formatEuro(runImpresa * (1 + it.vatRate / 100))]);
        } else {
          rows.push([g.name, '', '', `${it.title} — IVA esclusa`, '', '', '', formatEuro(runCliente)]);
          rows.push([g.name, '', '', `${it.title} — ${it.vatLabel}`, '', '', '', formatEuro(runCliente * (it.vatRate / 100))]);
          rows.push([g.name, '', '', `${it.title} — IVA inclusa`, '', '', '', formatEuro(runCliente * (1 + it.vatRate / 100))]);
        }
      } else {
        rows.push([g.name, '', '', `— ${it.title} —`]);
      }
      runImpresa = 0;
      runCliente = 0;
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

function PrintableComputo({ project, revision, clientOnly, studioSettings }) {
  const groups = buildComputoGroups(revision);
  const realItems = (revision.items || []).filter((it) => it.type !== 'subtotal');
  const impresaTot = sumImpresa(realItems);
  const clienteTot = sumCliente(realItems);
  const header = project.header || {};
  const { rate: vatRate, label: vatLabel } = getVatInfo(revision);
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
  const effectiveFont = hasCustomFont ? `'StudioFontPersonalizzato', ${FONT}` : FONT;
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

      <h1 style={{ fontSize: 18, fontWeight: 600, margin: '0 0 8px', color: '#1A1A1A' }}>{clientOnly ? 'Computo metrico — versione cliente' : 'Computo metrico'}</h1>
      <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>Progetto: {project.name} — Cliente: {project.client}</p>
      <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>Versione: {revision.customName || revision.label} — Modificata il {revision.dateModified}</p>
      {(header.descrizione || header.ubicazione) && (
        <p style={{ fontSize: 10.5, margin: '2px 0', color: '#5A5A5A' }}>{header.descrizione} {header.ubicazione && `— ${header.ubicazione}`}</p>
      )}
      {project.clientSheet && (
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
      {groups.map((g, gi) => {
        let runI = 0;
        let runC = 0;
        const thStyle = { textAlign: 'left', borderBottom: '1.5px solid #1A1A1A', padding: '4px 4px', fontSize: 8, textTransform: 'uppercase', letterSpacing: 0.5, color: '#8A8A8A', fontWeight: 600 };
        return (
        <div key={gi} style={{ marginTop: 24 }}>
          <h3 style={{
            fontSize: 11.5, fontWeight: 700, margin: '0 0 9px', padding: '1px 0 1px 10px',
            borderLeft: `3px solid ${C.maroon}`, textTransform: 'uppercase', letterSpacing: 0.6,
            color: '#1A1A1A', pageBreakAfter: 'avoid', breakAfter: 'avoid',
          }}>{g.name}</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 10 }}>
            <thead>
              <tr>
                <th style={thStyle}>Codice</th>
                <th style={thStyle}>Descrizione</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Qtà</th>
                <th style={thStyle}>U.M.</th>
                {!clientOnly && <th style={{ ...thStyle, textAlign: 'right' }}>Prezzo impresa</th>}
                {!clientOnly && <th style={{ ...thStyle, textAlign: 'right' }}>Totale impresa</th>}
                <th style={{ ...thStyle, textAlign: 'right' }}>Prezzo cliente</th>
                <th style={{ ...thStyle, textAlign: 'right' }}>Totale cliente</th>
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
                      <td style={{ padding: '5px 4px', color: '#5A5A5A' }}>{it.code}</td>
                      <td style={{ padding: '5px 4px' }}>{it.desc}</td>
                      <td style={{ padding: '5px 4px', textAlign: 'right' }}>{it.qty}</td>
                      <td style={{ padding: '5px 4px', color: '#8A8A8A' }}>{it.unit}</td>
                      {!clientOnly && <td style={{ padding: '5px 4px', textAlign: 'right', color: '#5A5A5A' }}>{it.unitPriceImpresa} €</td>}
                      {!clientOnly && <td style={{ padding: '5px 4px', textAlign: 'right', fontWeight: 600 }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>}
                      <td style={{ padding: '5px 4px', textAlign: 'right', color: C.maroon }}>{it.unitPriceCliente} €</td>
                      <td style={{ padding: '5px 4px', textAlign: 'right', fontWeight: 600, color: C.maroon }}>{formatEuro(parseEuro(it.unitPriceCliente) * parseEuro(it.qty))}</td>
                    </tr>
                    );
                  })}
                </React.Fragment>
              ))}
              {g.subtotalMarkers.map((it, ii) => {
                const hasVat = it.vatRate !== null && it.vatRate !== undefined;
                const rowEl = (
                  <tr key={ii} style={{ breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                    <td colSpan={clientOnly ? 6 : 8} style={{ padding: '8px 4px', fontWeight: 700, borderTop: '1px solid #E5E2DA' }}>
                      {it.title}
                      {hasVat && (
                        <span style={{ fontWeight: 400, color: '#8A8A8A' }}>
                          {' '}— IVA esclusa {formatEuro(clientOnly ? runC : runI)}, {it.vatLabel} {formatEuro((clientOnly ? runC : runI) * (it.vatRate / 100))}, IVA inclusa {formatEuro((clientOnly ? runC : runI) * (1 + it.vatRate / 100))}
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
        </div>
        );
      })}
      <div style={{ marginTop: 30, maxWidth: 320, marginLeft: 'auto', breakInside: 'avoid', pageBreakInside: 'avoid' }}>
        {!clientOnly && (
          <>
            <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>Totale generale IVA esclusa (impresa)</span><span>{formatEuro(impresaTot)}</span></p>
            <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>{vatLabel}</span><span>{formatEuro(impresaTot * (vatRate / 100))}</span></p>
            <p style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between', fontWeight: 700, margin: '6px 0 0', paddingTop: 6, borderTop: '1px solid #E5E2DA' }}><span>Totale generale IVA inclusa (impresa)</span><span>{formatEuro(impresaTot * (1 + vatRate / 100))}</span></p>
            <div style={{ height: 16 }} />
          </>
        )}
        <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>Totale generale IVA esclusa (cliente)</span><span>{formatEuro(clienteTot)}</span></p>
        <p style={{ fontSize: 10.5, display: 'flex', justifyContent: 'space-between', color: '#5A5A5A', margin: '4px 0' }}><span>{vatLabel}</span><span>{formatEuro(clienteTot * (vatRate / 100))}</span></p>
        <p style={{ fontSize: 13, display: 'flex', justifyContent: 'space-between', fontWeight: 700, color: C.maroon, margin: '6px 0 0', paddingTop: 6, borderTop: `1.5px solid ${C.maroon}` }}><span>Totale generale IVA inclusa (cliente)</span><span>{formatEuro(clienteTot * (1 + vatRate / 100))}</span></p>
      </div>

          </td></tr>
        </tbody>
      </table>
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
  const [dragOverTarget, setDragOverTarget] = useState(null); // "macroName", "macroName|categoriaName" o "macroName|categoriaName|sottoName" evidenziato durante il drag
  const [showImportPdf, setShowImportPdf] = useState(false);
  const [voceComputoCtx, setVoceComputoCtx] = useState(null); // { macroName, categoriaName, sottoName, initialItem }
  const [expandedItems, setExpandedItems] = useState({}); // { [itemId]: true } — dettaglio misurazioni/note aperto

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
  const { rate: vatRate, label: vatLabel } = getVatInfo(selectedRevision);

  // Raggruppamento a tre livelli (macrocategoria > categoria > sottocategoria), calcolato da
  // buildComputoGroups (condiviso anche con la stampa PDF e l'export Excel, così restano sempre
  // coerenti con quanto mostrato qui a schermo).
  const groupedSections = buildComputoGroups(selectedRevision);
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
  // Ricalcola sempre i codici delle voci col codice automatico, in modo che qualunque spostamento
  // (voce, sottocategoria o macrocategoria) tenga i codici coerenti con la posizione attuale.
  const applyRevisionChange = (updater) => {
    const patch = updater(selectedRevision);
    const nextItems = regenerateItemCodes(patch.items || items);
    const patchWithCodes = patch.items ? { ...patch, items: nextItems } : patch;
    const realNextItems = nextItems.filter((it) => it.type !== 'subtotal');
    const total = formatEuro(sumImpresa(realNextItems));
    const totalCliente = formatEuro(sumCliente(realNextItems));
    if (isEditingLatest) {
      const updatedRevisions = revisions.map((r) => (r.id === selectedRevision.id ? { ...r, ...patchWithCodes, dateModified: nowLabel(), total, totalCliente } : r));
      onUpdateProject({ ...project, revisions: updatedRevisions, value: total });
    } else {
      const newRev = { ...selectedRevision, ...patchWithCodes, id: Date.now(), label: `Revisione ${revisions.length + 1}`, customName: null, dateCreated: nowLabel(), dateModified: nowLabel(), status: STATUS_OPTIONS[0], total, totalCliente };
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

  // Aggiungere una voce dal Listino (trascinandola o toccando +) apre lo stesso pannello di misurazione
  // delle voci create da zero, precompilato con descrizione/unità/prezzi del listino: la modalità di calcolo
  // (ml/m²/m³, oppure quantità manuale per "a corpo"/"cadauna"/ecc.) si sceglie in base all'unità di misura.
  // La voce finisce come voce vera e propria nella macrocategoria/sottocategoria del listino di provenienza.
  const openVoceFromListino = (voce) => {
    setVoceComputoCtx({
      macroName: voce.macro || 'Voci varie',
      categoriaName: voce.categoria || 'Generale',
      sottoName: voce.sotto || 'Generale',
      initialItem: null,
      prefill: { desc: voce.desc, unit: voce.unit, priceImpresa: voce.priceImpresa, priceCliente: voce.priceCliente, note: voce.note },
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
      prefill: { desc: voce.desc, unit: voce.unit, priceImpresa: voce.priceImpresa, priceCliente: voce.priceCliente, note: voce.note },
    });
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
    if (!confirm('Eliminare questa voce?')) return;
    applyItemsChange((its) => its.filter((it) => it.id !== id));
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
    applyRevisionChange((rev) => ({
      items: (rev.items || []).map((it) => ((it.section || it.macro) === oldName ? { ...it, section: newName } : it)),
      extraSections: (rev.extraSections || []).map((n) => (n === oldName ? newName : n)),
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
      desc: voceData.desc, unit: voceData.unit, note: voceData.note || '',
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
          ...it, ...priceFields, unitaCalcolo: voceData.unitaCalcolo, misurazioni: voceData.misurazioni, qty,
        } : it));
      }
      const newItem = {
        id: Date.now() + Math.random(), code: '', autoCode: true, ...priceFields,
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
    supabase.from('app_state').select('data').eq('id', 1).maybeSingle().then(({ data, error }) => {
      setLoadingDsmClients(false);
      if (error) { setDsmClientsError('Impossibile leggere l\'anagrafica clienti da Desearq Studio Manager: ' + error.message); return; }
      setDsmClients((data?.data && data.data.clients) || []);
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

  const addTeamMember = () => {
    const name = prompt('Nome della persona da assegnare:');
    if (!name) return;
    const role = prompt('Ruolo:', 'Collaboratore') || 'Collaboratore';
    onUpdateProject({ ...project, team: [...(project.team || []), { name, role }] });
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

  const addFatturaCategoria = (categoria) => {
    const label = prompt('Descrizione fattura (es. "Fattura n.1", "SAL 1 lavori"):');
    if (!label) return;
    const importo = parseEuro(prompt('Importo (€):', '0') || '0');
    const current = cantiereSal[categoria] || { impresa: '', fatture: [] };
    const fattura = { id: Date.now(), label, importo, stato: 'Emessa' };
    onUpdateProject({ ...project, cantiereSal: { ...cantiereSal, [categoria]: { ...current, fatture: [...(current.fatture || []), fattura] } } });
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
            <button onClick={() => setShowImportPdf(true)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>📄 Nuova revisione da PDF</button>
            <button onClick={() => exportComputoExcel(project, selectedRevision, false)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬚ Scarica Excel completo</button>
            <button onClick={() => exportComputoExcel(project, selectedRevision, true)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬚ Scarica Excel solo cliente</button>
            <button onClick={() => requestPdf(project, selectedRevision, false)} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>⬇ Scarica PDF completo (impresa+cliente)</button>
            <button onClick={() => requestPdf(project, selectedRevision, true)} style={{ background: C.maroon, border: 'none', borderRadius: 999, padding: '9px 14px', fontSize: 12, fontWeight: 600, color: C.white, cursor: 'pointer' }}>⬇ Scarica PDF solo cliente</button>
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
            <div style={{ ...card, width: 300, maxWidth: '100%', flexShrink: 0, border: `2px solid ${C.maroon}`, position: 'sticky', top: 16, alignSelf: 'flex-start', maxHeight: 'calc(100vh - 32px)', display: 'flex', flexDirection: 'column' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4, flexShrink: 0 }}>
                <p style={{ fontSize: 14, fontWeight: 700, margin: 0, color: C.black, fontFamily: FONT }}>Listino</p>
              </div>
              <select value={listinoId} onChange={(e) => setListinoId(Number(e.target.value))} style={{ width: '100%', fontSize: 12, padding: '8px 10px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '6px 0 10px', flexShrink: 0 }}>
                {listini.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
              </select>
              <p style={{ fontSize: 11, color: C.gray, margin: '0 0 10px', flexShrink: 0 }}>Apri le categorie per trovare la voce giusta: trascinala nel computo a destra, oppure tocca + per aggiungerla subito (utile su tablet e smartphone).</p>
              <div style={{ flex: 1, minHeight: 0, overflowY: 'auto' }}>
                <DraggableCatalogTree listino={activeListino} onAdd={openVoceFromListino} />
              </div>
            </div>

            <div style={{ flex: 1, minWidth: 320 }}>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                <button onClick={addCustomSection} style={{ background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '7px 12px', fontSize: 12, fontWeight: 600, color: C.black, cursor: 'pointer' }}>+ Nuova macrosezione</button>
              </div>
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
                                      const detailRows = (it.misurazioni || []).flatMap((g) => g.rows || []);
                                      const hasDetail = detailRows.length > 0 || !!it.note;
                                      const isExpanded = !!expandedItems[it.id];
                                      return (
                                      <React.Fragment key={it.id}>
                                      <tr style={{ borderTop: `1px solid ${C.paleGray}` }}>
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
                                          </div>
                                        </td>
                                        <td style={{ padding: '8px 6px', fontWeight: 700, color: C.black }}>{it.code}</td>
                                        <td style={{ padding: '8px 6px', color: C.midGray }}>
                                          {it.desc}
                                          {it.note && <p style={{ margin: '3px 0 0', fontSize: 10.5, color: C.gray, fontStyle: 'italic' }}>📝 {it.note}</p>}
                                        </td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right' }}>
                                          {it.autoCode ? (
                                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                              <span style={{ fontWeight: 700 }}>{it.qty}</span>
                                              <button onClick={() => setVoceComputoCtx({ macroName: section.name, categoriaName: cat.name, sottoName: it.sottocategoria || 'Generale', initialItem: it })} style={{ ...iconBtn, width: 20, height: 20, fontSize: 10 }}>✎</button>
                                            </span>
                                          ) : (
                                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                              <input
                                                value={it.qty}
                                                onChange={(e) => updateQty(it.id, e.target.value)}
                                                style={{ width: 60, fontSize: 12, padding: '5px 6px', borderRadius: 6, border: `1px solid ${C.paleGray}`, textAlign: 'right' }}
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
                                        <td style={{ padding: '8px 6px', textAlign: 'right' }}>{it.unitPriceImpresa} €</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: C.black }}>{formatEuro(parseEuro(it.unitPriceImpresa) * parseEuro(it.qty))}</td>
                                        <td style={{ padding: '8px 6px', textAlign: 'right', color: C.maroon }}>{it.unitPriceCliente} €</td>
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
                                                        <td style={{ padding: '3px 6px', textAlign: 'right', fontWeight: 700 }}>{computeMisurazioneRowValue(r, it.unitaCalcolo).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
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

                        {markerRows.length > 0 && (
                          <div style={{ borderTop: `1px solid ${C.paleGray}` }}>
                            {markerRows.map(({ marker, runImp, runCli }) => {
                              const hasVat = marker.vatRate !== null && marker.vatRate !== undefined;
                              const ivaImpresaPart = hasVat ? runImp * (marker.vatRate / 100) : 0;
                              const ivaClientePart = hasVat ? runCli * (marker.vatRate / 100) : 0;
                              return (
                                <div key={marker.id} style={{ padding: '8px 14px', background: 'rgba(128,20,48,0.05)', borderTop: `2px solid ${C.paleGray}`, display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
                                  <div style={{ fontWeight: 700, color: C.maroon, fontSize: 12 }}>
                                    {marker.title}
                                    <button onClick={() => editPartialSubtotal(marker.id)} style={{ ...rowBtnStyle, marginLeft: 8, padding: '1px 6px', fontSize: 10 }}>✎</button>
                                  </div>
                                  <div style={{ fontSize: 12, textAlign: 'right' }}>
                                    {hasVat ? (
                                      <>
                                        <div>Impresa — IVA escl. {formatEuro(runImp)}, {marker.vatLabel} {formatEuro(ivaImpresaPart)}, IVA incl. {formatEuro(runImp + ivaImpresaPart)}</div>
                                        <div style={{ color: C.maroon }}>Cliente — IVA escl. {formatEuro(runCli)}, {marker.vatLabel} {formatEuro(ivaClientePart)}, IVA incl. {formatEuro(runCli + ivaClientePart)}</div>
                                      </>
                                    ) : (
                                      <div>Impresa {formatEuro(runImp)} · <span style={{ color: C.maroon }}>Cliente {formatEuro(runCli)}</span></div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
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
                {voceComputoCtx && (
                  <VoceComputoModal
                    macroName={voceComputoCtx.macroName}
                    categoriaName={voceComputoCtx.categoriaName}
                    sottoName={voceComputoCtx.sottoName}
                    initialItem={voceComputoCtx.initialItem}
                    prefill={voceComputoCtx.prefill}
                    mergeCandidates={(groupedSections.find((s) => s.name === voceComputoCtx.macroName)?.categorie.find((c) => c.name === (voceComputoCtx.categoriaName || 'Generale'))?.items || []).filter((it) => it.autoCode && it.id !== voceComputoCtx.initialItem?.id)}
                    onClose={() => setVoceComputoCtx(null)}
                    onSave={(voceData) => saveVoceComputo(voceComputoCtx.macroName, voceComputoCtx.categoriaName || 'Generale', voceComputoCtx.sottoName, voceData)}
                  />
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
                      <button onClick={() => addFatturaCategoria(s.name)} style={rowBtnStyle}>+ Fattura</button>
                    </div>
                  </div>

                  {fatture.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      {fatture.map((f) => (
                        <div key={f.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0', gap: 8, flexWrap: 'wrap' }}>
                          <span style={{ fontSize: 12, color: C.black }}>{f.label} — {formatEuro(f.importo)}</span>
                          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
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
                    <span style={{ color: C.gray }}>Fatturato: <strong style={{ color: C.black }}>{formatEuro(fatturato)}</strong></span>
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
        {(project.team || []).length === 0 && <p style={{ fontSize: 12, color: C.gray, margin: 0 }}>Nessuna persona assegnata a questo progetto.</p>}
        {(project.team || []).map((m, i) => (
          <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: i < (project.team || []).length - 1 ? `1px solid ${C.paleGray}` : 'none' }}>
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
    <div className="table-scroll">
    <table style={{ width: '100%', minWidth: 720, borderCollapse: 'collapse', fontSize: 12 }}>
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
    </div>
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
              borderRadius: 8, padding: '9px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
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
};

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

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10, padding: 16 }}>
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
  );
}

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
        <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 380, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
          <p style={{ fontSize: 13, color: C.gray, margin: '0 0 16px' }}>Nessun progetto ha ancora un computo creato. Apri un progetto e crea il primo computo prima di aggiungere forniture.</p>
          <button onClick={onClose} style={{ background: C.darkGray, color: C.white, border: 'none', padding: '9px 14px', borderRadius: 999, fontSize: 12, fontWeight: 600 }}>Chiudi</button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(5,5,5,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10 }}>
      <div style={{ background: C.white, borderRadius: 14, padding: 22, width: 420, maxWidth: 'calc(100vw - 32px)', maxHeight: '90vh', overflowY: 'auto' }}>
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
          <div className="table-scroll">
            <div style={{ minWidth: 420 }}>
              <div style={{ display: 'flex', flexWrap: 'nowrap', fontSize: 10, color: C.gray, textTransform: 'uppercase', padding: '2px 0' }}>
                <span style={{ flex: 1, minWidth: 120 }}>Fornitore</span>
                <span style={{ width: 70, textAlign: 'right', flexShrink: 0 }}>Listino</span>
                <span style={{ width: 70, textAlign: 'right', flexShrink: 0 }}>Scontato</span>
                <span style={{ width: 70, textAlign: 'right', flexShrink: 0 }}>Cliente</span>
                <span style={{ width: 56, flexShrink: 0 }}></span>
              </div>
              {[...prodotto.fornitori].sort((a, b) => parseEuro(a.prezzoScontato) - parseEuro(b.prezzoScontato)).map((f) => (
                <div key={f.id} style={{ display: 'flex', flexWrap: 'nowrap', alignItems: 'center', padding: '6px 0', borderTop: `1px solid ${C.paleGray}` }}>
                  <div style={{ flex: 1, minWidth: 120, display: 'flex', flexWrap: 'nowrap', alignItems: 'center', gap: 8 }}>
                    <span style={{ fontSize: 12, color: C.black, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{f.name}</span>
                    {parseEuro(f.prezzoScontato) === cheapestScontato && (
                      <span style={{ fontSize: 10, fontWeight: 700, color: C.maroon, background: 'rgba(128,20,48,0.1)', padding: '2px 7px', borderRadius: 999, whiteSpace: 'nowrap', flexShrink: 0 }}>Migliore offerta</span>
                    )}
                  </div>
                  <span style={{ width: 70, textAlign: 'right', fontSize: 12, color: C.gray, flexShrink: 0 }}>{f.prezzoListino} €</span>
                  <span style={{ width: 70, textAlign: 'right', fontSize: 12, fontWeight: 700, color: C.black, flexShrink: 0 }}>{f.prezzoScontato} €</span>
                  <span style={{ width: 70, textAlign: 'right', fontSize: 12, color: C.maroon, flexShrink: 0 }}>{f.prezzoCliente} €</span>
                  <span style={{ width: 56, display: 'flex', flexWrap: 'nowrap', gap: 4, justifyContent: 'flex-end', flexShrink: 0 }}>
                    <button onClick={() => editFornitore(f.id)} style={{ ...rowBtnStyle, padding: '2px 6px' }}>✎</button>
                    <button onClick={() => removeFornitore(f.id)} style={{ ...rowBtnStyle, padding: '2px 6px' }}>🗑</button>
                  </span>
                </div>
              ))}
            </div>
          </div>
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
    // La voce va nella categoria del fornitore scelto come fonte prezzo (non in un generico "FORNITURE"):
    // cosi' ogni categoria del computo raggruppa le voci di un solo fornitore, ed e' collegabile a
    // un'unica impresa/fattura in "Stati avanzamento pagamenti" senza mischiare fornitori diversi.
    // Solo se si sceglie il prezzo di listino del prodotto (nessun fornitore reale) si resta su "FORNITURE".
    const sectionName = priceSource.id !== 'listino' ? priceSource.name : 'FORNITURE';
    const newItem = {
      id: Date.now() + Math.random(),
      code: '',
      desc: prodotto.name,
      unit: 'cad',
      unitPriceImpresa: priceSource.prezzoScontato,
      unitPriceCliente: priceSource.prezzoCliente,
      listinoRef: priceSource.prezzoListino,
      qty,
      macro: sectionName,
      section: sectionName,
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

function TeamPage({ profile }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState('');
  const [name, setName] = useState('');
  const [password, setPassword] = useState('');
  const [role, setRole] = useState('Membro');
  const [error, setError] = useState('');
  const [creating, setCreating] = useState(false);
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
  // (funzione Edge "cea-admin-create-user", con service role) senza inviare alcuna email di conferma —
  // la persona può accedere da subito con "Accedi" usando queste credenziali.
  const handleCreate = async () => {
    const trimmedEmail = email.trim();
    if (!name.trim() || !trimmedEmail || !password) { setError('Nome, email e password sono obbligatori.'); return; }
    if (password.length < 6) { setError('La password deve avere almeno 6 caratteri.'); return; }
    setError(''); setCreating(true);
    const { data, error: err } = await supabase.functions.invoke('cea-admin-create-user', {
      body: { action: 'create', email: trimmedEmail, password, name: name.trim(), role },
    });
    setCreating(false);
    if (err) { setError(await extractFunctionErrorMessage(err, 'Creazione non riuscita.')); return; }
    if (data?.error) { setError(data.error); return; }
    setEmail(''); setName(''); setPassword(''); setRole('Membro');
    loadMembers();
  };

  const handleResetPassword = async (m) => {
    const newPassword = prompt(`Nuova password per ${m.name} (${m.email}):`);
    if (!newPassword) return;
    if (newPassword.length < 6) { alert('La password deve avere almeno 6 caratteri.'); return; }
    const { data, error: err } = await supabase.functions.invoke('cea-admin-create-user', {
      body: { action: 'reset_password', memberId: m.id, password: newPassword },
    });
    if (err) { alert(await extractFunctionErrorMessage(err, 'Operazione non riuscita.')); return; }
    if (data?.error) { alert(data.error); return; }
    alert('Password aggiornata: comunicala alla persona di persona o in chat privata, non via email.');
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
          <p style={{ fontSize: 12, color: C.black, margin: 0 }}>Sei collegato come <strong>membro</strong>: solo l'admin può creare o rimuovere accessi al team.</p>
        </div>
      )}

      {isAdmin && (
        <div style={{ ...card, marginBottom: 18 }}>
          <h2 style={{ fontSize: 18, margin: '0 0 12px', color: C.black, fontFamily: FONT }}>Crea un nuovo accesso</h2>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome" style={{ flex: 1, minWidth: 120, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, background: C.bg }} />
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="nome@studio.it"
              style={{ flex: 1, minWidth: 180, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, background: C.bg }}
            />
            <input
              type="text"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Password (min. 6 caratteri)"
              style={{ flex: 1, minWidth: 160, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, background: C.bg }}
            />
            <select value={role} onChange={(e) => setRole(e.target.value)} style={{ fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}` }}>
              <option>Membro</option>
              <option>Admin</option>
            </select>
            <button
              onClick={handleCreate}
              disabled={creating}
              style={{ background: C.maroon, color: C.white, border: 'none', padding: '9px 16px', borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: creating ? 'default' : 'pointer', opacity: creating ? 0.7 : 1 }}
            >
              {creating ? 'Un attimo…' : 'Crea accesso'}
            </button>
          </div>
          {error && <p style={{ fontSize: 12, color: C.maroon, margin: '10px 0 0' }}>{error}</p>}
          <p style={{ fontSize: 11, color: C.gray, margin: '10px 0 0' }}>
            L'account viene attivato subito: non parte nessuna email di conferma. Comunica tu stesso email e password alla persona (di persona, in chat privata…) — potrà accedere subito da "Accedi" con queste credenziali, senza passare da "Crea un account".
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
              {isAdmin && m.auth_user_id && (
                <button onClick={() => handleResetPassword(m)} style={rowBtnStyle}>🔑 Reimposta password</button>
              )}
              {isAdmin && m.id !== profile.id && (
                <button onClick={() => removeMember(m.id)} style={{ ...rowBtnStyle, color: C.maroon }}>🗑</button>
              )}
            </div>
          </div>
        ))}
      </div>

      <div style={{ background: 'rgba(128,20,48,0.06)', border: '1px solid rgba(128,20,48,0.18)', borderRadius: 10, padding: 14, fontSize: 12, color: C.midGray }}>
        <strong style={{ color: C.black }}>Come funziona la condivisione.</strong> Tutti i membri con un accesso creato qui entrano nello stesso workspace e vedono gli stessi progetti, computi e listino con un vero account: i dati restano salvati per sempre e sono visibili a tutto il team.
      </div>
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

  const labelStyle = { fontSize: 11, fontWeight: 700, color: C.midGray };
  const fieldStyle = { width: '100%', fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${C.paleGray}`, margin: '4px 0 14px', background: C.bg };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginBottom: 20 }}>
        <h1 style={h1Style}>Impostazioni studio</h1>
        <span style={freshBadge}>Workspace condiviso</span>
      </div>

      <div style={{ ...card, marginBottom: 18 }}>
        <h2 style={{ fontSize: 18, margin: '0 0 4px', color: C.black, fontFamily: FONT }}>Dati dello studio</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 16px' }}>Questi dati compaiono automaticamente in intestazione e piè di pagina di ogni computo metrico stampato in PDF, a meno di caricare un'intestazione/piè di pagina personalizzata qui sotto.</p>

        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
          <div style={{ flex: '1 1 260px' }}>
            <label style={labelStyle}>Nome studio</label>
            <input value={settings.nome} onChange={set('nome')} placeholder="Desearq Studio" style={fieldStyle} />
          </div>
          <div style={{ flex: '1 1 260px' }}>
            <label style={labelStyle}>Indirizzo</label>
            <input value={settings.indirizzo} onChange={set('indirizzo')} placeholder="Via, numero civico, città" style={fieldStyle} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
          <div style={{ flex: '1 1 160px' }}>
            <label style={labelStyle}>P.IVA</label>
            <input value={settings.piva} onChange={set('piva')} style={fieldStyle} />
          </div>
          <div style={{ flex: '1 1 160px' }}>
            <label style={labelStyle}>Codice fiscale</label>
            <input value={settings.cf} onChange={set('cf')} style={fieldStyle} />
          </div>
          <div style={{ flex: '1 1 160px' }}>
            <label style={labelStyle}>Telefono</label>
            <input value={settings.telefono} onChange={set('telefono')} style={fieldStyle} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
          <div style={{ flex: '1 1 200px' }}>
            <label style={labelStyle}>Email</label>
            <input value={settings.email} onChange={set('email')} style={fieldStyle} />
          </div>
          <div style={{ flex: '1 1 200px' }}>
            <label style={labelStyle}>Sito web</label>
            <input value={settings.sito} onChange={set('sito')} style={fieldStyle} />
          </div>
        </div>

        <label style={labelStyle}>Logo studio</label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '4px 0 4px' }}>
          {settings.logo && <img src={settings.logo} alt="Logo studio" style={{ height: 44, borderRadius: 6, border: `1px solid ${C.paleGray}` }} />}
          <label style={{ ...rowBtnStyle, cursor: 'pointer' }}>
            {settings.logo ? 'Sostituisci logo' : 'Carica logo'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('logo', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.logo && <button onClick={() => onUpdate({ logo: null })} style={rowBtnStyle}>🗑 Rimuovi</button>}
        </div>
      </div>

      <div style={{ ...card, marginBottom: 18 }}>
        <h2 style={{ fontSize: 18, margin: '0 0 4px', color: C.black, fontFamily: FONT }}>Intestazione personalizzata</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 12px' }}>Carica un'immagine (es. la tua carta intestata già impaginata) da usare al posto dei dati anagrafici in cima a ogni pagina del PDF, dentro il margine della pagina.</p>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: C.black, marginBottom: 10, cursor: 'pointer' }}>
          <input type="checkbox" checked={settings.usaIntestazionePersonalizzata} onChange={(e) => onUpdate({ usaIntestazionePersonalizzata: e.target.checked })} />
          Usa l'immagine personalizzata invece dei dati anagrafici
        </label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: settings.intestazioneImg ? 14 : 0 }}>
          {settings.intestazioneImg && <img src={settings.intestazioneImg} alt="Intestazione personalizzata" style={{ maxHeight: 70, maxWidth: 260, borderRadius: 6, border: `1px solid ${C.paleGray}` }} />}
          <label style={{ ...rowBtnStyle, cursor: 'pointer' }}>
            {settings.intestazioneImg ? 'Sostituisci immagine' : 'Carica immagine'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('intestazioneImg', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.intestazioneImg && <button onClick={() => onUpdate({ intestazioneImg: null })} style={rowBtnStyle}>🗑 Rimuovi</button>}
        </div>
        {settings.intestazioneImg && (
          <div>
            <label style={labelStyle}>Dimensione immagine ({settings.intestazioneScala || 100}% della larghezza disponibile)</label>
            <input type="range" min="25" max="200" step="5" value={settings.intestazioneScala || 100}
              onChange={(e) => onUpdate({ intestazioneScala: Number(e.target.value) })} style={{ width: '100%', maxWidth: 320, display: 'block', margin: '4px 0 0' }} />
          </div>
        )}
      </div>

      <div style={{ ...card, marginBottom: 18 }}>
        <h2 style={{ fontSize: 18, margin: '0 0 4px', color: C.black, fontFamily: FONT }}>Piè di pagina</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 12px' }}>Un testo breve (es. dati di contatto o un promemoria legale) ripetuto in fondo a ogni pagina, oppure un'immagine personalizzata al posto del testo, anche questa dentro il margine della pagina.</p>

        <label style={labelStyle}>Testo piè di pagina</label>
        <input value={settings.testoPiePagina} onChange={set('testoPiePagina')} placeholder="es. Desearq Studio — Via Roma 1, Milano — info@desearq.com" style={fieldStyle} />

        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: C.black, marginBottom: 10, cursor: 'pointer' }}>
          <input type="checkbox" checked={settings.usaPiePersonalizzato} onChange={(e) => onUpdate({ usaPiePersonalizzato: e.target.checked })} />
          Usa un'immagine personalizzata invece del testo
        </label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: settings.pieImg ? 14 : 0 }}>
          {settings.pieImg && <img src={settings.pieImg} alt="Piè di pagina personalizzato" style={{ maxHeight: 50, maxWidth: 260, borderRadius: 6, border: `1px solid ${C.paleGray}` }} />}
          <label style={{ ...rowBtnStyle, cursor: 'pointer' }}>
            {settings.pieImg ? 'Sostituisci immagine' : 'Carica immagine'}
            <input type="file" accept="image/*" onChange={(e) => uploadImage('pieImg', e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.pieImg && <button onClick={() => onUpdate({ pieImg: null })} style={rowBtnStyle}>🗑 Rimuovi</button>}
        </div>
        {settings.pieImg && (
          <div>
            <label style={labelStyle}>Dimensione immagine ({settings.pieScala || 100}% della larghezza disponibile)</label>
            <input type="range" min="25" max="200" step="5" value={settings.pieScala || 100}
              onChange={(e) => onUpdate({ pieScala: Number(e.target.value) })} style={{ width: '100%', maxWidth: 320, display: 'block', margin: '4px 0 0' }} />
          </div>
        )}
      </div>

      <div style={{ ...card }}>
        <h2 style={{ fontSize: 18, margin: '0 0 4px', color: C.black, fontFamily: FONT }}>Font personalizzato</h2>
        <p style={{ fontSize: 12, color: C.gray, margin: '0 0 12px' }}>Carica il font del tuo studio (.woff2, .woff, .ttf o .otf) da usare nel testo dei computi metrici stampati in PDF, al posto del font di sistema.</p>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: C.black, marginBottom: 10, cursor: 'pointer' }}>
          <input type="checkbox" checked={settings.usaFontPersonalizzato} disabled={!settings.fontPersonalizzato} onChange={(e) => onUpdate({ usaFontPersonalizzato: e.target.checked })} />
          Usa il font personalizzato nel PDF invece del font di sistema
        </label>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          {settings.fontPersonalizzatoNome && <span style={{ fontSize: 12, color: C.darkGray }}>{settings.fontPersonalizzatoNome}</span>}
          <label style={{ ...rowBtnStyle, cursor: 'pointer' }}>
            {settings.fontPersonalizzato ? 'Sostituisci font' : 'Carica font'}
            <input type="file" accept=".woff,.woff2,.ttf,.otf" onChange={(e) => uploadFont(e.target.files[0])} style={{ display: 'none' }} />
          </label>
          {settings.fontPersonalizzato && <button onClick={() => onUpdate({ fontPersonalizzato: null, fontPersonalizzatoNome: '', usaFontPersonalizzato: false })} style={rowBtnStyle}>🗑 Rimuovi</button>}
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
        .dc-login { --paper:#f7f5ef; --ink:#090909; --muted:#77756f; --line:#d8d4cc; --bordeaux:#6e2635; --dark:#171411; }
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
    if (err) { setError('Link o PIN non validi. Controlla con lo studio.'); return; }
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
    const items = [];
    (data.macros || []).forEach((m) => (m.categorie || []).forEach((c) => (c.sottocategorie || []).forEach((s) => (s.voci || []).forEach((v) => {
      const val = values[v.code];
      if (val && String(val.impresa).trim() !== '') {
        items.push({ code: v.code, desc: v.desc, costoImpresa: String(val.impresa).trim(), note: (val.note || '').trim() });
      }
    }))));
    if (items.length === 0) { alert('Inserisci almeno un costo impresa prima di inviare.'); return; }
    setSubmitting(true);
    const { error: err } = await cea.rpc('fornitore_submit', { p_token: token, p_pin: pin.trim(), p_items: items });
    setSubmitting(false);
    if (err) { alert('Invio non riuscito: ' + err.message); return; }
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

  const [page, setPage] = useState('dashboard');
  const [projects, setProjects] = useState(PROJECTS);
  const [selectedProjectId, setSelectedProjectId] = useState(null);
  const [openRevisionId, setOpenRevisionId] = useState(null);
  const [listini, setListini] = useState([{ id: 1, name: 'Listino standard 2026', macros: INITIAL_MACROS }]);
  const [activeListinoId, setActiveListinoId] = useState(1);
  const [fornitoriCatalog, setFornitoriCatalog] = useState(INITIAL_FORNITORI);
  const [studioSettings, setStudioSettings] = useState(DEFAULT_STUDIO_SETTINGS);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

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
      if (saved.studioSettings) setStudioSettings({ ...DEFAULT_STUDIO_SETTINGS, ...saved.studioSettings });
      setDataLoaded(true);
    });
  }, [profile]);

  React.useEffect(() => {
    if (!dataLoaded || !profile || profile.status !== 'Attivo') return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      cea.from('app_state').update({
        data: { projects, listini, activeListinoId, fornitoriCatalog, studioSettings },
        updated_at: new Date().toISOString(),
        updated_by: authUser?.id,
      }).eq('id', 1).then(({ error }) => { if (error) console.error('Salvataggio fallito:', error.message); });
    }, 900);
    return () => clearTimeout(saveTimer.current);
  }, [projects, listini, activeListinoId, fornitoriCatalog, studioSettings, dataLoaded]);

  const openProject = (id) => { setSelectedProjectId(id); setOpenRevisionId(null); setPage('progetto-dettaglio'); };
  const openRevisionInProject = (projectId, revisionId) => { setSelectedProjectId(projectId); setOpenRevisionId(revisionId); setPage('progetto-dettaglio'); };
  const selectedProject = projects.find((p) => p.id === selectedProjectId);

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
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        html, body, #root { max-width: 100%; overflow-x: clip; }
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
        .sidebar-item-btn:not(.active):hover { background: ${C.sidebarHover} !important; color: ${C.black} !important; }
        .btn-accent-pill:hover { background: #650F26 !important; }
        .input-focus:focus { outline: none; border-color: ${C.maroon} !important; box-shadow: 0 0 0 3px rgba(128,20,48,0.25); }

        /* --- Rendering mobile --- */
        input, select, textarea, button { max-width: 100%; }
        .table-scroll { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }
        .table-scroll > table { min-width: 620px; }
        .mobile-menu-btn { display: none; }
        .sidebar-backdrop { display: none; }
        .form-grid { display: grid; }

        /* Su schermi stretti, qualunque riga flessibile può andare a capo invece di traboccare */
        @media (max-width: 760px) {
          [style*="display:flex"], [style*="display: flex"] { flex-wrap: wrap; min-width: 0; }
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
      <div className="no-print app-shell" style={{ display: 'flex', minHeight: '100vh', fontFamily: FONT, background: PAGE_GRADIENT, backgroundAttachment: 'fixed' }}>
      <div className={`sidebar-backdrop${mobileNavOpen ? ' open' : ''}`} onClick={() => setMobileNavOpen(false)} />
      <aside className={`sidebar-aside${mobileNavOpen ? ' open' : ''}`} style={{ width: 280, flexShrink: 0, background: C.sidebar, color: C.darkGray, borderRight: `1px solid ${C.paleGray}`, display: 'flex', flexDirection: 'column', padding: '22px 14px' }}>
        <div style={{ padding: '0 6px 20px', borderBottom: `1px solid ${C.paleGray}`, marginBottom: 18, display: 'flex', alignItems: 'center', gap: 10 }}>
          <img src="/quant-logo.png" alt="Quant" style={{ width: 36, height: 36, borderRadius: 10, flexShrink: 0, objectFit: 'cover' }} />
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
                onClick={() => { setPage(item.key); setMobileNavOpen(false); }}
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

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, width: '100%' }}>
        <header className="top-header" style={{ background: 'transparent', padding: '18px 24px 0', display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: 12 }}>
          <button
            className="mobile-menu-btn"
            onClick={() => setMobileNavOpen(true)}
            aria-label="Apri il menu"
            style={{ fontSize: 16 }}
          >
            ☰
          </button>
          <div className="search-input-wrap" style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: C.white, border: `1px solid ${C.paleGray}`, borderRadius: 999, padding: '9px 18px', boxShadow: '0 1px 2px rgba(0,0,0,0.03), 0 4px 12px rgba(0,0,0,0.04)' }}>
            <Search size={16} strokeWidth={1.5} color={C.darkGray} />
            <input
              type="text"
              placeholder="Cerca progetto, cliente, voce o codice…"
              className="search-input"
              style={{ border: 'none', outline: 'none', background: 'transparent', fontSize: 13, color: C.darkGray, width: 'clamp(120px, 30vw, 260px)' }}
            />
          </div>
        </header>

        {/* overflowX qui era 'auto': su qualunque browser questo trasforma <main> in un "contenitore di scroll",
            che diventa il riferimento per il posizionamento "sticky" al suo interno — rompendo lo sticky della
            colonna Listino nel computo metrico (restava ancorato al box di <main>, che non scorre mai da solo,
            invece che alla pagina). 'clip' taglia comunque l'overflow orizzontale ma senza creare quel contenitore,
            quindi lo sticky funziona di nuovo; le singole tabelle larghe restano scorrevili con la classe .table-scroll. */}
        <main className="app-main" style={{ padding: 24, flex: 1, overflowX: 'clip', minWidth: 0, width: '100%' }}>
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
          {page === 'impostazioni' && <ImpostazioniPage settings={studioSettings} onUpdate={(patch) => setStudioSettings((s) => ({ ...s, ...patch }))} />}
        </main>
      </div>
      </div>
    </>
  );
}
