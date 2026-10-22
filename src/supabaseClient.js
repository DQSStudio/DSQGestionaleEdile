import { createClient } from '@supabase/supabase-js'

// Suite Desearq: un unico database per tutti i prodotti dello studio.
export const SUITE_URL = 'https://qudoygcfphqhhyenpwlr.supabase.co'
export const SUITE_KEY = 'sb_publishable_JuOo4pGtRT_C0Wq7dYOE8Q_XpIRFOEX'
export const supabase = createClient(SUITE_URL, SUITE_KEY)

// Vecchio Manager/Gestionale: serve solo per la copia dei dati al primo avvio.
export const OLD_URL = 'https://qgeiehavpnqdxqnggfzq.supabase.co'
export const OLD_KEY = 'sb_publishable_HgY7u9Q1ls9oHW07Qvcb-A_kgYntyPn'

// Le tabelle e le funzioni del Gestionale nella suite hanno il prefisso "quant_".
export const cea = {
  from: (t) => supabase.from('quant_' + t),
  rpc: (f, a) => supabase.rpc('quant_' + f, a),
}
