import { createClient } from '@supabase/supabase-js'

export const supabase = createClient(
  'https://qgeiehavpnqdxqnggfzq.supabase.co',
  'sb_publishable_HgY7u9Q1ls9oHW07Qvcb-A_kgYntyPn'
)

// Tutte le tabelle del Software di Computazione Edile vivono nello schema "cea"
// (separato dagli altri programmi che condividono lo stesso progetto Supabase).
export const cea = supabase.schema('cea')
