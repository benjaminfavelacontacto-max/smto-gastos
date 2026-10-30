/* ═══════════════════════════════════════════════════
   BORRADOR LOCAL (IndexedDB)
   Guarda el reporte en curso SOLO en este navegador para recuperarlo si se
   recarga o cierra la página. Nada sale de la computadora del usuario.

   - Un solo registro ('actual') en el store 'borradores'. Cada guardado es
     una transacción de IndexedDB: o se escribe completo o no se escribe
     (no hay borradores a medias).
   - IndexedDB clona File/Blob nativamente, así que los XML/PDF/imágenes
     viajan con su gasto sin convertirlos.
   - `SCHEMA` versiona la forma del registro; uno de otra versión se descarta
     en vez de intentar restaurarlo a ciegas.
   - Caduca a los CADUCIDAD_DIAS días.
   - Sin IndexedDB (modo privado estricto, almacenamiento bloqueado) todas
     las funciones fallan de forma controlada y la app sigue como antes.
═══════════════════════════════════════════════════ */

const DB_NAME = 'smto-gastos'
const STORE = 'borradores'
const KEY = 'actual'
const SCHEMA = 1
export const CADUCIDAD_DIAS = 30

let dbPromise = null
const abrirDB = () => {
  if (dbPromise) return dbPromise
  dbPromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') { reject(new Error('IndexedDB no disponible')); return }
    const req = indexedDB.open(DB_NAME, 1)
    req.onupgradeneeded = () => req.result.createObjectStore(STORE)
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error || new Error('No se pudo abrir IndexedDB'))
    req.onblocked = () => reject(new Error('IndexedDB bloqueado'))
  }).catch(err => { dbPromise = null; throw err })
  return dbPromise
}

const tx = async (mode, fn) => {
  const db = await abrirDB()
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode)
    const store = t.objectStore(STORE)
    let result
    const req = fn(store)
    if (req) req.onsuccess = () => { result = req.result }
    t.oncomplete = () => resolve(result)
    t.onerror = () => reject(t.error || req?.error)
    t.onabort = () => reject(t.error || new Error('Transacción abortada'))
  })
}

// Quita del gasto lo que es solo de UI (animación de fila nueva).
const limpiarGasto = ({ isNew, ...g }) => g

export async function guardarBorrador({ lista, colaborador, carpetaNombre, appVersion }) {
  const registro = {
    schema: SCHEMA,
    appVersion,
    savedAt: Date.now(),
    colaborador: colaborador ? { nombre: colaborador.nombre, categoria: colaborador.categoria } : null,
    carpetaNombre,
    count: lista.length,
    lista: lista.map(limpiarGasto),
  }
  await tx('readwrite', s => s.put(registro, KEY))
  return registro.savedAt
}

export const borrarBorrador = () => tx('readwrite', s => s.delete(KEY))

/* Lee el borrador y verifica su integridad antes de ofrecerlo.
   Devuelve null si no hay, caducó o no es de este esquema (y lo borra);
   si no, { ...registro, avisos[] } con lo que no cuadró. */
export async function leerBorrador() {
  const r = await tx('readonly', s => s.get(KEY))
  if (!r) return null
  const caducado = Date.now() - (r.savedAt || 0) > CADUCIDAD_DIAS * 86400000
  const valido = r.schema === SCHEMA && Array.isArray(r.lista) && r.count === r.lista.length
    && r.lista.every(g => g && typeof g.id === 'string')
  if (caducado || !valido || !r.lista.length) {
    await borrarBorrador().catch(() => {})
    return null
  }
  // Relación archivos ↔ gastos: un gasto que dice tener PDF debe traer el
  // archivo (File) o su dataURL. Si no, se avisa; no se inventa nada.
  const sinArchivo = r.lista.filter(g => g.tienePDF && !(g.pdfFile instanceof Blob) && !g.pdfDataURL).length
  const avisos = []
  if (sinArchivo) avisos.push(`${sinArchivo} ${sinArchivo === 1 ? 'gasto marca PDF pero el archivo no se recuperó' : 'gastos marcan PDF pero los archivos no se recuperaron'}.`)
  return { ...r, avisos }
}
