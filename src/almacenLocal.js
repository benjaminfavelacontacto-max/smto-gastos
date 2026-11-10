// src/almacenLocal.js
// Guarda archivos en IndexedDB de ESTE navegador; nunca salen de la
// computadora. Se usa para el último archivo de Saldos que se subió y el último
// Presupuesto Dinámico generado, así "Presupuesto Dinámico" sólo pide la Hoja
// de trabajo. Si el navegador no permite IndexedDB (p. ej. ventana privada),
// todo sigue funcionando: sólo se vuelven a pedir los archivos.
const DB = 'smto-gastos-archivos'
const STORE = 'archivos'

function abrir() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') return reject(new Error('Sin IndexedDB'))
    const req = indexedDB.open(DB, 1)
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE)
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
}

async function operacion(modo, fn) {
  const db = await abrir()
  try {
    return await new Promise((resolve, reject) => {
      const t = db.transaction(STORE, modo)
      const req = fn(t.objectStore(STORE))
      t.oncomplete = () => resolve(req?.result)
      t.onerror = () => reject(t.error)
      t.onabort = () => reject(t.error)
    })
  } finally {
    db.close()
  }
}

// Respaldo en memoria por si IndexedDB no está disponible: dura la sesión.
const memoria = new Map()

// Pide al navegador que no borre estos archivos aunque le falte espacio en
// disco (Chrome/Edge lo conceden solos según el uso del sitio; Firefox
// pregunta). No evita que el usuario los borre al limpiar los datos del sitio.
let persistenciaPedida = false
async function pedirPersistencia() {
  if (persistenciaPedida) return
  persistenciaPedida = true
  try {
    if (navigator.storage?.persist && !(await navigator.storage.persisted())) await navigator.storage.persist()
  } catch { /* sin Storage API */ }
}

// registro: { nombre, buffer: ArrayBuffer, ...metadatos }
export async function guardarArchivo(clave, registro) {
  const r = { ...registro, guardado: Date.now() }
  memoria.set(clave, r)
  try {
    await operacion('readwrite', s => s.put(r, clave))
    pedirPersistencia()
  } catch { /* queda en memoria */ }
}

export async function leerArchivo(clave) {
  try {
    const r = await operacion('readonly', s => s.get(clave))
    if (r) return r
  } catch { /* sin IndexedDB */ }
  return memoria.get(clave) || null
}
