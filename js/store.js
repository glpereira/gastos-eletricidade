// Guarda os dados no próprio dispositivo (IndexedDB). Nada sai do iPad.
const DB = 'gastos-eletricidade';
const LOJA = 'dados';

function abrir() {
  return new Promise((ok, erro) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(LOJA);
    r.onsuccess = () => ok(r.result);
    r.onerror = () => erro(r.error);
  });
}

export async function ler(chave, omissao) {
  try {
    const db = await abrir();
    return await new Promise((ok, erro) => {
      const r = db.transaction(LOJA).objectStore(LOJA).get(chave);
      r.onsuccess = () => ok(r.result ?? omissao);
      r.onerror = () => erro(r.error);
    });
  } catch {
    return omissao;
  }
}

export async function gravar(chave, valor) {
  const db = await abrir();
  return new Promise((ok, erro) => {
    const t = db.transaction(LOJA, 'readwrite');
    t.objectStore(LOJA).put(valor, chave);
    t.oncomplete = () => ok();
    t.onerror = () => erro(t.error);
  });
}

export async function pedirPersistencia() {
  try {
    if (navigator.storage?.persist) return await navigator.storage.persist();
  } catch {}
  return false;
}
