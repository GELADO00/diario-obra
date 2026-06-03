/**
 * Pure utility functions extracted from diario-obra.html.
 *
 * Functions that previously relied on `new Date()` or a global `usuarioAtual`
 * now accept those values as parameters so they can be tested deterministically.
 * The HTML file keeps its own identical global copies — these exports are the
 * testable mirror.
 */

export function uid() {
  return Date.now() + Math.floor(Math.random() * 9999);
}

/** @param {Date} [now] */
export function hoje(now = new Date()) {
  return now.toISOString().slice(0, 10);
}

export function toInputDate(s) {
  if (!s) return "";
  let str = String(s).trim();
  if (str.includes("T")) str = str.split("T")[0];
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
  return "";
}

export function fmtDate(s) {
  if (!s) return "—";
  let str = String(s).trim();
  if (str.includes("T")) str = str.split("T")[0];
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) {
    const [y, m, d] = str.split("-");
    return d + "/" + m + "/" + y;
  }
  return str;
}

/**
 * Calculates the final deadline date given a start date, base duration, and
 * optional additive days. The first day counts, so a 1-day obra that starts
 * on Jan 1 ends on Jan 1.
 *
 * NOTE: Uses `new Date(inicio)` which interprets ISO date strings as UTC
 * midnight, but then calls `setDate` in local time. In non-UTC environments
 * this can shift the result by one day. Behaviour is preserved as-is from
 * the original source.
 */
export function calcPrazoFinal(inicio, dias, aditivoDias) {
  if (!inicio || !dias) return "";
  try {
    const total = parseInt(dias) + (parseInt(aditivoDias) || 0);
    const d = new Date(inicio);
    d.setDate(d.getDate() + (total - 1));
    return d.toISOString().slice(0, 10);
  } catch (e) {
    return "";
  }
}

/**
 * Returns how many days remain until `prazo` (inclusive of the deadline day),
 * optionally extended by `aditivoDias`. Returns null for invalid input.
 * Negative means overdue; 0 means expired yesterday; 1 means due today.
 *
 * @param {string|null} prazo - ISO date "YYYY-MM-DD"
 * @param {string|number|null} aditivoDias
 * @param {Date} [now]
 */
export function diasR(prazo, aditivoDias, now = new Date()) {
  if (!prazo) return null;
  let str = String(prazo).trim();
  if (str.includes("T")) str = str.split("T")[0];
  if (!str.match(/^\d{4}-\d{2}-\d{2}$/)) return null;
  const [py, pm, pd] = str.split("-").map(Number);
  const prazoDate = Date.UTC(py, pm - 1, pd);
  const adDias = aditivoDias ? parseInt(aditivoDias) || 0 : 0;
  const prazoFinal = prazoDate + adDias * 86400000;
  const hojeUTC = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((prazoFinal - hojeUTC) / 86400000) + 1;
}

export function safeArr(v) {
  if (Array.isArray(v)) return v;
  try {
    return JSON.parse(v) || [];
  } catch (e) {
    return [];
  }
}

/** @param {{ perfil: string }|null} usuario */
export function isAdmin(usuario) {
  return usuario != null && String(usuario.perfil).toUpperCase() === "ADM";
}

/** @param {{ perfil: string }|null} usuario */
export function isGestor(usuario) {
  return usuario != null && String(usuario.perfil).toUpperCase() === "GESTOR";
}

/** @param {{ perfil: string }|null} usuario */
export function isTecnico(usuario) {
  return usuario != null && String(usuario.perfil).toUpperCase() === "TECNICO";
}

/** @param {{ perfil: string }|null} usuario */
export function podeEditar(usuario) {
  return !isGestor(usuario);
}

/**
 * Returns the key for the current fortnightly period: "YYYY-MM-15" for the
 * first half, or "YYYY-MM-<last-day>" for the second half.
 *
 * @param {Date} [now]
 */
export function quinzenaAtual(now = new Date()) {
  const dia = now.getDate();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  return dia <= 15
    ? `${y}-${m}-15`
    : `${y}-${m}-${new Date(y, now.getMonth() + 1, 0).getDate()}`;
}

/**
 * Returns a human-readable label for a quinzena key, e.g.
 * "1ª Quinzena de janeiro/2024".
 *
 * @param {string} data - quinzena key "YYYY-MM-DD"
 */
export function quinzenaLabel(data) {
  if (!data) return "";
  const [y, m, d] = data.split("-");
  const dia = parseInt(d);
  return `${dia <= 15 ? "1ª" : "2ª"} Quinzena de ${new Date(
    Number(y),
    parseInt(m) - 1,
    1
  ).toLocaleString("pt-BR", { month: "long" })}/${y}`;
}

/**
 * Returns true if the obra already has a quinzena comment for the current
 * fortnightly period.
 *
 * @param {{ comentariosQuinzenais: any }} o
 * @param {Date} [now]
 */
export function temComentarioQuinzena(o, now = new Date()) {
  const qa = quinzenaAtual(now);
  const cs = safeArr(o.comentariosQuinzenais);
  return cs.some((c) => c.data === qa);
}

/**
 * Returns true if the obra has at least one update entry within the last
 * 7 days (inclusive of the boundary).
 *
 * @param {{ atualizacoes: any }} o
 * @param {Date} [now]
 */
export function temRetornoSemanal(o, now = new Date()) {
  const atualizacoes = safeArr(o.atualizacoes);
  if (!atualizacoes.length) return false;
  const seteAtras = new Date(now);
  seteAtras.setDate(seteAtras.getDate() - 7);
  return atualizacoes.some((a) => {
    if (!a.data) return false;
    return new Date(a.data) >= seteAtras;
  });
}
