import { describe, it, expect } from "vitest";
import {
  toInputDate,
  fmtDate,
  calcPrazoFinal,
  diasR,
  safeArr,
  isAdmin,
  isGestor,
  podeEditar,
  quinzenaAtual,
  quinzenaLabel,
  temComentarioQuinzena,
  temRetornoSemanal,
} from "./utils.js";

// ── toInputDate ────────────────────────────────────────────────────────────
describe("toInputDate", () => {
  it("returns empty string for falsy inputs", () => {
    expect(toInputDate(null)).toBe("");
    expect(toInputDate(undefined)).toBe("");
    expect(toInputDate("")).toBe("");
    expect(toInputDate(0)).toBe("");
    expect(toInputDate(false)).toBe("");
  });

  it("accepts a valid ISO date unchanged", () => {
    expect(toInputDate("2024-01-15")).toBe("2024-01-15");
    expect(toInputDate("2000-12-31")).toBe("2000-12-31");
  });

  it("strips the time portion from an ISO datetime string", () => {
    expect(toInputDate("2024-06-01T10:30:00Z")).toBe("2024-06-01");
    expect(toInputDate("2024-06-01T00:00:00.000Z")).toBe("2024-06-01");
  });

  it("trims surrounding whitespace before validating", () => {
    expect(toInputDate("  2024-01-15  ")).toBe("2024-01-15");
  });

  it("returns empty string for non-ISO formats", () => {
    expect(toInputDate("15/01/2024")).toBe("");
    expect(toInputDate("Jan 15 2024")).toBe("");
    expect(toInputDate("20240115")).toBe("");
  });
});

// ── fmtDate ────────────────────────────────────────────────────────────────
describe("fmtDate", () => {
  it("returns em-dash for falsy inputs", () => {
    expect(fmtDate(null)).toBe("—");
    expect(fmtDate(undefined)).toBe("—");
    expect(fmtDate("")).toBe("—");
    expect(fmtDate(0)).toBe("—");
  });

  it("converts ISO date to DD/MM/YYYY", () => {
    expect(fmtDate("2024-01-15")).toBe("15/01/2024");
    expect(fmtDate("2000-12-31")).toBe("31/12/2000");
    expect(fmtDate("2024-02-29")).toBe("29/02/2024");
  });

  it("strips time before formatting an ISO datetime string", () => {
    expect(fmtDate("2024-06-01T10:30:00Z")).toBe("01/06/2024");
  });

  it("trims whitespace before processing", () => {
    expect(fmtDate("  2024-03-05  ")).toBe("05/03/2024");
  });

  it("returns the original string when format is unrecognised", () => {
    expect(fmtDate("15/01/2024")).toBe("15/01/2024");
    expect(fmtDate("some text")).toBe("some text");
  });
});

// ── calcPrazoFinal ─────────────────────────────────────────────────────────
describe("calcPrazoFinal", () => {
  it("returns empty string when inicio is missing", () => {
    expect(calcPrazoFinal("", 10, 0)).toBe("");
    expect(calcPrazoFinal(null, 10, 0)).toBe("");
  });

  it("returns empty string when dias is missing", () => {
    expect(calcPrazoFinal("2024-01-01", "", 0)).toBe("");
    expect(calcPrazoFinal("2024-01-01", null, 0)).toBe("");
  });

  it("a 1-day obra starts and ends on the same day (first day counts)", () => {
    expect(calcPrazoFinal("2024-01-01", 1, 0)).toBe("2024-01-01");
  });

  it("calculates the end date correctly within a month", () => {
    expect(calcPrazoFinal("2024-01-01", 10, 0)).toBe("2024-01-10");
    expect(calcPrazoFinal("2024-01-01", 31, 0)).toBe("2024-01-31");
  });

  it("crosses month boundaries correctly", () => {
    expect(calcPrazoFinal("2024-01-25", 10, 0)).toBe("2024-02-03");
  });

  it("handles leap year February correctly", () => {
    expect(calcPrazoFinal("2024-02-01", 29, 0)).toBe("2024-02-29");
  });

  it("adds additive days on top of base days", () => {
    expect(calcPrazoFinal("2024-01-01", 10, 5)).toBe("2024-01-15");
  });

  it("treats missing or zero aditivoDias as no extension", () => {
    expect(calcPrazoFinal("2024-01-01", 10, null)).toBe("2024-01-10");
    expect(calcPrazoFinal("2024-01-01", 10, undefined)).toBe("2024-01-10");
    expect(calcPrazoFinal("2024-01-01", 10, 0)).toBe("2024-01-10");
    expect(calcPrazoFinal("2024-01-01", 10, "")).toBe("2024-01-10");
  });

  it("accepts dias and aditivoDias as strings", () => {
    expect(calcPrazoFinal("2024-01-01", "10", "5")).toBe("2024-01-15");
  });

  it("crosses a year boundary", () => {
    expect(calcPrazoFinal("2024-12-28", 10, 0)).toBe("2025-01-06");
  });
});

// ── diasR ──────────────────────────────────────────────────────────────────
describe("diasR", () => {
  // All tests use a fixed "today" of 2024-06-03 so results are deterministic.
  const TODAY = new Date("2024-06-03T12:00:00Z");

  it("returns null for falsy prazo", () => {
    expect(diasR(null, 0, TODAY)).toBeNull();
    expect(diasR("", 0, TODAY)).toBeNull();
    expect(diasR(undefined, 0, TODAY)).toBeNull();
  });

  it("returns null for a non-ISO format", () => {
    expect(diasR("03/06/2024", 0, TODAY)).toBeNull();
    expect(diasR("June 3 2024", 0, TODAY)).toBeNull();
  });

  it("returns 1 when prazo equals today (due today)", () => {
    expect(diasR("2024-06-03", 0, TODAY)).toBe(1);
  });

  it("returns 2 when prazo is tomorrow", () => {
    expect(diasR("2024-06-04", 0, TODAY)).toBe(2);
  });

  it("returns 0 when prazo was yesterday (expired)", () => {
    expect(diasR("2024-06-02", 0, TODAY)).toBe(0);
  });

  it("returns a negative number for a past deadline", () => {
    expect(diasR("2024-05-20", 0, TODAY)).toBeLessThan(0);
  });

  it("extends the deadline by aditivoDias", () => {
    // prazo today + 5 additive days = 5 extra days remaining (1 + 5 = 6)
    expect(diasR("2024-06-03", 5, TODAY)).toBe(6);
  });

  it("strips time component before calculating", () => {
    expect(diasR("2024-06-03T00:00:00Z", 0, TODAY)).toBe(1);
  });

  it("treats missing or zero aditivoDias as no extension", () => {
    expect(diasR("2024-06-03", null, TODAY)).toBe(1);
    expect(diasR("2024-06-03", undefined, TODAY)).toBe(1);
    expect(diasR("2024-06-03", 0, TODAY)).toBe(1);
    expect(diasR("2024-06-03", "", TODAY)).toBe(1);
  });

  it("accepts aditivoDias as a string", () => {
    expect(diasR("2024-06-03", "10", TODAY)).toBe(11);
  });
});

// ── safeArr ────────────────────────────────────────────────────────────────
describe("safeArr", () => {
  it("returns an array as-is", () => {
    expect(safeArr([])).toEqual([]);
    expect(safeArr([1, 2, 3])).toEqual([1, 2, 3]);
    expect(safeArr([{ a: 1 }])).toEqual([{ a: 1 }]);
  });

  it("parses a valid JSON array string", () => {
    expect(safeArr("[]")).toEqual([]);
    expect(safeArr('[{"id":1}]')).toEqual([{ id: 1 }]);
    expect(safeArr("[1,2,3]")).toEqual([1, 2, 3]);
  });

  it("returns empty array for null or undefined", () => {
    expect(safeArr(null)).toEqual([]);
    expect(safeArr(undefined)).toEqual([]);
  });

  it("returns empty array for invalid JSON", () => {
    expect(safeArr("not-json")).toEqual([]);
    expect(safeArr("{broken")).toEqual([]);
  });

  it("returns empty array for JSON null", () => {
    expect(safeArr("null")).toEqual([]);
  });
});

// ── isAdmin / isGestor / podeEditar ───────────────────────────────────────
describe("isAdmin", () => {
  it("returns false for null or undefined", () => {
    expect(isAdmin(null)).toBe(false);
    expect(isAdmin(undefined)).toBe(false);
  });

  it("returns true when perfil is ADM (case-insensitive)", () => {
    expect(isAdmin({ perfil: "ADM" })).toBe(true);
    expect(isAdmin({ perfil: "adm" })).toBe(true);
    expect(isAdmin({ perfil: "Adm" })).toBe(true);
  });

  it("returns false for other profiles", () => {
    expect(isAdmin({ perfil: "GESTOR" })).toBe(false);
    expect(isAdmin({ perfil: "USER" })).toBe(false);
    expect(isAdmin({ perfil: "" })).toBe(false);
  });
});

describe("isGestor", () => {
  it("returns false for null or undefined", () => {
    expect(isGestor(null)).toBe(false);
    expect(isGestor(undefined)).toBe(false);
  });

  it("returns true when perfil is GESTOR (case-insensitive)", () => {
    expect(isGestor({ perfil: "GESTOR" })).toBe(true);
    expect(isGestor({ perfil: "gestor" })).toBe(true);
    expect(isGestor({ perfil: "Gestor" })).toBe(true);
  });

  it("returns false for other profiles", () => {
    expect(isGestor({ perfil: "ADM" })).toBe(false);
    expect(isGestor({ perfil: "" })).toBe(false);
  });
});

describe("podeEditar", () => {
  it("returns false for GESTOR (read-only role)", () => {
    expect(podeEditar({ perfil: "GESTOR" })).toBe(false);
  });

  it("returns true for ADM and any other role", () => {
    expect(podeEditar({ perfil: "ADM" })).toBe(true);
    expect(podeEditar({ perfil: "USER" })).toBe(true);
    expect(podeEditar(null)).toBe(true);
  });
});

// ── quinzenaAtual ─────────────────────────────────────────────────────────
describe("quinzenaAtual", () => {
  it("returns YYYY-MM-15 for any day in the first half (days 1–15)", () => {
    expect(quinzenaAtual(new Date("2024-01-01"))).toBe("2024-01-15");
    expect(quinzenaAtual(new Date("2024-01-08"))).toBe("2024-01-15");
    expect(quinzenaAtual(new Date("2024-01-15"))).toBe("2024-01-15");
  });

  it("returns YYYY-MM-<last-day> for any day in the second half (days 16–end)", () => {
    expect(quinzenaAtual(new Date("2024-01-16"))).toBe("2024-01-31");
    expect(quinzenaAtual(new Date("2024-01-31"))).toBe("2024-01-31");
  });

  it("handles February in a leap year (last day = 29)", () => {
    expect(quinzenaAtual(new Date("2024-02-16"))).toBe("2024-02-29");
    expect(quinzenaAtual(new Date("2024-02-29"))).toBe("2024-02-29");
  });

  it("handles February in a non-leap year (last day = 28)", () => {
    expect(quinzenaAtual(new Date("2023-02-16"))).toBe("2023-02-28");
    expect(quinzenaAtual(new Date("2023-02-28"))).toBe("2023-02-28");
  });

  it("handles months with 30 days", () => {
    expect(quinzenaAtual(new Date("2024-04-16"))).toBe("2024-04-30");
    expect(quinzenaAtual(new Date("2024-06-30"))).toBe("2024-06-30");
  });

  it("pads the month with a leading zero", () => {
    expect(quinzenaAtual(new Date("2024-03-01"))).toBe("2024-03-15");
    expect(quinzenaAtual(new Date("2024-09-20"))).toBe("2024-09-30");
  });
});

// ── quinzenaLabel ─────────────────────────────────────────────────────────
describe("quinzenaLabel", () => {
  it("returns empty string for falsy input", () => {
    expect(quinzenaLabel("")).toBe("");
    expect(quinzenaLabel(null)).toBe("");
    expect(quinzenaLabel(undefined)).toBe("");
  });

  it("labels first-half keys as '1ª Quinzena'", () => {
    expect(quinzenaLabel("2024-01-15")).toMatch(/^1ª Quinzena de /);
  });

  it("labels second-half keys as '2ª Quinzena'", () => {
    expect(quinzenaLabel("2024-01-31")).toMatch(/^2ª Quinzena de /);
    expect(quinzenaLabel("2024-02-29")).toMatch(/^2ª Quinzena de /);
  });

  it("includes the year in the label", () => {
    expect(quinzenaLabel("2024-06-15")).toContain("/2024");
    expect(quinzenaLabel("2025-01-31")).toContain("/2025");
  });

  it("includes the correct Portuguese month name", () => {
    expect(quinzenaLabel("2024-01-15")).toContain("janeiro");
    expect(quinzenaLabel("2024-06-15")).toContain("junho");
    expect(quinzenaLabel("2024-12-31")).toContain("dezembro");
  });
});

// ── temComentarioQuinzena ─────────────────────────────────────────────────
describe("temComentarioQuinzena", () => {
  const NOW = new Date("2024-06-10"); // 10th → first quinzena → key = "2024-06-15"

  it("returns false when there are no comments", () => {
    expect(temComentarioQuinzena({ comentariosQuinzenais: [] }, NOW)).toBe(false);
    expect(temComentarioQuinzena({ comentariosQuinzenais: null }, NOW)).toBe(false);
  });

  it("returns true when a comment exists for the current quinzena", () => {
    const o = { comentariosQuinzenais: [{ data: "2024-06-15", texto: "ok" }] };
    expect(temComentarioQuinzena(o, NOW)).toBe(true);
  });

  it("returns false when comments exist only for other quinzenas", () => {
    const o = {
      comentariosQuinzenais: [
        { data: "2024-05-15", texto: "anterior" },
        { data: "2024-06-30", texto: "segunda quinzena" },
      ],
    };
    expect(temComentarioQuinzena(o, NOW)).toBe(false);
  });

  it("accepts comentariosQuinzenais as a JSON string", () => {
    const o = {
      comentariosQuinzenais: JSON.stringify([{ data: "2024-06-15", texto: "ok" }]),
    };
    expect(temComentarioQuinzena(o, NOW)).toBe(true);
  });
});

// ── temRetornoSemanal ─────────────────────────────────────────────────────
describe("temRetornoSemanal", () => {
  // Use UTC midnight so "7 days ago" aligns exactly with the seteAtras boundary.
  // When `now` has a time component (e.g. noon), an update stored as a date-only
  // string ("YYYY-MM-DD" → midnight UTC) from exactly 7 days prior falls strictly
  // before seteAtras and would return false — a known time-sensitivity in the impl.
  const NOW = new Date("2024-06-10T00:00:00Z");

  it("returns false when there are no updates", () => {
    expect(temRetornoSemanal({ atualizacoes: [] }, NOW)).toBe(false);
    expect(temRetornoSemanal({ atualizacoes: null }, NOW)).toBe(false);
  });

  it("returns true for an update made today", () => {
    const o = { atualizacoes: [{ data: "2024-06-10", texto: "ok" }] };
    expect(temRetornoSemanal(o, NOW)).toBe(true);
  });

  it("returns true for an update made yesterday", () => {
    const o = { atualizacoes: [{ data: "2024-06-09", texto: "ok" }] };
    expect(temRetornoSemanal(o, NOW)).toBe(true);
  });

  it("returns true for an update exactly 7 days ago when now is midnight UTC", () => {
    // seteAtras = June 3 00:00 UTC; update "2024-06-03" = June 3 00:00 UTC → equal → true
    const o = { atualizacoes: [{ data: "2024-06-03", texto: "ok" }] };
    expect(temRetornoSemanal(o, NOW)).toBe(true);
  });

  it("returns false for an update 8 days ago", () => {
    const o = { atualizacoes: [{ data: "2024-06-02", texto: "ok" }] };
    expect(temRetornoSemanal(o, NOW)).toBe(false);
  });

  it("returns false when now has a time component and the update is exactly 7 days prior at midnight", () => {
    // This documents the time-sensitivity: noon-based seteAtras > midnight update
    const nowNoon = new Date("2024-06-10T12:00:00Z");
    const o = { atualizacoes: [{ data: "2024-06-03", texto: "ok" }] };
    expect(temRetornoSemanal(o, nowNoon)).toBe(false);
  });

  it("ignores update entries that have no data field", () => {
    const o = { atualizacoes: [{ texto: "sem data" }] };
    expect(temRetornoSemanal(o, NOW)).toBe(false);
  });

  it("returns true when at least one update is within the window", () => {
    const o = {
      atualizacoes: [
        { data: "2024-05-01", texto: "antigo" },
        { data: "2024-06-08", texto: "recente" },
      ],
    };
    expect(temRetornoSemanal(o, NOW)).toBe(true);
  });

  it("accepts atualizacoes as a JSON string", () => {
    const o = {
      atualizacoes: JSON.stringify([{ data: "2024-06-10", texto: "ok" }]),
    };
    expect(temRetornoSemanal(o, NOW)).toBe(true);
  });
});
