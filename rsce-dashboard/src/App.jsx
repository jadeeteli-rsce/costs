import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ReferenceLine,
} from "recharts";
import { Search, Download, Award, TrendingUp, TrendingDown, Minus, PawPrint, Save, CheckCircle, Loader } from "lucide-react";
import * as XLSX from "xlsx";


import { createClient } from "@supabase/supabase-js";
const SUPABASE_URL  = "https://kyexopavvsbyrdwtefca.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imt5ZXhvcGF2dnNieXJkd3RlZmNhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM5MTEzMDMsImV4cCI6MjA5OTQ4NzMwM30.NFVWMgeYSEOETDe_f8yH6TbULtxWuodFORgy8WImG_8";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);

const CT_LABELS = {
  Member: "Socios / Miembros",
  User: "Usuarios (No Socios)",
  Canine_Collaborator: "Colaboradoras Caninas",
};
const CT_COLORS = {
  Member: "#B98A3F",
  User: "#1C2B45",
  Canine_Collaborator: "#5C7A5E",
};
const CT_ORDER = ["Canine_Collaborator","Member", "User"];

function lastNonNull(series, field) {
  for (let i = series.length - 1; i >= 0; i--) {
    if (series[i][field] !== null && series[i][field] !== undefined) {
      return series[i];
    }
  }
  return null;
}

const APP_PASSWORD = "rsce2026"; // change this to whatever you want

function PasswordGate({ children }) {
  const [input, setInput] = useState("");
  const [unlocked, setUnlocked] = useState(() => {
    return sessionStorage.getItem("rsce_unlocked") === "true";
  });
  const [error, setError] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (input === APP_PASSWORD) {
      sessionStorage.setItem("rsce_unlocked", "true");
      setUnlocked(true);
      setError(false);
    } else {
      setError(true);
    }
  };

  if (unlocked) return children;

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: "#F6F3EC", fontFamily: "'Inter', system-ui, sans-serif",
    }}>
      <form onSubmit={handleSubmit} style={{
        background: "white", padding: "32px 28px", borderRadius: 10,
        border: "1px solid #E5DFD1", width: 320,
      }}>
        <div style={{ fontFamily: "'Georgia', serif", fontSize: 20, marginBottom: 4, color: "#1C2B45" }}>
          Real Sociedad Canina de España
        </div>
        <div style={{ fontSize: 13, opacity: 0.7, marginBottom: 20 }}>
          Introduce la contraseña para acceder al panel
        </div>
        <input
          type="password"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Contraseña"
          autoFocus
          style={{
            width: "100%", padding: "8px 10px", borderRadius: 6, border: "1px solid #D4CDBB",
            fontSize: 14, boxSizing: "border-box", marginBottom: 12, fontFamily: "inherit",
          }}
        />
        {error && (
          <div style={{ color: "#A6452E", fontSize: 12.5, marginBottom: 12 }}>
            Contraseña incorrecta.
          </div>
        )}
        <button type="submit" style={{
          width: "100%", padding: "8px 10px", borderRadius: 6, border: "none",
          background: "#B98A3F", color: "#1C2B45", fontWeight: 700, fontSize: 13.5,
          cursor: "pointer",
        }}>
          Entrar
        </button>
      </form>
    </div>
  );
}

function secondLastNonNull(series, field, beforeYear) {
  for (let i = series.length - 1; i >= 0; i--) {
    if (series[i].year < beforeYear && series[i][field] !== null && series[i][field] !== undefined) {
      return series[i];
    }
  }
  return null;
}

function prevYearEntry(series, beforeYear) {
  for (let i = series.length - 1; i >= 0; i--) {
    if (series[i].year < beforeYear) return series[i];
  }
  return null;
}

function valueForYear(series, field, year) {
  const entry = series.find((p) => p.year === year);
  if (!entry) return null;
  if (entry[field] === null || entry[field] === undefined) return null;
  return entry;
}

function pctChange(a, b) {
  if (a === null || a === undefined || a === 0) return null;
  return (b - a) / a;
}

function fmtEUR(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return (0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
  return v.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
}

function buildCodeToProduct(products) {
  const normalizedProducts = {};
  for (const p of products) {
    normalizedProducts[normalizeName(p)] = p;
  }
  const map = {};
  for (const { code, name } of CODE_MAP) {
    const match = normalizedProducts[normalizeName(name)];
    if (match) map[code] = match;
  }
  return map;
}

const normalizeName = (s) =>
  s.toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // strip accents
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

function fmtPct(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return "—";
  const s = (v * 100).toLocaleString("es-ES", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  return (v > 0 ? "+" : "") + s + "%";
}

function TrendIcon({ v }) {
  if (v === null || v === undefined) return <Minus size={14} style={{ opacity: 0.35 }} />;
  if (v > 0.0005) return <TrendingUp size={14} color="#5C7A5E" />;
  if (v < -0.0005) return <TrendingDown size={14} color="#A6452E" />;
  return <Minus size={14} style={{ opacity: 0.5 }} />;
}

// NOTE: CODE_MAP and RSCE_DATA are large constant arrays/objects from the original file.
// PLACEHOLDER_DATA_BLOCK

  function RSCEDashboard() {
  const [products, setProducts] = useState(() => [...RSCE_DATA.products]);
  const [categories, setCategories] = useState(() => [...RSCE_DATA.categories]);
  const [data, setData] = useState(() => JSON.parse(JSON.stringify(RSCE_DATA.data)));
  const [years, setYears] = useState(() => [...RSCE_DATA.years]);

  const [query, setQuery] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("Todas");
  const [selected, setSelected] = useState(products[0]);
  const [vatMode, setVatMode] = useState("with_vat"); // 'with_vat' | 'no_vat'
  const [cpiRate, setCpiRate] = useState(0.02);
  const [cpiDraft, setCpiDraft] = useState((0.02 * 100).toFixed(1)); // ← add this line

  const [overrides, setOverrides] = useState({}); // key `${product}::${ct}` -> number

  const [loading, setLoading] = useState(true);
const [saveStatus, setSaveStatus] = useState("idle"); // 'idle' | 'saving' | 'saved' | 'error'
const isFirstRun = useRef(true);


useEffect(() => {
  let cancelled = false;
  (async () => {
    const { data: row, error } = await supabase
      .from("rsce_dataset").select("*").eq("id", 1).maybeSingle();

    if (error) { console.error(error); setLoading(false); return; }

    if (row) {
      if (!cancelled) {
        setProducts(row.products);
        setCategories(row.categories);
        setData(row.data);
      }
    } else {
      // first run ever — seed Supabase from the bundled data
      await supabase.from("rsce_dataset").insert({
        id: 1, products: RSCE_DATA.products,
        categories: RSCE_DATA.categories, data: RSCE_DATA.data,
      });
    }

    const { data: overrideRows } = await supabase.from("rsce_overrides").select("*");
    if (overrideRows && !cancelled) {
      const ov = {};
      overrideRows.forEach((r) => { ov[`${r.product}::${r.client_type}`] = r.value; });
      setOverrides(ov);
    }
    if (!cancelled) setLoading(false);
  })();
  return () => { cancelled = true; };
}, []);



useEffect(() => {
  if (isFirstRun.current) { isFirstRun.current = false; return; }
  if (loading) return;
  setSaveStatus("saving");
  supabase.from("rsce_dataset")
    .upsert({ id: 1, products, categories, data, updated_at: new Date().toISOString() })
    .then(({ error }) => setSaveStatus(error ? "error" : "saved"));
}, [products, categories, data, loading]);

console.log("categoryFilter:", categoryFilter, "| sample category:", data[products[0]]?.category);

const [editYear, setEditYear] = useState(RSCE_DATA.years[RSCE_DATA.years.length - 1]);
  const [showAddProduct, setShowAddProduct] = useState(false);
const [newProductName, setNewProductName] = useState("");
const [newProductCategory, setNewProductCategory] = useState(categories[0] || "Sin categorizar");
const [addingNewCategory, setAddingNewCategory] = useState(false);
const [newCategoryInput, setNewCategoryInput] = useState("");

const [compareTypeA, setCompareTypeA] = useState("Member");
const [compareYearA, setCompareYearA] = useState(years[years.length - 2] ?? years[years.length - 1]);
const [compareTypeB, setCompareTypeB] = useState("Member");
const [compareYearB, setCompareYearB] = useState(years[years.length - 1]);

useEffect(() => {
  if (!years.includes(compareYearA)) setCompareYearA(years[years.length - 1]);
  if (!years.includes(compareYearB)) setCompareYearB(years[years.length - 1]);
}, [years, compareYearA, compareYearB]);

const handleAddProduct = useCallback(() => {
  const name = newProductName.trim();
  if (!name) return;
  if (data[name]) {
    alert("Ya existe un producto con ese nombre.");
    return;
  }
  const category = addingNewCategory
    ? (newCategoryInput.trim() || "Sin categorizar")
    : newProductCategory;

  setData((prev) => ({
    ...prev,
    [name]: { category, prices: { Member: [], User: [], Canine_Collaborator: [] } },
  }));
  setProducts((prev) => [...prev, name].sort((a, b) => a.localeCompare(b, "es")));
  setCategories((prev) =>
    prev.includes(category) ? prev : [...prev, category].sort((a, b) => a.localeCompare(b, "es"))
  );

  setSelected(name);
  setShowAddProduct(false);
  setNewProductName("");
  setNewCategoryInput("");
  setAddingNewCategory(false);
}, [newProductName, newProductCategory, newCategoryInput, addingNewCategory, data]);


const handleAddYear = useCallback(() => {
  setYears((prev) => {
    const next = Math.max(...prev) + 1;
    return prev.includes(next) ? prev : [...prev, next];
  });
}, []);

const handleRemoveYear = useCallback((yearToRemove) => {
  if (years.length <= 1) return;
  const ok = window.confirm(
    `¿Eliminar el año ${yearToRemove}? Esto borrará cualquier precio introducido para ese año en todos los productos.`
  );
  if (!ok) return;

  setYears((prev) => prev.filter((y) => y !== yearToRemove));
  setData((prev) => {
    const next = {};
    for (const [product, rec] of Object.entries(prev)) {
      const newPrices = {};
      for (const ct of CT_ORDER) {
        newPrices[ct] = (rec.prices[ct] || []).filter((p) => p.year !== yearToRemove);
      }
      next[product] = { ...rec, prices: newPrices };
    }
    return next;
  });
}, [years]);

const fileInputRef = useRef(null);



const handleImportExcel = useCallback((file) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: "array" });
    const sheetName = wb.SheetNames.includes("Data") ? "Data" : wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

    // map normalized old product names -> category, so we keep categories where possible
    const oldCategoryByName = {};
    for (const p of products) {
      oldCategoryByName[normalizeName(p)] = data[p]?.category || "Sin categorizar";
    }

    const newData = {};
    const yearSet = new Set();

    for (const row of rows) {
      const year = Number(row.Year);
      const product = String(row.Product || "").trim();
      const ct = row.Customer_Type;
      if (!product || !year || !CT_ORDER.includes(ct)) continue;

      if (!newData[product]) {
        const category = oldCategoryByName[normalizeName(product)] || "Sin categorizar";
        newData[product] = { category, prices: { Member: [], User: [], Canine_Collaborator: [] } };
      }
      newData[product].prices[ct].push({
        year,
        with_vat: row.Price_with_VAT === null ? null : Number(row.Price_with_VAT),
        no_vat: row.Price_no_VAT === null ? null : Number(row.Price_no_VAT),
      });
      yearSet.add(year);
    }

    // sort each product's series by year
    for (const p of Object.keys(newData)) {
      for (const ct of CT_ORDER) {
        newData[p].prices[ct].sort((a, b) => a.year - b.year);
      }
    }

    const newProducts = Object.keys(newData).sort((a, b) => a.localeCompare(b, "es"));
    const newCategories = [...new Set(newProducts.map((p) => newData[p].category))].sort((a, b) =>
      a.localeCompare(b, "es")
    );
    const newYears = [...yearSet].sort((a, b) => a - b);

    const ok = window.confirm(
      `Se importarán ${newProducts.length} productos (${newYears[0]}–${newYears[newYears.length - 1]}). ` +
      `Esto SUSTITUYE todos los datos actuales del panel (y de Supabase). ¿Continuar?`
    );
    if (!ok) return;

    setProducts(newProducts);
    setCategories(newCategories);
    setData(newData);
    setYears(newYears);
    setSelected(newProducts[0]);
    setOverrides({}); // old overrides won't map cleanly onto merged product names
    supabase.from("rsce_overrides").delete().neq("product", "__none__"); // clear stale overrides
  };
  reader.readAsArrayBuffer(file);
}, [products, data]);

const updatePrice = useCallback((product, ct, field, rawValue, year) => {
  const val = rawValue === "" ? null : parseFloat(String(rawValue).replace(",", "."));
  setData((prev) => {
    const rec = prev[product];
    if (!rec) return prev;
    const series = rec.prices[ct] || [];
    const idx = series.findIndex((p) => p.year === year);
    const newSeries = idx >= 0
      ? series.map((p, i) => (i === idx ? { ...p, [field]: val } : p))
      : [...series, { year, with_vat: null, no_vat: null, [field]: val }]
          .sort((a, b) => a.year - b.year);
    return { ...prev, [product]: { ...rec, prices: { ...rec.prices, [ct]: newSeries } } };
  });
}, []);

  useEffect(() => {
  setCpiDraft((cpiRate * 100).toFixed(1));
}, [cpiRate]);

  useEffect(() => {
  if (!years.includes(editYear)) {
    setEditYear(years[years.length - 1]);
  }
}, [years, editYear]);

const codeToProduct = useMemo(() => buildCodeToProduct(products), [products]);
const productToCode = useMemo(() => {
  const map = {};
  for (const [code, product] of Object.entries(codeToProduct)) {
    map[product] = code;
  }
  return map;
}, [codeToProduct]);

const [activeTab, setActiveTab] = useState("dashboard"); // 'dashboard' | 'tarifaWeb' | 'colaboradoras'
const [tarifaWebYear, setTarifaWebYear] = useState(years[years.length - 1]);
const [colaboradorasYear, setColaboradorasYear] = useState(years[years.length - 1]);

useEffect(() => {
  if (!years.includes(tarifaWebYear)) setTarifaWebYear(years[years.length - 1]);
  if (!years.includes(colaboradorasYear)) setColaboradorasYear(years[years.length - 1]);
}, [years, tarifaWebYear, colaboradorasYear]);

  const filteredProducts = useMemo(() => {
  const q = query.trim().toLowerCase();
  const matchedByCode = codeToProduct[query.trim()];
  return products.filter((p) => {
    const matchesQuery = q === "" ||
      p.toLowerCase().includes(q) ||
      p === matchedByCode;
    const matchesCat = categoryFilter === "Todas" || data[p].category === categoryFilter;
    return matchesQuery && matchesCat;
  });
}, [query, categoryFilter, products, data, codeToProduct]);

  const selectedRecord = data[selected];

  const perTypeStats = useMemo(() => {
    if (!selectedRecord) return {};
    const latestYear = years[years.length - 1];
    const out = {};
    for (const ct of CT_ORDER) {
      const series = selectedRecord.prices[ct] || [];
      const last = valueForYear(series, vatMode, years[years.length - 1]);
      const prev = last ? secondLastNonNull(series, vatMode, last.year) : null;
      const yoy = last && prev ? pctChange(prev[vatMode], last[vatMode]) : null;
      out[ct] = { last, prev, yoy };
    }
    return out;
  }, [selectedRecord, vatMode, years]);

  const crossTypeVariation = useMemo(() => {
    const m = perTypeStats.Member?.last?.[vatMode] ?? null;
    const u = perTypeStats.User?.last?.[vatMode] ?? null;
    const c = perTypeStats.Canine_Collaborator?.last?.[vatMode] ?? null;
    return {
      userVsMember: m && u ? pctChange(m, u) : null,
      collabVsMember: m && c ? pctChange(m, c) : null,
      collabVsUser: u && c ? pctChange(u, c) : null,
    };
  }, [perTypeStats, vatMode]);

  const customComparison = useMemo(() => {
  if (!selectedRecord) return null;
  const seriesA = selectedRecord.prices[compareTypeA] || [];
  const seriesB = selectedRecord.prices[compareTypeB] || [];
  const entryA = valueForYear(seriesA, vatMode, compareYearA);
  const entryB = valueForYear(seriesB, vatMode, compareYearB);
  const valA = entryA ? entryA[vatMode] : null;
  const valB = entryB ? entryB[vatMode] : null;
  const diff = (valA !== null && valB !== null) ? valB - valA : null;
  const pct = pctChange(valA, valB);
  return { valA, valB, diff, pct };
}, [selectedRecord, compareTypeA, compareYearA, compareTypeB, compareYearB, vatMode]);

  const chartData = useMemo(() => {
    if (!selectedRecord) return [];
    return years.map((y) => {
      const row = { year: y };
      for (const ct of CT_ORDER) {
        const point = (selectedRecord.prices[ct] || []).find((p) => p.year === y);
        row[ct] = point ? point[vatMode] : null;
      }
      return row;
    });
  }, [selectedRecord, years, vatMode]);

  const setOverride = useCallback((ct, value) => {
  const key = `${selected}::${ct}`;
  const numValue = value === "" || value === null ? null : parseFloat(value);

  setOverrides((prev) => {
    const next = { ...prev };
    if (numValue === null || Number.isNaN(numValue)) delete next[key];
    else next[key] = numValue;
    return next;
  });

  setSaveStatus("saving");
  const req = (numValue === null || Number.isNaN(numValue))
    ? supabase.from("rsce_overrides").delete().eq("product", selected).eq("client_type", ct)
    : supabase.from("rsce_overrides").upsert({
        product: selected, client_type: ct, value: numValue,
        updated_at: new Date().toISOString(),
      });

  req.then(({ error }) => setSaveStatus(error ? "error" : "saved"));
}, [selected]);

  const overrideKey = (p, ct) => `${p}::${ct}`;

  const handleExport = useCallback(() => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: raw long-format data
    const rawRows = [["Año", "Producto", "Categoría", "Tipo de Cliente", "Precio con IVA", "Precio sin IVA"]];
    for (const p of products) {
      const rec = data[p];
      for (const ct of CT_ORDER) {
        for (const point of rec.prices[ct]) {
          rawRows.push([point.year, p, rec.category, CT_LABELS[ct], point.with_vat, point.no_vat]);
        }
      }
    }
    const wsRaw = XLSX.utils.aoa_to_sheet(rawRows);
    XLSX.utils.book_append_sheet(wb, wsRaw, "Datos");

    // Sheet 2: summary + forecast snapshot (values, not live formulas)
    const summaryHeader = ["Producto", "Categoría"];
    for (const ct of CT_ORDER) {
      const label = CT_LABELS[ct];
      summaryHeader.push(
        `${label}: Año base`, `${label}: Precio base (€)`,
        `${label}: Variación interanual (%)`,
        `${label}: Previsión IPC (€)`, `${label}: Anulación manual (€)`, `${label}: Precio final (€)`
      );
    }
    const summaryRows = [summaryHeader];
    for (const p of products) {
      const rec = data[p];
      const row = [p, rec.category];
      for (const ct of CT_ORDER) {
        const series = rec.prices[ct] || [];
        const last = lastNonNull(series, vatMode);
        const prev = last ? secondLastNonNull(series, vatMode, last.year) : null;
        const yoy = last && prev ? pctChange(prev[vatMode], last[vatMode]) : null;
        const basePrice = last ? last[vatMode] : null;
        const forecast = basePrice !== null ? Math.round(basePrice * (1 + cpiRate) * 100) / 100 : null;
        const ov = overrides[overrideKey(p, ct)];
        const final = ov !== undefined && !Number.isNaN(ov) ? ov : forecast;
        row.push(
          last ? last.year : "",
          basePrice,
          yoy !== null ? Math.round(yoy * 1000) / 10 : "",
          forecast,
          ov !== undefined ? ov : "",
          final
        );
      }
      summaryRows.push(row);
    }
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Resumen_y_Prevision");

    // Sheet 3: assumptions note
    const wsNotes = XLSX.utils.aoa_to_sheet([
      ["Notas de la exportación"],
      [""],
      [`Tasa de IPC utilizada: ${(cpiRate * 100).toFixed(1)}%`],
      [`Modo de precio: ${vatMode === "with_vat" ? "Con IVA" : "Sin IVA"}`],
      [`Fecha de exportación: ${new Date().toLocaleDateString("es-ES")}`],
      [""],
      ["Este archivo es una instantánea de valores calculados en el panel interactivo."],
      ["A diferencia del libro Excel original (RSCE_Price_Analysis.xlsx), estas celdas"],
      ["NO son fórmulas en vivo — reflejan el estado del panel en el momento de exportar,"],
      ["incluyendo cualquier anulación manual introducida en esta sesión."],
    ]);
    XLSX.utils.book_append_sheet(wb, wsNotes, "Notas");

    XLSX.writeFile(wb, `RSCE_Dashboard_Export_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [products, data, vatMode, cpiRate, overrides]);

  return (
    <div style={{
      fontFamily: "'Inter', system-ui, sans-serif",
      background: "#F6F3EC",
      minHeight: "100vh",
      color: "#20242C",
    }}>
      {/* Header */}
      <div style={{
        background: "#1C2B45",
        color: "#F6F3EC",
        padding: "20px 28px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        borderBottom: "3px solid #B98A3F",
      }}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
  <div
    style={{
      width: 42,
      height: 42,
      borderRadius: "50%",
      background: "#B98A3F",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      flexShrink: 0,
      marginTop: 2,
    }}
  >
    <PawPrint size={22} color="#1C2B45" />
  </div>

  <div
    style={{
      display: "flex",
      flexDirection: "column",
      alignItems: "flex-start",
      justifyContent: "center",
    }}
  >
    <div
      style={{
        fontFamily: "'Georgia', serif",
        fontSize: 20,
        fontWeight: 700,
        lineHeight: 1.1,
        marginBottom: 3,
      }}
    >
      Real Sociedad Canina de España
    </div>

    <div
      style={{
        fontSize: 12.5,
        opacity: 0.75,
        marginLeft: 0,
        paddingLeft: 0,
        textAlign: "left",
        alignSelf: "flex-start",
      }}
    >
      Panel de análisis de tarifas
    </div>
  </div>
</div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button
            onClick={handleExport}
            style={{
              display: "flex", alignItems: "center", gap: 8,
              background: "#B98A3F", color: "#1C2B45", border: "none",
              padding: "10px 18px", borderRadius: 6, fontWeight: 600, fontSize: 14,
              cursor: "pointer",
            }}
          >
            <Download size={16} /> Exportar a Excel
          </button>
          {saveStatus === "saving" && <Loader size={14} className="animate-spin" style={{ opacity: 0.7, color: "#F6F3EC" }} />}
          {saveStatus === "saved" && <CheckCircle size={14} color="#5C7A5E" />}
          {saveStatus === "error" && <span style={{ fontSize: 11, color: "#A6452E" }}>Error al guardar</span>}
        </div>
      </div>

      {/* Tab navigation */}
      <div style={{
        maxWidth: 1400, margin: "0 auto", padding: "14px 32px 0 32px",
        display: "flex", gap: 6, borderBottom: "1px solid #DDD6C6",
      }}>
        {[
          { key: "dashboard", label: "Panel de producto" },
          { key: "tarifaWeb", label: "Tarifa publicada (Web)" },
          { key: "colaboradoras", label: "Impreso Colaboradoras" },
        ].map((tab) => (
          <button
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            style={{
              padding: "10px 16px", border: "none", borderBottom: activeTab === tab.key ? "3px solid #B98A3F" : "3px solid transparent",
              background: "transparent", fontSize: 13.5, fontWeight: 600, cursor: "pointer",
              color: activeTab === tab.key ? "#1C2B45" : "#8A8474", fontFamily: "inherit",
              marginBottom: -1,
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {activeTab === "dashboard" && (
      <div style={{ display: "flex", maxWidth: 1400, margin: "0 auto" }}>
        {/* Sidebar */}
        <div style={{
          width: 320, flexShrink: 0, borderRight: "1px solid #DDD6C6",
          padding: "20px 16px", background: "#FBFAF6", minHeight: "calc(100vh - 82px)",
        }}>
          <div style={{ position: "relative", marginBottom: 12 }}>
            <Search size={16} style={{ position: "absolute", left: 10, top: 10, opacity: 0.5 }} />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Buscar producto..."
              style={{
                width: "100%", padding: "8px 10px 8px 32px", borderRadius: 6,
                border: "1px solid #D4CDBB", fontSize: 13.5, boxSizing: "border-box",
                fontFamily: "inherit",
              }}
            />
          </div>
          <select
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            style={{
              width: "100%", padding: "7px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
              fontSize: 13, marginBottom: 10, fontFamily: "inherit", background: "white", color: "#20242C",
            }}
          >
            <option value="Todas" style={{ color: "#20242C", background: "white" }}>Todas las categorías</option>
            {categories.map((c) => (
              <option key={c} value={c} style={{ color: "#20242C", background: "white" }}>{c}</option>
            ))}
          </select>

          <button
            onClick={() => setShowAddProduct((v) => !v)}
            style={{
              width: "100%", padding: "7px 8px", borderRadius: 6, border: "1px dashed #B98A3F",
              fontSize: 12.5, fontWeight: 600, marginBottom: 14, cursor: "pointer",
              background: "transparent", color: "#8C6B2E",
            }}
          >
            {showAddProduct ? "Cancelar" : "+ Añadir producto"}
          </button>

          <input
            type="file"
            accept=".xlsx"
            ref={fileInputRef}
            style={{ display: "none" }}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleImportExcel(f);
              e.target.value = ""; // allow re-selecting the same file later
            }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              width: "100%", padding: "7px 8px", borderRadius: 6, border: "1px dashed #1C2B45",
              fontSize: 12.5, fontWeight: 600, marginBottom: 8, cursor: "pointer",
              background: "transparent", color: "#1C2B45",
            }}
          >
            ⤴ Importar Excel (hoja "Data")
          </button>

          {showAddProduct && (
            <div style={{
              background: "#FFFBEA", border: "1px solid #E5DFD1", borderRadius: 8,
              padding: 10, marginBottom: 14,
            }}>
              <input
                value={newProductName}
                onChange={(e) => setNewProductName(e.target.value)}
                placeholder="Nombre del producto"
                style={{
                  width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                  fontSize: 12.5, marginBottom: 8, fontFamily: "inherit", boxSizing: "border-box",
                }}
              />

              {!addingNewCategory ? (
                <select
                  value={newProductCategory}
                  onChange={(e) => {
                    if (e.target.value === "__new__") setAddingNewCategory(true);
                    else setNewProductCategory(e.target.value);
                  }}
                  style={{
                    width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                    fontSize: 12.5, marginBottom: 8, fontFamily: "inherit", background: "white", color: "#20242C",
                  }}
                >
                  {categories.map((c) => (
                    <option key={c} value={c} style={{ color: "#20242C", background: "white" }}>{c}</option>
                  ))}
                  <option value="__new__" style={{ color: "#20242C", background: "white" }}>+ Nueva categoría…</option>
                </select>
              ) : (
                <input
                  value={newCategoryInput}
                  onChange={(e) => setNewCategoryInput(e.target.value)}
                  placeholder="Nombre de la nueva categoría"
                  style={{
                    width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                    fontSize: 12.5, marginBottom: 8, fontFamily: "inherit", boxSizing: "border-box",
                  }}
                />
              )}

              <button
                onClick={handleAddProduct}
                style={{
                  width: "100%", padding: "7px 8px", borderRadius: 6, border: "none",
                  background: "#B98A3F", color: "#1C2B45", fontWeight: 700, fontSize: 12.5,
                  cursor: "pointer",
                }}
              >
                Guardar producto
              </button>
            </div>
          )}

          <div style={{ fontSize: 11.5, opacity: 0.6, marginBottom: 8, fontWeight: 600 }}>
            {filteredProducts.length} PRODUCTOS
          </div>
          <div style={{ maxHeight: "calc(100vh - 250px)", overflowY: "auto" }}>
            {filteredProducts.map((p) => (
              <div
                key={p}
                onClick={() => setSelected(p)}
                style={{
                  padding: "9px 10px",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 13,
                  marginBottom: 3,
                  lineHeight: 1.35,
                  background: p === selected ? "#1C2B45" : "transparent",
                  color: p === selected ? "#F6F3EC" : "#20242C",

                  textAlign: "left",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  whiteSpace: "normal",
                  wordBreak: "break-word",
}}
              >
                {p}
              </div>
            ))}
          </div>
        </div>

       {/* Main content */}
        <div style={{ flex: 1, padding: "24px 32px", minWidth: 0 }}>
          {selectedRecord && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, height: 118, overflow: "hidden" }}>                <div style={{ minWidth: 0, paddingRight: 16 }}>
                  <div style={{
                    fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8, color: "#B98A3F", fontWeight: 700,
                    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                  }}>
                    {selectedRecord.category}
                  </div>
                  <h1 style={{
                    fontFamily: "'Georgia', serif", fontSize: 24, margin: "2px 0 0 0", lineHeight: 1.25,
                    color: "#20242C",
                    height: 60, overflow: "hidden", display: "flex", alignItems: "flex-start",
                    display: "-webkit-box", WebkitLineClamp: 3, WebkitBoxOrient: "vertical",
                    whiteSpace: "normal", wordBreak: "break-word",
                  }}>
                    {selected}
                  </h1>
                </div>
                <div style={{ display: "flex", gap: 4, background: "#EFEAE0", borderRadius: 8, padding: 3, flexShrink: 0 }}>
                  <button
                    onClick={() => setVatMode("with_vat")}
                    style={{
                      padding: "6px 14px", borderRadius: 6, border: "none", fontSize: 12.5, fontWeight: 600,
                      cursor: "pointer",
                      background: vatMode === "with_vat" ? "#1C2B45" : "transparent",
                      color: vatMode === "with_vat" ? "#F6F3EC" : "#20242C",
                    }}
                  >Con IVA</button>
                  <button
                    onClick={() => setVatMode("no_vat")}
                    style={{
                      padding: "6px 14px", borderRadius: 6, border: "none", fontSize: 12.5, fontWeight: 600,
                      cursor: "pointer",
                      background: vatMode === "no_vat" ? "#1C2B45" : "transparent",
                      color: vatMode === "no_vat" ? "#F6F3EC" : "#20242C",
                    }}
                  >Sin IVA</button>
                </div>
              </div>

              {/* Stat cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, margin: "18px 0" }}>
                {CT_ORDER.map((ct) => {
                  const s = perTypeStats[ct];
                  return (
                    <div key={ct} style={{
                      background: "white", borderRadius: 10, padding: "14px 16px",
                      border: "1px solid #E5DFD1", borderTop: `3px solid ${CT_COLORS[ct]}`,
                    }}>
                      <div style={{ fontSize: 11.5, fontWeight: 700, opacity: 0.65, textTransform: "uppercase", letterSpacing: 0.4 }}>
                        {CT_LABELS[ct]}
                      </div>
                      <div style={{ fontSize: 22, fontWeight: 700, margin: "4px 0 2px 0", fontFamily: "'Georgia', serif" }}>
                        {fmtEUR(s?.last ? s.last[vatMode] : null)}
                      </div>
                      <div style={{ fontSize: 12, opacity: 0.6, marginBottom: 6 }}>
                        {s?.last ? `Año base: ${s.last.year}` : "—"}
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                        <TrendIcon v={s?.yoy} />
                        <span>{s?.yoy !== null && s?.yoy !== undefined ? `${fmtPct(s.yoy)} interanual` : "Sin comparación previa"}</span>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Chart */}
              <div style={{ background: "white", borderRadius: 10, border: "1px solid #E5DFD1", padding: "18px 20px", marginBottom: 18 }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, opacity: 0.75 }}>
                  Evolución de precios ({vatMode === "with_vat" ? "con IVA" : "sin IVA"})
                </div>
                <ResponsiveContainer width="100%" height={260}>
                  <LineChart data={chartData} margin={{ top: 5, right: 20, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#EEE8DA" />
                    <XAxis dataKey="year" tick={{ fontSize: 12 }} stroke="#8A8474" />
                    <YAxis tick={{ fontSize: 12 }} stroke="#8A8474" width={50} />
                    <Tooltip formatter={(v, name) => [fmtEUR(v), CT_LABELS[name] || name]} labelStyle={{ fontWeight: 600 }} />                    
                    <Legend formatter={(v) => CT_LABELS[v]} wrapperStyle={{ fontSize: 12.5 }} />
                    {CT_ORDER.map((ct) => (
                      <Line
                        key={ct} type="monotone" dataKey={ct} stroke={CT_COLORS[ct]}
                        strokeWidth={2.5} dot={{ r: 3 }} connectNulls={false}
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              </div>

              {/* Editar precios (año seleccionable) */}
              <div style={{ background: "white", borderRadius: 10, border: "1px solid #E5DFD1", padding: "18px 20px", marginBottom: 18 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, opacity: 0.75 }}>Editar precios</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <button
                      onClick={handleAddYear}
                      title="Crea el siguiente año en el histórico, disponible para todos los productos"
                      style={{
                        padding: "5px 10px", borderRadius: 6, border: "1px dashed #B98A3F",
                        fontSize: 12, fontWeight: 600, cursor: "pointer", background: "transparent",
                        color: "#8C6B2E",
                      }}
                    >
                      + Crear año {Math.max(...years) + 1}
                    </button>
                <label style={{ fontSize: 12.5, opacity: 0.7 }}>IPC:</label>
                <input
                  type="text"
                  inputMode="decimal"
                  value={cpiDraft}
                  onChange={(e) => setCpiDraft(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") e.target.blur();
                  }}
                  onBlur={() => {
                    const val = parseFloat(cpiDraft.replace(",", "."));
                    if (!Number.isNaN(val) && val >= 0) {
                      setCpiRate(val / 100);
                    } else {
                      setCpiDraft((cpiRate * 100).toFixed(1));
                    }
                  }}
                  style={{
                    width: 55, padding: "5px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                    fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C",
                    textAlign: "right",
                  }}
                />
                      <span style={{ fontSize: 12.5, opacity: 0.7 }}>%</span>                    
                      <select
                      value={editYear}
                      onChange={(e) => setEditYear(Number(e.target.value))}
                      style={{
                        padding: "5px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                        fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C",
                      }}
                    >
                      {years.map((y) => (
                        <option key={y} value={y} style={{ color: "#20242C", background: "white" }}>
                          {y}{y === years[years.length - 1] ? " (más reciente)" : ""}
                        </option>
                      ))}
                    </select>
                    <button
                      onClick={() => handleRemoveYear(editYear)}
                      disabled={years.length <= 1}
                      title="Elimina este año del histórico para todos los productos"
                      style={{
                        padding: "5px 10px", borderRadius: 6, border: "1px solid #D4CDBB",
                        fontSize: 12, fontWeight: 600,
                        cursor: years.length > 1 ? "pointer" : "not-allowed",
                        background: "white", color: years.length > 1 ? "#A6452E" : "#C9C2B0",
                      }}
                    >
                      Eliminar año
                    </button>
                  </div>
                </div>

              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, tableLayout: "fixed" }}>
                  <colgroup>
                    <col style={{ width: "30%" }} />
                    <col style={{ width: "14%" }} />
                    <col style={{ width: "14%" }} />
                    <col style={{ width: "42%" }} />
                  </colgroup>
                  <thead>
                    <tr style={{ textAlign: "left", opacity: 0.6, fontSize: 11.5, textTransform: "uppercase" }}>
                      <th style={{ paddingBottom: 6, paddingRight: 12, textAlign: "left" }}>Tipo</th>
                      <th style={{ paddingBottom: 6, paddingRight: 12, textAlign: "left" }}>Con IVA</th>
                      <th style={{ paddingBottom: 6, paddingRight: 12, textAlign: "left" }}>Sin IVA</th>
                      <th style={{ paddingBottom: 6, textAlign: "left" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {CT_ORDER.map((ct) => {
                      const series = selectedRecord.prices[ct] || [];
                      const entry = series.find((p) => p.year === editYear);
                      const prev = prevYearEntry(series, editYear);

                      const applyForecast = () => {
                        if (!prev) return;
                        if (prev.with_vat !== null && prev.with_vat !== undefined) {
                          updatePrice(selected, ct, "with_vat", (Math.round(prev.with_vat * (1 + cpiRate) * 100) / 100).toString(), editYear);
                        }
                        if (prev.no_vat !== null && prev.no_vat !== undefined) {
                          updatePrice(selected, ct, "no_vat", (Math.round(prev.no_vat * (1 + cpiRate) * 100) / 100).toString(), editYear);
                        }
                      };

                      return (
                        <tr key={ct} style={{ borderTop: "1px solid #F0EBDD" }}>
                          <td style={{ padding: "8px 0", fontWeight: 600 }}>
                            <span style={{ display: "inline-block", width: 8, height: 8, borderRadius: "50%", background: CT_COLORS[ct], marginRight: 7 }} />
                            {CT_LABELS[ct]}
                          </td>
                          <td style={{ paddingRight: 12 }}>
                            <input
                              key={`${selected}-${ct}-${editYear}-wv-${entry?.with_vat ?? "empty"}`}
                              type="text" inputMode="decimal"
                              defaultValue={entry?.with_vat ?? ""}
                              placeholder="—"
                              onBlur={(e) => updatePrice(selected, ct, "with_vat", e.target.value, editYear)}
                              style={{width: "100%", maxWidth: 90, boxSizing: "border-box",
                              padding: "5px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                              fontSize: 13, background: "#FFFBEA", fontFamily: "inherit", color: "#20242C", }}
                            />
                          </td>
                          <td style={{ paddingRight: 20 }}>
                            <input
                              key={`${selected}-${ct}-${editYear}-nv-${entry?.no_vat ?? "empty"}`}
                              type="text" inputMode="decimal"
                              defaultValue={entry?.no_vat ?? ""}
                              placeholder="—"
                              onBlur={(e) => updatePrice(selected, ct, "no_vat", e.target.value, editYear)}
                              style={{ width: 90, padding: "5px 8px", borderRadius: 6, border: "1px solid #D4CDBB",
                                fontSize: 13, background: "#FFFBEA", fontFamily: "inherit", color: "#20242C" }}
                            />
                          </td>
                          <td style={{ paddingLeft: 20 }}>
                            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                              <button
                                onClick={applyForecast}
                                disabled={!prev}
                                title={prev ? `Usar ${fmtEUR(prev.with_vat)} × (1+IPC)` : "Sin año anterior para calcular"}
                                style={{
                                  padding: "5px 10px", borderRadius: 6, border: "1px solid #D4CDBB",
                                  fontSize: 12, fontWeight: 600, cursor: prev ? "pointer" : "not-allowed",
                                  background: prev ? "#EFEAE0" : "#F5F2E9", color: prev ? "#1C2B45" : "#A8A08C",
                                  whiteSpace: "nowrap", flexShrink: 0,
                                }}
                              >
                                Usar previsión IPC
                              </button>
                              <button
                                onClick={() => {
                                  updatePrice(selected, ct, "with_vat", "", editYear);
                                  updatePrice(selected, ct, "no_vat", "", editYear);
                                }}
                                disabled={!entry}
                                title="Borra el precio de este tipo para este año (vuelve a quedar sin datos)"
                                style={{
                                  padding: "5px 10px", borderRadius: 6, border: "1px solid #D4CDBB",
                                  fontSize: 12, fontWeight: 600, cursor: entry ? "pointer" : "not-allowed",
                                  background: "white", color: entry ? "#A6452E" : "#C9C2B0",
                                  whiteSpace: "nowrap", flexShrink: 0,
                                }}
                              >
                                Vaciar
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Cross-type comparison + rosette badge */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr auto", gap: 12, marginBottom: 18, alignItems: "stretch" }}>
                <ComparisonCard label="Usuario vs Socio" value={crossTypeVariation.userVsMember} />
                <ComparisonCard label="Colaboradora vs Socio" value={crossTypeVariation.collabVsMember} />
                <ComparisonCard label="Colaboradora vs Usuario" value={crossTypeVariation.collabVsUser} />
                <RosetteBadge value={crossTypeVariation.userVsMember} />
              </div>

              {/* Comparación personalizada (full-width, matches Editar precios styling) */}
              <div style={{ background: "white", borderRadius: 10, border: "1px solid #E5DFD1", padding: "18px 20px", marginBottom: 18 }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14, opacity: 0.75 }}>
                  Comparación personalizada
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#FBFAF6", border: "1px solid #EEE8DA", borderRadius: 8, padding: "6px 10px" }}>
                    <select
                      value={compareTypeA}
                      onChange={(e) => setCompareTypeA(e.target.value)}
                      style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
                    >
                      {CT_ORDER.map((ct) => (
                        <option key={ct} value={ct} style={{ color: "#20242C", background: "white" }}>{CT_LABELS[ct]}</option>
                      ))}
                    </select>
                    <select
                      value={compareYearA}
                      onChange={(e) => setCompareYearA(Number(e.target.value))}
                      style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
                    >
                      {years.map((y) => (
                        <option key={y} value={y} style={{ color: "#20242C", background: "white" }}>{y}</option>
                      ))}
                    </select>
                  </div>

                  <span style={{ fontSize: 13, opacity: 0.6, fontWeight: 600 }}>vs</span>

                  <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#FBFAF6", border: "1px solid #EEE8DA", borderRadius: 8, padding: "6px 10px" }}>
                    <select
                      value={compareTypeB}
                      onChange={(e) => setCompareTypeB(e.target.value)}
                      style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
                    >
                      {CT_ORDER.map((ct) => (
                        <option key={ct} value={ct} style={{ color: "#20242C", background: "white" }}>{CT_LABELS[ct]}</option>
                      ))}
                    </select>
                    <select
                      value={compareYearB}
                      onChange={(e) => setCompareYearB(Number(e.target.value))}
                      style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
                    >
                      {years.map((y) => (
                        <option key={y} value={y} style={{ color: "#20242C", background: "white" }}>{y}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {customComparison && (
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ textAlign: "left", opacity: 0.6, fontSize: 11.5, textTransform: "uppercase" }}>
                        <th style={{ paddingBottom: 6, textAlign: "left" }}>{CT_LABELS[compareTypeA]} ({compareYearA})</th>
                        <th style={{ paddingBottom: 6, textAlign: "left" }}>{CT_LABELS[compareTypeB]} ({compareYearB})</th>
                        <th style={{ paddingBottom: 6, textAlign: "left" }}>Diferencia (€)</th>
                        <th style={{ paddingBottom: 6, textAlign: "left" }}>Variación (%)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ borderTop: "1px solid #F0EBDD" }}>
                        <td style={{ padding: "8px 0", fontWeight: 600 }}>{fmtEUR(customComparison.valA)}</td>
                        <td style={{ padding: "8px 0", fontWeight: 600 }}>{fmtEUR(customComparison.valB)}</td>
                        <td style={{ padding: "8px 0" }}>
                          {customComparison.diff !== null ? fmtEUR(customComparison.diff) : "—"}
                        </td>
                        <td style={{ padding: "8px 0", display: "flex", alignItems: "center", gap: 5 }}>
                          <TrendIcon v={customComparison.pct} />
                          {fmtPct(customComparison.pct)}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                )}
              </div>
            </>
          )}
        </div>
      </div>
      )}

      {activeTab === "tarifaWeb" && (
        <TarifaWebTab
          products={products}
          categories={categories}
          data={data}
          years={years}
          year={tarifaWebYear}
          setYear={setTarifaWebYear}
        />
      )}

      {activeTab === "colaboradoras" && (
        <ColaboradorasTab
          products={products}
          categories={categories}
          data={data}
          years={years}
          year={colaboradorasYear}
          setYear={setColaboradorasYear}
          productToCode={productToCode}
        />
      )}
    </div>
  );
}


     
function TarifaWebTab({ products, categories, data, years, year, setYear }) {
  const rowsByCategory = useMemo(() => {
    const byCat = {};
    for (const p of products) {
      const rec = data[p];
      const cat = rec.category || "Sin categorizar";
      const memberEntry = valueForYear(rec.prices.Member || [], "with_vat", year);
      const userEntry = valueForYear(rec.prices.User || [], "with_vat", year);
      if (!memberEntry && !userEntry) continue; // nothing to show for this year
      if (!byCat[cat]) byCat[cat] = [];
      byCat[cat].push({
        product: p,
        member: memberEntry ? memberEntry.with_vat : null,
        user: userEntry ? userEntry.with_vat : null,
      });
    }
    return byCat;
  }, [products, data, year]);

  const orderedCategories = categories.filter((c) => rowsByCategory[c] && rowsByCategory[c].length > 0);

  const handleExport = useCallback(() => {
    const rows = [["Concepto", "Socios RSCE", "Resto de Usuarios"]];
    for (const cat of orderedCategories) {
      rows.push([cat, "", ""]);
      for (const r of rowsByCategory[cat]) {
        rows.push([r.product, r.member ?? "", r.user ?? ""]);
      }
      rows.push(["", "", ""]);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tarifa Web");
    XLSX.writeFile(wb, `Tarifa_Publicada_Web_${year}.xlsx`);
  }, [orderedCategories, rowsByCategory, year]);

  return (
    <div style={{ maxWidth: 1400, margin: "0 auto", padding: "24px 32px" }}>
      <div style={{ background: "white", borderRadius: 10, border: "1px solid #E5DFD1", padding: "18px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "'Georgia', serif", color: "#20242C" }}>
              Tarifa publicada (Web)
            </div>
            <div style={{ fontSize: 12, opacity: 0.6, marginTop: 2 }}>
              Precios con IVA — vista de solo lectura, generada a partir de los datos actuales del panel
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <select
              value={year}
              onChange={(e) => setYear(Number(e.target.value))}
              style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
            >
              {years.map((y) => (
                <option key={y} value={y} style={{ color: "#20242C", background: "white" }}>{y}</option>
              ))}
            </select>
            <button
              onClick={handleExport}
              style={{
                display: "flex", alignItems: "center", gap: 8,
                background: "#B98A3F", color: "#1C2B45", border: "none",
                padding: "8px 14px", borderRadius: 6, fontWeight: 600, fontSize: 13,
                cursor: "pointer",
              }}
            >
              <Download size={15} /> Exportar a Excel
            </button>
          </div>
        </div>

        {orderedCategories.length === 0 ? (
          <div style={{ fontSize: 13, opacity: 0.6, padding: "20px 0" }}>
            No hay precios registrados para {year}.
          </div>
        ) : (
          orderedCategories.map((cat) => (
            <div key={cat} style={{ marginBottom: 20 }}>
              <div style={{
                fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5,
                color: "#B98A3F", padding: "8px 0", borderBottom: "2px solid #EFEAE0", marginBottom: 6,
              }}>
                {cat}
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ textAlign: "left", opacity: 0.55, fontSize: 11, textTransform: "uppercase" }}>
                    <th style={{ padding: "4px 0", width: "60%", textAlign: "left" }}>Concepto</th>
                    <th style={{ padding: "4px 0", textAlign: "left" }}>Socios RSCE</th>
                    <th style={{ padding: "4px 0", textAlign: "left" }}>Resto de Usuarios</th>
                  </tr>
                </thead>
                <tbody>
                  {rowsByCategory[cat].map((r) => (
                    <tr key={r.product} style={{ borderTop: "1px solid #F5F1E6" }}>
                      <td style={{ padding: "6px 0", paddingRight: 12 }}>{r.product}</td>
                      <td style={{ padding: "6px 0", fontWeight: 600 }}>{r.member !== null ? fmtEUR(r.member) : "—"}</td>
                      <td style={{ padding: "6px 0", fontWeight: 600 }}>{r.user !== null ? fmtEUR(r.user) : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

function ColaboradorasTab({ products, categories, data, years, year, setYear, productToCode }) {
  const rowsByCategory = useMemo(() => {
    const byCat = {};
    for (const p of products) {
      const rec = data[p];
      const cat = rec.category || "Sin categorizar";
      const entry = valueForYear(rec.prices.Canine_Collaborator || [], "no_vat", year);
      if (!entry) continue; // only list items that apply to Colaboradoras Caninas
      if (!byCat[cat]) byCat[cat] = [];
      byCat[cat].push({
        product: p,
        code: productToCode[p] || "—",
        price: entry.no_vat,
      });
    }
    return byCat;
  }, [products, data, year, productToCode]);

  const orderedCategories = categories.filter((c) => rowsByCategory[c] && rowsByCategory[c].length > 0);

  const handleExport = useCallback(() => {
    const rows = [["Concepto", "Código", "Precio"]];
    for (const cat of orderedCategories) {
      rows.push([cat, "", ""]);
      for (const r of rowsByCategory[cat]) {
        rows.push([r.product, r.code, r.price]);
      }
      rows.push(["", "", ""]);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Impreso Colaboradoras");
    XLSX.writeFile(wb, `Impreso_Colaboradoras_${year}.xlsx`);
  }, [orderedCategories, rowsByCategory, year]);

  return (
    <div style={{ maxWidth: 1400, margin: "0 auto", padding: "24px 32px" }}>
      <div style={{ background: "white", borderRadius: 10, border: "1px solid #E5DFD1", padding: "18px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "'Georgia', serif", color: "#20242C" }}>
              Impreso Colaboradoras
            </div>
            <div style={{ fontSize: 12, opacity: 0.6, marginTop: 2 }}>
              Precios sin IVA (Colaboradoras Caninas) — vista de solo lectura, generada a partir de los datos actuales del panel
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <select
              value={year}
              onChange={(e) => setYear(Number(e.target.value))}
              style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #D4CDBB", fontSize: 13, fontFamily: "inherit", background: "white", color: "#20242C" }}
            >
              {years.map((y) => (
                <option key={y} value={y} style={{ color: "#20242C", background: "white" }}>{y}</option>
              ))}
            </select>
            <button
              onClick={handleExport}
              style={{
                display: "flex", alignItems: "center", gap: 8,
                background: "#B98A3F", color: "#1C2B45", border: "none",
                padding: "8px 14px", borderRadius: 6, fontWeight: 600, fontSize: 13,
                cursor: "pointer",
              }}
            >
              <Download size={15} /> Exportar a Excel
            </button>
          </div>
        </div>

        {orderedCategories.length === 0 ? (
          <div style={{ fontSize: 13, opacity: 0.6, padding: "20px 0" }}>
            No hay precios de Colaboradoras Caninas registrados para {year}.
          </div>
        ) : (
          orderedCategories.map((cat) => (
            <div key={cat} style={{ marginBottom: 20 }}>
              <div style={{
                fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5,
                color: "#5C7A5E", padding: "8px 0", borderBottom: "2px solid #EFEAE0", marginBottom: 6,
              }}>
                {cat}
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ textAlign: "left", opacity: 0.55, fontSize: 11, textTransform: "uppercase" }}>
                    <th style={{ padding: "4px 0", width: "60%", textAlign: "left" }}>Concepto</th>
                    <th style={{ padding: "4px 0", textAlign: "left" }}>Código</th>
                    <th style={{ padding: "4px 0", textAlign: "left" }}>Precio</th>
                  </tr>
                </thead>
                <tbody>
                  {rowsByCategory[cat].map((r) => (
                    <tr key={r.product} style={{ borderTop: "1px solid #F5F1E6" }}>
                      <td style={{ padding: "6px 0", paddingRight: 12 }}>{r.product}</td>
                      <td style={{ padding: "6px 0", opacity: 0.75 }}>{r.code}</td>
                      <td style={{ padding: "6px 0", fontWeight: 600 }}>{fmtEUR(r.price)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

function ComparisonCard({ label, value }) {
  return (
    <div style={{
      background: "white", borderRadius: 10, border: "1px solid #E5DFD1",
      padding: "14px 16px", display: "flex", flexDirection: "column", justifyContent: "center",
    }}>
      <div style={{ fontSize: 11.5, opacity: 0.6, marginBottom: 6 }}>{label}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 18, fontWeight: 700, fontFamily: "'Georgia', serif" }}>
        <TrendIcon v={value} /> {fmtPct(value)}
      </div>
    </div>
  );
}

function RosetteBadge({ value }) {
  const pct = value !== null && value !== undefined ? Math.abs(value * 100).toFixed(0) : "—";
  return (
    <div style={{
      width: 92, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
      position: "relative",
    }}>
      <div style={{
        width: 78, height: 78, borderRadius: "50%",
        background: "linear-gradient(135deg, #B98A3F, #8C6B2E)",
        display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
        color: "#FBFAF6", boxShadow: "0 3px 8px rgba(28,43,69,0.25)",
      }}>
        <Award size={16} />
        <div style={{ fontSize: 15, fontWeight: 800, lineHeight: 1 }}>{pct}%</div>
      </div>
      <div style={{ fontSize: 9.5, textAlign: "center", opacity: 0.65, marginTop: 4, lineHeight: 1.2 }}>
        AHORRO
        <br />SOCIO
      </div>
    </div>
  );
}

export default function App() {
  return (
    <PasswordGate>
      <RSCEDashboard />
    </PasswordGate>
  );
}