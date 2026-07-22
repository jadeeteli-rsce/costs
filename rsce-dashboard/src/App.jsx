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

const CODE_MAP = [{"code": "11100001", "name": "INSCRIPCIÓN CACHORRO PREMIUM LOE/RRC"}, {"code": "11100002", "name": "INSCRIPCIÓN PERRO IMPORTADOS"}, {"code": "11100003", "name": "INSCRIPCION PERROS DE OTROS LIBROS GENEALÓGICOS"}, {"code": "11100004", "name": "INSCRIPCIÓN CACHORRO RAZAS ESPAÑOLAS BONIFICADAS"}, {"code": "11100005", "name": "INSCRIPCION CACHORRO GRUPOS ETNICOS"}, {"code": "11100006", "name": "BAJA DE UN PERRO EN EL L.O.E. / R.R.C."}, {"code": "11100007", "name": "INSCRIPCIÓN CACHORRO ACCESS LBO/RBR"}, {"code": "11200006", "name": "PEDIGREE_RSCE DE EXPORTACIÓN TRES GENERACIONES"}, {"code": "11200007", "name": "PEDIGREE_RSCE DE EXPORTACIÓN CUATRO GENERACIONES"}, {"code": "11200008", "name": "VISADO DE EXPORTACION"}, {"code": "11200011", "name": "DUPLICADO PEDIGREE_RSCE EXPORTACION HASTA TRES GENERACIONES"}, {"code": "11200012", "name": "DUPLICADO PEDIGREE_RSCE DE EXPORTACION CUATRO GENERACIONES"}, {"code": "11200019", "name": "PEDIGREE PREMIUM 3 GENERACIONES CON TRANSFERENCIA EN PAPEL"}, {"code": "11200020", "name": "PEDIGREE PREMIUM 4 GENERACIONES CON TRANSFERENCIA EN PAPEL"}, {"code": "11200021", "name": "PEDIGREE PREMIUM RRC CON TRANSFERENCIA EN PAPEL"}, {"code": "11200022", "name": "PEDIGREE PREMIUM 3 GENERACIONES SIN TRANSFERENCIA EN PAPEL"}, {"code": "11200023", "name": "PEDIGREE PREMIUM 4 GENERACIONES SIN TRANSFERENCIA EN PAPEL"}, {"code": "11200024", "name": "MEJORA A PEDIGREE PLUS 3 GENERACIONES CON TRANSFERENCIA EN P"}, {"code": "11200025", "name": "MAJORA A PEDIGREE PLUS 4 GENERACIONES CON TRANSFERENCIA EN P"}, {"code": "11200026", "name": "MEJORA A PEDIGREE PLUS RRC SIN TRANSFERENCIA EN PAPEL"}, {"code": "11200027", "name": "MEJORA A PEDIGREE PLUS 3 GENERACIONES SIN TRANSFERENCIA EN P"}, {"code": "11200028", "name": "MEJORA A PEDIGREE PLUS 4 GENERACIONES SIN TRANSFERENCIA EN P"}, {"code": "11200029", "name": "PEDIGREE ACCESS 3 GENERACIONES CON TRANSFERENCIA EN PAPEL"}, {"code": "11200030", "name": "PEDIGREE ACCESS RBR CON TRANSFERENCIA EN PAPEL"}, {"code": "11200031", "name": "PEDIGREE ACCESS SIN TRANSFERENCIA EN PAPEL"}, {"code": "11200032", "name": "MEJORA DE PEDIGREE ACCESS A PREMIUM 3 GEN. S/TRF. EN PAPEL"}, {"code": "11200033", "name": "MEJORA DE PEDIGREE ACCESS A PREMIUM 4 GEN. S/TRF. EN PAPEL"}, {"code": "11200034", "name": "MEJORA DE PEDIGREE ACCESS RBR A PREMIUM RRC S/TRF. EN PAPEL"}, {"code": "11200035", "name": "DUPLICADO PEDIGREE EN PAPEL"}, {"code": "11300001", "name": "TRANSFERENCIA PROPIEDAD EN REGISTRO INICIAL/PERRO IMPORTADO"}, {"code": "11300002", "name": "CESION TEMPORAL DE PROPIEDAD DE LOS REPRODUCTORES"}, {"code": "11300003", "name": "TRANSFERENCIA DE PROPIEDAD - GRUPOS ETNICOS"}, {"code": "11400001", "name": "RECARGO 100% INSCRIPCION CACHORRO PREMIUM LOE/RRC"}, {"code": "11400002", "name": "RECARGO 200% INSCRIPCION CACHORRO PREMIUM LOE/RRC"}, {"code": "11400003", "name": "RECARGO 300% INSCRIPCION CACHORRO PREMIUM LOE/RRC"}, {"code": "11400004", "name": "INSCRIPCIÓN PERRO MÁS 18 MESES - APORTANDO PRUEBA ADN"}, {"code": "11400005", "name": "RECARGO 100% INSCRIPCION CACHORRO ACCESS LBO/RBR"}, {"code": "11400006", "name": "RECARGO 200% INSCRIPCION CACHORRO ACCESS LBO/RBR"}, {"code": "11400007", "name": "RECARGO 300% INSCRIPCION CACHORRO ACCESS LBO/RBR"}, {"code": "11500001", "name": "SOLICITUD DE AFIJO"}, {"code": "11500002", "name": "DUPLICADO DE AFIJO"}, {"code": "11500003", "name": "RECTIFICACIÓN DE AFIJO"}, {"code": "11600001", "name": "REGISTRO INICIAL DE RAZAS ESPAÑOLAS"}, {"code": "11600002", "name": "REGISTRO INICIAL DE RAZA"}, {"code": "11600003", "name": "REGISTRO INICIAL - GRUPOS ETNICOS"}, {"code": "11600004", "name": "REGISTRO INICIAL RAZAS ESPAÑOLAS VULNERABLES"}, {"code": "11700001", "name": "CONFIRMACION DE RAZA (Perros de razas españolas)"}, {"code": "11700002", "name": "CONFIRMACION DE RAZA - GRUPOS ETNICOS"}, {"code": "11800001", "name": "IMPRESO ALTA CAMADA"}, {"code": "11800002", "name": "IMPRESO ALTA CAMADA GRUPOS ETNICOS"}, {"code": "11800003", "name": "IMPRESO DIPLOMA DE INICIACIÓN"}, {"code": "11800004", "name": "IMPRESO HOJA DE CALIFICACION - PRUEBAS PERRO DE AG"}, {"code": "11800005", "name": "IMPRESO PARA PRUEBAS DE C.A.C.T. - C.A.C.I.T."}, {"code": "11800006", "name": "IMPRESO PRUEBAS APTITUDES NATURALES"}, {"code": "11800007", "name": "INFORMES SOLICITADOS A LA R.S.C.E. (por folio impreso)"}, {"code": "11900001", "name": "FORMULARIO DE IDENTIFICACIÓN R.S.C.E"}, {"code": "11900002", "name": "FORMULARIO IDENTIFICACION EN ORÍGEN"}, {"code": "12000001", "name": "DUPLICADO JUSTIFICANTES DE INSCRIPCIÓN"}, {"code": "12100001", "name": "TRAMITACIÓN URGENTE INSCRIPCIONES, PEDIGRÍES Y TRANSFERENCIA"}, {"code": "12100002", "name": "PLUS POR TRAMITACIÓN URGENTE - CARTILLA"}, {"code": "12100003", "name": "PLUS POR TRAMITACIÓN URGENTE - CERTIFICADO COLA CORTA"}, {"code": "12100004", "name": "PLUS POR TRAMITACIÓN URGENTE - CERTIFICADO TRABAJO"}, {"code": "12200001", "name": "CANON POR PERRO INSCRITO EN CAC - MORFOLOGÍA"}, {"code": "12200013", "name": "CANON PUNTO OBLIGATORIO EN MONOGRAFICAS CON CAC"}, {"code": "12200014", "name": "CANON CAC A CLUBES RECONOCIDOS"}, {"code": "12300001", "name": "TITULO CAMPEÓN EN ESPAÑA DE LA RSCE JOVEN"}, {"code": "12300002", "name": "TITULO SUPERCAMPEON EN ESPAÑA DE LA RSCE"}, {"code": "12300003", "name": "TITULOS FCI JOVEN"}, {"code": "12300004", "name": "CERTIFICADO DE COLA CORTA"}, {"code": "12300005", "name": "TITULO LATIN CHAMPION"}, {"code": "12300006", "name": "TITULO LATIN WINNER"}, {"code": "12300007", "name": "ANOTACIONES EN EL LOE TITULOS CAMPEÓN DE ONC DE LA FCI"}, {"code": "12300008", "name": "TITULO DE CAMPEÓN ESPAÑA ADULTO"}, {"code": "12300009", "name": "TITULO DE CAMPEÓN ESPAÑA VETERANO"}, {"code": "12300010", "name": "TITULO DE CAMPEÓN IBERICO JOVEN"}, {"code": "12300011", "name": "TITULO DE CAMPEÓN IBERICO ADULTO"}, {"code": "12300012", "name": "TITULOS FCI ADULTO BELLEZA"}, {"code": "12300013", "name": "TITULOS FCI VETERANO"}, {"code": "12300014", "name": "TITULO CAMPEÓN ESPAÑA DE LA RSCE ADULTO"}, {"code": "12300015", "name": "TITULO CAMPEÓN EN ESPAÑA RSCE VETERANO"}, {"code": "12300016", "name": "TRAMITACION TITULOS FCI ADULTO EXPOSICION"}, {"code": "19900001", "name": "DERECHOS EXAMEN JUECES"}, {"code": "19900002", "name": "PLUS ENVIO CARTILLA POR MENSAJERIA"}, {"code": "19900003", "name": "REEMBOLSO GASTOS BANCARIOS"}, {"code": "19900004", "name": "REEMBOLSO GASTOS DE ENVIO EUROPA"}, {"code": "19900005", "name": "REEMBOLSO GASTOS DE ENVIO MUNDIAL"}, {"code": "19900006", "name": "REEMBOLSO GASTOS DE ENVIO PENINSULA"}, {"code": "19900007", "name": "TEMARIO JUECES"}, {"code": "19900008", "name": "REEMBOLSO GASTOS DE ENVIO BALEARES"}, {"code": "19900009", "name": "REEMBOLSO GASTOS DE ENVIO CANARIAS"}, {"code": "19900010", "name": "DERECHOS EXAMEN A JUEZ NACIONAL DE CONCURSO"}, {"code": "19900011", "name": "DERECHOS EXAMEN AMPLIACIÓN RAZAS"}, {"code": "19900012", "name": "DERECHOS EXAMEN JUEZ INTERNACIONAL PARA COMPLETAR GRUPO"}, {"code": "19900013", "name": "RECURSO COMISION LIBRO ORIGENES ESPAÑOL"}, {"code": "19900014", "name": "CURSO COMISARIO PRINCIPAL, RING O JUNIOR HANDLER"}, {"code": "19900015", "name": "CURSO JUECES"}, {"code": "19900016", "name": "DERECHOS EXAMEN JUECES IGP"}, {"code": "19900099", "name": "VARIOS"}, {"code": "19900999", "name": "VARIOS COPA ESPAÑA IGP"}, {"code": "19909999", "name": "VARIOS COPA ESPAÑA IBGH/StPr"}, {"code": "19910001", "name": "LIBRO DE RAZAS ESPAÑOLAS"}, {"code": "200000", "name": "INMOVILIZADO"}, {"code": "200001", "name": "INMOVILIZADO SIN IVA"}, {"code": "21000000", "name": "INSCRIPCIÓN PERRO/S EN EXPOSICIÓN INTERNACIONAL CANINA"}, {"code": "21000001", "name": "INSCRIPCIÓN 1º PERRO - LATIN WINNER"}, {"code": "21000002", "name": "INSCRIPCIÓN 2º PERRO - LATIN WINNER"}, {"code": "21000003", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES  - LATIN WINNER"}, {"code": "21000004", "name": "INSCRIPCIÓN 1º CACHORRO - LATIN WINNER"}, {"code": "21000005", "name": "INSCRIPCIÓN 2º CACHORRO - LATIN WINNER"}, {"code": "21000006", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES-LATIN WINNER"}, {"code": "21000007", "name": "INSCRIPCIÓN 1º MUY CACHORRO - LATIN WINNER"}, {"code": "21000008", "name": "INSCRIPCIÓN 2º MUY CACHORRO - LATIN WINNER"}, {"code": "21000009", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGS - LATIN WINNER"}, {"code": "21000010", "name": "INSCRIPCIÓN 1º VETERANO - LATIN WINNER"}, {"code": "21000011", "name": "INSCRIPCIÓN 2º VETERANO - LATIN WINNER"}, {"code": "21000012", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES-LATIN WINNER"}, {"code": "21000013", "name": "INSCRIPCIÓN PAREJAS - LATIN WINNER"}, {"code": "21000014", "name": "INSCRIPCIÓN GRUPOS DE CRÍA - LATIN WINNER"}, {"code": "21000015", "name": "REGISTRO INICIAL RAZAS ESPAÑOLAS - LATIN WINNER"}, {"code": "21000016", "name": "REGISTRO INICIAL RAZAS INTEGRADAS - LATIN WINNER"}, {"code": "21000017", "name": "DISPONIBLE"}, {"code": "21000018", "name": "INSCRIPCIÓN 1º PERRO - 2º PLAZO-LATIN WINNER"}, {"code": "21000019", "name": "INSCRIPCIÓN 2º PERRO - 2º PLAZO- LATIN WINNER"}, {"code": "21000020", "name": "INSCRIPCIÓN DESDE 3º PERRO- 2º PLAZO -LATIN WINNER"}, {"code": "21000021", "name": "INSCRIPCIÓN 1º CACHORRO- 2º PLAZO - LATIN WINNER"}, {"code": "21000022", "name": "INSCRIPCIÓN 2º CACHORRO - 2º PLAZO - LATIN WINNER"}, {"code": "21000023", "name": "INSCRIPCIÓN DESDE 3º CACHORRO - 2º PLAZO - LATIN WINNER"}, {"code": "21000024", "name": "INSCRIPCIÓN 1º MUY CACHORRO - 2º PLAZO - LATIN WINNER"}, {"code": "21000025", "name": "INSCRIPCIÓN 2º MUY CACHORRO - 2º PLAZO - LATIN WINNER"}, {"code": "21000026", "name": "INSCRIPCIÓN DESDE 3º MUY CACHORRO - 2º PLAZO - LATIN WINNER"}, {"code": "21000027", "name": "INSCRIPCIÓN 1º VETERANO - 2º PLAZO - LATIN WINNER"}, {"code": "21000028", "name": "INSCRIPCIÓN 2º VETERANO - 2º PLAZO - LATIN WINNER"}, {"code": "21000029", "name": "INSCRIPCIÓN DESDE 3º VETERANO - 2º PLAZO - LATIN WINNER"}, {"code": "21000030", "name": "INSCRIPCIÓN 1º PERRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000031", "name": "INSCRIPCIÓN 2º PERRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000032", "name": "INSCRIPCIÓN DESDE 3º PERRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000033", "name": "INSCRIPCIÓN 1º CACHORRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000034", "name": "INSCRIPCIÓN 2º CACHORRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000035", "name": "INSCRIPCIÓN DESDE 3º CACHORRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000036", "name": "INSCRIPCIÓN 1º MUY CACHORRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000037", "name": "INSCRIPCIÓN 2º MUY CACHORRO - 3º PLAZO - LATIN WINNER"}, {"code": "21000038", "name": "INSCRIPCIÓN DESDE 3º MUY CACHORRO-3º PLAZO-LATIN WINNER"}, {"code": "21000039", "name": "INSCRIPCIÓN 1º VETERANO - 3º PLAZO - LATIN WINNER"}, {"code": "21000040", "name": "INSCRIPCIÓN 2º VETERANO - 3º PLAZO - LATIN WINNER"}, {"code": "21000041", "name": "INSCRIPCIÓN DESDE 3º VETERANO - 3º PLAZO - LATIN WINNER"}, {"code": "21000042", "name": "INSCRIPCIÓN PAREJAS - 2º PLAZO - LATIN WINNER"}, {"code": "21000043", "name": "BONIFICACION RAZAS ESPAÑOLAS - LATIN WINNER"}, {"code": "21000044", "name": "BONIFICACION INSCRIPCIÓN VETERANOS - LATIN WINNER"}, {"code": "21000045", "name": "BONIFICACIÓN MULTIINSCRIPCIÓN - LATIN WINNER"}, {"code": "22000001", "name": "INSCRIPCIÓN 1º PERRO - RSCE WINNER"}, {"code": "22000002", "name": "INSCRIPCIÓN 2º PERRO - RSCE WINNER"}, {"code": "22000003", "name": "INSCRIPCIÓN DESDE 3º PERRO - RSCE WINNER"}, {"code": "22000004", "name": "INSCRIPCIÓN 1º CACHORRO - RSCE WINNER"}, {"code": "22000005", "name": "INSCRIPCIÓN 2º CACHORRO - RSCE WINNER"}, {"code": "22000006", "name": "INSCRIPCIÓN DESDE 3º CACHORRO - RSCE WINNER"}, {"code": "22000007", "name": "INSCRIPCIÓN 1º MUY CACHORRO - RSCE WINNER"}, {"code": "22000008", "name": "INSCRIPCIÓN 2º MUY CACHORRO - RSCE WINNER"}, {"code": "22000009", "name": "INSCRIPCIÓN DESDE 3º MUY CACHORRO - RSCE WINNER"}, {"code": "22000010", "name": "INSCRIPCIÓN 1º VETERANO - RSCE WINNER"}, {"code": "22000011", "name": "INSCRIPCIÓN 2º VETERANO - RSCE WINNER"}, {"code": "22000012", "name": "INSCRIPCIÓN DESDE 3º VETERANO - RSCE WINNER"}, {"code": "22000013", "name": "INSCRIPCIÓN PAREJAS - RSCE WINNER"}, {"code": "22000014", "name": "INSCRIPCIÓN GRUPOS DE CRÍA - RSCE WINNER"}, {"code": "22000015", "name": "REGISTRO INICIAL RAZAS ESPAÑOLAS - RSCE WINNER"}, {"code": "22000016", "name": "REGISTRO INICIAL RAZAS INTEGRADAS - RSCE WINNER"}, {"code": "22000017", "name": "DISPONIBLE"}, {"code": "22000018", "name": "INSCRIPCIÓN 1º PERRO - 2º PLAZO - RSCE WINNER"}, {"code": "22000019", "name": "INSCRIPCIÓN 2º PERRO - 2º PLAZO - RSCE WINNER"}, {"code": "22000020", "name": "INSCRIPCIÓN DESDE 3º PERRO - 2º PLAZO -RSCE WINNER"}, {"code": "22000021", "name": "INSCRIPCIÓN 1º CACHORRO- 2º PLAZO-RSCE WINNER"}, {"code": "22000022", "name": "INSCRIPCIÓN 2º CACHORRO- 2º PLAZO-RSCE WINNER"}, {"code": "22000023", "name": "INSCRIPCIÓN DESDE 3º CACHORRO-2º PLAZO-RSCE WINNER"}, {"code": "22000024", "name": "INSCRIPCIÓN 1º MUY CACHORRO -2º PLAZO-RSCE WINNER"}, {"code": "22000025", "name": "INSCRIPCIÓN 2º MUY CACHORRO - 2º PLAZO-RSCE WINNER"}, {"code": "22000026", "name": "INSCRIPCIÓN DESDE 3º MUY CACHORRO-2º PLAZO-RSCE WINNER"}, {"code": "22000027", "name": "INSCRIPCIÓN 1º VETERANO - 2º PLAZO - RSCE WINNER"}, {"code": "22000028", "name": "INSCRIPCIÓN 2º VETERANO - 2º PLAZO-RSCE WINNER"}, {"code": "22000029", "name": "INSCRIPCIÓN DESDE 3º VETERANO-2º PLAZO-RSCE WINNER"}, {"code": "22000030", "name": "INSCRIPCIÓN 1º PERRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000031", "name": "INSCRIPCIÓN 2º PERRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000032", "name": "INSCRIPCIÓN DESDE 3º PERRO - 3º PLAZO -RSCE WINNER"}, {"code": "22000033", "name": "INSCRIPCIÓN 1º CACHORRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000034", "name": "INSCRIPCIÓN 2º CACHORRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000035", "name": "INSCRIPCIÓN DESDE 3º CACHORRO - 3º PLAZO-RSCE WINNER"}, {"code": "22000036", "name": "INSCRIPCIÓN 1º MUY CACHORRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000037", "name": "INSCRIPCIÓN 2º MUY CACHORRO - 3º PLAZO - RSCE WINNER"}, {"code": "22000038", "name": "INSCRIPCIÓN DESDE 3º MUY CACHORRO-3º PLAZO-RSCE WINNER"}, {"code": "22000039", "name": "INSCRIPCIÓN 1º VETERANO - 3º PLAZO - RSCE WINNER"}, {"code": "22000040", "name": "INSCRIPCIÓN 2º VETERANO - 3º PLAZO - RSCE WINNER"}, {"code": "22000041", "name": "INSCRIPCIÓN DESDE 3º VETERANO - 3º PLAZ0-RSCE WINNER"}, {"code": "22000042", "name": "INSCRIPCIÓN PAREJAS- 2º PLAZO - RSCE WINNER"}, {"code": "22000043", "name": "BONIFICACION RAZAS ESPAÑOLAS - RSCE WINNER"}, {"code": "22000044", "name": "BONIFICACION INSCRIPCIÓN VETERANOS - RSCE WINNER"}, {"code": "22000045", "name": "BONIFICACION MULTIINSCRIPCIÓN - RSCE WINNER"}, {"code": "22000099", "name": "INSCRIPCIONES EXPOSICION INTERNACIONAL PINTO"}, {"code": "23000000", "name": "INSCRIPCIÓN PERRO/S EN EXPOSICION NACIONAL CANINA"}, {"code": "23000001", "name": "INSCRIPCIÓN 1º PERRO"}, {"code": "23000002", "name": "INSCRIPCIÓN 2º PERRO"}, {"code": "23000003", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES"}, {"code": "23000004", "name": "INSCRIPCIÓN 1º CACHORRO"}, {"code": "23000005", "name": "INSCRIPCIÓN 2º CACHORRO"}, {"code": "23000006", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES"}, {"code": "23000007", "name": "INSCRIPCIÓN 1º MUY CACHORRO"}, {"code": "23000008", "name": "INSCRIPCIÓN 2º MUY CACHORRO"}, {"code": "23000009", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGUIENTES"}, {"code": "23000010", "name": "INSCRIPCIÓN 1º VETERANO"}, {"code": "23000011", "name": "INSCRIPCIÓN 2º VETERANO"}, {"code": "23000012", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES"}, {"code": "23000013", "name": "INSCRIPCIÓN PAREJAS"}, {"code": "23000014", "name": "INSCRIPCIÓN GRUPOS DE CRÍA"}, {"code": "23000015", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS ESPAÑOLAS"}, {"code": "23000016", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS INTEGRADAS EN E"}, {"code": "23000017", "name": "INSCRIPCIÓN RECONOCIMIENTO DE RAZA"}, {"code": "23000018", "name": "INSCRIPCIÓN 1º PERRO - SEGUNDO PLAZO"}, {"code": "23000019", "name": "INSCRIPCIÓN 2º PERRO - SEGUNDO PLAZO"}, {"code": "23000020", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES - SEGUNDO PLAZO"}, {"code": "23000021", "name": "INSCRIPCIÓN 1º CACHORRO - SEGUNDO PLAZO"}, {"code": "23000022", "name": "INSCRIPCIÓN 2º CACHORRO - SEGUNDO PLAZO"}, {"code": "23000023", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES - SEGUNDO PLA"}, {"code": "23000024", "name": "INSCRIPCIÓN 1º MUY CACHORRO - SEGUNDO PLAZO"}, {"code": "23000025", "name": "INSCRIPCIÓN 2º MUY CACHORRO - SEGUNDO PLAZO"}, {"code": "23000026", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGUIENTES - SEGUNDO"}, {"code": "23000027", "name": "INSCRIPCIÓN 1º VETERANO - SEGUNDO PLAZO"}, {"code": "23000028", "name": "INSCRIPCIÓN 2º VETERANO - SEGUNDO PLAZO"}, {"code": "23000029", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES - SEGUNDO PLA"}, {"code": "23000030", "name": "INSCRIPCIÓN 1º PERRO - TERCER PLAZO"}, {"code": "23000031", "name": "INSCRIPCIÓN 2º PERRO - TERCER PLAZO"}, {"code": "23000032", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES - TERCER PLAZO"}, {"code": "23000033", "name": "INSCRIPCIÓN 1º CACHORRO - TERCER PLAZO"}, {"code": "23000034", "name": "INSCRIPCIÓN 2º CACHORRO - TERCER PLAZO"}, {"code": "23000035", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES - TERCER PLAZ"}, {"code": "23000036", "name": "INSCRIPCIÓN 1º MUY CACHORRO - TERCER PLAZO"}, {"code": "23000037", "name": "INSCRIPCIÓN 2º MUY CACHORRO - TERCER PLAZO"}, {"code": "23000038", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGUIENTES - TERCER"}, {"code": "23000039", "name": "INSCRIPCIÓN 1º VETERANO - TERCER PLAZO"}, {"code": "23000040", "name": "INSCRIPCIÓN 2º VETERANO - TERCER PLAZO"}, {"code": "23000041", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES - TERCER PLAZ"}, {"code": "23500000", "name": "INSCRIPCIÓN PERRO/S EN RAZAS ESPAÑOLAS"}, {"code": "24000000", "name": "INSCRIPCIÓN PERRO/S EN CONCURSO CANINO"}, {"code": "24000001", "name": "INSCRIPCIÓN 1º PERRO"}, {"code": "24000002", "name": "INSCRIPCIÓN 2º PERRO"}, {"code": "24000003", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES"}, {"code": "24000004", "name": "INSCRIPCIÓN 1º CACHORRO"}, {"code": "24000005", "name": "INSCRIPCIÓN 2º CACHORRO"}, {"code": "24000006", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES"}, {"code": "24000007", "name": "INSCRIPCIÓN 1º MUY CACHORRO"}, {"code": "24000008", "name": "INSCRIPCIÓN 2º MUY CACHORRO"}, {"code": "24000009", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGUIENTES"}, {"code": "24000010", "name": "INSCRIPCIÓN 1º VETERANO"}, {"code": "24000011", "name": "INSCRIPCIÓN 2º VETERANO"}, {"code": "24000012", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES"}, {"code": "24000013", "name": "INSCRIPCIÓN PAREJAS"}, {"code": "24000014", "name": "INSCRIPCIÓN GRUPOS DE CRÍA"}, {"code": "24000015", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS ESPAÑOLAS"}, {"code": "24000016", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS INTEGRADAS EN E"}, {"code": "24000017", "name": "INSCRIPCIÓN CONFIRMACIÓN DE RAZA (RAZAS ESPAÑOLAS)"}, {"code": "25000001", "name": "INSCRIPCIÓN 1º PERRO"}, {"code": "25000002", "name": "INSCRIPCIÓN 2º PERRO"}, {"code": "25000003", "name": "INSCRIPCIÓN 3º PERRO Y SIGUIENTES"}, {"code": "25000004", "name": "INSCRIPCIÓN 1º CACHORRO"}, {"code": "25000005", "name": "INSCRIPCIÓN 2º CACHORRO"}, {"code": "25000006", "name": "INSCRIPCIÓN 3º CACHORRO Y SIGUIENTES"}, {"code": "25000007", "name": "INSCRIPCIÓN 1º MUY CACHORRO"}, {"code": "25000008", "name": "INSCRIPCIÓN 2º MUY CACHORRO"}, {"code": "25000009", "name": "INSCRIPCIÓN 3º MUY CACHORRO Y SIGUIENTES"}, {"code": "25000010", "name": "INSCRIPCIÓN 1º VETERANO"}, {"code": "25000011", "name": "INSCRIPCIÓN 2º VETERANO"}, {"code": "25000012", "name": "INSCRIPCIÓN 3º VETERANO Y SIGUIENTES"}, {"code": "25000013", "name": "INSCRIPCIÓN PAREJAS"}, {"code": "25000014", "name": "INSCRIPCIÓN GRUPOS DE CRÍA"}, {"code": "25000015", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS ESPAÑOLAS"}, {"code": "25000016", "name": "INSCRIPCIÓN REGISTRO INICIAL RAZAS INTEGRADAS EN E"}, {"code": "25000017", "name": "INSCRIPCIÓN RECONOCIMIENTO DE RAZA"}, {"code": "30000002", "name": "1º PERRO - INSCRIPCIÓN PRUEBA APTITUDES NATURALES"}, {"code": "30000003", "name": "2º PERRO Y SIGUIENTES - PRUEBA APTITUDES NATURALES"}, {"code": "31000001", "name": "CARTILLA PRUEBAS PERROS DE MUESTRA"}, {"code": "31000004", "name": "ACTAS CACT-CACIT - PERROS DE MUESTRA"}, {"code": "31000005", "name": "IMPRESO DIPLOMA DE INICIACIÓN - PERRO DE MUESTRA"}, {"code": "31000006", "name": "CERTIFICADO DE TRABAJO PERROS DE MUESTRA"}, {"code": "31000007", "name": "IMPRESO APTITUDES NATURALES - PERROS DE MUESTRA"}, {"code": "31000008", "name": "ACTAS BARRAGE - PERROS DE MUESTRA"}, {"code": "31000009", "name": "TÍTULO DE CAMPEÓN - PERROS DE MUESTRA"}, {"code": "31000010", "name": "TÍTULO TRIALER - PERROS DE MUESTRA"}, {"code": "31000011", "name": "TÍTULO FCI - PERROS DE MUESTRA CONTINENTALES"}, {"code": "31000012", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES-PERROS DE MUESTRA"}, {"code": "31000013", "name": "INSCRIPCION COPA ESPAÑA CAZA PRACTICA"}, {"code": "31000014", "name": "CANON POR PERRO INSCRITO DISC. BASICAS-PERROS DE MUESTRA"}, {"code": "31000015", "name": "TASA POR DÍA DE PRUEBA DISC. BÁSICAS PERROS DE MUESTRA"}, {"code": "31000016", "name": "TASA POR DÍA CACT PERROS DE MUESTRA"}, {"code": "31000017", "name": "TASA POR DÍA CACIT PERROS DE MUESTRA"}, {"code": "31000018", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - PERROS DE MUESTRA"}, {"code": "31000019", "name": "ANULACION DE PRUEBAS CACT/DIA - PERROS DE MUESTRA"}, {"code": "31000020", "name": "ANULACION DE PRUEBAS CACT-CACIT/DIA - PERROS DE MUESTRA"}, {"code": "31000021", "name": "TÍTULO FCI - PERROS DE MUESTRA BRITANICOS"}, {"code": "31000022", "name": "INSCRIPCION COPA EXCELENCIA"}, {"code": "31000023", "name": "INSCRIPCION COPA EUROPA GRAN BUSQUEDA"}, {"code": "31000024", "name": "INSCRIPCION CAZA PRACTICA NATURAL EN PAREJA - BRITANICOS"}, {"code": "32000001", "name": "CERTIFICADO DE TRABAJO - SPANIEL"}, {"code": "32000002", "name": "CARTILLA PRUEBAS SPANIELS"}, {"code": "32000003", "name": "IMPRESO APTITUDES NATURALES - SPANIELS"}, {"code": "32000004", "name": "ACTAS CACT-CACIT - SPANIELS"}, {"code": "32000005", "name": "TÍTULO DE CAMPEÓN - SPANIELS"}, {"code": "32000006", "name": "TÍTULO DE TRIALES - SPANIELS"}, {"code": "32000007", "name": "TÍTULOS FCI - SPANIELS"}, {"code": "32000008", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES SPANIELS"}, {"code": "32000009", "name": "CANON POR PERRO INSCRITO DISC. BASICAS SPANIELS"}, {"code": "32000010", "name": "TASA POR DIA DE PRUEBA DISC. BÁSICAS SPANIELS"}, {"code": "32000011", "name": "TASA POR DIA DE PRUEBA DISC. SUPERIORES SPANIELS"}, {"code": "32000012", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - SPANIELS"}, {"code": "32000013", "name": "ANULACION DE PRUEBAS CACT/DIA - SPANIELS"}, {"code": "32000014", "name": "TASA POR DÍA CACT SPANIELS"}, {"code": "32000015", "name": "TASA POR DÍA CACIT SPANIELS"}, {"code": "33000001", "name": "CARTILLAS DE PRUEBAS - RETRIEVERS"}, {"code": "33000002", "name": "CERTIFICADO DE TRABAJO - RETRIEVERS"}, {"code": "33000003", "name": "IMPRESO APTITUDES NATURALES - RETRIEVERS"}, {"code": "33000004", "name": "ACTAS CACT-CACIT - RETRIEVERS"}, {"code": "33000005", "name": "TÍTULO DE CAMPEÓN - RETRIEVERS"}, {"code": "33000006", "name": "TÍTULO TRIALER - RETRIEVERS"}, {"code": "33000007", "name": "TÍTULOS FCI- RETRIEVERS"}, {"code": "33000008", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES RETRIEVERS"}, {"code": "33000009", "name": "CANON POR PERRO INSCRITO DISC. BASICAS RETRIEVERS"}, {"code": "33000010", "name": "TASA POR DIA DE PRUEBA DISC. BÁSICAS RETRIEVERS"}, {"code": "33000011", "name": "TASA POR DÍA CACT RETRIEVERS"}, {"code": "33000012", "name": "TASA POR DÍA CACIT RETRIEVERS"}, {"code": "33000014", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - RETRIEVERS"}, {"code": "33000015", "name": "ANULACION DE PRUEBAS CACT/DIA - RETRIEVERS"}, {"code": "33000016", "name": "ANULACION DE PRUEBAS CACT-CACIT/DIA - RETRIEVERS"}, {"code": "33000017", "name": "TASA POR DÍA PRUEBA DISC. SUPERIORES RETRIEVERS"}, {"code": "34000001", "name": "IMPRESO APTITUDES NATURALES - TECKEL Y TERRIERS"}, {"code": "34000002", "name": "CARTILLA PARA PRUEBAS - TECKELS Y TERRIERS"}, {"code": "34000003", "name": "CERTIFICADO DE TRABAJO - TECKELS Y TERRIERS"}, {"code": "34000004", "name": "ACTAS CACT-CACIT - TECKELS Y TERRIERS"}, {"code": "34000005", "name": "TÍTULO DE CAMPEÓN - TECKELS Y TERRIERS"}, {"code": "34000006", "name": "TÍTULO TRIALER - TECKELS Y TERRIERS"}, {"code": "34000007", "name": "TÍTULOS FCI - TECKELS Y TERRIERS"}, {"code": "34000008", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES TECKELS Y TERRIERS"}, {"code": "34000009", "name": "CANON POR PERRO INSCRITO DISC. BASICAS TECKELS Y TERRIERS"}, {"code": "34000010", "name": "TASA POR DIA DE PRUEBA DISC. BÁSICAS TECKELS Y TERRIERS"}, {"code": "34000011", "name": "TASA POR DIA DE PRUEBA DISC. SUPERIORES TECKELS Y TERRIERS"}, {"code": "34000012", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - TECKELS Y TERRIER"}, {"code": "34000013", "name": "ANULACION DE PRUEBAS CACT/DIA - TECKELS Y TERRIER"}, {"code": "34000014", "name": "TASA POR DÍA CACT TECKELS Y TERRIERS"}, {"code": "34000015", "name": "TASA POR DÍA CACIT TECKELS Y TERRIERS"}, {"code": "34010001", "name": "INSCRIPCIÓN COPA ESPAÑA RASTRO SANGRE ARTIFICIAL 20 HORAS"}, {"code": "34010002", "name": "INSCRIPCIÓN COPA ESPAÑA RASTRO SANGRE ARTIFICIAL 40 HORAS"}, {"code": "34010003", "name": "INSCRIPCIÓN PRUEBA APTITUDES NATURALES - RASTROS"}, {"code": "34010004", "name": "INSCRIPCIÓN COBRO EN AGUA"}, {"code": "35000001", "name": "CARTILLAS PARA PRUEBAS - PODENCOS"}, {"code": "35000002", "name": "CERTIFICADO DE TRABAJO - PODENCOS"}, {"code": "35000003", "name": "IMPRESO APTITUDES NATURALES - PODENCOS"}, {"code": "35000004", "name": "ACTAS CACT-CACIT"}, {"code": "35000005", "name": "TÍTULOS FCI - PODENCOS"}, {"code": "35000006", "name": "TÍTULO TRIALER - PODENCOS"}, {"code": "35000007", "name": "CANON POR PERRO INSCRITO DISC. BÁSICAS PODENCOS"}, {"code": "35000008", "name": "INSCRIPCION COPA ESPAÑA PODENCOS"}, {"code": "35000009", "name": "TASA POR DÍA DE PRUEBA DISC. BASICAS PODENCOS"}, {"code": "35000010", "name": "TASA POR DIA CACT PODENCOS"}, {"code": "35000011", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - PODENCOS"}, {"code": "35000012", "name": "ANULACION DE PRUEBAS CACT/DIA - PODENCOS"}, {"code": "35000013", "name": "ANULACION DE PRUEBAS CACT-CACIT/DIA - PODENCOS"}, {"code": "35000014", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES PODENCOS"}, {"code": "36000001", "name": "1ª CARTILLA JAURÍA PERROS RASTRO GRUPO 6º"}, {"code": "36000002", "name": "2º CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000003", "name": "3ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000004", "name": "4ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000005", "name": "CARTILLAS PARA PRUEBAS - RASTRO"}, {"code": "36000006", "name": "CERTIFICADO DE TRABAJO - RASTRO"}, {"code": "36000007", "name": "IMPRESO APTITUDES NATURALES - RASTRO"}, {"code": "36000008", "name": "ACTAS CACT-CACIT - RASTRO"}, {"code": "36000009", "name": "TITULO DE CAMPEÓN - RASTRO"}, {"code": "36000010", "name": "TÍTULO TRIALER - RASTRO"}, {"code": "36000011", "name": "TÍTULOS FCI - RASTRO"}, {"code": "36000012", "name": "CANON POR PERRO INSCRITO DISC. SUPERIORES PERROS DE RASTRO"}, {"code": "36000013", "name": "TASA POR DIA DE PRUEBA DISC. BÁSICAS PERROS DE RASTRO"}, {"code": "36000014", "name": "TASA POR DIA CACT PERROS DE RASTRO"}, {"code": "36000015", "name": "ANULACION DE PRUEBAS DISC. BASICAS/DIA - RASTRO"}, {"code": "36000016", "name": "ANULACION DE PRUEBAS CACT/DIA - RASTRO"}, {"code": "36000017", "name": "5ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000018", "name": "6ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000019", "name": "7ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000020", "name": "8ª CARTILLA JAURÍA PERROS DE RASTRO GRUPO 6º"}, {"code": "36000021", "name": "TASA POR DIA DE PRUEBA DISC. SUPERIORES PERROS DE RASTRO"}, {"code": "36000022", "name": "CANON POR PERRO INSCRITO DISC. BÁSICAS PERRO DE RASTRO"}, {"code": "36000023", "name": "TASA POR DÍA CACIT PERROS DE RASTRO"}, {"code": "36000024", "name": "CARNET DE TRABAJO PARA PERROS DE RASTRO DEL GRUPO 6º"}, {"code": "36010001", "name": "INSCRIPCIÓN COPA EN ESPAÑA DE RASTRO SOBRE LIEBRE - CAT. A"}, {"code": "36010002", "name": "INSCRIPCION COPA EN ESPAÑA DE RASTRO DE CONEJO Y COBRO MADR."}, {"code": "40000001", "name": "CUOTA ANUAL MULTIDISCIPLINAR"}, {"code": "40000002", "name": "TARJETA PRUEBA SOCIABILIDAD"}, {"code": "40210002", "name": "INSCRIPCIÓN CAMPEONATO DEL MUNDO DE IGP"}, {"code": "41000001", "name": "LICENCIA INICIAL AGILITY"}, {"code": "41000002", "name": "RENOVACION ANUAL LICENCIA AGILITY"}, {"code": "41000003", "name": "CUOTA ANUAL CLUB AGILITY"}, {"code": "41000004", "name": "CARTILLA DE AGILITY"}, {"code": "41000005", "name": "3º MEDICION EJEMPLARES PARA DETERMINAR CATEGORIA DE PARTICIP"}, {"code": "41000006", "name": "CANON POR PERRO INSCRITO - AGILITY"}, {"code": "41000007", "name": "CANON PERRO INSCRITO CAMPEONATO ESPAÑA - AGILITY"}, {"code": "41000008", "name": "CANON PERRO INSCRITO PRUEBAS REGIONALES - AGILITY"}, {"code": "41000009", "name": "MEDIDOR AGILITY"}, {"code": "41000010", "name": "CANON POR PERRO INSCRITO EN PRUEBA SELECTIVA AGILITY"}, {"code": "41000011", "name": "CANON POR PERRO INSCRITO EN COPAS AGILITY PARA RAZA"}, {"code": "41000012", "name": "CANON POR PERRO INSCRITO CAMPEONATO ESPAÑA EQUIPOS - AGILITY"}, {"code": "41000013", "name": "CANON POR PRUEBA ANULADA"}, {"code": "41000014", "name": "CURSO JUECES AGILITY"}, {"code": "42000001", "name": "CANON PERRO INSCRITO CACT - GRADO 3 - IGP"}, {"code": "42000002", "name": "CANON PETICIÓN CACT - IGP"}, {"code": "42000003", "name": "LICENCIA INICIAL IGP"}, {"code": "42000004", "name": "RENOVACIÓN ANUAL LICENCIA IGP"}, {"code": "42000005", "name": "CUOTA ANUAL GRUPO DE TRABAJO IGP"}, {"code": "42000006", "name": "CARTILLA DE TRABAJO - IGP"}, {"code": "42000007", "name": "CERTIFICADO DE TRABAJO IGP"}, {"code": "42010001", "name": "INSCRIPCIÓN COPA EN ESPAÑA DE IGP"}, {"code": "42020001", "name": "INSCRIPCIÓN COPA EN ESPAÑA IBGH/StPr"}, {"code": "43010001", "name": "INSCRIPCIÓN COPA EN ESPAÑA IGP-IFH"}, {"code": "44000001", "name": "CANON PERRO INSCRITO CACOB - CACIOB - GRADO 3 - OBEDIENCIA"}, {"code": "44000002", "name": "CANON PETICIÓN CACT - OBEDIENCIA"}, {"code": "44000003", "name": "LICENCIA INICIAL OBEDIENCIA"}, {"code": "44000004", "name": "RENOVACIÓN ANUAL LICENCIA OBEDIENCIA"}, {"code": "44000005", "name": "CUOTA ANUAL GRUPO DE TRABAJO OBEDICIENCIA"}, {"code": "44000006", "name": "CARTILLA DE TRABAJO OBEDIENCIA"}, {"code": "44000007", "name": "INSCRIPCIÓN COPA EN ESPAÑA DE OBEDIENCIA"}, {"code": "44000008", "name": "INSCRIPCIÓN COPA EN ESPAÑA OBEDIENCIA - RAZAS ESPAÑOLAS"}, {"code": "44000009", "name": "FORMACION JUECES RALLY-O"}, {"code": "44000010", "name": "CURSO REGLAMENTO RALLY-O"}, {"code": "44000011", "name": "INSCRIPCIÓN COPA EN ESPAÑA DE RALLY-O - 1º PLAZO"}, {"code": "44000012", "name": "INSCRIPCIÓN COPA EN ESPAÑA DE RALLY-O - 2º PLAZO"}, {"code": "44030001", "name": "INSCRIPCIÓN SELECTIVA DOBLE OBEDIENCIA"}, {"code": "45000001", "name": "CANON PERRO INSCRITO CACT - GRADO 3 - MONDIORING"}, {"code": "45000002", "name": "CANON PETICIÓN CACT MONDIORING"}, {"code": "45000003", "name": "LICENCIA INICIAL MONDIORING"}, {"code": "45000004", "name": "RENOVACIÓN ANUAL LICENCIA MONDIORING"}, {"code": "45000005", "name": "CUOTA ANUAL GRUPO DE TRABAJO MONDIORING"}, {"code": "45000006", "name": "CARTILLA DE TRABAJO MONDIORING"}, {"code": "45000007", "name": "CERTIFICADO DE TRABAJO MONDIORING"}, {"code": "45000008", "name": "CARTILLA ASISTENTE MONDIORING"}, {"code": "45000009", "name": "LICENCIA INICIAL PARA HA MONDIORING"}, {"code": "45000010", "name": "RENOVACION ANUAL LICENCIA PARA HA MONDIORING"}, {"code": "45000011", "name": "EXAMEN HA MONDIORING"}, {"code": "45010001", "name": "INSCRIPCION COPA EN ESPAÑA DE MONDIORING"}, {"code": "45010002", "name": "INSCRIPCION COPA EN ESPAÑA DE OBEDIENCIA DE MONDIORING"}, {"code": "46000001", "name": "CARTILLA PRUEBAS POLIVALENTES - PERRO DE AGUA"}, {"code": "47000001", "name": "LICENCIA INICIAL PRUEBAS BUSQUEDA Y RESCATE"}, {"code": "47000002", "name": "RENOVACIÓN ANUAL LICENCIA PRUEBAS RESCATE"}, {"code": "47000003", "name": "CUOTA ANUAL GRUPO DE RESCATE"}, {"code": "47000004", "name": "CARTILLA PRUEBAS DE RESCATE"}, {"code": "47000005", "name": "LICENCIA INICIAL PRUEBAS RESCATE ACUATICO DEPORTIVO"}, {"code": "47000006", "name": "RENOVACION ANUAL LICENCIA PRUEBAS RESCATE ACUATICO DEPORTIVO"}, {"code": "47000007", "name": "CANON PETICIÓN CACT - RESCATE ACUÁTICO"}, {"code": "50000001", "name": "ANOTACIÓN BASE DE DATOS PRUEBAS - CONVENIO RSCE"}, {"code": "50000002", "name": "ANOTACIÓN EN BASE DE DATOS PRUEBA ADN"}, {"code": "50000003", "name": "IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE"}, {"code": "50000005", "name": "ANOTACIÓN PRUEBA DISPLASIA"}, {"code": "50000006", "name": "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CODO"}, {"code": "50000007", "name": "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CADERA"}, {"code": "50000008", "name": "RECURSO AL DIAGNOSTICO DE DISPLASIA"}, {"code": "50000009", "name": "CERTIFICADO INTERNACIONAL DE DISPLASIA"}, {"code": "50000010", "name": "BONO 25 ADN"}, {"code": "50000011", "name": "BONO 50 ADN"}, {"code": "50000012", "name": "BONO 100 ADN"}, {"code": "50000013", "name": "TASA ADMINISTRATIVA - ANOTACIÓN ADN EN LOE"}, {"code": "50000014", "name": "PRE-REGISTRO PRUEBA DE ADN"}, {"code": "50000015", "name": "KIT MY.DOG.DNA"}, {"code": "50000016", "name": "RESULTADO VIALES DE LABORATORIO"}, {"code": "50000017", "name": "KIT MY.DOG.DNA - BLACK FRIDAY"}, {"code": "54000001", "name": "CARTILLA PARA PRUEBAS POLIVALENTES - PERRO DE AGUA"}, {"code": "55500000", "name": "FCI SUPLIDOS"}, {"code": "55500025", "name": "ENTIDADES COLABORADORAS"}, {"code": "55500125", "name": "CLUBES DE RAZA"}, {"code": "5550025", "name": "SUPLIDOS FCI"}, {"code": "6001002", "name": "MATERIAL OFICINA - TARJETAS SOCIO/IDENTIFICADORES"}, {"code": "6001004", "name": "MATERIAL OFICINA - TONER"}, {"code": "6001006", "name": "MATERIAL OFICINA - OTROS"}, {"code": "6001102", "name": "IMPRENTA - ENCUADERNACION"}, {"code": "6001103", "name": "IMPRENTA - JUSTIFICANTES DE INSCRIPCION"}, {"code": "6001104", "name": "IMPRENTA - PEDIGREES 3 GENERACIONES"}, {"code": "6001105", "name": "IMPRENTA - PEDIGREES 4 GENERACIONES"}, {"code": "6001108", "name": "IMPRENTA - IMPRESOS DE CAMADAS"}, {"code": "6001111", "name": "IMPRENTA - OTRAS IMPRESIONES (AFIJOS, SOBRES, TARJETAS...)"}, {"code": "6001112", "name": "IMPRENTA - CARTILLAS"}, {"code": "6002000", "name": "COMPRA DE MEDALLAS Y TROFEOS"}, {"code": "6002001", "name": "COMPRA OBSEQUIOS"}, {"code": "6002002", "name": "ESCARAPELAS CAC/CACIB Y CONCURSOS"}, {"code": "6002003", "name": "COMPRA OBSEQUIOS - IVA 10%"}, {"code": "6002004", "name": "OBSEQUIOS EXENTOS IVA"}, {"code": "6002005", "name": "COMPRA OBSEQUIOS - IVA 2%"}, {"code": "6009001", "name": "EQUIPACIONES"}, {"code": "6009002", "name": "ROLL-UP / CARTELES"}, {"code": "6009006", "name": "COMPRA PEQUEÑO MATERIAL"}, {"code": "6009007", "name": "COMPRAS AL 0%"}, {"code": "6009008", "name": "MERCHANDISING"}, {"code": "6070000", "name": "TROE - OTRAS GESTIONES"}, {"code": "6070001", "name": "TROE - GESTION ALARMA INTRUSION"}, {"code": "6070002", "name": "TROE - GESTION VETERINARIA EXPOSICIONES"}, {"code": "6070009", "name": "TROE - GESTION MUESTRAS ADN"}, {"code": "6070010", "name": "TROE - ANCLAJES EXPOSICIONES"}, {"code": "6070014", "name": "TROE - GESTION SISTEMA CONTRA INCENDIOS"}, {"code": "6070015", "name": "TROE - GESTION RIESGOS LABORALES"}, {"code": "6070016", "name": "TROE - GESTION RIESGOS LABORALES (IVA 0%)"}, {"code": "6070017", "name": "TROE - GESTION PROTECCION DATOS"}, {"code": "6070020", "name": "TROE - ACTIVIDADES EXPOSICIONES"}, {"code": "6070023", "name": "TROE - CONSULTORIAS TÉCNICAS"}, {"code": "6210002", "name": "ALQUILER SALAS"}, {"code": "6210003", "name": "ALQUILER CAMPOS"}, {"code": "6210004", "name": "ALQUILER IMPRESORAS/FOTOCOPIADORAS"}, {"code": "6212000", "name": "ALQUILER MAQUINARIA Y VEHICULO"}, {"code": "6213000", "name": "ALQUILER MATERIAL"}, {"code": "62130001", "name": "ALQUILER MOBILIARIO EVENTO"}, {"code": "6213002", "name": "ALQUILER GRADAS"}, {"code": "6213003", "name": "ALQUILER EQUIPO SONIDO"}, {"code": "6213004", "name": "ALQUILER PANTALLA GIGANTE"}, {"code": "6213005", "name": "ALQUILER CORTINAS RING HONOR"}, {"code": "6220000", "name": "REPARACIONES Y CONSERVACION"}, {"code": "6221001", "name": "MANTENIMIENTO SISTEMA ACCESO"}, {"code": "6221006", "name": "MANTENIMIENTO PLATAFORMAS WEB"}, {"code": "6221007", "name": "MANTENIMIENTO INSTALACIONES MALDONADO"}, {"code": "6221011", "name": "MANTENIMIENTO PROGRAMA CONTABILIDAD"}, {"code": "6221013", "name": "MANTENIMIENTO PROGRAMA NOMINAS"}, {"code": "6221017", "name": "MANTENIMIENTO FOTOCOPIADORAS"}, {"code": "6223000", "name": "LIMPIEZA"}, {"code": "6230001", "name": "PROFESIONALES INDEPENDIENTES - ASESORIA JURIDICA"}, {"code": "6230002", "name": "PROFESIONALES INDEPENDIENTES - AUDITORES"}, {"code": "6230003", "name": "PROFESIONALES INDEPENDIENTES - NOTARIOS Y REGISTRADORES"}, {"code": "6230004", "name": "PROFESIONALES INDEPENDIENTES - NOTARIOS - EXENTO IVA"}, {"code": "6240001", "name": "DESPLAZAMIENTOS URBANOS PERSONAL RSCE"}, {"code": "6241000", "name": "DESPLAZAMIENTOS Y GASTOS VIAJES PERSONAL RSCE"}, {"code": "6260000", "name": "COMISIONES BANCARIAS"}, {"code": "6271000", "name": "RELACIONES PUBLICAS (EXENTO IVA)"}, {"code": "6271001", "name": "RELACIONES PUBLICAS (IVA 10%)"}, {"code": "6271002", "name": "RELACIONES PUBLICAS (IVA 21%)"}, {"code": "6271003", "name": "RELACIONES PUBLICAS (IVA 4%)"}, {"code": "6280000", "name": "AGUA"}, {"code": "6281000", "name": "TELEFONO"}, {"code": "6281001", "name": "TELEFONIA MOVIL"}, {"code": "6282000", "name": "ELECTRICIDAD"}, {"code": "6282001", "name": "ELECTRICIDAD 10%"}, {"code": "6293000", "name": "CORREO POSTAL"}, {"code": "6293001", "name": "SERVICIO MENSAJERIA"}, {"code": "6293002", "name": "CORREO POSTAL - EXENTO"}, {"code": "6293400", "name": "FOTOCOPIAS"}, {"code": "6295000", "name": "CUOTAS FCI"}, {"code": "6298001", "name": "CURSO CAPACITACION DEL PERSONAL"}, {"code": "6298002", "name": "CURSO CAPACITACION DEL PERSONAL - EXENTO IVA"}, {"code": "6310000", "name": "TRIBUTOS Y OTROS ARBITRIOS"}, {"code": "6490001", "name": "FORMACION DE PERSONAL"}, {"code": "6530101", "name": "MEDIOS DE TRANSPORTE DELEGADOS - VOCALES"}, {"code": "6530102", "name": "MEDIOS TRANSPORTE JUECES"}, {"code": "6530103", "name": "MEDIOS TRANSPORTE COMISARIO GENERAL"}, {"code": "6530104", "name": "MEDIOS TRANSPORTE COMISARIOS PRINCIPALES"}, {"code": "6530105", "name": "MEDIOS TRANSPORTE COMISARIOS"}, {"code": "6530106", "name": "MEDIOS TRANSPORTE TRAZADORES"}, {"code": "6530107", "name": "MEDIOS TRANSPORTE FIGURANTES-ASISTENTES"}, {"code": "6530108", "name": "MEDIOS TRANSPORTE OTROS COLABORADORES"}, {"code": "6530109", "name": "MEDIOS TRANSPORTE OTROS COLABORADORES (CON IVA)"}, {"code": "6530302", "name": "MANUTENCION JUECES"}, {"code": "6530303", "name": "MANUTENCION COMISARIO GENERAL"}, {"code": "6530304", "name": "MANUTENCION COMISARIOS PRINCIPALES"}, {"code": "6530305", "name": "MANUTENCION COMISARIOS"}, {"code": "6530306", "name": "MANUTENCION TRAZADORES"}, {"code": "6530307", "name": "MANUTENCION FIGURANTES-ASISTENTES"}, {"code": "6530308", "name": "MANUTENCION - OTROS COLABORADORES"}, {"code": "6530402", "name": "HOSPEDAJE JUECES"}, {"code": "6530403", "name": "HOSPEDAJE COMISARIO GENERAL"}, {"code": "6530404", "name": "HOSPEDAJE COMISARIOS PRINCIPALES"}, {"code": "6530405", "name": "HOSPEDAJE COMISARIOS"}, {"code": "6530406", "name": "HOSPEDAJE TRAZADORES"}, {"code": "6530407", "name": "HOSPEDAJE FIGURANTES-ASISTENTES"}, {"code": "6530408", "name": "HOSPEDAJE OTROS COLABORADORES"}, {"code": "6540001", "name": "MEDIOS TRANSPORTE ORGANO DE GOBIERNO"}, {"code": "6540002", "name": "MEDIOS TRANSPORTE ORGANO DE GOBIERNO (IVA 10%)"}, {"code": "6540003", "name": "MANUTENCION ORGANO DE GOBIERNO"}, {"code": "6540004", "name": "HOSPEDAJE ORGANO DE GOBIERNO"}, {"code": "6540005", "name": "MANUTENCION ORGANO DE GOBIERNO (IVA 21%)"}, {"code": "6540014", "name": "RESIDENCIA MASCOTAS ORGANO DE GOBIERNO"}, {"code": "6541001", "name": "MEDIOS TRANSPORTE ORGANO DE GOBIERNO (EXENTO IVA)"}, {"code": "70000001", "name": "MOCHILA"}, {"code": "70000002", "name": "BOLSA"}, {"code": "70000003", "name": "LLAVERO"}, {"code": "70000004", "name": "GORRA RSCE"}, {"code": "70000005", "name": "MANTA AZUL"}, {"code": "70000006", "name": "CAMISETA RSCE"}, {"code": "70000007", "name": "CHALECO"}, {"code": "70000008", "name": "GORRA RSCE - BANDERA DE ESPAÑA"}, {"code": "70000009", "name": "CAMISERA RSCE JOVEN - INFANTIL"}, {"code": "70000010", "name": "CAMISETA RSCE JOVEN - JUVENIL"}, {"code": "70000011", "name": "LANYARD"}, {"code": "70000012", "name": "IMAN RSCE"}, {"code": "70000013", "name": "PACK JUEGOS (CARTAS + MEMORY)"}, {"code": "70000014", "name": "JUEGO DE CARTAS"}, {"code": "70000015", "name": "JUEGO MEMORY"}, {"code": "80000001", "name": "PATROCINIO PAGINA WEB"}, {"code": "80000002", "name": "PATROCINIO CRIADOR (1 raza)"}, {"code": "80000003", "name": "PATROCINIO CRIADORES (2 a 5 razas)"}, {"code": "80000004", "name": "PATROCINIO CRIADOR (6 a 9 razas)"}, {"code": "80000005", "name": "PATROCINIO CRIADOR (10 razas o más)"}, {"code": "80000006", "name": "PATROCINIO periodo de ejecución de 06/02/2025 al 08/02/2025"}, {"code": "80000007", "name": "PATROCINIO - COLABORADOR PREFERENTE"}, {"code": "80000008", "name": "PATROCINIO MyDogDNA"}, {"code": "80000009", "name": "PATROCINIO CAPTACIÓN Y FIDELIZACIÓN CRIADORES"}, {"code": "81000001", "name": "PATROCINIO CONCURSOS CANINOS"}, {"code": "81000002", "name": "PATROCINIO EXPOSICION LATIN WINNER"}, {"code": "81000003", "name": "PATROCINIO EXPOSICION RSCE WINNER"}, {"code": "81000004", "name": "PATROCINIO EXPOSICION NACIONAL CAC"}, {"code": "81000005", "name": "PATROCINIO EVENTO CANINO"}, {"code": "81000006", "name": "PATROCINIO EVENTO PERROS DE MUESTRA"}, {"code": "84104001", "name": "PATROCINIO EQUIPACIONES PARA EL EUROPEAN OPEN - AGILITY"}, {"code": "84203001", "name": "PATROCINIO EQUIPACIONES PARA EL CAMPEONATO DEL MUNDO IGP-FCI"}, {"code": "85000001", "name": "PATROCINIO COPA EN ESPAÑA IGP"}, {"code": "85000002", "name": "PATROCINIO CAMPEONATO DEL MUNDO FCI-IGP"}, {"code": "90000001", "name": "CUOTA INICIAL SOCIO"}, {"code": "90000002", "name": "CUOTA ANUAL SOCIO"}, {"code": "91000001", "name": "CUOTA INICIAL ABONADO"}, {"code": "91000002", "name": "CUOTA ANUAL ABONADO"}, {"code": "91000003", "name": "CUOTA ABONADO JOVEN"}, {"code": "99000001", "name": "ALQUILER OFICINA LAGASCA, 16"}, {"code": "99000002", "name": "GASTOS GENERALES Y SERVICIOS"}, {"code": "99000003", "name": "COMPENSACIÓN"}, {"code": "99000004", "name": "RECURSO A LA COMISIÓN LIBRO DE ORÍGENES ESPAÑOL"}, {"code": "99000005", "name": "INSCRIPCION RSCE WINNER - LATIN WINNER"}, {"code": "99000006", "name": "SUPLIDO"}];

const RSCE_DATA = {"products": ["1ª CARTILLA", "1ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º", "2ª CARTILLA", "2ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º", "3ª CARTILLA", "3ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º", "3ª MEDICION EJEMPLARES PARA DETERMINAR CATEGORIA DE PARTICIPACIÓN (según reglamento)", "3ª MEDICIÓN EJEMPLARES PARA DETERMINAR CATEGORIA DE PARTICIPACIÓN (según Reglamento)", "3ª MEDICIÓN EJEMPLARES PARA DETERMINAR CATEGORÍA DE PARTICIPACIÓN (según Reglamento)", "4ª CARTILLA Y SUCESIVAS HASTA EL NÚMERO MÁXIMO DE PERROS QUE COMPONEN UNA JAURÍA SEGÚN EL REGLAMENTO", "4ª CARTILLA Y SUCESIVAS HASTA EL NÚMERO MÁXIMO DE PERROS QUE PUEDAN COMPONER UNA JAURÍA (según Reglamento)", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PARTERNIDAD (por perro) - SALIVA", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PARTERNIDAD (por perro) - SANGRE", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PATERNIDAD (por perro) / SALIVA", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PATERNIDAD (por perro) / SANGRE", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro)", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro) / SALIVA", "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro) / SANGRE", "ANOTACIÓN DE TÍTULOS", "ANOTACIÓN EN BASE DE DATOS DE PRUEBA ADN SIN ACOGERSE AL CONVENIO RSCE – HISPALAB (por perro)", "ANOTACIÓN EN BASE DE DATOS DE PRUEBA GENÉTICA O DE DISPLASIA (por prueba)", "ANOTACIÓN EN BASE DE DATOS PRUEBA ADN", "ANOTACIÓN PRUEBA DE DISPLASIA", "BAJA DE UN PERRO EN EL L.O.E. / R.R.C. (para su posterior inscripción en otros libros genealógicos)", "BAJA DE UN PERRO EN EL L.O.E./R.R.C.", "BAJA DE UN PERRO EN EL L.O.E./R.R.C. (para su posterior inscripción de otros libros genealógicos)", "BAJA DE UN PERRO EN EL LOE/RRC (para su posterior inscripción de otros libros genealógicos)", "CAMBIO PEDIGREE CUATRO GENERACIONES A PEDIGREE ORO SIN TRANSFERENCIA DE PROPIEDAD", "CAMBIO PEDIGREE TRES GENERACIONES A PEDIGREE ORO SIN TRANSFERENCIA DE PROPIEDAD", "CARTILLA ASISTENTE", "CARTILLA DE ASISTENTE", "CARTILLA DE PRUEBAS DEPORTIVAS DE  CAZA", "CARTILLA DE PRUEBAS DEPORTIVAS DE AGILITY", "CARTILLA DE PRUEBAS DEPORTIVAS DE CAZA", "CARTILLA DE PRUEBAS DEPORTIVAS DE TRABAJO", "CARTILLA DE PRUEBAS DEPORTIVAS DE TRABAJO PARA PERROS SIN PEDIGREE", "CARTILLA DE PRUEBAS DEPORTIVAS PARA PERROS SIN PEDIGREE", "CARTILLA PARA PRUEBAS DE CAMPO", "CARTILLA PARA PRUEBAS DEPORTIVAS DE CAMPO", "CARTILLA PARA PRUEBAS DEPORTIVAS POLIVALENTES - PERRO DE AGUA ESPAÑOL", "CERTIFICADO  DE COLA CORTA", "CERTIFICADO DE COLA CORTA", "CERTIFICADO DE TRABAJO", "CERTIFICADO DE TRABAJO PERROS DE MUESTRA RASTRO Y MADRIGUERA", "CERTIFICADO DE TRABAJO PERROS DE MUESTRA, RASTRO Y MADRIGUERA", "CERTIFICADO INTERNACIONAL DE DISPLASIA", "CESIÓN TEMPORAL DE PROPIEDAD DE LA HEMBRA REPRODUCTORA", "CESIÓN TEMPORAL DE PROPIEDAD DE LOS REPRODUCTORES", "CONFIRMACION DE RAZA (PERROS RAZAS ESPAÑOLAS)", "CONFIRMACION DE RAZA (perros de razas españolas)", "CONFIRMACIÓN DE RAZA (PERROS GRUPOS ÉTINICOS)", "CONFIRMACIÓN DE RAZA (perros de razas españolas)", "CUOTA ABONADO", "CUOTA ABONO JOVEN", "CUOTA ANUAL CLUB AGILITY", "CUOTA ANUAL CLUB DE AGILITY", "CUOTA ANUAL GRUPO DE TRABAJO", "CUOTA ANUAL MULTIDISCIPLINAR", "CUOTA ANUAL SOCIO", "CUOTA INICIAL SOCIO (incluida la cuota del primer año natural)", "CUOTA INICIAL SOCIO, incluida la cuota del primer año natural (260 + 30)", "DUPLICADO AFIJO", "DUPLICADO DE AFIJO", "DUPLICADO DE PEDIGREE CUATRO GENERACIONES", "DUPLICADO DE PEDIGREE TRES GENERACIONES", "DUPLICADO EXPORT PEDIGREE CUATRO GENERACIONES", "DUPLICADO EXPORT PEDIGREE TRES GENERACIONES", "DUPLICADO JUSTIFICANTE DE INSCRIPCIÓN", "DUPLICADO PEDIGREE_RSCE", "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN 3 GENERACIONES", "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN 4 GENERACIONES", "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN CUATRO GENERACIONES EN PAPEL", "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN TRES GENERACIONES EN PAPEL", "DUPLICADO PEDIGREE_RSCE EN PAPEL", "EDIGREE_RSCE PREMIUM TRES GENERACIONES SIN TRANSFERENCIA", "EXCEPCION, RECARGO 300% (más de 12 meses y menos de 18)  Art, 5.k) Reglamento L.O.E.", "EXCEPCION: RECARGO 300% (más de 12 meses y menos de 18 meses)  Reglamento L.O.E.", "EXPEDICIÓN DE INFORMES (precio por folio impreso)", "EXPEDICIÓN TITULOS DE CAMPEONES", "EXPEDICIÓN TITULOS LATIN CHAMPION", "EXPEDICIÓN TÍTULOS \"LATIN CHAMPION\"", "EXPEDICIÓN TÍTULOS DE CAMPEONES", "EXPEDICIÓN TÍTULOS “LATIN CHAMPION”", "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento)", "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento) EN ORIGEN", "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento) RSCE", "IMPRESO ALTA CAMADA", "IMPRESO ALTA DE CAMADA", "IMPRESO ALTA DE CAMADA (GRUPOS ÉTNICOS)", "IMPRESO ALTA DE CAMADA (PERROS DE RAZA)", "IMPRESO DIPLOMA DE INICIACIÓN (original y 3 copias)", "IMPRESO HOJA DE CALIFICACIÓN – PRUEBAS PERRO DE AGUA ESPAÑOL", "IMPRESO PARA PRUEBAS CACT - CACIT", "IMPRESO PARA PRUEBAS DE C.A.C.T. - C.A.C.I.T.", "IMPRESO PARA PRUEBAS DE CACT – CACIT", "IMPRESO PRUEBAS APTITUDES NATURALES (original y 3 copias)", "INFORMES SOLICITADOS A LA R.S.C.E. (precio por folio impreso)", "INFORMES SOLICITADOS A LA RSCE (precio por folio impreso", "INFORMES SOLICITADOS A LA RSCE (precio por folio impreso)", "INSCRIPCION DE OTROS LIBROS GENEALOGICOS (NO ADMITE PLUS DE URGENCIA)", "INSCRIPCION PERROS DE OTROS LIBROS GENEALÓGICOS (ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA)", "INSCRIPCIÓN CACHORRO", "INSCRIPCIÓN CACHORRO ACCESS LBO/RBR", "INSCRIPCIÓN CACHORRO DE RAZAS ESPAÑOLAS BONIFICADAS", "INSCRIPCIÓN CACHORRO EN EL REGISTRO DE GRUPOS ÉTNICOS", "INSCRIPCIÓN CACHORRO PREMIUM LOE/RRC", "INSCRIPCIÓN CACHORRO RAZAS ESPAÑOLAS BONIFICADAS ( Del 01/05/2023 al 30/04/2024)", "INSCRIPCIÓN CACHORRO RAZAS ESPAÑOLAS BONIFICADAS ( Del 01/05/2024 al 30/04/2025)", "INSCRIPCIÓN IMPORTADOS", "INSCRIPCIÓN PERRO con más de 18 meses aportando prueba ADN progenitores y huella genética de compatibilidad", "INSCRIPCIÓN PERROS DE OTROS LIBROS GENEALÓGICOS", "INSCRIPCIÓN PERROS DE OTROS LIBROS GENEALÓGICOS - ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA", "INSCRIPCIÓN PERROS PROCEDENTES DE OTROS LIBROS GENEALÓGICOS", "INSCRIPCIÓN PERROS PROCEDENTES DE OTROS LIBROS GENEALÓGICOS (ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA)", "LICENCIA INCIAL", "LICENCIA INICIAL", "MEJORA A PEDIGREE_RSCE PLUS CUATRO GENERACIONES SIN TRANSFERENCIA", "MEJORA A PEDIGREE_RSCE PLUS RRC (sin genealogía completa) CON TRANSFERENCIA", "MEJORA A PEDIGREE_RSCE PLUS RRC (sin genealógica completa) SIN TRANSFERENCIA", "MEJORA A PEDIGREE_RSCE PLUS TRES GENERACIONES SIN TRANSFERENCIA", "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE CUATRO GENERACIONES SIN TRANSFERENCIA", "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL", "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE TRES GENERACIONES SIN TRANSFERENCIA", "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL", "MEJORA DE PEDIGREE_RSCE ACCESS RBR A PEDIGREE_RSCE PREMIUM RRC SIN TRANSFERENCIA", "MEJORA DE PEDIGREE_RSCE ACCESS RBR A PEDIGREE_RSCE PREMIUM RRC SIN TRANSFERENCIA EN PAPEL", "NOTIFICACIÓN NACIMIENTO CACHORRO (formulario de identificación)", "OTRAS ANOTACIONES EN EL LOE / RRC / RGE", "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES  (SI CONSTAN EN EL L.O.E.)", "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES  (si consta en el L.O.E..)", "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)", "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)", "PEDIGREE CON TRANSFERENCIA TRES GENERACIONES", "PEDIGREE DE EXPORTACION (4 generaciones)", "PEDIGREE DE EXPORTACIÓN CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)", "PEDIGREE DE EXPORTACIÓN CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)", "PEDIGREE DE EXPORTACIÓN TRES GENERACIONES", "PEDIGREE ORO", "PEDIGREE ORO CON TRANSFERENCIA DE PROPIEDAD", "PEDIGREE ORO SIN TRANSFERENCIA", "PEDIGREE PARA PERROS INSCRITOS EN EL R.R.C. (sin genealógica completa)", "PEDIGREE PARA PERROS INSCRITOS EN EL RRC (sin genealógia completa)", "PEDIGREE PARA PERROS INSCRITOS EN EL RRC (sin genealógica completa)", "PEDIGREE SIN TRANSFERENCIA", "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES  (SI CONSTAN EN EL L.O.E.)", "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES  (si consta en el  L.O.E.)", "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (Si la genealogía consta en la base de datos del LOE)", "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)", "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)", "PEDIGREE_RSCE ACCESS LBO CON TRANSFERENCIA", "PEDIGREE_RSCE ACCESS LBO SIN TRANSFERENCIA", "PEDIGREE_RSCE ACCESS RBR CON TRANSFERENCIA", "PEDIGREE_RSCE ACCESS RBR CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE ACCESS RBR SIN TRANSFERENCIA", "PEDIGREE_RSCE ACCESS SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE ACCESS TRES GENERACIONES CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE DE EXPORTACION 3 GENERACIONES", "PEDIGREE_RSCE DE EXPORTACION 4 GENERACIONES", "PEDIGREE_RSCE DE EXPORTACIÓN CUATRO GENERACIONES EN PAPEL", "PEDIGREE_RSCE DE EXPORTACIÓN TRES GENERACIONES EN PAPEL", "PEDIGREE_RSCE PLUS CUATRO GENERACIONES CON TRANSFERENCIA", "PEDIGREE_RSCE PLUS CUATRO GENERACIONES CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PLUS CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PLUS RRC (sin genealógica completa) SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PLUS RRC CON TRANSFERENCIA", "PEDIGREE_RSCE PLUS TRES GENERACIONES CON TRANSFERENCIA", "PEDIGREE_RSCE PLUS TRES GENERACIONES CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PLUS TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES CON TRANSFERENCIA", "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES SIN TRANSFERENCIA", "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) CON TRANSFERENCIA", "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) SIN TRANSFERENCIA", "PEDIGREE_RSCE PREMIUM TRES GENERACIONES CON TRANSFERENCIA", "PEDIGREE_RSCE PREMIUM TRES GENERACIONES CON TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUM TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL", "PEDIGREE_RSCE PREMIUN 3 GENERACIONES SIN TRANSFERENCIA", "PLUS ENVÍO CARTILLA POR MENSAJERÍA (hasta un máximo de 10 cartillas por guía)", "PLUS ENVÍO CARTILLA POR MENSAJERÍA (hasta un máximo de 10 cartillas por propietario)", "PLUS POR TRAMITACIÓN URGENTE DE CAMADAS, PEDIGREES Y PEDIGREES DE EXPORTACIÓN", "POR CADA ANOTACIÓN DE INFORME ADN O INFORME DE DISPLASIA EN BASE DE DATOS", "POR CADA ANOTACIÓN EN BASE DE DATOS DE PRUEBA GENÉTICA, INFORME DE ADN O INFORME DE DISPLASIA", "POR PERRO INSCRITO EN CADA PRUEBA CON CAC / CACT", "PRUEBA ADN / PRUEBAS PATERNIDAD, CON ANOTACIÓN EN BASE DE DATOS, ACOGIÉNDOSE AL CONVENIO RSCE – HISPALAB (por perro)", "PRUEBA ADN/PATERNIDAD Y ANOTACIÓN EN BASE DE DATOS, SEGÚN CONVENIO RSCE – HISPALAB (por perro)", "PRUEBA ADN/PATERNIDAD Y ANOTACIÓN EN BASE DE DATOS-CONVENIO RSCE – HISPALAB (por perro) RAZAS ESPAÑOLAS VULNERABLES", "RECARGO 100% (más de 6 meses y menos de 9 meses)  Reglamento L.O.E.", "RECARGO 100% (más de 6 meses y menos de 9) Art. 6.)13 Reglamento L.O.E.", "RECARGO 100% EN INSCRIPCION CACHORRO (más de 6 meses y menos de 9 meses de edad) Reglamento LOE)", "RECARGO 100% EN INSCRIPCIÓN CACHORRO (más de 6 meses y menos de 9 meses de edad)", "RECARGO 100% EN INSCRIPCIÓN CACHORRO (más de 6 meses y menos de 9 meses de edad) Art. 6.13 Reglamento L.O.E.", "RECARGO 200% (más de 9 meses, menos de 12 meses) Reglamento L.O.E.", "RECARGO 200% (más de 9 meses, menos de 12) Art. 6.)13 Reglamento L.O.E.", "RECARGO 200% EN INSCRIPCION CACHORRO (más de 9 meses y menos de 12 meses de edad)  Reglamento LOE)", "RECARGO 200% EN INSCRIPCIÓN CACHORRO (más de 9 meses y menos de 12 meses de edad)", "RECARGO 200% EN INSCRIPCIÓN CACHORRO (más de 9 meses y menos de 12 meses de edad) Art. 6.13 Reglamento L.O.E.", "RECARGO 300% EN INSCRIPCION CACHORRO (más de 12 meses y menos de 18 meses de edad) Reglamento LOE)", "RECARGO 300% EN INSCRIPCIÓN CACHORRO (más de 12 meses y menos de 18 meses de edad)", "RECARGO 300% EN INSCRIPCIÓN CACHORRO (más de 12 meses y menos de 18 meses de edad) Art. 6.13 Reglamento L.O.E.", "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/ RBR con más de 6 meses y menos de 9 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 12 meses y menos de 18 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 6 meses y menos de 9 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 9 meses y menos de 12 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 12 meses y menos de 18 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 6 meses y menos de 9 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 9 meses y menos de 12 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 12 meses y menos de 18 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 6 meses y menos de 9 meses de edad", "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 9 meses y menos de 12 meses de edad", "RECTIFICACION AFIJO", "RECTIFICACIÓN DE AFIJO", "RECURSO AL DIAGNOSTICO DE DISPLASIA", "RECURSO AL DIAGNÓSTICO DE DISPLASIA", "RECURSO COMISIÓN LIBRO DE ORÍGENES ESPAÑOL", "REGISTRO INCIAL DE RAZA", "REGISTRO INICIAL DE RAZA", "REGISTRO INICIAL GRUPO ÉTNICO", "REGISTRO INICIAL RAZAS ESPAÑOLAS", "REGISTRO INICIAL RAZAS ESPAÑOLAS VULNERABLES", "REGISTRO INICIAL RESTO DE RAZAS", "RENOVACIÓN ANUAL LICENCIA", "RENOVACIÓN ANUAL LICIENCIA", "SIN DESCRIPCIÓN (tasa fija 0,25€)", "SOLICITUD AFIJO", "SOLICITUD DE AFIJO", "TRAMITACION TITULOS FCI", "TRAMITACION URGENTE  CARTILLA DE ASISTENTE (exclusivamente a través de las oficinas de la RSCE)", "TRAMITACION URGENTE CARTILLAS DE PRUEBAS DEPORTIVAS (exclusivamente a través de las oficinas de la RSCE)", "TRAMITACION URGENTE CERTIFICADOS DE COLA CORTA (se realizará exclusivamente a través de las oficinas de la RSCE)", "TRAMITACION URGENTE CERTIFICADOS DE TRABAJO (se realizará exclusivamente a través de las oficinas de la RSCE)", "TRAMITACION URGENTE DE CAMADAS Y PEDIGREES (exclusivamente a través de las oficinas de la RSCE)", "TRAMITACIÓN TITULOS F.C.I.", "TRAMITACIÓN TITULOS FCI", "TRAMITACIÓN TÍTULOS FCI", "TRAMITACIÓN URGENTE DE CARTILLAS DE ASISTENTE", "TRAMITACIÓN URGENTE DE CARTILLAS DE PRUEBAS DEPORTIVAS", "TRAMITACIÓN URGENTE DE CERTIFICADOS DE COLA CORTA", "TRAMITACIÓN URGENTE DE CERTIFICADOS DE TRABAJO", "TRAMITACIÓN URGENTE DE INSCRIPCIONES EN LOE/RRC, PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "TRAMITACIÓN URGENTE DE PEDIGRÍES, TRANSFERENCIAS DE PROPIEDAD E INSCRIPCIONES EN LOE/RRC", "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CADERA", "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CODO", "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CODOS", "TRANSFERENCIA DE PROPIEDAD EN PEDIGREE, REGISTRO INICIAL O PERRO IMPORTADO", "TRANSFERENCIA DE PROPIEDAD GRUPOS ETNICOS", "TRANSFERENCIA DE PROPIEDAD REGISTRO INICIAL O PERRO IMPORTADO"], "categories": ["AFIJOS", "AGILITY", "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "DISPLASIA", "DUPLICADOS DE PEDIGRÍES", "IDENTIFICACIÓN GENÉTICA", "IMPRESOS", "INSCRIPCIONES", "INSCRIPCIONES ACCESS LBO/RBR", "INSCRIPCIONES EN EL LOE / RRC", "INSCRIPCIONES PREMIUM LOE/RRC", "MEJORA DE ACCESS A PREMIUM", "MUESTRA, RASTRO Y MADRIGUERA", "OTRAS INSCRIPCIONES", "OTROS SERVICIOS", "PEDIGRÍES ACCESS LBO/RBR", "PEDIGRÍES DE EXPORTACIÓN", "PEDIGRÍES PLUS", "PEDIGRÍES PREMIUM LOE/RRC", "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "PLUS POR TRAMITACIÓN URGENTE", "REGISTROS INICIALES Y CONFIRMACIONES DE RAZA", "SALUD DEL PERRO", "SOCIOS Y ABONADOS", "Sin categorizar", "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)"], "data": {"1ª CARTILLA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 39.6, "no_vat": 32.73}], "User": [{"year": 2019, "with_vat": 51.0, "no_vat": 42.15}], "Canine_Collaborator": []}}, "1ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2020, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2023, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2024, "with_vat": 21.7, "no_vat": 17.93}, {"year": 2025, "with_vat": 22.8, "no_vat": 18.84}, {"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2020, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2021, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2022, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2023, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2024, "with_vat": 28.9, "no_vat": 23.88}, {"year": 2025, "with_vat": 30.3, "no_vat": 25.04}, {"year": 2026, "with_vat": 32.0, "no_vat": 26.45}], "Canine_Collaborator": []}}, "2ª CARTILLA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 23.8, "no_vat": 19.67}], "User": [{"year": 2019, "with_vat": 30.4, "no_vat": 25.12}], "Canine_Collaborator": []}}, "2ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2020, "with_vat": 24.0, "no_vat": 19.83}, {"year": 2021, "with_vat": 24.0, "no_vat": 19.83}, {"year": 2022, "with_vat": 12.6, "no_vat": 10.41}, {"year": 2023, "with_vat": 12.6, "no_vat": 10.41}, {"year": 2024, "with_vat": 13.0, "no_vat": 10.74}, {"year": 2025, "with_vat": 13.7, "no_vat": 11.32}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "User": [{"year": 2020, "with_vat": 32.0, "no_vat": 26.45}, {"year": 2021, "with_vat": 32.0, "no_vat": 26.45}, {"year": 2022, "with_vat": 16.8, "no_vat": 13.88}, {"year": 2023, "with_vat": 16.8, "no_vat": 13.88}, {"year": 2024, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2025, "with_vat": 18.2, "no_vat": 15.04}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "Canine_Collaborator": []}}, "3ª CARTILLA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 15.9, "no_vat": 13.14}], "User": [{"year": 2019, "with_vat": 20.4, "no_vat": 16.86}], "Canine_Collaborator": []}}, "3ª CARTILLA PERROS DE RASTRO DEL GRUPO 6º": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2020, "with_vat": 16.0, "no_vat": 13.22}, {"year": 2021, "with_vat": 16.0, "no_vat": 13.22}, {"year": 2022, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2023, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2024, "with_vat": 8.7, "no_vat": 7.19}, {"year": 2025, "with_vat": 9.1, "no_vat": 7.52}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2020, "with_vat": 21.3, "no_vat": 17.6}, {"year": 2021, "with_vat": 21.3, "no_vat": 17.6}, {"year": 2022, "with_vat": 11.2, "no_vat": 9.26}, {"year": 2023, "with_vat": 11.2, "no_vat": 9.26}, {"year": 2024, "with_vat": 11.5, "no_vat": 9.5}, {"year": 2025, "with_vat": 12.1, "no_vat": 10.0}, {"year": 2026, "with_vat": 14.0, "no_vat": 11.57}], "Canine_Collaborator": [{"year": 2022, "with_vat": 21.54, "no_vat": 17.8}, {"year": 2023, "with_vat": 21.54, "no_vat": 17.8}, {"year": 2024, "with_vat": 22.23, "no_vat": 18.37}, {"year": 2025, "with_vat": 23.3, "no_vat": 19.26}, {"year": 2026, "with_vat": 23.98, "no_vat": 19.82}]}}, "3ª MEDICION EJEMPLARES PARA DETERMINAR CATEGORIA DE PARTICIPACIÓN (según reglamento)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 50.8, "no_vat": 41.98}], "User": [{"year": 2018, "with_vat": 52.4, "no_vat": 43.31}], "Canine_Collaborator": []}}, "3ª MEDICIÓN EJEMPLARES PARA DETERMINAR CATEGORIA DE PARTICIPACIÓN (según Reglamento)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 52.0, "no_vat": 42.98}], "User": [{"year": 2019, "with_vat": 53.7, "no_vat": 44.38}], "Canine_Collaborator": []}}, "3ª MEDICIÓN EJEMPLARES PARA DETERMINAR CATEGORÍA DE PARTICIPACIÓN (según Reglamento)": {"category": "AGILITY", "prices": {"Member": [{"year": 2020, "with_vat": 41.3, "no_vat": 34.13}, {"year": 2021, "with_vat": 41.3, "no_vat": 34.13}, {"year": 2022, "with_vat": 43.4, "no_vat": 35.87}, {"year": 2023, "with_vat": 43.4, "no_vat": 35.87}, {"year": 2024, "with_vat": 44.7, "no_vat": 36.94}, {"year": 2025, "with_vat": 46.9, "no_vat": 38.76}, {"year": 2026, "with_vat": 50.0, "no_vat": 41.32}], "User": [{"year": 2020, "with_vat": 55.0, "no_vat": 45.45}, {"year": 2021, "with_vat": 55.0, "no_vat": 45.45}, {"year": 2022, "with_vat": 57.8, "no_vat": 47.77}, {"year": 2023, "with_vat": 57.8, "no_vat": 47.77}, {"year": 2024, "with_vat": 59.6, "no_vat": 49.26}, {"year": 2025, "with_vat": 62.6, "no_vat": 51.74}, {"year": 2026, "with_vat": 67.0, "no_vat": 55.37}], "Canine_Collaborator": []}}, "4ª CARTILLA Y SUCESIVAS HASTA EL NÚMERO MÁXIMO DE PERROS QUE COMPONEN UNA JAURÍA SEGÚN EL REGLAMENTO": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 8.0, "no_vat": 6.61}], "User": [{"year": 2019, "with_vat": 10.3, "no_vat": 8.51}], "Canine_Collaborator": []}}, "4ª CARTILLA Y SUCESIVAS HASTA EL NÚMERO MÁXIMO DE PERROS QUE PUEDAN COMPONER UNA JAURÍA (según Reglamento)": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2020, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2021, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2022, "with_vat": 4.2, "no_vat": 3.47}, {"year": 2023, "with_vat": 4.2, "no_vat": 3.47}, {"year": 2024, "with_vat": 4.3, "no_vat": 3.55}, {"year": 2025, "with_vat": 4.5, "no_vat": 3.72}, {"year": 2026, "with_vat": 5.0, "no_vat": 4.13}], "User": [{"year": 2020, "with_vat": 10.7, "no_vat": 8.84}, {"year": 2021, "with_vat": 10.7, "no_vat": 8.84}, {"year": 2022, "with_vat": 5.7, "no_vat": 4.71}, {"year": 2023, "with_vat": 5.7, "no_vat": 4.71}, {"year": 2024, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2025, "with_vat": 6.2, "no_vat": 5.12}, {"year": 2026, "with_vat": 7.0, "no_vat": 5.79}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PARTERNIDAD (por perro) - SALIVA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 48.3, "no_vat": 39.92}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PARTERNIDAD (por perro) - SANGRE": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 39.0, "no_vat": 32.23}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PATERNIDAD (por perro) / SALIVA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 47.2, "no_vat": 39.01}], "User": [{"year": 2018, "with_vat": 47.2, "no_vat": 39.01}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA A.D.N. / PRUEBAS DE PATERNIDAD (por perro) / SANGRE": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 38.1, "no_vat": 31.49}], "User": [{"year": 2018, "with_vat": 38.1, "no_vat": 31.49}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro)": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2022, "with_vat": 34.7, "no_vat": 28.68}], "User": [{"year": 2022, "with_vat": 42.0, "no_vat": 34.71}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro) / SALIVA": {"category": "IDENTIFICACIÓN GENÉTICA", "prices": {"Member": [{"year": 2020, "with_vat": 49.3, "no_vat": 40.74}, {"year": 2021, "with_vat": 49.3, "no_vat": 40.74}], "User": [{"year": 2020, "with_vat": 49.3, "no_vat": 40.74}, {"year": 2021, "with_vat": 65.75, "no_vat": 54.34}], "Canine_Collaborator": []}}, "ANALÍTICA IDENTIFICACIÓN GENÉTICA CANINA ADN / PRUEBAS DE PATERNIDAD (por perro) / SANGRE": {"category": "IDENTIFICACIÓN GENÉTICA", "prices": {"Member": [{"year": 2020, "with_vat": 39.8, "no_vat": 32.89}, {"year": 2021, "with_vat": 39.8, "no_vat": 32.89}], "User": [{"year": 2020, "with_vat": 39.8, "no_vat": 32.89}, {"year": 2021, "with_vat": 53.1, "no_vat": 43.88}], "Canine_Collaborator": []}}, "ANOTACIÓN DE TÍTULOS": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2025, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2026, "with_vat": 6.5, "no_vat": 5.37}], "User": [{"year": 2025, "with_vat": 8.7, "no_vat": 7.19}, {"year": 2026, "with_vat": 8.7, "no_vat": 7.19}], "Canine_Collaborator": []}}, "ANOTACIÓN EN BASE DE DATOS DE PRUEBA ADN SIN ACOGERSE AL CONVENIO RSCE – HISPALAB (por perro)": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2023, "with_vat": 25.0, "no_vat": 20.66}, {"year": 2024, "with_vat": 25.8, "no_vat": 21.32}], "User": [{"year": 2023, "with_vat": 25.0, "no_vat": 20.66}, {"year": 2024, "with_vat": 25.8, "no_vat": 21.32}], "Canine_Collaborator": []}}, "ANOTACIÓN EN BASE DE DATOS DE PRUEBA GENÉTICA O DE DISPLASIA (por prueba)": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2023, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2024, "with_vat": 6.5, "no_vat": 5.37}], "User": [{"year": 2023, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2024, "with_vat": 8.7, "no_vat": 7.19}], "Canine_Collaborator": []}}, "ANOTACIÓN EN BASE DE DATOS PRUEBA ADN": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2025, "with_vat": 25.8, "no_vat": 21.32}, {"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "User": [{"year": 2025, "with_vat": 25.8, "no_vat": 21.32}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "ANOTACIÓN PRUEBA DE DISPLASIA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2025, "with_vat": 6.8, "no_vat": 5.62}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2025, "with_vat": 9.1, "no_vat": 7.52}, {"year": 2026, "with_vat": 14.0, "no_vat": 11.57}], "Canine_Collaborator": []}}, "BAJA DE UN PERRO EN EL L.O.E. / R.R.C. (para su posterior inscripción en otros libros genealógicos)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 26.4, "no_vat": 21.82}], "User": [{"year": 2019, "with_vat": 38.9, "no_vat": 32.15}], "Canine_Collaborator": []}}, "BAJA DE UN PERRO EN EL L.O.E./R.R.C.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 22.05, "no_vat": 18.22}, {"year": 2019, "with_vat": 22.57, "no_vat": 18.65}, {"year": 2020, "with_vat": 23.03, "no_vat": 19.03}, {"year": 2022, "with_vat": 24.19, "no_vat": 19.99}, {"year": 2023, "with_vat": 24.19, "no_vat": 19.99}, {"year": 2024, "with_vat": 24.94, "no_vat": 20.61}, {"year": 2025, "with_vat": 26.21, "no_vat": 21.66}, {"year": 2026, "with_vat": 30.15, "no_vat": 24.92}]}}, "BAJA DE UN PERRO EN EL L.O.E./R.R.C. (para su posterior inscripción de otros libros genealógicos)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 25.8, "no_vat": 21.32}], "User": [{"year": 2018, "with_vat": 38.0, "no_vat": 31.4}], "Canine_Collaborator": [{"year": 2021, "with_vat": 23.03, "no_vat": 19.03}]}}, "BAJA DE UN PERRO EN EL LOE/RRC (para su posterior inscripción de otros libros genealógicos)": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2020, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2021, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2022, "with_vat": 28.4, "no_vat": 23.47}, {"year": 2023, "with_vat": 28.4, "no_vat": 23.47}, {"year": 2024, "with_vat": 29.3, "no_vat": 24.21}, {"year": 2025, "with_vat": 30.8, "no_vat": 25.45}, {"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "User": [{"year": 2020, "with_vat": 39.7, "no_vat": 32.81}, {"year": 2021, "with_vat": 39.7, "no_vat": 32.81}, {"year": 2022, "with_vat": 41.7, "no_vat": 34.46}, {"year": 2023, "with_vat": 41.7, "no_vat": 34.46}, {"year": 2024, "with_vat": 43.0, "no_vat": 35.54}, {"year": 2025, "with_vat": 45.2, "no_vat": 37.36}, {"year": 2026, "with_vat": 52.0, "no_vat": 42.98}], "Canine_Collaborator": []}}, "CAMBIO PEDIGREE CUATRO GENERACIONES A PEDIGREE ORO SIN TRANSFERENCIA DE PROPIEDAD": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2021, "with_vat": 5.0, "no_vat": 4.13}], "User": [{"year": 2021, "with_vat": 7.5, "no_vat": 6.2}], "Canine_Collaborator": [{"year": 2022, "with_vat": 4.31, "no_vat": 3.56}, {"year": 2023, "with_vat": 4.31, "no_vat": 3.56}]}}, "CAMBIO PEDIGREE TRES GENERACIONES A PEDIGREE ORO SIN TRANSFERENCIA DE PROPIEDAD": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2021, "with_vat": 9.0, "no_vat": 7.44}], "User": [{"year": 2021, "with_vat": 13.3, "no_vat": 10.99}], "Canine_Collaborator": [{"year": 2022, "with_vat": 7.64, "no_vat": 6.31}, {"year": 2023, "with_vat": 7.64, "no_vat": 6.31}]}}, "CARTILLA ASISTENTE": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2019, "with_vat": 32.1, "no_vat": 26.53}, {"year": 2020, "with_vat": 33.7, "no_vat": 27.85}, {"year": 2021, "with_vat": 33.7, "no_vat": 27.85}, {"year": 2022, "with_vat": 35.4, "no_vat": 29.26}, {"year": 2023, "with_vat": 35.4, "no_vat": 29.26}, {"year": 2024, "with_vat": 36.5, "no_vat": 30.17}, {"year": 2025, "with_vat": 37.3, "no_vat": 30.83}, {"year": 2026, "with_vat": 40.0, "no_vat": 33.06}], "User": [{"year": 2019, "with_vat": 55.0, "no_vat": 45.45}, {"year": 2020, "with_vat": 45.0, "no_vat": 37.19}, {"year": 2021, "with_vat": 45.0, "no_vat": 37.19}, {"year": 2022, "with_vat": 47.3, "no_vat": 39.09}, {"year": 2023, "with_vat": 47.3, "no_vat": 39.09}, {"year": 2024, "with_vat": 48.8, "no_vat": 40.33}, {"year": 2025, "with_vat": 51.2, "no_vat": 42.31}, {"year": 2026, "with_vat": 55.0, "no_vat": 45.45}], "Canine_Collaborator": []}}, "CARTILLA DE ASISTENTE": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 31.3, "no_vat": 25.87}], "User": [{"year": 2018, "with_vat": 53.7, "no_vat": 44.38}], "Canine_Collaborator": []}}, "CARTILLA DE PRUEBAS DEPORTIVAS DE  CAZA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 38.7, "no_vat": 31.98}], "User": [{"year": 2018, "with_vat": 49.8, "no_vat": 41.16}], "Canine_Collaborator": []}}, "CARTILLA DE PRUEBAS DEPORTIVAS DE AGILITY": {"category": "AGILITY", "prices": {"Member": [{"year": 2018, "with_vat": 38.7, "no_vat": 31.98}, {"year": 2019, "with_vat": 39.6, "no_vat": 32.73}, {"year": 2020, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 42.0, "no_vat": 34.71}, {"year": 2023, "with_vat": 42.0, "no_vat": 34.71}, {"year": 2024, "with_vat": 43.3, "no_vat": 35.79}, {"year": 2025, "with_vat": 45.5, "no_vat": 37.6}, {"year": 2026, "with_vat": 48.0, "no_vat": 39.67}], "User": [{"year": 2018, "with_vat": 49.8, "no_vat": 41.16}, {"year": 2019, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2020, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2021, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2022, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2023, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2024, "with_vat": 57.7, "no_vat": 47.69}, {"year": 2025, "with_vat": 60.6, "no_vat": 50.08}, {"year": 2026, "with_vat": 64.0, "no_vat": 52.89}], "Canine_Collaborator": []}}, "CARTILLA DE PRUEBAS DEPORTIVAS DE CAZA": {"category": "MUESTRA, RASTRO Y MADRIGUERA", "prices": {"Member": [{"year": 2019, "with_vat": 39.6, "no_vat": 32.73}, {"year": 2020, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}], "User": [{"year": 2019, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2020, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2021, "with_vat": 53.3, "no_vat": 44.05}], "Canine_Collaborator": [{"year": 2024, "with_vat": 22.23, "no_vat": 18.37}, {"year": 2025, "with_vat": 23.3, "no_vat": 19.26}]}}, "CARTILLA DE PRUEBAS DEPORTIVAS DE TRABAJO": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": 46.9, "no_vat": 38.76}, {"year": 2019, "with_vat": 48.0, "no_vat": 39.67}, {"year": 2020, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 42.0, "no_vat": 34.71}, {"year": 2023, "with_vat": 42.0, "no_vat": 34.71}, {"year": 2024, "with_vat": 43.3, "no_vat": 35.79}, {"year": 2025, "with_vat": 45.5, "no_vat": 37.6}, {"year": 2026, "with_vat": 48.0, "no_vat": 39.67}], "User": [{"year": 2018, "with_vat": 64.5, "no_vat": 53.31}, {"year": 2019, "with_vat": 66.0, "no_vat": 54.55}, {"year": 2020, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2021, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2022, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2023, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2024, "with_vat": 57.7, "no_vat": 47.69}, {"year": 2025, "with_vat": 60.6, "no_vat": 50.08}, {"year": 2026, "with_vat": 64.0, "no_vat": 52.89}], "Canine_Collaborator": []}}, "CARTILLA DE PRUEBAS DEPORTIVAS DE TRABAJO PARA PERROS SIN PEDIGREE": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 5.3, "no_vat": 4.38}], "User": [{"year": 2019, "with_vat": 5.5, "no_vat": 4.55}], "Canine_Collaborator": []}}, "CARTILLA DE PRUEBAS DEPORTIVAS PARA PERROS SIN PEDIGREE": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2020, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2021, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2022, "with_vat": 6.3, "no_vat": 5.21}], "User": [{"year": 2020, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2021, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2022, "with_vat": 8.4, "no_vat": 6.94}], "Canine_Collaborator": []}}, "CARTILLA PARA PRUEBAS DE CAMPO": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2022, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2023, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2024, "with_vat": 21.7, "no_vat": 17.93}, {"year": 2025, "with_vat": 22.8, "no_vat": 18.84}, {"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2022, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2023, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2024, "with_vat": 28.9, "no_vat": 23.88}, {"year": 2025, "with_vat": 30.3, "no_vat": 25.04}, {"year": 2026, "with_vat": 32.0, "no_vat": 26.45}], "Canine_Collaborator": []}}, "CARTILLA PARA PRUEBAS DEPORTIVAS DE CAMPO": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2022, "with_vat": 21.54, "no_vat": 17.8}, {"year": 2023, "with_vat": 21.54, "no_vat": 17.8}]}}, "CARTILLA PARA PRUEBAS DEPORTIVAS POLIVALENTES - PERRO DE AGUA ESPAÑOL": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": 38.7, "no_vat": 31.98}, {"year": 2019, "with_vat": 39.6, "no_vat": 32.73}, {"year": 2020, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2023, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2024, "with_vat": 21.7, "no_vat": 17.93}, {"year": 2025, "with_vat": 22.8, "no_vat": 18.84}, {"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2018, "with_vat": 49.8, "no_vat": 41.16}, {"year": 2019, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2020, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2021, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2022, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2023, "with_vat": 28.0, "no_vat": 23.14}, {"year": 2024, "with_vat": 28.9, "no_vat": 23.88}, {"year": 2025, "with_vat": 30.3, "no_vat": 25.04}, {"year": 2026, "with_vat": 32.0, "no_vat": 26.45}], "Canine_Collaborator": [{"year": 2022, "with_vat": 21.54, "no_vat": 17.8}, {"year": 2023, "with_vat": 21.54, "no_vat": 17.8}, {"year": 2024, "with_vat": 22.23, "no_vat": 18.37}, {"year": 2025, "with_vat": 23.3, "no_vat": 19.26}]}}, "CERTIFICADO  DE COLA CORTA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 6.5, "no_vat": 5.37}], "User": [{"year": 2019, "with_vat": 10.0, "no_vat": 8.26}], "Canine_Collaborator": []}}, "CERTIFICADO DE COLA CORTA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2020, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2021, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2022, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2023, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2024, "with_vat": 8.1, "no_vat": 6.69}, {"year": 2025, "with_vat": 8.5, "no_vat": 7.02}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2020, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2021, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2022, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2023, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "CERTIFICADO DE TRABAJO": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2019, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2020, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2021, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2022, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2023, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2024, "with_vat": 8.1, "no_vat": 6.69}, {"year": 2025, "with_vat": 8.5, "no_vat": 7.02}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2018, "with_vat": 9.7, "no_vat": 8.02}, {"year": 2019, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2020, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2021, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2022, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2023, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "CERTIFICADO DE TRABAJO PERROS DE MUESTRA RASTRO Y MADRIGUERA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 6.5, "no_vat": 5.37}], "User": [{"year": 2019, "with_vat": 10.0, "no_vat": 8.26}], "Canine_Collaborator": []}}, "CERTIFICADO DE TRABAJO PERROS DE MUESTRA, RASTRO Y MADRIGUERA": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2018, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2020, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2021, "with_vat": 7.5, "no_vat": 6.2}, {"year": 2022, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2023, "with_vat": 7.9, "no_vat": 6.53}, {"year": 2024, "with_vat": 8.1, "no_vat": 6.69}, {"year": 2025, "with_vat": 8.5, "no_vat": 7.02}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2018, "with_vat": 9.7, "no_vat": 8.02}, {"year": 2020, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2021, "with_vat": 10.0, "no_vat": 8.26}, {"year": 2022, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2023, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "CERTIFICADO INTERNACIONAL DE DISPLASIA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2018, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2019, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2020, "with_vat": 6.7, "no_vat": 5.54}, {"year": 2021, "with_vat": 6.7, "no_vat": 5.54}, {"year": 2022, "with_vat": 7.0, "no_vat": 5.79}, {"year": 2023, "with_vat": 7.0, "no_vat": 5.79}, {"year": 2024, "with_vat": 7.2, "no_vat": 5.95}, {"year": 2025, "with_vat": 7.6, "no_vat": 6.28}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2018, "with_vat": 10.2, "no_vat": 8.43}, {"year": 2019, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2020, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2021, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2022, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2023, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2024, "with_vat": 11.7, "no_vat": 9.67}, {"year": 2025, "with_vat": 12.3, "no_vat": 10.17}, {"year": 2026, "with_vat": 17.0, "no_vat": 14.05}], "Canine_Collaborator": []}}, "CESIÓN TEMPORAL DE PROPIEDAD DE LA HEMBRA REPRODUCTORA": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2020, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2021, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2022, "with_vat": 18.2, "no_vat": 15.04}, {"year": 2023, "with_vat": 18.2, "no_vat": 15.04}, {"year": 2024, "with_vat": 18.8, "no_vat": 15.54}, {"year": 2025, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2026, "with_vat": 42.0, "no_vat": 34.71}], "User": [{"year": 2020, "with_vat": 20.9, "no_vat": 17.27}, {"year": 2021, "with_vat": 20.9, "no_vat": 17.27}, {"year": 2022, "with_vat": 21.9, "no_vat": 18.1}, {"year": 2023, "with_vat": 21.9, "no_vat": 18.1}, {"year": 2024, "with_vat": 22.6, "no_vat": 18.68}, {"year": 2025, "with_vat": 60.0, "no_vat": 49.59}, {"year": 2026, "with_vat": 63.0, "no_vat": 52.07}], "Canine_Collaborator": []}}, "CESIÓN TEMPORAL DE PROPIEDAD DE LOS REPRODUCTORES": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 16.5, "no_vat": 13.64}, {"year": 2019, "with_vat": 16.9, "no_vat": 13.97}], "User": [{"year": 2018, "with_vat": 19.9, "no_vat": 16.45}, {"year": 2019, "with_vat": 20.4, "no_vat": 16.86}], "Canine_Collaborator": [{"year": 2018, "with_vat": 11.52, "no_vat": 9.52}, {"year": 2019, "with_vat": 11.81, "no_vat": 9.76}, {"year": 2020, "with_vat": 12.1, "no_vat": 10.0}, {"year": 2021, "with_vat": 12.1, "no_vat": 10.0}, {"year": 2022, "with_vat": 12.68, "no_vat": 10.48}, {"year": 2023, "with_vat": 12.68, "no_vat": 10.48}, {"year": 2024, "with_vat": 13.08, "no_vat": 10.81}, {"year": 2025, "with_vat": 34.73, "no_vat": 28.7}, {"year": 2026, "with_vat": 36.47, "no_vat": 30.14}]}}, "CONFIRMACION DE RAZA (PERROS RAZAS ESPAÑOLAS)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 4.79, "no_vat": 3.96}, {"year": 2019, "with_vat": 5.0, "no_vat": 4.13}, {"year": 2020, "with_vat": 5.09, "no_vat": 4.21}, {"year": 2022, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2023, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2024, "with_vat": 5.6, "no_vat": 4.63}, {"year": 2025, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2026, "with_vat": 5.9, "no_vat": 4.88}]}}, "CONFIRMACION DE RAZA (perros de razas españolas)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 4.8, "no_vat": 3.97}], "User": [{"year": 2018, "with_vat": 4.8, "no_vat": 3.97}], "Canine_Collaborator": []}}, "CONFIRMACIÓN DE RAZA (PERROS GRUPOS ÉTINICOS)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 5.9, "no_vat": 4.88}]}}, "CONFIRMACIÓN DE RAZA (perros de razas españolas)": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2019, "with_vat": 5.0, "no_vat": 4.13}, {"year": 2020, "with_vat": 5.1, "no_vat": 4.21}, {"year": 2021, "with_vat": 5.1, "no_vat": 4.21}, {"year": 2022, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2023, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2024, "with_vat": 5.6, "no_vat": 4.63}, {"year": 2025, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2026, "with_vat": 5.9, "no_vat": 4.88}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 5.1, "no_vat": 4.21}, {"year": 2021, "with_vat": 5.1, "no_vat": 4.21}, {"year": 2022, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2023, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2024, "with_vat": 5.6, "no_vat": 4.63}, {"year": 2025, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2021, "with_vat": 5.19, "no_vat": 4.29}]}}, "CUOTA ABONADO": {"category": "SOCIOS Y ABONADOS", "prices": {"Member": [{"year": 2022, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2023, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2024, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2025, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2026, "with_vat": 105.0, "no_vat": 86.78}], "User": [{"year": 2022, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2023, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2024, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2025, "with_vat": 90.0, "no_vat": 74.38}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA ABONO JOVEN": {"category": "SOCIOS Y ABONADOS", "prices": {"Member": [{"year": 2024, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2025, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2026, "with_vat": 25.0, "no_vat": 20.66}], "User": [{"year": 2024, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2025, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA ANUAL CLUB AGILITY": {"category": "AGILITY", "prices": {"Member": [{"year": 2019, "with_vat": 73.5, "no_vat": 60.74}, {"year": 2020, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2021, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2022, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2023, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2024, "with_vat": 100.0, "no_vat": 82.64}, {"year": 2025, "with_vat": 105.0, "no_vat": 86.78}, {"year": 2026, "with_vat": 110.0, "no_vat": 90.91}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2021, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2022, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2023, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2024, "with_vat": 100.0, "no_vat": 82.64}, {"year": 2025, "with_vat": 105.0, "no_vat": 86.78}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA ANUAL CLUB DE AGILITY": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": null, "no_vat": null}], "User": [{"year": 2018, "with_vat": 71.8, "no_vat": 59.34}], "Canine_Collaborator": []}}, "CUOTA ANUAL GRUPO DE TRABAJO": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": null, "no_vat": null}, {"year": 2019, "with_vat": 73.5, "no_vat": 60.74}, {"year": 2020, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2021, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2022, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2023, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2024, "with_vat": 81.2, "no_vat": 67.11}, {"year": 2025, "with_vat": 85.3, "no_vat": 70.5}, {"year": 2026, "with_vat": 90.0, "no_vat": 74.38}], "User": [{"year": 2018, "with_vat": 71.8, "no_vat": 59.34}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2021, "with_vat": 75.0, "no_vat": 61.98}, {"year": 2022, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2023, "with_vat": 78.8, "no_vat": 65.12}, {"year": 2024, "with_vat": 81.2, "no_vat": 67.11}, {"year": 2025, "with_vat": 85.3, "no_vat": 70.5}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA ANUAL MULTIDISCIPLINAR": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": null, "no_vat": null}, {"year": 2019, "with_vat": 136.4, "no_vat": 112.73}, {"year": 2020, "with_vat": 138.7, "no_vat": 114.63}, {"year": 2021, "with_vat": 138.7, "no_vat": 114.63}, {"year": 2022, "with_vat": 145.6, "no_vat": 120.33}, {"year": 2023, "with_vat": 145.6, "no_vat": 120.33}, {"year": 2024, "with_vat": 150.1, "no_vat": 124.05}, {"year": 2025, "with_vat": 157.6, "no_vat": 130.25}, {"year": 2025, "with_vat": 157.6, "no_vat": 130.25}, {"year": 2026, "with_vat": 165.0, "no_vat": 136.36}, {"year": 2026, "with_vat": 165.0, "no_vat": 136.36}], "User": [{"year": 2018, "with_vat": 133.3, "no_vat": 110.17}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 138.7, "no_vat": 114.63}, {"year": 2021, "with_vat": 138.7, "no_vat": 114.63}, {"year": 2022, "with_vat": 145.6, "no_vat": 120.33}, {"year": 2023, "with_vat": 145.6, "no_vat": 120.33}, {"year": 2024, "with_vat": 150.1, "no_vat": 124.05}, {"year": 2025, "with_vat": 157.6, "no_vat": 130.25}, {"year": 2025, "with_vat": 157.6, "no_vat": 130.25}, {"year": 2026, "with_vat": null, "no_vat": null}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA ANUAL SOCIO": {"category": "SOCIOS Y ABONADOS", "prices": {"Member": [{"year": 2018, "with_vat": 30.0, "no_vat": 24.79}, {"year": 2019, "with_vat": 35.0, "no_vat": 28.93}, {"year": 2020, "with_vat": 35.0, "no_vat": 28.93}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2023, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2024, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2025, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2026, "with_vat": 45.0, "no_vat": 37.19}], "User": [{"year": 2018, "with_vat": null, "no_vat": null}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 35.0, "no_vat": 28.93}, {"year": 2021, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2022, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2023, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2024, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2025, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA INICIAL SOCIO (incluida la cuota del primer año natural)": {"category": "SOCIOS Y ABONADOS", "prices": {"Member": [{"year": 2019, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2020, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2021, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2022, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2023, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2024, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2025, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2026, "with_vat": 300.0, "no_vat": 247.93}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2021, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2022, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2023, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2024, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2025, "with_vat": 295.0, "no_vat": 243.8}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "CUOTA INICIAL SOCIO, incluida la cuota del primer año natural (260 + 30)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 290.0, "no_vat": 239.67}], "User": [{"year": 2018, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "DUPLICADO AFIJO": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 4.5, "no_vat": 3.72}], "User": [{"year": 2018, "with_vat": 5.2, "no_vat": 4.3}], "Canine_Collaborator": []}}, "DUPLICADO DE AFIJO": {"category": "AFIJOS", "prices": {"Member": [{"year": 2019, "with_vat": 4.7, "no_vat": 3.88}, {"year": 2020, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2021, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2022, "with_vat": 5.0, "no_vat": 4.13}, {"year": 2023, "with_vat": 5.0, "no_vat": 4.13}, {"year": 2024, "with_vat": 5.2, "no_vat": 4.3}, {"year": 2025, "with_vat": 5.5, "no_vat": 4.55}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2019, "with_vat": 5.4, "no_vat": 4.46}, {"year": 2020, "with_vat": 5.6, "no_vat": 4.63}, {"year": 2021, "with_vat": 5.6, "no_vat": 4.63}, {"year": 2022, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2023, "with_vat": 5.9, "no_vat": 4.88}, {"year": 2024, "with_vat": 6.1, "no_vat": 5.04}, {"year": 2025, "with_vat": 6.4, "no_vat": 5.29}, {"year": 2026, "with_vat": 30.0, "no_vat": 24.79}], "Canine_Collaborator": [{"year": 2018, "with_vat": 4.56, "no_vat": 3.77}, {"year": 2019, "with_vat": 4.74, "no_vat": 3.92}, {"year": 2020, "with_vat": 4.92, "no_vat": 4.07}, {"year": 2021, "with_vat": 4.92, "no_vat": 4.07}, {"year": 2022, "with_vat": 5.19, "no_vat": 4.29}, {"year": 2023, "with_vat": 5.19, "no_vat": 4.29}, {"year": 2024, "with_vat": 5.37, "no_vat": 4.44}, {"year": 2025, "with_vat": 5.64, "no_vat": 4.66}, {"year": 2026, "with_vat": 14.52, "no_vat": 12.0}]}}, "DUPLICADO DE PEDIGREE CUATRO GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 30.0, "no_vat": 24.79}, {"year": 2019, "with_vat": 30.7, "no_vat": 25.37}, {"year": 2020, "with_vat": 31.4, "no_vat": 25.95}, {"year": 2021, "with_vat": 31.4, "no_vat": 25.95}, {"year": 2022, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2023, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2024, "with_vat": 34.0, "no_vat": 28.1}], "User": [{"year": 2018, "with_vat": 41.3, "no_vat": 34.13}, {"year": 2019, "with_vat": 42.3, "no_vat": 34.96}, {"year": 2020, "with_vat": 43.2, "no_vat": 35.7}, {"year": 2021, "with_vat": 43.2, "no_vat": 35.7}, {"year": 2022, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2023, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2024, "with_vat": 46.8, "no_vat": 38.68}], "Canine_Collaborator": [{"year": 2018, "with_vat": 22.23, "no_vat": 18.37}, {"year": 2019, "with_vat": 22.76, "no_vat": 18.81}, {"year": 2020, "with_vat": 23.24, "no_vat": 19.21}, {"year": 2021, "with_vat": 23.24, "no_vat": 19.21}, {"year": 2022, "with_vat": 24.43, "no_vat": 20.19}, {"year": 2023, "with_vat": 24.43, "no_vat": 20.19}, {"year": 2024, "with_vat": 25.18, "no_vat": 20.81}]}}, "DUPLICADO DE PEDIGREE TRES GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2019, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2020, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2021, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2022, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2023, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2024, "with_vat": 29.9, "no_vat": 24.71}], "User": [{"year": 2018, "with_vat": 36.3, "no_vat": 30.0}, {"year": 2019, "with_vat": 37.2, "no_vat": 30.74}, {"year": 2020, "with_vat": 38.0, "no_vat": 31.4}, {"year": 2021, "with_vat": 38.0, "no_vat": 31.4}, {"year": 2022, "with_vat": 39.9, "no_vat": 32.98}, {"year": 2023, "with_vat": 39.9, "no_vat": 32.98}, {"year": 2024, "with_vat": 41.1, "no_vat": 33.97}], "Canine_Collaborator": [{"year": 2018, "with_vat": 20.99, "no_vat": 17.35}, {"year": 2019, "with_vat": 21.51, "no_vat": 17.78}, {"year": 2020, "with_vat": 21.97, "no_vat": 18.16}, {"year": 2021, "with_vat": 21.98, "no_vat": 18.16}, {"year": 2022, "with_vat": 23.07, "no_vat": 19.07}, {"year": 2023, "with_vat": 23.07, "no_vat": 19.07}, {"year": 2024, "with_vat": 23.76, "no_vat": 19.64}]}}, "DUPLICADO EXPORT PEDIGREE CUATRO GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2019, "with_vat": 57.1, "no_vat": 47.19}, {"year": 2020, "with_vat": 58.3, "no_vat": 48.18}, {"year": 2021, "with_vat": 58.3, "no_vat": 48.18}, {"year": 2022, "with_vat": 61.2, "no_vat": 50.58}, {"year": 2023, "with_vat": 61.2, "no_vat": 50.58}, {"year": 2024, "with_vat": 63.1, "no_vat": 52.15}], "User": [{"year": 2018, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 58.3, "no_vat": 48.18}, {"year": 2021, "with_vat": 58.3, "no_vat": 48.18}, {"year": 2022, "with_vat": 61.2, "no_vat": 50.58}, {"year": 2023, "with_vat": 61.2, "no_vat": 50.58}, {"year": 2024, "with_vat": 63.1, "no_vat": 52.15}], "Canine_Collaborator": [{"year": 2018, "with_vat": 37.24, "no_vat": 30.78}, {"year": 2019, "with_vat": 38.12, "no_vat": 31.5}, {"year": 2020, "with_vat": 38.91, "no_vat": 32.16}, {"year": 2021, "with_vat": 38.9, "no_vat": 32.15}, {"year": 2022, "with_vat": 40.85, "no_vat": 33.76}, {"year": 2023, "with_vat": 40.85, "no_vat": 33.76}, {"year": 2024, "with_vat": 42.12, "no_vat": 34.81}]}}, "DUPLICADO EXPORT PEDIGREE TRES GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 50.9, "no_vat": 42.07}, {"year": 2019, "with_vat": 52.1, "no_vat": 43.06}, {"year": 2020, "with_vat": 53.2, "no_vat": 43.97}, {"year": 2021, "with_vat": 53.2, "no_vat": 43.97}, {"year": 2022, "with_vat": 55.9, "no_vat": 46.2}, {"year": 2023, "with_vat": 55.9, "no_vat": 46.2}, {"year": 2024, "with_vat": 57.6, "no_vat": 47.6}], "User": [{"year": 2018, "with_vat": 50.9, "no_vat": 42.07}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 53.2, "no_vat": 43.97}, {"year": 2021, "with_vat": 53.2, "no_vat": 43.97}, {"year": 2022, "with_vat": 55.9, "no_vat": 46.2}, {"year": 2023, "with_vat": 55.9, "no_vat": 46.2}, {"year": 2024, "with_vat": 57.6, "no_vat": 47.6}], "Canine_Collaborator": [{"year": 2018, "with_vat": 34.28, "no_vat": 28.33}, {"year": 2019, "with_vat": 35.09, "no_vat": 29.0}, {"year": 2020, "with_vat": 35.83, "no_vat": 29.61}, {"year": 2021, "with_vat": 35.82, "no_vat": 29.6}, {"year": 2022, "with_vat": 37.64, "no_vat": 31.11}, {"year": 2023, "with_vat": 37.64, "no_vat": 31.11}, {"year": 2024, "with_vat": 38.79, "no_vat": 32.06}]}}, "DUPLICADO JUSTIFICANTE DE INSCRIPCIÓN": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2018, "with_vat": 4.3, "no_vat": 3.55}, {"year": 2019, "with_vat": 4.4, "no_vat": 3.64}, {"year": 2020, "with_vat": 4.5, "no_vat": 3.72}, {"year": 2021, "with_vat": 4.5, "no_vat": 3.72}, {"year": 2022, "with_vat": 4.7, "no_vat": 3.88}, {"year": 2023, "with_vat": 4.7, "no_vat": 3.88}, {"year": 2024, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2025, "with_vat": 5.0, "no_vat": 4.13}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2018, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2019, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2020, "with_vat": 6.2, "no_vat": 5.12}, {"year": 2021, "with_vat": 6.2, "no_vat": 5.12}, {"year": 2022, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2023, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2024, "with_vat": 6.7, "no_vat": 5.54}, {"year": 2025, "with_vat": 7.0, "no_vat": 5.79}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": [{"year": 2018, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2019, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2020, "with_vat": 6.21, "no_vat": 5.13}, {"year": 2021, "with_vat": 6.2, "no_vat": 5.13}, {"year": 2022, "with_vat": 6.51, "no_vat": 5.38}, {"year": 2023, "with_vat": 6.51, "no_vat": 5.38}, {"year": 2024, "with_vat": 6.72, "no_vat": 5.55}, {"year": 2025, "with_vat": 7.02, "no_vat": 5.8}, {"year": 2026, "with_vat": 8.47, "no_vat": 7.0}]}}, "DUPLICADO PEDIGREE_RSCE": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 17.54, "no_vat": 14.5}]}}, "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN 3 GENERACIONES": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 43.81, "no_vat": 36.21}]}}, "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN 4 GENERACIONES": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 46.73, "no_vat": 38.62}]}}, "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN CUATRO GENERACIONES EN PAPEL": {"category": "DUPLICADOS DE PEDIGRÍES", "prices": {"Member": [{"year": 2025, "with_vat": 66.3, "no_vat": 54.79}, {"year": 2026, "with_vat": 70.0, "no_vat": 57.85}], "User": [{"year": 2025, "with_vat": 66.3, "no_vat": 54.79}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2025, "with_vat": 44.26, "no_vat": 36.58}]}}, "DUPLICADO PEDIGREE_RSCE DE EXPORTACIÓN TRES GENERACIONES EN PAPEL": {"category": "DUPLICADOS DE PEDIGRÍES", "prices": {"Member": [{"year": 2025, "with_vat": 60.5, "no_vat": 50.0}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "User": [{"year": 2025, "with_vat": 60.5, "no_vat": 50.0}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2025, "with_vat": 40.78, "no_vat": 33.7}]}}, "DUPLICADO PEDIGREE_RSCE EN PAPEL": {"category": "DUPLICADOS DE PEDIGRÍES", "prices": {"Member": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2025, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2026, "with_vat": 27.0, "no_vat": 22.31}], "Canine_Collaborator": [{"year": 2025, "with_vat": 29.63, "no_vat": 24.49}]}}, "EDIGREE_RSCE PREMIUM TRES GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 32.0, "no_vat": 26.45}], "User": [{"year": 2026, "with_vat": 43.0, "no_vat": 35.54}], "Canine_Collaborator": []}}, "EXCEPCION, RECARGO 300% (más de 12 meses y menos de 18)  Art, 5.k) Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 29.85, "no_vat": 24.67}, {"year": 2020, "with_vat": 31.18, "no_vat": 25.77}, {"year": 2022, "with_vat": 32.74, "no_vat": 27.06}, {"year": 2023, "with_vat": 32.74, "no_vat": 27.06}, {"year": 2024, "with_vat": 33.75, "no_vat": 27.89}]}}, "EXCEPCION: RECARGO 300% (más de 12 meses y menos de 18 meses)  Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2019, "with_vat": 30.54, "no_vat": 25.24}]}}, "EXPEDICIÓN DE INFORMES (precio por folio impreso)": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2021, "with_vat": 11.1, "no_vat": 9.17}, {"year": 2022, "with_vat": 11.7, "no_vat": 9.67}, {"year": 2023, "with_vat": 11.7, "no_vat": 9.67}, {"year": 2024, "with_vat": 12.1, "no_vat": 10.0}, {"year": 2025, "with_vat": 12.7, "no_vat": 10.5}, {"year": 2026, "with_vat": 0.0, "no_vat": 0.0}], "User": [{"year": 2021, "with_vat": 11.1, "no_vat": 9.17}, {"year": 2022, "with_vat": 11.7, "no_vat": 9.67}, {"year": 2023, "with_vat": 11.7, "no_vat": 9.67}, {"year": 2024, "with_vat": 12.1, "no_vat": 10.0}, {"year": 2025, "with_vat": 12.7, "no_vat": 10.5}, {"year": 2026, "with_vat": 13.0, "no_vat": 10.74}], "Canine_Collaborator": []}}, "EXPEDICIÓN TITULOS DE CAMPEONES": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2018, "with_vat": 5.1, "no_vat": 4.21}, {"year": 2019, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2020, "with_vat": 5.5, "no_vat": 4.55}], "User": [{"year": 2018, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2019, "with_vat": 5.5, "no_vat": 4.55}, {"year": 2020, "with_vat": 5.7, "no_vat": 4.71}], "Canine_Collaborator": []}}, "EXPEDICIÓN TITULOS LATIN CHAMPION": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2020, "with_vat": 15.0, "no_vat": 12.4}], "User": [{"year": 2020, "with_vat": 20.0, "no_vat": 16.53}], "Canine_Collaborator": []}}, "EXPEDICIÓN TÍTULOS \"LATIN CHAMPION\"": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2021, "with_vat": 15.0, "no_vat": 12.4}, {"year": 2022, "with_vat": 15.8, "no_vat": 13.06}, {"year": 2023, "with_vat": 15.8, "no_vat": 13.06}, {"year": 2024, "with_vat": 16.3, "no_vat": 13.47}, {"year": 2025, "with_vat": 17.1, "no_vat": 14.13}], "User": [{"year": 2021, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2022, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2023, "with_vat": 21.0, "no_vat": 17.36}, {"year": 2024, "with_vat": 21.7, "no_vat": 17.93}, {"year": 2025, "with_vat": 22.8, "no_vat": 18.84}], "Canine_Collaborator": []}}, "EXPEDICIÓN TÍTULOS DE CAMPEONES": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2021, "with_vat": 5.5, "no_vat": 4.55}, {"year": 2022, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2023, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2024, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2025, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2021, "with_vat": 5.7, "no_vat": 4.71}, {"year": 2022, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2023, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2024, "with_vat": 6.2, "no_vat": 5.12}, {"year": 2025, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "EXPEDICIÓN TÍTULOS “LATIN CHAMPION”": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 17.1, "no_vat": 14.13}], "User": [{"year": 2026, "with_vat": 22.8, "no_vat": 18.84}], "Canine_Collaborator": []}}, "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento)": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2019, "with_vat": 2.6, "no_vat": 2.15}, {"year": 2020, "with_vat": 2.7, "no_vat": 2.23}, {"year": 2021, "with_vat": 2.7, "no_vat": 2.23}, {"year": 2022, "with_vat": 2.8, "no_vat": 2.31}, {"year": 2023, "with_vat": 2.8, "no_vat": 2.31}, {"year": 2024, "with_vat": 2.9, "no_vat": 2.4}, {"year": 2025, "with_vat": 3.0, "no_vat": 2.48}], "User": [{"year": 2019, "with_vat": 4.2, "no_vat": 3.47}, {"year": 2020, "with_vat": 4.3, "no_vat": 3.55}, {"year": 2021, "with_vat": 4.3, "no_vat": 3.55}, {"year": 2022, "with_vat": 4.5, "no_vat": 3.72}, {"year": 2023, "with_vat": 4.5, "no_vat": 3.72}, {"year": 2024, "with_vat": 4.6, "no_vat": 3.8}, {"year": 2025, "with_vat": 4.8, "no_vat": 3.97}], "Canine_Collaborator": [{"year": 2021, "with_vat": 2.32, "no_vat": 1.91}]}}, "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento) EN ORIGEN": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": null, "no_vat": null}], "User": [{"year": 2018, "with_vat": 4.1, "no_vat": 3.39}], "Canine_Collaborator": [{"year": 2018, "with_vat": 2.21, "no_vat": 1.83}, {"year": 2019, "with_vat": 2.26, "no_vat": 1.87}, {"year": 2020, "with_vat": 2.31, "no_vat": 1.91}, {"year": 2022, "with_vat": 2.4, "no_vat": 1.98}, {"year": 2023, "with_vat": 2.4, "no_vat": 1.98}, {"year": 2024, "with_vat": 2.4, "no_vat": 1.98}]}}, "FORMULARIO IDENTIFICACIÓN (por cachorro al notificar nacimiento) RSCE": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 2.5, "no_vat": 2.07}], "User": [{"year": 2018, "with_vat": 4.1, "no_vat": 3.39}], "Canine_Collaborator": [{"year": 2018, "with_vat": 3.06, "no_vat": 2.53}, {"year": 2019, "with_vat": 3.13, "no_vat": 2.59}, {"year": 2020, "with_vat": 3.21, "no_vat": 2.65}, {"year": 2021, "with_vat": 3.21, "no_vat": 2.65}, {"year": 2022, "with_vat": 3.33, "no_vat": 2.75}, {"year": 2023, "with_vat": 3.33, "no_vat": 2.75}, {"year": 2024, "with_vat": 3.45, "no_vat": 2.85}]}}, "IMPRESO ALTA CAMADA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 1.9, "no_vat": 1.57}], "User": [{"year": 2018, "with_vat": 1.9, "no_vat": 1.57}], "Canine_Collaborator": []}}, "IMPRESO ALTA DE CAMADA": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2019, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2020, "with_vat": 2.1, "no_vat": 1.74}, {"year": 2021, "with_vat": 2.1, "no_vat": 1.74}, {"year": 2022, "with_vat": 2.2, "no_vat": 1.82}, {"year": 2023, "with_vat": 2.2, "no_vat": 1.82}, {"year": 2024, "with_vat": 2.3, "no_vat": 1.9}, {"year": 2025, "with_vat": 2.4, "no_vat": 1.98}, {"year": 2026, "with_vat": 3.0, "no_vat": 2.48}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 2.1, "no_vat": 1.74}, {"year": 2021, "with_vat": 2.1, "no_vat": 1.74}, {"year": 2022, "with_vat": 2.2, "no_vat": 1.82}, {"year": 2023, "with_vat": 2.2, "no_vat": 1.82}, {"year": 2024, "with_vat": 2.3, "no_vat": 1.9}, {"year": 2025, "with_vat": 2.4, "no_vat": 1.98}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2018, "with_vat": 1.21, "no_vat": 1.0}, {"year": 2019, "with_vat": 1.27, "no_vat": 1.05}, {"year": 2020, "with_vat": 1.33, "no_vat": 1.1}, {"year": 2021, "with_vat": 1.33, "no_vat": 1.1}, {"year": 2022, "with_vat": 1.39, "no_vat": 1.15}, {"year": 2023, "with_vat": 1.39, "no_vat": 1.15}, {"year": 2024, "with_vat": 1.45, "no_vat": 1.2}, {"year": 2025, "with_vat": 1.51, "no_vat": 1.25}]}}, "IMPRESO ALTA DE CAMADA (GRUPOS ÉTNICOS)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 2.15, "no_vat": 1.78}]}}, "IMPRESO ALTA DE CAMADA (PERROS DE RAZA)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 2.15, "no_vat": 1.78}]}}, "IMPRESO DIPLOMA DE INICIACIÓN (original y 3 copias)": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": 1.7, "no_vat": 1.4}, {"year": 2020, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "User": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "Canine_Collaborator": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": 1.69, "no_vat": 1.4}, {"year": 2020, "with_vat": 1.79, "no_vat": 1.48}, {"year": 2021, "with_vat": 1.79, "no_vat": 1.48}, {"year": 2022, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2023, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2024, "with_vat": 1.98, "no_vat": 1.64}, {"year": 2025, "with_vat": 2.08, "no_vat": 1.72}, {"year": 2026, "with_vat": 3.0, "no_vat": 2.48}]}}, "IMPRESO HOJA DE CALIFICACIÓN – PRUEBAS PERRO DE AGUA ESPAÑOL": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "User": [{"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "Canine_Collaborator": [{"year": 2022, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2023, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2024, "with_vat": 1.92, "no_vat": 1.59}, {"year": 2025, "with_vat": 2.02, "no_vat": 1.67}, {"year": 2026, "with_vat": 3.0, "no_vat": 2.48}]}}, "IMPRESO PARA PRUEBAS CACT - CACIT": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 1.1, "no_vat": 0.91}], "User": [{"year": 2018, "with_vat": 1.1, "no_vat": 0.91}], "Canine_Collaborator": []}}, "IMPRESO PARA PRUEBAS DE C.A.C.T. - C.A.C.I.T.": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 1.2, "no_vat": 0.99}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2018, "with_vat": 1.1, "no_vat": 0.91}, {"year": 2019, "with_vat": 1.2, "no_vat": 0.99}, {"year": 2020, "with_vat": 1.29, "no_vat": 1.07}, {"year": 2021, "with_vat": 1.29, "no_vat": 1.07}, {"year": 2022, "with_vat": 1.39, "no_vat": 1.15}, {"year": 2023, "with_vat": 1.39, "no_vat": 1.15}, {"year": 2024, "with_vat": 1.96, "no_vat": 1.62}, {"year": 2025, "with_vat": 2.08, "no_vat": 1.72}, {"year": 2026, "with_vat": 3.0, "no_vat": 2.48}]}}, "IMPRESO PARA PRUEBAS DE CACT – CACIT": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2020, "with_vat": 1.3, "no_vat": 1.07}, {"year": 2021, "with_vat": 1.3, "no_vat": 1.07}, {"year": 2022, "with_vat": 1.4, "no_vat": 1.16}, {"year": 2023, "with_vat": 1.4, "no_vat": 1.16}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "User": [{"year": 2020, "with_vat": 1.3, "no_vat": 1.07}, {"year": 2021, "with_vat": 1.3, "no_vat": 1.07}, {"year": 2022, "with_vat": 1.4, "no_vat": 1.16}, {"year": 2023, "with_vat": 1.4, "no_vat": 1.16}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "Canine_Collaborator": []}}, "IMPRESO PRUEBAS APTITUDES NATURALES (original y 3 copias)": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": 1.7, "no_vat": 1.4}, {"year": 2020, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "User": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2021, "with_vat": 1.8, "no_vat": 1.49}, {"year": 2022, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2023, "with_vat": 1.9, "no_vat": 1.57}, {"year": 2024, "with_vat": 2.0, "no_vat": 1.65}, {"year": 2025, "with_vat": 2.1, "no_vat": 1.74}], "Canine_Collaborator": [{"year": 2018, "with_vat": 1.6, "no_vat": 1.32}, {"year": 2019, "with_vat": 1.69, "no_vat": 1.4}, {"year": 2020, "with_vat": 1.79, "no_vat": 1.48}, {"year": 2021, "with_vat": 1.79, "no_vat": 1.48}, {"year": 2022, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2023, "with_vat": 1.89, "no_vat": 1.56}, {"year": 2024, "with_vat": 1.98, "no_vat": 1.64}, {"year": 2025, "with_vat": 2.08, "no_vat": 1.72}, {"year": 2026, "with_vat": 3.0, "no_vat": 2.48}]}}, "INFORMES SOLICITADOS A LA R.S.C.E. (precio por folio impreso)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 10.8, "no_vat": 8.93}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "INFORMES SOLICITADOS A LA RSCE (precio por folio impreso": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 10.5, "no_vat": 8.68}], "User": [{"year": 2018, "with_vat": 10.5, "no_vat": 8.68}], "Canine_Collaborator": []}}, "INFORMES SOLICITADOS A LA RSCE (precio por folio impreso)": {"category": "IMPRESOS", "prices": {"Member": [{"year": 2020, "with_vat": 11.1, "no_vat": 9.17}], "User": [{"year": 2020, "with_vat": 11.1, "no_vat": 9.17}], "Canine_Collaborator": []}}, "INSCRIPCION DE OTROS LIBROS GENEALOGICOS (NO ADMITE PLUS DE URGENCIA)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 32.82, "no_vat": 27.12}, {"year": 2019, "with_vat": 33.57, "no_vat": 27.74}, {"year": 2020, "with_vat": 34.27, "no_vat": 28.32}, {"year": 2022, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2023, "with_vat": 24.22, "no_vat": 20.02}, {"year": 2024, "with_vat": 24.99, "no_vat": 20.65}, {"year": 2025, "with_vat": 26.27, "no_vat": 21.71}, {"year": 2026, "with_vat": 28.48, "no_vat": 23.54}]}}, "INSCRIPCION PERROS DE OTROS LIBROS GENEALÓGICOS (ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 56.1, "no_vat": 46.36}], "User": [{"year": 2019, "with_vat": 57.8, "no_vat": 47.77}], "Canine_Collaborator": [{"year": 2021, "with_vat": 34.26, "no_vat": 28.32}]}}, "INSCRIPCIÓN CACHORRO": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2018, "with_vat": 14.2, "no_vat": 11.74}, {"year": 2019, "with_vat": 14.6, "no_vat": 12.07}, {"year": 2020, "with_vat": 14.9, "no_vat": 12.31}, {"year": 2021, "with_vat": 14.9, "no_vat": 12.31}, {"year": 2022, "with_vat": 15.6, "no_vat": 12.89}, {"year": 2023, "with_vat": 15.6, "no_vat": 12.89}, {"year": 2024, "with_vat": 16.1, "no_vat": 13.31}], "User": [{"year": 2018, "with_vat": 17.2, "no_vat": 14.21}, {"year": 2019, "with_vat": 17.6, "no_vat": 14.55}, {"year": 2020, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2021, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2022, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2023, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2024, "with_vat": 19.5, "no_vat": 16.12}], "Canine_Collaborator": [{"year": 2018, "with_vat": 9.96, "no_vat": 8.23}, {"year": 2019, "with_vat": 10.19, "no_vat": 8.42}, {"year": 2020, "with_vat": 10.42, "no_vat": 8.61}, {"year": 2021, "with_vat": 10.42, "no_vat": 8.61}, {"year": 2022, "with_vat": 10.9, "no_vat": 9.01}, {"year": 2023, "with_vat": 10.9, "no_vat": 9.01}, {"year": 2024, "with_vat": 11.25, "no_vat": 9.3}]}}, "INSCRIPCIÓN CACHORRO ACCESS LBO/RBR": {"category": "INSCRIPCIONES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 11.0, "no_vat": 9.09}, {"year": 2026, "with_vat": 12.0, "no_vat": 9.92}], "User": [{"year": 2025, "with_vat": 15.0, "no_vat": 12.4}, {"year": 2026, "with_vat": 17.0, "no_vat": 14.05}], "Canine_Collaborator": []}}, "INSCRIPCIÓN CACHORRO DE RAZAS ESPAÑOLAS BONIFICADAS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2024, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2025, "with_vat": 8.8, "no_vat": 7.27}, {"year": 2026, "with_vat": 8.8, "no_vat": 7.27}], "User": [{"year": 2024, "with_vat": 9.7, "no_vat": 8.02}, {"year": 2025, "with_vat": 10.2, "no_vat": 8.43}, {"year": 2026, "with_vat": 10.2, "no_vat": 8.43}], "Canine_Collaborator": []}}, "INSCRIPCIÓN CACHORRO EN EL REGISTRO DE GRUPOS ÉTNICOS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2024, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2025, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2026, "with_vat": 8.4, "no_vat": 6.94}], "User": [{"year": 2024, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2025, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2024, "with_vat": 5.6, "no_vat": 4.63}]}}, "INSCRIPCIÓN CACHORRO PREMIUM LOE/RRC": {"category": "INSCRIPCIONES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 16.9, "no_vat": 13.97}, {"year": 2026, "with_vat": 18.0, "no_vat": 14.88}], "User": [{"year": 2025, "with_vat": 20.5, "no_vat": 16.94}, {"year": 2026, "with_vat": 22.0, "no_vat": 18.18}], "Canine_Collaborator": []}}, "INSCRIPCIÓN CACHORRO RAZAS ESPAÑOLAS BONIFICADAS ( Del 01/05/2023 al 30/04/2024)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2023, "with_vat": 5.46, "no_vat": 4.51}]}}, "INSCRIPCIÓN CACHORRO RAZAS ESPAÑOLAS BONIFICADAS ( Del 01/05/2024 al 30/04/2025)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2024, "with_vat": 5.6, "no_vat": 4.63}]}}, "INSCRIPCIÓN IMPORTADOS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2018, "with_vat": 25.8, "no_vat": 21.32}, {"year": 2019, "with_vat": 26.4, "no_vat": 21.82}, {"year": 2020, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2021, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2022, "with_vat": 28.4, "no_vat": 23.47}, {"year": 2023, "with_vat": 28.4, "no_vat": 23.47}, {"year": 2024, "with_vat": 29.3, "no_vat": 24.21}, {"year": 2025, "with_vat": 30.8, "no_vat": 25.45}, {"year": 2026, "with_vat": 33.0, "no_vat": 27.27}], "User": [{"year": 2018, "with_vat": 38.0, "no_vat": 31.4}, {"year": 2019, "with_vat": 38.9, "no_vat": 32.15}, {"year": 2020, "with_vat": 39.7, "no_vat": 32.81}, {"year": 2021, "with_vat": 39.7, "no_vat": 32.81}, {"year": 2022, "with_vat": 41.7, "no_vat": 34.46}, {"year": 2023, "with_vat": 41.7, "no_vat": 34.46}, {"year": 2024, "with_vat": 43.0, "no_vat": 35.54}, {"year": 2025, "with_vat": 45.2, "no_vat": 37.36}, {"year": 2026, "with_vat": 49.0, "no_vat": 40.5}], "Canine_Collaborator": [{"year": 2018, "with_vat": 22.05, "no_vat": 18.22}, {"year": 2019, "with_vat": 22.57, "no_vat": 18.65}, {"year": 2020, "with_vat": 23.03, "no_vat": 19.03}, {"year": 2021, "with_vat": 23.03, "no_vat": 19.03}, {"year": 2022, "with_vat": 24.22, "no_vat": 20.02}, {"year": 2023, "with_vat": 24.22, "no_vat": 20.02}, {"year": 2024, "with_vat": 24.99, "no_vat": 20.65}, {"year": 2025, "with_vat": 26.27, "no_vat": 21.71}, {"year": 2026, "with_vat": 28.48, "no_vat": 23.54}]}}, "INSCRIPCIÓN PERRO con más de 18 meses aportando prueba ADN progenitores y huella genética de compatibilidad": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2021, "with_vat": 89.4, "no_vat": 73.88}, {"year": 2022, "with_vat": 93.9, "no_vat": 77.6}, {"year": 2023, "with_vat": 93.9, "no_vat": 77.6}, {"year": 2024, "with_vat": 96.8, "no_vat": 80.0}, {"year": 2025, "with_vat": 101.6, "no_vat": 83.97}, {"year": 2026, "with_vat": 125.0, "no_vat": 103.31}], "User": [{"year": 2021, "with_vat": 108.0, "no_vat": 89.26}, {"year": 2022, "with_vat": 113.4, "no_vat": 93.72}, {"year": 2023, "with_vat": 113.4, "no_vat": 93.72}, {"year": 2024, "with_vat": 116.9, "no_vat": 96.61}, {"year": 2025, "with_vat": 122.7, "no_vat": 101.4}, {"year": 2026, "with_vat": 151.0, "no_vat": 124.79}], "Canine_Collaborator": []}}, "INSCRIPCIÓN PERROS DE OTROS LIBROS GENEALÓGICOS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2020, "with_vat": 57.3, "no_vat": 47.36}, {"year": 2021, "with_vat": 57.3, "no_vat": 47.36}, {"year": 2022, "with_vat": 60.2, "no_vat": 49.75}], "User": [{"year": 2020, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2021, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2022, "with_vat": 62.0, "no_vat": 51.24}], "Canine_Collaborator": []}}, "INSCRIPCIÓN PERROS DE OTROS LIBROS GENEALÓGICOS - ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 54.8, "no_vat": 45.29}], "User": [{"year": 2018, "with_vat": 56.5, "no_vat": 46.69}], "Canine_Collaborator": []}}, "INSCRIPCIÓN PERROS PROCEDENTES DE OTROS LIBROS GENEALÓGICOS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2023, "with_vat": 28.4, "no_vat": 23.47}, {"year": 2024, "with_vat": 29.3, "no_vat": 24.21}, {"year": 2025, "with_vat": 30.8, "no_vat": 25.45}], "User": [{"year": 2023, "with_vat": 41.7, "no_vat": 34.46}, {"year": 2024, "with_vat": 43.0, "no_vat": 35.54}, {"year": 2025, "with_vat": 45.2, "no_vat": 37.36}], "Canine_Collaborator": []}}, "INSCRIPCIÓN PERROS PROCEDENTES DE OTROS LIBROS GENEALÓGICOS (ESTE TRÁMITE NO ADMITE PLUS DE URGENCIA)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 33.0, "no_vat": 27.27}], "User": [{"year": 2026, "with_vat": 49.0, "no_vat": 40.5}], "Canine_Collaborator": []}}, "LICENCIA INCIAL": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 16.1, "no_vat": 13.31}], "User": [{"year": 2019, "with_vat": 33.1, "no_vat": 27.36}], "Canine_Collaborator": []}}, "LICENCIA INICIAL": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2018, "with_vat": 15.7, "no_vat": 12.98}, {"year": 2018, "with_vat": 15.7, "no_vat": 12.98}, {"year": 2019, "with_vat": 16.1, "no_vat": 13.31}, {"year": 2020, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2020, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2021, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2021, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2022, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2022, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2023, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2023, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2024, "with_vat": 19.5, "no_vat": 16.12}, {"year": 2024, "with_vat": 19.5, "no_vat": 16.12}, {"year": 2025, "with_vat": 20.5, "no_vat": 16.94}, {"year": 2025, "with_vat": 20.5, "no_vat": 16.94}, {"year": 2026, "with_vat": 22.0, "no_vat": 18.18}, {"year": 2026, "with_vat": 22.0, "no_vat": 18.18}], "User": [{"year": 2018, "with_vat": 32.3, "no_vat": 26.69}, {"year": 2018, "with_vat": 32.3, "no_vat": 26.69}, {"year": 2019, "with_vat": 33.1, "no_vat": 27.36}, {"year": 2020, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2020, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2021, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2021, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2022, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2022, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2023, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2023, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2024, "with_vat": 39.0, "no_vat": 32.23}, {"year": 2024, "with_vat": 39.0, "no_vat": 32.23}, {"year": 2025, "with_vat": 41.0, "no_vat": 33.88}, {"year": 2025, "with_vat": 41.0, "no_vat": 33.88}, {"year": 2026, "with_vat": 44.0, "no_vat": 36.36}, {"year": 2026, "with_vat": 44.0, "no_vat": 36.36}], "Canine_Collaborator": []}}, "MEJORA A PEDIGREE_RSCE PLUS CUATRO GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 28.0, "no_vat": 23.14}], "User": [{"year": 2026, "with_vat": 38.0, "no_vat": 31.4}], "Canine_Collaborator": [{"year": 2026, "with_vat": 24.7, "no_vat": 20.41}]}}, "MEJORA A PEDIGREE_RSCE PLUS RRC (sin genealogía completa) CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 39.0, "no_vat": 32.23}]}}, "MEJORA A PEDIGREE_RSCE PLUS RRC (sin genealógica completa) SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2026, "with_vat": 27.0, "no_vat": 22.31}], "Canine_Collaborator": [{"year": 2026, "with_vat": 17.54, "no_vat": 14.5}]}}, "MEJORA A PEDIGREE_RSCE PLUS TRES GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2026, "with_vat": 33.0, "no_vat": 27.27}], "Canine_Collaborator": [{"year": 2026, "with_vat": 21.45, "no_vat": 17.73}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE CUATRO GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 27.0, "no_vat": 22.31}], "User": [{"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "Canine_Collaborator": [{"year": 2026, "with_vat": 22.75, "no_vat": 18.8}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "MEJORA DE ACCESS A PREMIUM", "prices": {"Member": [{"year": 2025, "with_vat": 25.0, "no_vat": 20.66}], "User": [{"year": 2025, "with_vat": 32.0, "no_vat": 26.45}], "Canine_Collaborator": [{"year": 2025, "with_vat": 20.8, "no_vat": 17.19}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE TRES GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 21.0, "no_vat": 17.36}], "User": [{"year": 2026, "with_vat": 29.0, "no_vat": 23.97}], "Canine_Collaborator": [{"year": 2026, "with_vat": 18.84, "no_vat": 15.57}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS LBO A PEDIGREE_RSCE PREMIUM LOE TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "MEJORA DE ACCESS A PREMIUM", "prices": {"Member": [{"year": 2025, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2025, "with_vat": 27.0, "no_vat": 22.31}], "Canine_Collaborator": [{"year": 2025, "with_vat": 17.54, "no_vat": 14.5}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS RBR A PEDIGREE_RSCE PREMIUM RRC SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 19.0, "no_vat": 15.7}], "User": [{"year": 2026, "with_vat": 27.0, "no_vat": 22.31}], "Canine_Collaborator": [{"year": 2026, "with_vat": 16.84, "no_vat": 13.92}]}}, "MEJORA DE PEDIGREE_RSCE ACCESS RBR A PEDIGREE_RSCE PREMIUM RRC SIN TRANSFERENCIA EN PAPEL": {"category": "MEJORA DE ACCESS A PREMIUM", "prices": {"Member": [{"year": 2025, "with_vat": 18.0, "no_vat": 14.88}], "User": [{"year": 2025, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": [{"year": 2025, "with_vat": 15.6, "no_vat": 12.89}]}}, "NOTIFICACIÓN NACIMIENTO CACHORRO (formulario de identificación)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 3.2, "no_vat": 2.64}], "User": [{"year": 2026, "with_vat": 5.1, "no_vat": 4.21}], "Canine_Collaborator": []}}, "OTRAS ANOTACIONES EN EL LOE / RRC / RGE": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2022, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2023, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2024, "with_vat": 6.5, "no_vat": 5.37}], "User": [{"year": 2022, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2023, "with_vat": 8.4, "no_vat": 6.94}, {"year": 2024, "with_vat": 8.7, "no_vat": 7.19}], "Canine_Collaborator": []}}, "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES  (SI CONSTAN EN EL L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 28.27, "no_vat": 23.36}, {"year": 2020, "with_vat": 29.55, "no_vat": 24.42}, {"year": 2022, "with_vat": 31.0, "no_vat": 25.62}, {"year": 2023, "with_vat": 31.0, "no_vat": 25.62}, {"year": 2024, "with_vat": 31.99, "no_vat": 26.44}]}}, "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES  (si consta en el L.O.E..)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2019, "with_vat": 28.97, "no_vat": 23.94}]}}, "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 33.1, "no_vat": 27.36}], "User": [{"year": 2018, "with_vat": 48.7, "no_vat": 40.25}], "Canine_Collaborator": [{"year": 2021, "with_vat": 29.55, "no_vat": 24.42}]}}, "PEDIGREE CON TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2019, "with_vat": 33.9, "no_vat": 28.02}, {"year": 2020, "with_vat": 34.6, "no_vat": 28.6}, {"year": 2021, "with_vat": 34.6, "no_vat": 28.6}, {"year": 2022, "with_vat": 36.3, "no_vat": 30.0}, {"year": 2023, "with_vat": 36.3, "no_vat": 30.0}, {"year": 2024, "with_vat": 37.4, "no_vat": 30.91}], "User": [{"year": 2019, "with_vat": 49.9, "no_vat": 41.24}, {"year": 2020, "with_vat": 50.9, "no_vat": 42.07}, {"year": 2021, "with_vat": 50.9, "no_vat": 42.07}, {"year": 2022, "with_vat": 53.4, "no_vat": 44.13}, {"year": 2023, "with_vat": 53.4, "no_vat": 44.13}, {"year": 2024, "with_vat": 55.1, "no_vat": 45.54}], "Canine_Collaborator": []}}, "PEDIGREE CON TRANSFERENCIA TRES GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 29.3, "no_vat": 24.21}, {"year": 2019, "with_vat": 30.0, "no_vat": 24.79}, {"year": 2020, "with_vat": 30.6, "no_vat": 25.29}, {"year": 2021, "with_vat": 30.6, "no_vat": 25.29}, {"year": 2022, "with_vat": 32.1, "no_vat": 26.53}, {"year": 2023, "with_vat": 32.1, "no_vat": 26.53}, {"year": 2024, "with_vat": 33.1, "no_vat": 27.36}], "User": [{"year": 2018, "with_vat": 43.2, "no_vat": 35.7}, {"year": 2019, "with_vat": 44.2, "no_vat": 36.53}, {"year": 2020, "with_vat": 45.1, "no_vat": 37.27}, {"year": 2021, "with_vat": 45.1, "no_vat": 37.27}, {"year": 2022, "with_vat": 47.4, "no_vat": 39.17}, {"year": 2023, "with_vat": 47.4, "no_vat": 39.17}, {"year": 2024, "with_vat": 48.9, "no_vat": 40.41}], "Canine_Collaborator": [{"year": 2018, "with_vat": 25.06, "no_vat": 20.71}, {"year": 2019, "with_vat": 25.64, "no_vat": 21.19}, {"year": 2020, "with_vat": 26.16, "no_vat": 21.62}, {"year": 2021, "with_vat": 26.16, "no_vat": 21.62}, {"year": 2022, "with_vat": 27.49, "no_vat": 22.72}, {"year": 2023, "with_vat": 27.49, "no_vat": 22.72}, {"year": 2024, "with_vat": 28.36, "no_vat": 23.44}]}}, "PEDIGREE DE EXPORTACION (4 generaciones)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 41.59, "no_vat": 34.37}, {"year": 2019, "with_vat": 42.59, "no_vat": 35.2}, {"year": 2020, "with_vat": 43.46, "no_vat": 35.92}, {"year": 2022, "with_vat": 45.6, "no_vat": 37.69}, {"year": 2023, "with_vat": 45.6, "no_vat": 37.69}, {"year": 2024, "with_vat": 47.01, "no_vat": 38.85}]}}, "PEDIGREE DE EXPORTACIÓN CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 62.1, "no_vat": 51.32}], "User": [{"year": 2018, "with_vat": 62.1, "no_vat": 51.32}], "Canine_Collaborator": [{"year": 2021, "with_vat": 43.45, "no_vat": 35.91}]}}, "PEDIGREE DE EXPORTACIÓN CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2019, "with_vat": 63.6, "no_vat": 52.56}, {"year": 2020, "with_vat": 64.9, "no_vat": 53.64}, {"year": 2021, "with_vat": 64.9, "no_vat": 53.64}, {"year": 2022, "with_vat": 68.1, "no_vat": 56.28}, {"year": 2023, "with_vat": 68.1, "no_vat": 56.28}, {"year": 2024, "with_vat": 70.2, "no_vat": 58.02}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 64.9, "no_vat": 53.64}, {"year": 2021, "with_vat": 64.9, "no_vat": 53.64}, {"year": 2022, "with_vat": 68.1, "no_vat": 56.28}, {"year": 2023, "with_vat": 68.1, "no_vat": 56.28}, {"year": 2024, "with_vat": 70.2, "no_vat": 58.02}], "Canine_Collaborator": []}}, "PEDIGREE DE EXPORTACIÓN TRES GENERACIONES": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 56.5, "no_vat": 46.69}, {"year": 2019, "with_vat": 57.8, "no_vat": 47.77}, {"year": 2020, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2021, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2022, "with_vat": 62.0, "no_vat": 51.24}, {"year": 2023, "with_vat": 62.0, "no_vat": 51.24}, {"year": 2024, "with_vat": 63.9, "no_vat": 52.81}], "User": [{"year": 2018, "with_vat": 56.5, "no_vat": 46.69}, {"year": 2019, "with_vat": null, "no_vat": null}, {"year": 2020, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2021, "with_vat": 59.0, "no_vat": 48.76}, {"year": 2022, "with_vat": 62.0, "no_vat": 51.24}, {"year": 2023, "with_vat": 62.0, "no_vat": 51.24}, {"year": 2024, "with_vat": 63.9, "no_vat": 52.81}], "Canine_Collaborator": [{"year": 2018, "with_vat": 38.05, "no_vat": 31.45}, {"year": 2019, "with_vat": 38.93, "no_vat": 32.17}, {"year": 2020, "with_vat": 39.74, "no_vat": 32.84}, {"year": 2021, "with_vat": 39.72, "no_vat": 32.82}, {"year": 2022, "with_vat": 41.76, "no_vat": 34.51}, {"year": 2023, "with_vat": 41.76, "no_vat": 34.51}, {"year": 2024, "with_vat": 43.04, "no_vat": 35.57}]}}, "PEDIGREE ORO": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2022, "with_vat": 33.38, "no_vat": 27.59}, {"year": 2023, "with_vat": 33.38, "no_vat": 27.59}]}}, "PEDIGREE ORO CON TRANSFERENCIA DE PROPIEDAD": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2021, "with_vat": 39.6, "no_vat": 32.73}], "User": [{"year": 2021, "with_vat": 58.4, "no_vat": 48.26}], "Canine_Collaborator": []}}, "PEDIGREE ORO SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2022, "with_vat": 28.37, "no_vat": 23.45}, {"year": 2023, "with_vat": 28.37, "no_vat": 23.45}]}}, "PEDIGREE PARA PERROS INSCRITOS EN EL R.R.C. (sin genealógica completa)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 17.0, "no_vat": 14.05}], "User": [{"year": 2018, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": [{"year": 2021, "with_vat": 15.21, "no_vat": 12.57}]}}, "PEDIGREE PARA PERROS INSCRITOS EN EL RRC (sin genealógia completa)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 17.4, "no_vat": 14.38}], "User": [{"year": 2019, "with_vat": 25.6, "no_vat": 21.16}], "Canine_Collaborator": [{"year": 2018, "with_vat": 14.51, "no_vat": 11.99}, {"year": 2019, "with_vat": 14.86, "no_vat": 12.28}, {"year": 2020, "with_vat": 15.21, "no_vat": 12.57}, {"year": 2022, "with_vat": 15.96, "no_vat": 13.19}, {"year": 2023, "with_vat": 15.96, "no_vat": 13.19}, {"year": 2024, "with_vat": 16.48, "no_vat": 13.62}]}}, "PEDIGREE PARA PERROS INSCRITOS EN EL RRC (sin genealógica completa)": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2020, "with_vat": 17.8, "no_vat": 14.71}, {"year": 2021, "with_vat": 17.8, "no_vat": 14.71}, {"year": 2022, "with_vat": 18.7, "no_vat": 15.45}, {"year": 2023, "with_vat": 18.7, "no_vat": 15.45}, {"year": 2024, "with_vat": 19.3, "no_vat": 15.95}], "User": [{"year": 2020, "with_vat": 26.2, "no_vat": 21.65}, {"year": 2021, "with_vat": 26.2, "no_vat": 21.65}, {"year": 2022, "with_vat": 27.5, "no_vat": 22.73}, {"year": 2023, "with_vat": 27.5, "no_vat": 22.73}, {"year": 2024, "with_vat": 28.4, "no_vat": 23.47}], "Canine_Collaborator": []}}, "PEDIGREE SIN TRANSFERENCIA": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2019, "with_vat": 27.0, "no_vat": 22.31}, {"year": 2020, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2021, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2022, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2023, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2024, "with_vat": 29.9, "no_vat": 24.71}], "User": [{"year": 2018, "with_vat": 36.3, "no_vat": 30.0}, {"year": 2019, "with_vat": 37.2, "no_vat": 30.74}, {"year": 2020, "with_vat": 38.0, "no_vat": 31.4}, {"year": 2021, "with_vat": 38.0, "no_vat": 31.4}, {"year": 2022, "with_vat": 39.9, "no_vat": 32.98}, {"year": 2023, "with_vat": 39.9, "no_vat": 32.98}, {"year": 2024, "with_vat": 41.1, "no_vat": 33.97}], "Canine_Collaborator": [{"year": 2018, "with_vat": 20.99, "no_vat": 17.35}, {"year": 2019, "with_vat": 21.51, "no_vat": 17.78}, {"year": 2020, "with_vat": 21.97, "no_vat": 18.16}, {"year": 2021, "with_vat": 21.98, "no_vat": 18.16}, {"year": 2022, "with_vat": 23.07, "no_vat": 19.07}, {"year": 2023, "with_vat": 23.07, "no_vat": 19.07}, {"year": 2024, "with_vat": 23.76, "no_vat": 19.64}]}}, "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES  (SI CONSTAN EN EL L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 22.23, "no_vat": 18.37}, {"year": 2020, "with_vat": 23.24, "no_vat": 19.21}, {"year": 2022, "with_vat": 24.43, "no_vat": 20.19}, {"year": 2023, "with_vat": 24.43, "no_vat": 20.19}, {"year": 2024, "with_vat": 25.18, "no_vat": 20.81}]}}, "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES  (si consta en el  L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2019, "with_vat": 22.76, "no_vat": 18.81}]}}, "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (Si la genealogía consta en la base de datos del LOE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 30.7, "no_vat": 25.37}], "User": [{"year": 2019, "with_vat": 42.3, "no_vat": 34.96}], "Canine_Collaborator": []}}, "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del L.O.E.)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 30.0, "no_vat": 24.79}], "User": [{"year": 2018, "with_vat": 41.3, "no_vat": 34.13}], "Canine_Collaborator": [{"year": 2021, "with_vat": 23.24, "no_vat": 19.21}]}}, "PEDIGREE SIN TRANSFERENCIA CUATRO GENERACIONES (si la genealogía consta en la base de datos del LOE)": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2020, "with_vat": 31.4, "no_vat": 25.95}, {"year": 2021, "with_vat": 31.4, "no_vat": 25.95}, {"year": 2022, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2023, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2024, "with_vat": 34.0, "no_vat": 28.1}], "User": [{"year": 2020, "with_vat": 43.2, "no_vat": 35.7}, {"year": 2021, "with_vat": 43.2, "no_vat": 35.7}, {"year": 2022, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2023, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2024, "with_vat": 46.8, "no_vat": 38.68}], "Canine_Collaborator": []}}, "PEDIGREE_RSCE ACCESS LBO CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 32.0, "no_vat": 26.45}], "User": [{"year": 2026, "with_vat": 43.0, "no_vat": 35.54}], "Canine_Collaborator": [{"year": 2026, "with_vat": 27.95, "no_vat": 23.1}]}}, "PEDIGREE_RSCE ACCESS LBO SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 27.0, "no_vat": 22.31}], "User": [{"year": 2026, "with_vat": 38.0, "no_vat": 31.4}], "Canine_Collaborator": [{"year": 2026, "with_vat": 24.7, "no_vat": 20.41}]}}, "PEDIGREE_RSCE ACCESS RBR CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 23.0, "no_vat": 19.01}], "User": [{"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "Canine_Collaborator": [{"year": 2026, "with_vat": 17.34, "no_vat": 14.33}]}}, "PEDIGREE_RSCE ACCESS RBR CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}], "Canine_Collaborator": [{"year": 2025, "with_vat": 16.25, "no_vat": 13.43}]}}, "PEDIGREE_RSCE ACCESS RBR SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 18.0, "no_vat": 14.88}], "User": [{"year": 2026, "with_vat": 29.0, "no_vat": 23.97}], "Canine_Collaborator": [{"year": 2026, "with_vat": 15.72, "no_vat": 12.99}]}}, "PEDIGREE_RSCE ACCESS SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 25.0, "no_vat": 20.66}], "User": [{"year": 2025, "with_vat": 35.0, "no_vat": 28.93}], "Canine_Collaborator": [{"year": 2025, "with_vat": 22.75, "no_vat": 18.8}]}}, "PEDIGREE_RSCE ACCESS TRES GENERACIONES CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}], "User": [{"year": 2025, "with_vat": 40.0, "no_vat": 33.06}], "Canine_Collaborator": [{"year": 2025, "with_vat": 26.0, "no_vat": 21.49}]}}, "PEDIGREE_RSCE DE EXPORTACION 3 GENERACIONES": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 43.15, "no_vat": 35.66}]}}, "PEDIGREE_RSCE DE EXPORTACION 4 GENERACIONES": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 46.84, "no_vat": 38.71}]}}, "PEDIGREE_RSCE DE EXPORTACIÓN CUATRO GENERACIONES EN PAPEL": {"category": "PEDIGRÍES DE EXPORTACIÓN", "prices": {"Member": [{"year": 2025, "with_vat": 73.7, "no_vat": 60.91}, {"year": 2026, "with_vat": 78.0, "no_vat": 64.46}], "User": [{"year": 2025, "with_vat": 73.7, "no_vat": 60.91}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2025, "with_vat": 29.44, "no_vat": 24.33}]}}, "PEDIGREE_RSCE DE EXPORTACIÓN TRES GENERACIONES EN PAPEL": {"category": "PEDIGRÍES DE EXPORTACIÓN", "prices": {"Member": [{"year": 2025, "with_vat": 67.1, "no_vat": 55.45}, {"year": 2026, "with_vat": 71.0, "no_vat": 58.68}], "User": [{"year": 2025, "with_vat": 67.1, "no_vat": 55.45}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2025, "with_vat": 40.78, "no_vat": 33.7}]}}, "PEDIGREE_RSCE PLUS CUATRO GENERACIONES CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 63.0, "no_vat": 52.07}], "User": [{"year": 2026, "with_vat": 74.0, "no_vat": 61.16}], "Canine_Collaborator": [{"year": 2026, "with_vat": 48.1, "no_vat": 39.75}]}}, "PEDIGREE_RSCE PLUS CUATRO GENERACIONES CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PLUS", "prices": {"Member": [{"year": 2025, "with_vat": 60.0, "no_vat": 49.59}], "User": [{"year": 2025, "with_vat": 70.0, "no_vat": 57.85}], "Canine_Collaborator": [{"year": 2025, "with_vat": 45.5, "no_vat": 37.6}]}}, "PEDIGREE_RSCE PLUS CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PLUS", "prices": {"Member": [{"year": 2025, "with_vat": 26.0, "no_vat": 21.49}], "User": [{"year": 2025, "with_vat": 35.0, "no_vat": 28.93}], "Canine_Collaborator": [{"year": 2025, "with_vat": 22.75, "no_vat": 18.8}]}}, "PEDIGREE_RSCE PLUS RRC (sin genealógica completa) SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PLUS", "prices": {"Member": [{"year": 2025, "with_vat": 19.0, "no_vat": 15.7}], "User": [{"year": 2025, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": [{"year": 2025, "with_vat": 16.25, "no_vat": 13.43}]}}, "PEDIGREE_RSCE PLUS RRC CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 50.0, "no_vat": 41.32}], "User": [{"year": 2026, "with_vat": 60.0, "no_vat": 49.59}], "Canine_Collaborator": []}}, "PEDIGREE_RSCE PLUS TRES GENERACIONES CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 53.0, "no_vat": 43.8}], "User": [{"year": 2026, "with_vat": 64.0, "no_vat": 52.89}], "Canine_Collaborator": [{"year": 2026, "with_vat": 41.6, "no_vat": 34.38}]}}, "PEDIGREE_RSCE PLUS TRES GENERACIONES CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PLUS", "prices": {"Member": [{"year": 2025, "with_vat": 50.0, "no_vat": 41.32}], "User": [{"year": 2025, "with_vat": 60.0, "no_vat": 49.59}], "Canine_Collaborator": [{"year": 2025, "with_vat": 39.0, "no_vat": 32.23}]}}, "PEDIGREE_RSCE PLUS TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PLUS", "prices": {"Member": [{"year": 2025, "with_vat": 22.0, "no_vat": 18.18}], "User": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}], "Canine_Collaborator": [{"year": 2025, "with_vat": 19.51, "no_vat": 16.12}]}}, "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 42.0, "no_vat": 34.71}], "User": [{"year": 2026, "with_vat": 62.0, "no_vat": 51.24}], "Canine_Collaborator": [{"year": 2026, "with_vat": 35.8, "no_vat": 29.59}]}}, "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 39.3, "no_vat": 32.48}], "User": [{"year": 2025, "with_vat": 57.9, "no_vat": 47.85}], "Canine_Collaborator": [{"year": 2025, "with_vat": 33.43, "no_vat": 27.63}]}}, "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 37.0, "no_vat": 30.58}], "User": [{"year": 2026, "with_vat": 48.0, "no_vat": 39.67}], "Canine_Collaborator": [{"year": 2026, "with_vat": 28.44, "no_vat": 23.5}]}}, "PEDIGREE_RSCE PREMIUM CUATRO GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 35.0, "no_vat": 28.93}], "User": [{"year": 2025, "with_vat": 45.0, "no_vat": 37.19}], "Canine_Collaborator": [{"year": 2025, "with_vat": 29.25, "no_vat": 24.17}]}}, "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 29.0, "no_vat": 23.97}], "User": [{"year": 2026, "with_vat": 44.0, "no_vat": 36.36}], "Canine_Collaborator": []}}, "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 20.3, "no_vat": 16.78}], "User": [{"year": 2025, "with_vat": 29.8, "no_vat": 24.63}], "Canine_Collaborator": []}}, "PEDIGREE_RSCE PREMIUM RRC (sin genealógica completa) SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2026, "with_vat": 39.0, "no_vat": 32.23}], "Canine_Collaborator": []}}, "PEDIGREE_RSCE PREMIUM TRES GENERACIONES CON TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 37.0, "no_vat": 30.58}], "User": [{"year": 2026, "with_vat": 55.0, "no_vat": 45.45}], "Canine_Collaborator": [{"year": 2026, "with_vat": 31.9, "no_vat": 26.36}, {"year": 2026, "with_vat": 19.14, "no_vat": 15.82}]}}, "PEDIGREE_RSCE PREMIUM TRES GENERACIONES CON TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 34.8, "no_vat": 28.76}], "User": [{"year": 2025, "with_vat": 51.3, "no_vat": 42.4}], "Canine_Collaborator": [{"year": 2025, "with_vat": 29.75, "no_vat": 24.59}, {"year": 2025, "with_vat": 17.29, "no_vat": 14.29}]}}, "PEDIGREE_RSCE PREMIUM TRES GENERACIONES SIN TRANSFERENCIA EN PAPEL": {"category": "PEDIGRÍES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}], "User": [{"year": 2025, "with_vat": 40.0, "no_vat": 33.06}], "Canine_Collaborator": [{"year": 2025, "with_vat": 26.0, "no_vat": 21.49}]}}, "PEDIGREE_RSCE PREMIUN 3 GENERACIONES SIN TRANSFERENCIA": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 25.41, "no_vat": 21.0}]}}, "PLUS ENVÍO CARTILLA POR MENSAJERÍA (hasta un máximo de 10 cartillas por guía)": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2026, "with_vat": 8.0, "no_vat": 6.61}, {"year": 2026, "with_vat": 8.0, "no_vat": 6.61}], "User": [{"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2026, "with_vat": null, "no_vat": null}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "PLUS ENVÍO CARTILLA POR MENSAJERÍA (hasta un máximo de 10 cartillas por propietario)": {"category": "CAMPO Y TRABAJO (MUESTRA, SPANIELS, RETRIEVERS, TECKEL, TERRIER, RASTRO)", "prices": {"Member": [{"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2026, "with_vat": 8.0, "no_vat": 6.61}], "User": [{"year": 2023, "with_vat": 4.8, "no_vat": 3.97}, {"year": 2024, "with_vat": 4.9, "no_vat": 4.05}, {"year": 2025, "with_vat": 5.3, "no_vat": 4.38}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "PLUS POR TRAMITACIÓN URGENTE DE CAMADAS, PEDIGREES Y PEDIGREES DE EXPORTACIÓN": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 51.0, "no_vat": 42.15}], "User": [{"year": 2018, "with_vat": 54.5, "no_vat": 45.04}], "Canine_Collaborator": [{"year": 2018, "with_vat": 49.23, "no_vat": 40.69}]}}, "POR CADA ANOTACIÓN DE INFORME ADN O INFORME DE DISPLASIA EN BASE DE DATOS": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2021, "with_vat": 6.0, "no_vat": 4.96}], "User": [{"year": 2021, "with_vat": 8.0, "no_vat": 6.61}], "Canine_Collaborator": []}}, "POR CADA ANOTACIÓN EN BASE DE DATOS DE PRUEBA GENÉTICA, INFORME DE ADN O INFORME DE DISPLASIA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2022, "with_vat": 6.3, "no_vat": 5.21}], "User": [{"year": 2022, "with_vat": 8.4, "no_vat": 6.94}], "Canine_Collaborator": []}}, "POR PERRO INSCRITO EN CADA PRUEBA CON CAC / CACT": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 0.5, "no_vat": 0.41}, {"year": 2019, "with_vat": 0.5, "no_vat": 0.41}, {"year": 2020, "with_vat": 0.75, "no_vat": 0.62}, {"year": 2021, "with_vat": 0.75, "no_vat": 0.62}, {"year": 2022, "with_vat": 0.79, "no_vat": 0.65}, {"year": 2023, "with_vat": 0.79, "no_vat": 0.65}, {"year": 2024, "with_vat": 1.0, "no_vat": 0.83}, {"year": 2025, "with_vat": 1.1, "no_vat": 0.91}, {"year": 2026, "with_vat": 1.21, "no_vat": 1.0}]}}, "PRUEBA ADN / PRUEBAS PATERNIDAD, CON ANOTACIÓN EN BASE DE DATOS, ACOGIÉNDOSE AL CONVENIO RSCE – HISPALAB (por perro)": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2023, "with_vat": 34.7, "no_vat": 28.68}], "User": [{"year": 2023, "with_vat": 42.0, "no_vat": 34.71}], "Canine_Collaborator": []}}, "PRUEBA ADN/PATERNIDAD Y ANOTACIÓN EN BASE DE DATOS, SEGÚN CONVENIO RSCE – HISPALAB (por perro)": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2024, "with_vat": 35.8, "no_vat": 29.59}], "User": [{"year": 2024, "with_vat": 43.3, "no_vat": 35.79}], "Canine_Collaborator": []}}, "PRUEBA ADN/PATERNIDAD Y ANOTACIÓN EN BASE DE DATOS-CONVENIO RSCE – HISPALAB (por perro) RAZAS ESPAÑOLAS VULNERABLES": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2024, "with_vat": 25.0, "no_vat": 20.66}], "User": [{"year": 2024, "with_vat": 30.0, "no_vat": 24.79}], "Canine_Collaborator": []}}, "RECARGO 100% (más de 6 meses y menos de 9 meses)  Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2019, "with_vat": 10.19, "no_vat": 8.42}]}}, "RECARGO 100% (más de 6 meses y menos de 9) Art. 6.)13 Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 9.96, "no_vat": 8.23}, {"year": 2020, "with_vat": 10.42, "no_vat": 8.61}, {"year": 2022, "with_vat": 10.9, "no_vat": 9.01}, {"year": 2023, "with_vat": 10.9, "no_vat": 9.01}, {"year": 2024, "with_vat": 11.25, "no_vat": 9.3}]}}, "RECARGO 100% EN INSCRIPCION CACHORRO (más de 6 meses y menos de 9 meses de edad) Reglamento LOE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 14.6, "no_vat": 12.07}], "User": [{"year": 2019, "with_vat": 17.6, "no_vat": 14.55}], "Canine_Collaborator": []}}, "RECARGO 100% EN INSCRIPCIÓN CACHORRO (más de 6 meses y menos de 9 meses de edad)": {"category": "INSCRIPCIONES EN EL LOE / RRC", "prices": {"Member": [{"year": 2020, "with_vat": 14.9, "no_vat": 12.31}], "User": [{"year": 2020, "with_vat": 18.0, "no_vat": 14.88}], "Canine_Collaborator": []}}, "RECARGO 100% EN INSCRIPCIÓN CACHORRO (más de 6 meses y menos de 9 meses de edad) Art. 6.13 Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 14.2, "no_vat": 11.74}], "User": [{"year": 2018, "with_vat": 17.2, "no_vat": 14.21}], "Canine_Collaborator": [{"year": 2021, "with_vat": 10.42, "no_vat": 8.61}]}}, "RECARGO 200% (más de 9 meses, menos de 12 meses) Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2019, "with_vat": 20.35, "no_vat": 16.82}]}}, "RECARGO 200% (más de 9 meses, menos de 12) Art. 6.)13 Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 19.89, "no_vat": 16.44}, {"year": 2020, "with_vat": 20.81, "no_vat": 17.2}, {"year": 2022, "with_vat": 21.85, "no_vat": 18.06}, {"year": 2023, "with_vat": 21.85, "no_vat": 18.06}, {"year": 2024, "with_vat": 22.49, "no_vat": 18.59}]}}, "RECARGO 200% EN INSCRIPCION CACHORRO (más de 9 meses y menos de 12 meses de edad)  Reglamento LOE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 29.2, "no_vat": 24.13}], "User": [{"year": 2019, "with_vat": 35.2, "no_vat": 29.09}], "Canine_Collaborator": []}}, "RECARGO 200% EN INSCRIPCIÓN CACHORRO (más de 9 meses y menos de 12 meses de edad)": {"category": "INSCRIPCIONES EN EL LOE / RRC", "prices": {"Member": [{"year": 2020, "with_vat": 29.8, "no_vat": 24.63}], "User": [{"year": 2020, "with_vat": 36.0, "no_vat": 29.75}], "Canine_Collaborator": []}}, "RECARGO 200% EN INSCRIPCIÓN CACHORRO (más de 9 meses y menos de 12 meses de edad) Art. 6.13 Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 28.3, "no_vat": 23.39}], "User": [{"year": 2018, "with_vat": 34.4, "no_vat": 28.43}], "Canine_Collaborator": [{"year": 2021, "with_vat": 20.81, "no_vat": 17.2}]}}, "RECARGO 300% EN INSCRIPCION CACHORRO (más de 12 meses y menos de 18 meses de edad) Reglamento LOE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 43.8, "no_vat": 36.2}], "User": [{"year": 2019, "with_vat": 52.8, "no_vat": 43.64}], "Canine_Collaborator": []}}, "RECARGO 300% EN INSCRIPCIÓN CACHORRO (más de 12 meses y menos de 18 meses de edad)": {"category": "INSCRIPCIONES EN EL LOE / RRC", "prices": {"Member": [{"year": 2020, "with_vat": 44.7, "no_vat": 36.94}], "User": [{"year": 2020, "with_vat": 53.9, "no_vat": 44.55}], "Canine_Collaborator": []}}, "RECARGO 300% EN INSCRIPCIÓN CACHORRO (más de 12 meses y menos de 18 meses de edad) Art. 6.13 Reglamento L.O.E.": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 42.5, "no_vat": 35.12}], "User": [{"year": 2018, "with_vat": 51.6, "no_vat": 42.64}], "Canine_Collaborator": [{"year": 2021, "with_vat": 31.18, "no_vat": 25.77}]}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/ RBR con más de 6 meses y menos de 9 meses de edad": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2026, "with_vat": 12.0, "no_vat": 9.92}], "User": [{"year": 2026, "with_vat": 17.0, "no_vat": 14.05}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 12 meses y menos de 18 meses de edad": {"category": "INSCRIPCIONES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2026, "with_vat": 36.0, "no_vat": 29.75}], "User": [{"year": 2025, "with_vat": 45.0, "no_vat": 37.19}, {"year": 2026, "with_vat": 51.0, "no_vat": 42.15}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 6 meses y menos de 9 meses de edad": {"category": "INSCRIPCIONES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 11.0, "no_vat": 9.09}], "User": [{"year": 2025, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO ACCESS LBO/RBR con más de 9 meses y menos de 12 meses de edad": {"category": "INSCRIPCIONES ACCESS LBO/RBR", "prices": {"Member": [{"year": 2025, "with_vat": 22.0, "no_vat": 18.18}, {"year": 2026, "with_vat": 24.0, "no_vat": 19.83}], "User": [{"year": 2025, "with_vat": 30.0, "no_vat": 24.79}, {"year": 2026, "with_vat": 34.0, "no_vat": 28.1}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 12 meses y menos de 18 meses de edad": {"category": "INSCRIPCIONES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 50.7, "no_vat": 41.9}, {"year": 2026, "with_vat": 54.0, "no_vat": 44.63}], "User": [{"year": 2025, "with_vat": 61.5, "no_vat": 50.83}, {"year": 2026, "with_vat": 66.0, "no_vat": 54.55}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 6 meses y menos de 9 meses de edad": {"category": "INSCRIPCIONES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 16.9, "no_vat": 13.97}, {"year": 2026, "with_vat": 18.0, "no_vat": 14.88}], "User": [{"year": 2025, "with_vat": 20.5, "no_vat": 16.94}, {"year": 2026, "with_vat": 22.0, "no_vat": 18.18}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO PREMIUM LOE/RRC con más de 9 meses y menos de 12 meses de edad": {"category": "INSCRIPCIONES PREMIUM LOE/RRC", "prices": {"Member": [{"year": 2025, "with_vat": 33.8, "no_vat": 27.93}, {"year": 2026, "with_vat": 36.0, "no_vat": 29.75}], "User": [{"year": 2025, "with_vat": 41.0, "no_vat": 33.88}, {"year": 2026, "with_vat": 44.0, "no_vat": 36.36}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 12 meses y menos de 18 meses de edad": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2021, "with_vat": 44.7, "no_vat": 36.94}, {"year": 2022, "with_vat": 46.8, "no_vat": 38.68}, {"year": 2023, "with_vat": 46.8, "no_vat": 38.68}, {"year": 2024, "with_vat": 48.3, "no_vat": 39.92}], "User": [{"year": 2021, "with_vat": 53.9, "no_vat": 44.55}, {"year": 2022, "with_vat": 56.7, "no_vat": 46.86}, {"year": 2023, "with_vat": 56.7, "no_vat": 46.86}, {"year": 2024, "with_vat": 58.5, "no_vat": 48.35}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 6 meses y menos de 9 meses de edad": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2021, "with_vat": 14.9, "no_vat": 12.31}, {"year": 2022, "with_vat": 15.6, "no_vat": 12.89}, {"year": 2023, "with_vat": 15.6, "no_vat": 12.89}, {"year": 2024, "with_vat": 16.1, "no_vat": 13.31}], "User": [{"year": 2021, "with_vat": 18.0, "no_vat": 14.88}, {"year": 2022, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2023, "with_vat": 18.9, "no_vat": 15.62}, {"year": 2024, "with_vat": 19.5, "no_vat": 16.12}], "Canine_Collaborator": []}}, "RECARGO POR INSCRIPCIÓN DE CACHORRO con más de 9 meses y menos de 12 meses de edad": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2021, "with_vat": 29.8, "no_vat": 24.63}, {"year": 2022, "with_vat": 31.2, "no_vat": 25.79}, {"year": 2023, "with_vat": 31.2, "no_vat": 25.79}, {"year": 2024, "with_vat": 32.2, "no_vat": 26.61}], "User": [{"year": 2021, "with_vat": 36.0, "no_vat": 29.75}, {"year": 2022, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2023, "with_vat": 37.8, "no_vat": 31.24}, {"year": 2024, "with_vat": 39.0, "no_vat": 32.23}], "Canine_Collaborator": []}}, "RECTIFICACION AFIJO": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 35.5, "no_vat": 29.34}], "User": [{"year": 2018, "with_vat": 45.7, "no_vat": 37.77}], "Canine_Collaborator": []}}, "RECTIFICACIÓN DE AFIJO": {"category": "AFIJOS", "prices": {"Member": [{"year": 2019, "with_vat": 36.4, "no_vat": 30.08}, {"year": 2020, "with_vat": 37.2, "no_vat": 30.74}, {"year": 2021, "with_vat": 37.2, "no_vat": 30.74}, {"year": 2022, "with_vat": 39.1, "no_vat": 32.31}, {"year": 2023, "with_vat": 39.1, "no_vat": 32.31}, {"year": 2024, "with_vat": 40.3, "no_vat": 33.31}, {"year": 2025, "with_vat": 42.3, "no_vat": 34.96}, {"year": 2026, "with_vat": 50.0, "no_vat": 41.32}], "User": [{"year": 2019, "with_vat": 46.8, "no_vat": 38.68}, {"year": 2020, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2021, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2022, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2023, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2024, "with_vat": 51.8, "no_vat": 42.81}, {"year": 2025, "with_vat": 54.4, "no_vat": 44.96}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "Canine_Collaborator": [{"year": 2018, "with_vat": 30.6, "no_vat": 25.29}, {"year": 2019, "with_vat": 31.34, "no_vat": 25.9}, {"year": 2020, "with_vat": 32.0, "no_vat": 26.45}, {"year": 2021, "with_vat": 32.01, "no_vat": 26.45}, {"year": 2022, "with_vat": 33.61, "no_vat": 27.78}, {"year": 2023, "with_vat": 33.61, "no_vat": 27.78}, {"year": 2024, "with_vat": 34.69, "no_vat": 28.67}, {"year": 2025, "with_vat": 36.43, "no_vat": 30.11}, {"year": 2026, "with_vat": 43.54, "no_vat": 35.98}]}}, "RECURSO AL DIAGNOSTICO DE DISPLASIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 76.2, "no_vat": 62.98}, {"year": 2019, "with_vat": 78.0, "no_vat": 64.46}], "User": [{"year": 2018, "with_vat": 101.5, "no_vat": 83.88}, {"year": 2019, "with_vat": 103.9, "no_vat": 85.87}], "Canine_Collaborator": []}}, "RECURSO AL DIAGNÓSTICO DE DISPLASIA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2020, "with_vat": 79.6, "no_vat": 65.79}, {"year": 2021, "with_vat": 79.6, "no_vat": 65.79}, {"year": 2022, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2023, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2024, "with_vat": 86.2, "no_vat": 71.24}, {"year": 2025, "with_vat": 90.5, "no_vat": 74.79}, {"year": 2026, "with_vat": 100.0, "no_vat": 82.64}], "User": [{"year": 2020, "with_vat": 106.0, "no_vat": 87.6}, {"year": 2021, "with_vat": 106.0, "no_vat": 87.6}, {"year": 2022, "with_vat": 111.3, "no_vat": 91.98}, {"year": 2023, "with_vat": 111.3, "no_vat": 91.98}, {"year": 2024, "with_vat": 114.8, "no_vat": 94.88}, {"year": 2025, "with_vat": 120.5, "no_vat": 99.59}, {"year": 2026, "with_vat": 134.0, "no_vat": 110.74}], "Canine_Collaborator": []}}, "RECURSO COMISIÓN LIBRO DE ORÍGENES ESPAÑOL": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2025, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2025, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2026, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": []}}, "REGISTRO INCIAL DE RAZA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 41.0, "no_vat": 33.88}], "User": [{"year": 2019, "with_vat": null, "no_vat": null}], "Canine_Collaborator": []}}, "REGISTRO INICIAL DE RAZA": {"category": "REGISTROS INICIALES Y CONFIRMACIONES DE RAZA", "prices": {"Member": [{"year": 2018, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2020, "with_vat": 41.9, "no_vat": 34.63}, {"year": 2021, "with_vat": 41.9, "no_vat": 34.63}], "User": [{"year": 2018, "with_vat": 40.0, "no_vat": 33.06}, {"year": 2020, "with_vat": 41.9, "no_vat": 34.63}, {"year": 2021, "with_vat": 41.9, "no_vat": 34.63}], "Canine_Collaborator": [{"year": 2018, "with_vat": 17.59, "no_vat": 14.54}, {"year": 2019, "with_vat": 18.03, "no_vat": 14.9}, {"year": 2020, "with_vat": 18.43, "no_vat": 15.23}, {"year": 2021, "with_vat": 18.42, "no_vat": 15.22}]}}, "REGISTRO INICIAL GRUPO ÉTNICO": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2021, "with_vat": 25.0, "no_vat": 20.66}, {"year": 2022, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2023, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": 11.3, "no_vat": 9.34}], "User": [{"year": 2021, "with_vat": 25.0, "no_vat": 20.66}, {"year": 2022, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2023, "with_vat": 10.5, "no_vat": 8.68}, {"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2024, "with_vat": 10.79, "no_vat": 8.92}, {"year": 2025, "with_vat": 11.29, "no_vat": 9.33}, {"year": 2026, "with_vat": 11.29, "no_vat": 9.33}]}}, "REGISTRO INICIAL RAZAS ESPAÑOLAS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2022, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2023, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2024, "with_vat": 27.1, "no_vat": 22.4}, {"year": 2025, "with_vat": 28.5, "no_vat": 23.55}, {"year": 2026, "with_vat": 28.5, "no_vat": 23.55}], "User": [{"year": 2022, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2023, "with_vat": 26.3, "no_vat": 21.74}, {"year": 2024, "with_vat": 27.1, "no_vat": 22.4}, {"year": 2025, "with_vat": 28.5, "no_vat": 23.55}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2022, "with_vat": 14.0, "no_vat": 11.57}, {"year": 2023, "with_vat": 14.0, "no_vat": 11.57}, {"year": 2024, "with_vat": 14.42, "no_vat": 11.92}, {"year": 2025, "with_vat": 15.17, "no_vat": 12.54}, {"year": 2026, "with_vat": 15.17, "no_vat": 12.54}]}}, "REGISTRO INICIAL RAZAS ESPAÑOLAS VULNERABLES": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": 11.3, "no_vat": 9.34}], "User": [{"year": 2024, "with_vat": 10.8, "no_vat": 8.93}, {"year": 2025, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2024, "with_vat": 10.79, "no_vat": 8.92}, {"year": 2025, "with_vat": 11.29, "no_vat": 9.33}, {"year": 2026, "with_vat": 11.29, "no_vat": 9.33}]}}, "REGISTRO INICIAL RESTO DE RAZAS": {"category": "INSCRIPCIONES", "prices": {"Member": [{"year": 2022, "with_vat": 44.0, "no_vat": 36.36}, {"year": 2023, "with_vat": 44.0, "no_vat": 36.36}, {"year": 2024, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2025, "with_vat": 65.0, "no_vat": 53.72}, {"year": 2026, "with_vat": 70.0, "no_vat": 57.85}], "User": [{"year": 2022, "with_vat": 44.0, "no_vat": 36.36}, {"year": 2023, "with_vat": 44.0, "no_vat": 36.36}, {"year": 2024, "with_vat": 45.4, "no_vat": 37.52}, {"year": 2025, "with_vat": 65.0, "no_vat": 53.72}, {"year": 2026, "with_vat": null, "no_vat": null}], "Canine_Collaborator": [{"year": 2022, "with_vat": 19.35, "no_vat": 15.99}, {"year": 2023, "with_vat": 19.35, "no_vat": 15.99}, {"year": 2024, "with_vat": 19.96, "no_vat": 16.5}, {"year": 2025, "with_vat": 28.58, "no_vat": 23.62}, {"year": 2026, "with_vat": 30.78, "no_vat": 25.44}]}}, "RENOVACIÓN ANUAL LICENCIA": {"category": "TRABAJO (IGP / IGP-IFH / OBEDIENCIA / MONDIORING)", "prices": {"Member": [{"year": 2019, "with_vat": 15.8, "no_vat": 13.06}, {"year": 2019, "with_vat": 15.8, "no_vat": 13.06}, {"year": 2020, "with_vat": 16.5, "no_vat": 13.64}, {"year": 2020, "with_vat": 16.5, "no_vat": 13.64}, {"year": 2021, "with_vat": 16.5, "no_vat": 13.64}, {"year": 2021, "with_vat": 16.5, "no_vat": 13.64}, {"year": 2022, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2022, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2023, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2023, "with_vat": 17.3, "no_vat": 14.3}, {"year": 2024, "with_vat": 17.8, "no_vat": 14.71}, {"year": 2024, "with_vat": 17.8, "no_vat": 14.71}, {"year": 2025, "with_vat": 18.7, "no_vat": 15.45}, {"year": 2025, "with_vat": 18.7, "no_vat": 15.45}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}, {"year": 2026, "with_vat": 20.0, "no_vat": 16.53}], "User": [{"year": 2019, "with_vat": 32.5, "no_vat": 26.86}, {"year": 2019, "with_vat": 32.5, "no_vat": 26.86}, {"year": 2020, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2020, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2021, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2021, "with_vat": 33.0, "no_vat": 27.27}, {"year": 2022, "with_vat": 34.7, "no_vat": 28.68}, {"year": 2022, "with_vat": 34.7, "no_vat": 28.68}, {"year": 2023, "with_vat": 34.7, "no_vat": 28.68}, {"year": 2023, "with_vat": 34.7, "no_vat": 28.68}, {"year": 2024, "with_vat": 35.8, "no_vat": 29.59}, {"year": 2024, "with_vat": 35.8, "no_vat": 29.59}, {"year": 2025, "with_vat": 37.6, "no_vat": 31.07}, {"year": 2025, "with_vat": 37.6, "no_vat": 31.07}, {"year": 2026, "with_vat": 41.0, "no_vat": 33.88}, {"year": 2026, "with_vat": 41.0, "no_vat": 33.88}], "Canine_Collaborator": []}}, "RENOVACIÓN ANUAL LICIENCIA": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 15.4, "no_vat": 12.73}, {"year": 2018, "with_vat": 15.4, "no_vat": 12.73}], "User": [{"year": 2018, "with_vat": 31.75, "no_vat": 26.24}, {"year": 2018, "with_vat": 31.75, "no_vat": 26.24}], "Canine_Collaborator": []}}, "SIN DESCRIPCIÓN (tasa fija 0,25€)": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2018, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2019, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2020, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2022, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2023, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2024, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2025, "with_vat": 0.25, "no_vat": 0.21}, {"year": 2026, "with_vat": 0.25, "no_vat": 0.21}]}}, "SOLICITUD AFIJO": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 103.8, "no_vat": 85.79}], "User": [{"year": 2018, "with_vat": 133.7, "no_vat": 110.5}], "Canine_Collaborator": []}}, "SOLICITUD DE AFIJO": {"category": "AFIJOS", "prices": {"Member": [{"year": 2019, "with_vat": 106.2, "no_vat": 87.77}, {"year": 2020, "with_vat": 108.4, "no_vat": 89.59}, {"year": 2021, "with_vat": 108.4, "no_vat": 89.59}, {"year": 2022, "with_vat": 113.8, "no_vat": 94.05}, {"year": 2023, "with_vat": 113.8, "no_vat": 94.05}, {"year": 2024, "with_vat": 117.3, "no_vat": 96.94}, {"year": 2025, "with_vat": 175.0, "no_vat": 144.63}, {"year": 2026, "with_vat": 185.0, "no_vat": 152.89}], "User": [{"year": 2019, "with_vat": 136.8, "no_vat": 113.06}, {"year": 2020, "with_vat": 139.6, "no_vat": 115.37}, {"year": 2021, "with_vat": 139.6, "no_vat": 115.37}, {"year": 2022, "with_vat": 146.6, "no_vat": 121.16}, {"year": 2023, "with_vat": 146.6, "no_vat": 121.16}, {"year": 2024, "with_vat": 151.1, "no_vat": 124.88}, {"year": 2025, "with_vat": 225.0, "no_vat": 185.95}, {"year": 2026, "with_vat": 238.0, "no_vat": 196.69}], "Canine_Collaborator": [{"year": 2018, "with_vat": 89.59, "no_vat": 74.04}, {"year": 2019, "with_vat": 91.67, "no_vat": 75.76}, {"year": 2020, "with_vat": 93.55, "no_vat": 77.31}, {"year": 2021, "with_vat": 93.55, "no_vat": 77.31}, {"year": 2022, "with_vat": 98.24, "no_vat": 81.19}, {"year": 2023, "with_vat": 98.24, "no_vat": 81.19}, {"year": 2024, "with_vat": 101.25, "no_vat": 83.68}, {"year": 2025, "with_vat": 150.78, "no_vat": 124.61}, {"year": 2026, "with_vat": 159.49, "no_vat": 131.81}]}}, "TRAMITACION TITULOS FCI": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2018, "with_vat": 5.1, "no_vat": 4.21}], "User": [{"year": 2018, "with_vat": 5.3, "no_vat": 4.38}], "Canine_Collaborator": []}}, "TRAMITACION URGENTE  CARTILLA DE ASISTENTE (exclusivamente a través de las oficinas de la RSCE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 46.8, "no_vat": 38.68}], "User": [{"year": 2019, "with_vat": 50.0, "no_vat": 41.32}], "Canine_Collaborator": []}}, "TRAMITACION URGENTE CARTILLAS DE PRUEBAS DEPORTIVAS (exclusivamente a través de las oficinas de la RSCE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 46.8, "no_vat": 38.68}], "User": [{"year": 2019, "with_vat": 50.0, "no_vat": 41.32}], "Canine_Collaborator": []}}, "TRAMITACION URGENTE CERTIFICADOS DE COLA CORTA (se realizará exclusivamente a través de las oficinas de la RSCE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 23.4, "no_vat": 19.34}], "User": [{"year": 2019, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": []}}, "TRAMITACION URGENTE CERTIFICADOS DE TRABAJO (se realizará exclusivamente a través de las oficinas de la RSCE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 23.4, "no_vat": 19.34}], "User": [{"year": 2019, "with_vat": 25.0, "no_vat": 20.66}], "Canine_Collaborator": []}}, "TRAMITACION URGENTE DE CAMADAS Y PEDIGREES (exclusivamente a través de las oficinas de la RSCE)": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 52.2, "no_vat": 43.14}], "User": [{"year": 2019, "with_vat": 55.8, "no_vat": 46.12}], "Canine_Collaborator": []}}, "TRAMITACIÓN TITULOS F.C.I.": {"category": "Sin categorizar", "prices": {"Member": [{"year": 2019, "with_vat": 5.3, "no_vat": 4.38}], "User": [{"year": 2019, "with_vat": 5.5, "no_vat": 4.55}], "Canine_Collaborator": []}}, "TRAMITACIÓN TITULOS FCI": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2020, "with_vat": 5.5, "no_vat": 4.55}], "User": [{"year": 2020, "with_vat": 5.7, "no_vat": 4.71}], "Canine_Collaborator": []}}, "TRAMITACIÓN TÍTULOS FCI": {"category": "OTROS SERVICIOS", "prices": {"Member": [{"year": 2021, "with_vat": 5.5, "no_vat": 4.55}, {"year": 2022, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2023, "with_vat": 5.8, "no_vat": 4.79}, {"year": 2024, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2025, "with_vat": 6.3, "no_vat": 5.21}, {"year": 2026, "with_vat": 10.0, "no_vat": 8.26}], "User": [{"year": 2021, "with_vat": 5.7, "no_vat": 4.71}, {"year": 2022, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2023, "with_vat": 6.0, "no_vat": 4.96}, {"year": 2024, "with_vat": 6.2, "no_vat": 5.12}, {"year": 2025, "with_vat": 6.5, "no_vat": 5.37}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE CARTILLAS DE ASISTENTE": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2020, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2021, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2022, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2023, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2024, "with_vat": 51.8, "no_vat": 42.81}, {"year": 2025, "with_vat": 54.4, "no_vat": 44.96}, {"year": 2026, "with_vat": 40.0, "no_vat": 33.06}], "User": [{"year": 2020, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2021, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2022, "with_vat": 53.6, "no_vat": 44.3}, {"year": 2023, "with_vat": 53.6, "no_vat": 44.3}, {"year": 2024, "with_vat": 55.3, "no_vat": 45.7}, {"year": 2025, "with_vat": 58.1, "no_vat": 48.02}, {"year": 2026, "with_vat": 55.0, "no_vat": 45.45}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE CARTILLAS DE PRUEBAS DEPORTIVAS": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2020, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2021, "with_vat": 47.8, "no_vat": 39.5}, {"year": 2022, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2023, "with_vat": 50.2, "no_vat": 41.49}, {"year": 2024, "with_vat": 51.8, "no_vat": 42.81}, {"year": 2025, "with_vat": 54.4, "no_vat": 44.96}, {"year": 2026, "with_vat": 60.0, "no_vat": 49.59}], "User": [{"year": 2020, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2021, "with_vat": 51.0, "no_vat": 42.15}, {"year": 2022, "with_vat": 53.6, "no_vat": 44.3}, {"year": 2023, "with_vat": 53.6, "no_vat": 44.3}, {"year": 2024, "with_vat": 55.3, "no_vat": 45.7}, {"year": 2025, "with_vat": 58.1, "no_vat": 48.02}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE CERTIFICADOS DE COLA CORTA": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2020, "with_vat": 23.9, "no_vat": 19.75}, {"year": 2021, "with_vat": 23.9, "no_vat": 19.75}, {"year": 2022, "with_vat": 25.1, "no_vat": 20.74}, {"year": 2023, "with_vat": 25.1, "no_vat": 20.74}, {"year": 2024, "with_vat": 25.9, "no_vat": 21.4}, {"year": 2025, "with_vat": 27.2, "no_vat": 22.48}, {"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "User": [{"year": 2020, "with_vat": 25.5, "no_vat": 21.07}, {"year": 2021, "with_vat": 25.5, "no_vat": 21.07}, {"year": 2022, "with_vat": 26.8, "no_vat": 22.15}, {"year": 2023, "with_vat": 26.8, "no_vat": 22.15}, {"year": 2024, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2025, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2026, "with_vat": 38.0, "no_vat": 31.4}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE CERTIFICADOS DE TRABAJO": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2020, "with_vat": 23.9, "no_vat": 19.75}, {"year": 2021, "with_vat": 23.9, "no_vat": 19.75}, {"year": 2022, "with_vat": 25.1, "no_vat": 20.74}, {"year": 2023, "with_vat": 25.1, "no_vat": 20.74}, {"year": 2024, "with_vat": 25.9, "no_vat": 21.4}, {"year": 2025, "with_vat": 27.2, "no_vat": 22.48}, {"year": 2026, "with_vat": 35.0, "no_vat": 28.93}], "User": [{"year": 2020, "with_vat": 25.5, "no_vat": 21.07}, {"year": 2021, "with_vat": 25.5, "no_vat": 21.07}, {"year": 2022, "with_vat": 26.8, "no_vat": 22.15}, {"year": 2023, "with_vat": 26.8, "no_vat": 22.15}, {"year": 2024, "with_vat": 27.6, "no_vat": 22.81}, {"year": 2025, "with_vat": 29.0, "no_vat": 23.97}, {"year": 2026, "with_vat": 38.0, "no_vat": 31.4}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE INSCRIPCIONES EN LOE/RRC, PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2021, "with_vat": 53.3, "no_vat": 44.05}, {"year": 2022, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2023, "with_vat": 56.0, "no_vat": 46.28}, {"year": 2024, "with_vat": 57.7, "no_vat": 47.69}, {"year": 2025, "with_vat": 60.6, "no_vat": 50.08}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "User": [{"year": 2021, "with_vat": 57.0, "no_vat": 47.11}, {"year": 2022, "with_vat": 59.9, "no_vat": 49.5}, {"year": 2023, "with_vat": 59.9, "no_vat": 49.5}, {"year": 2024, "with_vat": 61.8, "no_vat": 51.07}, {"year": 2025, "with_vat": 64.9, "no_vat": 53.64}, {"year": 2026, "with_vat": 70.0, "no_vat": 57.85}], "Canine_Collaborator": []}}, "TRAMITACIÓN URGENTE DE PEDIGRÍES, TRANSFERENCIAS DE PROPIEDAD E INSCRIPCIONES EN LOE/RRC": {"category": "PLUS POR TRAMITACIÓN URGENTE", "prices": {"Member": [{"year": 2020, "with_vat": 53.3, "no_vat": 44.05}], "User": [{"year": 2020, "with_vat": 57.0, "no_vat": 47.11}], "Canine_Collaborator": []}}, "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CADERA": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2018, "with_vat": 50.8, "no_vat": 41.98}, {"year": 2019, "with_vat": 52.0, "no_vat": 42.98}, {"year": 2020, "with_vat": 53.1, "no_vat": 43.88}, {"year": 2021, "with_vat": 53.1, "no_vat": 43.88}, {"year": 2022, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2023, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2024, "with_vat": 57.5, "no_vat": 47.52}, {"year": 2025, "with_vat": 60.4, "no_vat": 49.92}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "User": [{"year": 2018, "with_vat": 76.2, "no_vat": 62.98}, {"year": 2019, "with_vat": 78.0, "no_vat": 64.46}, {"year": 2020, "with_vat": 79.6, "no_vat": 65.79}, {"year": 2021, "with_vat": 79.6, "no_vat": 65.79}, {"year": 2022, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2023, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2024, "with_vat": 86.2, "no_vat": 71.24}, {"year": 2025, "with_vat": 90.5, "no_vat": 74.79}, {"year": 2026, "with_vat": 98.0, "no_vat": 80.99}], "Canine_Collaborator": []}}, "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CODO": {"category": "SALUD DEL PERRO", "prices": {"Member": [{"year": 2021, "with_vat": 53.1, "no_vat": 43.88}, {"year": 2022, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2023, "with_vat": 55.8, "no_vat": 46.12}, {"year": 2024, "with_vat": 57.5, "no_vat": 47.52}, {"year": 2025, "with_vat": 60.4, "no_vat": 49.92}, {"year": 2026, "with_vat": 65.0, "no_vat": 53.72}], "User": [{"year": 2021, "with_vat": 79.6, "no_vat": 65.79}, {"year": 2022, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2023, "with_vat": 83.6, "no_vat": 69.09}, {"year": 2024, "with_vat": 86.2, "no_vat": 71.24}, {"year": 2025, "with_vat": 90.5, "no_vat": 74.79}, {"year": 2026, "with_vat": 98.0, "no_vat": 80.99}], "Canine_Collaborator": []}}, "TRAMITACIÓN VALORACIÓN DEL GRADO DE DISPLASIA DE CODOS": {"category": "DISPLASIA", "prices": {"Member": [{"year": 2018, "with_vat": 50.8, "no_vat": 41.98}, {"year": 2019, "with_vat": 52.0, "no_vat": 42.98}, {"year": 2020, "with_vat": 53.1, "no_vat": 43.88}], "User": [{"year": 2018, "with_vat": 76.2, "no_vat": 62.98}, {"year": 2019, "with_vat": 78.0, "no_vat": 64.46}, {"year": 2020, "with_vat": 79.6, "no_vat": 65.79}], "Canine_Collaborator": []}}, "TRANSFERENCIA DE PROPIEDAD EN PEDIGREE, REGISTRO INICIAL O PERRO IMPORTADO": {"category": "PEDIGRÍES Y TRANSFERENCIAS DE PROPIEDAD", "prices": {"Member": [{"year": 2018, "with_vat": 11.0, "no_vat": 9.09}, {"year": 2019, "with_vat": 11.3, "no_vat": 9.34}, {"year": 2020, "with_vat": 11.6, "no_vat": 9.59}, {"year": 2021, "with_vat": 11.6, "no_vat": 9.59}, {"year": 2022, "with_vat": 12.2, "no_vat": 10.08}, {"year": 2023, "with_vat": 12.2, "no_vat": 10.08}, {"year": 2024, "with_vat": 12.6, "no_vat": 10.41}], "User": [{"year": 2018, "with_vat": 13.3, "no_vat": 10.99}, {"year": 2019, "with_vat": 13.7, "no_vat": 11.32}, {"year": 2020, "with_vat": 14.0, "no_vat": 11.57}, {"year": 2021, "with_vat": 14.0, "no_vat": 11.57}, {"year": 2022, "with_vat": 14.7, "no_vat": 12.15}, {"year": 2023, "with_vat": 14.7, "no_vat": 12.15}, {"year": 2024, "with_vat": 15.2, "no_vat": 12.56}], "Canine_Collaborator": [{"year": 2018, "with_vat": 7.76, "no_vat": 6.41}, {"year": 2019, "with_vat": 7.99, "no_vat": 6.6}, {"year": 2020, "with_vat": 8.16, "no_vat": 6.74}, {"year": 2021, "with_vat": 8.16, "no_vat": 6.74}, {"year": 2022, "with_vat": 8.57, "no_vat": 7.08}, {"year": 2023, "with_vat": 8.57, "no_vat": 7.08}, {"year": 2024, "with_vat": 8.86, "no_vat": 7.32}]}}, "TRANSFERENCIA DE PROPIEDAD GRUPOS ETNICOS": {"category": "Sin categorizar", "prices": {"Member": [], "User": [], "Canine_Collaborator": [{"year": 2026, "with_vat": 9.91, "no_vat": 8.19}]}}, "TRANSFERENCIA DE PROPIEDAD REGISTRO INICIAL O PERRO IMPORTADO": {"category": "OTRAS INSCRIPCIONES", "prices": {"Member": [{"year": 2025, "with_vat": 13.2, "no_vat": 10.91}, {"year": 2026, "with_vat": 15.0, "no_vat": 12.4}], "User": [{"year": 2025, "with_vat": 16.0, "no_vat": 13.22}, {"year": 2026, "with_vat": 19.0, "no_vat": 15.7}], "Canine_Collaborator": [{"year": 2025, "with_vat": 9.33, "no_vat": 7.71}, {"year": 2026, "with_vat": 11.08, "no_vat": 9.16}]}}}, "years": [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]};


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