import { createFileRoute } from "@tanstack/react-router";
import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Bar,
  BarChart,
  Cell,
  LabelList,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  SlidersHorizontal,
  Upload,
  Download,
  Copy as CopyIcon,
  AlertTriangle,
  Flame,
  Filter,
  Clock,
  MapPin,
  X,
  Search,
  Maximize2,
  Crosshair,
  Zap,
  TrendingUp,
  Route as RouteIcon,
  Flag,
  StickyNote,
  Plus,
} from "lucide-react";
import logoHeader from "@/assets/logo-branca.png";
import { NavVoltarHome } from "@/components/nav-voltar-home";
import rawData from "@/data/backlog.json";
import elevatoriasData from "@/data/elevatorias.json";
import rawEquipeOverrides from "@/data/equipe-overrides.json";
import type { RouteStart, RouteStop } from "@/components/backlog-map";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Switch } from "@/components/ui/switch";
import { supabase } from "@/lib/supabase";
import { useAuth } from "@/lib/auth";
import { toast } from "sonner";
import {
  computeEquipe,
  computeResponsabilidade,
  type Equipe,
  type Responsabilidade,
} from "@/data/responsabilidade-rules";

// Leaflet importa `window` no topo → carrega só no cliente para evitar crash de SSR.
const BacklogMap = lazy(() => import("@/components/backlog-map"));

export const Route = createFileRoute("/backlog")({
  head: () => ({
    meta: [
      { title: "Eletromecânica · Backlog BI" },
      {
        name: "description",
        content: "Backlog de Ordens de Manutenção (Field/SAP) com SLA, mapa e programação semanal.",
      },
    ],
  }),
  component: BacklogPage,
});

type Row = {
  "Ordem de Manutenção": string | null;
  NOTA: string | null;
  "Status da Atividade": string | null;
  "Início do SLA": string | null;
  "Fim do SLA": string | null;
  "TEXTO BREVE": string | null;
  "TEXTO LONGO": string | null;
  PLANTA: string | null;
  Endereço: string | null;
  BAIRRO: string | null;
  Cidade: string | null;
  Estado: string | null;
  "Coordenada Y": string | null;
  "Coordenada X": string | null;
  "Bucket de Origem da OS": string | null;
  PRIORIDADE: string | null;
  "DESCRIÇÃO EQUIPAMENTO": string | null;
  "Tipo de Atividade": string | null;
};

const DATA = rawData as unknown as Row[];

type BacklogObservacao = {
  id: number;
  om: string;
  texto: string;
  autor_id: string | null;
  autor_nome: string | null;
  criado_em: string | null;
};
const EQUIPE_OVERRIDES = rawEquipeOverrides as Record<string, string>;
const RESP_OVERRIDES = rawEquipeOverrides as Record<string, string>;

const plantaToElevatoriaMap = new Map<string, string>();
(elevatoriasData as Array<{ PLANTA: string | null; ELEVATORIAS: string | null }>).forEach(
  (item) => {
    if (item.PLANTA) {
      plantaToElevatoriaMap.set(item.PLANTA.trim().toUpperCase(), item.ELEVATORIAS || "");
    }
  },
);

export function getElevatoriaName(plantaStr: string | null | undefined): string {
  if (!plantaStr) return "—";
  const code = plantaStr.split(" - ")[0].trim().toUpperCase();
  const found = plantaToElevatoriaMap.get(code);
  if (found) return found;

  const parts = plantaStr.split(" - ");
  if (parts.length > 1) {
    return parts[1].replace(/^(EAT|ETE|EAB|EEA|ETA)\s+/i, "").trim();
  }
  return "—";
}

export function abbreviateAtividade(name: string | null | undefined): string {
  if (!name) return "";
  let s = String(name).trim();
  s = s.replace(/MANUTENÇÃO PREVENTIVA/gi, "M. Preventiva");
  s = s.replace(/MANUTENÇÃO PREDITIVA/gi, "M. Preditiva");
  s = s.replace(/MANUTENÇÃO CORRETIVA/gi, "M. Corretiva");
  s = s.replace(/EMERGENCIAL/gi, "Emerg.");
  s = s.replace(/PROGRAMADA/gi, "Prog.");
  s = s.replace(/FREQUENCIA/gi, "Frequência");
  s = s.replace(/FREQUÊNCIA/gi, "Frequência");
  s = s.replace(/ENGENHARIA DE MANUTENÇÃO/gi, "Eng. Manutenção");
  return s;
}

const STORAGE_KEY = "backlog_data_v1";
const VIEW_STORAGE_KEY = "backlog_saved_views_v1";

const BLUE = "#1f7ad6";
const BLUE_DARK = "#0b3a73";

// Paleta para o pie de "Tipo de Atividade" — cores nitidamente distintas.
const PIE_COLORS = [
  "#0b3a73", // azul escuro
  "#1f7ad6", // azul médio (marca)
  "#22c1c3", // ciano
  "#8b5cf6", // roxo
  "#10b981", // verde-azulado
  "#f59e0b", // âmbar
  "#ef4444", // vermelho
  "#64748b", // cinza-azulado
  "#ec4899", // rosa
  "#0ea5e9", // azul céu
];

// ---------- parsers ----------

// "DD/MM/YY HH:MM" ou "DD/MM/YYYY HH:MM" → Date (assumido horário local).
// Detecta e corrige datas no padrão US (MM/DD/YYYY) que eventualmente
// venham do Excel como texto.
function parseFieldDate(s: string | null | undefined): Date | null {
  if (!s) return null;
  const m = String(s)
    .trim()
    .match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) {
    const iso = new Date(s);
    return isNaN(iso.getTime()) ? null : iso;
  }
  let dd = +m[1];
  let mm = +m[2];
  let yy = +m[3];
  if (yy < 100) yy += 2000;
  // Se mm > 12, a data está em formato US (MM/DD) → inverte
  if (mm > 12 && dd <= 12) {
    [dd, mm] = [mm, dd];
  }
  const hh = m[4] ? +m[4] : 0;
  const mi = m[5] ? +m[5] : 0;
  return new Date(yy, mm - 1, dd, hh, mi, 0);
}

function fmtDate(d: Date | null): string {
  if (!d) return "";
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

// Converte coordenada do Field/SAP para graus decimais.
// O export do Field traz o valor como inteiro escalado com separador de
// milhar (ex.: "-2.281.507" = -22.81507) ou já em graus ("-22.8113").
function parseCoord(raw: string | null | undefined, min: number, max: number): number | null {
  if (raw === null || raw === undefined || String(raw).trim() === "") return null;
  const s = String(raw).trim().replace(",", ".");
  const n = parseFloat(s);
  if (!Number.isFinite(n)) return null;
  if (n >= min && n <= max) return n;
  const intStr = s.replace(/\./g, "");
  if (/^-?\d+$/.test(intStr)) {
    const iv = parseInt(intStr, 10);
    for (const div of [100000, 10000, 1000, 100, 10]) {
      const v = iv / div;
      if (v >= min && v <= max) return v;
    }
  }
  return null;
}

// Latitude: -25 a -20 (RJ).
function parseLat(raw: string | null | undefined): number | null {
  return parseCoord(raw, -25, -20);
}

// Longitude: -45 a -40 (RJ).
function parseLon(raw: string | null | undefined): number | null {
  return parseCoord(raw, -45, -40);
}

// ---------- faixa ----------

type Faixa = "0-3 dias" | "4-7 dias" | "8-15 dias" | "15+ dias";
const FAIXAS: Faixa[] = ["0-3 dias", "4-7 dias", "8-15 dias", "15+ dias"];
function toFaixa(diasAberto: number): Faixa {
  if (diasAberto <= 3) return "0-3 dias";
  if (diasAberto <= 7) return "4-7 dias";
  if (diasAberto <= 15) return "8-15 dias";
  return "15+ dias";
}

// ---------- enrichment ----------

type Enriched = {
  r: Row;
  om: string;
  planta: string;
  plantaShort: string;
  fimSla: Date | null;
  inicioSla: Date | null;
  slaStatus: "ATRASADO" | "NO PRAZO" | "—";
  diasAberto: number;
  faixa: Faixa;
  responsabilidade: Responsabilidade;
  equipe: Equipe;
  lat: number | null;
  lon: number | null;
};

function enrich(rows: Row[], now: Date): Enriched[] {
  return rows.map((r) => {
    const planta = r.PLANTA || "";
    const plantaShort = planta.split(" - ")[0].trim() || planta;
    const inicioSla = parseFieldDate(r["Início do SLA"]);
    const fimSla = parseFieldDate(r["Fim do SLA"]);
    const diasAberto = inicioSla
      ? Math.max(0, Math.floor((now.getTime() - inicioSla.getTime()) / 86_400_000))
      : 0;
    const slaStatus: Enriched["slaStatus"] = fimSla
      ? now > fimSla
        ? "ATRASADO"
        : "NO PRAZO"
      : "—";
    const responsabilidade = computeResponsabilidade({
      om: r["Ordem de Manutenção"] || "",
      planta,
      textoBreve: r["TEXTO BREVE"] || "",
    });
    const equipe = computeEquipe({
      responsabilidade,
      descricaoEquipamento: r["DESCRIÇÃO EQUIPAMENTO"] || "",
      om: r["Ordem de Manutenção"] || "",
    });
    return {
      r,
      om: r["Ordem de Manutenção"] || "",
      planta,
      plantaShort,
      inicioSla,
      fimSla,
      slaStatus,
      diasAberto,
      faixa: toFaixa(diasAberto),
      responsabilidade,
      equipe,
      lat: parseLat(r["Coordenada Y"]),
      lon: parseLon(r["Coordenada X"]),
    };
  });
}

// ---------- multi-select simples ----------
function MultiSelect({
  label,
  options,
  value,
  onChange,
  hideInlineLabel,
}: {
  label: string;
  options: string[];
  value: string[];
  onChange: (v: string[]) => void;
  hideInlineLabel?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const toggle = (o: string) => {
    if (value.includes(o)) onChange(value.filter((v) => v !== o));
    else onChange([...value, o]);
  };
  return (
    <div className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex min-h-11 w-full items-center justify-between rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 py-2 text-left text-[15px] shadow-sm hover:border-[#1f7ad6] cursor-pointer"
      >
        <span className="truncate">
          {!hideInlineLabel && (
            <span className="mr-1 text-xs text-slate-500 dark:text-slate-400">{label}:</span>
          )}
          <span className="font-medium text-slate-800 dark:text-slate-100">
            {value.length === 0
              ? "Todos"
              : value.length === 1
                ? value[0]
                : `${value.length} selecionados`}
          </span>
        </span>
        <Filter className="ml-2 h-4 w-4 shrink-0 text-slate-400 dark:text-slate-400" />
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-30" onClick={() => setOpen(false)} />
          <div className="absolute z-40 mt-1 max-h-72 w-full min-w-[220px] overflow-auto rounded-md border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-1 shadow-lg">
            <div className="flex items-center justify-between border-b px-2 py-1 text-xs text-slate-500 dark:text-slate-400">
              <span>{value.length} selecionados</span>
              <button className="text-[#1f7ad6] hover:underline" onClick={() => onChange([])}>
                limpar
              </button>
            </div>
            {options.map((o) => (
              <label
                key={o}
                className="flex cursor-pointer items-center gap-2 rounded px-2 py-1 text-sm hover:bg-slate-50 dark:hover:bg-slate-700"
              >
                <input type="checkbox" checked={value.includes(o)} onChange={() => toggle(o)} />
                <span className="truncate">{o}</span>
              </label>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ---------- combobox com busca (para lista longa de plantas) ----------
function ComboboxSearch({
  label,
  options,
  value,
  onChange,
  allLabel = "Todas",
  hideInlineLabel,
}: {
  label: string;
  options: string[];
  value: string;
  onChange: (v: string) => void;
  allLabel?: string;
  hideInlineLabel?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const filtered = q ? options.filter((o) => o.toLowerCase().includes(q.toLowerCase())) : options;
  const selectedLabel = value === "TODAS" ? allLabel : value;
  return (
    <div className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex min-h-11 w-full items-center justify-between rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 py-2 text-left text-[15px] shadow-sm hover:border-[#1f7ad6] cursor-pointer"
      >
        <span className="truncate">
          {!hideInlineLabel && (
            <span className="mr-1 text-xs text-slate-500 dark:text-slate-400">{label}:</span>
          )}
          <span className="font-medium text-slate-800 dark:text-slate-100">{selectedLabel}</span>
        </span>
        <Search className="ml-2 h-4 w-4 shrink-0 text-slate-400 dark:text-slate-400" />
      </button>
      {open && (
        <>
          <div
            className="fixed inset-0 z-30"
            onClick={() => {
              setOpen(false);
              setQ("");
            }}
          />
          <div className="absolute z-40 mt-1 w-full min-w-[260px] overflow-hidden rounded-md border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 shadow-lg">
            <div className="flex items-center gap-2 border-b border-slate-100 dark:border-slate-700 px-2 py-1">
              <Search className="h-4 w-4 text-slate-400 dark:text-slate-400" />
              <input
                autoFocus
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder="Buscar…"
                className="min-h-9 w-full border-none text-[14px] outline-none"
              />
              {q && (
                <button
                  onClick={() => setQ("")}
                  className="text-slate-400 dark:text-slate-400 hover:text-slate-700 dark:text-slate-200"
                >
                  <X className="h-4 w-4" />
                </button>
              )}
            </div>
            <div className="max-h-72 overflow-auto p-1">
              <button
                onClick={() => {
                  onChange("TODAS");
                  setOpen(false);
                  setQ("");
                }}
                className={`block w-full truncate rounded px-2 py-1 text-left text-sm hover:bg-slate-50 dark:hover:bg-slate-700 ${value === "TODAS" ? "bg-[#eaf3fb] font-semibold text-[#0b3a73] dark:text-white" : ""}`}
              >
                {allLabel}
              </button>
              {filtered.length === 0 && (
                <div className="px-2 py-2 text-xs text-slate-400 dark:text-slate-400">
                  Nenhum resultado.
                </div>
              )}
              {filtered.map((o) => (
                <button
                  key={o}
                  onClick={() => {
                    onChange(o);
                    setOpen(false);
                    setQ("");
                  }}
                  className={`block w-full truncate rounded px-2 py-1 text-left text-sm hover:bg-slate-50 dark:hover:bg-slate-700 ${value === o ? "bg-[#eaf3fb] font-semibold text-[#0b3a73] dark:text-white" : ""}`}
                >
                  {o}
                </button>
              ))}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ---------- página ----------
function BacklogPage() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [data, setData] = useState<Row[]>(DATA);
  const [hasCustomData, setHasCustomData] = useState(false);
  const [now, setNow] = useState(() => new Date());
  const [autoRefresh, setAutoRefresh] = useState(true);
  const [mounted, setMounted] = useState(false);
  const [activeRouteTab, setActiveRouteTab] = useState<number>(0);
  useEffect(() => setMounted(true), []);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as Row[];
        if (Array.isArray(parsed) && parsed.length) {
          setData(parsed);
          setHasCustomData(true);
        }
      }
    } catch {
      // ignore
    }
  }, []);

  // Recalcula "agora" a cada minuto para atualizar SLA/Dias em aberto.
  useEffect(() => {
    if (!autoRefresh) return;
    const id = setInterval(() => setNow(new Date()), 60_000);
    return () => clearInterval(id);
  }, [autoRefresh]);

  const [equipeOverrides, setEquipeOverrides] = useState<Record<string, Equipe>>(() => {
    try {
      const stored = localStorage.getItem("equipeOverrides");
      if (stored) return JSON.parse(stored) as Record<string, Equipe>;
    } catch {}
    return EQUIPE_OVERRIDES as Record<string, Equipe>;
  });
  // Persiste no localStorage sempre que mudar
  useEffect(() => {
    try {
      localStorage.setItem("equipeOverrides", JSON.stringify(equipeOverrides));
    } catch {}
  }, [equipeOverrides]);
  // Carrega overrides do Supabase (persiste pra todo mundo)
  useEffect(() => {
    supabase
      .from("equipe_overrides")
      .select("*")
      .then(({ data, error }) => {
        if (error) return;
        if (data?.length) {
          const map: Record<string, Equipe> = {};
          data.forEach((row: { om: string; equipe: string }) => {
            if (row.equipe === "EMEC" || row.equipe === "Automação") {
              map[row.om] = row.equipe;
            }
          });
          setEquipeOverrides((prev) => ({ ...prev, ...map }) as Record<string, Equipe>);
        }
      });
  }, []);

  const [responsabilidadeOverrides, setResponsabilidadeOverrides] = useState<
    Record<string, Responsabilidade>
  >(() => {
    try {
      const stored = localStorage.getItem("responsabilidadeOverrides");
      if (stored) return JSON.parse(stored) as Record<string, Responsabilidade>;
    } catch {}
    return RESP_OVERRIDES as Record<string, Responsabilidade>;
  });
  useEffect(() => {
    try {
      localStorage.setItem("responsabilidadeOverrides", JSON.stringify(responsabilidadeOverrides));
    } catch {}
  }, [responsabilidadeOverrides]);
  useEffect(() => {
    supabase
      .from("responsabilidade_overrides")
      .select("*")
      .then(({ data, error }) => {
        if (error) return;
        if (data?.length) {
          const map: Record<string, Responsabilidade> = {};
          data.forEach((row: { om: string; responsabilidade: string }) => {
            const valid: Responsabilidade[] = [
              "Planta Inativa",
              "Não atendemos",
              "CDA",
              "Baixada 1",
              "Baixada 2",
              "Outra SUP",
              "Ainda não identificado",
            ];
            if (valid.includes(row.responsabilidade as Responsabilidade)) {
              map[row.om] = row.responsabilidade as Responsabilidade;
            }
          });
          setResponsabilidadeOverrides(
            (prev) => ({ ...prev, ...map }) as Record<string, Responsabilidade>,
          );
        }
      });
  }, []);

  const enriched = useMemo(
    () =>
      enrich(data, now).map((e) => {
        const eqOverride = equipeOverrides[e.om];
        const respOverride = responsabilidadeOverrides[e.om];
        let result = e;
        if (eqOverride) result = { ...result, equipe: eqOverride };
        if (respOverride) result = { ...result, responsabilidade: respOverride };
        return result;
      }),
    [data, now, equipeOverrides, responsabilidadeOverrides],
  );

  // ---------- filtros ----------
  const [fPlantas, setFPlantas] = useState<string[]>([]);
  const [fResp, setFResp] = useState<string[]>(["Baixada 2"]);
  const [fStatus, setFStatus] = useState<string>("TODOS");
  const [fEquipe, setFEquipe] = useState<string>("TODAS");
  const [fCidade, setFCidade] = useState<string>("TODAS");
  const [fFaixa, setFFaixa] = useState<string>("TODAS");
  const [fTipo, setFTipo] = useState<string>("TODOS");
  const [fSearch, setFSearch] = useState("");
  const [onlyLate, setOnlyLate] = useState(false);
  const [onlyEmerg, setOnlyEmerg] = useState(false);
  const [fSlaBefore, setFSlaBefore] = useState<string>("");
  const [showFilters, setShowFilters] = useState(false);
  const [mapOpen, setMapOpen] = useState(false);
  const [mapFitSignal, setMapFitSignal] = useState(0);
  const [tableExpanded, setTableExpanded] = useState(false);
  const [copiedOm, setCopiedOm] = useState<string | null>(null);

  // ---------- filtro em cascata: cada dropdown considera os demais filtros ----------
  type FilterKey =
    | "planta"
    | "resp"
    | "status"
    | "equipe"
    | "cidade"
    | "faixa"
    | "late"
    | "emerg"
    | "sla"
    | "tipo";
  const applyFilters = (rows: Enriched[], skip?: FilterKey) => {
    const slaLimit = fSlaBefore ? new Date(fSlaBefore + "T00:00:00") : null;
    const q = fSearch.toLowerCase().trim();
    return rows.filter((e) => {
      if (q) {
        const haystack = [
          e.om,
          e.planta,
          e.r["TEXTO BREVE"],
          e.r.Cidade,
          e.r["Status da Atividade"],
          e.responsabilidade,
          e.equipe,
          e.r["Tipo de Atividade"],
          e.r.PRIORIDADE,
          e.r.NOTA,
          e.r["DESCRIÇÃO EQUIPAMENTO"],
          e.r.Endereço,
          e.r.BAIRRO,
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(q)) return false;
      }
      if (skip !== "planta" && fPlantas.length && !fPlantas.includes(e.planta)) return false;
      if (skip !== "resp" && fResp.length && !fResp.includes(e.responsabilidade)) return false;
      if (
        skip !== "status" &&
        fStatus !== "TODOS" &&
        (e.r["Status da Atividade"] || "").trim() !== fStatus
      )
        return false;
      if (skip !== "equipe" && fEquipe !== "TODAS" && e.equipe !== fEquipe) return false;
      if (skip !== "cidade" && fCidade !== "TODAS" && e.r.Cidade !== fCidade) return false;
      if (skip !== "faixa" && fFaixa !== "TODAS" && e.faixa !== fFaixa) return false;
      if (skip !== "tipo" && fTipo !== "TODOS" && e.r["Tipo de Atividade"] !== fTipo) return false;
      if (skip !== "late" && onlyLate && e.slaStatus !== "ATRASADO") return false;
      if (skip !== "emerg" && onlyEmerg && (e.r.PRIORIDADE || "").toUpperCase() !== "EMERGÊNCIA")
        return false;
      if (skip !== "sla" && slaLimit && e.fimSla && e.fimSla >= slaLimit) return false;
      return true;
    });
  };
  const filtered = useMemo(
    () => applyFilters(enriched),
    [
      enriched,
      fSearch,
      fPlantas,
      fResp,
      fStatus,
      fEquipe,
      fCidade,
      fFaixa,
      fTipo,
      onlyLate,
      onlyEmerg,
      fSlaBefore,
    ],
  );

  const uniq = <T,>(arr: T[]) => Array.from(new Set(arr));
  const OPT_PLANTA = useMemo(
    () =>
      uniq(
        applyFilters(enriched, "planta")
          .map((e) => e.planta)
          .filter(Boolean),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fResp, fStatus, fEquipe, fCidade, fFaixa, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_RESP = useMemo(
    () => uniq(applyFilters(enriched, "resp").map((e) => e.responsabilidade)).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fStatus, fEquipe, fCidade, fFaixa, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_STATUS = useMemo(
    () =>
      uniq(
        applyFilters(enriched, "status")
          .map((e) => (e.r["Status da Atividade"] || "").trim())
          .filter(Boolean),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fResp, fEquipe, fCidade, fFaixa, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_EQUIPE = useMemo(
    () => uniq(applyFilters(enriched, "equipe").map((e) => e.equipe)).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fResp, fStatus, fCidade, fFaixa, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_CIDADE = useMemo(
    () =>
      uniq(
        applyFilters(enriched, "cidade")
          .map((e) => e.r.Cidade)
          .filter(Boolean),
      ).sort() as string[],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fResp, fStatus, fEquipe, fFaixa, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_FAIXA = useMemo(
    () => FAIXAS.filter((f) => applyFilters(enriched, "faixa").some((e) => e.faixa === f)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fResp, fStatus, fEquipe, fCidade, fTipo, onlyLate, onlyEmerg, fSlaBefore],
  );
  const OPT_TIPO = useMemo(
    () =>
      uniq(
        applyFilters(enriched, "tipo")
          .map((e) => e.r["Tipo de Atividade"])
          .filter((x): x is string => x !== null),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, fPlantas, fResp, fStatus, fEquipe, fCidade, fFaixa, onlyLate, onlyEmerg, fSlaBefore],
  );

  // ---------- KPIs ----------
  const kTotal = filtered.length;
  const kLate = filtered.filter((e) => e.slaStatus === "ATRASADO").length;
  const kEmerg = filtered.filter(
    (e) => (e.r.PRIORIDADE || "").toUpperCase() === "EMERGÊNCIA",
  ).length;
  const kPct = kTotal ? Math.round((kLate / kTotal) * 100) : 0;

  // ---------- charts data ----------
  const dataFaixa = useMemo(
    () =>
      FAIXAS.map((f) => ({
        name: f,
        value: filtered.filter((e) => e.faixa === f).length,
      })),
    [filtered],
  );

  const dataTipoAtividade = useMemo(() => {
    const m = new Map<string, number>();
    for (const e of filtered) {
      const k = e.r["Tipo de Atividade"] || "—";
      m.set(k, (m.get(k) || 0) + 1);
    }
    return Array.from(m.entries())
      .map(([name, value]) => ({ name, displayName: abbreviateAtividade(name), value }))
      .sort((a, b) => b.value - a.value);
  }, [filtered]);

  const dataCidade = useMemo(() => {
    const m = new Map<string, number>();
    for (const e of filtered) {
      const k = e.r.Cidade || "—";
      m.set(k, (m.get(k) || 0) + 1);
    }
    const sorted = Array.from(m.entries())
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
    const TOP = 10;
    if (sorted.length <= TOP) return sorted;
    const top = sorted.slice(0, TOP);
    const outras = sorted.slice(TOP).reduce((s, x) => s + x.value, 0);
    return [...top, { name: "Outras", value: outras }];
  }, [filtered]);

  // ---------- mapa: 1 marker por PLANTA única ----------
  const mapMarkers = useMemo(() => {
    const m = new Map<
      string,
      { lat: number; lon: number; planta: string; count: number; late: number; emerg: number }
    >();
    for (const e of filtered) {
      if (e.lat === null || e.lon === null) continue;
      const key = e.planta;
      const isEmerg = (e.r["Tipo de Atividade"] || "")
        .toUpperCase()
        .includes("MANUTENÇÃO CORRETIVA EMERGENCIAL");
      const cur = m.get(key);
      if (cur) {
        cur.count += 1;
        if (e.slaStatus === "ATRASADO") cur.late += 1;
        if (isEmerg) cur.emerg += 1;
      } else {
        m.set(key, {
          lat: e.lat,
          lon: e.lon,
          planta: e.planta,
          count: 1,
          late: e.slaStatus === "ATRASADO" ? 1 : 0,
          emerg: isEmerg ? 1 : 0,
        });
      }
    }
    return Array.from(m.values());
  }, [filtered]);

  const togglePlanta = (planta: string) =>
    setFPlantas((cur) =>
      cur.includes(planta) ? cur.filter((p) => p !== planta) : [...cur, planta],
    );

  // ---------- Ações recomendadas ----------
  const emergSemProgramacao = useMemo(
    () =>
      filtered.filter(
        (e) =>
          (e.r.PRIORIDADE || "").toUpperCase() === "EMERGÊNCIA" &&
          (e.r["Status da Atividade"] || "").toLowerCase().includes("pendente"),
      ).length,
    [filtered],
  );
  const topPlantasBacklog = useMemo(() => {
    const m = new Map<string, number>();
    for (const e of filtered)
      if (e.faixa === "15+ dias") m.set(e.planta, (m.get(e.planta) || 0) + 1);
    return Array.from(m.entries())
      .map(([planta, count]) => ({ planta, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
  }, [filtered]);
  const applyEmergPendente = () => {
    setOnlyEmerg(true);
    setFStatus(OPT_STATUS.find((s) => s.toLowerCase().includes("pendente")) || "TODOS");
  };

  // ---------- ordenação da tabela ----------
  type SortKey = "om" | "textoBreve" | "planta" | "inicioSla" | "tipo";
  const [sortKey, setSortKey] = useState<SortKey>("inicioSla");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const sortedRows = useMemo(() => {
    const arr = [...filtered];
    arr.sort((a, b) => {
      let av: string | number = "";
      let bv: string | number = "";
      switch (sortKey) {
        case "om":
          av = a.om;
          bv = b.om;
          break;
        case "textoBreve":
          av = a.r["TEXTO BREVE"] || "";
          bv = b.r["TEXTO BREVE"] || "";
          break;
        case "planta":
          av = a.planta;
          bv = b.planta;
          break;
        case "inicioSla":
          av = a.inicioSla ? a.inicioSla.getTime() : 0;
          bv = b.inicioSla ? b.inicioSla.getTime() : 0;
          break;
        case "tipo":
          av = a.r["Tipo de Atividade"] || "";
          bv = b.r["Tipo de Atividade"] || "";
          break;
      }
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sortKey, sortDir]);
  const toggleSort = (k: SortKey) => {
    if (sortKey === k) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else {
      setSortKey(k);
      setSortDir("asc");
    }
  };

  // ---------- observações por O.S. ----------
  const { user, profile } = useAuth();
  const [obsPorOm, setObsPorOm] = useState<Record<string, BacklogObservacao[]>>({});
  const [obsDialogOm, setObsDialogOm] = useState<string | null>(null);
  const [novoObsTexto, setNovoObsTexto] = useState("");
  const [obsEnviando, setObsEnviando] = useState(false);
  const [obsUnica, setObsUnica] = useState<Record<string, string>>({});
  const obsUnicaLoaded = useRef<Set<string>>(new Set());
  const obsUnicaTimer = useRef<Record<string, number>>({});

  useEffect(() => {
    let active = true;
    const faltantes = Array.from(new Set(sortedRows.map((e) => e.om).filter(Boolean))).filter(
      (om) => !obsUnicaLoaded.current.has(om),
    );
    if (faltantes.length === 0) return;
    const CHUNK = 200;
    const run = async () => {
      for (let i = 0; i < faltantes.length; i += CHUNK) {
        const chunk = faltantes.slice(i, i + CHUNK);
        const { data, error } = await supabase
          .from("backlog_obs_unica")
          .select("om, obs")
          .in("om", chunk);
        if (error) {
          console.warn("Falha ao carregar observações do backlog", error);
          break;
        }
        if (!active) return;
        setObsUnica((prev) => {
          const next = { ...prev };
          for (const r of data ?? []) next[r.om] = r.obs ?? "";
          return next;
        });
        for (const om of chunk) obsUnicaLoaded.current.add(om);
      }
    };
    run();
    return () => {
      active = false;
    };
  }, [sortedRows]);

  const salvarObsUnica = async (om: string, valor: string) => {
    const { error } = await supabase
      .from("backlog_obs_unica")
      .upsert({ om, obs: valor.trim() || null, atualizado_em: new Date().toISOString() });
    if (error) console.warn("Falha ao salvar observação", error);
  };

  const atualizarObsUnica = (om: string, valor: string) => {
    setObsUnica((prev) => ({ ...prev, [om]: valor }));
    if (obsUnicaTimer.current[om]) window.clearTimeout(obsUnicaTimer.current[om]);
    obsUnicaTimer.current[om] = window.setTimeout(() => salvarObsUnica(om, valor), 700);
  };

  const flushObsUnica = (om: string) => {
    if (obsUnicaTimer.current[om]) {
      window.clearTimeout(obsUnicaTimer.current[om]);
      delete obsUnicaTimer.current[om];
    }
    salvarObsUnica(om, obsUnica[om] ?? "");
  };

  useEffect(() => {
    let active = true;
    const faltantes = Array.from(new Set(sortedRows.map((e) => e.om).filter(Boolean))).filter(
      (om) => !(om in obsPorOm),
    );
    if (faltantes.length === 0) return;
    const CHUNK = 200;
    const run = async () => {
      const todas: BacklogObservacao[] = [];
      for (let i = 0; i < faltantes.length; i += CHUNK) {
        const chunk = faltantes.slice(i, i + CHUNK);
        const { data, error } = await supabase
          .from("backlog_observacoes")
          .select("*, profiles:autor_id(nome_completo)")
          .in("om", chunk);
        if (error) {
          console.warn("Falha ao carregar observações do backlog", error);
          break;
        }
        for (const r of data ?? []) {
          todas.push({
            ...(r as Omit<BacklogObservacao, "autor_nome">),
            autor_nome: (r.profiles as { nome_completo?: string } | null)?.nome_completo ?? null,
          });
        }
      }
      if (!active) return;
      setObsPorOm((prev) => {
        const next = { ...prev };
        for (const om of faltantes) if (!(om in next)) next[om] = [];
        for (const obs of todas) next[obs.om] = [...(next[obs.om] ?? []), obs];
        return next;
      });
    };
    run();
    return () => {
      active = false;
    };
  }, [sortedRows, obsPorOm]);

  const adicionarObservacao = async () => {
    if (!obsDialogOm || !novoObsTexto.trim() || obsEnviando) return;
    setObsEnviando(true);
    const { data, error } = await supabase
      .from("backlog_observacoes")
      .insert({
        om: obsDialogOm,
        texto: novoObsTexto.trim(),
        autor_id: user?.id ?? null,
      })
      .select("*, profiles:autor_id(nome_completo)")
      .single();
    if (error) {
      toast.error("Erro ao adicionar observação: " + error.message);
      setObsEnviando(false);
      return;
    }
    const criada: BacklogObservacao = {
      ...(data as Omit<BacklogObservacao, "autor_nome">),
      autor_nome: (data.profiles as { nome_completo?: string } | null)?.nome_completo ?? null,
    };
    setObsPorOm((prev) => ({
      ...prev,
      [obsDialogOm]: [criada, ...(prev[obsDialogOm] ?? [])],
    }));
    setNovoObsTexto("");
    setObsEnviando(false);
  };

  // ---------- upload ----------
  const handleUpload = async (file: File) => {
    try {
      const name = file.name.toLowerCase();
      const isCSV = name.endsWith(".csv");
      let rows: Record<string, unknown>[];

      if (isCSV) {
        // CSV: lê como texto e parseia manualmente para evitar que o xlsx
        // interprete datas BR como US e converta para serial number.
        const text = await file.text();
        const lines = text.split(/\r?\n/).filter(Boolean);
        if (lines.length < 2) {
          alert("CSV vazio ou inválido.");
          return;
        }
        // Detecta delimitador (; ou ,)
        const delim = lines[0].includes(";") ? ";" : ",";
        const parseLine = (line: string) =>
          line.split(delim).map((s) => {
            s = s.trim();
            if (s.startsWith('"') && s.endsWith('"')) s = s.slice(1, -1);
            return s;
          });
        const headers = parseLine(lines[0]);
        rows = lines.slice(1).map((line) => {
          const cols = parseLine(line);
          const r: Record<string, unknown> = {};
          headers.forEach((h, i) => {
            r[h] = cols[i] ?? null;
          });
          return r;
        });
      } else {
        // XLSX/XLS: usa a biblioteca com cellDates:false para obter valores crus
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array", cellDates: false });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null });
      }

      if (!rows.length) {
        alert("Planilha vazia ou inválida.");
        return;
      }

      // Normaliza nome da coluna para matching (ignora acentos, encoding, BOM, espaços)
      function normKey(s: string): string {
        return s
          .replace(/^\uFEFF/, "") // BOM
          .replace(/\n/g, "")
          .trim()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "") // remove acentos
          .replace(/\u00A0/g, " ") // non-breaking space → espaço
          .replace(/\s+/g, " ") // whitespace → um espaço
          .toLowerCase();
      }
      // Mapeia qualquer variação do nome da coluna para o nome padrão
      const dateFieldMap: Record<string, string> = {};
      for (const name of ["Início do SLA", "Fim do SLA"]) {
        dateFieldMap[normKey(name)] = name;
      }

      function serialToBR(v: number): string {
        const adjusted = v > 60 ? v - 1 : v;
        const d = new Date(Date.UTC(1899, 11, 30) + adjusted * 86_400_000);
        const hh = String(d.getUTCHours()).padStart(2, "0");
        const mi = String(d.getUTCMinutes()).padStart(2, "0");
        return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()} ${hh}:${mi}`;
      }

      const dateStrRe = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/;
      function normalizeDateStr(s: string): string {
        const m = s.trim().match(dateStrRe);
        if (!m) return s;
        let a = +m[1],
          b = +m[2];
        if (b > 12 && a <= 12) [a, b] = [b, a];
        const y = +m[3] < 100 ? +m[3] + 2000 : +m[3];
        const hh = s.trim().match(/(\d{1,2}):(\d{2})$/);
        const time = hh ? `${hh[1]}:${hh[2]}` : "00:00";
        return `${String(a).padStart(2, "0")}/${String(b).padStart(2, "0")}/${y} ${time}`;
      }

      const norm = rows.map((r) => {
        const out: Record<string, unknown> = {};
        for (let [k, v] of Object.entries(r)) {
          k = k.replace(/\n/g, "").trim();
          const canon = dateFieldMap[normKey(k)];
          if (canon) {
            out[canon] =
              typeof v === "number"
                ? serialToBR(v)
                : typeof v === "string"
                  ? normalizeDateStr(v)
                  : v;
          } else {
            out[k] = v;
          }
        }
        return out;
      });
      localStorage.setItem(STORAGE_KEY, JSON.stringify(norm));
      setData(norm as Row[]);
      setHasCustomData(true);
      alert(`Bucket atualizado com ${norm.length} registros.`);
    } catch (err) {
      console.error(err);
      alert("Falha ao ler o arquivo.");
    }
  };

  // ---------- exportar CSV ----------
  const exportCSV = () => {
    const headers = [
      "Ordem",
      "Nota",
      "Planta",
      "Cidade",
      "Prioridade",
      "Status",
      "Tipo",
      "Início SLA",
      "Fim SLA",
      "Dias Aberto",
      "SLA",
      "Responsabilidade",
      "Equipe",
      "TEXTO BREVE",
    ];
    const rows = sortedRows.map((e) => [
      e.om,
      e.r.NOTA ?? "",
      e.planta,
      e.r.Cidade ?? "",
      e.r.PRIORIDADE ?? "",
      e.r["Status da Atividade"] ?? "",
      e.r["Tipo de Atividade"] ?? "",
      fmtDate(e.inicioSla),
      fmtDate(e.fimSla),
      String(e.diasAberto),
      e.slaStatus,
      e.responsabilidade,
      e.equipe,
      e.r["TEXTO BREVE"] ?? "",
    ]);
    const csv = [headers, ...rows]
      .map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(";"))
      .join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `backlog-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const copyResumo = async () => {
    const txt =
      `📊 *Backlog BI*\n` +
      `Total: ${kTotal} O.S.\n` +
      `Atrasadas: ${kLate} (${kPct}%)\n` +
      `Emergenciais: ${kEmerg}\n\n` +
      `Por faixa:\n${dataFaixa.map((d) => `• ${d.name}: ${d.value}`).join("\n")}\n\n` +
      `Top cidades:\n${dataCidade
        .slice(0, 5)
        .map((d) => `• ${d.name}: ${d.value}`)
        .join("\n")}`;
    try {
      await navigator.clipboard.writeText(txt);
      alert("Resumo copiado! Cole no WhatsApp.");
    } catch {
      alert("Não foi possível copiar.");
    }
  };

  // ---------- views salvas ----------
  const saveView = () => {
    const name = prompt("Nome da view:");
    if (!name) return;
    const v = {
      name,
      fPlantas,
      fResp,
      fStatus,
      fEquipe,
      fCidade,
      fFaixa,
      onlyLate,
      onlyEmerg,
      fSlaBefore,
    };
    const raw = localStorage.getItem(VIEW_STORAGE_KEY);
    const list: Array<typeof v> = raw ? JSON.parse(raw) : [];
    const idx = list.findIndex((x) => x.name === name);
    if (idx >= 0) list[idx] = v;
    else list.push(v);
    localStorage.setItem(VIEW_STORAGE_KEY, JSON.stringify(list));
    alert("View salva.");
  };
  const [savedViews, setSavedViews] = useState<
    Array<{
      name: string;
      fPlantas: string[];
      fResp: string[];
      fStatus: string;
      fEquipe: string;
      fCidade: string;
      fFaixa: string;
      onlyLate: boolean;
      onlyEmerg: boolean;
      fSlaBefore: string;
    }>
  >([]);
  useEffect(() => {
    try {
      const raw = localStorage.getItem(VIEW_STORAGE_KEY);
      if (raw) setSavedViews(JSON.parse(raw));
    } catch {
      /* ignore */
    }
  }, [now]);
  const loadView = (n: string) => {
    const v = savedViews.find((x) => x.name === n);
    if (!v) return;
    setFPlantas(v.fPlantas);
    setFResp(v.fResp);
    setFStatus(v.fStatus);
    setFEquipe(v.fEquipe);
    setFCidade(v.fCidade);
    setFFaixa(v.fFaixa);
    setOnlyLate(v.onlyLate);
    setOnlyEmerg(v.onlyEmerg);
    setFSlaBefore(v.fSlaBefore);
  };

  const clearAllFilters = () => {
    setFSearch("");
    setFPlantas([]);
    setFResp([]);
    setFStatus("TODOS");
    setFEquipe("TODAS");
    setFCidade("TODAS");
    setFFaixa("TODAS");
    setOnlyLate(false);
    setOnlyEmerg(false);
    setFSlaBefore("");
  };

  // ============================================================
  // Route Builder — Montador Automático de Rota
  // ============================================================
  const DEFAULT_START_HINT = "PL-RJB-EAT1005";
  const AVG_KMH = 38;
  const URGENCY_WEIGHT_KM_PER_DAY = 0.35;
  const findDefaultStart = () =>
    enriched.find((e) => e.planta.toUpperCase().includes(DEFAULT_START_HINT))?.planta || "";

  const allTipos = useMemo(
    () =>
      Array.from(
        new Set(enriched.map((e) => e.r["Tipo de Atividade"] || "").filter(Boolean)),
      ).sort(),
    [enriched],
  );
  const allResps = useMemo(
    () => Array.from(new Set(enriched.map((e) => e.responsabilidade))).sort(),
    [enriched],
  );
  const allPlantas = useMemo(
    () => Array.from(new Set(enriched.map((e) => e.planta).filter(Boolean))).sort(),
    [enriched],
  );
  const allCidades = useMemo(
    () =>
      Array.from(new Set(enriched.map((e) => e.r.Cidade || "").filter(Boolean))).sort() as string[],
    [enriched],
  );

  const ROUTE_COLORS = ["#0b3a73", "#ef4444", "#10b981", "#f59e0b", "#8b5cf6"];

  const [routeDialogOpen, setRouteDialogOpen] = useState(false);
  const [rbSlaBefore, setRbSlaBefore] = useState<string>("");
  const [rbTipos, setRbTipos] = useState<string[]>([]);
  const [rbResps, setRbResps] = useState<string[]>(["Baixada 2"]);
  const [rbStart, setRbStart] = useState<string>("");
  const [rbMaxStops, setRbMaxStops] = useState<number>(20);
  const [rbTolerance, setRbTolerance] = useState<number>(3);
  const [rbElevatorias, setRbElevatorias] = useState<string[]>([]);
  const [rbCidades, setRbCidades] = useState<string[]>([]);
  const [rbRouteCount, setRbRouteCount] = useState<number>(1);
  const [routeError, setRouteError] = useState<string>("");
  const [rbEquipe, setRbEquipe] = useState<string>("TODAS");
  const [rbUseIndividualConfig, setRbUseIndividualConfig] = useState(false);
  const [rbRouteConfigs, setRbRouteConfigs] = useState<
    Array<{ equipe: string; cidades: string[]; maxStops: number }>
  >([]);

  // Sincroniza rbRouteConfigs com rbRouteCount
  useEffect(() => {
    setRbRouteConfigs((prev) => {
      const next = [...prev];
      while (next.length < rbRouteCount) {
        next.push({ equipe: "TODAS", cidades: [], maxStops: rbMaxStops });
      }
      while (next.length > rbRouteCount) {
        next.pop();
      }
      return next;
    });
  }, [rbRouteCount, rbMaxStops]);

  useEffect(() => {
    if (!rbStart && allPlantas.length) setRbStart(findDefaultStart() || allPlantas[0]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allPlantas.length]);

  // Cascade filters for Montar Rota dialog — cada dropdown só mostra opções compatíveis
  type RbFilterKey = "tipo" | "resp" | "elevatoria" | "cidade";
  const applyRbFilters = (rows: Enriched[], skip?: RbFilterKey) => {
    const slaLimit = rbSlaBefore ? new Date(rbSlaBefore) : null;
    return rows.filter((e) => {
      if (slaLimit && (!e.fimSla || e.fimSla >= slaLimit)) return false;
      if (e.lat === null || e.lon === null) return false;
      if (skip !== "tipo" && rbTipos.length && !rbTipos.includes(e.r["Tipo de Atividade"] || ""))
        return false;
      if (skip !== "resp" && rbResps.length && !rbResps.includes(e.responsabilidade)) return false;
      if (skip !== "elevatoria" && rbElevatorias.length && !rbElevatorias.includes(e.planta))
        return false;
      if (skip !== "cidade" && rbCidades.length && !rbCidades.includes(e.r.Cidade || ""))
        return false;
      return true;
    });
  };
  const rbOptTipos = useMemo(
    () =>
      Array.from(
        new Set(
          applyRbFilters(enriched, "tipo")
            .map((e) => e.r["Tipo de Atividade"] || "")
            .filter(Boolean),
        ),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, rbResps, rbElevatorias, rbCidades, rbSlaBefore],
  );
  const rbOptResps = useMemo(
    () =>
      Array.from(new Set(applyRbFilters(enriched, "resp").map((e) => e.responsabilidade))).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, rbTipos, rbElevatorias, rbCidades, rbSlaBefore],
  );
  const rbOptElevatorias = useMemo(
    () =>
      Array.from(
        new Set(
          applyRbFilters(enriched, "elevatoria")
            .map((e) => e.planta)
            .filter(Boolean),
        ),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, rbTipos, rbResps, rbCidades, rbSlaBefore],
  );
  const rbOptCidades = useMemo(
    () =>
      Array.from(
        new Set(
          applyRbFilters(enriched, "cidade")
            .map((e) => e.r.Cidade || "")
            .filter(Boolean),
        ),
      ).sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [enriched, rbTipos, rbResps, rbElevatorias, rbSlaBefore],
  );

  type GeneratedRoute = {
    start: RouteStart;
    stops: RouteStop[];
    details: Array<{
      ordem: number;
      planta: string;
      plantaShort: string;
      cidade: string;
      distKm: number;
      cumKm: number;
      oss: Enriched[];
      oldestFimSla: Date | null;
    }>;
    totalKm: number;
    etaMin: number;
    totalOs: number;
    limitConfig: { max: number; tolerance: number };
    color?: string;
  };
  const [generatedRoutes, setGeneratedRoutes] = useState<GeneratedRoute[]>([]);

  function haversineKm(a: { lat: number; lon: number }, b: { lat: number; lon: number }) {
    const R = 6371;
    const toRad = (x: number) => (x * Math.PI) / 180;
    const dLat = toRad(b.lat - a.lat);
    const dLon = toRad(b.lon - a.lon);
    const s =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLon / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(s));
  }

  // Multi-route optimization: tenta trocar paradas entre rotas pra reduzir distância total
  function optimizeRoutes(routes: GeneratedRoute[]): GeneratedRoute[] {
    if (routes.length < 2) return routes;
    let improved = true;
    let iterations = 0;
    const MAX_ITER = 20;
    while (improved && iterations < MAX_ITER) {
      improved = false;
      iterations++;
      for (let i = 0; i < routes.length; i++) {
        for (let j = i + 1; j < routes.length; j++) {
          const ri = routes[i];
          const rj = routes[j];
          if (ri.details.length < 1 || rj.details.length < 1) continue;

          const iFurthestIdx = ri.details.length - 1;
          for (let jIdx = 0; jIdx < rj.details.length; jIdx++) {
            const stopI = ri.details[iFurthestIdx];
            const stopJ = rj.details[jIdx];

            const origDistI = stopI.distKm;
            const origDistJ = stopJ.distKm;

            const prevForJinI =
              ri.details.length > 1
                ? haversineKm(
                    {
                      lat: ri.details[ri.details.length - 2].oss[0].lat!,
                      lon: ri.details[ri.details.length - 2].oss[0].lon!,
                    },
                    { lat: stopJ.oss[0].lat!, lon: stopJ.oss[0].lon! },
                  )
                : haversineKm(ri.start, { lat: stopJ.oss[0].lat!, lon: stopJ.oss[0].lon! });
            const prevForIinJ =
              jIdx > 0
                ? haversineKm(
                    {
                      lat: rj.details[jIdx - 1].oss[0].lat!,
                      lon: rj.details[jIdx - 1].oss[0].lon!,
                    },
                    { lat: stopI.oss[0].lat!, lon: stopI.oss[0].lon! },
                  )
                : haversineKm(rj.start, { lat: stopI.oss[0].lat!, lon: stopI.oss[0].lon! });
            const nextForIinJ =
              jIdx < rj.details.length - 1
                ? haversineKm(
                    { lat: stopI.oss[0].lat!, lon: stopI.oss[0].lon! },
                    {
                      lat: rj.details[jIdx + 1].oss[0].lat!,
                      lon: rj.details[jIdx + 1].oss[0].lon!,
                    },
                  )
                : 0;

            const newDistJ = prevForIinJ + nextForIinJ;
            const totalOrig = origDistI + origDistJ;
            const totalNew = prevForJinI + newDistJ;

            if (totalNew < totalOrig - 0.5) {
              // Recalculate distances properly for new route I
              const newDetailsI = ri.details.slice(0, -1).map((d) => ({ ...d }));
              const lastPrev =
                newDetailsI.length > 0
                  ? {
                      lat: newDetailsI[newDetailsI.length - 1].oss[0].lat!,
                      lon: newDetailsI[newDetailsI.length - 1].oss[0].lon!,
                    }
                  : ri.start;
              const dToJ = haversineKm(lastPrev, {
                lat: stopJ.oss[0].lat!,
                lon: stopJ.oss[0].lon!,
              });
              newDetailsI.push({
                ...stopJ,
                ordem: newDetailsI.length + 1,
                distKm: dToJ,
                cumKm:
                  (newDetailsI.length > 0 ? newDetailsI[newDetailsI.length - 1].cumKm : 0) + dToJ,
              });

              // Recalculate distances for new route J
              const newDetailsJ: GeneratedRoute["details"] = [];
              for (let k = 0; k < rj.details.length; k++) {
                if (k === jIdx) {
                  const prev =
                    newDetailsJ.length > 0
                      ? {
                          lat: newDetailsJ[newDetailsJ.length - 1].oss[0].lat!,
                          lon: newDetailsJ[newDetailsJ.length - 1].oss[0].lon!,
                        }
                      : rj.start;
                  const d = haversineKm(prev, { lat: stopI.oss[0].lat!, lon: stopI.oss[0].lon! });
                  newDetailsJ.push({
                    ...stopI,
                    ordem: k + 1,
                    distKm: d,
                    cumKm:
                      (newDetailsJ.length > 0 ? newDetailsJ[newDetailsJ.length - 1].cumKm : 0) + d,
                  });
                } else {
                  const prev =
                    newDetailsJ.length > 0
                      ? {
                          lat: newDetailsJ[newDetailsJ.length - 1].oss[0].lat!,
                          lon: newDetailsJ[newDetailsJ.length - 1].oss[0].lon!,
                        }
                      : rj.start;
                  const d = haversineKm(prev, {
                    lat: rj.details[k].oss[0].lat!,
                    lon: rj.details[k].oss[0].lon!,
                  });
                  newDetailsJ.push({
                    ...rj.details[k],
                    ordem: k + 1,
                    distKm: d,
                    cumKm:
                      (newDetailsJ.length > 0 ? newDetailsJ[newDetailsJ.length - 1].cumKm : 0) + d,
                  });
                }
              }

              const cumI = newDetailsI.reduce((s, d) => s + d.distKm, 0);
              const cumJ = newDetailsJ.reduce((s, d) => s + d.distKm, 0);

              const stopsI: RouteStop[] = newDetailsI.map((d) => ({
                planta: d.planta,
                lat: d.oss[0].lat!,
                lon: d.oss[0].lon!,
                ordem: d.ordem,
                osCount: d.oss.length,
              }));
              const stopsJ: RouteStop[] = newDetailsJ.map((d) => ({
                planta: d.planta,
                lat: d.oss[0].lat!,
                lon: d.oss[0].lon!,
                ordem: d.ordem,
                osCount: d.oss.length,
              }));

              routes[i] = {
                ...ri,
                details: newDetailsI,
                stops: stopsI,
                totalKm: cumI,
                etaMin: Math.round((cumI / AVG_KMH) * 60),
              };
              routes[j] = {
                ...rj,
                details: newDetailsJ,
                stops: stopsJ,
                totalKm: cumJ,
                etaMin: Math.round((cumJ / AVG_KMH) * 60),
              };
              improved = true;
              break;
            }
          }
          if (improved) break;
        }
        if (improved) break;
      }
    }
    return routes;
  }

  const generateRoute = () => {
    setRouteError("");
    if (!rbSlaBefore) {
      setRouteError("Informe a data/hora limite do Fim do SLA.");
      return;
    }
    if (!rbStart) {
      setRouteError("Selecione o ponto de partida.");
      return;
    }
    if (!rbMaxStops || rbMaxStops < 1) {
      setRouteError("Máximo de paradas deve ser ≥ 1.");
      return;
    }
    if (rbRouteCount < 1) {
      setRouteError("Quantidade de rotas deve ser no mínimo 1.");
      return;
    }

    const slaLimit = new Date(rbSlaBefore);
    if (isNaN(slaLimit.getTime())) {
      setRouteError("Data/hora limite inválida.");
      return;
    }

    // 1) candidatos (filtros globais)
    const candidates = enriched.filter((e) => {
      if (!e.fimSla || e.fimSla >= slaLimit) return false;
      if (rbTipos.length && !rbTipos.includes(e.r["Tipo de Atividade"] || "")) return false;
      if (rbResps.length && !rbResps.includes(e.responsabilidade)) return false;
      if (rbElevatorias.length && !rbElevatorias.includes(e.planta)) return false;
      if (!rbUseIndividualConfig) {
        if (rbCidades.length && !rbCidades.includes(e.r.Cidade || "")) return false;
        if (rbEquipe !== "TODAS" && e.equipe !== rbEquipe) return false;
      }
      if (e.lat === null || e.lon === null) return false;
      return true;
    });
    if (!candidates.length) {
      setRouteError("Nenhuma O.S. atende aos critérios informados.");
      return;
    }

    // 2) ponto de partida
    const startRow = enriched.find((e) => e.planta === rbStart && e.lat !== null && e.lon !== null);
    if (!startRow) {
      setRouteError("Ponto de partida não possui coordenadas.");
      return;
    }
    const startPt = { lat: startRow.lat!, lon: startRow.lon!, planta: rbStart };

    const routesList: GeneratedRoute[] = [];
    const visitedPlantas = new Set<string>();

    for (let rIdx = 0; rIdx < rbRouteCount; rIdx++) {
      const routeCfg = rbUseIndividualConfig ? rbRouteConfigs[rIdx] : null;
      const routeMaxStops = routeCfg?.maxStops ?? rbMaxStops;

      // Candidatos ainda não visitados
      let routeCandidates = candidates.filter((e) => !visitedPlantas.has(e.planta));

      // Filtros individuais por rota
      if (routeCfg) {
        routeCandidates = routeCandidates.filter((e) => {
          if (routeCfg.equipe !== "TODAS" && e.equipe !== routeCfg.equipe) return false;
          if (routeCfg.cidades.length && !routeCfg.cidades.includes(e.r.Cidade || "")) return false;
          return true;
        });
      }

      if (!routeCandidates.length) continue;

      const activeCandidates = routeCandidates;

      // 3) agrupa por planta
      type Group = {
        planta: string;
        lat: number;
        lon: number;
        oss: Enriched[];
        oldestFimSla: Date;
      };
      const groupMap = new Map<string, Group>();
      for (const e of activeCandidates) {
        const cur = groupMap.get(e.planta);
        if (!cur) {
          groupMap.set(e.planta, {
            planta: e.planta,
            lat: e.lat!,
            lon: e.lon!,
            oss: [e],
            oldestFimSla: e.fimSla!,
          });
        } else {
          cur.oss.push(e);
          if (e.fimSla! < cur.oldestFimSla) cur.oldestFimSla = e.fimSla!;
        }
      }
      const groups = Array.from(groupMap.values());
      if (!groups.length) break;

      // 4) Nearest-neighbor — prioridade é a distância, urgência é desempate
      const maxDaysAtraso = Math.max(
        1,
        ...groups.map((g) =>
          Math.max(0, (slaLimit.getTime() - g.oldestFimSla.getTime()) / 86_400_000),
        ),
      );
      const remaining = new Set(groups.map((g) => g.planta));
      const orderedGroups: Group[] = [];
      let cursor: { lat: number; lon: number } = startPt;

      while (remaining.size) {
        let best: Group | null = null;
        let bestScore = Infinity;
        for (const g of groups) {
          if (!remaining.has(g.planta)) continue;
          const d = haversineKm(cursor, g);
          const days = Math.max(0, (slaLimit.getTime() - g.oldestFimSla.getTime()) / 86_400_000);
          const norm = days / maxDaysAtraso;
          // Penaliza severamente saltos > 10 km (inadmissível)
          const hopPenalty = d > 10 ? (d - 10) * 100 : 0;
          // Urgência conta muito pouco no score (máximo ~0.5 km de desconto)
          const urgencyDiscount = norm * 0.5;
          const score = d + hopPenalty - urgencyDiscount;
          if (score < bestScore) {
            bestScore = score;
            best = g;
          }
        }
        if (!best) break;
        orderedGroups.push(best);
        remaining.delete(best.planta);
        cursor = { lat: best.lat, lon: best.lon };
      }

      // 5) Corte pelo máximo de paradas (com tolerância)
      const chosen: Group[] = [];
      let acc = 0;
      for (const g of orderedGroups) {
        if (acc >= routeMaxStops) break;
        const next = acc + g.oss.length;
        if (next <= routeMaxStops) {
          chosen.push(g);
          acc = next;
        } else {
          const overflow = next - routeMaxStops;
          if (overflow <= rbTolerance) {
            chosen.push(g);
            acc = next;
          }
          break;
        }
      }
      if (!chosen.length) continue;

      // 6) monta stops + distâncias
      let cumKm = 0;
      let prev: { lat: number; lon: number } = startPt;
      const details: GeneratedRoute["details"] = [];
      const stops: RouteStop[] = [];
      let totalOs = 0;

      chosen.forEach((g, idx) => {
        visitedPlantas.add(g.planta); // Marca planta como visitada nesta rota
        const d = haversineKm(prev, g);
        cumKm += d;
        const ordem = idx + 1;
        details.push({
          ordem,
          planta: g.planta,
          plantaShort: g.planta.split(" - ")[0] || g.planta,
          cidade: g.oss[0]?.r.Cidade || "",
          distKm: d,
          cumKm,
          oss: g.oss,
          oldestFimSla: g.oldestFimSla,
        });
        stops.push({ planta: g.planta, lat: g.lat, lon: g.lon, ordem, osCount: g.oss.length });
        totalOs += g.oss.length;
        prev = { lat: g.lat, lon: g.lon };
      });

      const totalKm = cumKm;
      const etaMin = Math.round((totalKm / AVG_KMH) * 60);

      routesList.push({
        start: { lat: startPt.lat, lon: startPt.lon, label: rbStart },
        stops,
        details,
        totalKm,
        etaMin,
        totalOs,
        limitConfig: { max: routeMaxStops, tolerance: rbTolerance },
        color: ROUTE_COLORS[rIdx % ROUTE_COLORS.length],
      });
    }

    if (!routesList.length) {
      setRouteError("Nenhuma rota pôde ser gerada com as O.S. restantes.");
      return;
    }

    // Otimização entre rotas: tenta trocar paradas entre rotas pra reduzir distância total
    const optimized = optimizeRoutes(routesList);

    setGeneratedRoutes(optimized);
    setRouteDialogOpen(false);
    setTimeout(() => setMapFitSignal((n) => n + 1), 100);
  };

  const clearRoute = () => setGeneratedRoutes([]);

  const exportRouteCSV = (routeIdx?: number) => {
    if (!generatedRoutes.length) return;
    const headers = [
      "Rota",
      "Parada",
      "Planta",
      "Cidade",
      "Distância (km)",
      "Acumulado (km)",
      "Ordem",
      "Nota",
      "Fim SLA",
      "Prioridade",
      "Tipo",
      "TEXTO BREVE",
    ];
    const rows: string[][] = [];

    generatedRoutes.forEach((route, rIdx) => {
      if (routeIdx !== undefined && routeIdx !== rIdx) return;
      for (const d of route.details) {
        for (const os of d.oss) {
          rows.push([
            `Rota ${rIdx + 1}`,
            String(d.ordem),
            d.planta,
            d.cidade,
            d.distKm.toFixed(2),
            d.cumKm.toFixed(2),
            os.om,
            os.r.NOTA ?? "",
            fmtDate(os.fimSla),
            os.r.PRIORIDADE ?? "",
            os.r["Tipo de Atividade"] ?? "",
            os.r["TEXTO BREVE"] ?? "",
          ]);
        }
      }
    });

    const csv = [headers, ...rows]
      .map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(";"))
      .join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `rotas-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const copyRouteResumo = async () => {
    if (!generatedRoutes.length) return;
    const lines = [
      `🗺️ *Programação de Rotas* — ${generatedRoutes.reduce((acc, r) => acc + r.totalOs, 0)} O.S. no total`,
      "",
    ];

    generatedRoutes.forEach((route, rIdx) => {
      lines.push(
        `📍 *Rota ${rIdx + 1}* — ${route.totalOs} O.S. em ${route.stops.length} paradas`,
        `📏 ${route.totalKm.toFixed(1)} km · ⏱️ ~${route.etaMin} min de deslocamento`,
        `🚩 Partida: ${route.start.label.split(" - ")[0]}`,
        ...route.details.map(
          (d) =>
            `• Parada ${d.ordem}: ${d.plantaShort} (${d.oss.length} O.S., +${d.distKm.toFixed(1)} km)`,
        ),
        "",
      );
    });

    try {
      await navigator.clipboard.writeText(lines.join("\n"));
      alert("Resumo de todas as rotas copiado! Cole no WhatsApp.");
    } catch {
      alert("Não foi possível copiar.");
    }
  };

  const mapCard = (heightPx: number, showToolbar = true) => (
    <>
      {showToolbar && (
        <div className="mb-2 flex items-center justify-between">
          <div className="text-sm font-semibold text-[#0b3a73] dark:text-white">
            <MapPin className="mr-1 inline h-4 w-4" /> Mapa de elevatórias ({mapMarkers.length})
          </div>
          <div className="flex items-center gap-1">
            {fPlantas.length > 0 && (
              <button
                onClick={() => setFPlantas([])}
                className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[11px] font-medium text-red-600 hover:bg-red-100"
              >
                Limpar {fPlantas.length} planta{fPlantas.length > 1 ? "s" : ""}
              </button>
            )}
            <button
              onClick={() => setMapFitSignal((n) => n + 1)}
              title="Centralizar nas plantas visíveis"
              className="inline-flex items-center gap-1 rounded border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 py-1 text-[11px] font-medium text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
            >
              <Crosshair className="h-3.5 w-3.5" /> Centralizar
            </button>
            <button
              onClick={() => setMapOpen(true)}
              title="Expandir mapa"
              className="inline-flex items-center gap-1 rounded border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 py-1 text-[11px] font-medium text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
            >
              <Maximize2 className="h-3.5 w-3.5" /> Expandir
            </button>
          </div>
        </div>
      )}
      <div
        style={{ height: heightPx, width: "100%" }}
        className="overflow-hidden rounded-md bg-slate-100 dark:bg-slate-800"
      >
        {mounted ? (
          <Suspense
            fallback={
              <div className="flex h-full items-center justify-center text-xs text-slate-400 dark:text-slate-400">
                Carregando mapa…
              </div>
            }
          >
            <BacklogMap
              markers={mapMarkers}
              onSelect={togglePlanta}
              selectedPlantas={fPlantas}
              fitSignal={mapFitSignal}
              route={
                generatedRoutes.length > 0
                  ? generatedRoutes[activeRouteTab] || generatedRoutes[0]
                  : undefined
              }
            />
          </Suspense>
        ) : (
          <div className="flex h-full items-center justify-center text-xs text-slate-400 dark:text-slate-400">
            Carregando mapa…
          </div>
        )}
      </div>
    </>
  );

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-900 p-3 md:p-6">
      {/* Header */}
      <div className="mb-4 overflow-hidden rounded-2xl border border-white/10 bg-gradient-to-r from-[#002d74] via-[#003087] to-[#00AEEF] p-4 shadow-[0_18px_40px_-24px_rgba(0,0,0,0.6)]">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-3">
            <div className="flex h-14 shrink-0 items-center justify-center rounded-2xl">
              <img
                src={logoHeader}
                alt="Águas do Rio - Eletromecânica"
                className="h-14 w-auto object-contain"
                loading="eager"
              />
            </div>
            <div className="min-w-0 text-white">
              <p className="truncate text-lg font-semibold">Águas do Rio</p>
              <p className="truncate text-sm text-cyan-50/90">
                Eletromecânica · Backlog e programação
              </p>
            </div>
          </div>
          <NavVoltarHome />
        </div>
      </div>

      {/* Title + actions */}
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-bold text-[#0b3a73] dark:text-white sm:text-2xl">
            Backlog BI
          </h1>
          <p className="text-xs text-slate-500 dark:text-slate-400">
            Bucket Field/SAP · {data.length} O.S. · atualizado {fmtDate(now)}
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex min-h-11 items-center gap-1 rounded-md bg-[#0b3a73] px-3 py-2 text-[13px] font-semibold text-white shadow hover:bg-[#1f7ad6]"
          >
            <Upload className="h-4 w-4" /> Importar bucket do Field
          </button>
          <button
            onClick={() => setRouteDialogOpen(true)}
            className="inline-flex min-h-11 items-center gap-1 rounded-md bg-gradient-to-r from-[#f59e0b] to-[#ef4444] px-3 py-2 text-[13px] font-semibold text-white shadow hover:opacity-95"
            title="Gerar rota otimizada"
          >
            <RouteIcon className="h-4 w-4" /> Montar Rota
          </button>
          {hasCustomData && (
            <button
              onClick={() => {
                localStorage.removeItem(STORAGE_KEY);
                setData(DATA);
                setHasCustomData(false);
              }}
              className="rounded border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 py-2 text-xs text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700"
            >
              restaurar exemplo
            </button>
          )}
          <button
            onClick={exportCSV}
            className="inline-flex min-h-11 items-center gap-1 rounded-md border border-[#1f7ad6] bg-white dark:bg-slate-800 px-3 py-2 text-[13px] font-semibold text-[#0b3a73] dark:text-white hover:bg-[#eaf3fb]"
          >
            <Download className="h-4 w-4" /> Exportar
          </button>
          <button
            onClick={copyResumo}
            className="inline-flex min-h-11 items-center gap-1 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 py-2 text-[13px] font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
          >
            <CopyIcon className="h-4 w-4" /> Copiar resumo
          </button>
          <button
            onClick={saveView}
            className="inline-flex min-h-11 items-center gap-1 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 py-2 text-[13px] font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
          >
            Salvar view
          </button>
          <label className="inline-flex items-center gap-2 text-xs text-slate-600 dark:text-slate-300">
            <input
              type="checkbox"
              checked={autoRefresh}
              onChange={(e) => setAutoRefresh(e.target.checked)}
            />
            Auto-refresh SLA
          </label>
        </div>
      </div>

      {/* Barra de busca */}
      <div className="mb-4">
        <div className="relative">
          <Search className="pointer-events-none absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400 dark:text-slate-400" />
          <input
            type="text"
            value={fSearch}
            onChange={(e) => setFSearch(e.target.value)}
            placeholder="Buscar por O.S., planta, cidade, texto breve, responsabilidade, equipe…"
            className="w-full rounded-xl border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 py-3.5 pl-12 pr-10 text-[15px] shadow-sm placeholder:text-slate-400 focus:border-[#1f7ad6] focus:outline-none dark:text-white"
          />
          {fSearch && (
            <button
              type="button"
              onClick={() => setFSearch("")}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 dark:text-slate-200 cursor-pointer"
            >
              <X className="h-5 w-5" />
            </button>
          )}
        </div>
      </div>

      {/* KPIs */}
      <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <button
          onClick={clearAllFilters}
          className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-4 text-left shadow-sm transition hover:-translate-y-0.5 hover:border-[#1f7ad6] hover:shadow-md"
        >
          <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            O.S. no bucket
          </div>
          <div className="mt-1 text-3xl font-bold text-[#0b3a73] dark:text-white">{kTotal}</div>
          <div className="text-[11px] text-slate-400 dark:text-slate-400">
            clique para limpar filtros
          </div>
        </button>
        <button
          onClick={() => setOnlyLate((v) => !v)}
          className={`relative rounded-xl border p-4 text-left shadow-sm transition hover:-translate-y-0.5 hover:shadow-md ${onlyLate ? "border-red-500 bg-red-50" : "border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 hover:border-red-400"}`}
        >
          <div className="flex items-center gap-1 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            <AlertTriangle className="h-3 w-3" /> O.S. Atrasadas
          </div>
          <div className="mt-1 text-3xl font-bold text-red-600">{kLate}</div>
          {kLate > 0 && (
            <span className="absolute right-3 top-3 inline-flex h-2 w-2 animate-pulse rounded-full bg-red-500" />
          )}
        </button>
        <button
          onClick={() => setOnlyEmerg((v) => !v)}
          className={`relative rounded-xl border p-4 text-left shadow-sm transition hover:-translate-y-0.5 hover:shadow-md ${onlyEmerg ? "border-orange-500 bg-orange-50" : "border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 hover:border-orange-400"}`}
        >
          <div className="flex items-center gap-1 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            <Flame className="h-3 w-3" /> Emergenciais
          </div>
          <div className="mt-1 text-3xl font-bold text-orange-600">{kEmerg}</div>
          {kEmerg > 0 && (
            <span className="absolute right-3 top-3 inline-flex h-2 w-2 animate-pulse rounded-full bg-orange-500" />
          )}
        </button>
        <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-4 shadow-sm">
          <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            % SLA Atrasado
          </div>
          <div className="mt-1 text-3xl font-bold text-[#1f7ad6]">{kPct}%</div>
          <div className="mt-2 h-2 w-full rounded-full bg-slate-100 dark:bg-slate-700">
            <div
              className="h-full rounded-full bg-gradient-to-r from-[#1f7ad6] to-red-500"
              style={{ width: `${kPct}%` }}
            />
          </div>
        </div>
      </div>

      {/* Filtros */}
      <div className="mb-4 rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-4 shadow-sm">
        <button
          type="button"
          onClick={() => setShowFilters((v) => !v)}
          className="mb-2 flex w-full items-center justify-between text-sm font-semibold text-[#0b3a73] dark:text-white sm:hidden cursor-pointer"
        >
          <span className="flex items-center gap-2">
            <SlidersHorizontal className="h-4 w-4" /> Filtros do Backlog
          </span>
          <span>{showFilters ? "ocultar" : "mostrar"}</span>
        </button>
        <div
          className={`${showFilters ? "grid" : "hidden"} gap-4 sm:!grid sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5`}
        >
          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Elevatória
            </label>
            <MultiSelect
              label="Elevatória"
              options={OPT_PLANTA}
              value={fPlantas}
              onChange={setFPlantas}
              hideInlineLabel
            />
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Responsabilidade
            </label>
            <MultiSelect
              label="Responsabilidade"
              options={OPT_RESP}
              value={fResp}
              onChange={setFResp}
              hideInlineLabel
            />
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Status
            </label>
            <select
              value={fStatus}
              onChange={(e) => setFStatus(e.target.value)}
              className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
            >
              <option value="TODOS">Todos</option>
              {OPT_STATUS.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Equipe
            </label>
            <select
              value={fEquipe}
              onChange={(e) => setFEquipe(e.target.value)}
              className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
            >
              <option value="TODAS">Todas</option>
              {OPT_EQUIPE.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Cidade
            </label>
            <select
              value={fCidade}
              onChange={(e) => setFCidade(e.target.value)}
              className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
            >
              <option value="TODAS">Todas</option>
              {OPT_CIDADE.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Tipo de Atividade
            </label>
            <select
              value={fTipo}
              onChange={(e) => setFTipo(e.target.value)}
              className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
            >
              <option value="TODOS">Todos</option>
              {OPT_TIPO.map((p) => (
                <option key={p} value={p}>
                  {abbreviateAtividade(p)}
                </option>
              ))}
            </select>
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Tempo Aberto
            </label>
            <select
              value={fFaixa}
              onChange={(e) => setFFaixa(e.target.value)}
              className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
            >
              <option value="TODAS">Todas as faixas</option>
              {OPT_FAIXA.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </div>

          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
              Fim SLA anterior a
            </label>
            <div className="relative flex min-h-11 items-center rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 shadow-sm">
              <Clock className="h-4 w-4 text-slate-400 dark:text-slate-400 shrink-0 mr-2" />
              <input
                type="date"
                value={fSlaBefore}
                onChange={(e) => setFSlaBefore(e.target.value)}
                className="w-full border-none text-[13px] outline-none bg-transparent cursor-pointer"
              />
            </div>
          </div>

          {savedViews.length > 0 && (
            <div className="flex flex-col gap-1.5">
              <label className="text-xs font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
                Views Salvas
              </label>
              <select
                onChange={(e) => e.target.value && loadView(e.target.value)}
                className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 text-[14px] shadow-sm focus:border-[#1f7ad6] focus:outline-none cursor-pointer"
                defaultValue=""
              >
                <option value="">Escolha uma view…</option>
                {savedViews.map((v) => (
                  <option key={v.name} value={v.name}>
                    {v.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="flex items-end">
            <button
              type="button"
              onClick={clearAllFilters}
              className="min-h-11 w-full rounded-md border border-slate-300 dark:border-slate-600 bg-slate-50 dark:bg-slate-800 px-3 text-[13px] font-bold text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 hover:text-slate-800 dark:text-slate-100 transition cursor-pointer"
            >
              Limpar filtros
            </button>
          </div>
        </div>
        {fPlantas.length > 0 && (
          <div className="mt-3 inline-flex items-center gap-2 rounded-full bg-[#eaf3fb] px-3 py-1 text-xs text-[#0b3a73] dark:text-white">
            <MapPin className="h-3 w-3" /> {fPlantas.length} planta{fPlantas.length > 1 ? "s" : ""}{" "}
            selecionada{fPlantas.length > 1 ? "s" : ""}
            <button type="button" onClick={() => setFPlantas([])} className="cursor-pointer">
              <X className="h-3 w-3" />
            </button>
          </div>
        )}
      </div>

      {/* Charts row */}
      <div className="mb-4 grid gap-3 lg:grid-cols-3">
        {/* Faixa */}
        <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 shadow-sm">
          <div className="mb-2 text-sm font-semibold text-[#0b3a73] dark:text-white">
            Backlog por Tempo em Aberto
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={dataFaixa} layout="vertical" margin={{ left: 10, right: 30 }}>
              <XAxis type="number" hide />
              <YAxis type="category" dataKey="name" width={80} tick={{ fontSize: 12 }} />
              <Tooltip />
              <Bar
                dataKey="value"
                fill={BLUE}
                radius={[0, 4, 4, 0]}
                onClick={(data) => {
                  if (data && data.name) {
                    setFFaixa((prev) => (prev === data.name ? "TODAS" : data.name));
                  }
                }}
                className="cursor-pointer"
              >
                <LabelList
                  dataKey="value"
                  position="right"
                  style={{ fontSize: 12, fill: BLUE_DARK, fontWeight: 600 }}
                />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Pie tipo atividade */}
        <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 shadow-sm">
          <div className="mb-2 text-sm font-semibold text-[#0b3a73] dark:text-white">
            O.S. por Tipo de Atividade
          </div>
          {(() => {
            const totalTipo = dataTipoAtividade.reduce((s, d) => s + d.value, 0);
            return (
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-5">
                <div className="relative sm:col-span-2" style={{ height: 200 }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={dataTipoAtividade}
                        dataKey="value"
                        nameKey="name"
                        innerRadius={50}
                        outerRadius={80}
                        paddingAngle={1}
                        stroke="#fff"
                        strokeWidth={2}
                        startAngle={90}
                        endAngle={-270}
                        isAnimationActive={false}
                      >
                        {dataTipoAtividade.map((d, i) => (
                          <Cell
                            key={i}
                            fill={PIE_COLORS[i % PIE_COLORS.length]}
                            onClick={() => {
                              setFTipo((prev) => (prev === d.name ? "TODOS" : d.name));
                            }}
                            className="cursor-pointer"
                          />
                        ))}
                      </Pie>
                      <Tooltip
                        formatter={(v: number, n: string) => [
                          `${v} (${totalTipo ? Math.round((v / totalTipo) * 100) : 0}%)`,
                          abbreviateAtividade(n),
                        ]}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center">
                    <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-400 dark:text-slate-400">
                      Total
                    </div>
                    <div className="text-2xl font-bold text-[#0b3a73] dark:text-white">
                      {totalTipo}
                    </div>
                  </div>
                </div>
                <ul className="sm:col-span-3 grid grid-cols-1 gap-1 self-center text-[11px] xl:grid-cols-2">
                  {dataTipoAtividade.map((d, i) => {
                    const pct = totalTipo ? Math.round((d.value / totalTipo) * 100) : 0;
                    const isSelected = fTipo === d.name;
                    return (
                      <li key={d.name} className="truncate">
                        <button
                          type="button"
                          onClick={() => setFTipo((prev) => (prev === d.name ? "TODOS" : d.name))}
                          className={`flex items-center gap-2 truncate text-left w-full hover:bg-slate-50 dark:hover:bg-slate-700 p-1 rounded cursor-pointer transition-colors ${
                            isSelected
                              ? "bg-[#eaf3fb] font-semibold text-[#0b3a73] dark:text-white"
                              : "text-slate-750"
                          }`}
                          title={d.name}
                        >
                          <span
                            className="inline-block h-2.5 w-2.5 shrink-0 rounded-full"
                            style={{ background: PIE_COLORS[i % PIE_COLORS.length] }}
                          />
                          <span className="truncate">{d.displayName}</span>
                          <span className="ml-auto shrink-0 text-[10px] font-bold text-slate-500 dark:text-slate-400">
                            {d.value} ({pct}%)
                          </span>
                        </button>
                      </li>
                    );
                  })}
                </ul>
              </div>
            );
          })()}
        </div>

        {/* Cidade */}
        <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 shadow-sm">
          <div className="mb-2 text-sm font-semibold text-[#0b3a73] dark:text-white">
            Distribuição por Cidade
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={dataCidade} layout="vertical" margin={{ left: 10, right: 30 }}>
              <XAxis type="number" hide />
              <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11 }} />
              <Tooltip />
              <Bar
                dataKey="value"
                fill={BLUE_DARK}
                radius={[0, 4, 4, 0]}
                onClick={(data) => {
                  if (data && data.name) {
                    setFCidade((prev) => (prev === data.name ? "TODAS" : data.name));
                  }
                }}
                className="cursor-pointer"
              >
                <LabelList
                  dataKey="value"
                  position="right"
                  style={{ fontSize: 11, fill: BLUE_DARK, fontWeight: 600 }}
                />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Mapa + Programar */}
      <div className="mb-4 grid gap-3 lg:grid-cols-3">
        <div className="lg:col-span-2 rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 shadow-sm">
          {mapCard(240)}
        </div>

        <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 shadow-sm">
          <div className="mb-2 flex items-center gap-1 text-sm font-semibold text-[#0b3a73] dark:text-white">
            <Zap className="h-4 w-4" /> Ações recomendadas
          </div>

          <button
            onClick={applyEmergPendente}
            className="mb-3 flex w-full items-center justify-between rounded-lg border border-red-200 bg-red-50 p-3 text-left transition hover:border-red-400 hover:bg-red-100"
          >
            <div>
              <div className="text-[11px] font-semibold uppercase text-red-700">
                Emergenciais pendentes
              </div>
              <div className="text-2xl font-bold text-red-600">{emergSemProgramacao}</div>
              <div className="text-[10px] text-red-500">clique para filtrar</div>
            </div>
            <Flame className="h-8 w-8 text-red-400" />
          </button>

          <div className="mb-1 flex items-center gap-1 text-[11px] font-semibold uppercase text-slate-500 dark:text-slate-400">
            <TrendingUp className="h-3 w-3" /> Top plantas (15+ dias)
          </div>
          <ul className="space-y-1 text-xs">
            {topPlantasBacklog.length === 0 && (
              <li className="text-slate-400 dark:text-slate-400">
                Nenhuma planta com backlog crítico.
              </li>
            )}
            {topPlantasBacklog.map((p) => (
              <li key={p.planta}>
                <button
                  onClick={() => togglePlanta(p.planta)}
                  className={`flex w-full items-center justify-between rounded border p-2 text-left transition hover:border-[#1f7ad6] hover:bg-[#eaf3fb] ${fPlantas.includes(p.planta) ? "border-[#1f7ad6] bg-[#eaf3fb]" : "border-slate-100 dark:border-slate-700"}`}
                >
                  <span className="truncate font-medium text-[#0b3a73] dark:text-white">
                    {p.planta.split(" - ")[0]}
                  </span>
                  <span className="ml-2 shrink-0 rounded bg-red-100 px-1.5 py-0.5 text-[10px] font-semibold text-red-700">
                    {p.count} O.S.
                  </span>
                </button>
              </li>
            ))}
          </ul>
        </div>
      </div>

      <Dialog open={mapOpen} onOpenChange={setMapOpen}>
        <DialogContent className="max-w-6xl">
          <DialogHeader>
            <DialogTitle className="text-[#0b3a73] dark:text-white">
              <MapPin className="mr-1 inline h-4 w-4" /> Mapa de elevatórias ({mapMarkers.length})
            </DialogTitle>
          </DialogHeader>
          <div className="flex items-center justify-end gap-1 pb-2">
            {fPlantas.length > 0 && (
              <button
                onClick={() => setFPlantas([])}
                className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[11px] font-medium text-red-600 hover:bg-red-100"
              >
                Limpar {fPlantas.length} planta{fPlantas.length > 1 ? "s" : ""}
              </button>
            )}
            <button
              onClick={() => setMapFitSignal((n) => n + 1)}
              className="inline-flex items-center gap-1 rounded border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 py-1 text-[11px] font-medium text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
            >
              <Crosshair className="h-3.5 w-3.5" /> Centralizar
            </button>
          </div>
          <div
            style={{ height: "85vh", width: "100%" }}
            className="overflow-hidden rounded-md bg-slate-100 dark:bg-slate-800"
          >
            {mapOpen && mounted && (
              <Suspense fallback={null}>
                <BacklogMap
                  markers={mapMarkers}
                  onSelect={togglePlanta}
                  selectedPlantas={fPlantas}
                  fitSignal={mapFitSignal}
                  route={
                    generatedRoutes.length > 0
                      ? generatedRoutes[activeRouteTab] || generatedRoutes[0]
                      : undefined
                  }
                />
              </Suspense>
            )}
          </div>
        </DialogContent>
      </Dialog>

      {/* Route Builder — Dialog de configuração */}
      <Dialog open={routeDialogOpen} onOpenChange={setRouteDialogOpen}>
        <DialogContent className="max-w-2xl">
          <DialogHeader>
            <DialogTitle className="text-[#0b3a73] dark:text-white">
              <RouteIcon className="mr-1 inline h-4 w-4" /> Montar Rota Otimizada
            </DialogTitle>
          </DialogHeader>
          <div className="grid gap-3 py-2 text-sm">
            <div className="grid gap-2 sm:grid-cols-2">
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                  Fim SLA anterior a *
                </span>
                <input
                  type="datetime-local"
                  value={rbSlaBefore}
                  onChange={(e) => setRbSlaBefore(e.target.value)}
                  className="min-h-11 rounded-md border border-slate-300 dark:border-slate-600 px-2 text-[14px] shadow-sm"
                />
                <span className="text-[10px] text-slate-400 dark:text-slate-400">
                  critério de corte e de urgência (mais antigo = mais prioritário)
                </span>
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                  Ponto de partida *
                </span>
                <select
                  value={rbStart}
                  onChange={(e) => setRbStart(e.target.value)}
                  className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 text-[14px] shadow-sm"
                >
                  <option value="">Selecione…</option>
                  {allPlantas.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
                <span className="text-[10px] text-slate-400 dark:text-slate-400">
                  só origem do trajeto — não é atendida
                </span>
              </label>
            </div>

            <div className="grid gap-2 sm:grid-cols-2">
              <MultiSelect
                label="Tipo de Atividade *"
                options={rbOptTipos}
                value={rbTipos}
                onChange={setRbTipos}
              />
              <MultiSelect
                label="Responsabilidade *"
                options={rbOptResps}
                value={rbResps}
                onChange={setRbResps}
              />
            </div>

            <div className="grid gap-2 sm:grid-cols-3">
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                  Quantidade de rotas *
                </span>
                <input
                  type="number"
                  min={1}
                  max={10}
                  value={rbRouteCount}
                  onChange={(e) => setRbRouteCount(Math.max(1, Number(e.target.value) || 1))}
                  className="min-h-11 rounded-md border border-slate-300 dark:border-slate-600 px-2 text-[14px] shadow-sm"
                />
                <span className="text-[10px] text-slate-400 dark:text-slate-400">
                  divide as O.S. entre N rotas
                </span>
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                  Máximo de paradas (O.S.) *
                </span>
                <input
                  type="number"
                  min={1}
                  value={rbMaxStops}
                  onChange={(e) => setRbMaxStops(Number(e.target.value) || 0)}
                  className="min-h-11 rounded-md border border-slate-300 dark:border-slate-600 px-2 text-[14px] shadow-sm"
                />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                  Tolerância de estouro (O.S.)
                </span>
                <input
                  type="number"
                  min={0}
                  value={rbTolerance}
                  onChange={(e) => setRbTolerance(Number(e.target.value) || 0)}
                  className="min-h-11 rounded-md border border-slate-300 dark:border-slate-600 px-2 text-[14px] shadow-sm"
                />
                <span className="text-[10px] text-slate-400 dark:text-slate-400">
                  um grupo é incluído inteiro se estourar até isso
                </span>
              </label>
            </div>

            <div className="grid gap-2 sm:grid-cols-2">
              <MultiSelect
                label="Elevatórias (opcional)"
                options={rbOptElevatorias}
                value={rbElevatorias}
                onChange={setRbElevatorias}
              />
              {!rbUseIndividualConfig && (
                <MultiSelect
                  label="Cidade (opcional)"
                  options={rbOptCidades}
                  value={rbCidades}
                  onChange={setRbCidades}
                />
              )}
            </div>

            {!rbUseIndividualConfig && (
              <div className="grid gap-2 sm:grid-cols-2">
                <div className="flex flex-col gap-1">
                  <span className="text-xs font-semibold text-slate-600 dark:text-slate-300">
                    Equipe
                  </span>
                  <select
                    value={rbEquipe}
                    onChange={(e) => setRbEquipe(e.target.value)}
                    className="min-h-11 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 text-[14px] shadow-sm"
                  >
                    <option value="TODAS">Todas</option>
                    <option value="EMEC">EMEC</option>
                    <option value="Automação">Automação</option>
                  </select>
                </div>
              </div>
            )}

            {rbRouteCount > 1 && (
              <div className="flex items-center gap-2 pt-1">
                <Switch
                  checked={rbUseIndividualConfig}
                  onCheckedChange={setRbUseIndividualConfig}
                />
                <span className="text-sm font-medium text-slate-700 dark:text-slate-200">
                  Configuração individual por rota
                </span>
              </div>
            )}

            {rbUseIndividualConfig && rbRouteCount > 1 && (
              <div className="space-y-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 p-3">
                <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                  Configurações de cada rota
                </span>
                {rbRouteConfigs.map((cfg, idx) => (
                  <div
                    key={idx}
                    className="grid gap-3 rounded-md border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 p-3 sm:grid-cols-4"
                  >
                    <span className="text-xs font-bold text-[#0b3a73] dark:text-white flex items-center">
                      Rota {idx + 1}
                    </span>
                    <label className="flex flex-col gap-1">
                      <span className="text-[10px] font-semibold text-slate-500 dark:text-slate-400">
                        Equipe
                      </span>
                      <select
                        value={cfg.equipe}
                        onChange={(e) => {
                          const next = [...rbRouteConfigs];
                          next[idx] = { ...next[idx], equipe: e.target.value };
                          setRbRouteConfigs(next);
                        }}
                        className="min-h-9 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 text-[13px] shadow-sm"
                      >
                        <option value="TODAS">Todas</option>
                        <option value="EMEC">EMEC</option>
                        <option value="Automação">Automação</option>
                      </select>
                    </label>
                    <label className="flex flex-col gap-1">
                      <span className="text-[10px] font-semibold text-slate-500 dark:text-slate-400">
                        Cidade
                      </span>
                      <select
                        value={cfg.cidades[0] || ""}
                        onChange={(e) => {
                          const next = [...rbRouteConfigs];
                          next[idx] = {
                            ...next[idx],
                            cidades: e.target.value ? [e.target.value] : [],
                          };
                          setRbRouteConfigs(next);
                        }}
                        className="min-h-9 rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-2 text-[13px] shadow-sm"
                      >
                        <option value="">Todas</option>
                        {allCidades.map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="flex flex-col gap-1">
                      <span className="text-[10px] font-semibold text-slate-500 dark:text-slate-400">
                        Max paradas
                      </span>
                      <input
                        type="number"
                        min={1}
                        value={cfg.maxStops}
                        onChange={(e) => {
                          const next = [...rbRouteConfigs];
                          next[idx] = {
                            ...next[idx],
                            maxStops: Number(e.target.value) || 0,
                          };
                          setRbRouteConfigs(next);
                        }}
                        className="min-h-9 rounded-md border border-slate-300 dark:border-slate-600 px-2 text-[13px] shadow-sm"
                      />
                    </label>
                  </div>
                ))}
              </div>
            )}

            {routeError && (
              <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-700">
                {routeError}
              </div>
            )}

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setRouteDialogOpen(false)}
                className="rounded-md border border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-800 px-3 py-2 text-[13px] font-medium text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700"
              >
                Cancelar
              </button>
              <button
                onClick={generateRoute}
                className="inline-flex items-center gap-1 rounded-md bg-[#0b3a73] px-4 py-2 text-[13px] font-semibold text-white shadow hover:bg-[#1f7ad6]"
              >
                <RouteIcon className="h-4 w-4" /> Gerar Rota
              </button>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      {/* Resultado da rota */}
      {generatedRoutes.length > 0 && (
        <div className="mb-4 overflow-hidden rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 shadow-md">
          {/* Abas das rotas */}
          <div className="flex flex-wrap border-b border-slate-100 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 px-2 pt-2 gap-1">
            {generatedRoutes.map((route, rIdx) => {
              const isActive = activeRouteTab === rIdx;
              return (
                <button
                  key={rIdx}
                  type="button"
                  onClick={() => setActiveRouteTab(rIdx)}
                  className={`flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-t-lg transition border-b-2 -mb-[1px] cursor-pointer ${
                    isActive
                      ? "bg-white dark:bg-slate-800 border-[#0b3a73] text-[#0b3a73] dark:text-white"
                      : "text-slate-500 dark:text-slate-400 border-transparent hover:text-slate-850 hover:bg-white/40"
                  }`}
                >
                  <span
                    className="h-2 w-2 rounded-full inline-block shrink-0"
                    style={{ backgroundColor: route.color || "#0b3a73" }}
                  />
                  Rota {rIdx + 1}
                </button>
              );
            })}
            <div className="ml-auto flex flex-wrap items-center gap-2 pb-2 pr-2">
              <button
                onClick={copyRouteResumo}
                className="inline-flex items-center gap-1 rounded border border-slate-350 dark:border-slate-600 bg-white dark:bg-slate-800 px-2.5 py-1.5 text-xs font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 cursor-pointer"
              >
                <CopyIcon className="h-3.5 w-3.5" /> Copiar Todas (WhatsApp)
              </button>
              <button
                onClick={() => exportRouteCSV()}
                className="inline-flex items-center gap-1 rounded border border-slate-350 dark:border-slate-600 bg-white dark:bg-slate-800 px-2.5 py-1.5 text-xs font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 cursor-pointer"
              >
                <Download className="h-3.5 w-3.5" /> Exportar Todas (CSV)
              </button>
              <button
                onClick={clearRoute}
                className="inline-flex items-center gap-1 rounded border border-red-200 bg-red-50 px-2.5 py-1.5 text-xs font-semibold text-red-700 hover:bg-red-100 cursor-pointer"
              >
                <X className="h-3.5 w-3.5" /> Limpar
              </button>
            </div>
          </div>

          {/* Conteúdo da Rota Ativa */}
          {(() => {
            const activeRoute = generatedRoutes[activeRouteTab] || generatedRoutes[0];
            if (!activeRoute) return null;
            return (
              <div>
                <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 dark:border-slate-700 bg-[#eaf3fb]/20 px-4 py-3">
                  <div className="flex items-center gap-2 text-[#0b3a73] dark:text-white">
                    <RouteIcon
                      className="h-5 w-5"
                      style={{ color: activeRoute.color || "#0b3a73" }}
                    />
                    <div>
                      <div className="text-sm font-bold">Detalhes da Rota {activeRouteTab + 1}</div>
                      <div className="text-[11px] text-slate-600 dark:text-slate-300">
                        {activeRoute.totalOs} / {activeRoute.limitConfig.max} O.S. em{" "}
                        {activeRoute.stops.length} paradas
                        {activeRoute.totalOs > activeRoute.limitConfig.max &&
                          ` (estouro de ${activeRoute.totalOs - activeRoute.limitConfig.max}, tolerância ${activeRoute.limitConfig.tolerance})`}
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-3 text-xs">
                    <div className="rounded bg-white dark:bg-slate-800 px-2 py-1 shadow-sm border border-slate-100 dark:border-slate-700">
                      <span className="text-slate-500 dark:text-slate-400">Distância:</span>{" "}
                      <span className="font-bold text-[#0b3a73] dark:text-white">
                        {activeRoute.totalKm.toFixed(1)} km
                      </span>
                    </div>
                    <div className="rounded bg-white dark:bg-slate-800 px-2 py-1 shadow-sm border border-slate-100 dark:border-slate-700">
                      <span className="text-slate-500 dark:text-slate-400">Tempo est.:</span>{" "}
                      <span className="font-bold text-[#0b3a73] dark:text-white">
                        ~{activeRoute.etaMin} min
                      </span>
                    </div>
                    <button
                      onClick={() => exportRouteCSV(activeRouteTab)}
                      className="inline-flex items-center gap-1 rounded border border-[#0b3a73] bg-white dark:bg-slate-800 px-2.5 py-1 text-[11px] font-semibold text-[#0b3a73] dark:text-white hover:bg-[#eaf3fb] cursor-pointer"
                    >
                      <Download className="h-3 w-3" /> CSV Rota {activeRouteTab + 1}
                    </button>
                  </div>
                </div>

                <div className="grid gap-3 p-3 lg:grid-cols-2">
                  <div className="rounded border border-slate-150 dark:border-slate-700 bg-white dark:bg-slate-800 p-2 text-xs">
                    <div className="mb-2 flex items-center gap-1 font-semibold text-[#0b3a73] dark:text-white">
                      <Flag
                        className="h-3.5 w-3.5"
                        style={{ color: activeRoute.color || "#0b3a73" }}
                      />{" "}
                      Ponto de partida:{" "}
                      <span className="font-normal text-slate-700 dark:text-slate-200">
                        {activeRoute.start.label}
                      </span>
                    </div>
                    <ol className="space-y-1 max-h-[380px] overflow-auto pr-1">
                      {activeRoute.details.map((d) => (
                        <li
                          key={d.ordem}
                          className="rounded border border-slate-100 dark:border-slate-700 p-2 hover:border-[#1f7ad6]"
                        >
                          <div className="flex items-start justify-between gap-2">
                            <div className="flex items-center gap-2">
                              <span
                                className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-[11px] font-bold text-white"
                                style={{ backgroundColor: activeRoute.color || "#0b3a73" }}
                              >
                                {d.ordem}
                              </span>
                              <div>
                                <div className="font-semibold text-[#0b3a73] dark:text-white">
                                  {d.plantaShort}
                                </div>
                                <div className="text-[10px] text-slate-500 dark:text-slate-400">
                                  {d.cidade}
                                </div>
                              </div>
                            </div>
                            <div className="text-right text-[10px] text-slate-500 dark:text-slate-400">
                              <div>+{d.distKm.toFixed(1)} km</div>
                              <div className="text-slate-400 dark:text-slate-400">
                                Σ {d.cumKm.toFixed(1)} km
                              </div>
                            </div>
                            <div className="flex items-center gap-1 ml-2">
                              <button
                                onClick={() => {
                                  const routes = [...generatedRoutes];
                                  const r = { ...routes[activeRouteTab] };
                                  const idx = r.details.findIndex((x) => x.planta === d.planta);
                                  if (idx === -1) return;
                                  const newDetails = r.details.filter((_, i) => i !== idx);
                                  const newStops = r.stops.filter((_, i) => i !== idx);
                                  // Recalculate cumulative distances
                                  let cumKm = 0;
                                  const prev = activeRoute.start;
                                  const recalcDetails = newDetails.map((x, i) => {
                                    const dist =
                                      i === 0
                                        ? haversineKm(prev, {
                                            lat: x.oss[0].lat!,
                                            lon: x.oss[0].lon!,
                                          })
                                        : haversineKm(
                                            {
                                              lat: newDetails[i - 1].oss[0].lat!,
                                              lon: newDetails[i - 1].oss[0].lon!,
                                            },
                                            { lat: x.oss[0].lat!, lon: x.oss[0].lon! },
                                          );
                                    cumKm += dist;
                                    return { ...x, ordem: i + 1, distKm: dist, cumKm };
                                  });
                                  const totalKm = cumKm;
                                  routes[activeRouteTab] = {
                                    ...r,
                                    details: recalcDetails,
                                    stops: newStops.map((s, i) => ({ ...s, ordem: i + 1 })),
                                    totalKm,
                                    etaMin: Math.round((totalKm / AVG_KMH) * 60),
                                  };
                                  setGeneratedRoutes(routes);
                                }}
                                className="rounded bg-red-50 px-1 py-0.5 text-[10px] text-red-600 hover:bg-red-100 cursor-pointer"
                                title="Remover elevatória desta rota"
                              >
                                ✕
                              </button>
                              {generatedRoutes.length > 1 && (
                                <select
                                  value=""
                                  onChange={(e) => {
                                    const targetIdx = Number(e.target.value);
                                    if (isNaN(targetIdx) || targetIdx === activeRouteTab) return;
                                    const routes = [...generatedRoutes];
                                    // Remove from current route
                                    const src = { ...routes[activeRouteTab] };
                                    const idx = src.details.findIndex((x) => x.planta === d.planta);
                                    if (idx === -1) return;
                                    const movedDetail = src.details[idx];
                                    src.details = src.details.filter((_, i) => i !== idx);
                                    src.stops = src.stops.filter((_, i) => i !== idx);
                                    // Recalc source route
                                    let cumKm = 0;
                                    const prev = activeRoute.start;
                                    src.details = src.details.map((x, i) => {
                                      const dist =
                                        i === 0
                                          ? haversineKm(prev, {
                                              lat: x.oss[0].lat!,
                                              lon: x.oss[0].lon!,
                                            })
                                          : haversineKm(
                                              {
                                                lat: src.details[i - 1].oss[0].lat!,
                                                lon: src.details[i - 1].oss[0].lon!,
                                              },
                                              { lat: x.oss[0].lat!, lon: x.oss[0].lon! },
                                            );
                                      cumKm += dist;
                                      return { ...x, ordem: i + 1, distKm: dist, cumKm };
                                    });
                                    src.totalKm = cumKm;
                                    src.etaMin = Math.round((cumKm / AVG_KMH) * 60);
                                    src.stops = src.stops.map((s, i) => ({ ...s, ordem: i + 1 }));

                                    // Add to target route at the end
                                    const tgt = { ...routes[targetIdx] };
                                    const tgtPrev =
                                      tgt.details.length > 0
                                        ? {
                                            lat: tgt.details[tgt.details.length - 1].oss[0].lat!,
                                            lon: tgt.details[tgt.details.length - 1].oss[0].lon!,
                                          }
                                        : tgt.start;
                                    const distToNew = haversineKm(tgtPrev, {
                                      lat: movedDetail.oss[0].lat!,
                                      lon: movedDetail.oss[0].lon!,
                                    });
                                    const newDetail = {
                                      ...movedDetail,
                                      ordem: tgt.details.length + 1,
                                      distKm: distToNew,
                                      cumKm:
                                        (tgt.details.length > 0
                                          ? tgt.details[tgt.details.length - 1].cumKm
                                          : 0) + distToNew,
                                    };
                                    tgt.details = [...tgt.details, newDetail];
                                    tgt.stops = [
                                      ...tgt.stops,
                                      {
                                        planta: movedDetail.planta,
                                        lat: movedDetail.oss[0].lat!,
                                        lon: movedDetail.oss[0].lon!,
                                        ordem: tgt.stops.length + 1,
                                        osCount: movedDetail.oss.length,
                                      },
                                    ];
                                    tgt.totalKm = tgt.details.reduce((s, x) => s + x.distKm, 0);
                                    tgt.etaMin = Math.round((tgt.totalKm / AVG_KMH) * 60);

                                    routes[activeRouteTab] = src;
                                    routes[targetIdx] = tgt;
                                    setGeneratedRoutes(routes);
                                    setActiveRouteTab(targetIdx);
                                  }}
                                  className="rounded border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 px-1 py-0.5 text-[10px] text-slate-600 dark:text-slate-300 cursor-pointer"
                                  title="Transferir para outra rota"
                                >
                                  <option value="">→</option>
                                  {generatedRoutes.map((_, ri) =>
                                    ri !== activeRouteTab ? (
                                      <option key={ri} value={ri}>
                                        Rota {ri + 1}
                                      </option>
                                    ) : null,
                                  )}
                                </select>
                              )}
                            </div>
                          </div>
                          <ul className="mt-1 ml-8 space-y-0.5">
                            {d.oss.map((os) => (
                              <li
                                key={os.om}
                                className="flex items-center gap-2 text-[11px] text-slate-600 dark:text-slate-300"
                              >
                                <span className="font-mono text-[#1f7ad6]">{os.om}</span>
                                <span className="truncate">{os.r["TEXTO BREVE"]}</span>
                                {os.slaStatus === "ATRASADO" && (
                                  <span className="ml-auto shrink-0 rounded bg-red-100 px-1 text-[9px] font-semibold text-red-700">
                                    {os.diasAberto}d
                                  </span>
                                )}
                              </li>
                            ))}
                          </ul>
                        </li>
                      ))}
                    </ol>
                  </div>
                  <div className="min-h-[500px] rounded border border-slate-150 dark:border-slate-700 bg-slate-100 dark:bg-slate-700">
                    <div className="h-full min-h-[500px]" style={{ height: 500 }}>
                      {mounted ? (
                        <Suspense
                          fallback={
                            <div className="flex h-full items-center justify-center text-xs text-slate-400 dark:text-slate-400">
                              Carregando mapa…
                            </div>
                          }
                        >
                          <BacklogMap
                            markers={mapMarkers}
                            onSelect={togglePlanta}
                            selectedPlantas={fPlantas}
                            fitSignal={mapFitSignal}
                            route={activeRoute}
                          />
                        </Suspense>
                      ) : null}
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}
        </div>
      )}

      {/* Tabela */}
      <div className="rounded-xl border border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 shadow-sm">
        <div className="border-b border-slate-100 dark:border-slate-700 p-3 text-sm font-semibold text-[#0b3a73] dark:text-white flex items-center justify-between">
          <span>O.S. filtradas ({sortedRows.length})</span>
          <button
            type="button"
            onClick={() => setTableExpanded(true)}
            className="inline-flex items-center gap-1 rounded border border-slate-350 dark:border-slate-600 bg-white dark:bg-slate-800 px-2.5 py-1 text-xs font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 cursor-pointer"
            title="Expandir tabela"
          >
            <Maximize2 className="h-3.5 w-3.5" /> Expandir
          </button>
        </div>
        <div className="max-h-[500px] overflow-auto">
          <table className="min-w-[900px] w-full text-left text-[13px]">
            <thead className="sticky top-0 bg-[#eaf3fb] text-[12px] text-[#0b3a73] dark:text-white">
              <tr>
                {(
                  [
                    ["om", "Ordem"],
                    ["textoBreve", "Texto Breve"],
                    ["planta", "Planta"],
                    ["inicioSla", "Início SLA"],
                    ["tipo", "Tipo de Atividade"],
                  ] as Array<[SortKey, string]>
                ).map(([k, label]) => (
                  <th
                    key={k}
                    className="cursor-pointer whitespace-nowrap px-2 py-2 font-semibold hover:underline"
                    onClick={() => toggleSort(k)}
                  >
                    {label} {sortKey === k ? (sortDir === "asc" ? "▲" : "▼") : ""}
                  </th>
                ))}
                <th className="px-2 py-2 font-semibold">Elevatória</th>
                <th className="px-2 py-2 font-semibold">SLA</th>
                <th className="px-2 py-2 font-semibold">Resp.</th>
                <th className="px-2 py-2 font-semibold">Equipe</th>
                <th className="px-2 py-2 font-semibold">Observação</th>
                <th className="px-2 py-2 font-semibold">Comentários</th>
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((e, i) => (
                <tr
                  key={`${e.om}-${i}`}
                  className="border-b border-slate-100 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700"
                >
                  <td className="whitespace-nowrap px-2 py-1 font-mono text-[12px]">
                    <button
                      type="button"
                      onClick={async () => {
                        await navigator.clipboard.writeText(e.om);
                        setCopiedOm(e.om);
                        setTimeout(() => setCopiedOm(null), 1500);
                      }}
                      className="hover:underline text-[#1f7ad6] hover:text-[#0b3a73] dark:text-white font-bold text-left cursor-pointer flex items-center gap-1.5"
                      title="Clique para copiar a O.S."
                    >
                      {e.om}
                      {copiedOm === e.om && (
                        <span className="text-[10px] bg-emerald-100 text-emerald-700 px-1 rounded animate-fade-in font-normal font-sans shrink-0">
                          Copiado!
                        </span>
                      )}
                    </button>
                  </td>
                  <td className="px-2 py-1">{e.r["TEXTO BREVE"]}</td>
                  <td className="whitespace-nowrap px-2 py-1">{e.plantaShort}</td>
                  <td className="whitespace-nowrap px-2 py-1">{fmtDate(e.inicioSla)}</td>
                  <td className="px-2 py-1">{abbreviateAtividade(e.r["Tipo de Atividade"])}</td>
                  <td className="whitespace-nowrap px-2 py-1 font-semibold text-slate-800 dark:text-slate-100">
                    {getElevatoriaName(e.planta)}
                  </td>
                  <td className="whitespace-nowrap px-2 py-1">
                    <span
                      className={`rounded px-1 text-[11px] font-semibold ${e.slaStatus === "ATRASADO" ? "bg-red-100 text-red-700" : e.slaStatus === "NO PRAZO" ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-400"}`}
                    >
                      {e.slaStatus}
                    </span>
                  </td>
                  <td className="whitespace-nowrap px-2 py-1 text-[12px]">
                    <span className="inline-flex items-center gap-1">
                      <select
                        value={e.responsabilidade}
                        onChange={(ev) => {
                          const val = ev.target.value;
                          if (!val) return;
                          setResponsabilidadeOverrides((prev) => ({
                            ...prev,
                            [e.om]: val as Responsabilidade,
                          }));
                          supabase
                            .from("responsabilidade_overrides")
                            .upsert(
                              { om: e.om, responsabilidade: val },
                              { ignoreDuplicates: false },
                            )
                            .then(
                              ({ error }) => error && console.warn("Falha ao salvar resp", error),
                            );
                        }}
                        className="min-h-7 rounded border px-1.5 py-0.5 text-[11px] font-medium shadow-sm cursor-pointer border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 text-slate-700 dark:text-slate-200"
                      >
                        {(
                          [
                            "Planta Inativa",
                            "Não atendemos",
                            "CDA",
                            "Baixada 1",
                            "Baixada 2",
                            "Outra SUP",
                            "Ainda não identificado",
                          ] as const
                        ).map((opt) => (
                          <option key={opt} value={opt}>
                            {opt}
                          </option>
                        ))}
                      </select>
                      {responsabilidadeOverrides[e.om] && (
                        <button
                          onClick={() => {
                            setResponsabilidadeOverrides((prev) => {
                              const next = { ...prev };
                              delete next[e.om];
                              return next;
                            });
                            supabase
                              .from("responsabilidade_overrides")
                              .delete()
                              .eq("om", e.om)
                              .then(
                                ({ error }) =>
                                  error && console.warn("Falha ao remover resp override", error),
                              );
                          }}
                          className="rounded bg-slate-100 dark:bg-slate-700 px-1 py-0.5 text-[10px] text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-600 cursor-pointer"
                          title="Reverter ao cálculo automático"
                        >
                          ↺
                        </button>
                      )}
                    </span>
                  </td>
                  <td className="whitespace-nowrap px-2 py-1 text-[12px]">
                    {e.responsabilidade === "Baixada 2" ? (
                      <span className="inline-flex items-center gap-1">
                        <select
                          value={e.equipe === "Não analisado" ? "" : e.equipe}
                          onChange={(ev) => {
                            const val = ev.target.value;
                            if (!val) return;
                            setEquipeOverrides((prev) => ({ ...prev, [e.om]: val as Equipe }));
                            supabase
                              .from("equipe_overrides")
                              .upsert({ om: e.om, equipe: val }, { ignoreDuplicates: false })
                              .then(({ error }) => error && console.warn("Falha ao salvar", error));
                          }}
                          className={`min-h-7 rounded border px-1.5 py-0.5 text-[11px] font-medium shadow-sm cursor-pointer ${
                            e.equipe === "EMEC"
                              ? "border-blue-200 bg-blue-50 text-blue-700"
                              : e.equipe === "Automação"
                                ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                                : "border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 text-slate-400 dark:text-slate-400"
                          }`}
                        >
                          <option value="" disabled>
                            {e.equipe === "Não analisado" ? "Selecionar…" : e.equipe}
                          </option>
                          <option value="EMEC">EMEC</option>
                          <option value="Automação">Automação</option>
                        </select>
                        {equipeOverrides[e.om] && (
                          <button
                            onClick={() => {
                              setEquipeOverrides((prev) => {
                                const next = { ...prev };
                                delete next[e.om];
                                return next;
                              });
                              supabase
                                .from("equipe_overrides")
                                .delete()
                                .eq("om", e.om)
                                .then(
                                  ({ error }) =>
                                    error && console.warn("Falha ao remover override", error),
                                );
                            }}
                            className="rounded bg-slate-100 dark:bg-slate-700 px-1 py-0.5 text-[10px] text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-600 cursor-pointer"
                            title="Reverter ao cálculo automático"
                          >
                            ↺
                          </button>
                        )}
                      </span>
                    ) : (
                      <span className="text-slate-500 dark:text-slate-400">{e.equipe}</span>
                    )}
                  </td>
                  <td className="min-w-[160px] px-2 py-1">
                    <input
                      value={obsUnica[e.om] ?? ""}
                      onChange={(ev) => atualizarObsUnica(e.om, ev.target.value)}
                      onBlur={() => flushObsUnica(e.om)}
                      placeholder="Escreva uma observação..."
                      className="w-full rounded border border-transparent bg-transparent px-1.5 py-1 text-[12px] text-slate-700 placeholder:text-slate-400 hover:border-slate-300 focus:border-[#1f7ad6] focus:bg-white focus:outline-none dark:text-slate-200 dark:hover:border-slate-600 dark:focus:bg-slate-800"
                    />
                  </td>
                  <td className="whitespace-nowrap px-2 py-1 text-center">
                    <button
                      type="button"
                      onClick={() => setObsDialogOm(e.om)}
                      title="Ver/Adicionar observações"
                      className="relative inline-flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 bg-white text-slate-500 hover:bg-slate-100 hover:text-[#0b3a73] dark:border-slate-600 dark:bg-slate-800 dark:text-slate-400 dark:hover:bg-slate-700 cursor-pointer"
                    >
                      <StickyNote className="h-4 w-4" />
                      {(obsPorOm[e.om]?.length ?? 0) > 0 && (
                        <span className="absolute -right-1.5 -top-1.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-orange-500 px-1 text-[9px] font-bold text-white">
                          {obsPorOm[e.om]!.length}
                        </span>
                      )}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Dialog da Tabela Expandida */}
      <Dialog open={tableExpanded} onOpenChange={setTableExpanded}>
        <DialogContent className="max-w-7xl">
          <DialogHeader className="flex flex-row items-center justify-between border-b pb-2">
            <DialogTitle className="text-[#0b3a73] dark:text-white font-bold">
              Todas as O.S. Filtradas ({sortedRows.length})
            </DialogTitle>
          </DialogHeader>
          <div className="max-h-[70vh] overflow-auto">
            <table className="w-full text-left text-[13px]">
              <thead className="sticky top-0 bg-[#eaf3fb] text-[12px] text-[#0b3a73] dark:text-white z-10">
                <tr>
                  {(
                    [
                      ["om", "Ordem"],
                      ["textoBreve", "Texto Breve"],
                      ["planta", "Planta"],
                      ["inicioSla", "Início SLA"],
                      ["tipo", "Tipo de Atividade"],
                    ] as Array<[SortKey, string]>
                  ).map(([k, label]) => (
                    <th
                      key={k}
                      className="cursor-pointer whitespace-nowrap px-2 py-2 font-semibold hover:underline"
                      onClick={() => toggleSort(k)}
                    >
                      {label} {sortKey === k ? (sortDir === "asc" ? "▲" : "▼") : ""}
                    </th>
                  ))}
                  <th className="px-2 py-2 font-semibold">Elevatória</th>
                  <th className="px-2 py-2 font-semibold">SLA</th>
                  <th className="px-2 py-2 font-semibold">Resp.</th>
                  <th className="px-2 py-2 font-semibold">Equipe</th>
                  <th className="px-2 py-2 font-semibold">Observação</th>
                  <th className="px-2 py-2 font-semibold">Comentários</th>
                </tr>
              </thead>
              <tbody>
                {sortedRows.map((e, i) => (
                  <tr
                    key={`${e.om}-${i}`}
                    className="border-b border-slate-100 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700"
                  >
                    <td className="whitespace-nowrap px-2 py-1 font-mono text-[12px]">
                      <button
                        type="button"
                        onClick={async () => {
                          await navigator.clipboard.writeText(e.om);
                          setCopiedOm(e.om);
                          setTimeout(() => setCopiedOm(null), 1500);
                        }}
                        className="hover:underline text-[#1f7ad6] hover:text-[#0b3a73] dark:text-white font-bold text-left cursor-pointer flex items-center gap-1.5"
                        title="Clique para copiar a O.S."
                      >
                        {e.om}
                        {copiedOm === e.om && (
                          <span className="text-[10px] bg-emerald-100 text-emerald-700 px-1 rounded animate-fade-in font-normal font-sans shrink-0">
                            Copiado!
                          </span>
                        )}
                      </button>
                    </td>
                    <td className="px-2 py-1">{e.r["TEXTO BREVE"]}</td>
                    <td className="whitespace-nowrap px-2 py-1">{e.plantaShort}</td>
                    <td className="whitespace-nowrap px-2 py-1">{fmtDate(e.inicioSla)}</td>
                    <td className="px-2 py-1">{abbreviateAtividade(e.r["Tipo de Atividade"])}</td>
                    <td className="whitespace-nowrap px-2 py-1 font-semibold text-slate-800 dark:text-slate-100">
                      {getElevatoriaName(e.planta)}
                    </td>
                    <td className="whitespace-nowrap px-2 py-1">
                      <span
                        className={`rounded px-1 text-[11px] font-semibold ${e.slaStatus === "ATRASADO" ? "bg-red-100 text-red-700" : e.slaStatus === "NO PRAZO" ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-400"}`}
                      >
                        {e.slaStatus}
                      </span>
                    </td>
                    <td className="whitespace-nowrap px-2 py-1 text-[12px]">
                      <span className="inline-flex items-center gap-1">
                        <select
                          value={e.responsabilidade}
                          onChange={(ev) => {
                            const val = ev.target.value;
                            if (!val) return;
                            setResponsabilidadeOverrides((prev) => ({
                              ...prev,
                              [e.om]: val as Responsabilidade,
                            }));
                            supabase
                              .from("responsabilidade_overrides")
                              .upsert(
                                { om: e.om, responsabilidade: val },
                                { ignoreDuplicates: false },
                              )
                              .then(
                                ({ error }) => error && console.warn("Falha ao salvar resp", error),
                              );
                          }}
                          className="min-h-7 rounded border px-1.5 py-0.5 text-[11px] font-medium shadow-sm cursor-pointer border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 text-slate-700 dark:text-slate-200"
                        >
                          {(
                            [
                              "Planta Inativa",
                              "Não atendemos",
                              "CDA",
                              "Baixada 1",
                              "Baixada 2",
                              "Outra SUP",
                              "Ainda não identificado",
                            ] as const
                          ).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                        {responsabilidadeOverrides[e.om] && (
                          <button
                            onClick={() => {
                              setResponsabilidadeOverrides((prev) => {
                                const next = { ...prev };
                                delete next[e.om];
                                return next;
                              });
                              supabase
                                .from("responsabilidade_overrides")
                                .delete()
                                .eq("om", e.om)
                                .then(
                                  ({ error }) =>
                                    error && console.warn("Falha ao remover resp override", error),
                                );
                            }}
                            className="rounded bg-slate-100 dark:bg-slate-700 px-1 py-0.5 text-[10px] text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-600 cursor-pointer"
                            title="Reverter ao cálculo automático"
                          >
                            ↺
                          </button>
                        )}
                      </span>
                    </td>
                    <td className="whitespace-nowrap px-2 py-1 text-[12px]">
                      {e.responsabilidade === "Baixada 2" ? (
                        <span className="inline-flex items-center gap-1">
                          <select
                            value={e.equipe === "Não analisado" ? "" : e.equipe}
                            onChange={(ev) => {
                              const val = ev.target.value;
                              if (!val) return;
                              setEquipeOverrides((prev) => ({ ...prev, [e.om]: val as Equipe }));
                              supabase
                                .from("equipe_overrides")
                                .upsert({ om: e.om, equipe: val }, { ignoreDuplicates: false })
                                .then(
                                  ({ error }) => error && console.warn("Falha ao salvar", error),
                                );
                            }}
                            className={`min-h-7 rounded border px-1.5 py-0.5 text-[11px] font-medium shadow-sm cursor-pointer ${
                              e.equipe === "EMEC"
                                ? "border-blue-200 bg-blue-50 text-blue-700"
                                : e.equipe === "Automação"
                                  ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                                  : "border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800 text-slate-400 dark:text-slate-400"
                            }`}
                          >
                            <option value="" disabled>
                              {e.equipe === "Não analisado" ? "Selecionar…" : e.equipe}
                            </option>
                            <option value="EMEC">EMEC</option>
                            <option value="Automação">Automação</option>
                          </select>
                          {equipeOverrides[e.om] && (
                            <button
                              onClick={() => {
                                setEquipeOverrides((prev) => {
                                  const next = { ...prev };
                                  delete next[e.om];
                                  return next;
                                });
                                supabase
                                  .from("equipe_overrides")
                                  .delete()
                                  .eq("om", e.om)
                                  .then(
                                    ({ error }) =>
                                      error && console.warn("Falha ao remover override", error),
                                  );
                              }}
                              className="rounded bg-slate-100 dark:bg-slate-700 px-1 py-0.5 text-[10px] text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-600 cursor-pointer"
                              title="Reverter ao cálculo automático"
                            >
                              ↺
                            </button>
                          )}
                        </span>
                      ) : (
                        <span className="text-slate-500 dark:text-slate-400">{e.equipe}</span>
                      )}
                    </td>
                    <td className="min-w-[160px] px-2 py-1">
                      <input
                        value={obsUnica[e.om] ?? ""}
                        onChange={(ev) => atualizarObsUnica(e.om, ev.target.value)}
                        onBlur={() => flushObsUnica(e.om)}
                        placeholder="Escreva uma observação..."
                        className="w-full rounded border border-transparent bg-transparent px-1.5 py-1 text-[12px] text-slate-700 placeholder:text-slate-400 hover:border-slate-300 focus:border-[#1f7ad6] focus:bg-white focus:outline-none dark:text-slate-200 dark:hover:border-slate-600 dark:focus:bg-slate-800"
                      />
                    </td>
                    <td className="whitespace-nowrap px-2 py-1 text-center">
                      <button
                        type="button"
                        onClick={() => setObsDialogOm(e.om)}
                        title="Ver/Adicionar observações"
                        className="relative inline-flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 bg-white text-slate-500 hover:bg-slate-100 hover:text-[#0b3a73] dark:border-slate-600 dark:bg-slate-800 dark:text-slate-400 dark:hover:bg-slate-700 cursor-pointer"
                      >
                        <StickyNote className="h-4 w-4" />
                        {(obsPorOm[e.om]?.length ?? 0) > 0 && (
                          <span className="absolute -right-1.5 -top-1.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-orange-500 px-1 text-[9px] font-bold text-white">
                            {obsPorOm[e.om]!.length}
                          </span>
                        )}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </DialogContent>
      </Dialog>

      {/* Dialog Observações da O.S. */}
      <Dialog open={obsDialogOm !== null} onOpenChange={(o) => !o && setObsDialogOm(null)}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 font-bold text-[#0b3a73] dark:text-white">
              <StickyNote className="h-4 w-4 text-orange-500" /> Observações · {obsDialogOm}
            </DialogTitle>
          </DialogHeader>
          {obsDialogOm && (
            <div className="space-y-3 text-sm">
              <div className="max-h-[40vh] space-y-2 overflow-y-auto pr-1">
                {(obsPorOm[obsDialogOm] ?? []).length === 0 ? (
                  <p className="py-6 text-center text-sm text-slate-400">
                    Nenhuma observação para esta O.S.
                  </p>
                ) : (
                  (obsPorOm[obsDialogOm] ?? []).map((o) => (
                    <div
                      key={o.id}
                      className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 dark:border-slate-600 dark:bg-slate-700/40"
                    >
                      <p className="whitespace-pre-wrap text-[13px] text-slate-700 dark:text-slate-200">
                        {o.texto}
                      </p>
                      <p className="mt-1 text-[10px] text-slate-400">
                        {o.criado_em ? new Date(o.criado_em).toLocaleString("pt-BR") : ""}
                        {o.autor_nome ? ` · ${o.autor_nome}` : ""}
                      </p>
                    </div>
                  ))
                )}
              </div>
              <div className="flex gap-2">
                <textarea
                  value={novoObsTexto}
                  onChange={(e) => setNovoObsTexto(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && !e.shiftKey) {
                      e.preventDefault();
                      adicionarObservacao();
                    }
                  }}
                  rows={2}
                  placeholder="Digite uma observação sobre a O.S...."
                  className="flex-1 resize-none rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#1f7ad6] focus:outline-none dark:border-slate-600 dark:bg-slate-800 dark:text-slate-200"
                />
                <button
                  onClick={adicionarObservacao}
                  disabled={!novoObsTexto.trim() || obsEnviando}
                  className="inline-flex min-h-10 items-center gap-1 self-end rounded-lg bg-[#0b3a73] px-4 py-2 text-sm font-semibold text-white transition hover:bg-[#1f7ad6] disabled:opacity-50"
                >
                  <Plus className="h-4 w-4" /> Adicionar
                </button>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>

      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.xlsx,.xls"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleUpload(f);
          e.target.value = "";
        }}
      />

      <p className="mt-4 text-center text-xs text-slate-500 dark:text-slate-400">
        Águas do Rio · Eletromecânica · Backlog Field/SAP
      </p>
    </div>
  );
}
