import React, { useState, useEffect, useMemo, useRef } from "react";
import { Plus, X, Clock, Phone, Scissors, Calendar, Users, ChevronLeft, ChevronRight, ChevronDown, Trash2, Pencil, Search, CalendarDays, DollarSign, CheckCircle2, CircleDot, XCircle, Circle, AlertCircle, Undo2, TrendingUp, Download, Copy, Settings } from "lucide-react";
import * as XLSX from "xlsx";

const COLORS = {
  bg: "#efe9e4",
  surface: "#f3f0e7",
  ink: "#2a2529",
  inkSoft: "#7A6870",
  wine: "#2a2529",
  wineDark: "#1c1a1d",
  wineTint: "#E6DCCF",
  gold: "#C9932F",
  goldTint: "#F7ECD6",
  taupe: "#D9CFC2",
  taupeStrong: "#B9A99C",
  taupeLight: "#E9E1D5",
  danger: "#B0555F",
};

const cardStyle = (accent) => ({ background: COLORS.surface, border: `2px solid ${accent || COLORS.ink}`, boxShadow: `4px 4px 0px 0px ${accent || COLORS.ink}` });

const INPUT_STYLE = { borderColor: COLORS.taupeStrong, background: COLORS.bg };
const INPUT_ERROR_STYLE = { borderColor: COLORS.danger, background: "#FBEEEF" };

const SERVICIOS = [
  { id: "corte_hombre", label: "Corte Hombre", color: "#C0596A" },
  { id: "corte_mujer", label: "Corte Mujer", color: "#E0899C" },
  { id: "color", label: "Color", color: "#C9932F" },
  { id: "tratamiento", label: "Tratamiento", color: "#6E8560" },
  { id: "peinado", label: "Peinado", color: "#8574A1" },
  { id: "otro", label: "Otro", color: "#7A6870" },
];

const PRECIOS_POR_DEFECTO = { corte_hombre: 12, corte_mujer: 15, color: 20, tratamiento: 25, peinado: 15 };

const ESTADOS = [
  { id: "pendiente", label: "Pendiente", color: "#C9932F", icon: Circle },
  { id: "confirmada", label: "Confirmada", color: "#5B7DB1", icon: CircleDot },
  { id: "completada", label: "Completada", color: "#6E8560", icon: CheckCircle2 },
  { id: "cancelada", label: "Cancelada", color: COLORS.danger, icon: XCircle },
];

const servicioInfo = (id) => SERVICIOS.find((s) => s.id === id) || SERVICIOS[4];
const estadoInfo = (id) => ESTADOS.find((s) => s.id === id) || ESTADOS[0];



function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function isoFromDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function dateFromISO(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function todayISO() {
  return isoFromDate(new Date());
}

function formatFechaLarga(iso) {
  return dateFromISO(iso).toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" });
}

function formatFechaDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y}`;
}

function formatRelativo(iso) {
  const diffDias = Math.round((dateFromISO(iso) - dateFromISO(todayISO())) / 86400000);
  if (diffDias === 1) return "Mañana";
  if (diffDias > 1 && diffDias <= 6) return `En ${diffDias} días`;
  return null;
}

function formatHoraSlot(hourInt) {
  const suffix = hourInt >= 12 ? "pm" : "am";
  let h12 = hourInt % 12;
  if (h12 === 0) h12 = 12;
  return `${String(h12).padStart(2, "0")}:00 ${suffix}`;
}

function formatFechaCorta(iso) {
  return dateFromISO(iso).toLocaleDateString("es-ES", { day: "numeric", month: "short" });
}

function formatMoneda(v) {
  const n = parseFloat(v);
  if (isNaN(n)) return null;
  return `$${n.toFixed(2)}`;
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function startOfWeek(date) {
  const d = new Date(date);
  const offset = (d.getDay() + 6) % 7;
  return addDays(d, -offset);
}

function generarDiasMes(mesActual) {
  const year = mesActual.getFullYear();
  const month = mesActual.getMonth();
  const primerDia = new Date(year, month, 1);
  const ultimoDia = new Date(year, month + 1, 0);
  const startOffset = (primerDia.getDay() + 6) % 7;
  const dias = [];
  for (let i = 0; i < startOffset; i++) dias.push(null);
  for (let d = 1; d <= ultimoDia.getDate(); d++) dias.push(new Date(year, month, d));
  return dias;
}

const EMPTY_FORM = { cliente: "", telefono: "", servicio: "corte_hombre", estado: "pendiente", fecha: todayISO(), hora: "10:00", precio: "", notas: "" };
const DIAS_SEMANA = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];

export default function App() {
  const [citas, setCitas] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [vista, setVista] = useState("hoy");
  const [modoCalendario, setModoCalendario] = useState("mes");
  const [mesActual, setMesActual] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const [semanaActual, setSemanaActual] = useState(() => startOfWeek(new Date()));
  const [diaExpandidoSemana, setDiaExpandidoSemana] = useState(todayISO());
  const [diaModalMes, setDiaModalMes] = useState(null);
  const [modoIngresos, setModoIngresos] = useState("semana");
  const [diaIngresos, setDiaIngresos] = useState(() => todayISO());
  const [mesIngresos, setMesIngresos] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const [semanaIngresos, setSemanaIngresos] = useState(() => startOfWeek(new Date()));
  const [filtroServicio, setFiltroServicio] = useState("todos");
  const [busqueda, setBusqueda] = useState("");
  const [modalAbierto, setModalAbierto] = useState(false);
  const [editandoId, setEditandoId] = useState(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [errorCliente, setErrorCliente] = useState(false);
  const [formInicial, setFormInicial] = useState(EMPTY_FORM);
  const [confirmarDescartar, setConfirmarDescartar] = useState(false);
  const [mostrarSugerenciasCliente, setMostrarSugerenciasCliente] = useState(false);
  const [confirmarBorrado, setConfirmarBorrado] = useState(null);
  const [papelera, setPapelera] = useState(null);
  const toastTimerRef = useRef(null);
  const [precios, setPrecios] = useState(PRECIOS_POR_DEFECTO);
  const [preciosCargados, setPreciosCargados] = useState(false);
  const [modalPreciosAbierto, setModalPreciosAbierto] = useState(false);
  const [formPrecios, setFormPrecios] = useState(PRECIOS_POR_DEFECTO);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("citas", false);
        if (res && res.value) setCitas(JSON.parse(res.value));
      } catch (e) {}
      finally { setLoaded(true); }
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("citas", JSON.stringify(citas), false).catch(() => {});
  }, [citas, loaded]);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("precios_sugeridos", false);
        if (res && res.value) setPrecios({ ...PRECIOS_POR_DEFECTO, ...JSON.parse(res.value) });
      } catch (e) {}
      finally { setPreciosCargados(true); }
    })();
  }, []);

  useEffect(() => {
    if (!preciosCargados) return;
    window.storage.set("precios_sugeridos", JSON.stringify(precios), false).catch(() => {});
  }, [precios, preciosCargados]);

  const abrirModalPrecios = () => {
    setFormPrecios(precios);
    setModalPreciosAbierto(true);
  };

  const guardarPrecios = () => {
    setPrecios(formPrecios);
    setModalPreciosAbierto(false);
  };

  const abrirNueva = (fechaPref, horaPref) => {
    setEditandoId(null);
    setErrorCliente(false);
    const nuevoForm = { ...EMPTY_FORM, fecha: fechaPref || todayISO(), hora: horaPref || EMPTY_FORM.hora };
    setForm(nuevoForm);
    setFormInicial(nuevoForm);
    setModalAbierto(true);
  };

  const abrirEditar = (cita) => {
    setEditandoId(cita.id);
    setErrorCliente(false);
    const formCita = {
      cliente: cita.cliente, telefono: cita.telefono, servicio: cita.servicio,
      estado: cita.estado || "pendiente", fecha: cita.fecha, hora: cita.hora,
      precio: cita.precio ?? "", notas: cita.notas,
    };
    setForm(formCita);
    setFormInicial(formCita);
    setModalAbierto(true);
  };

  const duplicarCita = (cita) => {
    setEditandoId(null);
    setErrorCliente(false);
    const formDuplicado = {
      cliente: cita.cliente, telefono: cita.telefono, servicio: cita.servicio,
      estado: "pendiente", fecha: todayISO(), hora: cita.hora,
      precio: cita.precio ?? "", notas: cita.notas,
    };
    setForm(formDuplicado);
    setFormInicial(formDuplicado);
    setModalAbierto(true);
  };

  const cerrarModalDefinitivo = () => {
    setModalAbierto(false);
    setErrorCliente(false);
    setConfirmarDescartar(false);
  };

  const intentarCerrar = () => {
    const huboCambios = JSON.stringify(form) !== JSON.stringify(formInicial);
    if (huboCambios) {
      setConfirmarDescartar(true);
    } else {
      cerrarModalDefinitivo();
    }
  };

  const guardar = () => {
    if (!form.cliente.trim()) {
      setErrorCliente(true);
      return;
    }
    if (editandoId) {
      setCitas((prev) => prev.map((c) => (c.id === editandoId ? { ...c, ...form } : c)));
    } else {
      setCitas((prev) => [...prev, { id: uid(), ...form, creado: Date.now() }]);
    }
    cerrarModalDefinitivo();
  };

  const cambiarEstadoRapido = (id, estado) => {
    setCitas((prev) => prev.map((c) => (c.id === id ? { ...c, estado } : c)));
  };

  const eliminar = (id) => {
    const citaBorrada = citas.find((c) => c.id === id);
    setCitas((prev) => prev.filter((c) => c.id !== id));
    setConfirmarBorrado(null);
    if (citaBorrada) {
      setPapelera(citaBorrada);
      if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
      toastTimerRef.current = setTimeout(() => setPapelera(null), 6000);
    }
  };

  const exportarExcel = () => {
    const filas = [...citas]
      .sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora))
      .map((c) => ({
        Cliente: c.cliente,
        Teléfono: c.telefono || "",
        Servicio: servicioInfo(c.servicio).label,
        Estado: estadoInfo(c.estado).label,
        Fecha: formatFechaDMY(c.fecha),
        Hora: c.hora,
        Precio: c.precio ? parseFloat(c.precio) : "",
        Notas: c.notas || "",
      }));
    const wsCitas = XLSX.utils.json_to_sheet(filas);
    wsCitas["!cols"] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 10 }, { wch: 30 }];

    const completadas = citas.filter((c) => c.estado === "completada");
    const totalGeneral = completadas.reduce((s, c) => s + (parseFloat(c.precio) || 0), 0);

    const porMes = {};
    completadas.forEach((c) => {
      const key = c.fecha.slice(0, 7);
      if (!porMes[key]) porMes[key] = { total: 0, citas: 0 };
      porMes[key].total += parseFloat(c.precio) || 0;
      porMes[key].citas += 1;
    });
    const mesesOrdenados = Object.keys(porMes).sort().map((key) => {
      const label = dateFromISO(`${key}-01`).toLocaleDateString("es-ES", { month: "long", year: "numeric" });
      return { label: label.charAt(0).toUpperCase() + label.slice(1), ...porMes[key] };
    });

    const porServicio = {};
    completadas.forEach((c) => {
      const key = c.servicio;
      if (!porServicio[key]) porServicio[key] = { total: 0, citas: 0 };
      porServicio[key].total += parseFloat(c.precio) || 0;
      porServicio[key].citas += 1;
    });
    const serviciosOrdenados = SERVICIOS.filter((s) => porServicio[s.id]).sort((a, b) => porServicio[b.id].total - porServicio[a.id].total).map((s) => ({ label: s.label, ...porServicio[s.id] }));

    const filasIngresos = [
      ["Resumen de ingresos"],
      ["Total ingresos (citas completadas)", totalGeneral],
      ["Cantidad de citas completadas", completadas.length],
      [],
      ["Ingresos por mes"],
      ["Mes", "Total", "Citas"],
      ...mesesOrdenados.map((m) => [m.label, m.total, m.citas]),
      [],
      ["Ingresos por servicio"],
      ["Servicio", "Total", "Citas"],
      ...serviciosOrdenados.map((s) => [s.label, s.total, s.citas]),
    ];
    const wsIngresos = XLSX.utils.aoa_to_sheet(filasIngresos);
    wsIngresos["!cols"] = [{ wch: 26 }, { wch: 14 }, { wch: 10 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsIngresos, "Ingresos");
    XLSX.utils.book_append_sheet(wb, wsCitas, "Citas");
    XLSX.writeFile(wb, `citas-respaldo-${todayISO()}.xlsx`);
  };

  const deshacerEliminar = () => {
    if (!papelera) return;
    setCitas((prev) => [...prev, papelera]);
    setPapelera(null);
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
  };

  const citaEncimada = useMemo(() => {
    if (!form.fecha || !form.hora) return null;
    return citas.find((c) => c.fecha === form.fecha && c.hora === form.hora && c.id !== editandoId && c.estado !== "cancelada") || null;
  }, [citas, form.fecha, form.hora, editandoId]);

  const citasFiltradas = useMemo(() => citas.filter((c) => filtroServicio === "todos" || c.servicio === filtroServicio), [citas, filtroServicio]);

  const citasHoy = useMemo(() => citasFiltradas.filter((c) => c.fecha === todayISO()).sort((a, b) => a.hora.localeCompare(b.hora)), [citasFiltradas]);
  const citasHoyTotal = useMemo(() => citas.filter((c) => c.fecha === todayISO()), [citas]);

  const citasProximas = useMemo(() => {
    const hoy = todayISO();
    return citasFiltradas.filter((c) => c.fecha > hoy).sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
  }, [citasFiltradas]);

  const proximasPorDia = useMemo(() => {
    const grupos = {};
    citasProximas.forEach((c) => { if (!grupos[c.fecha]) grupos[c.fecha] = []; grupos[c.fecha].push(c); });
    return Object.keys(grupos).sort().map((fecha) => ({ fecha, citas: grupos[fecha] }));
  }, [citasProximas]);

  const totalMesSeleccionado = useMemo(() => {
    const prefijo = `${mesActual.getFullYear()}-${String(mesActual.getMonth() + 1).padStart(2, "0")}`;
    return citasFiltradas.filter((c) => c.fecha.startsWith(prefijo) && c.estado !== "cancelada").reduce((sum, c) => sum + (parseFloat(c.precio) || 0), 0);
  }, [citasFiltradas, mesActual]);

  const clientesAgrupados = useMemo(() => {
    const grupos = {};
    citasFiltradas.forEach((c) => {
      const key = c.cliente.trim().toLowerCase();
      if (!grupos[key]) grupos[key] = { nombre: c.cliente, telefono: c.telefono, citas: [] };
      grupos[key].citas.push(c);
    });
    let lista = Object.values(grupos).map((g) => ({ ...g, citas: g.citas.sort((a, b) => (b.fecha + b.hora).localeCompare(a.fecha + a.hora)) }));
    if (busqueda.trim()) lista = lista.filter((g) => g.nombre.toLowerCase().includes(busqueda.trim().toLowerCase()));
    return lista.sort((a, b) => a.nombre.localeCompare(b.nombre));
  }, [citasFiltradas, busqueda]);

  const clientesUnicos = useMemo(() => {
    const map = {};
    citas.forEach((c) => {
      const key = c.cliente.trim().toLowerCase();
      if (key) map[key] = { nombre: c.cliente.trim(), telefono: c.telefono };
    });
    return Object.values(map).sort((a, b) => a.nombre.localeCompare(b.nombre));
  }, [citas]);

  const sugerenciasCliente = useMemo(() => {
    const q = form.cliente.trim().toLowerCase();
    if (!q) return [];
    return clientesUnicos.filter((c) => c.nombre.toLowerCase().includes(q) && c.nombre.toLowerCase() !== q).slice(0, 5);
  }, [form.cliente, clientesUnicos]);

  // --- Ingresos ---
  const citasCompletadas = useMemo(() => citasFiltradas.filter((c) => c.estado === "completada"), [citasFiltradas]);

  const diasSemanaIngresos = useMemo(() => Array.from({ length: 7 }, (_, i) => addDays(semanaIngresos, i)), [semanaIngresos]);

  const citasEnRango = useMemo(() => {
    if (modoIngresos === "dia") {
      return citasCompletadas.filter((c) => c.fecha === diaIngresos);
    }
    if (modoIngresos === "semana") {
      const isos = diasSemanaIngresos.map(isoFromDate);
      return citasCompletadas.filter((c) => isos.includes(c.fecha));
    }
    const prefijo = `${mesIngresos.getFullYear()}-${String(mesIngresos.getMonth() + 1).padStart(2, "0")}`;
    return citasCompletadas.filter((c) => c.fecha.startsWith(prefijo));
  }, [citasCompletadas, modoIngresos, diaIngresos, diasSemanaIngresos, mesIngresos]);

  const totalPeriodo = useMemo(() => citasEnRango.reduce((s, c) => s + (parseFloat(c.precio) || 0), 0), [citasEnRango]);

  const desglosePorServicio = useMemo(() => {
    const map = {};
    citasEnRango.forEach((c) => { map[c.servicio] = (map[c.servicio] || 0) + (parseFloat(c.precio) || 0); });
    return SERVICIOS.map((s) => ({ ...s, total: map[s.id] || 0 })).filter((s) => s.total > 0).sort((a, b) => b.total - a.total);
  }, [citasEnRango]);

  const barrasIngresos = useMemo(() => {
    if (modoIngresos === "dia") {
      const base = dateFromISO(diaIngresos);
      const dias7 = Array.from({ length: 7 }, (_, i) => addDays(base, i - 6));
      return dias7.map((d) => {
        const iso = isoFromDate(d);
        const total = citasCompletadas.filter((c) => c.fecha === iso).reduce((s, c) => s + (parseFloat(c.precio) || 0), 0);
        return { label: DIAS_SEMANA[(d.getDay() + 6) % 7], total };
      });
    }
    if (modoIngresos === "semana") {
      return diasSemanaIngresos.map((d, i) => {
        const iso = isoFromDate(d);
        const total = citasEnRango.filter((c) => c.fecha === iso).reduce((s, c) => s + (parseFloat(c.precio) || 0), 0);
        return { label: DIAS_SEMANA[i], total };
      });
    }
    const dias = generarDiasMes(mesIngresos).filter(Boolean);
    const buckets = [];
    for (let i = 0; i < dias.length; i += 7) {
      const chunk = dias.slice(i, i + 7);
      const isos = chunk.map(isoFromDate);
      const total = citasEnRango.filter((c) => isos.includes(c.fecha)).reduce((s, c) => s + (parseFloat(c.precio) || 0), 0);
      buckets.push({ label: `Sem ${buckets.length + 1}`, total });
    }
    return buckets;
  }, [modoIngresos, diaIngresos, diasSemanaIngresos, mesIngresos, citasEnRango, citasCompletadas]);

  const maxBarraIngresos = Math.max(1, ...barrasIngresos.map((b) => b.total));

  const tituloPeriodoIngresos =
    modoIngresos === "dia"
      ? formatFechaCorta(diaIngresos)
      : modoIngresos === "semana"
      ? `${formatFechaCorta(isoFromDate(diasSemanaIngresos[0]))} – ${formatFechaCorta(isoFromDate(diasSemanaIngresos[6]))}`
      : mesIngresos.toLocaleDateString("es-ES", { month: "long", year: "numeric" });

  return (
    <div style={{ background: COLORS.bg, minHeight: "100%", fontFamily: "'Inter', sans-serif", color: COLORS.ink }} className="w-full">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .font-display { font-family: 'Inter', sans-serif; font-weight: 800; letter-spacing: -0.01em; }
        .font-mono { font-family: 'IBM Plex Mono', monospace; }
        .tab-underline { position: relative; }
        .tab-underline.active::after { content: ''; position: absolute; left: 0; right: 0; bottom: -9px; height: 2px; background: ${COLORS.wine}; }
        input::placeholder, textarea::placeholder { color: #B7A6AC; }
        input:focus, textarea:focus, select:focus { outline: 2px solid ${COLORS.wine}; outline-offset: 1px; }
        button:focus-visible { outline: 2px solid ${COLORS.wine}; outline-offset: 2px; }
        .icon-edit { transition: color 0.15s ease; }
        .icon-edit:hover { color: ${COLORS.gold} !important; }
        .icon-delete { transition: color 0.15s ease; }
        .icon-delete:hover { color: ${COLORS.danger} !important; }
        .icon-duplicate { transition: color 0.15s ease; }
        .icon-duplicate:hover { color: #5B7DB1 !important; }
      `}</style>

      <div className="max-w-3xl mx-auto px-5 pt-8 pb-24">
        <div className="flex items-start justify-between mb-7">
          <div>
            <div className="flex items-center gap-2 mb-1" style={{ color: COLORS.wine }}>
              <Scissors size={18} strokeWidth={2.2} />
              <span className="font-mono text-xs tracking-widest uppercase" style={{ color: COLORS.inkSoft }}>Libro de citas</span>
            </div>
            <h1 className="font-display text-3xl uppercase" style={{ color: COLORS.ink }}>Tu agenda</h1>
            <span className="inline-block font-mono text-xs mt-2 px-2.5 py-1 rounded-full" style={{ border: `1px solid ${COLORS.ink}`, color: COLORS.ink }}>
              {citasHoyTotal.length === 0 ? "Sin citas hoy" : citasHoyTotal.length === 1 ? "1 cita hoy" : `${citasHoyTotal.length} citas hoy`}
            </span>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={exportarExcel} aria-label="Exportar a Excel" className="flex items-center justify-center w-10 h-10 rounded-full shrink-0" style={{ background: COLORS.surface, border: `2px solid ${COLORS.ink}`, color: COLORS.ink }}>
              <Download size={17} />
            </button>
            <button onClick={() => abrirNueva()} className="flex items-center gap-1.5 px-4 py-2.5 rounded-full text-sm font-medium" style={{ background: COLORS.wine, color: "#fff" }}>
              <Plus size={16} strokeWidth={2.5} /> Nueva cita
            </button>
          </div>
        </div>

        <div className="mb-5 -mx-1 overflow-x-auto">
          <div className="inline-flex items-center gap-1 rounded-full p-1" style={{ border: `2px solid ${COLORS.ink}`, background: COLORS.surface }}>
            {[
              { id: "hoy", label: "Hoy", icon: Clock },
              { id: "proximas", label: "Próximas", icon: Calendar },
              { id: "calendario", label: "Calendario", icon: CalendarDays },
              { id: "ingresos", label: "Ingresos", icon: TrendingUp },
              { id: "clientes", label: "Por cliente", icon: Users },
            ].map((t) => (
              <button key={t.id} onClick={() => setVista(t.id)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-full text-sm font-medium whitespace-nowrap" style={{ background: vista === t.id ? COLORS.ink : "transparent", color: vista === t.id ? "#fff" : COLORS.ink }}>
                <t.icon size={15} /> {t.label}
              </button>
            ))}
          </div>
        </div>

        <div className="flex flex-wrap gap-2 mb-6">
          <Chip active={filtroServicio === "todos"} onClick={() => setFiltroServicio("todos")} label="Todos" color={COLORS.inkSoft} />
          {SERVICIOS.map((s) => (
            <Chip key={s.id} active={filtroServicio === s.id} onClick={() => setFiltroServicio(s.id)} label={s.label} color={s.color} />
          ))}
        </div>

        {vista === "hoy" && (
          <div>
            <p className="font-mono text-xs uppercase tracking-widest mb-3" style={{ color: COLORS.inkSoft }}>{formatFechaLarga(todayISO())}</p>
            <AgendaPorHoras
              citas={citasHoy}
              onEditar={abrirEditar}
              onEliminar={setConfirmarBorrado}
              onDuplicar={duplicarCita}
              onAgregar={(h) => abrirNueva(todayISO(), `${String(h).padStart(2, "0")}:00`)}
            />
          </div>
        )}

        {vista === "proximas" && (
          <div>
            {proximasPorDia.length === 0 ? (
              <EmptyState texto="No hay próximas citas con este filtro." />
            ) : (
              proximasPorDia.map(({ fecha, citas: citasDia }, idx) => (
                <div key={fecha} className={idx > 0 ? "mt-6" : ""}>
                  <div className="flex items-center gap-2 mb-3">
                    <span className="font-mono text-xs uppercase tracking-widest px-2 py-1 rounded-full" style={{ background: COLORS.gold, color: "#fff" }}>
                      {formatRelativo(fecha) || dateFromISO(fecha).toLocaleDateString("es-ES", { weekday: "short" }).toUpperCase()}
                    </span>
                    <p className="font-mono text-xs uppercase tracking-widest" style={{ color: COLORS.inkSoft }}>{formatFechaLarga(fecha)}</p>
                  </div>
                  <AgendaPorHoras
                    citas={citasDia}
                    onEditar={abrirEditar}
                    onEliminar={setConfirmarBorrado}
                    onDuplicar={duplicarCita}
                    onAgregar={(h) => abrirNueva(fecha, `${String(h).padStart(2, "0")}:00`)}
                  />
                </div>
              ))
            )}
          </div>
        )}

        {vista === "calendario" && (
          <div>
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-1 rounded-full p-1" style={{ background: COLORS.surface, border: `2px solid ${COLORS.ink}` }}>
                {["semana", "mes"].map((m) => (
                  <button key={m} onClick={() => setModoCalendario(m)} className="px-3 py-1 rounded-full text-xs font-medium capitalize" style={{ background: modoCalendario === m ? COLORS.wine : "transparent", color: modoCalendario === m ? "#fff" : COLORS.inkSoft }}>
                    {m}
                  </button>
                ))}
              </div>
              {modoCalendario === "mes" && (
                <p className="font-mono text-xs px-2 py-1 rounded-full" style={{ background: COLORS.goldTint, color: COLORS.gold }}>
                  Mes: {formatMoneda(totalMesSeleccionado) || "$0.00"}
                </p>
              )}
            </div>

            {modoCalendario === "mes" ? (
              <VistaMes mesActual={mesActual} setMesActual={setMesActual} citas={citasFiltradas} onAbrirDia={(iso) => setDiaModalMes(iso)} />
            ) : (
              <VistaSemana
                semanaActual={semanaActual}
                setSemanaActual={setSemanaActual}
                citas={citasFiltradas}
                diaExpandido={diaExpandidoSemana}
                onToggleDia={(iso) => setDiaExpandidoSemana((prev) => (prev === iso ? null : iso))}
                onAgregar={abrirNueva}
                onEditar={abrirEditar}
                onEliminar={setConfirmarBorrado}
                onDuplicar={duplicarCita}
                onEstado={cambiarEstadoRapido}
              />
            )}
          </div>
        )}

        {vista === "ingresos" && (
          <div>
            <div className="flex items-center justify-between mb-5 flex-wrap gap-2">
              <div className="flex items-center gap-1 rounded-full p-1" style={{ background: COLORS.surface, border: `2px solid ${COLORS.ink}` }}>
                {[{ id: "dia", label: "Día" }, { id: "semana", label: "Semana" }, { id: "mes", label: "Mes" }].map((m) => (
                  <button key={m.id} onClick={() => setModoIngresos(m.id)} className="px-3 py-1 rounded-full text-xs font-medium" style={{ background: modoIngresos === m.id ? COLORS.wine : "transparent", color: modoIngresos === m.id ? "#fff" : COLORS.inkSoft }}>
                    {m.label}
                  </button>
                ))}
              </div>
              <div className="flex items-center gap-1">
                <button
                  onClick={() => {
                    if (modoIngresos === "dia") setDiaIngresos(isoFromDate(addDays(dateFromISO(diaIngresos), -1)));
                    else if (modoIngresos === "semana") setSemanaIngresos(addDays(semanaIngresos, -7));
                    else setMesIngresos(new Date(mesIngresos.getFullYear(), mesIngresos.getMonth() - 1, 1));
                  }}
                  className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}
                ><ChevronLeft size={16} /></button>
                <p className="font-mono text-xs capitalize w-32 text-center" style={{ color: COLORS.inkSoft }}>{tituloPeriodoIngresos}</p>
                <button
                  onClick={() => {
                    if (modoIngresos === "dia") setDiaIngresos(isoFromDate(addDays(dateFromISO(diaIngresos), 1)));
                    else if (modoIngresos === "semana") setSemanaIngresos(addDays(semanaIngresos, 7));
                    else setMesIngresos(new Date(mesIngresos.getFullYear(), mesIngresos.getMonth() + 1, 1));
                  }}
                  className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}
                ><ChevronRight size={16} /></button>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2 mb-5">
              <ResumenCard label="Total" valor={formatMoneda(totalPeriodo) || "$0.00"} />
              <ResumenCard label="Citas" valor={citasEnRango.length} />
            </div>

            <div className="rounded-2xl p-4 mb-5" style={cardStyle()}>
              <p className="font-mono text-xs uppercase tracking-widest mb-4" style={{ color: COLORS.inkSoft }}>
                {modoIngresos === "dia" ? "Últimos 7 días" : modoIngresos === "semana" ? "Ingresos por día" : "Ingresos por semana"}
              </p>
              <div className="flex items-end gap-2" style={{ height: 120 }}>
                {barrasIngresos.map((b, i) => (
                  <div key={i} className="flex-1 flex flex-col items-center gap-1.5 h-full justify-end">
                    <span className="font-mono text-xs" style={{ color: COLORS.inkSoft, opacity: b.total > 0 ? 1 : 0 }}>{b.total > 0 ? `$${Math.round(b.total)}` : ""}</span>
                    <div className="w-full rounded-t-md" style={{ height: `${Math.max((b.total / maxBarraIngresos) * 100, 3)}%`, background: b.total > 0 ? COLORS.wine : COLORS.taupe }} />
                    <span className="font-mono text-xs" style={{ color: COLORS.inkSoft }}>{b.label}</span>
                  </div>
                ))}
              </div>
            </div>

            <div className="rounded-2xl p-4" style={cardStyle()}>
              <p className="font-mono text-xs uppercase tracking-widest mb-3" style={{ color: COLORS.inkSoft }}>Por servicio</p>
              {desglosePorServicio.length === 0 ? (
                <p className="text-xs" style={{ color: COLORS.inkSoft }}>Sin ingresos en este periodo.</p>
              ) : (
                <div className="space-y-2.5">
                  {desglosePorServicio.map((s) => (
                    <div key={s.id} className="flex items-center gap-3">
                      <span className="text-xs w-20 shrink-0" style={{ color: COLORS.ink }}>{s.label}</span>
                      <div className="flex-1 h-2 rounded-full overflow-hidden" style={{ background: COLORS.taupeLight }}>
                        <div className="h-full rounded-full" style={{ width: `${(s.total / totalPeriodo) * 100}%`, background: s.color }} />
                      </div>
                      <span className="font-mono text-xs shrink-0 w-14 text-right" style={{ color: COLORS.gold }}>{formatMoneda(s.total)}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {vista === "clientes" && (
          <div>
            <div className="relative mb-5">
              <Search size={15} style={{ position: "absolute", left: 12, top: 12, color: COLORS.inkSoft }} />
              <input value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar cliente..." className="w-full pl-9 pr-3 py-2.5 rounded-xl text-sm border" style={INPUT_STYLE} />
            </div>
            {clientesAgrupados.length === 0 && <EmptyState texto="Aún no hay clientes registrados." />}
            <div className="space-y-4">
              {clientesAgrupados.map((g) => {
                const totalGastado = g.citas.filter((c) => c.estado !== "cancelada").reduce((s, c) => s + (parseFloat(c.precio) || 0), 0);
                return (
                  <div key={g.nombre} className="rounded-2xl p-4" style={cardStyle()}>
                    <div className="flex items-center justify-between mb-3">
                      <div>
                        <p className="font-display text-lg" style={{ fontWeight: 600 }}>{g.nombre}</p>
                        {g.telefono && <p className="text-xs flex items-center gap-1 mt-0.5" style={{ color: COLORS.inkSoft }}><Phone size={11} /> {g.telefono}</p>}
                      </div>
                      <div className="text-right">
                        <span className="font-mono text-xs px-2 py-1 rounded-full block mb-1" style={{ background: COLORS.taupeLight, color: COLORS.inkSoft }}>{g.citas.length} {g.citas.length === 1 ? "cita" : "citas"}</span>
                        {totalGastado > 0 && <span className="font-mono text-xs" style={{ color: COLORS.gold }}>{formatMoneda(totalGastado)} total</span>}
                      </div>
                    </div>
                    <div className="space-y-2">
                      {g.citas.slice(0, 4).map((c) => (
                        <CitaMini key={c.id} cita={c} onEditar={() => abrirEditar(c)} onEliminar={() => setConfirmarBorrado(c.id)} onDuplicar={() => duplicarCita(c)} />
                      ))}
                      {g.citas.length > 4 && <p className="text-xs pl-1" style={{ color: COLORS.inkSoft }}>+{g.citas.length - 4} más en el historial</p>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* Modal día (vista mensual) */}
      {diaModalMes && (
        <ModalDiaDetalle
          iso={diaModalMes}
          citas={citasFiltradas.filter((c) => c.fecha === diaModalMes).sort((a, b) => a.hora.localeCompare(b.hora))}
          onClose={() => setDiaModalMes(null)}
          onAgregar={() => { const f = diaModalMes; setDiaModalMes(null); abrirNueva(f); }}
          onEditar={(c) => { setDiaModalMes(null); abrirEditar(c); }}
          onEliminar={(id) => setConfirmarBorrado(id)}
          onDuplicar={(c) => { setDiaModalMes(null); duplicarCita(c); }}
          onEstado={cambiarEstadoRapido}
        />
      )}

      {/* Modal nueva/editar cita */}
      {modalAbierto && (
        <div className="fixed inset-0 flex items-end sm:items-center justify-center p-0 sm:p-4 z-50" style={{ background: "rgba(43,30,36,0.5)" }} onClick={intentarCerrar}>
          <div className="w-full sm:max-w-md rounded-t-3xl sm:rounded-3xl overflow-hidden flex flex-col" style={{ ...cardStyle(), maxHeight: "90vh" }} onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center justify-between px-6 pt-6 pb-4 shrink-0" style={{ borderBottom: `1px solid ${COLORS.taupe}` }}>
              <h2 className="font-display text-xl" style={{ fontWeight: 600 }}>{editandoId ? "Editar cita" : "Nueva cita"}</h2>
              <button onClick={intentarCerrar} aria-label="Cerrar" className="w-9 h-9 flex items-center justify-center rounded-full" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>
                <X size={18} />
              </button>
            </div>

            <div className="px-6 py-5 overflow-y-auto space-y-4">
              <Campo label="Cliente *">
                <div className="relative">
                  <input
                    autoFocus
                    value={form.cliente}
                    onChange={(e) => { setForm({ ...form, cliente: e.target.value }); if (errorCliente) setErrorCliente(false); setMostrarSugerenciasCliente(true); }}
                    onFocus={() => setMostrarSugerenciasCliente(true)}
                    onBlur={() => setTimeout(() => setMostrarSugerenciasCliente(false), 120)}
                    placeholder="Nombre del cliente"
                    className="w-full px-3 py-2.5 rounded-xl text-sm border"
                    style={errorCliente ? INPUT_ERROR_STYLE : INPUT_STYLE}
                  />
                  {mostrarSugerenciasCliente && sugerenciasCliente.length > 0 && (
                    <div className="absolute left-0 right-0 mt-1.5 rounded-xl overflow-hidden z-10" style={{ background: COLORS.surface, border: `2px solid ${COLORS.ink}`, boxShadow: `3px 3px 0px 0px ${COLORS.ink}` }}>
                      {sugerenciasCliente.map((c) => (
                        <button
                          key={c.nombre}
                          onMouseDown={(e) => { e.preventDefault(); setForm((f) => ({ ...f, cliente: c.nombre, telefono: c.telefono || f.telefono })); setMostrarSugerenciasCliente(false); }}
                          className="w-full flex items-center justify-between gap-2 px-3 py-2 text-left text-sm"
                          style={{ color: COLORS.ink }}
                        >
                          <span className="truncate">{c.nombre}</span>
                          {c.telefono && <span className="font-mono text-xs shrink-0" style={{ color: COLORS.inkSoft }}>{c.telefono}</span>}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                {errorCliente && (
                  <span className="flex items-center gap-1 text-xs mt-1.5" style={{ color: COLORS.danger }}>
                    <AlertCircle size={12} /> Escribe el nombre del cliente para poder agendar.
                  </span>
                )}
              </Campo>

              <Campo label="Teléfono (opcional)">
                <input value={form.telefono} onChange={(e) => setForm({ ...form, telefono: e.target.value })} placeholder="099 999 9999" className="w-full px-3 py-2.5 rounded-xl text-sm border" style={INPUT_STYLE} />
              </Campo>

              <Campo label="Servicio">
                <div className="flex flex-wrap gap-2">
                  {SERVICIOS.map((s) => (
                    <button key={s.id} onClick={() => setForm({ ...form, servicio: s.id })} className="px-3 py-1.5 rounded-full text-xs font-medium border" style={{ borderColor: s.color, background: form.servicio === s.id ? s.color : "transparent", color: form.servicio === s.id ? "#fff" : s.color }}>
                      {s.label}
                    </button>
                  ))}
                </div>
              </Campo>

              <Campo label="Estado">
                <div className="flex flex-wrap gap-2">
                  {ESTADOS.map((s) => (
                    <button key={s.id} onClick={() => setForm({ ...form, estado: s.id })} className="flex items-center gap-1 px-3 py-1.5 rounded-full text-xs font-medium border" style={{ borderColor: s.color, background: form.estado === s.id ? s.color : "transparent", color: form.estado === s.id ? "#fff" : s.color }}>
                      <s.icon size={12} /> {s.label}
                    </button>
                  ))}
                </div>
              </Campo>

              <div className="grid grid-cols-2 gap-3">
                <Campo label="Fecha">
                  <input type="date" value={form.fecha} onChange={(e) => setForm({ ...form, fecha: e.target.value })} className="w-full px-3 py-2.5 rounded-xl text-sm border font-mono" style={INPUT_STYLE} />
                </Campo>
                <Campo label="Hora">
                  <input type="time" value={form.hora} onChange={(e) => setForm({ ...form, hora: e.target.value })} className="w-full px-3 py-2.5 rounded-xl text-sm border font-mono" style={INPUT_STYLE} />
                </Campo>
              </div>

              <Campo label="Precio (opcional)">
                <div className="relative">
                  <DollarSign size={14} style={{ position: "absolute", left: 10, top: 12, color: COLORS.inkSoft }} />
                  <input type="number" min="0" step="0.01" value={form.precio} onChange={(e) => setForm({ ...form, precio: e.target.value })} placeholder="0.00" className="w-full pl-7 pr-3 py-2.5 rounded-xl text-sm border font-mono" style={INPUT_STYLE} />
                </div>
                <div className="flex flex-wrap items-center gap-1.5 mt-2">
                  {SERVICIOS.filter((s) => precios[s.id] != null).map((s) => (
                    <button
                      key={s.id}
                      onClick={() => setForm({ ...form, precio: String(precios[s.id]), servicio: s.id })}
                      className="px-2.5 py-1 rounded-full text-xs font-medium border"
                      style={{ borderColor: s.color, color: s.color }}
                    >
                      ${precios[s.id]} · {s.label}
                    </button>
                  ))}
                  <button onClick={abrirModalPrecios} className="flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium" style={{ color: COLORS.inkSoft }}>
                    <Settings size={11} /> Editar precios
                  </button>
                </div>
              </Campo>

              <Campo label="Especificaciones (opcional)">
                <textarea value={form.notas} onChange={(e) => setForm({ ...form, notas: e.target.value })} placeholder="Ej: rubio ceniza, alergia a amoniaco, trae referencia..." rows={3} className="w-full px-3 py-2.5 rounded-xl text-sm border resize-none" style={INPUT_STYLE} />
              </Campo>

              {citaEncimada && (
                <div className="flex items-start gap-2 px-3 py-2.5 rounded-xl" style={{ background: COLORS.goldTint }}>
                  <AlertCircle size={15} style={{ color: COLORS.gold, flexShrink: 0, marginTop: 1 }} />
                  <p className="text-xs" style={{ color: COLORS.wineDark }}>
                    Ya tienes a <strong>{citaEncimada.cliente}</strong> agendado/a a esta misma hora. Puedes guardar de todas formas si trabajas con ambos a la vez.
                  </p>
                </div>
              )}
            </div>

            <div className="flex gap-3 px-6 py-4 shrink-0" style={{ borderTop: `1px solid ${COLORS.taupe}` }}>
              <button onClick={intentarCerrar} className="flex-1 py-3 rounded-xl text-sm font-semibold" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>
                Cerrar
              </button>
              <button onClick={guardar} className="flex-1 py-3 rounded-xl text-sm font-semibold" style={{ background: COLORS.wine, color: "#fff" }}>
                {editandoId ? "Guardar cambios" : "Agendar cita"}
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmarDescartar && (
        <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(43,30,36,0.5)", zIndex: 60 }} onClick={() => setConfirmarDescartar(false)}>
          <div className="w-full max-w-xs rounded-2xl p-5" style={cardStyle(COLORS.danger)} onClick={(e) => e.stopPropagation()}>
            <p className="text-sm mb-4" style={{ color: COLORS.ink }}>Tienes cambios sin guardar. ¿Quieres descartarlos?</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmarDescartar(false)} className="flex-1 py-2 rounded-lg text-sm font-medium" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>Seguir editando</button>
              <button onClick={cerrarModalDefinitivo} className="flex-1 py-2 rounded-lg text-sm font-medium" style={{ background: COLORS.wine, color: "#fff" }}>Descartar</button>
            </div>
          </div>
        </div>
      )}

      {modalPreciosAbierto && (
        <div className="fixed inset-0 flex items-end sm:items-center justify-center p-0 sm:p-4 z-50" style={{ background: "rgba(43,30,36,0.5)" }} onClick={() => setModalPreciosAbierto(false)}>
          <div className="w-full sm:max-w-sm rounded-t-3xl sm:rounded-3xl overflow-hidden flex flex-col" style={{ ...cardStyle(), maxHeight: "85vh" }} onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center justify-between px-6 pt-6 pb-4 shrink-0" style={{ borderBottom: `1px solid ${COLORS.taupe}` }}>
              <h2 className="font-display text-xl" style={{ fontWeight: 600 }}>Precios sugeridos</h2>
              <button onClick={() => setModalPreciosAbierto(false)} aria-label="Cerrar" className="w-9 h-9 flex items-center justify-center rounded-full" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>
                <X size={18} />
              </button>
            </div>
            <div className="px-6 py-5 overflow-y-auto space-y-3">
              <p className="text-xs" style={{ color: COLORS.inkSoft }}>Estos son los montos que aparecen como sugerencia rápida al agendar. "Otro" no tiene precio fijo porque varía.</p>
              {SERVICIOS.filter((s) => s.id !== "otro").map((s) => (
                <Campo key={s.id} label={s.label}>
                  <div className="relative">
                    <DollarSign size={14} style={{ position: "absolute", left: 10, top: 12, color: COLORS.inkSoft }} />
                    <input
                      type="number" min="0" step="0.01"
                      value={formPrecios[s.id] ?? ""}
                      onChange={(e) => setFormPrecios({ ...formPrecios, [s.id]: e.target.value === "" ? undefined : parseFloat(e.target.value) })}
                      placeholder="Sin precio sugerido"
                      className="w-full pl-7 pr-3 py-2.5 rounded-xl text-sm border font-mono"
                      style={INPUT_STYLE}
                    />
                  </div>
                </Campo>
              ))}
            </div>
            <div className="flex gap-3 px-6 py-4 shrink-0" style={{ borderTop: `1px solid ${COLORS.taupe}` }}>
              <button onClick={() => setModalPreciosAbierto(false)} className="flex-1 py-3 rounded-xl text-sm font-semibold" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>Cancelar</button>
              <button onClick={guardarPrecios} className="flex-1 py-3 rounded-xl text-sm font-semibold" style={{ background: COLORS.wine, color: "#fff" }}>Guardar precios</button>
            </div>
          </div>
        </div>
      )}

      {confirmarBorrado && (
        <div className="fixed inset-0 flex items-center justify-center p-4 z-50" style={{ background: "rgba(43,30,36,0.5)" }} onClick={() => setConfirmarBorrado(null)}>
          <div className="w-full max-w-xs rounded-2xl p-5" style={cardStyle(COLORS.danger)} onClick={(e) => e.stopPropagation()}>
            <p className="text-sm mb-4" style={{ color: COLORS.ink }}>¿Eliminar esta cita? No se puede deshacer.</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmarBorrado(null)} className="flex-1 py-2 rounded-lg text-sm font-medium" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>Cancelar</button>
              <button onClick={() => eliminar(confirmarBorrado)} className="flex-1 py-2 rounded-lg text-sm font-medium" style={{ background: COLORS.wine, color: "#fff" }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {papelera && (
        <div className="fixed bottom-5 left-1/2 z-50 flex items-center gap-3 px-4 py-3 rounded-full shadow-lg" style={{ background: COLORS.ink, color: "#fff", transform: "translateX(-50%)" }}>
          <span className="text-sm">Cita de {papelera.cliente} eliminada</span>
          <button onClick={deshacerEliminar} className="flex items-center gap-1 text-sm font-semibold" style={{ color: COLORS.goldTint }}>
            <Undo2 size={14} /> Deshacer
          </button>
        </div>
      )}
    </div>
  );
}

function AgendaPorHoras({ citas, onEditar, onEliminar, onDuplicar, onAgregar }) {
  const horasConCita = citas.map((c) => parseInt(c.hora.split(":")[0], 10));
  const inicio = Math.min(8, ...(horasConCita.length ? horasConCita : [8]));
  const fin = Math.max(20, ...(horasConCita.length ? horasConCita : [20]));
  const horas = [];
  for (let h = inicio; h <= fin; h++) horas.push(h);

  return (
    <div className="rounded-2xl px-4" style={cardStyle()}>
      {horas.map((h, i) => {
        const citasHora = citas.filter((c) => parseInt(c.hora.split(":")[0], 10) === h).sort((a, b) => a.hora.localeCompare(b.hora));
        return (
          <div key={h} className="flex items-start gap-3 py-2.5" style={{ borderBottom: i < horas.length - 1 ? `1px solid ${COLORS.taupeLight}` : "none" }}>
            <span className="font-mono text-xs w-16 shrink-0 pt-1.5" style={{ color: COLORS.inkSoft }}>{formatHoraSlot(h)}</span>
            <div className="flex-1 min-w-0 space-y-1.5">
              {citasHora.length === 0 ? (
                <button onClick={() => onAgregar(h)} className="w-full text-left py-1.5" aria-label="Agregar cita a esta hora">
                  <span className="font-mono text-xs tracking-widest" style={{ color: COLORS.taupeStrong }}>----------</span>
                </button>
              ) : (
                citasHora.map((c) => (
                  <CitaEnHora key={c.id} cita={c} onEditar={() => onEditar(c)} onEliminar={() => onEliminar(c.id)} onDuplicar={onDuplicar ? () => onDuplicar(c) : undefined} />
                ))
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function CitaEnHora({ cita, onEditar, onEliminar, onDuplicar }) {
  const s = servicioInfo(cita.servicio);
  const precio = formatMoneda(cita.precio);
  const filaStyle =
    cita.estado === "cancelada"
      ? { background: "#E9CDD1", opacity: 0.72 }
      : cita.estado === "completada"
      ? { background: "#D3DCC9" }
      : { background: COLORS.taupeLight };
  return (
    <div className="flex items-center justify-between gap-2 pl-2.5 pr-1.5 py-1.5 rounded-lg" style={{ background: filaStyle.background, opacity: filaStyle.opacity ?? 1 }}>
      <button onClick={onEditar} className="flex items-center gap-2 min-w-0 flex-1 text-left">
        <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
        <span className="font-mono text-xs shrink-0" style={{ color: COLORS.inkSoft }}>{cita.hora}</span>
        <span className="text-sm font-medium truncate" style={{ color: COLORS.ink }}>{cita.cliente}</span>
        {precio && <span className="font-mono text-xs shrink-0" style={{ color: COLORS.gold }}>{precio}</span>}
      </button>
      <div className="flex items-center gap-0.5 shrink-0">
        {onDuplicar && <button onClick={onDuplicar} aria-label="Duplicar cita" className="icon-duplicate p-2 rounded" style={{ color: COLORS.inkSoft }}><Copy size={14} /></button>}
        <button onClick={onEliminar} className="icon-delete p-2 rounded" style={{ color: COLORS.wine }}><Trash2 size={15} /></button>
      </div>
    </div>
  );
}

function Chip({ active, onClick, label, color }) {
  return (
    <button onClick={onClick} className="px-3 py-1.5 rounded-full text-xs font-medium border transition-colors" style={{ borderColor: active ? color : COLORS.taupe, background: active ? color : "transparent", color: active ? "#fff" : color }}>
      {label}
    </button>
  );
}

function Campo({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium mb-1.5" style={{ color: COLORS.inkSoft }}>{label}</span>
      {children}
    </label>
  );
}

function EmptyState({ texto }) {
  return (
    <div className="text-center py-16 rounded-2xl" style={{ border: `1px dashed ${COLORS.taupe}` }}>
      <Scissors size={22} style={{ color: COLORS.taupe, margin: "0 auto 10px" }} />
      <p className="text-sm" style={{ color: COLORS.inkSoft }}>{texto}</p>
    </div>
  );
}

function CitaCard({ cita, onEditar, onEliminar, onDuplicar, onEstado, mostrarFecha }) {
  const s = servicioInfo(cita.servicio);
  const precio = formatMoneda(cita.precio);
  const estadoStyle =
    cita.estado === "cancelada"
      ? { background: "#E9CDD1", accent: "#8C4750", opacity: 0.72 }
      : cita.estado === "completada"
      ? { background: "#D3DCC9", accent: "#6E8560" }
      : { background: COLORS.surface, accent: COLORS.ink };
  return (
    <div className="relative rounded-2xl pl-6 pr-4 py-4" style={{ background: estadoStyle.background, border: `2px solid ${estadoStyle.accent}`, boxShadow: `4px 4px 0px 0px ${estadoStyle.accent}`, opacity: estadoStyle.opacity ?? 1 }}>
      <div style={{ position: "absolute", left: 0, top: 14, bottom: 14, width: 4, borderRadius: 2, background: s.color }} />
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <p className="font-display text-base" style={{ fontWeight: 600, color: COLORS.ink }}>{cita.cliente}</p>
            <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ background: s.color + "22", color: s.color }}>{s.label}</span>
          </div>
          <div className="flex items-center gap-3 mt-1 font-mono text-xs flex-wrap" style={{ color: COLORS.inkSoft }}>
            {mostrarFecha && <span>{formatFechaDMY(cita.fecha)}</span>}
            <span className="flex items-center gap-1"><Clock size={11} /> {cita.hora}</span>
            {cita.telefono && <span className="flex items-center gap-1"><Phone size={11} /> {cita.telefono}</span>}
            {precio && <span style={{ color: COLORS.gold }}>{precio}</span>}
          </div>
          {cita.notas && <p className="text-sm mt-2" style={{ color: COLORS.ink }}>{cita.notas}</p>}
          {onEstado && (
            <div className="flex flex-wrap gap-1.5 mt-2.5">
              {ESTADOS.map((es) => (
                <button key={es.id} onClick={() => onEstado(cita.id, es.id)} className="flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium border" style={{ borderColor: es.color, background: cita.estado === es.id ? es.color : "transparent", color: cita.estado === es.id ? "#fff" : es.color }}>
                  <es.icon size={9} /> {es.label}
                </button>
              ))}
            </div>
          )}
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {onDuplicar && <button onClick={onDuplicar} aria-label="Duplicar cita" className="icon-duplicate p-2 rounded-lg" style={{ color: COLORS.inkSoft }}><Copy size={15} /></button>}
          <button onClick={onEditar} className="icon-edit p-2 rounded-lg" style={{ color: COLORS.inkSoft }}><Pencil size={16} /></button>
          <button onClick={onEliminar} className="icon-delete p-2 rounded-lg" style={{ color: COLORS.wine }}><Trash2 size={16} /></button>
        </div>
      </div>
    </div>
  );
}

function CitaMini({ cita, onEditar, onEliminar, onDuplicar }) {
  const s = servicioInfo(cita.servicio);
  const e = estadoInfo(cita.estado);
  const precio = formatMoneda(cita.precio);
  const filaStyle =
    cita.estado === "cancelada"
      ? { background: "#E9CDD1", opacity: 0.72 }
      : cita.estado === "completada"
      ? { background: "#D3DCC9" }
      : { background: COLORS.taupeLight };
  return (
    <div className="flex items-center justify-between gap-2 py-1.5 px-2 rounded-lg" style={{ background: filaStyle.background, opacity: filaStyle.opacity ?? 1 }}>
      <div className="flex items-center gap-2 min-w-0">
        <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
        <span className="font-mono text-xs shrink-0" style={{ color: COLORS.inkSoft }}>{formatFechaDMY(cita.fecha)} · {cita.hora}</span>
        <span className="text-xs truncate" style={{ color: COLORS.ink }}>{cita.notas || s.label}</span>
        <span className="text-xs font-medium shrink-0" style={{ color: e.color }}>{e.label}</span>
        {precio && <span className="font-mono text-xs shrink-0" style={{ color: COLORS.gold }}>{precio}</span>}
      </div>
      <div className="flex items-center gap-1 shrink-0">
        {onDuplicar && <button onClick={onDuplicar} aria-label="Duplicar cita" className="icon-duplicate p-1.5 rounded" style={{ color: COLORS.inkSoft }}><Copy size={13} /></button>}
        <button onClick={onEditar} className="icon-edit p-1.5 rounded" style={{ color: COLORS.inkSoft }}><Pencil size={14} /></button>
        <button onClick={onEliminar} className="icon-delete p-1.5 rounded" style={{ color: COLORS.wine }}><Trash2 size={14} /></button>
      </div>
    </div>
  );
}

function ResumenCard({ label, valor }) {
  return (
    <div className="rounded-2xl px-3 py-3 text-center" style={cardStyle()}>
      <p className="font-display text-lg" style={{ fontWeight: 600, color: COLORS.ink }}>{valor}</p>
      <p className="text-xs mt-0.5" style={{ color: COLORS.inkSoft }}>{label}</p>
    </div>
  );
}

function VistaMes({ mesActual, setMesActual, citas, onAbrirDia }) {
  const dias = generarDiasMes(mesActual);
  const hoy = todayISO();
  const nombreMes = mesActual.toLocaleDateString("es-ES", { month: "long", year: "numeric" });

  return (
    <div className="rounded-2xl p-4" style={cardStyle()}>
      <div className="flex items-center justify-between mb-4">
        <button onClick={() => setMesActual(new Date(mesActual.getFullYear(), mesActual.getMonth() - 1, 1))} className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}><ChevronLeft size={16} /></button>
        <p className="font-display text-sm capitalize" style={{ fontWeight: 600 }}>{nombreMes}</p>
        <button onClick={() => setMesActual(new Date(mesActual.getFullYear(), mesActual.getMonth() + 1, 1))} className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}><ChevronRight size={16} /></button>
      </div>

      <div className="grid grid-cols-7 gap-1 mb-1">
        {DIAS_SEMANA.map((d) => (
          <div key={d} className="text-center font-mono text-xs uppercase" style={{ color: COLORS.inkSoft }}>{d}</div>
        ))}
      </div>

      <div className="grid grid-cols-7 gap-1">
        {dias.map((d, i) => {
          if (!d) return <div key={i} />;
          const iso = isoFromDate(d);
          const citasDia = citas.filter((c) => c.fecha === iso);
          const esHoy = iso === hoy;
          const serviciosUnicos = [...new Set(citasDia.map((c) => c.servicio))].slice(0, 3);
          return (
            <button
              key={i}
              onClick={() => onAbrirDia(iso)}
              className="aspect-square rounded-xl flex flex-col items-center justify-center gap-0.5"
              style={{ background: esHoy ? COLORS.wineTint : "transparent", border: esHoy ? `1px solid ${COLORS.wine}` : `1px solid transparent` }}
            >
              <span className="text-xs font-medium" style={{ color: COLORS.ink }}>{d.getDate()}</span>
              <div className="flex items-center gap-0.5" style={{ minHeight: 6 }}>
                {serviciosUnicos.map((sid) => (
                  <span key={sid} className="w-1.5 h-1.5 rounded-full" style={{ background: servicioInfo(sid).color }} />
                ))}
              </div>
            </button>
          );
        })}
      </div>
      <p className="text-xs text-center mt-3" style={{ color: COLORS.inkSoft }}>Toca un día para ver o agregar citas.</p>
    </div>
  );
}

function ModalDiaDetalle({ iso, citas, onClose, onAgregar, onEditar, onEliminar, onDuplicar, onEstado }) {
  return (
    <div className="fixed inset-0 flex items-end sm:items-center justify-center p-0 sm:p-4 z-50" style={{ background: "rgba(43,30,36,0.5)" }} onClick={onClose}>
      <div className="w-full sm:max-w-md rounded-t-3xl sm:rounded-3xl overflow-hidden flex flex-col" style={{ ...cardStyle(), maxHeight: "85vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 pt-6 pb-4 shrink-0" style={{ borderBottom: `1px solid ${COLORS.taupe}` }}>
          <h2 className="font-display text-lg capitalize" style={{ fontWeight: 600 }}>{formatFechaLarga(iso)}</h2>
          <button onClick={onClose} aria-label="Cerrar" className="w-9 h-9 flex items-center justify-center rounded-full" style={{ background: COLORS.taupeLight, color: COLORS.ink }}>
            <X size={18} />
          </button>
        </div>

        <div className="px-6 py-4 overflow-y-auto space-y-3">
          {citas.length === 0 ? (
            <EmptyState texto="Sin citas este día." />
          ) : (
            citas.map((c) => (
              <CitaCard key={c.id} cita={c} onEditar={() => onEditar(c)} onEliminar={() => onEliminar(c.id)} onDuplicar={onDuplicar ? () => onDuplicar(c) : undefined} onEstado={onEstado} />
            ))
          )}
        </div>

        <div className="px-6 py-4 shrink-0" style={{ borderTop: `1px solid ${COLORS.taupe}` }}>
          <button onClick={onAgregar} className="w-full flex items-center justify-center gap-1.5 py-3 rounded-xl text-sm font-semibold" style={{ background: COLORS.wine, color: "#fff" }}>
            <Plus size={15} /> Agregar cita este día
          </button>
        </div>
      </div>
    </div>
  );
}

function VistaSemana({ semanaActual, setSemanaActual, citas, diaExpandido, onToggleDia, onAgregar, onEditar, onEliminar, onDuplicar, onEstado }) {
  const dias = Array.from({ length: 7 }, (_, i) => addDays(semanaActual, i));
  const hoy = todayISO();
  const rango = `${formatFechaCorta(isoFromDate(dias[0]))} – ${formatFechaCorta(isoFromDate(dias[6]))}`;

  return (
    <div className="rounded-2xl p-4" style={cardStyle()}>
      <div className="flex items-center justify-between mb-4">
        <button onClick={() => setSemanaActual(addDays(semanaActual, -7))} className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}><ChevronLeft size={16} /></button>
        <p className="font-display text-sm" style={{ fontWeight: 600 }}>{rango}</p>
        <button onClick={() => setSemanaActual(addDays(semanaActual, 7))} className="p-1.5 rounded-lg" style={{ color: COLORS.inkSoft }}><ChevronRight size={16} /></button>
      </div>

      <div className="space-y-2">
        {dias.map((d, i) => {
          const iso = isoFromDate(d);
          const citasDia = citas.filter((c) => c.fecha === iso).sort((a, b) => a.hora.localeCompare(b.hora));
          const esHoy = iso === hoy;
          const abierto = diaExpandido === iso;
          return (
            <div key={i} className="rounded-xl overflow-hidden" style={{ background: COLORS.taupeLight }}>
              <button onClick={() => onToggleDia(iso)} className="w-full flex items-center justify-between gap-3 px-3 py-2.5 text-left">
                <div className="flex items-center gap-2">
                  <span className="font-mono text-xs uppercase w-8" style={{ color: esHoy ? COLORS.wine : COLORS.inkSoft }}>{DIAS_SEMANA[i]}</span>
                  <span className="text-sm font-medium" style={{ color: esHoy ? COLORS.wine : COLORS.ink }}>{d.getDate()}</span>
                  {citasDia.length > 0 && (
                    <span className="font-mono text-xs px-1.5 py-0.5 rounded-full" style={{ background: COLORS.surface, color: COLORS.inkSoft }}>{citasDia.length}</span>
                  )}
                </div>
                <div className="flex items-center gap-2">
                  <div className="flex items-center gap-1">
                    {citasDia.slice(0, 5).map((c) => (
                      <span key={c.id} className="w-2 h-2 rounded-full" style={{ background: servicioInfo(c.servicio).color }} />
                    ))}
                  </div>
                  <ChevronDown size={15} style={{ color: COLORS.inkSoft, transform: abierto ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                </div>
              </button>

              {abierto && (
                <div className="px-3 pb-3 space-y-2">
                  {citasDia.length === 0 ? (
                    <p className="text-xs px-1 pb-1" style={{ color: COLORS.inkSoft }}>Sin citas este día.</p>
                  ) : (
                    citasDia.map((c) => (
                      <CitaMini key={c.id} cita={c} onEditar={() => onEditar(c)} onEliminar={() => onEliminar(c.id)} onDuplicar={onDuplicar ? () => onDuplicar(c) : undefined} />
                    ))
                  )}
                  <button onClick={() => onAgregar(iso)} className="flex items-center gap-1 text-xs font-medium px-1" style={{ color: COLORS.wine }}>
                    <Plus size={12} /> Agregar cita
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
