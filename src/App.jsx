import React, { Fragment, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { controlRecords, controlSource } from './controlData.js';
import {
  AlertTriangle,
  BadgeCheck,
  Building2,
  CalendarDays,
  ChevronDown,
  Clock3,
  Pencil,
  FileSpreadsheet,
  RadioTower,
  Save,
  Search,
  ShieldCheck,
  Siren,
  X,
  TrainFront,
  Upload,
  UsersRound,
} from 'lucide-react';

const emovaLogo = new URL('../logo-emova-transparente.png', import.meta.url).href;
const policeLogo = new URL('../logo-policia-ciudad-transparente.png', import.meta.url).href;
const morningShift = '05:30 A 13:30 - Turno Mañana';
const afternoonShift = '16:30 A 00:00 - Turno Tarde';
const allDaysOption = 'Fechas';
const allLinesOption = 'Líneas';
const allStatusesOption = 'Estados';
const allShiftsOption = 'Turnos';
const shiftOptions = [allShiftsOption, morningShift, afternoonShift];
const networkLines = ['Línea A', 'Línea B', 'Línea C', 'Línea D', 'Línea E', 'Línea H'];

const sampleRecords = [
  {
    id: 'CP-001',
    fecha: '2026-05-15',
    hora: '06:00',
    dni: '41.285.945',
    nombre: 'MELGAREJO, CLAUDIA MARLENE',
    dniCubre: '-',
    nombreCubre: '-',
    linea: 'Línea B',
    estacion: 'Carlos Pellegrini',
    turno: morningShift,
    ingreso: '05:26',
    egreso: '13:30',
    fiscalizado: 'SI',
    dependencia: 'PFA',
    dotacion: 1,
    puesto: 'Andén y combinación',
    estado: 'Operativo',
    controlCmi: 'Se observa al adicional en puesto.',
    novedad: 'Refuerzo en hora pico y control preventivo de accesos.',
  },
  {
    id: 'CP-002',
    fecha: '2026-05-15',
    hora: '08:30',
    dni: '36.570.634',
    nombre: 'CABAÑA GOMEZ, MICAELA',
    dniCubre: '-',
    nombreCubre: 'GARCIA, JESICA',
    linea: 'Línea D',
    estacion: 'Catedral',
    turno: morningShift,
    ingreso: '05:14',
    egreso: 'S/I',
    fiscalizado: 'NO',
    dependencia: 'Policía de la Ciudad',
    dotacion: 2,
    puesto: 'Molinete sur',
    estado: 'Observación',
    controlCmi: 'Pendiente de confirmación de egreso.',
    novedad: 'Demora por alta concentración de pasajeros.',
  },
  {
    id: 'CP-003',
    fecha: '2026-05-15',
    hora: '14:00',
    dni: '36.646.673',
    nombre: 'CASTRO, MARIA ALEJANDRA',
    dniCubre: '-',
    nombreCubre: 'CELARIO, DALILA',
    linea: 'Línea A',
    estacion: 'Plaza de Mayo',
    turno: afternoonShift,
    ingreso: '16:30',
    egreso: '00:00',
    fiscalizado: 'SI',
    dependencia: 'PFA',
    dotacion: 2,
    puesto: 'Hall central',
    estado: 'Operativo',
    controlCmi: 'Control visual normal.',
    novedad: 'Control de rutina sin novedades relevantes.',
  },
  {
    id: 'CP-004',
    fecha: '2026-05-15',
    hora: '18:10',
    dni: '39.742.740',
    nombre: 'MALDONADO, MANUEL DAVID',
    dniCubre: '-',
    nombreCubre: 'ENRIQUEZ, ADRIAN',
    linea: 'Línea C',
    estacion: 'Constitución',
    turno: afternoonShift,
    ingreso: '16:34',
    egreso: 'S/I',
    fiscalizado: 'SI',
    dependencia: 'Policía de la Ciudad',
    dotacion: 2,
    puesto: 'Acceso principal',
    estado: 'Crítico',
    controlCmi: 'Fiscalización informa ausencia temporal del adicional.',
    novedad: 'Intervención por incidente menor, seguimiento desde CMI.',
  },
  {
    id: 'CP-005',
    fecha: '2026-05-15',
    hora: '22:00',
    dni: '30.046.398',
    nombre: 'ACOSTA, AGUSTIN LUIS',
    dniCubre: '-',
    nombreCubre: '-',
    linea: 'Línea H',
    estacion: 'Once',
    turno: afternoonShift,
    ingreso: '22:00',
    egreso: 'S/I',
    fiscalizado: '-',
    dependencia: 'Gendarmería',
    dotacion: 1,
    puesto: 'Combinación',
    estado: 'Operativo',
    controlCmi: 'Cobertura activa en monitoreo.',
    novedad: 'Cobertura nocturna activa.',
  },
  {
    id: 'CP-006',
    fecha: '2026-05-15',
    hora: '07:10',
    dni: '34.918.276',
    nombre: 'FERNANDEZ, MARCELO JAVIER',
    dniCubre: '-',
    nombreCubre: 'SOSA, PAULA',
    linea: 'Línea E',
    estacion: 'Bolívar',
    turno: morningShift,
    ingreso: '05:42',
    egreso: '13:34',
    fiscalizado: 'SI',
    dependencia: 'Policía de la Ciudad',
    dotacion: 2,
    puesto: 'Hall boletería',
    estado: 'Operativo',
    controlCmi: 'Adicional observado en puesto asignado.',
    novedad: 'Cobertura normal de Línea E.',
  },
];

const fieldAliases = {
  fecha: ['fecha', 'dia', 'día', 'date'],
  hora: ['hora', 'horario', 'time'],
  dni: ['dni'],
  nombre: ['apellido y nombre', 'nombre', 'titular'],
  dniCubre: ['dni2', 'dni cubre', 'dni reemplazo'],
  nombreCubre: ['apellido y nombre2', 'cubre puesto', 'nombre2', 'reemplazo'],
  linea: ['linea', 'línea', 'ramal', 'subte'],
  estacion: ['estacion', 'estación', 'station', 'ubicacion', 'ubicación'],
  turno: ['turno', 'franja'],
  ingreso: ['ingreso', 'entrada'],
  egreso: ['egreso', 'salida'],
  fiscalizado: ['fiscalizado', 'fiscalizacion', 'fiscalización'],
  dependencia: ['dependencia', 'fuerza', 'organismo', 'policia', 'policía'],
  dotacion: ['dotacion', 'dotación', 'cantidad', 'personal', 'efectivos'],
  puesto: ['puesto', 'cubre puesto', 'sector', 'posición', 'posicion'],
  estado: ['estado', 'situacion', 'situación', 'status'],
  novedad: ['novedad', 'novedades', 'observacion', 'observación', 'observaciones'],
  controlCmi: ['control cmi', 'cmi', 'monitoreo'],
};

const lineOptions = [allLinesOption, ...networkLines, 'Premetro'];
const statusOptions = [allStatusesOption, 'Operativo', 'Observación', 'Crítico'];

function normalizeHeader(value) {
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function resolveField(row, field) {
  const aliases = fieldAliases[field].map(normalizeHeader);
  const entries = Object.entries(row);
  const found = entries.find(([key]) => aliases.some((alias) => normalizeHeader(key).includes(alias)));
  return found?.[1] ?? '';
}

function toDate(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }
  const text = String(value);
  const match = text.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (match) {
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    return `${year}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
  }
  return text.slice(0, 10);
}

function toTime(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
  if (typeof value === 'number' && value > 0 && value < 1) {
    const totalMinutes = Math.round(value * 24 * 60);
    return `${String(Math.floor(totalMinutes / 60)).padStart(2, '0')}:${String(totalMinutes % 60).padStart(2, '0')}`;
  }
  const match = String(value).match(/\d{1,2}:\d{2}/);
  return match?.[0] ?? String(value);
}

function normalizeShift(value, fallbackIndex = 0, ingreso = '') {
  const text = String(value || '').toLowerCase();
  const entryHour = Number(String(ingreso || '').match(/\d{1,2}/)?.[0]);
  if (text.includes('16:30') || text.includes('tarde') || (entryHour >= 14 && entryHour <= 23)) {
    return afternoonShift;
  }
  if (text.includes('05:30') || text.includes('mañana') || text.includes('manana') || entryHour < 14) {
    return morningShift;
  }
  return fallbackIndex % 2 === 0 ? morningShift : afternoonShift;
}

function normalizeStatus(value, index) {
  const text = String(value || '').toLowerCase();
  if (text.includes('crit') || text.includes('incid') || text.includes('alert')) return 'Crítico';
  if (text.includes('obs') || text.includes('pend') || text.includes('demora')) return 'Observación';
  return index % 11 === 0 ? 'Observación' : 'Operativo';
}

function normalizePoliceStatus({ fiscalizado, egreso, observaciones, controlCmi }, index) {
  const text = `${fiscalizado} ${egreso} ${observaciones} ${controlCmi}`.toLowerCase();
  if (text.includes('no encontro') || text.includes('no encontró') || text.includes('ausente')) return 'Crítico';
  if (text.includes('seguridad no fiscalizo') || text.includes('no fiscalizo') || fiscalizado === 'NO') return 'Observación';
  if (String(egreso).toUpperCase() === 'S/I') return 'Observación';
  return normalizeStatus(text, index);
}

function normalizeGridRows(rows, sheetName) {
  const sheetDate = sheetName?.match(/\d{1,2}-\d{1,2}-\d{4}/)?.[0]?.replace(/(\d{2})-(\d{2})-(\d{4})/, '$3-$2-$1') ?? '';
  const cleanRows = rows
    .map((row) => row.map((cell) => String(cell ?? '').trim()))
    .filter((row) => row.some(Boolean));

  return cleanRows
    .filter((row) => {
      const text = row.join(' ').toLowerCase();
      const isTitle = text.includes('centro de monitoreo') || text.includes('control personal policial');
      const isHeaderOnly = ['titular', 'cubre puesto', 'franco', 'ausente'].some((word) => text === word);
      return !isTitle && !isHeaderOnly;
    })
    .map((row, index) => {
      const text = row.join(' | ');
      const likelyLine = row.find((cell) => /\blinea\b|\blínea\b|\b[abcdeh]\b/i.test(cell)) ?? '';
      const likelyStation = row.find((cell) => /estaci[oó]n|and[eé]n|hall|acceso|combinaci[oó]n/i.test(cell)) ?? '';
      const likelyShift = row.find((cell) => /mañana|manana|tarde|05:30|16:30/i.test(cell)) ?? '';
      const likelyTime = row.find((cell) => /\d{1,2}:\d{2}/.test(cell)) ?? '';
      const staffCells = row.filter((cell) => cell && !/\d{1,2}:\d{2}/.test(cell));

      return {
        id: `XL-${sheetName}-${String(index + 1).padStart(3, '0')}`,
        fecha: sheetDate,
        hora: toTime(likelyTime) || '--:--',
        dni: '',
        nombre: row[2] || row[1] || 'Personal sin identificar',
        dniCubre: row[4] || '',
        nombreCubre: row[5] || '',
        linea: normalizeLine(likelyLine || text, index),
        estacion: likelyStation || 'Red SubTE',
        turno: normalizeShift(likelyShift, index, row[8]),
        ingreso: row[8] || '',
        egreso: row[9] || '',
        fiscalizado: row[10] || '-',
        dependencia: /ciudad/i.test(text) ? 'Policía de la Ciudad' : /pfa|federal/i.test(text) ? 'PFA' : 'Control policial',
        dotacion: Math.max(1, Math.min(12, staffCells.length || 1)),
        puesto: row[5] || row[4] || row[1] || 'Cobertura operativa',
        estado: normalizeStatus(text, index),
        controlCmi: row[12] || '',
        novedad: text.length > 150 ? `${text.slice(0, 150)}...` : text || 'Registro importado desde Excel.',
      };
    });
}

function normalizeLine(value, index) {
  const text = String(value || '').toUpperCase();
  const match = text.match(/\b(A|B|C|D|E|H)\b/);
  if (match) return `Línea ${match[1]}`;
  return lineOptions[(index % 6) + 1];
}

function normalizeRecords(rows, sheetName) {
  return rows
    .filter((row) => Object.values(row).some((value) => String(value ?? '').trim()))
    .map((row, index) => {
      const fallbackDate = sheetName?.match(/\d{1,2}-\d{1,2}-\d{4}/)?.[0]?.replace(/(\d{2})-(\d{2})-(\d{4})/, '$3-$2-$1') ?? '';
      const nombreCubre = String(resolveField(row, 'nombreCubre') || '');
      const dotacion = Number(resolveField(row, 'dotacion')) || (nombreCubre && nombreCubre !== '-' ? 2 : 1);
      const fiscalizado = String(resolveField(row, 'fiscalizado') || '-').toUpperCase();
      const egreso = toTime(resolveField(row, 'egreso')) || '';
      const observaciones = String(resolveField(row, 'novedad') || '');
      const controlCmi = String(resolveField(row, 'controlCmi') || '');
      return {
        id: `XL-${String(index + 1).padStart(3, '0')}`,
        fecha: toDate(resolveField(row, 'fecha')) || fallbackDate,
        hora: toTime(resolveField(row, 'hora') || resolveField(row, 'ingreso')) || '--:--',
        dni: String(resolveField(row, 'dni') || ''),
        nombre: String(resolveField(row, 'nombre') || 'Personal sin identificar'),
        dniCubre: String(resolveField(row, 'dniCubre') || ''),
        nombreCubre,
        linea: normalizeLine(resolveField(row, 'linea'), index),
        estacion: String(resolveField(row, 'estacion') || resolveField(row, 'puesto') || 'Estación sin especificar'),
        turno: normalizeShift(resolveField(row, 'turno'), index, resolveField(row, 'ingreso')),
        ingreso: toTime(resolveField(row, 'ingreso')) || '',
        egreso,
        fiscalizado,
        dependencia: String(resolveField(row, 'dependencia') || 'Control policial'),
        dotacion,
        puesto: String(resolveField(row, 'puesto') || 'Cobertura operativa'),
        estado: normalizePoliceStatus({ fiscalizado, egreso, observaciones, controlCmi }, index),
        controlCmi,
        novedad: observaciones || 'Registro importado desde Excel.',
      };
    });
}

function App() {
  const [showSplash, setShowSplash] = useState(true);
  const [records, setRecords] = useState(controlRecords.length ? controlRecords : sampleRecords);
  const [query, setQuery] = useState('');
  const [line, setLine] = useState(allLinesOption);
  const [status, setStatus] = useState(allStatusesOption);
  const [shift, setShift] = useState(allShiftsOption);
  const [day, setDay] = useState(allDaysOption);
  const [sourceName, setSourceName] = useState(controlRecords.length ? controlSource : 'Datos demo');
  const [importError, setImportError] = useState('');
  const [editingRecord, setEditingRecord] = useState(null);
  const fileInputRef = useRef(null);

  const filteredRecords = useMemo(() => {
    const term = query.toLowerCase().trim();
    return records.filter((record) => {
      const matchesTerm = !term || Object.values(record).join(' ').toLowerCase().includes(term);
      const matchesLine = line === allLinesOption || record.linea === line;
      const matchesStatus = status === allStatusesOption || record.estado === status;
      const matchesShift = shift === allShiftsOption || record.turno === shift;
      const recordDay = record.dia || record.fecha || 'Sin día';
      const matchesDay = day === allDaysOption || recordDay === day;
      return matchesTerm && matchesLine && matchesStatus && matchesShift && matchesDay;
    });
  }, [day, line, query, records, shift, status]);

  const stats = useMemo(() => {
    const totalStaff = filteredRecords.reduce((sum, record) => sum + Number(record.dotacion || 0), 0);
    const critical = filteredRecords.filter((record) => record.estado === 'Crítico').length;
    const observations = filteredRecords.filter((record) => record.estado === 'Observación').length;
    const activeLines = networkLines.length;
    const audited = filteredRecords.filter((record) => String(record.fiscalizado).toUpperCase() === 'SI').length;
    const noExit = filteredRecords.filter((record) => String(record.egreso).toUpperCase() === 'S/I').length;
    return { totalStaff, critical, observations, activeLines, audited, noExit, total: filteredRecords.length };
  }, [filteredRecords]);

  const dayGroups = useMemo(() => {
    const groups = new Map();
    filteredRecords.forEach((record) => {
      const day = record.dia || record.fecha || 'Sin día';
      if (!groups.has(day)) groups.set(day, []);
      groups.get(day).push(record);
    });
    return Array.from(groups.entries()).map(([day, items]) => ({ day, items }));
  }, [filteredRecords]);

  async function handleFileUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setImportError('');

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const imported = workbook.SheetNames.flatMap((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const objectRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        const normalizedObjects = normalizeRecords(objectRows, sheetName).filter((record) => {
          const text = `${record.estacion} ${record.puesto} ${record.novedad}`.toLowerCase();
          return !text.includes('centro de monitoreo integral') && !text.includes('gerencia de seguridad');
        });

        if (normalizedObjects.length > 3) {
          return normalizedObjects;
        }

        const gridRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
        return normalizeGridRows(gridRows, sheetName);
      });

      if (!imported.length) {
        setImportError('No se detectaron registros con datos en el archivo.');
        return;
      }

      setRecords(imported);
      setEditingRecord(null);
      setSourceName(file.name);
      event.target.value = '';
    } catch (error) {
      setImportError('No se pudo leer el Excel. Probá con un archivo .xlsx o .xls válido.');
    }
  }

  function updateEditingField(field, value) {
    setEditingRecord((current) => ({ ...current, [field]: value }));
  }

  function saveEditingRecord() {
    if (!editingRecord) return;
    setRecords((current) =>
      current.map((record) =>
        record.id === editingRecord.id
          ? { ...editingRecord, dotacion: Number(editingRecord.dotacion) || 1 }
          : record,
      ),
    );
    setEditingRecord(null);
  }

  const shifts = shiftOptions;
  const days = [allDaysOption, ...Array.from(new Set(records.map((record) => record.dia || record.fecha || 'Sin día')))];

  useEffect(() => {
    const timer = window.setTimeout(() => setShowSplash(false), 2300);
    return () => window.clearTimeout(timer);
  }, []);

  return (
    <main className="monitor-grid min-h-screen px-4 py-5 sm:px-6 lg:px-8">
      {showSplash ? <SplashScreen onSkip={() => setShowSplash(false)} /> : null}
      <section className="mx-auto flex max-w-7xl flex-col gap-5">
        <header className="glass-panel rounded-lg p-4 sm:p-5">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex items-start gap-4">
              <div className="grid h-12 w-12 shrink-0 place-items-center rounded-md border border-emova/40 bg-emova/15 shadow-glow">
                <RadioTower className="h-6 w-6 text-emova" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2 text-xs font-semibold uppercase tracking-[0.24em] text-subte">
                  <span>EMOVA</span>
                  <span className="h-1 w-1 rounded-full bg-subte" />
                  <span>SubTE</span>
                  <span className="h-1 w-1 rounded-full bg-subte" />
                  <span>CMI</span>
                </div>
                <h1 className="mt-2 text-2xl font-semibold tracking-normal text-white sm:text-3xl">
                  Control Policial Operativo
                </h1>
                <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-300">
                  Panel de monitoreo para seguimiento de coberturas, dotaciones, estaciones y novedades del servicio.
                </p>
              </div>
            </div>

            <div className="flex flex-col gap-3">
              <InstitutionalLogos />
              <div className="grid grid-cols-2 gap-3 text-sm sm:flex sm:items-center">
                <StatusPill icon={ShieldCheck} label="Sistema" value="En línea" tone="green" />
                <StatusPill icon={CalendarDays} label="Fuente" value={sourceName} tone="blue" />
              </div>
            </div>
          </div>
        </header>

        <section className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
          <StatCard icon={UsersRound} label="Efectivos filtrados" value={stats.totalStaff} detail={`${stats.total} registros activos`} />
          <StatCard icon={TrainFront} label="Líneas de la red" value={stats.activeLines} detail="A, B, C, D, E y H" />
          <StatCard icon={BadgeCheck} label="Fiscalizados" value={stats.audited} detail={`${stats.noExit} sin egreso informado`} tone="amber" />
          <StatCard icon={Siren} label="Alertas" value={stats.critical + stats.observations} detail="Observación y críticos" tone="red" />
        </section>

        <section className="grid gap-4 xl:grid-cols-[1fr_360px]">
          <div className="glass-panel rounded-lg p-4">
            <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <h2 className="text-lg font-semibold text-white">Registros de cobertura</h2>
                <p className="mt-1 text-sm text-slate-400">Búsqueda, filtros y lectura operativa de la dotación policial.</p>
              </div>
              <div className="flex flex-col gap-2 sm:flex-row">
                <button
                  className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-emova px-4 text-sm font-semibold text-ink transition hover:bg-cyan-300"
                  onClick={() => fileInputRef.current?.click()}
                  type="button"
                >
                  <Upload className="h-4 w-4" />
                  Importar Excel
                </button>
                <input
                  ref={fileInputRef}
                  className="hidden"
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={handleFileUpload}
                />
              </div>
            </div>

            <div className="mt-4 grid gap-3 lg:grid-cols-[1.3fr_0.75fr_0.75fr_0.75fr_0.75fr]">
              <label className="relative">
                  <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500" />
                  <input
                    className="h-11 w-full rounded-md border border-line bg-ink/70 pl-10 pr-3 text-sm text-white outline-none transition placeholder:text-slate-500 focus:border-emova"
                    placeholder="Buscar por estación, fuerza, puesto o novedad"
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                  />
              </label>
              <SelectControl label="Fecha" icon={CalendarDays} value={day} onChange={setDay} options={days} />
              <SelectControl label="Línea" icon={TrainFront} value={line} onChange={setLine} options={lineOptions} />
              <SelectControl label="Estado" icon={BadgeCheck} value={status} onChange={setStatus} options={statusOptions} />
              <SelectControl label="Turno" icon={Clock3} value={shift} onChange={setShift} options={shifts} />
            </div>

            {importError ? (
              <div className="mt-3 rounded-md border border-alert/40 bg-alert/10 px-3 py-2 text-sm text-red-100">
                {importError}
              </div>
            ) : null}

            <div className="mt-4 overflow-hidden rounded-lg border border-line">
              <div className="hidden max-h-[560px] overflow-auto md:block">
                <table className="min-w-[1180px] w-full table-fixed border-collapse text-left text-[13px]">
                  <colgroup>
                    <col className="w-[78px]" />
                    <col className="w-[92px]" />
                    <col className="w-[180px]" />
                    <col className="w-[160px]" />
                    <col className="w-[76px]" />
                    <col className="w-[142px]" />
                    <col className="w-[150px]" />
                    <col className="w-[118px]" />
                    <col className="w-[104px]" />
                    <col className="w-[96px]" />
                    <col />
                  </colgroup>
                  <thead className="sticky top-0 z-10 bg-[#0b1a22] text-[10px] uppercase tracking-[0.12em] text-slate-400">
                    <tr>
                      <Th>Acción</Th>
                      <Th>Fecha</Th>
                      <Th>Personal</Th>
                      <Th>Cubre puesto</Th>
                      <Th>Línea</Th>
                      <Th>Estación</Th>
                      <Th>Horario</Th>
                      <Th>Dependencia</Th>
                      <Th>Fiscalización</Th>
                      <Th>Estado</Th>
                      <Th>Control CMI / Novedad</Th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-line/80">
                    {dayGroups.map(({ day, items }) => (
                      <Fragment key={day}>
                        <tr className="bg-emova/15">
                          <td colSpan={11} className="px-3 py-2 text-xs font-semibold uppercase tracking-[0.12em] text-cyan-100">
                            {day} · {items.length} registros
                          </td>
                        </tr>
                        {items.map((record) => (
                          <tr key={`${record.id}-${record.estacion}-${record.hora}`} className="bg-panel/40 transition hover:bg-emova/10">
                            <Td>
                              <button
                                className="inline-flex h-8 items-center justify-center gap-1.5 rounded-md border border-emova/35 bg-emova/10 px-2 text-[11px] font-semibold text-cyan-100 transition hover:border-emova hover:bg-emova/20"
                                type="button"
                                onClick={() => setEditingRecord(record)}
                              >
                                <Pencil className="h-3 w-3" />
                                Editar
                              </button>
                            </Td>
                            <Td>
                              <div className="font-medium leading-4 text-white">{record.fecha || 'Sin fecha'}</div>
                              <div className="text-[11px] leading-4 text-slate-500">{record.hora}</div>
                            </Td>
                            <Td>
                              <div className="font-semibold leading-4 text-white">{record.nombre}</div>
                              <div className="text-[11px] leading-4 text-slate-500">DNI {record.dni || 'S/D'}</div>
                            </Td>
                            <Td>
                              <div className="font-medium leading-4 text-slate-200">{record.nombreCubre && record.nombreCubre !== '-' ? record.nombreCubre : 'Sin reemplazo'}</div>
                              <div className="text-[11px] leading-4 text-slate-500">{record.dniCubre && record.dniCubre !== '-' ? `DNI ${record.dniCubre}` : 'Titular en puesto'}</div>
                            </Td>
                            <Td>
                              <span className="inline-flex rounded bg-subte/15 px-2 py-0.5 text-[11px] font-semibold text-subte">{record.linea}</span>
                            </Td>
                            <Td>
                              <div className="font-medium leading-4 text-white">{record.estacion}</div>
                              <div className="text-[11px] leading-4 text-slate-500">{record.puesto}</div>
                            </Td>
                            <Td>
                              <div className="font-medium leading-4 text-slate-200">{record.turno}</div>
                              <div className="mt-1 flex flex-col gap-0.5 text-[11px] leading-4 text-slate-500">
                                <TimeValue label="Ingreso" value={record.ingreso} />
                                <TimeValue label="Egreso" value={record.egreso} />
                              </div>
                            </Td>
                            <Td>{record.dependencia}</Td>
                            <Td>
                              <FiscalBadge value={record.fiscalizado} />
                              <div className="mt-1 text-xs text-slate-500">Dotación {record.dotacion}</div>
                            </Td>
                            <Td>
                              <StatusBadge status={record.estado} />
                            </Td>
                            <Td>
                              <p className="max-w-[360px] whitespace-normal leading-4 text-slate-300">{record.controlCmi || record.novedad}</p>
                              {record.controlCmi && record.novedad ? (
                                <p className="mt-1 max-w-[360px] whitespace-normal text-[11px] leading-4 text-slate-500">{record.novedad}</p>
                              ) : null}
                            </Td>
                          </tr>
                        ))}
                      </Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="max-h-[620px] space-y-3 overflow-auto bg-ink/30 p-3 md:hidden">
                {dayGroups.map(({ day, items }) => (
                  <section key={day} className="space-y-3">
                    <div className="sticky top-0 z-10 rounded-md border border-emova/25 bg-[#0b1a22] px-3 py-2 text-sm font-semibold text-cyan-100">
                      {day} · {items.length} registros
                    </div>
                    {items.map((record) => (
                      <article key={`${record.id}-mobile`} className="rounded-md border border-line bg-panel/70 p-2.5">
                        <div className="flex items-start justify-between gap-3">
                          <div>
                            <div className="text-xs text-slate-500">{record.fecha}</div>
                            <h3 className="mt-1 text-sm font-semibold leading-4 text-white">{record.nombre}</h3>
                            <p className="text-xs text-slate-500">DNI {record.dni || 'S/D'}</p>
                          </div>
                          <StatusBadge status={record.estado} />
                        </div>

                        <div className="mt-2 grid grid-cols-2 gap-2 text-xs">
                          <MobileFact label="Línea" value={record.linea} />
                          <MobileFact label="Estación" value={record.estacion} />
                          <MobileFact label="Turno" value={record.turno} wide />
                          <div className="rounded-md border border-line bg-ink/60 px-2.5 py-1.5">
                            <span className="block text-slate-500">Ingreso</span>
                            <TimeValue label="" value={record.ingreso} />
                          </div>
                          <div className="rounded-md border border-line bg-ink/60 px-2.5 py-1.5">
                            <span className="block text-slate-500">Egreso</span>
                            <TimeValue label="" value={record.egreso} />
                          </div>
                          <MobileFact label="Fiscalizado" value={record.fiscalizado || '-'} />
                          <MobileFact label="Cubre puesto" value={record.nombreCubre && record.nombreCubre !== '-' ? record.nombreCubre : 'Sin reemplazo'} wide />
                        </div>

                        {(record.controlCmi || record.novedad) ? (
                          <p className="mt-2 rounded-md border border-line bg-ink/50 px-2.5 py-1.5 text-xs leading-4 text-slate-300">
                            {record.controlCmi || record.novedad}
                          </p>
                        ) : null}

                        <button
                          className="mt-3 inline-flex h-10 w-full items-center justify-center gap-2 rounded-md border border-emova/35 bg-emova/10 px-3 text-xs font-semibold text-cyan-100"
                          type="button"
                          onClick={() => setEditingRecord(record)}
                        >
                          <Pencil className="h-3.5 w-3.5" />
                          Editar registro
                        </button>
                      </article>
                    ))}
                  </section>
                ))}
              </div>
            </div>
          </div>

          <aside className="flex flex-col gap-4">
            <div className="glass-panel rounded-lg p-4">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-white">Mapa operativo</h2>
                <Building2 className="h-5 w-5 text-emova" />
              </div>
              <div className="mt-5 space-y-3">
                {lineOptions.slice(1).map((lineName, index) => {
                  const count = filteredRecords.filter((record) => record.linea === lineName).length;
                  const width = stats.total ? Math.max(8, (count / stats.total) * 100) : 8;
                  return (
                    <div key={lineName}>
                      <div className="mb-1 flex items-center justify-between text-xs text-slate-400">
                        <span>{lineName}</span>
                        <span>{count} registros</span>
                      </div>
                      <div className="h-2 rounded-full bg-ink">
                        <div
                          className="h-2 rounded-full"
                          style={{
                            width: `${width}%`,
                            background: index % 2 ? '#00a0df' : '#f6c141',
                          }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="glass-panel rounded-lg p-4">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="h-5 w-5 text-subte" />
                <h2 className="text-lg font-semibold text-white">Importación</h2>
              </div>
              <p className="mt-3 text-sm leading-6 text-slate-300">
                El panel acepta Excel con una o varias hojas. Detecta titular, cubre puesto, ingreso, egreso, fiscalización, observaciones y control CMI.
              </p>
              <div className="mt-4 grid grid-cols-2 gap-2 text-xs text-slate-400">
                {['DNI', 'Nombre', 'Cubre puesto', 'Línea', 'Estación', 'Ingreso', 'Egreso', 'Control CMI'].map((item) => (
                  <div key={item} className="rounded border border-line bg-ink/50 px-3 py-2">
                    {item}
                  </div>
                ))}
              </div>
            </div>
          </aside>
        </section>
      </section>

      {editingRecord ? (
        <EditRecordPanel
          record={editingRecord}
          onChange={updateEditingField}
          onCancel={() => setEditingRecord(null)}
          onSave={saveEditingRecord}
        />
      ) : null}
    </main>
  );
}

function SplashScreen({ onSkip }) {
  return (
    <button
      className="fixed inset-0 z-[60] grid cursor-pointer place-items-center border-0 bg-ink/95 p-6 text-left backdrop-blur-md"
      type="button"
      onClick={onSkip}
      aria-label="Ingresar al dashboard"
    >
      <div className="flex w-full max-w-3xl flex-col items-center gap-7">
        <div className="flex w-full items-center justify-center gap-4">
          <div className="flex h-28 w-56 items-center justify-center p-4">
            <img src={emovaLogo} alt="Logo EMOVA" className="max-h-20 w-full object-contain" />
          </div>
          <div className="text-4xl font-light text-slate-500">|</div>
          <div className="flex h-28 w-40 items-center justify-center p-4">
            <img src={policeLogo} alt="Logo Policía de la Ciudad de Buenos Aires" className="max-h-24 w-full object-contain" />
          </div>
        </div>
        <div className="text-center">
          <p className="text-xs font-semibold uppercase tracking-[0.3em] text-subte">Centro de monitoreo integral</p>
          <h2 className="mt-3 text-2xl font-semibold text-white sm:text-4xl">Control Policial Emova | Subte</h2>
          <p className="mt-3 text-sm text-slate-400">Cargando registros operativos...</p>
        </div>
      </div>
    </button>
  );
}

function StatusPill({ icon: Icon, label, value, tone }) {
  const color = tone === 'green' ? 'text-emerald-300 bg-emerald-400/10 border-emerald-400/30' : 'text-cyan-200 bg-emova/10 border-emova/30';
  return (
    <div className={`rounded-md border px-3 py-2 ${color}`}>
      <div className="flex items-center gap-2">
        <Icon className="h-4 w-4 shrink-0" />
        <div className="min-w-0">
          <div className="text-[10px] uppercase tracking-[0.18em] opacity-70">{label}</div>
          <div className="truncate text-sm font-semibold">{value}</div>
        </div>
      </div>
    </div>
  );
}

function InstitutionalLogos() {
  return (
    <div className="flex items-center justify-center gap-3 p-2">
      <div className="flex min-h-16 w-44 items-center justify-center bg-transparent px-2 py-2">
        <img
          src={emovaLogo}
          alt="Logo EMOVA"
          className="max-h-14 w-full object-contain"
        />
      </div>

      <div className="text-2xl font-light text-slate-500">|</div>

      <div className="flex min-h-16 w-24 items-center justify-center bg-transparent px-2 py-2">
        <img
          src={policeLogo}
          alt="Logo Policía de la Ciudad de Buenos Aires"
          className="max-h-14 w-full object-contain"
        />
      </div>
    </div>
  );
}

function StatCard({ icon: Icon, label, value, detail, tone = 'blue' }) {
  const toneClass = {
    blue: 'text-emova bg-emova/10 border-emova/30',
    amber: 'text-subte bg-subte/10 border-subte/30',
    red: 'text-red-300 bg-alert/10 border-alert/30',
  }[tone];

  return (
    <article className="glass-panel rounded-lg p-4">
      <div className="flex items-center justify-between gap-3">
        <div>
          <p className="text-sm text-slate-400">{label}</p>
          <p className="mt-2 text-3xl font-semibold text-white">{value}</p>
          <p className="mt-1 text-xs text-slate-500">{detail}</p>
        </div>
        <div className={`grid h-12 w-12 place-items-center rounded-md border ${toneClass}`}>
          <Icon className="h-6 w-6" />
        </div>
      </div>
    </article>
  );
}

function SelectControl({ label, icon: Icon, value, onChange, options }) {
  return (
    <label className="relative">
        <Icon className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500" />
        <select
          aria-label={label}
          className="h-11 w-full appearance-none rounded-md border border-line bg-ink/70 pl-10 pr-9 text-sm text-white outline-none transition focus:border-emova"
          value={value}
          onChange={(event) => onChange(event.target.value)}
        >
          {options.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
        <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500" />
    </label>
  );
}

function StatusBadge({ status }) {
  const styles = {
    Operativo: 'border-emerald-400/30 bg-emerald-400/10 text-emerald-200',
    Observación: 'border-subte/30 bg-subte/10 text-yellow-100',
    Crítico: 'border-alert/40 bg-alert/10 text-red-100',
  };

  return (
    <span className={`inline-flex whitespace-nowrap rounded px-2 py-0.5 text-[11px] font-semibold ${styles[status] ?? styles.Operativo}`}>
      {status}
    </span>
  );
}

function FiscalBadge({ value }) {
  const normalized = String(value || '-').toUpperCase();
  const className = normalized === 'SI'
    ? 'border-emerald-400/30 bg-emerald-400/10 text-emerald-200'
    : normalized === 'NO'
      ? 'border-alert/40 bg-alert/10 text-red-100'
      : 'border-slate-500/30 bg-slate-500/10 text-slate-300';

  return <span className={`inline-flex rounded px-2 py-0.5 text-[11px] font-semibold ${className}`}>{normalized}</span>;
}

function TimeValue({ label, value }) {
  const isMissing = !value || String(value).toUpperCase() === 'S/I';
  return (
    <span>
      {label ? `${label}: ` : ''}{isMissing ? <small className="text-[10px] uppercase tracking-[0.1em] text-slate-500">sin informar</small> : value}
    </span>
  );
}

function MobileFact({ label, value, wide = false }) {
  return (
    <div className={`rounded-md border border-line bg-ink/60 px-2.5 py-1.5 ${wide ? 'col-span-2' : ''}`}>
      <span className="block text-slate-500">{label}</span>
      <span className="mt-1 block font-medium text-slate-200">{value || 'S/D'}</span>
    </div>
  );
}

function EditRecordPanel({ record, onChange, onCancel, onSave }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/60 p-3 backdrop-blur-sm sm:items-center">
      <section className="glass-panel max-h-[92vh] w-full max-w-5xl overflow-auto rounded-lg p-4 sm:p-5">
        <div className="flex flex-col gap-3 border-b border-line pb-4 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h2 className="text-xl font-semibold text-white">Editar registro policial</h2>
            <p className="mt-1 text-sm text-slate-400">Los cambios se aplican al panel actual y a los filtros de la tabla.</p>
          </div>
          <div className="flex gap-2">
            <button
              className="inline-flex h-10 items-center justify-center gap-2 rounded-md border border-line bg-ink px-4 text-sm font-semibold text-slate-200 transition hover:border-slate-500"
              type="button"
              onClick={onCancel}
            >
              <X className="h-4 w-4" />
              Cancelar
            </button>
            <button
              className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-emova px-4 text-sm font-semibold text-ink transition hover:bg-cyan-300"
              type="button"
              onClick={onSave}
            >
              <Save className="h-4 w-4" />
              Guardar
            </button>
          </div>
        </div>

        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          <EditField label="Fecha" value={record.fecha} onChange={(value) => onChange('fecha', value)} />
          <EditField label="Hora" value={record.hora} onChange={(value) => onChange('hora', value)} />
          <EditField label="DNI titular" value={record.dni} onChange={(value) => onChange('dni', value)} />
          <EditField label="Apellido y nombre" value={record.nombre} onChange={(value) => onChange('nombre', value)} />
          <EditField label="DNI cubre puesto" value={record.dniCubre} onChange={(value) => onChange('dniCubre', value)} />
          <EditField label="Cubre puesto" value={record.nombreCubre} onChange={(value) => onChange('nombreCubre', value)} />
          <EditSelect label="Línea" value={record.linea} options={lineOptions.filter((item) => item !== allLinesOption)} onChange={(value) => onChange('linea', value)} />
          <EditField label="Estación" value={record.estacion} onChange={(value) => onChange('estacion', value)} />
          <EditSelect label="Turno" value={record.turno} options={shiftOptions.filter((item) => item !== allShiftsOption)} onChange={(value) => onChange('turno', value)} />
          <EditField label="Ingreso" value={record.ingreso} onChange={(value) => onChange('ingreso', value)} />
          <EditField label="Egreso" value={record.egreso} onChange={(value) => onChange('egreso', value)} />
          <EditSelect label="Fiscalizado" value={record.fiscalizado} options={['SI', 'NO', '-']} onChange={(value) => onChange('fiscalizado', value)} />
          <EditField label="Dependencia" value={record.dependencia} onChange={(value) => onChange('dependencia', value)} />
          <EditField label="Dotación" type="number" value={record.dotacion} onChange={(value) => onChange('dotacion', value)} />
          <EditSelect label="Estado" value={record.estado} options={statusOptions.filter((item) => item !== allStatusesOption)} onChange={(value) => onChange('estado', value)} />
          <EditField label="Puesto" value={record.puesto} onChange={(value) => onChange('puesto', value)} />
          <EditArea label="Control CMI" value={record.controlCmi} onChange={(value) => onChange('controlCmi', value)} />
          <EditArea label="Observaciones" value={record.novedad} onChange={(value) => onChange('novedad', value)} />
        </div>
      </section>
    </div>
  );
}

function EditField({ label, value, onChange, type = 'text' }) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.14em] text-slate-500">{label}</span>
      <input
        className="h-11 w-full rounded-md border border-line bg-ink/80 px-3 text-sm text-white outline-none transition focus:border-emova"
        type={type}
        value={value ?? ''}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function EditSelect({ label, value, options, onChange }) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.14em] text-slate-500">{label}</span>
      <select
        className="h-11 w-full rounded-md border border-line bg-ink/80 px-3 text-sm text-white outline-none transition focus:border-emova"
        value={value ?? ''}
        onChange={(event) => onChange(event.target.value)}
      >
        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </label>
  );
}

function EditArea({ label, value, onChange }) {
  return (
    <label className="block xl:col-span-3">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.14em] text-slate-500">{label}</span>
      <textarea
        className="min-h-24 w-full rounded-md border border-line bg-ink/80 px-3 py-2 text-sm text-white outline-none transition focus:border-emova"
        value={value ?? ''}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function Th({ children }) {
  return <th className="px-3 py-2 font-semibold">{children}</th>;
}

function Td({ children }) {
  return <td className="px-3 py-2 align-top leading-4 text-slate-300">{children}</td>;
}

export default App;
