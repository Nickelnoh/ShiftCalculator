const MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
];

const HOLIDAYS_BY_YEAR = {
    2026: new Set([
        "2026-01-01", "2026-01-02", "2026-01-03", "2026-01-04",
        "2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08",
        "2026-02-23", "2026-03-08", "2026-05-01", "2026-05-09",
        "2026-06-12", "2026-11-04"
    ])
};

const ANNUAL_NORM_HOURS_BY_YEAR = {
    2026: 1972
};

const STORAGE_KEYS = {
    year: "ssc_year",
    graphs: "ssc_graphs",
    activeId: "ssc_activeId",
    absences: "ssc_absences",
    cellHeight: "ssc_cellHeight",
    overrides: "ssc_cellOverrides"
};

const SUPABASE_TABLES = {
    settings: "app_settings",
    graphs: "graphs",
    shifts: "shifts",
    absences: "absences",
    overrides: "cell_overrides"
};

const CELL_HEIGHT_MIN = 44;
const CELL_HEIGHT_MAX = 140;
const CELL_HEIGHT_STEP = 8;
const DEFAULT_CELL_HEIGHT = 72;
const DEFAULT_YEAR = 2026;
const MS_IN_DAY = 24 * 60 * 60 * 1000;

let currentYear = DEFAULT_YEAR;
let graphs = [];
let activeGraphId = "";
let absences = [];
let cellOverrides = [];
let currentCellHeight = DEFAULT_CELL_HEIGHT;
let currentZoom = 100;
let fullscreenModalInstance = null;
let shareModalInstance = null;
let cellEditModalInstance = null;

let supabaseClient = null;
let remoteSaveTimer = null;
let remoteSavePromise = Promise.resolve();
let hasRemoteStateLoaded = false;
let currentAppKey = "";
let availableRemoteTables = new Set([SUPABASE_TABLES.settings, SUPABASE_TABLES.graphs, SUPABASE_TABLES.shifts, SUPABASE_TABLES.absences, SUPABASE_TABLES.overrides]);
let remoteCapabilities = { shiftType: true };

let isReadOnlyMode = false;
let readOnlyGraphId = "";
let pendingCreatedGraphId = "";
let editingCellContext = null;
let currentAbsenceCategory = "planned";
const AUTO_BALANCE_PREFIX = "autobal";


function uid(prefix = "id") {
    return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
}

function mod(value, base) {
    return ((value % base) + base) % base;
}

function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
}

function safeParseJSON(raw, fallback) {
    try {
        return raw ? JSON.parse(raw) : fallback;
    } catch {
        return fallback;
    }
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function toId(value) {
    return String(value ?? "");
}

function buildIsoDate(year, monthIndex, day) {
    return `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseIsoDateParts(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ""))) {
        return null;
    }
    const [year, month, day] = iso.split("-").map(Number);
    if (!year || !month || !day) {
        return null;
    }
    return { year, month, day };
}

function toUtcDayNumber(iso) {
    const parts = parseIsoDateParts(iso);
    if (!parts) {
        return null;
    }
    return Math.floor(Date.UTC(parts.year, parts.month - 1, parts.day) / MS_IN_DAY);
}

function normalizeIsoDate(iso, fallbackYear = currentYear) {
    const parts = parseIsoDateParts(iso);
    if (parts) {
        return buildIsoDate(parts.year, parts.month - 1, parts.day);
    }
    return `${fallbackYear}-01-01`;
}

function getDaysInMonth(year, monthIndex) {
    return new Date(Date.UTC(year, monthIndex + 1, 0)).getUTCDate();
}

function getWeekday(year, monthIndex, day) {
    return new Date(Date.UTC(year, monthIndex, day)).getUTCDay();
}

function isLeapYear(year) {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function isHoliday(dateStr) {
    const year = Number(String(dateStr).slice(0, 4));
    return HOLIDAYS_BY_YEAR[year]?.has(dateStr) || false;
}

function isPreholiday(dateStr) {
    const dayNumber = toUtcDayNumber(dateStr);
    if (dayNumber === null) return false;
    const nextDateStr = toIsoFromUtcDayNumber(dayNumber + 1);
    if (!nextDateStr) return false;
    const parts = parseIsoDateParts(dateStr);
    if (!parts) return false;
    const weekday = getWeekday(parts.year, parts.month - 1, parts.day);
    return !isHoliday(dateStr) && ![0, 6].includes(weekday) && isHoliday(nextDateStr);
}

function toIsoFromUtcDayNumber(dayNumber) {
    const date = new Date(dayNumber * MS_IN_DAY);
    return buildIsoDate(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function normalizeShiftName(value, fallbackIndex = 0) {
    const raw = String(value ?? "").trim();
    if (!raw) {
        return `Смена ${fallbackIndex + 1}`;
    }

    const match = raw.match(/^\s*Смена\s+(\d+)/i);
    if (match) {
        return `Смена ${match[1]}`;
    }

    return raw;
}

function normalizeShiftKind(value) {
    return String(value || "").toLowerCase() === "reserve" ? "reserve" : "regular";
}

function getShiftKindLabel(kind) {
    return normalizeShiftKind(kind) === "reserve" ? "Подменная" : "Основная";
}

function createDefaultShifts(count = 4) {
    return Array.from({ length: count }, (_, index) => ({
        id: uid("shift"),
        name: `Смена ${index + 1}`,
        shiftKind: "regular"
    }));
}

function createDefaultGraph() {
    return {
        id: uid("graph"),
        name: "Цех №1",
        type: "24",
        firstShiftDate: `${currentYear}-01-01`,
        smeny: createDefaultShifts(4)
    };
}

function normalizeGraph(graph, index) {
    const type = graph?.type === "12" ? "12" : "24";
    const smeny = Array.isArray(graph?.smeny) && graph.smeny.length
        ? graph.smeny.map((item, itemIndex) => ({
            id: toId(item?.id || uid("shift")),
            name: normalizeShiftName(item?.name, itemIndex),
            shiftKind: normalizeShiftKind(item?.shiftKind)
        }))
        : createDefaultShifts(4);

    return {
        id: toId(graph?.id || `graph-${index + 1}`),
        name: String(graph?.name || `График ${index + 1}`),
        type,
        firstShiftDate: normalizeIsoDate(graph?.firstShiftDate, currentYear),
        smeny
    };
}

function normalizeAbsenceType(type) {
    if (type === "Срочный больничный") return "Срочный больничный";
    if (type === "Больничный") return "Больничный";
    if (type === "Отгул") return "Отгул";
    return "Отпуск";
}

function isShortAbsenceType(type) {
    return normalizeAbsenceType(type) === "Срочный больничный" || normalizeAbsenceType(type) === "Отгул";
}

function getAbsenceVisualMeta(type) {
    const normalized = normalizeAbsenceType(type);
    if (normalized === "Отпуск") {
        return { className: "absence-otpusk", cellClass: "absence-vacation-cell", icon: "🏖️" };
    }
    if (normalized === "Отгул") {
        return { className: "absence-otgul", cellClass: "absence-dayoff-cell", icon: "🟢" };
    }
    if (normalized === "Срочный больничный") {
        return { className: "absence-srochny", cellClass: "absence-urgent-cell", icon: "🚑" };
    }
    return { className: "absence-bolnichny", cellClass: "absence-sick-cell", icon: "🏥" };
}

function normalizeAbsence(absence) {
    return {
        id: toId(absence?.id || uid("absence")),
        graphId: toId(absence?.graphId),
        smenaId: toId(absence?.smenaId),
        type: normalizeAbsenceType(absence?.type),
        start: normalizeIsoDate(absence?.start, currentYear),
        end: normalizeIsoDate(absence?.end, currentYear)
    };
}

function isAutoBalanceOverride(override) {
    return toId(override?.id).startsWith(`${AUTO_BALANCE_PREFIX}-`);
}

function buildAutoBalanceOverrideId(graphId, smenaId, dateStr) {
    return `${AUTO_BALANCE_PREFIX}-${toId(graphId)}-${toId(smenaId)}-${toId(dateStr)}`;
}

function normalizeOverride(override) {
    return {
        id: toId(override?.id || uid("override")),
        graphId: toId(override?.graphId),
        smenaId: toId(override?.smenaId),
        date: normalizeIsoDate(override?.date, currentYear),
        mode: String(override?.mode || "auto"),
        hours: clamp(parseInt(override?.hours, 10) || 0, 0, 24),
        night: clamp(parseInt(override?.night, 10) || 0, 0, 24),
        absenceType: override?.absenceType ? normalizeAbsenceType(override.absenceType) : ""
    };
}

function readLegacyLocalState() {
    const localYear = clamp(parseInt(localStorage.getItem(STORAGE_KEYS.year), 10) || DEFAULT_YEAR, 2000, 2100);
    const localCellHeight = clamp(parseInt(localStorage.getItem(STORAGE_KEYS.cellHeight), 10) || DEFAULT_CELL_HEIGHT, CELL_HEIGHT_MIN, CELL_HEIGHT_MAX);

    const rawGraphs = safeParseJSON(localStorage.getItem(STORAGE_KEYS.graphs), []);
    const localGraphs = Array.isArray(rawGraphs) && rawGraphs.length
        ? rawGraphs.map(normalizeGraph)
        : [createDefaultGraph()];

    const savedActiveId = toId(localStorage.getItem(STORAGE_KEYS.activeId));
    const localActiveId = localGraphs.some(graph => graph.id === savedActiveId) ? savedActiveId : localGraphs[0].id;

    const rawAbsences = safeParseJSON(localStorage.getItem(STORAGE_KEYS.absences), []);
    const localAbsences = Array.isArray(rawAbsences)
        ? rawAbsences.map(normalizeAbsence).filter(item => item.start <= item.end)
        : [];

    const rawOverrides = safeParseJSON(localStorage.getItem(STORAGE_KEYS.overrides), []);
    const localOverrides = Array.isArray(rawOverrides)
        ? rawOverrides.map(normalizeOverride)
        : [];

    return {
        year: localYear,
        graphs: localGraphs,
        activeId: localActiveId,
        absences: localAbsences,
        cellHeight: localCellHeight,
        overrides: localOverrides
    };
}

function applySnapshot(snapshot) {
    const safeState = snapshot && typeof snapshot === "object" ? snapshot : {};
    currentYear = clamp(parseInt(safeState.year, 10) || DEFAULT_YEAR, 2000, 2100);
    currentCellHeight = clamp(parseInt(safeState.cellHeight, 10) || DEFAULT_CELL_HEIGHT, CELL_HEIGHT_MIN, CELL_HEIGHT_MAX);

    graphs = Array.isArray(safeState.graphs) && safeState.graphs.length
        ? safeState.graphs.map(normalizeGraph)
        : [createDefaultGraph()];

    const savedActiveId = toId(safeState.activeId);
    activeGraphId = graphs.some(graph => graph.id === savedActiveId) ? savedActiveId : graphs[0].id;

    const knownShiftIds = new Set(graphs.flatMap(graph => graph.smeny.map(smena => smena.id)));
    const knownGraphIds = new Set(graphs.map(graph => graph.id));

    absences = Array.isArray(safeState.absences)
        ? safeState.absences
            .map(normalizeAbsence)
            .filter(item => item.start <= item.end && knownGraphIds.has(item.graphId) && knownShiftIds.has(item.smenaId))
        : [];

    cellOverrides = Array.isArray(safeState.overrides)
        ? safeState.overrides
            .map(normalizeOverride)
            .filter(item => knownGraphIds.has(item.graphId) && knownShiftIds.has(item.smenaId))
        : [];
}

function normalizeState() {
    applySnapshot(readLegacyLocalState());
    save({ skipRemote: true });
}

function writeLocalSnapshot() {
    try {
        localStorage.setItem(STORAGE_KEYS.year, String(currentYear));
        localStorage.setItem(STORAGE_KEYS.graphs, JSON.stringify(graphs));
        localStorage.setItem(STORAGE_KEYS.activeId, activeGraphId);
        localStorage.setItem(STORAGE_KEYS.absences, JSON.stringify(absences));
        localStorage.setItem(STORAGE_KEYS.cellHeight, String(currentCellHeight));
        localStorage.setItem(STORAGE_KEYS.overrides, JSON.stringify(cellOverrides));
    } catch (error) {
        console.error("Ошибка сохранения в localStorage", error);
        alert("Не удалось сохранить данные в браузере.");
    }
}

function getSupabaseConfig() {
    const raw = window.SUPABASE_CONFIG || {};
    return {
        url: String(raw.url || "").trim(),
        anonKey: String(raw.anonKey || "").trim(),
        appKey: String(raw.appKey || "shiftmaster-main").trim() || "shiftmaster-main"
    };
}

function isSupabaseConfigured() {
    const config = getSupabaseConfig();
    return Boolean(config.url && config.anonKey && window.supabase?.createClient);
}

function ensureSupabaseClient() {
    if (!isSupabaseConfigured()) {
        return null;
    }

    if (!supabaseClient) {
        const config = getSupabaseConfig();
        supabaseClient = window.supabase.createClient(config.url, config.anonKey, {
            auth: { persistSession: false }
        });
        currentAppKey = config.appKey;
    }

    return supabaseClient;
}

function isMissingRelationError(error) {
    if (!error) return false;
    const text = `${error.message || ""} ${error.details || ""} ${error.hint || ""}`.toLowerCase();
    return error.code === "42P01" || error.code === "PGRST205" || text.includes("relation") && text.includes("does not exist") || text.includes("could not find the table");
}

function isMissingColumnError(error, columnName) {
    if (!error) return false;
    const text = `${error.message || ""} ${error.details || ""} ${error.hint || ""}`.toLowerCase();
    const column = String(columnName || "").toLowerCase();
    return error.code === "42703" || error.code === "PGRST204" || (column && text.includes(column) && (text.includes("column") || text.includes("schema cache")));
}

async function selectRemoteRows(client, tableName, columns, appKey, orderColumn = "sort_order") {
    const query = client.from(tableName).select(columns).eq("app_key", appKey);
    const response = orderColumn ? await query.order(orderColumn, { ascending: true }) : await query;
    if (response.error) {
        if (isMissingRelationError(response.error)) {
            availableRemoteTables.delete(tableName);
            return [];
        }
        throw response.error;
    }
    availableRemoteTables.add(tableName);
    return response.data || [];
}

async function maybeSingleRemote(client, tableName, columns, appKey) {
    const response = await client.from(tableName).select(columns).eq("app_key", appKey).maybeSingle();
    if (response.error) {
        if (isMissingRelationError(response.error)) {
            availableRemoteTables.delete(tableName);
            return null;
        }
        throw response.error;
    }
    availableRemoteTables.add(tableName);
    return response.data || null;
}

async function loadFromSupabase() {
    const client = ensureSupabaseClient();
    if (!client) {
        return false;
    }

    const config = getSupabaseConfig();
    currentAppKey = config.appKey;

    try {
        const settingsRow = await maybeSingleRemote(client, SUPABASE_TABLES.settings, "year, active_graph_id, cell_height", config.appKey);
        const graphRows = await selectRemoteRows(client, SUPABASE_TABLES.graphs, "id, name, type, first_shift_date, sort_order", config.appKey, "sort_order");
        let shiftRows = [];
        try {
            shiftRows = await selectRemoteRows(client, SUPABASE_TABLES.shifts, "id, graph_id, name, shift_type, sort_order", config.appKey, "sort_order");
            remoteCapabilities.shiftType = true;
        } catch (error) {
            if (!isMissingColumnError(error, "shift_type")) {
                throw error;
            }
            remoteCapabilities.shiftType = false;
            shiftRows = await selectRemoteRows(client, SUPABASE_TABLES.shifts, "id, graph_id, name, sort_order", config.appKey, "sort_order");
        }
        const absenceRows = await selectRemoteRows(client, SUPABASE_TABLES.absences, "id, graph_id, shift_id, absence_type, start_date, end_date", config.appKey, "start_date");
        const overrideRows = await selectRemoteRows(client, SUPABASE_TABLES.overrides, "id, graph_id, shift_id, work_date, mode, hours, night, absence_type", config.appKey, "work_date");

        if (!settingsRow && !graphRows.length && !shiftRows.length && !absenceRows.length && !overrideRows.length) {
            hasRemoteStateLoaded = true;
            return false;
        }

        const remoteGraphs = graphRows.map((graphRow, graphIndex) => ({
            id: toId(graphRow.id),
            name: String(graphRow.name || `График ${graphIndex + 1}`),
            type: graphRow.type === "12" ? "12" : "24",
            firstShiftDate: normalizeIsoDate(graphRow.first_shift_date, currentYear),
            smeny: shiftRows
                .filter(shiftRow => toId(shiftRow.graph_id) === toId(graphRow.id))
                .map((shiftRow, shiftIndex) => ({
                    id: toId(shiftRow.id),
                    name: normalizeShiftName(shiftRow.name, shiftIndex),
                    shiftKind: normalizeShiftKind(shiftRow.shift_type)
                }))
        }));

        const remoteAbsences = absenceRows.map((absenceRow) => ({
            id: toId(absenceRow.id),
            graphId: toId(absenceRow.graph_id),
            smenaId: toId(absenceRow.shift_id),
            type: normalizeAbsenceType(absenceRow.absence_type),
            start: normalizeIsoDate(absenceRow.start_date, currentYear),
            end: normalizeIsoDate(absenceRow.end_date, currentYear)
        }));

        const remoteOverrides = overrideRows.map((overrideRow) => ({
            id: toId(overrideRow.id),
            graphId: toId(overrideRow.graph_id),
            smenaId: toId(overrideRow.shift_id),
            date: normalizeIsoDate(overrideRow.work_date, currentYear),
            mode: String(overrideRow.mode || "auto"),
            hours: clamp(parseInt(overrideRow.hours, 10) || 0, 0, 24),
            night: clamp(parseInt(overrideRow.night, 10) || 0, 0, 24),
            absenceType: normalizeAbsenceType(overrideRow.absence_type || "")
        }));

        applySnapshot({
            year: settingsRow?.year ?? currentYear,
            activeId: settingsRow?.active_graph_id ?? "",
            cellHeight: settingsRow?.cell_height ?? currentCellHeight,
            graphs: remoteGraphs,
            absences: remoteAbsences,
            overrides: remoteOverrides
        });

        writeLocalSnapshot();
        hasRemoteStateLoaded = true;
        return true;
    } catch (error) {
        console.error("Ошибка загрузки из Supabase", error);
        return false;
    }
}

async function deleteMissingRows(client, tableName, appKey, currentIds) {
    if (!availableRemoteTables.has(tableName)) {
        return;
    }

    const { data, error } = await client.from(tableName).select("id").eq("app_key", appKey);
    if (error) {
        if (isMissingRelationError(error)) {
            availableRemoteTables.delete(tableName);
            return;
        }
        throw error;
    }

    const existingIds = (data || []).map(item => toId(item.id));
    const idsToDelete = existingIds.filter(id => !currentIds.includes(id));
    if (!idsToDelete.length) {
        return;
    }

    const { error: deleteError } = await client.from(tableName).delete().eq("app_key", appKey).in("id", idsToDelete);
    if (deleteError && !isMissingRelationError(deleteError)) {
        throw deleteError;
    }
}

async function upsertMaybe(client, tableName, rows, onConflict) {
    if (!rows.length) return;
    const { error } = await client.from(tableName).upsert(rows, { onConflict });
    if (error) {
        if (isMissingRelationError(error)) {
            availableRemoteTables.delete(tableName);
            return;
        }
        throw error;
    }
    availableRemoteTables.add(tableName);
}

async function syncStateToSupabase() {
    const client = ensureSupabaseClient();
    if (!client) {
        return;
    }

    const config = getSupabaseConfig();
    currentAppKey = config.appKey;
    const now = new Date().toISOString();

    const graphRows = graphs.map((graph, graphIndex) => ({
        id: graph.id,
        app_key: config.appKey,
        name: graph.name,
        type: graph.type,
        first_shift_date: graph.firstShiftDate,
        sort_order: graphIndex,
        updated_at: now
    }));

    const shiftRows = graphs.flatMap(graph => graph.smeny.map((smena, shiftIndex) => ({
        id: smena.id,
        app_key: config.appKey,
        graph_id: graph.id,
        name: normalizeShiftName(smena.name, shiftIndex),
        shift_type: normalizeShiftKind(smena.shiftKind),
        sort_order: shiftIndex,
        updated_at: now
    })));

    const absenceRows = absences.map(absence => ({
        id: absence.id,
        app_key: config.appKey,
        graph_id: absence.graphId,
        shift_id: absence.smenaId,
        absence_type: absence.type,
        start_date: absence.start,
        end_date: absence.end,
        updated_at: now
    }));

    const overrideRows = cellOverrides.map(item => ({
        id: item.id,
        app_key: config.appKey,
        graph_id: item.graphId,
        shift_id: item.smenaId,
        work_date: item.date,
        mode: item.mode,
        hours: item.hours,
        night: item.night,
        absence_type: item.absenceType || null,
        updated_at: now
    }));

    try {
        const { error: settingsError } = await client.from(SUPABASE_TABLES.settings).upsert({
            app_key: config.appKey,
            year: currentYear,
            active_graph_id: activeGraphId || null,
            cell_height: currentCellHeight,
            updated_at: now
        }, { onConflict: "app_key" });
        if (settingsError && !isMissingRelationError(settingsError)) throw settingsError;

        await upsertMaybe(client, SUPABASE_TABLES.graphs, graphRows, "id");
        await deleteMissingRows(client, SUPABASE_TABLES.graphs, config.appKey, graphRows.map(row => row.id));

        if (remoteCapabilities.shiftType) {
            try {
                await upsertMaybe(client, SUPABASE_TABLES.shifts, shiftRows, "id");
            } catch (error) {
                if (!isMissingColumnError(error, "shift_type")) {
                    throw error;
                }
                remoteCapabilities.shiftType = false;
                const fallbackShiftRows = shiftRows.map(({ shift_type, ...row }) => row);
                await upsertMaybe(client, SUPABASE_TABLES.shifts, fallbackShiftRows, "id");
            }
        } else {
            const fallbackShiftRows = shiftRows.map(({ shift_type, ...row }) => row);
            await upsertMaybe(client, SUPABASE_TABLES.shifts, fallbackShiftRows, "id");
        }
        await deleteMissingRows(client, SUPABASE_TABLES.shifts, config.appKey, shiftRows.map(row => row.id));

        await upsertMaybe(client, SUPABASE_TABLES.absences, absenceRows, "id");
        await deleteMissingRows(client, SUPABASE_TABLES.absences, config.appKey, absenceRows.map(row => row.id));

        await upsertMaybe(client, SUPABASE_TABLES.overrides, overrideRows, "id");
        await deleteMissingRows(client, SUPABASE_TABLES.overrides, config.appKey, overrideRows.map(row => row.id));
    } catch (error) {
        console.error("Ошибка синхронизации с Supabase", error);
    }
}

function queueRemoteSave() {
    if (!isSupabaseConfigured() || !hasRemoteStateLoaded || isReadOnlyMode) {
        return;
    }

    clearTimeout(remoteSaveTimer);
    remoteSaveTimer = setTimeout(() => {
        remoteSavePromise = remoteSavePromise.then(() => syncStateToSupabase());
    }, 350);
}

function save(options = {}) {
    writeLocalSnapshot();
    if (!options.skipRemote) {
        queueRemoteSave();
    }
}

function getActiveGraph() {
    const source = getVisibleGraphs();
    return source.find(graph => graph.id === activeGraphId) || source[0] || null;
}

function getVisibleGraphs() {
    if (isReadOnlyMode && readOnlyGraphId) {
        return graphs.filter(graph => graph.id === readOnlyGraphId);
    }
    return graphs;
}

function getPattern(type) {
    return type === "12"
        ? [
            { hours: 12, night: 0, worked: true },
            { hours: 4, night: 2, worked: true },
            { hours: 8, night: 6, worked: true },
            { hours: 0, night: 0, worked: false }
        ]
        : [
            { hours: 16, night: 2, worked: true },
            { hours: 8, night: 6, worked: true },
            { hours: 0, night: 0, worked: false },
            { hours: 0, night: 0, worked: false }
        ];
}

function getShiftPhaseOffset(smenaIndex) {
    const alignedOffsets = [2, 1, 0, 3];
    return alignedOffsets[smenaIndex] ?? mod(smenaIndex, 4);
}

function getCellSchedule(graph, smena, smenaIndex, dateStr) {
    if (normalizeShiftKind(smena?.shiftKind) === "reserve") {
        const parts = parseIsoDateParts(dateStr);
        const weekend = [0, 6].includes(getWeekday(parts.year, parts.month - 1, parts.day));
        const holiday = isHoliday(dateStr);
        if (weekend || holiday) {
            return { hours: 0, night: 0, worked: false };
        }
        return { hours: 8, night: 0, worked: true };
    }

    const pattern = getPattern(graph.type);
    const startDay = toUtcDayNumber(graph.firstShiftDate);
    const currentDay = toUtcDayNumber(dateStr);
    const diffDays = currentDay - startDay;
    const phase = mod(diffDays + getShiftPhaseOffset(smenaIndex), pattern.length);
    return pattern[phase];
}

function findAbsence(graphId, smenaId, dateStr) {
    return absences.find(item => item.graphId === graphId && item.smenaId === smenaId && item.start <= dateStr && item.end >= dateStr) || null;
}

function findOverride(graphId, smenaId, dateStr) {
    return cellOverrides.find(item => item.graphId === graphId && item.smenaId === smenaId && item.date === dateStr) || null;
}

function getAbsenceCode(type) {
    if (type === "Срочный больничный") return "СБ";
    if (type === "Больничный") return "БЛ";
    if (type === "Отгул") return "ОГ";
    return "ОТ";
}

function getOverrideAppliedState(override, fallbackSchedule) {
    if (!override) return null;
    switch (override.mode) {
        case "off":
            return { hours: 0, night: 0, worked: false, code: "", absence: null, manualClass: "" };
        case "work16":
            return { hours: 16, night: 2, worked: true, code: "", absence: null, manualClass: "manual-work-cell" };
        case "work8":
            return { hours: 8, night: 6, worked: true, code: "", absence: null, manualClass: "manual-work-cell" };
        case "work12":
            return { hours: 12, night: 0, worked: true, code: "", absence: null, manualClass: "manual-work-cell" };
        case "work24":
            return { hours: 24, night: 8, worked: true, code: "", absence: null, manualClass: "manual-work-cell" };
        case "work8day":
            return { hours: 8, night: 0, worked: true, code: "", absence: null, manualClass: "manual-work-cell" };
        case "custom":
            return {
                hours: clamp(parseInt(override.hours, 10) || 0, 0, 24),
                night: clamp(parseInt(override.night, 10) || 0, 0, 24),
                worked: clamp(parseInt(override.hours, 10) || 0, 0, 24) > 0,
                code: "",
                absence: null,
                manualClass: "manual-work-cell"
            };
        case "vacation":
            return { hours: 0, night: 0, worked: false, code: "ОТ", absence: { type: "Отпуск" }, manualClass: "" };
        case "sick":
            return { hours: 0, night: 0, worked: false, code: "БЛ", absence: { type: "Больничный" }, manualClass: "" };
        case "urgent_sick":
            return { hours: 0, night: 0, worked: false, code: "СБ", absence: { type: "Срочный больничный" }, manualClass: "" };
        case "dayoff":
            return { hours: 0, night: 0, worked: false, code: "ОГ", absence: { type: "Отгул" }, manualClass: "" };
        default:
            return {
                hours: fallbackSchedule.hours,
                night: fallbackSchedule.night,
                worked: fallbackSchedule.hours > 0,
                code: "",
                absence: null,
                manualClass: ""
            };
    }
}

function getProductionMonthStats(year, monthIndex) {
    const daysInMonth = getDaysInMonth(year, monthIndex);
    let workDays = 0;
    let workHours = 0;
    let offDays = 0;

    for (let day = 1; day <= daysInMonth; day += 1) {
        const dateStr = buildIsoDate(year, monthIndex, day);
        const weekend = [0, 6].includes(getWeekday(year, monthIndex, day));
        const holiday = isHoliday(dateStr);
        if (weekend || holiday) {
            offDays += 1;
        } else {
            workDays += 1;
            workHours += 8;
        }
    }

    return { workDays, workHours, offDays };
}

function buildMonthRowData(graph, smena, smenaIndex, monthIndex) {
    const daysInMonth = getDaysInMonth(currentYear, monthIndex);
    const rowStats = {
        workedDays: 0,
        hours: 0,
        night: 0
    };

    const cells = [];

    for (let day = 1; day <= 31; day += 1) {
        if (day > daysInMonth) {
            cells.push({ kind: "empty" });
            continue;
        }

        const dateStr = buildIsoDate(currentYear, monthIndex, day);
        const weekend = [0, 6].includes(getWeekday(currentYear, monthIndex, day));
        const holiday = isHoliday(dateStr);
        const preholiday = isPreholiday(dateStr);
        const schedule = getCellSchedule(graph, smena, smenaIndex, dateStr);
        const rangeAbsence = findAbsence(graph.id, smena.id, dateStr);
        const override = findOverride(graph.id, smena.id, dateStr);

        let hours = schedule.hours;
        let night = schedule.night;
        let worked = hours > 0;
        let code = "";
        let absence = rangeAbsence;
        let manualClass = "";

        if (rangeAbsence) {
            hours = 0;
            night = 0;
            worked = false;
            code = getAbsenceCode(rangeAbsence.type);
        }

        if (override) {
            const applied = getOverrideAppliedState(override, schedule);
            hours = applied.hours;
            night = applied.night;
            worked = applied.worked;
            code = applied.code;
            absence = applied.absence;
            manualClass = applied.manualClass;
        }

        if (hours > 0) {
            rowStats.workedDays += 1;
            rowStats.hours += hours;
            rowStats.night += night;
        }

        cells.push({
            kind: "day",
            dateStr,
            weekend,
            holiday,
            preholiday,
            absence,
            hours,
            night,
            code,
            worked,
            graphId: graph.id,
            smenaId: smena.id,
            smenaName: smena.name,
            override,
            manualClass,
            isManual: Boolean(override),
            shiftKind: normalizeShiftKind(smena.shiftKind)
        });
    }

    return {
        monthIndex,
        monthName: MONTHS[monthIndex],
        smena,
        cells,
        rowStats,
        productionStats: getProductionMonthStats(currentYear, monthIndex)
    };
}

function getAnnualProductionStats() {
    const totals = { workDays: 0, workHours: 0, offDays: 0 };
    for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
        const current = getProductionMonthStats(currentYear, monthIndex);
        totals.workDays += current.workDays;
        totals.workHours += current.workHours;
        totals.offDays += current.offDays;
    }
    return totals;
}

function getAnnualNormHours(year = currentYear) {
    return ANNUAL_NORM_HOURS_BY_YEAR[year] || getAnnualProductionStats().workHours;
}

function buildAnnualStats(graph) {
    const baseNormHours = getAnnualNormHours(currentYear);
    return graph.smeny.map((smena, smenaIndex) => {
        const stats = {
            workedDays: 0,
            hours: 0,
            night: 0,
            absenceWorkingDays: 0,
            absenceHours: 0,
            plannedAbsenceDays: 0,
            shortAbsenceDays: 0
        };

        for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
            const rowData = buildMonthRowData(graph, smena, smenaIndex, monthIndex);
            stats.workedDays += rowData.rowStats.workedDays;
            stats.hours += rowData.rowStats.hours;
            stats.night += rowData.rowStats.night;

            rowData.cells.forEach((cell) => {
                if (cell.kind !== "day" || !cell.absence || cell.weekend || cell.holiday) {
                    return;
                }
                stats.absenceWorkingDays += 1;
                stats.absenceHours += 8;
                if (isShortAbsenceType(cell.absence.type)) {
                    stats.shortAbsenceDays += 1;
                } else {
                    stats.plannedAbsenceDays += 1;
                }
            });
        }

        const normHours = Math.max(0, baseNormHours - stats.absenceHours);
        const diffHours = stats.hours - normHours;
        return { smena, diffHours, normHours, baseNormHours, ...stats };
    });
}

function renderGraphs() {
    const container = document.getElementById("graphsContainer");
    const visibleGraphs = getVisibleGraphs();
    container.innerHTML = visibleGraphs.map((graph) => `
        <div class="graph-card ${graph.id === activeGraphId ? "active" : ""}" onclick="selectGraph('${escapeHtml(graph.id)}')">
            <div class="d-flex justify-content-between gap-2 align-items-start">
                <div>
                    <div class="graph-name">${escapeHtml(graph.name)}</div>
                    <div class="graph-meta">${graph.type === "24" ? "24 часа" : "12 часов"} · ${graph.smeny.length} смен</div>
                </div>
                <div class="graph-card-actions">
                    <button class="btn btn-sm btn-link text-primary p-0 graph-link-btn" onclick="openShareModal('${escapeHtml(graph.id)}', event)" title="Ссылка только для чтения">
                        <i class="bi bi-share"></i>
                    </button>
                    ${isReadOnlyMode ? "" : `
                    <button class="btn btn-sm btn-link text-danger p-0 graph-delete-btn" onclick="deleteGraph(event, '${escapeHtml(graph.id)}')" title="Удалить график">
                        <i class="bi bi-x-circle"></i>
                    </button>`}
                </div>
            </div>
            <div class="mt-2">
                <label class="small text-muted d-block mb-1">Старт цикла</label>
                <input
                    type="date"
                    class="form-control form-control-sm ${isReadOnlyMode ? "readonly-input" : ""}"
                    value="${escapeHtml(graph.firstShiftDate)}"
                    onclick="event.stopPropagation()"
                    onchange="updateGraphDate('${escapeHtml(graph.id)}', this.value)"
                    ${isReadOnlyMode ? "disabled" : ""}
                >
            </div>
        </div>
    `).join("");
}

function renderSmeny() {
    const graph = getActiveGraph();
    const container = document.getElementById("smenyList");
    if (!graph) {
        container.innerHTML = "";
        return;
    }

    container.innerHTML = graph.smeny.map((smena, index) => `
        <div class="employee-item employee-item-editable ${normalizeShiftKind(smena.shiftKind) === "reserve" ? "employee-item-reserve" : ""}">
            <div class="employee-badge">${index + 1}</div>
            <div class="employee-main-fields">
                <input
                    type="text"
                    class="form-control form-control-sm ${isReadOnlyMode ? "readonly-input" : ""}"
                    value="${escapeHtml(smena.name)}"
                    onchange="renameSmena('${escapeHtml(smena.id)}', this.value)"
                    onblur="renameSmena('${escapeHtml(smena.id)}', this.value)"
                    ${isReadOnlyMode ? "disabled" : ""}
                >
                <div class="employee-kind-row">
                    <select
                        class="form-select form-select-sm shift-kind-select ${isReadOnlyMode ? "readonly-input" : ""}"
                        onchange="updateShiftKind('${escapeHtml(smena.id)}', this.value)"
                        ${isReadOnlyMode ? "disabled" : ""}
                    >
                        <option value="regular" ${normalizeShiftKind(smena.shiftKind) === "regular" ? "selected" : ""}>Основная</option>
                        <option value="reserve" ${normalizeShiftKind(smena.shiftKind) === "reserve" ? "selected" : ""}>Подменная</option>
                    </select>
                    <span class="shift-kind-badge ${normalizeShiftKind(smena.shiftKind) === "reserve" ? "reserve" : "regular"}">${getShiftKindLabel(smena.shiftKind)}</span>
                </div>
            </div>
            ${isReadOnlyMode ? "" : `
            <button onclick="deleteSmena(event, '${escapeHtml(smena.id)}')" class="btn btn-sm text-danger" title="Удалить смену">
                <i class="bi bi-trash"></i>
            </button>`}
        </div>
    `).join("");
}

function renderAbsences() {
    const graph = getActiveGraph();
    const plannedContainer = document.getElementById("plannedAbsencesList");
    const shortContainer = document.getElementById("shortAbsencesList");
    if (!graph) {
        if (plannedContainer) plannedContainer.innerHTML = "";
        if (shortContainer) shortContainer.innerHTML = "";
        return;
    }

    const list = absences.filter(item => item.graphId === graph.id);
    const planned = list.filter(item => !isShortAbsenceType(item.type));
    const short = list.filter(item => isShortAbsenceType(item.type));

    const renderAbsenceItems = (items, emptyText) => {
        if (!items.length) {
            return `<div class="empty-note">${emptyText}</div>`;
        }

        return items.map((absence) => {
            const smena = graph.smeny.find(item => item.id === absence.smenaId);
            const meta = getAbsenceVisualMeta(absence.type);
            return `
                <div class="employee-item small ${meta.className}">
                    <div><b>${meta.icon} ${escapeHtml(absence.type)}</b></div>
                    <div>${escapeHtml(smena?.name || "Удаленная смена")}</div>
                    <div>${escapeHtml(absence.start)} — ${escapeHtml(absence.end)}</div>
                    ${isReadOnlyMode ? "" : `<button onclick="deleteAbsence(event, '${escapeHtml(absence.id)}')" class="btn btn-sm btn-link text-danger p-0 mt-1">Удалить</button>`}
                </div>
            `;
        }).join("");
    };

    if (plannedContainer) {
        plannedContainer.innerHTML = renderAbsenceItems(planned, "Отвлечения не добавлены.");
    }
    if (shortContainer) {
        shortContainer.innerHTML = renderAbsenceItems(short, "Кратковременные отвлечения не добавлены.");
    }
}

function renderActiveGraphInfo(graph) {
    const container = document.getElementById("activeGraphInfo");
    const modeBanner = document.getElementById("modeBanner");
    if (!graph) {
        container.innerHTML = "";
        modeBanner.classList.add("d-none");
        return;
    }

    const patternLabel = graph.type === "24" ? "24 часа" : "12 часов";
    container.innerHTML = `
        <div class="info-strip">
            <span class="info-strip-name">${escapeHtml(graph.name)}</span>
            <span class="info-strip-meta">${patternLabel}</span>
            <span class="info-strip-meta">Старт: ${escapeHtml(graph.firstShiftDate)}</span>
            <span class="info-strip-meta">Смен: ${graph.smeny.length}</span>
            <span class="info-strip-meta">Год: ${currentYear}${isLeapYear(currentYear) ? " · високосный" : ""}</span>
            ${isReadOnlyMode ? '<span class="info-strip-badge"><i class="bi bi-eye me-1"></i>Только чтение</span>' : ''}
        </div>
    `;

    if (isReadOnlyMode) {
        modeBanner.textContent = "Открыт режим чтения: редактирование отключено, доступен только просмотр и экспорт текущего графика.";
        modeBanner.classList.remove("d-none");
    } else {
        modeBanner.classList.add("d-none");
    }
}

function renderAnnualSummary(graph) {
    const container = document.getElementById("annualSummary");
    if (!graph) {
        container.innerHTML = "";
        return;
    }

    const annualStats = buildAnnualStats(graph);
    container.innerHTML = annualStats.map((item) => {
        const diffLabel = item.diffHours > 0
            ? `<span class="summary-pill over"><i class="bi bi-arrow-up-right"></i>Переработка: +${item.diffHours} ч</span>`
            : item.diffHours < 0
                ? `<span class="summary-pill under"><i class="bi bi-arrow-down-right"></i>Недоработка: ${item.diffHours} ч</span>`
                : `<span class="summary-pill norm"><i class="bi bi-check2"></i>Норма выполнена</span>`;

        return `
            <div class="summary-card">
                <div class="summary-card-title">${escapeHtml(item.smena.name)}</div>
                <div class="summary-card-values">
                    <div><span class="summary-value">${item.workedDays}</span><span class="summary-label">дней</span></div>
                    <div><span class="summary-value">${item.hours}</span><span class="summary-label">часов</span></div>
                    <div><span class="summary-value">${item.night}</span><span class="summary-label">ночных</span></div>
                </div>
                <div class="summary-card-extra">
                    <span class="summary-pill norm"><i class="bi bi-clock"></i>Норма: ${item.normHours} ч</span>
                    <span class="summary-pill norm"><i class="bi bi-calendar-x"></i>Отвлечения: ${item.absenceHours} ч</span>
                    ${item.plannedAbsenceDays ? `<span class="summary-pill norm"><i class="bi bi-suitcase"></i>Плановые: ${item.plannedAbsenceDays} дн.</span>` : ""}
                    ${item.shortAbsenceDays ? `<span class="summary-pill norm"><i class="bi bi-lightning-charge"></i>Краткосрочные: ${item.shortAbsenceDays} дн.</span>` : ""}
                    ${diffLabel}
                </div>
            </div>
        `;
    }).join("");
}

function renderTable(graph) {
    const wrapper = document.getElementById("scheduleTableWrapper");
    const fullscreenWrapper = document.getElementById("fullscreenTableWrapper");
    if (!graph) {
        wrapper.innerHTML = "";
        fullscreenWrapper.innerHTML = "";
        return;
    }

    const rows = [];
    for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
        graph.smeny.forEach((smena, smenaIndex) => {
            rows.push(buildMonthRowData(graph, smena, smenaIndex, monthIndex));
        });
    }

    const headerDays = Array.from({ length: 31 }, (_, index) => `<th class="day-head">${index + 1}</th>`).join("");
    const bodyHtml = rows.map((row) => {
        const dayCells = row.cells.map(renderDayCell).join("");
        return `
            <tr>
                <td class="sticky-col month-col">${escapeHtml(row.monthName)}</td>
                <td class="sticky-col shift-col">${escapeHtml(row.smena.name)}</td>
                ${dayCells}
                <td class="stat-col stat-days">${row.rowStats.workedDays}</td>
                <td class="stat-col stat-hours">
                    <div class="stat-stack">
                        <div>${row.rowStats.hours}</div>
                        <div class="stat-sub">${row.rowStats.night}</div>
                    </div>
                </td>
                <td class="prod-col">${row.productionStats.workDays}</td>
                <td class="prod-col">${row.productionStats.workHours}</td>
                <td class="prod-col">${row.productionStats.offDays}</td>
            </tr>
        `;
    }).join("");

    const tableHtml = `
        <table class="schedule-table table table-bordered align-middle mb-0">
            <thead>
                <tr>
                    <th rowspan="2" class="sticky-col month-col head-main">Месяц</th>
                    <th rowspan="2" class="sticky-col shift-col head-main">Смена</th>
                    <th colspan="31" class="head-group">Часов за день, в том числе ночных</th>
                    <th colspan="2" class="head-group">По графику</th>
                    <th colspan="3" class="head-group">По производств. календарю</th>
                </tr>
                <tr>
                    ${headerDays}
                    <th class="mini-head">дней</th>
                    <th class="mini-head">часов<br><span class="mini-head-sub">ночных</span></th>
                    <th class="mini-head">дней</th>
                    <th class="mini-head">часов</th>
                    <th class="mini-head">нераб.</th>
                </tr>
            </thead>
            <tbody>${bodyHtml}</tbody>
        </table>
    `;

    wrapper.innerHTML = tableHtml;
    fullscreenWrapper.innerHTML = tableHtml;
    applyCellHeight();
}

function renderDayCell(cell) {
    if (cell.kind === "empty") {
        return `
            <td class="day-cell day-cell-empty">
                <div class="cell-lines cell-lines-empty"><span></span></div>
            </td>
        `;
    }

    const classes = ["day-cell"];
    if (!isReadOnlyMode) classes.push("day-cell-clickable");
    if (cell.isManual) classes.push("is-manual");
    if (cell.manualClass) classes.push(cell.manualClass);

    if (cell.absence) {
        const absenceClass = getAbsenceVisualMeta(cell.absence.type).cellClass;
        classes.push("is-absence", absenceClass);
    } else {
        if (cell.worked) {
            classes.push("is-worked");
        } else {
            classes.push("is-off");
        }

        if (cell.holiday) {
            classes.push("is-holiday");
        } else if (cell.preholiday) {
            classes.push("is-preholiday");
        } else if (cell.weekend) {
            classes.push("is-weekend");
        }
    }

    const tooltipParts = [cell.dateStr];
    if (cell.worked) tooltipParts.push(`Часы: ${cell.hours}`, `Ночные: ${cell.night}`);
    if (cell.absence) tooltipParts.push(cell.absence.type);
    if (cell.isManual) tooltipParts.push("Ручная корректировка");
    if (cell.holiday) tooltipParts.push("Праздник");
    else if (cell.weekend) tooltipParts.push("Выходной день календаря");

    const topValue = cell.absence ? cell.code : (cell.hours || "");
    const bottomValue = cell.absence ? "" : (cell.night || "");
    const tag = "";
    const action = !isReadOnlyMode
        ? `onclick="openCellEditModal('${escapeHtml(cell.graphId)}','${escapeHtml(cell.smenaId)}','${escapeHtml(cell.dateStr)}')"`
        : "";

    return `
        <td class="${classes.join(" ")}" title="${escapeHtml(tooltipParts.join(" · "))}" ${action}>
            <div class="cell-lines">
                <div class="cell-top">${topValue || tag}</div>
                <div class="cell-bottom">${bottomValue}</div>
            </div>
        </td>
    `;
}

function getDayAdjustmentPriority(cell) {
    if (cell.holiday) return 4;
    if (cell.weekend) return 3;
    if (cell.preholiday) return 2;
    return 1;
}

function createAutoOverride(graphId, smenaId, dateStr, mode, hours = 0, night = 0, absenceType = "") {
    return {
        id: buildAutoBalanceOverrideId(graphId, smenaId, dateStr),
        graphId,
        smenaId,
        date: dateStr,
        mode,
        hours,
        night,
        absenceType
    };
}

function removeAutoBalanceOverrides(graphId = "") {
    cellOverrides = cellOverrides.filter((item) => {
        if (!isAutoBalanceOverride(item)) return true;
        return graphId ? item.graphId !== toId(graphId) : false;
    });
}

function collectGraphCells(graph, smena, smenaIndex) {
    const result = [];
    for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
        const rowData = buildMonthRowData(graph, smena, smenaIndex, monthIndex);
        rowData.cells.forEach((cell) => {
            if (cell.kind === "day") {
                result.push(cell);
            }
        });
    }
    return result;
}

function getAdditionPlan(graphType, remaining) {
    if (remaining <= 0) {
        return { mode: "custom", hours: 0, night: 0 };
    }

    if (remaining >= 24) return { mode: "work24", hours: 24, night: 8 };
    if (graphType === "12") {
        if (remaining >= 12) return { mode: "work12", hours: 12, night: 0 };
        if (remaining >= 8) return { mode: "work8day", hours: 8, night: 0 };
        return { mode: "custom", hours: remaining, night: 0 };
    }

    if (remaining >= 16) return { mode: "work16", hours: 16, night: 2 };
    if (remaining >= 8) return { mode: "work8day", hours: 8, night: 0 };
    return { mode: "custom", hours: remaining, night: 0 };
}

function getOverrideForBalancingCandidate(cell) {
    if (!cell.override) return null;
    return isAutoBalanceOverride(cell.override) ? cell.override : null;
}

function getResolvedCellStateForDate(graph, smena, smenaIndex, dateStr) {
    const parts = parseIsoDateParts(dateStr);
    if (!parts) return null;

    const weekend = [0, 6].includes(getWeekday(parts.year, parts.month - 1, parts.day));
    const holiday = isHoliday(dateStr);
    const preholiday = isPreholiday(dateStr);
    const schedule = getCellSchedule(graph, smena, smenaIndex, dateStr);
    const rangeAbsence = findAbsence(graph.id, smena.id, dateStr);
    const override = findOverride(graph.id, smena.id, dateStr);

    let hours = schedule.hours;
    let night = schedule.night;
    let worked = hours > 0;
    let code = "";
    let absence = rangeAbsence;
    let manualClass = "";

    if (rangeAbsence) {
        hours = 0;
        night = 0;
        worked = false;
        code = getAbsenceCode(rangeAbsence.type);
    }

    if (override) {
        const applied = getOverrideAppliedState(override, schedule);
        hours = applied.hours;
        night = applied.night;
        worked = applied.worked;
        code = applied.code;
        absence = applied.absence;
        manualClass = applied.manualClass;
    }

    return {
        kind: "day",
        dateStr,
        weekend,
        holiday,
        preholiday,
        absence,
        hours,
        night,
        code,
        worked,
        graphId: graph.id,
        smenaId: smena.id,
        smenaName: smena.name,
        override,
        manualClass,
        isManual: Boolean(override),
        shiftKind: normalizeShiftKind(smena.shiftKind)
    };
}

function getWorkedRegularShiftIdsForDate(graph, dateStr, excludeSmenaId = "") {
    const ids = [];
    graph.smeny.forEach((smena, smenaIndex) => {
        if (toId(smena.id) === toId(excludeSmenaId)) return;
        if (normalizeShiftKind(smena.shiftKind) !== "regular") return;
        const state = getResolvedCellStateForDate(graph, smena, smenaIndex, dateStr);
        if (state?.worked) {
            ids.push(smena.id);
        }
    });
    return ids;
}

function getRegularShiftEntries(graph) {
    const stats = buildAnnualStats(graph);
    return graph.smeny
        .map((smena, index) => ({ smena, index, stats: stats[index] }))
        .filter((entry) => entry && normalizeShiftKind(entry.smena.shiftKind) === "regular");
}

function chooseCoverageKeeperForDate(workedEntries) {
    return workedEntries.slice().sort((a, b) => {
        const aManual = a.override && !isAutoBalanceOverride(a.override) ? 0 : 1;
        const bManual = b.override && !isAutoBalanceOverride(b.override) ? 0 : 1;
        if (aManual !== bManual) return aManual - bManual;

        const aDiff = Number(a.entry?.stats?.diffHours || 0);
        const bDiff = Number(b.entry?.stats?.diffHours || 0);
        if (aDiff !== bDiff) return aDiff - bDiff;

        if ((b.hours || 0) !== (a.hours || 0)) return (b.hours || 0) - (a.hours || 0);
        return String(a.smenaName || "").localeCompare(String(b.smenaName || ""));
    })[0] || null;
}

function collapseDailyRegularCoverage(graph, shiftEntries) {
    if (!graph || !shiftEntries.length) return;

    for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
        const daysInMonth = getDaysInMonth(currentYear, monthIndex);
        for (let day = 1; day <= daysInMonth; day += 1) {
            const dateStr = buildIsoDate(currentYear, monthIndex, day);
            const workedEntries = shiftEntries
                .map((entry) => {
                    const state = getResolvedCellStateForDate(graph, entry.smena, entry.index, dateStr);
                    return state?.worked ? { ...state, entry } : null;
                })
                .filter(Boolean);

            if (workedEntries.length <= 1) {
                continue;
            }

            const keeper = chooseCoverageKeeperForDate(workedEntries);
            if (!keeper) {
                continue;
            }

            const totalHours = workedEntries.reduce((sum, item) => sum + (item.hours || 0), 0);
            const totalNight = workedEntries.reduce((sum, item) => sum + (item.night || 0), 0);
            const keeperMode = getAutoModeForHours(totalHours, totalNight);

            for (const item of workedEntries) {
                item.entry.stats.hours -= item.hours || 0;
                item.entry.stats.diffHours -= item.hours || 0;
            }

            upsertOverride(createAutoOverride(graph.id, keeper.smenaId, dateStr, keeperMode.mode, keeperMode.hours, keeperMode.night));
            keeper.entry.stats.hours += keeperMode.hours || 0;
            keeper.entry.stats.diffHours += keeperMode.hours || 0;

            workedEntries.forEach((item) => {
                if (toId(item.smenaId) === toId(keeper.smenaId)) return;
                upsertOverride(createAutoOverride(graph.id, item.smenaId, dateStr, "off"));
            });
        }
    }
}

function getAutoModeForHours(hours, night) {
    const safeHours = clamp(parseInt(hours, 10) || 0, 0, 24);
    const safeNight = clamp(parseInt(night, 10) || 0, 0, 24);
    if (safeHours === 0) return { mode: "off", hours: 0, night: 0 };
    if (safeHours === 24 && safeNight === 8) return { mode: "work24", hours: 24, night: 8 };
    if (safeHours === 16 && safeNight === 2) return { mode: "work16", hours: 16, night: 2 };
    if (safeHours === 8 && safeNight === 6) return { mode: "work8", hours: 8, night: 6 };
    if (safeHours === 12 && safeNight === 0) return { mode: "work12", hours: 12, night: 0 };
    if (safeHours === 8 && safeNight === 0) return { mode: "work8day", hours: 8, night: 0 };
    return { mode: "custom", hours: safeHours, night: safeNight };
}

function canUseCellForBalancing(cell) {
    return !cell.absence && !cell.override;
}

function applyOverworkBalancing(graph, smena, diffHours, workedCells) {
    let remaining = diffHours;
    const candidates = workedCells
        .filter((cell) => canUseCellForBalancing(cell) && cell.hours > 0)
        .sort((a, b) => {
            const priorityDiff = getDayAdjustmentPriority(a) - getDayAdjustmentPriority(b);
            if (priorityDiff !== 0) return priorityDiff;
            if (b.hours !== a.hours) return b.hours - a.hours;
            return String(b.dateStr).localeCompare(String(a.dateStr));
        });

    for (const cell of candidates) {
        if (remaining <= 0) break;

        if (remaining >= cell.hours) {
            upsertOverride(createAutoOverride(graph.id, smena.id, cell.dateStr, "off"));
            remaining -= cell.hours;
            continue;
        }

        const newHours = Math.max(0, cell.hours - remaining);
        const newNight = Math.min(cell.night || 0, newHours);
        const nextMode = getAutoModeForHours(newHours, newNight);
        upsertOverride(createAutoOverride(graph.id, smena.id, cell.dateStr, nextMode.mode, nextMode.hours, nextMode.night));
        remaining = 0;
        break;
    }

    return remaining;
}

function applyUnderworkBalancing(graph, smena, deficitHours, freeCells) {
    let remaining = deficitHours;
    const candidates = freeCells
        .filter((cell) => canUseCellForBalancing(cell) && !cell.worked)
        .filter((cell) => getWorkedRegularShiftIdsForDate(graph, cell.dateStr, smena.id).length === 0)
        .sort((a, b) => {
            const priorityDiff = getDayAdjustmentPriority(a) - getDayAdjustmentPriority(b);
            if (priorityDiff !== 0) return priorityDiff;
            return String(a.dateStr).localeCompare(String(b.dateStr));
        });

    for (const cell of candidates) {
        if (remaining <= 0) break;
        const plan = getAdditionPlan(graph.type, remaining);
        upsertOverride(createAutoOverride(graph.id, smena.id, cell.dateStr, plan.mode, plan.hours, plan.night));
        remaining -= plan.hours;
    }

    return remaining;
}

function transferWorkBetweenShifts(graph, sourceEntry, targetEntry, cell) {
    const sourceDiff = sourceEntry.stats.diffHours;
    const targetDeficit = Math.abs(targetEntry.stats.diffHours);
    if (sourceDiff < cell.hours || targetDeficit < cell.hours) {
        return false;
    }

    const workedIds = getWorkedRegularShiftIdsForDate(graph, cell.dateStr);
    if (workedIds.some((id) => toId(id) !== toId(sourceEntry.smena.id))) {
        return false;
    }

    const targetState = getResolvedCellStateForDate(graph, targetEntry.smena, targetEntry.index, cell.dateStr);
    if (!targetState || targetState.absence || targetState.worked || targetState.override) {
        return false;
    }

    upsertOverride(createAutoOverride(graph.id, sourceEntry.smena.id, cell.dateStr, "off"));
    const targetMode = getAutoModeForHours(cell.hours, cell.night);
    upsertOverride(createAutoOverride(graph.id, targetEntry.smena.id, cell.dateStr, targetMode.mode, targetMode.hours, targetMode.night));

    sourceEntry.stats.diffHours -= cell.hours;
    sourceEntry.stats.hours -= cell.hours;
    targetEntry.stats.diffHours += cell.hours;
    targetEntry.stats.hours += cell.hours;
    return true;
}

function rebalanceGraphToAnnualNorm(graph) {
    if (!graph) return;

    removeAutoBalanceOverrides(graph.id);

    const shiftEntries = getRegularShiftEntries(graph);
    collapseDailyRegularCoverage(graph, shiftEntries);

    for (const sourceEntry of shiftEntries.filter((entry) => entry.stats?.diffHours > 0)) {
        let sourceCells = collectGraphCells(graph, sourceEntry.smena, sourceEntry.index)
            .filter((cell) => canUseCellForBalancing(cell) && cell.worked && cell.hours > 0)
            .sort((a, b) => {
                const priorityDiff = getDayAdjustmentPriority(a) - getDayAdjustmentPriority(b);
                if (priorityDiff !== 0) return priorityDiff;
                if (b.hours !== a.hours) return b.hours - a.hours;
                return String(b.dateStr).localeCompare(String(a.dateStr));
            });

        for (const cell of sourceCells) {
            if (sourceEntry.stats.diffHours <= 0) break;
            const targets = shiftEntries
                .filter((entry) => toId(entry.smena.id) !== toId(sourceEntry.smena.id) && entry.stats?.diffHours < 0)
                .sort((a, b) => a.stats.diffHours - b.stats.diffHours);

            for (const targetEntry of targets) {
                if (transferWorkBetweenShifts(graph, sourceEntry, targetEntry, cell)) {
                    break;
                }
            }
        }
    }

    graph.smeny.forEach((smena, smenaIndex) => {
        if (normalizeShiftKind(smena.shiftKind) !== "regular") return;
        const initialStats = buildAnnualStats(graph)[smenaIndex];
        if (!initialStats || initialStats.diffHours === 0) return;

        const cells = collectGraphCells(graph, smena, smenaIndex);
        if (initialStats.diffHours > 0) {
            applyOverworkBalancing(graph, smena, initialStats.diffHours, cells);
        }
    });

    const finalShiftEntries = getRegularShiftEntries(graph);
    collapseDailyRegularCoverage(graph, finalShiftEntries);

    const finalStats = buildAnnualStats(graph).filter((item, index) => normalizeShiftKind(graph.smeny[index]?.shiftKind) === "regular");
    const unresolved = finalStats.filter((item) => item.diffHours !== 0);
    if (unresolved.length) {
        console.warn("Не удалось полностью уравнять часы по норме без добавления дополнительных смен в те же даты. Для недоработки используйте подменную смену или ручную корректировку.", unresolved);
    }
}

function renderAll() {
    const graph = getActiveGraph();
    document.getElementById("currentYear").value = currentYear;
    document.getElementById("cellHeightDisplay").textContent = `${currentCellHeight}px`;
    renderGraphs();
    renderSmeny();
    renderAbsences();
    renderActiveGraphInfo(graph);
    renderAnnualSummary(graph);
    renderTable(graph);
    applyUiMode();
}

function calculateAndRender() {
    const graph = getActiveGraph();
    if (!graph) {
        renderAll();
        return;
    }

    rebalanceGraphToAnnualNorm(graph);
    save();
    renderAll();
}

function applyUiMode() {
    document.body.classList.toggle("read-only-mode", isReadOnlyMode);
    if (isReadOnlyMode) {
        closeSidebar();
    }

    const editableIds = ["createGraphButton", "addShiftButton", "addAbsenceButton"];
    editableIds.forEach((id) => {
        const element = document.getElementById(id);
        if (element) {
            element.classList.toggle("readonly-hide", isReadOnlyMode);
        }
    });

    const yearInput = document.getElementById("currentYear");
    if (yearInput) {
        yearInput.disabled = isReadOnlyMode;
        yearInput.classList.toggle("readonly-input", isReadOnlyMode);
    }
}

function updateYear(value) {
    if (isReadOnlyMode) return;
    currentYear = clamp(parseInt(value, 10) || DEFAULT_YEAR, 2000, 2100);
    save();
    renderAll();
}

function selectGraph(id) {
    const targetId = toId(id);
    if (isReadOnlyMode && targetId !== readOnlyGraphId) {
        return;
    }
    activeGraphId = targetId;
    save();
    renderAll();
    closeSidebar();
}

function createGraphWithType(type) {
    if (isReadOnlyMode) return;
    const nameInput = document.getElementById("newGraphName");
    const dateInput = document.getElementById("newFirstShiftDate");
    const name = nameInput.value.trim() || "Новый график";
    const date = dateInput.value || `${currentYear}-01-01`;

    const newGraph = {
        id: uid("graph"),
        name,
        type: type === "12" ? "12" : "24",
        firstShiftDate: normalizeIsoDate(date, currentYear),
        smeny: createDefaultShifts(4)
    };

    graphs.push(newGraph);
    activeGraphId = newGraph.id;
    pendingCreatedGraphId = newGraph.id;
    save();
    renderAll();

    nameInput.value = "";
    dateInput.value = `${currentYear}-01-01`;

    bootstrap.Modal.getInstance(document.getElementById("createModal"))?.hide();
    openShareModal(newGraph.id);
}

function deleteGraph(event, id) {
    event.stopPropagation();
    if (isReadOnlyMode) return;
    if (graphs.length === 1) {
        alert("Нельзя удалить последний график.");
        return;
    }
    if (!confirm("Удалить этот график?")) {
        return;
    }

    graphs = graphs.filter(graph => graph.id !== toId(id));
    absences = absences.filter(item => item.graphId !== toId(id));
    cellOverrides = cellOverrides.filter(item => item.graphId !== toId(id));
    activeGraphId = graphs[0]?.id || "";
    save();
    renderAll();
}

function updateGraphDate(id, value) {
    if (isReadOnlyMode) return;
    const graph = graphs.find(item => item.id === toId(id));
    if (!graph) return;
    graph.firstShiftDate = normalizeIsoDate(value, currentYear);
    save();
    renderAll();
}

function addSmena() {
    if (isReadOnlyMode) return;
    const graph = getActiveGraph();
    if (!graph) return;
    graph.smeny.push({ id: uid("shift"), name: `Смена ${graph.smeny.length + 1}`, shiftKind: "regular" });
    save();
    renderAll();
}

function updateShiftKind(smenaId, value) {
    if (isReadOnlyMode) return;
    const graph = getActiveGraph();
    if (!graph) return;
    const smena = graph.smeny.find(item => item.id === toId(smenaId));
    if (!smena) return;
    smena.shiftKind = normalizeShiftKind(value);
    save();
    renderAll();
}

function renameSmena(smenaId, value) {
    if (isReadOnlyMode) return;
    const graph = getActiveGraph();
    if (!graph) return;
    const smena = graph.smeny.find(item => item.id === toId(smenaId));
    if (!smena) return;
    smena.name = value.trim() || smena.name;
    save();
    renderAll();
}

function deleteSmena(event, smenaId) {
    event.stopPropagation();
    if (isReadOnlyMode) return;
    const graph = getActiveGraph();
    if (!graph) return;
    if (graph.smeny.length === 1) {
        alert("Нельзя удалить последнюю смену.");
        return;
    }
    if (!confirm("Удалить эту смену?")) {
        return;
    }

    graph.smeny = graph.smeny.filter(item => item.id !== toId(smenaId));
    absences = absences.filter(item => item.smenaId !== toId(smenaId));
    cellOverrides = cellOverrides.filter(item => item.smenaId !== toId(smenaId));
    save();
    renderAll();
}

function addAbsenceModal(category = "planned") {
    const graph = getActiveGraph();
    if (!graph || !graph.smeny.length) {
        alert("Сначала добавьте хотя бы одну смену.");
        return;
    }

    currentAbsenceCategory = category === "short" ? "short" : "planned";

    document.getElementById("modalSmena").innerHTML = graph.smeny
        .map(item => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name)}</option>`)
        .join("");

    const typeSelect = document.getElementById("modalType");
    const titleNode = document.getElementById("absenceModalTitle");
    const options = currentAbsenceCategory === "short"
        ? ["Срочный больничный", "Отгул"]
        : ["Отпуск", "Больничный"];

    typeSelect.innerHTML = options.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`).join("");
    titleNode.textContent = currentAbsenceCategory === "short" ? "Добавить кратковременное отвлечение" : "Добавить отвлечение";

    document.getElementById("modalStart").value = "";
    document.getElementById("modalEnd").value = "";
    new bootstrap.Modal(document.getElementById("absenceModal")).show();
}

function saveAbsence() {
    if (isReadOnlyMode) return;
    const graph = getActiveGraph();
    if (!graph) return;

    const smenaId = toId(document.getElementById("modalSmena").value);
    const type = normalizeAbsenceType(document.getElementById("modalType").value);
    const start = document.getElementById("modalStart").value;
    const end = document.getElementById("modalEnd").value;

    if (!smenaId || !start || !end) {
        alert("Заполните все поля отвлечения.");
        return;
    }

    if (start > end) {
        alert("Дата начала не может быть позже даты окончания.");
        return;
    }

    absences.push(normalizeAbsence({
        id: uid("absence"),
        graphId: graph.id,
        smenaId,
        type,
        start,
        end
    }));

    save();
    renderAll();
    bootstrap.Modal.getInstance(document.getElementById("absenceModal"))?.hide();
}

function deleteAbsence(event, absenceId) {
    event.stopPropagation();
    if (isReadOnlyMode) return;
    if (!confirm("Удалить это отвлечение?")) {
        return;
    }
    absences = absences.filter(item => item.id !== toId(absenceId));
    save();
    renderAll();
}

function getReadOnlyLink(graphId) {
    const url = new URL(window.location.href);
    url.searchParams.set("mode", "read");
    url.searchParams.set("graph", graphId);
    return url.toString();
}

function openShareModal(graphId, event) {
    event?.stopPropagation?.();
    const graph = graphs.find(item => item.id === toId(graphId));
    if (!graph) return;
    document.getElementById("shareLinkField").value = getReadOnlyLink(graph.id);
    shareModalInstance ||= new bootstrap.Modal(document.getElementById("shareModal"));
    shareModalInstance.show();
}

async function copyShareLink() {
    const field = document.getElementById("shareLinkField");
    field.select();
    field.setSelectionRange(0, field.value.length);
    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(field.value);
        } else {
            document.execCommand("copy");
        }
        alert("Ссылка скопирована.");
    } catch {
        alert("Не удалось автоматически скопировать ссылку. Скопируйте её вручную.");
    }
}

function openCellEditModal(graphId, smenaId, dateStr) {
    if (isReadOnlyMode) return;
    const graph = graphs.find(item => item.id === toId(graphId));
    const smena = graph?.smeny.find(item => item.id === toId(smenaId));
    if (!graph || !smena) return;

    editingCellContext = { graphId: graph.id, smenaId: smena.id, date: dateStr };
    const existingOverride = findOverride(graph.id, smena.id, dateStr);
    document.getElementById("cellEditInfo").innerHTML = `
        <b>${escapeHtml(graph.name)}</b><br>
        ${escapeHtml(smena.name)} · ${escapeHtml(dateStr)}
    `;

    document.getElementById("manualMode").value = existingOverride?.mode || "auto";
    document.getElementById("manualHours").value = existingOverride?.hours ?? 0;
    document.getElementById("manualNight").value = existingOverride?.night ?? 0;
    handleManualModeChange();

    cellEditModalInstance ||= new bootstrap.Modal(document.getElementById("cellEditModal"));
    cellEditModalInstance.show();
}

function handleManualModeChange() {
    const mode = document.getElementById("manualMode").value;
    document.getElementById("manualHoursRow").classList.toggle("d-none", mode !== "custom");
}

function upsertOverride(override) {
    const normalized = normalizeOverride(override);
    const existingIndex = cellOverrides.findIndex((item) => item.graphId === normalized.graphId && item.smenaId === normalized.smenaId && item.date === normalized.date);
    if (existingIndex >= 0) {
        cellOverrides[existingIndex] = normalized;
    } else {
        cellOverrides.push(normalized);
    }
}

function clearCellOverride() {
    if (!editingCellContext || isReadOnlyMode) return;
    cellOverrides = cellOverrides.filter((item) => !(item.graphId === editingCellContext.graphId && item.smenaId === editingCellContext.smenaId && item.date === editingCellContext.date));
    save();
    renderAll();
    cellEditModalInstance?.hide();
}

function saveCellOverride() {
    if (!editingCellContext || isReadOnlyMode) return;
    const mode = document.getElementById("manualMode").value;
    if (mode === "auto") {
        clearCellOverride();
        return;
    }

    const override = {
        id: uid("override"),
        graphId: editingCellContext.graphId,
        smenaId: editingCellContext.smenaId,
        date: editingCellContext.date,
        mode,
        hours: 0,
        night: 0,
        absenceType: ""
    };

    if (mode === "custom") {
        override.hours = clamp(parseInt(document.getElementById("manualHours").value, 10) || 0, 0, 24);
        override.night = clamp(parseInt(document.getElementById("manualNight").value, 10) || 0, 0, 24);
    }
    if (mode === "vacation") override.absenceType = "Отпуск";
    if (mode === "sick") override.absenceType = "Больничный";
    if (mode === "urgent_sick") override.absenceType = "Срочный больничный";
    if (mode === "dayoff") override.absenceType = "Отгул";

    const existing = findOverride(editingCellContext.graphId, editingCellContext.smenaId, editingCellContext.date);
    if (existing) override.id = existing.id;
    upsertOverride(override);
    save();
    renderAll();
    cellEditModalInstance?.hide();
}

function increaseCellHeight() {
    currentCellHeight = clamp(currentCellHeight + CELL_HEIGHT_STEP, CELL_HEIGHT_MIN, CELL_HEIGHT_MAX);
    save();
    applyCellHeight();
}

function decreaseCellHeight() {
    currentCellHeight = clamp(currentCellHeight - CELL_HEIGHT_STEP, CELL_HEIGHT_MIN, CELL_HEIGHT_MAX);
    save();
    applyCellHeight();
}

function applyCellHeight() {
    const cellHeight = `${currentCellHeight}px`;
    document.documentElement.style.setProperty("--cell-height", cellHeight);
    document.querySelectorAll(".schedule-table").forEach((table) => {
        table.style.setProperty("--cell-height", cellHeight);
    });
    document.getElementById("cellHeightDisplay").textContent = cellHeight;
}

function openFullscreen() {
    const graph = getActiveGraph();
    if (!graph) return;
    renderTable(graph);
    document.getElementById("fullscreenTitle").textContent = `${graph.name} — весь экран`;
    currentZoom = 100;
    updateZoom();
    fullscreenModalInstance ||= new bootstrap.Modal(document.getElementById("fullscreenModal"));
    fullscreenModalInstance.show();
}

function zoomIn() {
    currentZoom = clamp(currentZoom + 10, 50, 300);
    updateZoom();
}

function zoomOut() {
    currentZoom = clamp(currentZoom - 10, 50, 300);
    updateZoom();
}

function zoomReset() {
    currentZoom = 100;
    updateZoom();
}

function updateZoom() {
    const container = document.getElementById("fullscreenContainer");
    container.style.transform = `scale(${currentZoom / 100})`;
    document.getElementById("zoomLevel").textContent = currentZoom;
}

function openSidebar() {
    if (window.innerWidth > 991) return;
    document.body.classList.add("sidebar-open");
}

function closeSidebar() {
    document.body.classList.remove("sidebar-open");
}

function handleResize() {
    if (window.innerWidth > 991) {
        closeSidebar();
    }
}

function parseReadOnlyMode() {
    const params = new URLSearchParams(window.location.search);
    const mode = params.get("mode");
    const graphId = toId(params.get("graph"));
    isReadOnlyMode = mode === "read";
    readOnlyGraphId = graphId;
}

function applyReadOnlySelection() {
    if (!isReadOnlyMode) return;
    const visibleGraphs = getVisibleGraphs();
    if (!visibleGraphs.length) {
        isReadOnlyMode = false;
        readOnlyGraphId = "";
        return;
    }
    activeGraphId = visibleGraphs[0].id;
}

window.addEventListener("load", async () => {
    parseReadOnlyMode();
    normalizeState();

    if (isSupabaseConfigured()) {
        const loaded = await loadFromSupabase();
        if (!loaded) {
            hasRemoteStateLoaded = true;
            if (!isReadOnlyMode) {
                queueRemoteSave();
            }
        }
    } else {
        hasRemoteStateLoaded = true;
    }

    applyReadOnlySelection();
    document.getElementById("newFirstShiftDate").value = `${currentYear}-01-01`;
    renderAll();

    document.getElementById("createModal").addEventListener("show.bs.modal", () => {
        document.getElementById("newFirstShiftDate").value = `${currentYear}-01-01`;
    });
});
