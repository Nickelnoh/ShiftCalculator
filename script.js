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

const STORAGE_KEYS = {
    year: "ssc_year",
    graphs: "ssc_graphs",
    activeId: "ssc_activeId",
    absences: "ssc_absences",
    cellHeight: "ssc_cellHeight"
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
let currentCellHeight = DEFAULT_CELL_HEIGHT;
let currentZoom = 100;
let fullscreenModalInstance = null;

const SUPABASE_TABLES = {
    settings: "app_settings",
    graphs: "graphs",
    shifts: "shifts",
    absences: "absences"
};

let supabaseClient = null;
let remoteSaveTimer = null;
let remoteSavePromise = Promise.resolve();
let hasRemoteStateLoaded = false;
let currentAppKey = "";

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

function createDefaultShifts(count = 4) {
    return Array.from({ length: count }, (_, index) => ({
        id: uid("shift"),
        name: `Смена ${index + 1}`
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
            name: normalizeShiftName(item?.name, itemIndex)
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

function normalizeAbsence(absence) {
    return {
        id: toId(absence?.id || uid("absence")),
        graphId: toId(absence?.graphId),
        smenaId: toId(absence?.smenaId),
        type: absence?.type === "Больничный" ? "Больничный" : "Отпуск",
        start: normalizeIsoDate(absence?.start, currentYear),
        end: normalizeIsoDate(absence?.end, currentYear)
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

    return {
        year: localYear,
        graphs: localGraphs,
        activeId: localActiveId,
        absences: localAbsences,
        cellHeight: localCellHeight
    };
}

function buildStateSnapshot() {
    return {
        year: currentYear,
        graphs: graphs.map((graph) => ({
            id: graph.id,
            name: graph.name,
            type: graph.type,
            firstShiftDate: graph.firstShiftDate,
            smeny: graph.smeny.map((smena) => ({
                id: smena.id,
                name: smena.name
            }))
        })),
        activeId: activeGraphId,
        absences: absences.map((item) => ({
            id: item.id,
            graphId: item.graphId,
            smenaId: item.smenaId,
            type: item.type,
            start: item.start,
            end: item.end
        })),
        cellHeight: currentCellHeight
    };
}

function applySnapshot(snapshot) {
    const safeState = snapshot && typeof snapshot === "object" ? snapshot : {};
    currentYear = clamp(parseInt(safeState.year, 10) || DEFAULT_YEAR, 2000, 2100);
    currentCellHeight = clamp(parseInt(safeState.cellHeight, 10) || DEFAULT_CELL_HEIGHT, CELL_HEIGHT_MIN, CELL_HEIGHT_MAX);

    const incomingGraphs = Array.isArray(safeState.graphs) && safeState.graphs.length
        ? safeState.graphs.map(normalizeGraph)
        : [createDefaultGraph()];

    graphs = incomingGraphs;

    const savedActiveId = toId(safeState.activeId);
    activeGraphId = graphs.some(graph => graph.id === savedActiveId) ? savedActiveId : graphs[0].id;

    const knownShiftIds = new Set(graphs.flatMap((graph) => graph.smeny.map((smena) => smena.id)));
    const knownGraphIds = new Set(graphs.map((graph) => graph.id));
    absences = Array.isArray(safeState.absences)
        ? safeState.absences
            .map(normalizeAbsence)
            .filter((item) => item.start <= item.end && knownGraphIds.has(item.graphId) && knownShiftIds.has(item.smenaId))
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

async function loadFromSupabase() {
    const client = ensureSupabaseClient();
    if (!client) {
        return false;
    }

    const config = getSupabaseConfig();
    currentAppKey = config.appKey;

    try {
        const [settingsResponse, graphsResponse, shiftsResponse, absencesResponse] = await Promise.all([
            client.from(SUPABASE_TABLES.settings).select("year, active_graph_id, cell_height").eq("app_key", config.appKey).maybeSingle(),
            client.from(SUPABASE_TABLES.graphs).select("id, name, type, first_shift_date, sort_order").eq("app_key", config.appKey).order("sort_order", { ascending: true }),
            client.from(SUPABASE_TABLES.shifts).select("id, graph_id, name, sort_order").eq("app_key", config.appKey).order("sort_order", { ascending: true }),
            client.from(SUPABASE_TABLES.absences).select("id, graph_id, shift_id, absence_type, start_date, end_date").eq("app_key", config.appKey).order("start_date", { ascending: true })
        ]);

        if (settingsResponse.error) throw settingsResponse.error;
        if (graphsResponse.error) throw graphsResponse.error;
        if (shiftsResponse.error) throw shiftsResponse.error;
        if (absencesResponse.error) throw absencesResponse.error;

        const graphRows = graphsResponse.data || [];
        const shiftRows = shiftsResponse.data || [];
        const absenceRows = absencesResponse.data || [];
        const settingsRow = settingsResponse.data;

        if (!settingsRow && !graphRows.length && !shiftRows.length && !absenceRows.length) {
            hasRemoteStateLoaded = true;
            return false;
        }

        const remoteGraphs = graphRows.map((graphRow, graphIndex) => ({
            id: toId(graphRow.id),
            name: String(graphRow.name || `График ${graphIndex + 1}`),
            type: graphRow.type === "12" ? "12" : "24",
            firstShiftDate: normalizeIsoDate(graphRow.first_shift_date, currentYear),
            smeny: shiftRows
                .filter((shiftRow) => toId(shiftRow.graph_id) === toId(graphRow.id))
                .map((shiftRow, shiftIndex) => ({
                    id: toId(shiftRow.id),
                    name: normalizeShiftName(shiftRow.name, shiftIndex)
                }))
        }));

        const remoteAbsences = absenceRows.map((absenceRow) => ({
            id: toId(absenceRow.id),
            graphId: toId(absenceRow.graph_id),
            smenaId: toId(absenceRow.shift_id),
            type: absenceRow.absence_type === "Больничный" ? "Больничный" : "Отпуск",
            start: normalizeIsoDate(absenceRow.start_date, currentYear),
            end: normalizeIsoDate(absenceRow.end_date, currentYear)
        }));

        applySnapshot({
            year: settingsRow?.year ?? currentYear,
            activeId: settingsRow?.active_graph_id ?? "",
            cellHeight: settingsRow?.cell_height ?? currentCellHeight,
            graphs: remoteGraphs,
            absences: remoteAbsences
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
    const { data, error } = await client.from(tableName).select("id").eq("app_key", appKey);
    if (error) throw error;

    const existingIds = (data || []).map((item) => toId(item.id));
    const idsToDelete = existingIds.filter((id) => !currentIds.includes(id));

    if (!idsToDelete.length) {
        return;
    }

    const { error: deleteError } = await client.from(tableName).delete().eq("app_key", appKey).in("id", idsToDelete);
    if (deleteError) throw deleteError;
}

async function syncStateToSupabase() {
    const client = ensureSupabaseClient();
    if (!client) {
        return;
    }

    const config = getSupabaseConfig();
    currentAppKey = config.appKey;

    const graphRows = graphs.map((graph, graphIndex) => ({
        id: graph.id,
        app_key: config.appKey,
        name: graph.name,
        type: graph.type,
        first_shift_date: graph.firstShiftDate,
        sort_order: graphIndex,
        updated_at: new Date().toISOString()
    }));

    const shiftRows = graphs.flatMap((graph) => graph.smeny.map((smena, shiftIndex) => ({
        id: smena.id,
        app_key: config.appKey,
        graph_id: graph.id,
        name: normalizeShiftName(smena.name, shiftIndex),
        sort_order: shiftIndex,
        updated_at: new Date().toISOString()
    })));

    const absenceRows = absences.map((absence) => ({
        id: absence.id,
        app_key: config.appKey,
        graph_id: absence.graphId,
        shift_id: absence.smenaId,
        absence_type: absence.type,
        start_date: absence.start,
        end_date: absence.end,
        updated_at: new Date().toISOString()
    }));

    try {
        const { error: settingsError } = await client.from(SUPABASE_TABLES.settings).upsert({
            app_key: config.appKey,
            year: currentYear,
            active_graph_id: activeGraphId || null,
            cell_height: currentCellHeight,
            updated_at: new Date().toISOString()
        }, { onConflict: "app_key" });
        if (settingsError) throw settingsError;

        if (graphRows.length) {
            const { error: graphUpsertError } = await client.from(SUPABASE_TABLES.graphs).upsert(graphRows, { onConflict: "id" });
            if (graphUpsertError) throw graphUpsertError;
        }
        await deleteMissingRows(client, SUPABASE_TABLES.graphs, config.appKey, graphRows.map((row) => row.id));

        if (shiftRows.length) {
            const { error: shiftUpsertError } = await client.from(SUPABASE_TABLES.shifts).upsert(shiftRows, { onConflict: "id" });
            if (shiftUpsertError) throw shiftUpsertError;
        }
        await deleteMissingRows(client, SUPABASE_TABLES.shifts, config.appKey, shiftRows.map((row) => row.id));

        if (absenceRows.length) {
            const { error: absenceUpsertError } = await client.from(SUPABASE_TABLES.absences).upsert(absenceRows, { onConflict: "id" });
            if (absenceUpsertError) throw absenceUpsertError;
        }
        await deleteMissingRows(client, SUPABASE_TABLES.absences, config.appKey, absenceRows.map((row) => row.id));
    } catch (error) {
        console.error("Ошибка синхронизации с Supabase", error);
    }
}

function queueRemoteSave() {
    if (!isSupabaseConfigured() || !hasRemoteStateLoaded) {
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
    return graphs.find(graph => graph.id === activeGraphId) || graphs[0] || null;
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

function getCellSchedule(graph, smenaIndex, dateStr) {
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
        const absence = findAbsence(graph.id, smena.id, dateStr);
        const schedule = getCellSchedule(graph, smenaIndex, dateStr);

        let hours = schedule.hours;
        let night = schedule.night;
        let code = "";

        if (absence) {
            hours = 0;
            night = 0;
            code = absence.type === "Больничный" ? "БЛ" : "ОТ";
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
            absence,
            hours,
            night,
            code,
            worked: hours > 0
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

function buildAnnualStats(graph) {
    return graph.smeny.map((smena, smenaIndex) => {
        const stats = { workedDays: 0, hours: 0, night: 0 };
        for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
            const rowData = buildMonthRowData(graph, smena, smenaIndex, monthIndex);
            stats.workedDays += rowData.rowStats.workedDays;
            stats.hours += rowData.rowStats.hours;
            stats.night += rowData.rowStats.night;
        }
        return { smena, ...stats };
    });
}

function renderGraphs() {
    const container = document.getElementById("graphsContainer");
    container.innerHTML = graphs.map((graph) => `
        <div class="graph-card ${graph.id === activeGraphId ? "active" : ""}" onclick="selectGraph('${escapeHtml(graph.id)}')">
            <div class="d-flex justify-content-between gap-2 align-items-start">
                <div>
                    <div class="graph-name">${escapeHtml(graph.name)}</div>
                    <div class="graph-meta">${graph.type === "24" ? "24 часа" : "12 часов"} · ${graph.smeny.length} смен</div>
                </div>
                <button class="btn btn-sm btn-link text-danger p-0" onclick="deleteGraph(event, '${escapeHtml(graph.id)}')" title="Удалить график">
                    <i class="bi bi-x-circle"></i>
                </button>
            </div>
            <div class="mt-2">
                <label class="small text-muted d-block mb-1">Старт цикла</label>
                <input type="date" class="form-control form-control-sm" value="${escapeHtml(graph.firstShiftDate)}" onclick="event.stopPropagation()" onchange="updateGraphDate('${escapeHtml(graph.id)}', this.value)">
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
        <div class="employee-item employee-item-editable">
            <div class="employee-badge">${index + 1}</div>
            <input
                type="text"
                class="form-control form-control-sm"
                value="${escapeHtml(smena.name)}"
                onchange="renameSmena('${escapeHtml(smena.id)}', this.value)"
                onblur="renameSmena('${escapeHtml(smena.id)}', this.value)"
            >
            <button onclick="deleteSmena(event, '${escapeHtml(smena.id)}')" class="btn btn-sm text-danger" title="Удалить смену">
                <i class="bi bi-trash"></i>
            </button>
        </div>
    `).join("");
}

function renderAbsences() {
    const graph = getActiveGraph();
    const container = document.getElementById("absencesList");

    if (!graph) {
        container.innerHTML = "";
        return;
    }

    const list = absences.filter(item => item.graphId === graph.id);

    if (!list.length) {
        container.innerHTML = '<div class="empty-note">Отсутствия не добавлены.</div>';
        return;
    }

    container.innerHTML = list.map((absence) => {
        const smena = graph.smeny.find(item => item.id === absence.smenaId);
        const className = absence.type === "Больничный" ? "absence-bolnichny" : "absence-otpusk";
        const icon = absence.type === "Больничный" ? "🏥" : "🏖️";
        return `
            <div class="employee-item small ${className}">
                <div><b>${icon} ${escapeHtml(absence.type)}</b></div>
                <div>${escapeHtml(smena?.name || "Удаленная смена")}</div>
                <div>${escapeHtml(absence.start)} — ${escapeHtml(absence.end)}</div>
                <button onclick="deleteAbsence(event, '${escapeHtml(absence.id)}')" class="btn btn-sm btn-link text-danger p-0 mt-1">Удалить</button>
            </div>
        `;
    }).join("");
}

function renderActiveGraphInfo(graph) {
    const container = document.getElementById("activeGraphInfo");
    if (!graph) {
        container.innerHTML = "";
        return;
    }

    const patternLabel = graph.type === "24"
        ? "24 часа"
        : "12 часов";

    container.innerHTML = `
        <div class="info-strip">
            <span class="info-strip-name">${escapeHtml(graph.name)}</span>
            <span class="info-strip-meta">${patternLabel}</span>
            <span class="info-strip-meta">Старт: ${escapeHtml(graph.firstShiftDate)}</span>
            <span class="info-strip-meta">Смен: ${graph.smeny.length}</span>
            <span class="info-strip-meta">Год: ${currentYear}${isLeapYear(currentYear) ? " · високосный" : ""}</span>
        </div>
    `;
}

function renderAnnualSummary(graph) {
    const container = document.getElementById("annualSummary");
    if (!graph) {
        container.innerHTML = "";
        return;
    }

    const annualStats = buildAnnualStats(graph);

    container.innerHTML = annualStats.map((item) => `
        <div class="summary-card">
            <div class="summary-card-title">${escapeHtml(item.smena.name)}</div>
            <div class="summary-card-values">
                <div><span class="summary-value">${item.workedDays}</span><span class="summary-label">дней</span></div>
                <div><span class="summary-value">${item.hours}</span><span class="summary-label">часов</span></div>
                <div><span class="summary-value">${item.night}</span><span class="summary-label">ночных</span></div>
            </div>
        </div>
    `).join("");
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
            const rowData = buildMonthRowData(graph, smena, smenaIndex, monthIndex);
            rows.push(rowData);
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
                <div class="cell-lines cell-lines-empty"><span>×</span></div>
            </td>
        `;
    }

    const classes = ["day-cell"];

    if (cell.absence) {
        classes.push("is-absence", cell.absence.type === "Больничный" ? "absence-sick-cell" : "absence-vacation-cell");
    } else if (cell.holiday) {
        classes.push("is-holiday");
    } else if (cell.weekend) {
        classes.push("is-weekend");
    } else if (cell.worked) {
        classes.push("is-worked");
    } else {
        classes.push("is-off");
    }

    const tooltipParts = [cell.dateStr];
    if (cell.worked) tooltipParts.push(`Часы: ${cell.hours}`, `Ночные: ${cell.night}`);
    if (cell.absence) tooltipParts.push(cell.absence.type);
    if (cell.holiday) tooltipParts.push("Праздник");
    else if (cell.weekend) tooltipParts.push("Выходной день календаря");

    const topValue = cell.absence ? cell.code : (cell.hours || "");
    const bottomValue = cell.absence ? "" : (cell.night || "");

    return `
        <td class="${classes.join(" ")}" title="${escapeHtml(tooltipParts.join(" · "))}">
            <div class="cell-lines">
                <div class="cell-top">${topValue}</div>
                <div class="cell-bottom">${bottomValue}</div>
            </div>
        </td>
    `;
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
}

function calculateAndRender() {
    renderAll();
}

function updateYear(value) {
    const parsed = clamp(parseInt(value, 10) || DEFAULT_YEAR, 2000, 2100);
    currentYear = parsed;
    save();
    renderAll();
}

function selectGraph(id) {
    activeGraphId = toId(id);
    save();
    renderAll();
}

function createGraphWithType(type) {
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
    save();
    renderAll();

    nameInput.value = "";
    dateInput.value = `${currentYear}-01-01`;

    bootstrap.Modal.getInstance(document.getElementById("createModal"))?.hide();
}

function deleteGraph(event, id) {
    event.stopPropagation();
    if (graphs.length === 1) {
        alert("Нельзя удалить последний график.");
        return;
    }
    if (!confirm("Удалить этот график?")) {
        return;
    }

    graphs = graphs.filter(graph => graph.id !== toId(id));
    absences = absences.filter(item => item.graphId !== toId(id));
    activeGraphId = graphs[0]?.id || "";
    save();
    renderAll();
}

function updateGraphDate(id, value) {
    const graph = graphs.find(item => item.id === toId(id));
    if (!graph) {
        return;
    }
    graph.firstShiftDate = normalizeIsoDate(value, currentYear);
    save();
    renderAll();
}

function addSmena() {
    const graph = getActiveGraph();
    if (!graph) {
        return;
    }

    graph.smeny.push({
        id: uid("shift"),
        name: `Смена ${graph.smeny.length + 1}`
    });
    save();
    renderAll();
}

function renameSmena(smenaId, value) {
    const graph = getActiveGraph();
    if (!graph) {
        return;
    }
    const smena = graph.smeny.find(item => item.id === toId(smenaId));
    if (!smena) {
        return;
    }
    smena.name = value.trim() || smena.name;
    save();
    renderGraphs();
    renderActiveGraphInfo(graph);
    renderAnnualSummary(graph);
    renderTable(graph);
    renderAbsences();
    renderSmeny();
}

function deleteSmena(event, smenaId) {
    event.stopPropagation();
    const graph = getActiveGraph();
    if (!graph) {
        return;
    }
    if (graph.smeny.length === 1) {
        alert("Нельзя удалить последнюю смену.");
        return;
    }
    if (!confirm("Удалить эту смену?")) {
        return;
    }

    graph.smeny = graph.smeny.filter(item => item.id !== toId(smenaId));
    absences = absences.filter(item => item.smenaId !== toId(smenaId));
    save();
    renderAll();
}

function addAbsenceModal() {
    const graph = getActiveGraph();
    if (!graph || !graph.smeny.length) {
        alert("Сначала добавьте хотя бы одну смену.");
        return;
    }

    document.getElementById("modalSmena").innerHTML = graph.smeny
        .map(item => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name)}</option>`)
        .join("");

    document.getElementById("modalStart").value = "";
    document.getElementById("modalEnd").value = "";
    new bootstrap.Modal(document.getElementById("absenceModal")).show();
}

function saveAbsence() {
    const graph = getActiveGraph();
    if (!graph) {
        return;
    }

    const smenaId = toId(document.getElementById("modalSmena").value);
    const type = document.getElementById("modalType").value === "Больничный" ? "Больничный" : "Отпуск";
    const start = document.getElementById("modalStart").value;
    const end = document.getElementById("modalEnd").value;

    if (!smenaId || !start || !end) {
        alert("Заполните все поля отсутствия.");
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
    if (!confirm("Удалить это отсутствие?")) {
        return;
    }
    absences = absences.filter(item => item.id !== toId(absenceId));
    save();
    renderAll();
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
    if (!graph) {
        return;
    }
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


window.addEventListener("load", async () => {
    normalizeState();
    document.getElementById("newFirstShiftDate").value = `${currentYear}-01-01`;
    renderAll();

    if (isSupabaseConfigured()) {
        const loaded = await loadFromSupabase();
        if (loaded) {
            renderAll();
        } else {
            queueRemoteSave();
        }
    } else {
        hasRemoteStateLoaded = true;
    }

    document.getElementById("createModal").addEventListener("show.bs.modal", () => {
        document.getElementById("newFirstShiftDate").value = `${currentYear}-01-01`;
    });
});
