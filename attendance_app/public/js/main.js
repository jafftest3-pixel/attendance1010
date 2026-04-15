/* global Papa, XLSX, AttStorage */
(function () {
  const { db, getSetting, setSetting, getOrCreateEmployeeId, getOrCreateDriverId, upsertShift, upsertAttendance, upsertDriverAttendance } =
    AttStorage;

  const SHIFT_ALIASES = {
    employee_name: new Set(["employeename", "employee", "name", "staffname", "fullname"]),
    date: new Set(["date", "workdate", "day", "shiftdate"]),
    shift_start: new Set(["shiftstart", "start", "starttime", "shiftin", "timein"]),
    shift_end: new Set(["shiftend", "end", "endtime", "shiftout", "timeout"]),
  };

  const ATTENDANCE_ALIASES = {
    employee_name: new Set(["employeename", "employee", "name", "staffname", "fullname"]),
    date: new Set(["date", "workdate", "day", "attendancedate"]),
    check_in: new Set(["checkin", "in", "intime", "timein", "clockin"]),
    check_out: new Set(["checkout", "out", "outtime", "timeout", "clockout"]),
  };

  const RIDER_SHIFT_ALIASES = {
    date: new Set(["date", "day", "workdate", "shiftdate"]),
    shift_window: new Set(["shiftwindow", "shift", "shifttime", "window", "column2"]),
    zone_code: new Set(["zonecode", "routecode", "code", "truckcode", "column3"]),
    area_name: new Set(["areaname", "area", "location", "zone", "column4"]),
    driver_name: new Set(["drivername", "ridername", "name", "column5"]),
    assignment_type: new Set(["assignmenttype", "type", "status", "statusyes", "column7"]),
  };

  const DRIVER_ATTENDANCE_ALIASES = {
    driver_name: new Set(["drivername", "driver", "ridername", "name"]),
    date: new Set(["date", "workdate", "day", "attendancedate"]),
    check_in: new Set(["checkin", "in", "intime", "timein", "clockin"]),
    check_out: new Set(["checkout", "out", "outtime", "timeout", "clockout"]),
  };

  function aliasesToObject(setObj) {
    const o = {};
    for (const [k, s] of Object.entries(setObj)) o[k] = s;
    return o;
  }

  function normalizeColumnName(name) {
    return String(name || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");
  }

  function mapColumnsRows(rows, aliases) {
    if (!rows.length) return rows;
    const norm = {};
    Object.keys(rows[0]).forEach((k) => {
      norm[normalizeColumnName(k)] = k;
    });
    const rename = {};
    for (const [target, allowed] of Object.entries(aliases)) {
      for (const candidate of allowed) {
        if (norm[candidate]) {
          rename[norm[candidate]] = target;
          break;
        }
      }
    }
    return rows.map((row) => {
      const o = {};
      for (const [k, v] of Object.entries(row)) {
        const nk = rename[k] || k;
        o[nk] = v;
      }
      return o;
    });
  }

  function requiredColumnsFeedback(rows, aliases) {
    const mapped = mapColumnsRows(rows, aliases);
    const keys = new Set();
    mapped.forEach((r) => Object.keys(r).forEach((k) => keys.add(k)));
    const missing = Object.keys(aliases).filter((k) => !keys.has(k));
    return { mapped, missing, available: [...keys] };
  }

  function formatDateValue(value) {
    if (value == null || value === "") return "";
    if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0, 10);
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    const s = String(value).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return s.slice(0, 10);
    return s;
  }

  function formatTimeValue(value) {
    if (value == null || value === "") return "";
    if (value instanceof Date && !isNaN(value)) {
      const h = value.getHours();
      const m = value.getMinutes();
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }
    if (typeof value === "number") {
      const dayFrac = value % 1;
      if (dayFrac >= 0 && dayFrac <= 1) {
        const totalM = Math.round(dayFrac * 24 * 60);
        const h = Math.floor(totalM / 60) % 24;
        const m = totalM % 60;
        return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
      }
    }
    const s = String(value).trim();
    const ampm = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])$/);
    if (ampm) {
      let h = parseInt(ampm[1], 10);
      const m = parseInt(ampm[2], 10);
      const ap = ampm[4].toUpperCase();
      if (ap === "PM" && h < 12) h += 12;
      if (ap === "AM" && h === 12) h = 0;
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }
    const hm = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (hm) {
      return `${String(parseInt(hm[1], 10)).padStart(2, "0")}:${String(parseInt(hm[2], 10)).padStart(2, "0")}`;
    }
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
    }
    return s;
  }

  function parseTime(value) {
    const t = formatTimeValue(value);
    const [hh, mm] = t.split(":").map((x) => parseInt(x, 10));
    return new Date(2000, 0, 1, hh || 0, mm || 0, 0);
  }

  function calculateHours(checkIn, checkOut) {
    const a = parseTime(checkIn);
    const b = parseTime(checkOut);
    let ms = b - a;
    if (ms < 0) ms += 24 * 3600000;
    return Math.round((ms / 3600000) * 100) / 100;
  }

  function splitShiftWindow(shiftWindow) {
    if (!shiftWindow) return ["-", "-"];
    let raw = String(shiftWindow).trim().replace(/\n/g, " ");
    raw = raw.replace(/[\u064B-\u065F]/g, "");
    const timeMatches = raw.match(/\d{1,2}:\d{2}\s*[AaPp][Mm]/g);
    if (timeMatches && timeMatches.length >= 2) {
      return [formatTimeValue(timeMatches[0]), formatTimeValue(timeMatches[1])];
    }
    const parts = raw.split(/\s*[-–]\s*/);
    if (parts.length !== 2) return ["-", "-"];
    const start = formatTimeValue(parts[0]).trim();
    const end = formatTimeValue(parts[1]).trim();
    if (!start || !end) return ["-", "-"];
    return [start, end];
  }

  function normalizePersonName(name) {
    return String(name || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
  }

  function chooseShiftWindowForDriver(driverName, attendanceDate, riderShiftRows) {
    const driverNorm = normalizePersonName(driverName);
    const dateCandidates = [];
    const fallbackCandidates = [];
    for (const row of riderShiftRows) {
      const shiftNorm = normalizePersonName(row.driverName);
      const same =
        driverNorm === shiftNorm || driverNorm.includes(shiftNorm) || shiftNorm.includes(driverNorm);
      if (!same) continue;
      const rowDate = String(row.date);
      const sw = row.shiftWindow;
      if (rowDate <= attendanceDate) dateCandidates.push([rowDate, sw]);
      fallbackCandidates.push([rowDate, sw]);
    }
    if (dateCandidates.length) {
      dateCandidates.sort((a, b) => b[0].localeCompare(a[0]));
      return dateCandidates[0][1];
    }
    if (fallbackCandidates.length) {
      fallbackCandidates.sort((a, b) => b[0].localeCompare(a[0]));
      return fallbackCandidates[0][1];
    }
    return null;
  }

  function resolveReportRange(period, startDate, endDate) {
    const today = new Date();
    const iso = (d) => d.toISOString().slice(0, 10);
    if (startDate && endDate) return [startDate, endDate];
    if (period === "day") return [iso(today), iso(today)];
    if (period === "week") {
      const start = new Date(today);
      start.setDate(today.getDate() - today.getDay());
      const end = new Date(start);
      end.setDate(start.getDate() + 6);
      return [iso(start), iso(end)];
    }
    if (period === "month") {
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      const next = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      const monthEnd = new Date(next - 86400000);
      return [iso(monthStart), iso(monthEnd)];
    }
    if (startDate) return [startDate, endDate || startDate];
    if (endDate) return [startDate || endDate, endDate];
    return [null, null];
  }

  async function tableFromFile(file) {
    const name = (file.name || "").toLowerCase();
    const buf = await file.arrayBuffer();
    if (name.endsWith(".csv") || name.endsWith(".txt")) {
      const text = new TextDecoder().decode(buf);
      const r = Papa.parse(text, { header: true, skipEmptyLines: true });
      return r.data || [];
    }
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  }

  function googleSheetToCsvUrl(url, sheetName) {
    const cleaned = String(url || "").trim();
    if (sheetName) {
      const m = cleaned.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (m) {
        const id = m[1];
        return `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
      }
    }
    if (cleaned.includes("output=csv") || cleaned.includes("/export?")) return cleaned;
    const m2 = cleaned.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!m2) return cleaned;
    const id = m2[1];
    const gidM = cleaned.match(/gid=([0-9]+)/);
    const gid = gidM ? gidM[1] : "0";
    return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
  }

  async function fetchSheetCsv(targetUrl) {
    const proxy = `/.netlify/functions/sheet-proxy?url=${encodeURIComponent(targetUrl)}`;
    const res = await fetch(proxy);
    if (!res.ok) throw new Error(`Sheet fetch ${res.status}: ${(await res.text()).slice(0, 200)}`);
    return res.text();
  }

  async function parseCsvText(text) {
    const r = Papa.parse(text, { header: true, skipEmptyLines: true });
    return r.data || [];
  }

  async function loadGoogleSheetRows(url, tabCandidates) {
    let lastErr;
    for (const tab of tabCandidates) {
      try {
        const csvUrl = googleSheetToCsvUrl(url, tab);
        const text = await fetchSheetCsv(csvUrl);
        const rows = await parseCsvText(text);
        if (rows.length && Object.keys(rows[0]).length) return rows;
      } catch (e) {
        lastErr = e;
      }
    }
    throw lastErr || new Error("Unable to read Google Sheet tabs");
  }

  function showNotice(el, msg, isError) {
    if (!el) return;
    el.textContent = msg || "";
    el.className = "notice" + (isError ? " notice-error" : "");
    el.style.display = msg ? "block" : "none";
  }

  async function reportRows(graceMinutes, startDate, endDate) {
    const attendance = await db.attendance.toArray();
    const employees = await db.employees.toArray();
    const shifts = await db.shifts.toArray();
    const empMap = Object.fromEntries(employees.map((e) => [e.id, e.name]));
    const graceMs = graceMinutes * 60 * 1000;

    function resolveShift(employeeId, date) {
      const exact = shifts.find((s) => s.employeeId === employeeId && s.date === date);
      if (exact) return { start: exact.shiftStart, end: exact.shiftEnd };
      const before = shifts
        .filter((s) => s.employeeId === employeeId && s.date <= date)
        .sort((a, b) => b.date.localeCompare(a.date));
      if (before.length) return { start: before[0].shiftStart, end: before[0].shiftEnd };
      const any = shifts
        .filter((s) => s.employeeId === employeeId)
        .sort((a, b) => b.date.localeCompare(a.date));
      if (any.length) return { start: any[0].shiftStart, end: any[0].shiftEnd };
      return { start: null, end: null };
    }

    const out = [];
    for (const a of attendance) {
      if (startDate && a.date < startDate) continue;
      if (endDate && a.date > endDate) continue;
      const { start: shiftStart, end: shiftEnd } = resolveShift(a.employeeId, a.date);
      const workedHours = calculateHours(a.checkIn, a.checkOut);
      let status = "No shift assigned";
      let overtimeHours = 0;
      if (shiftStart) {
        const checkInDt = parseTime(a.checkIn);
        const shiftStartDt = parseTime(shiftStart);
        status = checkInDt.getTime() > shiftStartDt.getTime() + graceMs ? "Late" : "On time";
        try {
          const shiftHours = calculateHours(shiftStart, shiftEnd);
          overtimeHours = Math.round(Math.max(0, workedHours - shiftHours) * 100) / 100;
        } catch {
          overtimeHours = 0;
        }
      }
      out.push({
        employee_name: empMap[a.employeeId] || "?",
        date: a.date,
        check_in: a.checkIn,
        check_out: a.checkOut,
        shift_start: shiftStart || "-",
        shift_end: shiftEnd || "-",
        worked_hours: workedHours,
        overtime_hours: overtimeHours,
        status,
      });
    }
    out.sort((a, b) => b.date.localeCompare(a.date) || a.employee_name.localeCompare(b.employee_name));
    return out;
  }

  async function driverReportData(resolvedStart, resolvedEnd) {
    const driverRows = await db.driverAttendance.toArray();
    const drivers = await db.drivers.toArray();
    const riderShifts = await db.riderShifts.toArray();
    const driverMap = Object.fromEntries(drivers.map((d) => [d.id, d.name]));

    const filtered = driverRows.filter((r) => {
      if (resolvedStart && r.date < resolvedStart) return false;
      if (resolvedEnd && r.date > resolvedEnd) return false;
      return true;
    });

    const riderShiftRows = riderShifts.map((x) => ({
      date: x.date,
      driverName: x.driverName,
      shiftWindow: x.shiftWindow,
    }));

    const rows = [];
    const totals = {};
    const overtimeTotals = {};
    let lateCount = 0;
    let onTimeCount = 0;

    for (const row of filtered) {
      const driverName = driverMap[row.driverId] || "?";
      const hours = calculateHours(row.checkIn, row.checkOut);
      const matchedShift = chooseShiftWindowForDriver(driverName, row.date, riderShiftRows);
      const [shiftStart, shiftEnd] = splitShiftWindow(matchedShift || "");
      let status = "No shift assigned";
      let overtimeHours = 0;

      if (shiftStart !== "-" && shiftEnd !== "-") {
        const checkInDt = parseTime(row.checkIn);
        const shiftStartDt = parseTime(shiftStart);
        status = checkInDt.getTime() > shiftStartDt.getTime() + 5 * 60 * 1000 ? "Late" : "On time";
        const shiftHours = calculateHours(shiftStart, shiftEnd);
        overtimeHours = Math.round(Math.max(0, hours - shiftHours) * 100) / 100;
        if (status === "Late") lateCount++;
        else if (status === "On time") onTimeCount++;
      }

      rows.push({
        driver_name: driverName,
        date: row.date,
        check_in: row.checkIn,
        check_out: row.checkOut,
        worked_hours: hours,
        shift_start: shiftStart,
        shift_end: shiftEnd,
        overtime_hours: overtimeHours,
        status,
      });
      totals[driverName] = Math.round((totals[driverName] || 0) + hours * 100) / 100;
      overtimeTotals[driverName] = Math.round(((overtimeTotals[driverName] || 0) + overtimeHours) * 100) / 100;
    }

    rows.sort((a, b) => b.date.localeCompare(a.date) || a.driver_name.localeCompare(b.driver_name));

    const summary = Object.keys(totals)
      .sort()
      .map((k) => ({
        driver_name: k,
        total_worked_hours: totals[k],
        total_overtime_hours: overtimeTotals[k] || 0,
      }));

    const driverCount = new Set(filtered.map((r) => r.driverId)).size;

    return { rows, summary, stats: { driverCount, attendanceCount: rows.length, lateCount, onTimeCount } };
  }

  function setNavActive(hash) {
    const hNorm = hash === "" || hash === "#" ? "#shifts" : hash;
    document.querySelectorAll(".sidebar-nav .nav-item").forEach((a) => {
      const h = a.getAttribute("href");
      a.classList.toggle("active", h === hNorm);
    });
  }

  function showOnlyView(hash) {
    const id = (hash.replace("#", "") || "shifts").toLowerCase();
    document.querySelectorAll("[data-view]").forEach((el) => {
      el.style.display = el.getAttribute("data-view") === id ? "" : "none";
    });
    setNavActive(!hash || hash === "#" ? "#shifts" : hash);
  }

  async function refreshEmployeesSelect(selectEl) {
    const list = await db.employees.orderBy("name").toArray();
    const v = selectEl.value;
    selectEl.innerHTML = '<option value="">Select employee</option>';
    list.forEach((e) => {
      const o = document.createElement("option");
      o.value = e.id;
      o.textContent = e.name;
      selectEl.appendChild(o);
    });
    selectEl.value = v;
  }

  async function refreshDriversSelect(selectEl) {
    const list = await db.drivers.orderBy("name").toArray();
    const v = selectEl.value;
    selectEl.innerHTML = '<option value="">Select driver</option>';
    list.forEach((e) => {
      const o = document.createElement("option");
      o.value = e.id;
      o.textContent = e.name;
      selectEl.appendChild(o);
    });
    selectEl.value = v;
  }

  async function renderShiftsView() {
    const tbody = document.querySelector("#shifts-table tbody");
    const rows = await db.shifts.toArray();
    const emps = await db.employees.toArray();
    const idToName = Object.fromEntries(emps.map((e) => [e.id, e.name]));
    const joined = rows
      .map((s) => ({
        employee_name: idToName[s.employeeId] || "?",
        date: s.date,
        shift_start: s.shiftStart,
        shift_end: s.shiftEnd,
      }))
      .sort((a, b) => b.date.localeCompare(a.date) || a.employee_name.localeCompare(b.employee_name));
    tbody.innerHTML =
      joined.length === 0
        ? "<tr><td colspan='4'>No shifts yet.</td></tr>"
        : joined
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.employee_name)}</td><td>${r.date}</td><td>${r.shift_start}</td><td>${r.shift_end}</td></tr>`
            )
            .join("");

    document.getElementById("emp-shift-url").textContent = (await getSetting("linked_employee_sheet_url", "")) || "Not set";
    document.getElementById("emp-shift-tab").textContent = await getSetting("linked_employee_sheet_tab", "Employee Shift");
    document.getElementById("emp-shift-sync").textContent = (await getSetting("linked_employee_last_sync", "")) || "Never";
    const st = await getSetting("linked_employee_sync_status", "idle");
    const badge = document.getElementById("emp-shift-status");
    badge.textContent = st === "ok" ? "Connected" : st === "error" ? "Sync Error" : "Not Synced";
    badge.className = "status-badge status-" + st;

    const urlInput = document.getElementById("linked_employee_sheet_url");
    const tabInput = document.getElementById("linked_employee_sheet_tab");
    if (urlInput) urlInput.value = await getSetting("linked_employee_sheet_url", "");
    if (tabInput) tabInput.value = await getSetting("linked_employee_sheet_tab", "Employee Shift");

    const shiftEmployeeSelect = document.getElementById("shift-employee");
    if (shiftEmployeeSelect) await refreshEmployeesSelect(shiftEmployeeSelect);
  }

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }

  async function renderAttendanceView() {
    const tbody = document.querySelector("#attendance-table tbody");
    const rows = await db.attendance.toArray();
    const emps = await db.employees.toArray();
    const idToName = Object.fromEntries(emps.map((e) => [e.id, e.name]));
    const joined = rows
      .map((a) => ({
        employee_name: idToName[a.employeeId] || "?",
        date: a.date,
        check_in: a.checkIn,
        check_out: a.checkOut,
      }))
      .sort((a, b) => b.date.localeCompare(a.date) || a.employee_name.localeCompare(b.employee_name));
    tbody.innerHTML =
      joined.length === 0
        ? "<tr><td colspan='4'>No attendance yet.</td></tr>"
        : joined
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.employee_name)}</td><td>${r.date}</td><td>${r.check_in}</td><td>${r.check_out}</td></tr>`
            )
            .join("");

    const st = await getSetting("linked_employee_attendance_sync_status", "idle");
    const badge = document.getElementById("emp-att-status");
    badge.textContent = st === "ok" ? "Connected" : st === "error" ? "Sync Error" : "Not Synced";
    badge.className = "status-badge status-" + st;
    document.getElementById("emp-att-sync").textContent = (await getSetting("linked_employee_attendance_last_sync", "")) || "Never";
    document.getElementById("emp-att-tab").textContent = await getSetting(
      "linked_employee_attendance_sheet_tab",
      "eMPLOYEE aTEENDANCE"
    );
    document.getElementById("linked_emp_att_url").value = await getSetting(
      "linked_employee_attendance_sheet_url",
      (await getSetting("linked_employee_sheet_url", "")) || (await getSetting("linked_rider_sheet_url", ""))
    );
    document.getElementById("linked_emp_att_tab").value = await getSetting(
      "linked_employee_attendance_sheet_tab",
      "eMPLOYEE aTEENDANCE"
    );
    await refreshEmployeesSelect(document.getElementById("att-employee"));
  }

  async function renderRidersView() {
    const tbody = document.querySelector("#riders-table tbody");
    const rows = await db.riderShifts.orderBy("date").reverse().toArray();
    tbody.innerHTML =
      rows.length === 0
        ? "<tr><td colspan='6'>No rider shifts yet.</td></tr>"
        : rows
            .map(
              (r) =>
                `<tr><td>${r.date}</td><td>${escapeHtml(r.shiftWindow)}</td><td>${escapeHtml(r.zoneCode)}</td><td>${escapeHtml(
                  r.areaName
                )}</td><td>${escapeHtml(r.driverName)}</td><td>${escapeHtml(r.assignmentType)}</td></tr>`
            )
            .join("");
    const st = await getSetting("linked_rider_sync_status", "idle");
    const badge = document.getElementById("rider-status");
    badge.textContent = st === "ok" ? "Connected" : st === "error" ? "Sync Error" : "Not Synced";
    badge.className = "status-badge status-" + st;
    document.getElementById("rider-sync-time").textContent = (await getSetting("linked_rider_last_sync", "")) || "Never";
    document.getElementById("linked_rider_url").value = await getSetting("linked_rider_sheet_url", "");
  }

  async function renderDriversAttView() {
    const tbody = document.querySelector("#drivers-att-table tbody");
    const rows = await db.driverAttendance.toArray();
    const drivers = await db.drivers.toArray();
    const idToName = Object.fromEntries(drivers.map((d) => [d.id, d.name]));
    const joined = rows
      .map((r) => ({
        driver_name: idToName[r.driverId] || "?",
        date: r.date,
        check_in: r.checkIn,
        check_out: r.checkOut,
      }))
      .sort((a, b) => b.date.localeCompare(a.date) || a.driver_name.localeCompare(b.driver_name));
    tbody.innerHTML =
      joined.length === 0
        ? "<tr><td colspan='4'>No driver attendance yet.</td></tr>"
        : joined
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.driver_name)}</td><td>${r.date}</td><td>${r.check_in}</td><td>${r.check_out}</td></tr>`
            )
            .join("");
    const st = await getSetting("linked_driver_attendance_sync_status", "idle");
    const badge = document.getElementById("drv-att-status");
    badge.textContent = st === "ok" ? "Connected" : st === "error" ? "Sync Error" : "Not Synced";
    badge.className = "status-badge status-" + st;
    document.getElementById("drv-att-sync").textContent = (await getSetting("linked_driver_attendance_last_sync", "")) || "Never";
    document.getElementById("drv-att-tab").textContent = await getSetting(
      "linked_driver_attendance_sheet_tab",
      "Driver aTTENDANCE"
    );
    document.getElementById("linked_drv_att_url").value = await getSetting(
      "linked_driver_attendance_sheet_url",
      await getSetting("linked_rider_sheet_url", "")
    );
    document.getElementById("linked_drv_att_tab").value = await getSetting(
      "linked_driver_attendance_sheet_tab",
      "Driver aTTENDANCE"
    );
    await refreshDriversSelect(document.getElementById("drv-employee"));
  }

  function readReportFilters(prefix) {
    const period = document.getElementById(prefix + "-period").value;
    const start_date = document.getElementById(prefix + "-start").value;
    const end_date = document.getElementById(prefix + "-end").value;
    return { period, start_date, end_date };
  }

  async function renderEmployeeReports() {
    const { period, start_date, end_date } = readReportFilters("er");
    const [rs, re] = resolveReportRange(period, start_date, end_date);
    const rows = await reportRows(5, rs, re);
    const totals = {};
    let late = 0;
    let onTime = 0;
    rows.forEach((r) => {
      totals[r.employee_name] = Math.round(((totals[r.employee_name] || 0) + r.worked_hours) * 100) / 100;
      if (r.status === "Late") late++;
      if (r.status === "On time") onTime++;
    });
    const summary = Object.keys(totals)
      .sort()
      .map((k) => ({ employee_name: k, total_worked_hours: totals[k] }));

    const empCount = await db.employees.count();

    document.getElementById("er-stat-emp").textContent = empCount;
    document.getElementById("er-stat-att").textContent = rows.length;
    document.getElementById("er-stat-late").textContent = late;
    document.getElementById("er-stat-ontime").textContent = onTime;

    document.querySelector("#er-summary tbody").innerHTML =
      summary.length === 0
        ? "<tr><td colspan='2'>No report data yet.</td></tr>"
        : summary.map((r) => `<tr><td>${escapeHtml(r.employee_name)}</td><td>${r.total_worked_hours}</td></tr>`).join("");

    document.querySelector("#er-detail tbody").innerHTML =
      rows.length === 0
        ? "<tr><td colspan='9'>No report data yet.</td></tr>"
        : rows
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.employee_name)}</td><td>${r.date}</td><td>${r.check_in}</td><td>${r.check_out}</td><td>${r.shift_start}</td><td>${r.shift_end}</td><td>${r.worked_hours}</td><td>${r.overtime_hours}</td><td>${r.status}</td></tr>`
            )
            .join("");
  }

  async function renderDriverReports() {
    const { period, start_date, end_date } = readReportFilters("dr");
    const [rs, re] = resolveReportRange(period, start_date, end_date);
    const { rows, summary, stats } = await driverReportData(rs, re);

    document.getElementById("dr-stat-drv").textContent = stats.driverCount;
    document.getElementById("dr-stat-att").textContent = stats.attendanceCount;
    document.getElementById("dr-stat-late").textContent = stats.lateCount;
    document.getElementById("dr-stat-ontime").textContent = stats.onTimeCount;

    document.querySelector("#dr-summary tbody").innerHTML =
      summary.length === 0
        ? "<tr><td colspan='3'>No driver report data yet.</td></tr>"
        : summary
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.driver_name)}</td><td>${r.total_worked_hours}</td><td>${r.total_overtime_hours}</td></tr>`
            )
            .join("");

    document.querySelector("#dr-detail tbody").innerHTML =
      rows.length === 0
        ? "<tr><td colspan='9'>No driver report data yet.</td></tr>"
        : rows
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.driver_name)}</td><td>${r.date}</td><td>${r.check_in}</td><td>${r.check_out}</td><td>${r.shift_start}</td><td>${r.shift_end}</td><td>${r.worked_hours}</td><td>${r.overtime_hours}</td><td>${r.status}</td></tr>`
            )
            .join("");
  }

  async function exportEmployeeExcel() {
    const { period, start_date, end_date } = readReportFilters("er");
    const [rs, re] = resolveReportRange(period, start_date, end_date);
    const rows = await reportRows(5, rs, re);
    const summary = {};
    rows.forEach((r) => {
      summary[r.employee_name] = (summary[r.employee_name] || 0) + r.worked_hours;
    });
    const summaryRows = Object.keys(summary)
      .sort()
      .map((k) => ({ employee_name: k, total_worked_hours: Math.round(summary[k] * 100) / 100 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Detailed");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Summary");
    XLSX.writeFile(wb, "attendance_report.xlsx");
  }

  async function exportDriverExcel() {
    const { period, start_date, end_date } = readReportFilters("dr");
    const [rs, re] = resolveReportRange(period, start_date, end_date);
    const { rows, summary } = await driverReportData(rs, re);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Detailed");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Summary");
    XLSX.writeFile(wb, "driver_attendance_report.xlsx");
  }

  async function route() {
    let hash = location.hash || "#shifts";
    if (hash === "#home" || hash === "#") hash = "#shifts";
    showOnlyView(hash);
    if (hash === "#shifts") await renderShiftsView();
    else if (hash === "#attendance") await renderAttendanceView();
    else if (hash === "#riders-shifts") await renderRidersView();
    else if (hash === "#drivers-attendance") await renderDriversAttView();
    else if (hash === "#reports") await renderEmployeeReports();
    else if (hash === "#driver-reports") await renderDriverReports();
  }

  async function importShiftsFromRows(mappedRows, replaceAll) {
    let n = 0;
    if (replaceAll) await db.shifts.clear();
    const defaultDate = "1900-01-01";
    for (const row of mappedRows) {
      const employeeName = String(row.employee_name || "").trim();
      if (!employeeName) continue;
      const rawDate = formatDateValue(row.date);
      const date = rawDate || (replaceAll ? defaultDate : "");
      const shiftStart = formatTimeValue(row.shift_start);
      const shiftEnd = formatTimeValue(row.shift_end);
      if (!date || !shiftStart || !shiftEnd) continue;
      const eid = await getOrCreateEmployeeId(employeeName);
      await upsertShift(eid, date, shiftStart, shiftEnd);
      n++;
    }
    return n;
  }

  async function importAttendanceFromRows(mappedRows, replaceAll) {
    let n = 0;
    if (replaceAll) await db.attendance.clear();
    for (const row of mappedRows) {
      const employeeName = String(row.employee_name || "").trim();
      if (!employeeName) continue;
      const date = formatDateValue(row.date);
      const checkIn = formatTimeValue(row.check_in);
      const checkOut = formatTimeValue(row.check_out);
      if (!date || !checkIn || !checkOut) continue;
      const eid = await getOrCreateEmployeeId(employeeName);
      await upsertAttendance(eid, date, checkIn, checkOut);
      n++;
    }
    return n;
  }

  async function importRiderShiftsFromRows(mappedRows, appendOnly) {
    let n = 0;
    if (!appendOnly) await db.riderShifts.clear();
    for (const row of mappedRows) {
      const date = formatDateValue(row.date);
      const shiftWindow = String(row.shift_window || "").trim();
      const zoneCode = String(row.zone_code || "").trim();
      const areaName = String(row.area_name || "").trim();
      const driverName = String(row.driver_name || "").trim();
      const assignmentType = String(row.assignment_type || "").trim();
      if (!date || !shiftWindow || !zoneCode || !areaName || !driverName || !assignmentType) continue;
      await db.riderShifts.add({ date, shiftWindow, zoneCode, areaName, driverName, assignmentType });
      await getOrCreateDriverId(driverName);
      n++;
    }
    return n;
  }

  async function importDriverAttendanceFromRows(mappedRows, replaceAll) {
    let n = 0;
    if (replaceAll) await db.driverAttendance.clear();
    for (const row of mappedRows) {
      const driverName = String(row.driver_name || "").trim();
      if (!driverName) continue;
      const date = formatDateValue(row.date);
      const checkIn = formatTimeValue(row.check_in);
      const checkOut = formatTimeValue(row.check_out);
      if (!date || !checkIn || !checkOut) continue;
      const did = await getOrCreateDriverId(driverName);
      await upsertDriverAttendance(did, date, checkIn, checkOut);
      n++;
    }
    return n;
  }

  function bindForms() {
    const formAddEmployee = document.getElementById("form-add-employee");
    if (formAddEmployee) {
      formAddEmployee.addEventListener("submit", async (e) => {
        e.preventDefault();
        const name = document.getElementById("new-employee-name").value.trim();
        const notice = document.getElementById("shift-notice");
        if (!name) return showNotice(notice, "Enter a name", true);
        await getOrCreateEmployeeId(name);
        document.getElementById("new-employee-name").value = "";
        showNotice(notice, "Employee added");
        await refreshEmployeesSelect(document.getElementById("shift-employee"));
      });
    }

    const formShift = document.getElementById("form-shift");
    if (formShift) {
      formShift.addEventListener("submit", async (e) => {
        e.preventDefault();
        const notice = document.getElementById("shift-notice");
        const employeeId = parseInt(document.getElementById("shift-employee").value, 10);
        const date = document.getElementById("shift-date").value;
        const shiftStart = document.getElementById("shift-start").value;
        const shiftEnd = document.getElementById("shift-end").value;
        if (!employeeId || !date || !shiftStart || !shiftEnd) return showNotice(notice, "Fill all fields", true);
        const [sh, sm] = shiftStart.split(":");
        const [eh, em] = shiftEnd.split(":");
        await upsertShift(employeeId, date, `${sh}:${sm}`, `${eh}:${em}`);
        showNotice(notice, "Shift saved");
        await renderShiftsView();
      });
    }

    document.getElementById("form-shift-sheet").addEventListener("submit", async (e) => {
      e.preventDefault();
      const url = document.getElementById("linked_employee_sheet_url").value.trim();
      const tab = document.getElementById("linked_employee_sheet_tab").value.trim() || "Employee Shift";
      const notice = document.getElementById("shift-notice");
      if (!url) return showNotice(notice, "Enter URL", true);
      await setSetting("linked_employee_sheet_url", url);
      await setSetting("linked_employee_sheet_tab", tab);
      await setSetting("linked_employee_sync_status", "idle");
      showNotice(notice, "Google Sheet link saved");
      await renderShiftsView();
    });

    document.getElementById("form-shift-upload").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("shift-notice");
      const f = document.getElementById("shift-file").files[0];
      if (!f) return showNotice(notice, "Choose a file", true);
      try {
        const rows = await tableFromFile(f);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(SHIFT_ALIASES));
        if (missing.length) {
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importShiftsFromRows(mapped, false);
        showNotice(notice, `Imported ${n} shift rows`);
        await renderShiftsView();
      } catch (err) {
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-shift-sync").addEventListener("click", async () => {
      const notice = document.getElementById("shift-notice");
      const url = await getSetting("linked_employee_sheet_url", "");
      const tab = await getSetting("linked_employee_sheet_tab", "Employee Shift");
      if (!url) return showNotice(notice, "Save sheet URL first", true);
      try {
        const csvUrl = googleSheetToCsvUrl(url, tab);
        const text = await fetchSheetCsv(csvUrl);
        const rows = await parseCsvText(text);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(SHIFT_ALIASES));
        if (missing.length) {
          await setSetting("linked_employee_sync_status", "error");
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importShiftsFromRows(mapped, true);
        await setSetting("linked_employee_last_sync", new Date().toISOString().slice(0, 19).replace("T", " "));
        await setSetting("linked_employee_sync_status", "ok");
        showNotice(notice, `Employee sheet synced: ${n} shift rows`);
        await renderShiftsView();
      } catch (err) {
        await setSetting("linked_employee_sync_status", "error");
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-shift-clear").addEventListener("click", async () => {
      if (!confirm("Clear all shifts data?")) return;
      await db.shifts.clear();
      document.getElementById("shift-notice").style.display = "none";
      await renderShiftsView();
    });

    document.getElementById("form-att").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("att-notice");
      const employeeId = parseInt(document.getElementById("att-employee").value, 10);
      const date = document.getElementById("att-date").value;
      const checkIn = document.getElementById("att-in").value;
      const checkOut = document.getElementById("att-out").value;
      if (!employeeId || !date || !checkIn || !checkOut) return showNotice(notice, "Fill all fields", true);
      const [ih, im] = checkIn.split(":");
      const [oh, om] = checkOut.split(":");
      await upsertAttendance(employeeId, date, `${ih}:${im}`, `${oh}:${om}`);
      showNotice(notice, "Attendance saved");
      await renderAttendanceView();
    });

    document.getElementById("form-att-sheet").addEventListener("submit", async (e) => {
      e.preventDefault();
      const url = document.getElementById("linked_emp_att_url").value.trim();
      const tab = document.getElementById("linked_emp_att_tab").value.trim() || "eMPLOYEE aTEENDANCE";
      const notice = document.getElementById("att-notice");
      if (!url) return showNotice(notice, "Enter URL", true);
      await setSetting("linked_employee_attendance_sheet_url", url);
      await setSetting("linked_employee_attendance_sheet_tab", tab);
      await setSetting("linked_employee_attendance_sync_status", "idle");
      showNotice(notice, "Google Sheet link saved");
    });

    document.getElementById("btn-att-sync").addEventListener("click", async () => {
      const notice = document.getElementById("att-notice");
      let url = await getSetting("linked_employee_attendance_sheet_url", "");
      if (!url) url = await getSetting("linked_employee_sheet_url", "");
      if (!url) url = await getSetting("linked_rider_sheet_url", "");
      const tab = await getSetting("linked_employee_attendance_sheet_tab", "Employee attendance");
      if (!url) return showNotice(notice, "Save sheet URL first", true);
      try {
        const rows = await loadGoogleSheetRows(url, [tab, "Employee attendance", "Employee Attendance", "eMPLOYEE aTEENDANCE"]);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(ATTENDANCE_ALIASES));
        if (missing.length) {
          await setSetting("linked_employee_attendance_sync_status", "error");
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importAttendanceFromRows(mapped, true);
        await setSetting("linked_employee_attendance_last_sync", new Date().toISOString().slice(0, 19).replace("T", " "));
        await setSetting("linked_employee_attendance_sync_status", "ok");
        showNotice(notice, `Employee attendance synced: ${n} rows`);
        await renderAttendanceView();
      } catch (err) {
        await setSetting("linked_employee_attendance_sync_status", "error");
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("form-att-upload").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("att-notice");
      const f = document.getElementById("att-file").files[0];
      if (!f) return showNotice(notice, "Choose a file", true);
      try {
        const rows = await tableFromFile(f);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(ATTENDANCE_ALIASES));
        if (missing.length) {
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importAttendanceFromRows(mapped, false);
        showNotice(notice, `Imported ${n} attendance rows`);
        await renderAttendanceView();
      } catch (err) {
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-att-clear").addEventListener("click", async () => {
      if (!confirm("Clear all attendance?")) return;
      await db.attendance.clear();
      await renderAttendanceView();
    });

    document.getElementById("form-rider").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("rider-notice");
      const date = document.getElementById("rider-date").value;
      const shiftWindow = document.getElementById("rider-window").value.trim();
      const zoneCode = document.getElementById("rider-zone").value.trim();
      const areaName = document.getElementById("rider-area").value.trim();
      const driverName = document.getElementById("rider-driver").value.trim();
      const assignmentType = document.getElementById("rider-type").value.trim();
      if (!date || !shiftWindow || !zoneCode || !areaName || !driverName || !assignmentType) {
        return showNotice(notice, "Fill all fields", true);
      }
      await db.riderShifts.add({ date, shiftWindow, zoneCode, areaName, driverName, assignmentType });
      await getOrCreateDriverId(driverName);
      showNotice(notice, "Rider shift saved");
      await renderRidersView();
    });

    document.getElementById("form-rider-sheet").addEventListener("submit", async (e) => {
      e.preventDefault();
      const url = document.getElementById("linked_rider_url").value.trim();
      const notice = document.getElementById("rider-notice");
      if (!url) return showNotice(notice, "Enter URL", true);
      await setSetting("linked_rider_sheet_url", url);
      await setSetting("linked_rider_sync_status", "idle");
      showNotice(notice, "Google Sheet link saved");
      await renderRidersView();
    });

    document.getElementById("form-rider-upload").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("rider-notice");
      const f = document.getElementById("rider-file").files[0];
      if (!f) return showNotice(notice, "Choose a file", true);
      try {
        const rows = await tableFromFile(f);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(RIDER_SHIFT_ALIASES));
        if (missing.length) {
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importRiderShiftsFromRows(mapped, true);
        showNotice(notice, `Imported ${n} rider shift rows`);
        await renderRidersView();
      } catch (err) {
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-rider-sync").addEventListener("click", async () => {
      const notice = document.getElementById("rider-notice");
      const url = await getSetting("linked_rider_sheet_url", "");
      if (!url) return showNotice(notice, "Save sheet URL first", true);
      try {
        const csvUrl = googleSheetToCsvUrl(url);
        const text = await fetchSheetCsv(csvUrl);
        const rows = await parseCsvText(text);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(RIDER_SHIFT_ALIASES));
        if (missing.length) {
          await setSetting("linked_rider_sync_status", "error");
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importRiderShiftsFromRows(mapped, false);
        await setSetting("linked_rider_last_sync", new Date().toISOString().slice(0, 19).replace("T", " "));
        await setSetting("linked_rider_sync_status", "ok");
        showNotice(notice, `Google Sheet synced: ${n} rider shift rows`);
        await renderRidersView();
      } catch (err) {
        await setSetting("linked_rider_sync_status", "error");
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-rider-clear").addEventListener("click", async () => {
      if (!confirm("Clear all rider shifts?")) return;
      await db.riderShifts.clear();
      await renderRidersView();
    });

    document.getElementById("form-drv-att").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("drv-notice");
      const driverId = parseInt(document.getElementById("drv-employee").value, 10);
      const date = document.getElementById("drv-date").value;
      const checkIn = document.getElementById("drv-in").value;
      const checkOut = document.getElementById("drv-out").value;
      if (!driverId || !date || !checkIn || !checkOut) return showNotice(notice, "Fill all fields", true);
      const [ih, im] = checkIn.split(":");
      const [oh, om] = checkOut.split(":");
      await upsertDriverAttendance(driverId, date, `${ih}:${im}`, `${oh}:${om}`);
      showNotice(notice, "Driver attendance saved");
      await renderDriversAttView();
    });

    document.getElementById("form-drv-sheet").addEventListener("submit", async (e) => {
      e.preventDefault();
      const url = document.getElementById("linked_drv_att_url").value.trim();
      const tab = document.getElementById("linked_drv_att_tab").value.trim() || "Driver aTTENDANCE";
      const notice = document.getElementById("drv-notice");
      if (!url) return showNotice(notice, "Enter URL", true);
      await setSetting("linked_driver_attendance_sheet_url", url);
      await setSetting("linked_driver_attendance_sheet_tab", tab);
      await setSetting("linked_driver_attendance_sync_status", "idle");
      showNotice(notice, "Google Sheet link saved");
    });

    document.getElementById("btn-drv-sync").addEventListener("click", async () => {
      const notice = document.getElementById("drv-notice");
      let url = await getSetting("linked_driver_attendance_sheet_url", "");
      if (!url) url = await getSetting("linked_rider_sheet_url", "");
      const tab = await getSetting("linked_driver_attendance_sheet_tab", "Driver Ateendance");
      if (!url) return showNotice(notice, "Save sheet URL first", true);
      try {
        const rows = await loadGoogleSheetRows(url, [tab, "Driver Ateendance", "Driver Attendance", "Driver aTTENDANCE"]);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(DRIVER_ATTENDANCE_ALIASES));
        if (missing.length) {
          await setSetting("linked_driver_attendance_sync_status", "error");
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importDriverAttendanceFromRows(mapped, true);
        await setSetting("linked_driver_attendance_last_sync", new Date().toISOString().slice(0, 19).replace("T", " "));
        await setSetting("linked_driver_attendance_sync_status", "ok");
        showNotice(notice, `Driver attendance synced: ${n} rows`);
        await renderDriversAttView();
      } catch (err) {
        await setSetting("linked_driver_attendance_sync_status", "error");
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("form-drv-upload").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notice = document.getElementById("drv-notice");
      const f = document.getElementById("drv-file").files[0];
      if (!f) return showNotice(notice, "Choose a file", true);
      try {
        const rows = await tableFromFile(f);
        const { mapped, missing, available } = requiredColumnsFeedback(rows, aliasesToObject(DRIVER_ATTENDANCE_ALIASES));
        if (missing.length) {
          showNotice(notice, `Missing columns: ${missing.join(", ")} | Found: ${available.join(", ")}`, true);
          return;
        }
        const n = await importDriverAttendanceFromRows(mapped, false);
        showNotice(notice, `Imported ${n} driver attendance rows`);
        await renderDriversAttView();
      } catch (err) {
        showNotice(notice, String(err.message || err), true);
      }
    });

    document.getElementById("btn-drv-clear").addEventListener("click", async () => {
      if (!confirm("Clear all driver attendance?")) return;
      await db.driverAttendance.clear();
      await renderDriversAttView();
    });

    document.getElementById("form-er-filter").addEventListener("submit", async (e) => {
      e.preventDefault();
      await renderEmployeeReports();
    });
    document.getElementById("btn-er-export").addEventListener("click", () => exportEmployeeExcel());

    document.getElementById("form-dr-filter").addEventListener("submit", async (e) => {
      e.preventDefault();
      await renderDriverReports();
    });
    document.getElementById("btn-dr-export").addEventListener("click", () => exportDriverExcel());
  }

  async function init() {
    await db.open();
    bindForms();
    window.addEventListener("hashchange", route);
    if (!location.hash || location.hash === "#home" || location.hash === "#") location.hash = "#shifts";
    await route();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
