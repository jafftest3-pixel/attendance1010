/* global Dexie */
(function (global) {
  const db = new Dexie("AttendanceNetlify");

  db.version(1).stores({
    employees: "++id, &name",
    shifts: "++id, [employeeId+date], employeeId, date",
    attendance: "++id, [employeeId+date], employeeId, date",
    drivers: "++id, &name",
    riderShifts: "++id, date, driverName",
    driverAttendance: "++id, [driverId+date], driverId, date",
    settings: "key",
  });

  const apiRoot = "/api/settings";

  async function fetchRemoteSetting(key) {
    try {
      const response = await fetch(`${apiRoot}/${encodeURIComponent(key)}`);
      if (!response.ok) throw new Error("Remote settings unavailable");
      const data = await response.json();
      if (data && Object.prototype.hasOwnProperty.call(data, "value")) {
        return data.value;
      }
    } catch (error) {
      return null;
    }
    return null;
  }

  async function saveRemoteSetting(key, value) {
    try {
      const response = await fetch(`${apiRoot}/${encodeURIComponent(key)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ value: String(value) }),
      });
      return response.ok;
    } catch (error) {
      return false;
    }
  }

  async function getSetting(key, defaultValue) {
    const remoteValue = await fetchRemoteSetting(key);
    if (remoteValue !== null) {
      return remoteValue;
    }
    const row = await db.settings.get(key);
    return row ? row.value : defaultValue;
  }

  async function setSetting(key, value) {
    await db.settings.put({ key, value: String(value) });
    await saveRemoteSetting(key, value);
  }

  async function getOrCreateEmployeeId(name) {
    const n = String(name || "").trim();
    if (!n) throw new Error("Employee name required");
    let id = await db.employees.where("name").equals(n).first();
    if (id) return id.id;
    return db.employees.add({ name: n });
  }

  async function getOrCreateDriverId(name) {
    const n = String(name || "").trim();
    if (!n) throw new Error("Driver name required");
    let id = await db.drivers.where("name").equals(n).first();
    if (id) return id.id;
    return db.drivers.add({ name: n });
  }

  async function upsertShift(employeeId, date, shiftStart, shiftEnd) {
    await db.shifts.where("[employeeId+date]").equals([employeeId, date]).delete();
    await db.shifts.add({
      employeeId,
      date,
      shiftStart,
      shiftEnd,
    });
  }

  async function upsertAttendance(employeeId, date, checkIn, checkOut) {
    await db.attendance.where("[employeeId+date]").equals([employeeId, date]).delete();
    await db.attendance.add({ employeeId, date, checkIn, checkOut });
  }

  async function upsertDriverAttendance(driverId, date, checkIn, checkOut) {
    await db.driverAttendance.where("[driverId+date]").equals([driverId, date]).delete();
    await db.driverAttendance.add({ driverId, date, checkIn, checkOut });
  }

  global.AttStorage = {
    db,
    getSetting,
    setSetting,
    getOrCreateEmployeeId,
    getOrCreateDriverId,
    upsertShift,
    upsertAttendance,
    upsertDriverAttendance,
  };
})(typeof window !== "undefined" ? window : globalThis);
