(function () {
  "use strict";

  const STORAGE = Object.freeze({
    vehicles: "tire.monthly.vehicles.v1",
    drivers: "tire.monthly.drivers.v1"
  });

  const JA_COLLATOR = new Intl.Collator("ja", {
    usage: "sort",
    sensitivity: "base",
    numeric: true,
    ignorePunctuation: true
  });
  const KATAKANA_RE = /[\u30A1-\u30F6]/g;
  const DRIVER_WITH_READING_RE = /^(.*?)\s*[（(]\s*([^（）()]+?)\s*[）)]\s*$/;

  function normalizeText(value) {
    return String(value == null ? "" : value).trim();
  }

  function safeReadJson(key, fallback) {
    try {
      const raw = window.localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  }

  function safeWriteJson(key, value) {
    window.localStorage.setItem(key, JSON.stringify(value));
  }

  function parseDriverEntry(value) {
    const raw = normalizeText(value);
    if (!raw) {
      return { name: "", reading: "" };
    }

    const match = raw.match(DRIVER_WITH_READING_RE);
    if (!match) {
      return { name: raw, reading: "" };
    }

    return {
      name: normalizeText(match[1]) || raw,
      reading: normalizeText(match[2])
    };
  }

  function normalizeDriverName(value) {
    return parseDriverEntry(value).name;
  }

  function toHiragana(value) {
    return normalizeText(value)
      .normalize("NFKC")
      .replace(KATAKANA_RE, function (char) {
        return String.fromCharCode(char.charCodeAt(0) - 0x60);
      })
      .replace(/\s+/g, "");
  }

  function normalizeDriverReading(value) {
    return toHiragana(value);
  }

  function normalizeDriverEntry(value) {
    const parsed = parseDriverEntry(value);
    if (!parsed.name) {
      return "";
    }
    if (!parsed.reading) {
      return parsed.name;
    }
    return parsed.name + "（" + normalizeDriverReading(parsed.reading) + "）";
  }

  function driverSortKey(value) {
    const parsed = parseDriverEntry(value);
    return normalizeDriverReading(parsed.reading || parsed.name);
  }

  function normalizeVehicles(rows) {
    if (!Array.isArray(rows)) {
      return [];
    }

    const unique = [];
    rows.forEach(function (item) {
      const value = normalizeText(item);
      if (!value || unique.includes(value)) {
        return;
      }
      unique.push(value);
    });

    return unique.sort(function (left, right) {
      return JA_COLLATOR.compare(left, right);
    });
  }

  function normalizeDrivers(rows) {
    if (!Array.isArray(rows)) {
      return [];
    }

    const unique = [];
    rows.forEach(function (item) {
      const value = normalizeDriverEntry(item);
      if (!value || unique.includes(value)) {
        return;
      }
      unique.push(value);
    });

    return unique.sort(function (left, right) {
      const keyResult = JA_COLLATOR.compare(driverSortKey(left), driverSortKey(right));
      if (keyResult !== 0) {
        return keyResult;
      }
      return JA_COLLATOR.compare(normalizeDriverName(left), normalizeDriverName(right));
    });
  }

  function readState() {
    return {
      vehicles: normalizeVehicles(safeReadJson(STORAGE.vehicles, [])),
      drivers: normalizeDrivers(safeReadJson(STORAGE.drivers, []))
    };
  }

  function ensureState() {
    const state = readState();
    safeWriteJson(STORAGE.vehicles, state.vehicles);
    safeWriteJson(STORAGE.drivers, state.drivers);
    return state;
  }

  function saveVehicles(rows) {
    const vehicles = normalizeVehicles(rows);
    safeWriteJson(STORAGE.vehicles, vehicles);
    return ensureState();
  }

  function saveDrivers(rows) {
    const drivers = normalizeDrivers(rows);
    safeWriteJson(STORAGE.drivers, drivers);
    return ensureState();
  }

  function addVehicle(value) {
    const next = ensureState().vehicles.concat(normalizeText(value));
    return saveVehicles(next);
  }

  function addDriver(name, reading) {
    const state = ensureState();
    const normalizedName = normalizeDriverName(name);
    if (!normalizedName) {
      return state;
    }

    const nextEntry = normalizeDriverEntry(
      reading ? normalizedName + "（" + normalizeDriverReading(reading) + "）" : normalizedName
    );
    const nextDrivers = state.drivers.slice();
    const existingIndex = nextDrivers.findIndex(function (entry) {
      return normalizeDriverName(entry) === normalizedName;
    });

    if (existingIndex >= 0) {
      nextDrivers.splice(existingIndex, 1, nextEntry);
    } else {
      nextDrivers.push(nextEntry);
    }

    return saveDrivers(nextDrivers);
  }

  window.SharedAppSettings = Object.freeze({
    STORAGE: STORAGE,
    readState: readState,
    ensureState: ensureState,
    saveVehicles: saveVehicles,
    saveDrivers: saveDrivers,
    addVehicle: addVehicle,
    addDriver: addDriver,
    normalizeText: normalizeText,
    normalizeVehicles: normalizeVehicles,
    normalizeDrivers: normalizeDrivers,
    normalizeDriverName: normalizeDriverName,
    normalizeDriverEntry: normalizeDriverEntry,
    normalizeDriverReading: normalizeDriverReading
  });
})();
