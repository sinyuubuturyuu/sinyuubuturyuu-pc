(function () {
  "use strict";

  const DEFAULT_COLLECTION = "monthly_tire_autosave";
  const TIRE_FIELDS = ["maker", "type", "groove", "wear", "damage", "pressure"];
  const KNOWN_TRUCK_TYPES = ["low12", "ten10"];
  const TIRE_OPTION_ORDER = {
    maker: ["ミシュラン", "サイルン", "チャオヤン", "トーヨー", "ジンユー", "ブリヂストン", "ダンロップ", "ヨコハマ"],
    type: ["ノーマル", "スタッドレス", "再生", "リグ"],
    groove: ["○", "△", "☓"],
    wear: ["○", "☓"],
    damage: ["○", "☓"],
    pressure: ["○", "☓"]
  };
  const SPARE_OPTION_ORDER = {
    maker: [...TIRE_OPTION_ORDER.maker],
    type: [...TIRE_OPTION_ORDER.type],
    condition: [...TIRE_OPTION_ORDER.groove]
  };

  const TRUCK_TYPE_LABELS = {
    low12: "大型低床",
    ten10: "大型10輪"
  };
  const DRIVER_SETTINGS_DOC_IDS = [
    "monthly_tire_company_settings_backup_dribers_slot1",
    "monthly_tire_company_settings_backup_drivers_slot1"
  ];
  const VEHICLE_SETTINGS_DOC_IDS = [
    "monthly_tire_company_settings_backup_vehicles_slot1"
  ];
  const SETTINGS_DOC_IDS = [...new Set([
    ...DRIVER_SETTINGS_DOC_IDS,
    ...VEHICLE_SETTINGS_DOC_IDS
  ])];
  const CSV_COLUMNS = [
    { key: "inspectionDate", label: "点検日" },
    { key: "vehicleNumber", label: "車両番号" },
    { key: "driverName", label: "乗務員" },
    { key: "truckType", label: "車種" },
    { key: "reportNote", label: "報告事項" },
    { key: "tireNo", label: "タイヤ番号" },
    { key: "maker", label: "メーカー" },
    { key: "type", label: "種類" },
    { key: "groove", label: "溝" },
    { key: "wear", label: "偏摩耗" },
    { key: "damage", label: "キズ" },
    { key: "pressure", label: "空気圧" },
    { key: "status", label: "状態" }
  ];
  const IMPORT_BASE_KEYS = ["inspectionDate", "vehicleNumber", "driverName", "truckType", "reportNote"];
  const sharedSettings = window.SharedAppSettings || null;

  const ui = {
    fromYear: document.getElementById("filterFromYear"),
    fromMonth: document.getElementById("filterFromMonth"),
    toYear: document.getElementById("filterToYear"),
    toMonth: document.getElementById("filterToMonth"),
    vehicle: document.getElementById("filterVehicle"),
    driver: document.getElementById("filterDriver"),
    report: document.getElementById("filterReport"),
    sortButtons: document.querySelectorAll(".sort-toggle"),
    reloadBtn: document.getElementById("reloadBtn"),
    csvBtn: document.getElementById("csvBtn"),
    csvImportBtn: document.getElementById("csvImportBtn"),
    csvImportInput: document.getElementById("csvImportInput"),
    printBtn: document.getElementById("printBtn"),
    statusText: document.getElementById("statusText"),
    errorText: document.getElementById("errorText"),
    resultBody: document.getElementById("resultBody"),
    printSheets: document.getElementById("printSheets"),
    editModal: document.getElementById("editModal"),
    editForm: document.getElementById("editForm"),
    editInspectionDate: document.getElementById("editInspectionDate"),
    editVehicleNumber: document.getElementById("editVehicleNumber"),
    editDriverName: document.getElementById("editDriverName"),
    editTruckType: document.getElementById("editTruckType"),
    editReportNote: document.getElementById("editReportNote"),
    editTireBody: document.getElementById("editTireBody"),
    editErrorText: document.getElementById("editErrorText"),
    saveEditBtn: document.getElementById("saveEditBtn"),
    cancelEditBtn: document.getElementById("cancelEditBtn")
  };

  const state = {
    rows: [],
    filteredRows: [],
    expandedDocIds: new Set(),
    db: null,
    collection: DEFAULT_COLLECTION,
    sort: {
      field: "inspectionDate",
      order: "desc"
    },
    editingDocId: null,
    editingCurrent: null,
    fieldOptions: null,
    dateFilterInitialized: false
  };

  function setError(message) {
    ui.errorText.textContent = message || "";
  }

  function setStatus(message) {
    ui.statusText.textContent = message;
  }

  function setEditError(message) {
    ui.editErrorText.textContent = message || "";
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function escapeHtmlWithLineBreaks(value) {
    return escapeHtml(value)
      .replaceAll("\r\n", "<br>")
      .replaceAll("\n", "<br>")
      .replaceAll("\r", "<br>");
  }

  function asDate(value) {
    if (!value) return null;
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (typeof value.toDate === "function") {
      const d = value.toDate();
      return Number.isNaN(d.getTime()) ? null : d;
    }
    if (typeof value === "object" && typeof value.seconds === "number") {
      const d = new Date(value.seconds * 1000);
      return Number.isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function formatDateTime(value) {
    const date = asDate(value);
    if (!date) return "-";
    return new Intl.DateTimeFormat("ja-JP", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }

  function normalizeText(value) {
    return String(value || "").trim();
  }

  function toHiraganaText(value) {
    const text = normalizeText(value);
    if (!text) return "";

    return text
      .normalize("NFKC")
      .replaceAll(" ", "")
      .replaceAll("　", "")
      .replace(/[\u30A1-\u30F6]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0x60));
  }

  function normalizeDriverNameKey(value) {
    const text = normalizeText(value);
    if (!text) return "";
    return text
      .normalize("NFKC")
      .replaceAll("　", " ")
      .replace(/\s+/g, " ");
  }

  function normalizeDriverNameCompactKey(value) {
    return normalizeDriverNameKey(value).replaceAll(" ", "");
  }

  function parseDriverNameWithReading(value) {
    const text = normalizeText(value);
    if (!text) {
      return { name: "", reading: "" };
    }

    const match = text.match(/^(.*?)\s*[（(]\s*([^（）()]+?)\s*[）)]\s*$/);
    if (!match) {
      return { name: text, reading: "" };
    }

    const name = normalizeText(match[1]) || text;
    const reading = toHiraganaText(match[2]);
    return { name, reading };
  }

  function addDriverReadingMapEntry(map, nameValue, readingValue) {
    const name = normalizeText(nameValue);
    const reading = toHiraganaText(readingValue);
    if (!name || !reading) return;

    const nameKey = normalizeDriverNameKey(name);
    const compactNameKey = normalizeDriverNameCompactKey(name);
    if (nameKey && !map.has(nameKey)) {
      map.set(nameKey, reading);
    }
    if (compactNameKey && !map.has(compactNameKey)) {
      map.set(compactNameKey, reading);
    }
  }

  function collectDriverReadingCandidates(source) {
    const safe = asPlainObject(source);
    const driver = asPlainObject(safe.driver);
    return [
      safe.driverNameFurigana,
      safe.driverFurigana,
      safe.driverNameHiragana,
      safe.driverNameKana,
      safe.driverKana,
      safe.driverReading,
      safe.driverNameReading,
      safe.furigana,
      safe.hurigana,
      safe.hiragana,
      safe.kana,
      safe.reading,
      safe.yomi,
      safe["ふりがな"],
      safe["フリガナ"],
      safe["読み"],
      safe["よみ"],
      driver.furigana,
      driver.hurigana,
      driver.hiragana,
      driver.kana,
      driver.reading,
      driver.yomi,
      driver.nameFurigana,
      driver.nameHiragana,
      driver.nameKana,
      driver["ふりがな"],
      driver["フリガナ"],
      driver["読み"],
      driver["よみ"]
    ];
  }

  function pickFirstNonEmpty(values, normalizer) {
    const normalize = typeof normalizer === "function" ? normalizer : normalizeText;
    for (const value of values || []) {
      const text = normalize(value);
      if (text) {
        return text;
      }
    }
    return "";
  }

  function extractDriverNameAndReading(source) {
    const safe = asPlainObject(source);
    const driver = asPlainObject(safe.driver);
    const rawName = pickFirstNonEmpty([
      safe.driverName,
      safe.driver,
      safe.name,
      safe.displayName,
      safe.label,
      safe.value,
      safe["乗務員"],
      safe["氏名"],
      safe["名前"],
      driver.driverName,
      driver.name,
      driver.displayName,
      driver["氏名"],
      driver["名前"]
    ]);
    const rawReading = pickFirstNonEmpty(collectDriverReadingCandidates(safe), toHiraganaText);
    const parsed = parseDriverNameWithReading(rawName);

    return {
      name: parsed.name || rawName,
      reading: rawReading || parsed.reading
    };
  }

  function collectDriverReadingMapFromValue(value, map, depth) {
    if (depth > 8 || value == null) return;

    if (Array.isArray(value)) {
      value.forEach((item) => collectDriverReadingMapFromValue(item, map, depth + 1));
      return;
    }

    if (typeof value === "string") {
      const parsed = parseDriverNameWithReading(value);
      addDriverReadingMapEntry(map, parsed.name, parsed.reading);
      return;
    }

    if (typeof value !== "object") return;

    const safe = asPlainObject(value);
    const entry = extractDriverNameAndReading(safe);
    addDriverReadingMapEntry(map, entry.name, entry.reading);

    Object.values(safe).forEach((child) => {
      collectDriverReadingMapFromValue(child, map, depth + 1);
    });
  }

  function collectDriverReadingMap(docSnaps) {
    const map = new Map();
    (docSnaps || []).forEach((docSnap) => {
      const data = asPlainObject(docSnap && typeof docSnap.data === "function" ? docSnap.data() : {});
      collectDriverReadingMapFromValue(data, map, 0);
    });
    return map;
  }

  async function loadSettingsDocsFromFirestore() {
    if (!state.db || !SETTINGS_DOC_IDS.length) {
      return [];
    }

    const settingsSnaps = await Promise.all(
      SETTINGS_DOC_IDS.map((docId) => state.db.collection(state.collection).doc(docId).get())
    );

    return settingsSnaps.filter((docSnap) => docSnap && docSnap.exists);
  }

  function extractVehicleNumber(source) {
    const safe = asPlainObject(source);
    return pickFirstNonEmpty([
      safe.vehicleNumber,
      safe.vehicleNo,
      safe.vehicleCode,
      safe.vehicle,
      safe.number,
      safe.label,
      safe.value,
      safe.name,
      safe.code,
      safe["車両番号"],
      safe["車番"],
      safe["車両"],
      safe["号車"]
    ]);
  }

  function normalizeMasterFieldKey(value) {
    return normalizeText(value)
      .normalize("NFKC")
      .replaceAll(" ", "")
      .replaceAll("　", "")
      .toLowerCase();
  }

  function isVehicleMasterContainerKey(value) {
    const key = normalizeMasterFieldKey(value);
    return [
      "items",
      "list",
      "entries",
      "values",
      "options",
      "data",
      "vehicles",
      "vehiclenumbers",
      "車両",
      "車両番号",
      "車番",
      "号車"
    ].includes(key);
  }

  function isDriverMasterContainerKey(value) {
    const key = normalizeMasterFieldKey(value);
    return [
      "items",
      "list",
      "entries",
      "values",
      "options",
      "data",
      "drivers",
      "drivernames",
      "乗務員",
      "氏名",
      "名前"
    ].includes(key);
  }

  function collectVehicleOptionsFromValue(value, targetSet, depth, allowStringValue) {
    if (depth > 8 || value == null) return;

    if (Array.isArray(value)) {
      value.forEach((item) => collectVehicleOptionsFromValue(item, targetSet, depth + 1, true));
      return;
    }

    if (typeof value === "string") {
      if (allowStringValue) {
        addOption(targetSet, value);
      }
      return;
    }

    if (typeof value !== "object") return;

    const safe = asPlainObject(value);
    addOption(targetSet, extractVehicleNumber(safe));
    Object.entries(safe).forEach(([key, child]) => {
      collectVehicleOptionsFromValue(child, targetSet, depth + 1, allowStringValue || isVehicleMasterContainerKey(key));
    });
  }

  function collectDriverEntriesFromValue(value, entryMap, depth, allowStringValue) {
    if (depth > 8 || value == null) return;

    if (Array.isArray(value)) {
      value.forEach((item) => collectDriverEntriesFromValue(item, entryMap, depth + 1, true));
      return;
    }

    if (typeof value === "string") {
      if (allowStringValue) {
        const parsed = parseDriverNameWithReading(value);
        const name = normalizeText(parsed.name);
        if (!name) return;
        const reading = toHiraganaText(parsed.reading);
        if (!entryMap.has(name)) {
          entryMap.set(name, reading);
        } else if (reading && !entryMap.get(name)) {
          entryMap.set(name, reading);
        }
      }
      return;
    }

    if (typeof value !== "object") return;

    const safe = asPlainObject(value);
    const entry = extractDriverNameAndReading(safe);
    const name = normalizeText(entry.name);
    const reading = toHiraganaText(entry.reading);
    if (name) {
      if (!entryMap.has(name)) {
        entryMap.set(name, reading);
      } else if (reading && !entryMap.get(name)) {
        entryMap.set(name, reading);
      }
    }

    Object.entries(safe).forEach(([key, child]) => {
      collectDriverEntriesFromValue(child, entryMap, depth + 1, allowStringValue || isDriverMasterContainerKey(key));
    });
  }

  function compareDriverOption(left, right, readingMap) {
    const leftName = normalizeText(left);
    const rightName = normalizeText(right);
    const leftReading = resolveDriverOptionReading(leftName, readingMap);
    const rightReading = resolveDriverOptionReading(rightName, readingMap);
    const readingResult = leftReading.localeCompare(rightReading, "ja", { sensitivity: "base" });
    if (readingResult !== 0) {
      return readingResult;
    }
    return leftName.localeCompare(rightName, "ja", { numeric: true, sensitivity: "base" });
  }

  function resolveDriverOptionReading(driverName, readingMap) {
    const name = normalizeText(driverName);
    const parsed = parseDriverNameWithReading(name);
    const baseName = normalizeText(parsed.name || name);
    const normalizedName = normalizeDriverNameKey(baseName);
    const compactName = normalizeDriverNameCompactKey(baseName);

    if (readingMap instanceof Map) {
      const mappedReading = toHiraganaText(
        readingMap.get(baseName) ||
        readingMap.get(normalizedName) ||
        readingMap.get(compactName)
      );
      if (mappedReading) {
        return mappedReading;
      }
    }

    return toHiraganaText(parsed.reading || baseName);
  }

  function collectMasterFieldOptions(settingsDocSnaps) {
    const vehicleSet = new Set();
    const driverEntryMap = new Map();

    (settingsDocSnaps || []).forEach((docSnap) => {
      const docId = normalizeText(docSnap && docSnap.id);
      const data = asPlainObject(docSnap && typeof docSnap.data === "function" ? docSnap.data() : {});

      if (VEHICLE_SETTINGS_DOC_IDS.includes(docId)) {
        collectVehicleOptionsFromValue(data, vehicleSet, 0, false);
      }
      if (DRIVER_SETTINGS_DOC_IDS.includes(docId)) {
        collectDriverEntriesFromValue(data, driverEntryMap, 0, false);
      }
    });

    const driverNames = Array.from(driverEntryMap.keys()).sort((a, b) => compareDriverOption(a, b, driverEntryMap));
    const localSharedOptions = getLocalSharedMasterOptions();

    return {
      vehicleNumber: mergePreferredOptions(localSharedOptions.vehicleNumber, sortOptions(vehicleSet)),
      driverName: sortDriverOptions(
        mergePreferredOptions(localSharedOptions.driverName, driverNames),
        mergeDriverReadingMaps(localSharedOptions.driverReadingMap, driverEntryMap)
      )
    };
  }

  function getLocalSharedMasterOptions() {
    if (!sharedSettings || typeof sharedSettings.ensureState !== "function") {
      return {
        vehicleNumber: [],
        driverName: [],
        driverReadingMap: new Map()
      };
    }

    const sharedState = sharedSettings.ensureState();
    const driverReadingMap = new Map();

    (sharedState.drivers || []).forEach((entry) => {
      const parsed = parseDriverNameWithReading(entry);
      const driverName = normalizeText(parsed.name);
      const driverReading = toHiraganaText(parsed.reading);
      if (!driverName || !driverReading) {
        return;
      }

      driverReadingMap.set(normalizeDriverNameKey(driverName), driverReading);
      driverReadingMap.set(normalizeDriverNameCompactKey(driverName), driverReading);
    });

    return {
      vehicleNumber: sortOptions(new Set(sharedState.vehicles || [])),
      driverName: sortDriverOptions(
        new Set((sharedState.drivers || []).map((entry) => parseDriverNameWithReading(entry).name).filter(Boolean)),
        driverReadingMap
      ),
      driverReadingMap
    };
  }

  function mergeDriverReadingMaps() {
    const merged = new Map();

    Array.from(arguments).forEach((source) => {
      if (!(source instanceof Map)) {
        return;
      }

      source.forEach((value, key) => {
        if (!merged.has(key)) {
          merged.set(key, value);
        }
      });
    });

    return merged;
  }

  function mergePreferredOptions(preferredValues, fallbackValues) {
    const merged = [];
    const seen = new Set();

    [preferredValues || [], fallbackValues || []].forEach((values) => {
      values.forEach((value) => {
        const text = normalizeText(value);
        if (!text || seen.has(text)) return;
        seen.add(text);
        merged.push(text);
      });
    });

    return merged;
  }

  function sortDriverOptions(values, driverReadingMap) {
    return Array.from(values || []).sort((a, b) => compareDriverOption(a, b, driverReadingMap || new Map()));
  }

  function resolveDriverSortKey(data, payload, current, driverName, driverReadingMap) {
    const parsedDriverName = parseDriverNameWithReading(driverName);
    const normalizedDriverName = normalizeDriverNameKey(parsedDriverName.name || driverName);
    const compactDriverName = normalizeDriverNameCompactKey(parsedDriverName.name || driverName);
    if (driverReadingMap instanceof Map) {
      const mappedReading = toHiraganaText(
        driverReadingMap.get(normalizedDriverName) || driverReadingMap.get(compactDriverName)
      );
      if (mappedReading) {
        return mappedReading;
      }
    }

    const sources = [current, payload, data];
    for (const source of sources) {
      const candidates = collectDriverReadingCandidates(source);
      for (const candidate of candidates) {
        const reading = toHiraganaText(candidate);
        if (reading) {
          return reading;
        }
      }
    }

    return parsedDriverName.reading || toHiraganaText(parsedDriverName.name || driverName);
  }

  function parseInspectionDateParts(value) {
    const text = normalizeText(value);
    if (!text) return null;

    const match = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!match) return null;

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;

    return {
      year,
      month,
      day,
      ym: year * 100 + month
    };
  }

  function getCurrentYearMonth() {
    const now = new Date();
    return {
      year: String(now.getFullYear()),
      month: String(now.getMonth() + 1).padStart(2, "0")
    };
  }

  function asPlainObject(value) {
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
  }

  function cloneValue(value) {
    if (value == null) {
      return {};
    }
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (error) {
      if (typeof structuredClone === "function") {
        try {
          return structuredClone(value);
        } catch (cloneError) {
          return asPlainObject(value);
        }
      }
      return asPlainObject(value);
    }
  }

  function createEmptyFieldOptionSets() {
    return {
      vehicleNumber: new Set(),
      driverName: new Set(),
      truckType: new Set(KNOWN_TRUCK_TYPES),
      spare: {
        maker: new Set(),
        type: new Set(),
        condition: new Set()
      },
      tires: {
        maker: new Set(),
        type: new Set(),
        groove: new Set(),
        wear: new Set(),
        damage: new Set(),
        pressure: new Set()
      }
    };
  }

  function addOption(targetSet, value) {
    const text = normalizeText(value);
    if (text) {
      targetSet.add(text);
    }
  }

  function sortOptions(values) {
    return Array.from(values).sort((a, b) => String(a).localeCompare(String(b), "ja"));
  }

  function collectFieldOptions(rows, masterFieldOptions, driverReadingMap) {
    const optionSets = createEmptyFieldOptionSets();

    rows.forEach((row) => {
      addOption(optionSets.vehicleNumber, row.vehicleNumber);
      addOption(optionSets.driverName, row.driverName);
      addOption(optionSets.truckType, row.truckType);

      const current = asPlainObject(row.current);
      const spare = asPlainObject(current.spare);
      addOption(optionSets.spare.maker, spare.maker);
      addOption(optionSets.spare.type, spare.type);
      addOption(optionSets.spare.condition, spare.condition);

      const tires = asPlainObject(current.tires);
      Object.values(tires).forEach((tire) => {
        const safeTire = asPlainObject(tire);
        TIRE_FIELDS.forEach((field) => {
          addOption(optionSets.tires[field], safeTire[field]);
        });
      });
    });

    const rowVehicleOptions = sortOptions(optionSets.vehicleNumber);
    const rowDriverOptions = sortDriverOptions(optionSets.driverName, driverReadingMap);
    const masterVehicles = masterFieldOptions ? masterFieldOptions.vehicleNumber : [];
    const masterDrivers = masterFieldOptions ? masterFieldOptions.driverName : [];

    return {
      vehicleNumber: mergePreferredOptions(masterVehicles, rowVehicleOptions),
      driverName: sortDriverOptions(mergePreferredOptions(masterDrivers, rowDriverOptions), driverReadingMap),
      truckType: sortOptions(optionSets.truckType),
      spare: {
        maker: [...SPARE_OPTION_ORDER.maker],
        type: [...SPARE_OPTION_ORDER.type],
        condition: [...SPARE_OPTION_ORDER.condition]
      },
      tires: {
        maker: [...TIRE_OPTION_ORDER.maker],
        type: [...TIRE_OPTION_ORDER.type],
        groove: [...TIRE_OPTION_ORDER.groove],
        wear: [...TIRE_OPTION_ORDER.wear],
        damage: [...TIRE_OPTION_ORDER.damage],
        pressure: [...TIRE_OPTION_ORDER.pressure]
      }
    };
  }

  function buildSelectOptionsHtml(values, currentValue, placeholderText) {
    const selected = normalizeText(currentValue);
    const ordered = [];
    const unique = new Set();

    (values || []).forEach((value) => {
      const text = normalizeText(value);
      if (!text || unique.has(text)) return;
      unique.add(text);
      ordered.push(text);
    });

    if (selected && !unique.has(selected)) {
      unique.add(selected);
      ordered.push(selected);
    }

    const placeholder = escapeHtml(placeholderText || "選択してください");

    return [
      `<option value="">${placeholder}</option>`,
      ...ordered.map((value) => (
        `<option value="${escapeHtml(value)}"${value === selected ? " selected" : ""}>${escapeHtml(value)}</option>`
      ))
    ].join("");
  }

  function canonicalizeTireValue(field, value) {
    const text = normalizeText(value);
    if (!text) return "";

    const compact = text.replaceAll(" ", "").replaceAll("　", "");
    const lower = compact.toLowerCase();

    const isCircle = ["○", "◯", "o", "まる", "丸"].includes(compact) || lower === "o";
    const isCross = ["☓", "×", "✕", "x", "バツ", "ばつ"].includes(compact) || lower === "x";
    const isTriangle = ["△", "▲", "三角", "さんかく"].includes(compact);

    if (field === "groove") {
      if (isCircle) return "○";
      if (isTriangle) return "△";
      if (isCross) return "☓";
      return text;
    }

    if (field === "wear" || field === "damage" || field === "pressure") {
      if (isCircle) return "○";
      if (isCross) return "☓";
      return text;
    }

    return text;
  }

  function fillSelect(selectNode, values, currentValue, placeholderText) {
    if (!selectNode) return;
    selectNode.innerHTML = buildSelectOptionsHtml(values, currentValue, placeholderText);
  }

  function refreshFilterSelectOptions() {
    const vehicleOptions = state.fieldOptions ? state.fieldOptions.vehicleNumber : [];
    const driverOptions = state.fieldOptions ? state.fieldOptions.driverName : [];
    const currentVehicle = normalizeText(ui.vehicle.value);
    const currentDriver = normalizeText(ui.driver.value);
    const selectedVehicle = vehicleOptions.includes(currentVehicle) ? currentVehicle : "";
    const selectedDriver = driverOptions.includes(currentDriver) ? currentDriver : "";

    fillSelect(ui.vehicle, vehicleOptions, selectedVehicle, "すべて");
    fillSelect(ui.driver, driverOptions, selectedDriver, "すべて");
  }

  function collectInspectionDateIndex(rows) {
    const yearSet = new Set();
    const allMonthSet = new Set();
    const monthSetsByYear = new Map();

    rows.forEach((row) => {
      const parts = parseInspectionDateParts(row.inspectionDate);
      if (!parts) return;

      const year = String(parts.year);
      const month = String(parts.month).padStart(2, "0");
      yearSet.add(year);
      allMonthSet.add(month);

      if (!monthSetsByYear.has(year)) {
        monthSetsByYear.set(year, new Set());
      }
      monthSetsByYear.get(year).add(month);
    });

    const years = Array.from(yearSet).sort((a, b) => Number(a) - Number(b));
    const allMonths = Array.from(allMonthSet).sort((a, b) => Number(a) - Number(b));
    const monthsByYear = new Map();

    monthSetsByYear.forEach((monthSet, year) => {
      monthsByYear.set(year, Array.from(monthSet).sort((a, b) => Number(a) - Number(b)));
    });

    return { years, allMonths, monthsByYear };
  }

  function getMonthOptionsForYear(dateIndex, yearValue) {
    const year = normalizeText(yearValue);
    if (!year) {
      return dateIndex.allMonths;
    }
    return dateIndex.monthsByYear.get(year) || [];
  }

  function buildMonthOptionsHtml(monthValues, selectedValue) {
    const selected = normalizeText(selectedValue);
    return [
      '<option value="">すべて</option>',
      ...(monthValues || []).map((value) => {
        const monthLabel = `${Number(value)}月`;
        return `<option value="${value}"${value === selected ? " selected" : ""}>${monthLabel}</option>`;
      })
    ].join("");
  }

  function syncMonthSelectDisabledStates() {
    const fromYear = normalizeText(ui.fromYear.value);
    const toYear = normalizeText(ui.toYear.value);

    ui.fromMonth.disabled = !fromYear;
    ui.toMonth.disabled = !toYear;

    if (!fromYear) {
      ui.fromMonth.value = "";
    }
    if (!toYear) {
      ui.toMonth.value = "";
    }
  }

  function refreshDateFilterSelectOptions() {
    const dateIndex = collectInspectionDateIndex(state.rows);
    const currentYm = getCurrentYearMonth();
    const rawFromYear = state.dateFilterInitialized
      ? normalizeText(ui.fromYear.value)
      : currentYm.year;
    const rawToYear = state.dateFilterInitialized
      ? normalizeText(ui.toYear.value)
      : currentYm.year;
    const rawFromMonth = state.dateFilterInitialized
      ? normalizeText(ui.fromMonth.value)
      : currentYm.month;
    const rawToMonth = state.dateFilterInitialized
      ? normalizeText(ui.toMonth.value)
      : currentYm.month;
    const selectedFromYear = dateIndex.years.includes(rawFromYear) ? rawFromYear : "";
    const selectedToYear = dateIndex.years.includes(rawToYear) ? rawToYear : "";
    const fromMonthOptions = getMonthOptionsForYear(dateIndex, selectedFromYear);
    const toMonthOptions = getMonthOptionsForYear(dateIndex, selectedToYear);
    const selectedFromMonth = fromMonthOptions.includes(rawFromMonth) ? rawFromMonth : "";
    const selectedToMonth = toMonthOptions.includes(rawToMonth) ? rawToMonth : "";

    fillSelect(ui.fromYear, dateIndex.years, selectedFromYear, "すべて");
    fillSelect(ui.toYear, dateIndex.years, selectedToYear, "すべて");

    ui.fromMonth.innerHTML = buildMonthOptionsHtml(fromMonthOptions, selectedFromMonth);
    ui.toMonth.innerHTML = buildMonthOptionsHtml(toMonthOptions, selectedToMonth);
    syncMonthSelectDisabledStates();
    state.dateFilterInitialized = true;
  }

  function fillTruckTypeSelect(currentValue) {
    const options = state.fieldOptions ? state.fieldOptions.truckType : KNOWN_TRUCK_TYPES;
    const selected = normalizeText(currentValue);
    const unique = new Set(options || []);
    KNOWN_TRUCK_TYPES.forEach((value) => unique.add(value));
    if (selected) {
      unique.add(selected);
    }

    const sorted = sortOptions(unique);
    ui.editTruckType.innerHTML = [
      '<option value="">選択してください</option>',
      ...sorted.map((value) => {
        const label = TRUCK_TYPE_LABELS[value] || value;
        return `<option value="${escapeHtml(value)}"${value === selected ? " selected" : ""}>${escapeHtml(label)}</option>`;
      })
    ].join("");
  }

  function getActiveTireIds(truckType, tires) {
    if (truckType === "ten10") return Array.from({ length: 10 }, (_, i) => String(i + 1));
    if (truckType === "low12") return Array.from({ length: 12 }, (_, i) => String(i + 1));

    return Object.keys(tires || {})
      .map((key) => String(key))
      .sort((a, b) => Number(a) - Number(b));
  }

  const HIDDEN_DOC_IDS = new Set(SETTINGS_DOC_IDS);

  function isSettingsBackupDocId(docId) {
    const normalizedId = String(docId || "");
    return HIDDEN_DOC_IDS.has(normalizedId) || normalizedId.includes("_settings_backup_");
  }

  function extractRow(docSnap, driverReadingMap) {
    const data = docSnap.data() || {};
    const payload = data.state || {};
    const current = payload.current || data.current || {};

    const updatedAt = data.updatedAt || payload.savedAt || current.updatedAt || null;
    const inspectionDate = normalizeText(current.inspectionDate);
    const vehicleNumber = normalizeText(current.vehicleNumber);
    const driverName = normalizeText(current.driverName);
    const driverSortKey = resolveDriverSortKey(data, payload, current, driverName, driverReadingMap);
    const truckType = normalizeText(current.truckType);
    const reportNote = normalizeText(current.reportNote);

    return {
      docId: docSnap.id,
      updatedAt,
      inspectionDate,
      vehicleNumber,
      driverName,
      driverSortKey,
      truckType,
      truckTypeLabel: TRUCK_TYPE_LABELS[truckType] || truckType || "-",
      reportNote,
      payload,
      current
    };
  }

  function toTimestampMs(value) {
    const date = asDate(value);
    return date ? date.getTime() : Number.NEGATIVE_INFINITY;
  }

  function toStableComparableValue(value) {
    if (Array.isArray(value)) {
      return value.map((item) => toStableComparableValue(item));
    }

    if (value && typeof value === "object") {
      const sorted = {};
      Object.keys(value)
        .sort((a, b) => a.localeCompare(b, "ja"))
        .forEach((key) => {
          sorted[key] = toStableComparableValue(value[key]);
        });
      return sorted;
    }

    return value;
  }

  function buildMonthlyOverwriteKey(row) {
    const parts = parseInspectionDateParts(row.inspectionDate);
    if (!parts) return "";

    const monthKey = `${parts.year}-${String(parts.month).padStart(2, "0")}`;
    const current = cloneValue(row.current || {});
    if (current && typeof current === "object") {
      delete current.inspectionDate;
      delete current.vehicleNumber;
      delete current.driverName;
      delete current.truckType;
      delete current.reportNote;
      delete current.updatedAt;
      delete current.savedAt;
    }

    const keyPayload = toStableComparableValue({
      month: monthKey,
      vehicleNumber: normalizeText(row.vehicleNumber),
      driverName: normalizeText(row.driverName),
      truckType: normalizeText(row.truckType),
      reportNote: normalizeText(row.reportNote),
      current
    });

    return JSON.stringify(keyPayload);
  }

  async function removeMonthlyDuplicateRows(rows) {
    if (!rows.length) return 0;

    const groups = new Map();
    rows.forEach((row) => {
      const key = buildMonthlyOverwriteKey(row);
      if (!key) return;
      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key).push(row);
    });

    const toDelete = [];
    groups.forEach((group) => {
      if (group.length <= 1) return;
      const sorted = group
        .slice()
        .sort((a, b) => {
          const timeDiff = toTimestampMs(b.updatedAt) - toTimestampMs(a.updatedAt);
          if (timeDiff !== 0) return timeDiff;
          return String(a.docId || "").localeCompare(String(b.docId || ""), "ja");
        });
      toDelete.push(...sorted.slice(1));
    });

    if (!toDelete.length) return 0;

    const collectionRef = state.db.collection(state.collection);
    const chunkSize = 450;
    for (let i = 0; i < toDelete.length; i += chunkSize) {
      const batch = state.db.batch();
      const chunk = toDelete.slice(i, i + chunkSize);
      chunk.forEach((row) => {
        batch.delete(collectionRef.doc(row.docId));
      });
      await batch.commit();
    }

    return toDelete.length;
  }

  function getRowByDocId(docId) {
    return state.rows.find((row) => row.docId === docId) || null;
  }

  function renderEditTireInputs() {
    const current = asPlainObject(state.editingCurrent);
    const tires = asPlainObject(current.tires);
    const spare = asPlainObject(current.spare);
    current.tires = tires;
    current.spare = spare;
    state.editingCurrent = current;

    const truckType = normalizeText(ui.editTruckType.value) || normalizeText(current.truckType);
    const ids = getActiveTireIds(truckType, tires);

    const tireRows = ids.map((id) => {
      const tire = asPlainObject(tires[id]);
      tires[id] = tire;

      const tireInputs = TIRE_FIELDS.map((field) => {
        const values = state.fieldOptions ? state.fieldOptions.tires[field] : [];
        const currentValue = normalizeText(tire[field]);
        const canonicalValue = canonicalizeTireValue(field, currentValue);
        const selectedValue = values.includes(canonicalValue) ? canonicalValue : currentValue;
        return (
          `<td><select data-tire-id="${escapeHtml(id)}" data-tire-field="${field}">` +
          `${buildSelectOptionsHtml(values, selectedValue, "選択")}` +
          "</select></td>"
        );
      }).join("");

      return `<tr><td>${escapeHtml(id)}</td>${tireInputs}<td></td></tr>`;
    });

    if (!tireRows.length) {
      tireRows.push('<tr><td colspan="8">タイヤデータがありません。</td></tr>');
    }

    const spareMaker = normalizeText(spare.maker);
    const spareType = normalizeText(spare.type);
    const spareCondition = canonicalizeTireValue("groove", spare.condition);
    const spareRow = [
      '<tr class="spare-row">',
      "<td>スペア</td>",
      `<td><select data-spare-field="maker">${buildSelectOptionsHtml(state.fieldOptions ? state.fieldOptions.spare.maker : [], spareMaker, "選択")}</select></td>`,
      `<td><select data-spare-field="type">${buildSelectOptionsHtml(state.fieldOptions ? state.fieldOptions.spare.type : [], spareType, "選択")}</select></td>`,
      "<td></td>",
      "<td></td>",
      "<td></td>",
      "<td></td>",
      `<td><select data-spare-field="condition">${buildSelectOptionsHtml(state.fieldOptions ? state.fieldOptions.spare.condition : [], spareCondition, "選択")}</select></td>`,
      "</tr>"
    ].join("");

    ui.editTireBody.innerHTML = [...tireRows, spareRow].join("");
  }

  function syncEditingCurrentFromForm() {
    const current = asPlainObject(state.editingCurrent);
    current.inspectionDate = normalizeText(ui.editInspectionDate.value);
    current.vehicleNumber = normalizeText(ui.editVehicleNumber.value);
    current.driverName = normalizeText(ui.editDriverName.value);
    current.truckType = normalizeText(ui.editTruckType.value);
    current.reportNote = normalizeText(ui.editReportNote.value);

    const spare = asPlainObject(current.spare);
    const spareMakerNode = ui.editTireBody.querySelector('[data-spare-field="maker"]');
    const spareTypeNode = ui.editTireBody.querySelector('[data-spare-field="type"]');
    const spareConditionNode = ui.editTireBody.querySelector('[data-spare-field="condition"]');
    spare.maker = normalizeText(spareMakerNode ? spareMakerNode.value : spare.maker);
    spare.type = normalizeText(spareTypeNode ? spareTypeNode.value : spare.type);
    spare.condition = canonicalizeTireValue("groove", spareConditionNode ? spareConditionNode.value : spare.condition);
    current.spare = spare;

    const tires = asPlainObject(current.tires);
    ui.editTireBody.querySelectorAll("[data-tire-id][data-tire-field]").forEach((node) => {
      const tireId = normalizeText(node.dataset.tireId);
      const tireField = normalizeText(node.dataset.tireField);
      if (!tireId || !tireField) return;
      const selectedValue = normalizeText(node.value);
      if (!selectedValue) return;
      const tire = asPlainObject(tires[tireId]);
      tire[tireField] = canonicalizeTireValue(tireField, selectedValue);
      tires[tireId] = tire;
    });
    current.tires = tires;

    state.editingCurrent = current;
    return current;
  }

  function openEditModal(docId) {
    const row = getRowByDocId(docId);
    if (!row) {
      setError("編集対象データが見つかりません。");
      return;
    }

    const current = cloneValue(row.current || {});
    current.tires = asPlainObject(current.tires);
    current.spare = asPlainObject(current.spare);

    state.editingDocId = docId;
    state.editingCurrent = current;
    setEditError("");
    ui.editInspectionDate.value = normalizeText(current.inspectionDate);
    fillSelect(ui.editVehicleNumber, state.fieldOptions ? state.fieldOptions.vehicleNumber : [], current.vehicleNumber, "選択");
    fillSelect(ui.editDriverName, state.fieldOptions ? state.fieldOptions.driverName : [], current.driverName, "選択");
    fillTruckTypeSelect(current.truckType);
    ui.editReportNote.value = normalizeText(current.reportNote);
    renderEditTireInputs();
    ui.editModal.hidden = false;
    ui.editVehicleNumber.focus();
  }

  function closeEditModal() {
    state.editingDocId = null;
    state.editingCurrent = null;
    setEditError("");
    ui.editForm.reset();
    ui.editTireBody.innerHTML = "";
    ui.editModal.hidden = true;
  }

  async function saveEditedRow() {
    const docId = state.editingDocId;
    if (!docId) return;

    const fieldValue = window.firebase.firestore.FieldValue.serverTimestamp();
    const updatedCurrent = cloneValue(syncEditingCurrentFromForm());
    const docRef = state.db.collection(state.collection).doc(docId);
    const updateCandidates = [
      {
        updatedAt: fieldValue,
        "state.savedAt": fieldValue,
        "state.current": updatedCurrent,
        current: updatedCurrent
      },
      {
        updatedAt: fieldValue,
        "state.savedAt": fieldValue,
        "state.current": updatedCurrent
      },
      {
        updatedAt: fieldValue,
        current: updatedCurrent
      }
    ];

    ui.saveEditBtn.disabled = true;
    setError("");
    setEditError("");

    try {
      let lastError = null;
      let updated = false;
      for (const payload of updateCandidates) {
        try {
          await docRef.update(payload);
          updated = true;
          break;
        } catch (error) {
          lastError = error;
        }
      }
      if (!updated) {
        throw lastError || new Error("更新に失敗しました。");
      }

      closeEditModal();
      setStatus("保存しました。");
      await loadRows();
    } finally {
      ui.saveEditBtn.disabled = false;
    }
  }

  async function deleteRow(docId) {
    const row = getRowByDocId(docId);
    const summary = row
      ? `点検日: ${row.inspectionDate || "-"} / 車両番号: ${row.vehicleNumber || "-"} / 乗務員: ${row.driverName || "-"}`
      : `DocID: ${docId}`;

    if (!window.confirm(`このデータを削除しますか？\n${summary}\n\n削除後は元に戻せません。`)) {
      return;
    }

    setError("");
    await state.db.collection(state.collection).doc(docId).delete();
    state.expandedDocIds.delete(docId);
    if (state.editingDocId === docId) {
      closeEditModal();
    }
    await loadRows();
  }

  function renderDetailHtml(row) {
    const current = row.current || {};
    const tires = current.tires || {};
    const ids = getActiveTireIds(current.truckType, tires);

    const tireRows = ids.map((id) => {
      const t = tires[id] || {};
      return [
        `<td>${escapeHtml(id)}</td>`,
        `<td>${escapeHtml(t.maker || "-")}</td>`,
        `<td>${escapeHtml(t.type || "-")}</td>`,
        `<td>${escapeHtml(t.groove || "-")}</td>`,
        `<td>${escapeHtml(t.wear || "-")}</td>`,
        `<td>${escapeHtml(t.damage || "-")}</td>`,
        `<td>${escapeHtml(t.pressure || "-")}</td>`,
        "<td></td>"
      ].join("");
    });

    const spare = current.spare || {};
    const spareCondition = canonicalizeTireValue("groove", spare.condition);
    const spareRow = [
      "<tr class=\"spare-row\">",
      "<td>スペア</td>",
      `<td>${escapeHtml(spare.maker || "-")}</td>`,
      `<td>${escapeHtml(spare.type || "-")}</td>`,
      "<td></td>",
      "<td></td>",
      "<td></td>",
      "<td></td>",
      `<td>${escapeHtml(spareCondition || "-")}</td>`,
      "</tr>"
    ].join("");
    const tireBodyRows = tireRows.map((r) => `<tr>${r}</tr>`).join("") || "<tr><td colspan=\"8\">データなし</td></tr>";

    return [
      '<div class="detail">',
      "<h3>詳細</h3>",
      '<table class="detail-table">',
      "<thead><tr><th>タイヤ番号</th><th>メーカー</th><th>種類</th><th>溝</th><th>偏摩耗</th><th>キズ</th><th>空気圧</th><th>状態</th></tr></thead>",
      `<tbody>${tireBodyRows}${spareRow}</tbody>`,
      "</table>",
      `<p><strong>報告事項:</strong> ${escapeHtml(row.reportNote || "-")}</p>`,
      "</div>"
    ].join("");
  }

  function getPrintTireRows(row) {
    const current = asPlainObject(row.current);
    const tires = asPlainObject(current.tires);
    const ids = getActiveTireIds(current.truckType, tires);
    const rows = ids.map((id) => {
      const tire = asPlainObject(tires[id]);
      const status = normalizeText(tire.state || tire.condition);
      return {
        tireNo: id,
        maker: normalizeText(tire.maker) || "-",
        type: normalizeText(tire.type) || "-",
        groove: normalizeText(tire.groove) || "-",
        wear: normalizeText(tire.wear) || "-",
        damage: normalizeText(tire.damage) || "-",
        pressure: normalizeText(tire.pressure) || "-",
        status: status || "-"
      };
    });

    const spare = asPlainObject(current.spare);
    const spareCondition = canonicalizeTireValue("groove", spare.condition);
    rows.push({
      tireNo: "スペア",
      maker: normalizeText(spare.maker) || "-",
      type: normalizeText(spare.type) || "-",
      groove: "-",
      wear: "-",
      damage: "-",
      pressure: "-",
      status: normalizeText(spareCondition) || "-"
    });
    return rows;
  }

  function renderPrintSheetHtml(row) {
    const tireRows = getPrintTireRows(row);
    const tireRowsHtml = tireRows.map((tire) => (
      "<tr>" +
        `<td>${escapeHtml(tire.tireNo)}</td>` +
        `<td>${escapeHtml(tire.maker)}</td>` +
        `<td>${escapeHtml(tire.type)}</td>` +
        `<td>${escapeHtml(tire.groove)}</td>` +
        `<td>${escapeHtml(tire.wear)}</td>` +
        `<td>${escapeHtml(tire.damage)}</td>` +
        `<td>${escapeHtml(tire.pressure)}</td>` +
        `<td>${escapeHtml(tire.status)}</td>` +
      "</tr>"
    )).join("");

    return [
      '<section class="print-sheet-page">',
      '<div class="print-basic-grid">',
      `<p><span class="print-label">点検日</span><span class="print-value">${escapeHtml(row.inspectionDate || "-")}</span></p>`,
      `<p><span class="print-label">車両番号</span><span class="print-value">${escapeHtml(row.vehicleNumber || "-")}</span></p>`,
      `<p><span class="print-label">乗務員</span><span class="print-value">${escapeHtml(row.driverName || "-")}</span></p>`,
      `<p><span class="print-label">車種</span><span class="print-value">${escapeHtml(row.truckTypeLabel || "-")}</span></p>`,
      "</div>",
      '<div class="print-report-block">',
      '<div class="print-label">報告事項</div>',
      `<div class="print-report-text">${escapeHtmlWithLineBreaks(row.reportNote || "-")}</div>`,
      "</div>",
      '<table class="print-tire-table">',
      "<thead><tr><th>タイヤ番号</th><th>メーカー</th><th>種類</th><th>溝</th><th>偏摩耗</th><th>キズ</th><th>空気圧</th><th>状態</th></tr></thead>",
      `<tbody>${tireRowsHtml}</tbody>`,
      "</table>",
      "</section>"
    ].join("");
  }

  function renderPrintSheets() {
    if (!ui.printSheets) return false;

    if (!state.filteredRows.length) {
      ui.printSheets.innerHTML = "";
      return false;
    }

    ui.printSheets.innerHTML = state.filteredRows.map((row) => renderPrintSheetHtml(row)).join("");
    return true;
  }

  function renderTable() {
    const html = [];

    state.filteredRows.forEach((row) => {
      const expanded = state.expandedDocIds.has(row.docId);
      html.push(
        `<tr class="data-row" data-doc-id="${escapeHtml(row.docId)}">` +
          `<td>${escapeHtml(formatDateTime(row.updatedAt))}</td>` +
          `<td>${escapeHtml(row.inspectionDate || "-")}</td>` +
          `<td>${escapeHtml(row.vehicleNumber || "-")}</td>` +
          `<td>${escapeHtml(row.driverName || "-")}</td>` +
          `<td>${escapeHtml(row.truckTypeLabel)}</td>` +
          `<td>${escapeHtml(row.reportNote || "-")}</td>` +
          `<td class="actions-cell no-print"><div class="row-actions">` +
            `<button type="button" class="btn-small" data-action="edit" data-doc-id="${escapeHtml(row.docId)}">編集</button>` +
            `<button type="button" class="btn-small btn-danger" data-action="delete" data-doc-id="${escapeHtml(row.docId)}">削除</button>` +
          `</div></td>` +
        "</tr>"
      );

      if (expanded) {
        html.push(
          `<tr class="detail-row"><td colspan="7">${renderDetailHtml(row)}</td></tr>`
        );
      }
    });

    ui.resultBody.innerHTML = html.join("");

    if (!state.filteredRows.length) {
      ui.resultBody.innerHTML = '<tr><td colspan="7">該当データがありません。</td></tr>';
    }
  }

  function containsIgnoreCase(text, keyword) {
    if (!keyword) return true;
    return String(text || "").toLowerCase().includes(keyword.toLowerCase());
  }

  function compareNullableText(a, b) {
    const left = normalizeText(a);
    const right = normalizeText(b);
    if (!left && !right) return 0;
    if (!left) return 1;
    if (!right) return -1;
    return left.localeCompare(right, "ja", { numeric: true, sensitivity: "base" });
  }

  function updateSortButtons() {
    ui.sortButtons.forEach((button) => {
      const field = normalizeText(button.dataset.sortField);
      const active = state.sort.field === field;
      const symbol = active ? (state.sort.order === "asc" ? "▲" : "▼") : "↕";
      button.textContent = symbol;
      button.setAttribute("aria-pressed", active ? "true" : "false");
    });
  }

  function toggleSort(field) {
    const targetField = normalizeText(field);
    if (!targetField) return;

    if (state.sort.field === targetField) {
      state.sort.order = state.sort.order === "asc" ? "desc" : "asc";
    } else {
      state.sort.field = targetField;
      state.sort.order = "desc";
    }
    applyFilters();
  }

  function sortFilteredRows() {
    const sortField = state.sort.field || "inspectionDate";
    const sortMultiplier = state.sort.order === "asc" ? 1 : -1;

    state.filteredRows.sort((a, b) => {
      let result = 0;
      if (sortField === "inspectionDate") {
        result = compareNullableText(a.inspectionDate, b.inspectionDate);
      } else if (sortField === "vehicleNumber") {
        result = compareNullableText(a.vehicleNumber, b.vehicleNumber);
      } else if (sortField === "driverName") {
        result = compareNullableText(a.driverSortKey, b.driverSortKey);
        if (result === 0) {
          result = compareNullableText(a.driverName, b.driverName);
        }
      }

      if (result === 0) {
        const dateA = asDate(a.updatedAt);
        const dateB = asDate(b.updatedAt);
        if (dateA && dateB) {
          result = dateB.getTime() - dateA.getTime();
        } else if (dateA) {
          result = -1;
        } else if (dateB) {
          result = 1;
        } else {
          result = String(a.docId || "").localeCompare(String(b.docId || ""), "ja");
        }
      }

      return result * sortMultiplier;
    });
  }

  function buildYearMonthBoundary(yearValue, monthValue, fallbackMonth) {
    const yearText = normalizeText(yearValue);
    if (!yearText) return null;

    const year = Number(yearText);
    const month = Number(normalizeText(monthValue) || fallbackMonth);
    if (!Number.isInteger(year) || !Number.isInteger(month)) return null;
    if (month < 1 || month > 12) return null;

    return year * 100 + month;
  }

  function applyFilters() {
    const fromYear = normalizeText(ui.fromYear.value);
    const fromMonth = normalizeText(ui.fromMonth.value);
    const toYear = normalizeText(ui.toYear.value);
    const toMonth = normalizeText(ui.toMonth.value);
    const vehicle = normalizeText(ui.vehicle.value);
    const driver = normalizeText(ui.driver.value);
    const report = normalizeText(ui.report.value);
    const fromYm = buildYearMonthBoundary(fromYear, fromMonth, 1);
    const toYm = buildYearMonthBoundary(toYear, toMonth, 12);

    state.filteredRows = state.rows.filter((row) => {
      const rowDate = parseInspectionDateParts(row.inspectionDate);
      if (fromYm != null && (!rowDate || rowDate.ym < fromYm)) return false;
      if (toYm != null && (!rowDate || rowDate.ym > toYm)) return false;
      if (vehicle && row.vehicleNumber !== vehicle) return false;
      if (driver && row.driverName !== driver) return false;
      if (!containsIgnoreCase(row.reportNote, report)) return false;
      return true;
    });

    sortFilteredRows();
    updateSortButtons();
    setStatus(`表示件数: ${state.filteredRows.length} / 全${state.rows.length}件`);
    renderTable();
    renderPrintSheets();
  }

  function csvEscape(value) {
    const text = String(value == null ? "" : value).replaceAll('"', '""');
    return `"${text}"`;
  }

  function normalizeCsvHeaderKey(value) {
    return normalizeText(value)
      .replace(/^\uFEFF/, "")
      .normalize("NFKC")
      .replaceAll(" ", "")
      .replaceAll("　", "");
  }

  function normalizeCsvLiteral(value) {
    const text = normalizeText(value);
    if (!text) return "";

    const normalized = text.normalize("NFKC");
    if (normalized === "-" || normalized === "－") {
      return "";
    }
    return text;
  }

  function normalizeTruckTypeValue(value) {
    const text = normalizeCsvLiteral(value);
    if (!text) return "";
    if (KNOWN_TRUCK_TYPES.includes(text)) return text;

    const normalized = text
      .normalize("NFKC")
      .replaceAll(" ", "")
      .replaceAll("　", "")
      .toLowerCase();

    for (const [code, label] of Object.entries(TRUCK_TYPE_LABELS)) {
      const normalizedLabel = String(label)
        .normalize("NFKC")
        .replaceAll(" ", "")
        .replaceAll("　", "")
        .toLowerCase();
      if (normalized === normalizedLabel || normalized === code.toLowerCase()) {
        return code;
      }
    }

    return text;
  }

  function normalizeImportedInspectionDate(value) {
    const text = normalizeCsvLiteral(value);
    if (!text) return "";

    const normalized = text
      .normalize("NFKC")
      .replaceAll(".", "/")
      .replaceAll("年", "/")
      .replaceAll("月", "/")
      .replaceAll("日", "");
    const parts = parseInspectionDateParts(normalized);
    if (!parts) {
      return normalized;
    }

    return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
  }

  function normalizeImportedTireNo(value) {
    const text = normalizeCsvLiteral(value);
    if (!text) return "";

    const normalized = text
      .normalize("NFKC")
      .replaceAll(" ", "")
      .replaceAll("　", "");
    const lower = normalized.toLowerCase();

    if (normalized === "スペア" || lower === "spare") {
      return "スペア";
    }
    if (/^\d+$/.test(normalized)) {
      return String(Number(normalized));
    }

    return text;
  }

  function parseCsvRows(text) {
    const source = String(text == null ? "" : text).replace(/^\uFEFF/, "");
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;

    for (let i = 0; i < source.length; i += 1) {
      const char = source[i];

      if (inQuotes) {
        if (char === '"') {
          if (source[i + 1] === '"') {
            cell += '"';
            i += 1;
          } else {
            inQuotes = false;
          }
        } else {
          cell += char;
        }
        continue;
      }

      if (char === '"') {
        inQuotes = true;
        continue;
      }
      if (char === ",") {
        row.push(cell);
        cell = "";
        continue;
      }
      if (char === "\r") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
        if (source[i + 1] === "\n") {
          i += 1;
        }
        continue;
      }
      if (char === "\n") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
        continue;
      }

      cell += char;
    }

    if (inQuotes) {
      throw new Error("CSVの引用符が閉じられていません。");
    }
    if (cell || row.length) {
      row.push(cell);
      rows.push(row);
    }

    return rows;
  }

  function getCsvColumnIndexes(headerRow) {
    const normalizedHeaders = (headerRow || []).map((value) => normalizeCsvHeaderKey(value));
    const columnIndexes = {};
    const missingLabels = [];

    CSV_COLUMNS.forEach((column) => {
      const index = normalizedHeaders.indexOf(normalizeCsvHeaderKey(column.label));
      if (index === -1) {
        missingLabels.push(column.label);
        return;
      }
      columnIndexes[column.key] = index;
    });

    if (missingLabels.length) {
      throw new Error(`CSVヘッダーが不足しています: ${missingLabels.join("、")}`);
    }

    return columnIndexes;
  }

  function readCsvRowData(columns, columnIndexes) {
    const rowData = {};
    CSV_COLUMNS.forEach((column) => {
      const index = columnIndexes[column.key];
      rowData[column.key] = index == null ? "" : String(columns[index] == null ? "" : columns[index]);
    });
    return rowData;
  }

  function createImportedRecord(baseFields, lineNumber) {
    return {
      lineNumber,
      inspectionDate: normalizeImportedInspectionDate(baseFields.inspectionDate),
      vehicleNumber: normalizeCsvLiteral(baseFields.vehicleNumber),
      driverName: normalizeCsvLiteral(baseFields.driverName),
      truckType: normalizeTruckTypeValue(baseFields.truckType),
      reportNote: normalizeText(baseFields.reportNote),
      tires: {},
      spare: {}
    };
  }

  function mergeImportedBaseFields(record, baseFields) {
    const inspectionDate = normalizeImportedInspectionDate(baseFields.inspectionDate);
    const vehicleNumber = normalizeCsvLiteral(baseFields.vehicleNumber);
    const driverName = normalizeCsvLiteral(baseFields.driverName);
    const truckType = normalizeTruckTypeValue(baseFields.truckType);
    const reportNote = normalizeText(baseFields.reportNote);

    if (inspectionDate) record.inspectionDate = inspectionDate;
    if (vehicleNumber) record.vehicleNumber = vehicleNumber;
    if (driverName) record.driverName = driverName;
    if (truckType) record.truckType = truckType;
    if (reportNote) record.reportNote = reportNote;
  }

  function importedBaseFieldsDiffer(record, baseFields) {
    const incomingValues = {
      inspectionDate: normalizeImportedInspectionDate(baseFields.inspectionDate),
      vehicleNumber: normalizeCsvLiteral(baseFields.vehicleNumber),
      driverName: normalizeCsvLiteral(baseFields.driverName),
      truckType: normalizeTruckTypeValue(baseFields.truckType),
      reportNote: normalizeText(baseFields.reportNote)
    };

    return IMPORT_BASE_KEYS.some((key) => {
      const incoming = incomingValues[key];
      return incoming && incoming !== normalizeText(record[key]);
    });
  }

  function appendImportedTireRow(record, rowData, lineNumber) {
    const tireNo = normalizeImportedTireNo(rowData.tireNo);
    const maker = normalizeCsvLiteral(rowData.maker);
    const type = normalizeCsvLiteral(rowData.type);
    const groove = canonicalizeTireValue("groove", normalizeCsvLiteral(rowData.groove));
    const wear = canonicalizeTireValue("wear", normalizeCsvLiteral(rowData.wear));
    const damage = canonicalizeTireValue("damage", normalizeCsvLiteral(rowData.damage));
    const pressure = canonicalizeTireValue("pressure", normalizeCsvLiteral(rowData.pressure));
    const status = normalizeCsvLiteral(rowData.status);
    const hasTireData = [tireNo, maker, type, groove, wear, damage, pressure, status].some(Boolean);

    if (!hasTireData) return;
    if (!tireNo) {
      throw new Error(`${lineNumber}行目: タイヤ番号がありません。`);
    }

    if (tireNo === "スペア") {
      const spare = asPlainObject(record.spare);
      if (maker) spare.maker = maker;
      if (type) spare.type = type;
      if (status) spare.condition = canonicalizeTireValue("groove", status);
      record.spare = spare;
      return;
    }

    const tire = asPlainObject(record.tires[tireNo]);
    if (maker) tire.maker = maker;
    if (type) tire.type = type;
    if (groove) tire.groove = groove;
    if (wear) tire.wear = wear;
    if (damage) tire.damage = damage;
    if (pressure) tire.pressure = pressure;
    if (status) tire.state = status;
    record.tires[tireNo] = tire;
  }

  function finalizeImportedRecord(records, record) {
    if (!record) return;

    const inspectionDate = normalizeText(record.inspectionDate);
    const vehicleNumber = normalizeText(record.vehicleNumber);
    const driverName = normalizeText(record.driverName);
    const truckType = normalizeTruckTypeValue(record.truckType);
    const reportNote = normalizeText(record.reportNote);
    const parsedInspectionDate = parseInspectionDateParts(inspectionDate);
    const tires = asPlainObject(record.tires);
    const spare = asPlainObject(record.spare);
    const hasSpareData = ["maker", "type", "condition"].some((field) => normalizeText(spare[field]));

    if (!parsedInspectionDate) {
      throw new Error(`${record.lineNumber}行目: 点検日が不正です。`);
    }
    if (!vehicleNumber) {
      throw new Error(`${record.lineNumber}行目: 車両番号がありません。`);
    }
    if (!driverName) {
      throw new Error(`${record.lineNumber}行目: 乗務員がありません。`);
    }
    if (!Object.keys(tires).length && !hasSpareData) {
      throw new Error(`${record.lineNumber}行目: タイヤデータがありません。`);
    }

    records.push({
      inspectionDate: `${parsedInspectionDate.year}-${String(parsedInspectionDate.month).padStart(2, "0")}-${String(parsedInspectionDate.day).padStart(2, "0")}`,
      vehicleNumber,
      driverName,
      truckType,
      reportNote,
      tires,
      spare
    });
  }

  function parseImportedCsv(text) {
    const parsedRows = parseCsvRows(text);
    if (!parsedRows.length) {
      throw new Error("CSVファイルが空です。");
    }

    const headerRowIndex = parsedRows.findIndex((row) => row.some((cell) => normalizeText(cell)));
    if (headerRowIndex === -1) {
      throw new Error("CSVファイルが空です。");
    }

    const columnIndexes = getCsvColumnIndexes(parsedRows[headerRowIndex]);
    const records = [];
    let currentRecord = null;

    parsedRows.slice(headerRowIndex + 1).forEach((columns, offset) => {
      const lineNumber = headerRowIndex + offset + 2;
      const rowData = readCsvRowData(columns, columnIndexes);
      const hasAnyValue = Object.values(rowData).some((value) => normalizeText(value));
      if (!hasAnyValue) {
        return;
      }

      const baseFields = {
        inspectionDate: rowData.inspectionDate,
        vehicleNumber: rowData.vehicleNumber,
        driverName: rowData.driverName,
        truckType: rowData.truckType,
        reportNote: rowData.reportNote
      };
      const hasBaseData = IMPORT_BASE_KEYS.some((key) => normalizeText(baseFields[key]));

      if (!currentRecord) {
        if (!hasBaseData) {
          throw new Error(`${lineNumber}行目: 先頭データに基本情報がありません。`);
        }
        currentRecord = createImportedRecord(baseFields, lineNumber);
      } else if (hasBaseData && importedBaseFieldsDiffer(currentRecord, baseFields)) {
        finalizeImportedRecord(records, currentRecord);
        currentRecord = createImportedRecord(baseFields, lineNumber);
      } else if (hasBaseData) {
        mergeImportedBaseFields(currentRecord, baseFields);
      }

      appendImportedTireRow(currentRecord, rowData, lineNumber);
    });

    finalizeImportedRecord(records, currentRecord);

    if (!records.length) {
      throw new Error("CSVに取込可能なデータがありません。");
    }

    return records;
  }

  async function saveImportedRows(rows) {
    if (!rows.length) return;
    if (!state.db) {
      await initFirebase();
    }

    const collectionRef = state.db.collection(state.collection);
    const chunkSize = 400;

    for (let i = 0; i < rows.length; i += chunkSize) {
      const batch = state.db.batch();
      const chunk = rows.slice(i, i + chunkSize);

      chunk.forEach((current) => {
        const timestamp = window.firebase.firestore.FieldValue.serverTimestamp();
        const docRef = collectionRef.doc();
        batch.set(docRef, {
          updatedAt: timestamp,
          state: {
            savedAt: timestamp,
            current
          },
          current
        });
      });

      await batch.commit();
    }
  }

  async function importCsvFile(file) {
    if (!file) return;

    setError("");
    setStatus(`CSV読込中: ${file.name}`);

    const text = await file.text();
    const importedRows = parseImportedCsv(text);
    const confirmed = window.confirm(
      `CSVから ${importedRows.length} 件を取り込みますか？\n重複する同月データは再読込時に上書き整理されます。`
    );
    if (!confirmed) {
      setStatus("CSV取込をキャンセルしました。");
      return;
    }

    setStatus(`CSV取込中: ${importedRows.length}件を保存しています...`);
    await saveImportedRows(importedRows);
    await loadRows();
    setStatus(`CSV取込完了: ${importedRows.length}件。表示件数: ${state.filteredRows.length} / 全${state.rows.length}件`);
  }

  async function downloadCsv() {
    if (!state.filteredRows.length) {
      setError("CSV保存できるデータがありません。");
      return;
    }

    const header = CSV_COLUMNS.map((column) => column.label);
    const rows = state.filteredRows.flatMap((row) => {
      const tireRows = getPrintTireRows(row);
      return tireRows.map((tire, index) => [
        index === 0 ? row.inspectionDate : "",
        index === 0 ? row.vehicleNumber : "",
        index === 0 ? row.driverName : "",
        index === 0 ? row.truckTypeLabel : "",
        index === 0 ? row.reportNote : "",
        tire.tireNo,
        tire.maker,
        tire.type,
        tire.groove,
        tire.wear,
        tire.damage,
        tire.pressure,
        tire.status
      ]);
    });

    const csv = [header, ...rows]
      .map((cols) => cols.map(csvEscape).join(","))
      .join("\r\n");

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
    const filename = `tire_report_${stamp}.csv`;
    const csvBlob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" });

    if (typeof window.showSaveFilePicker !== "function") {
      throw new Error("このブラウザは保存先選択に対応していません。Chrome/Edge を使用してください。");
    }

    try {
      const fileHandle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [
          {
            description: "CSVファイル",
            accept: {
              "text/csv": [".csv"]
            }
          }
        ]
      });
      const writable = await fileHandle.createWritable();
      await writable.write(csvBlob);
      await writable.close();
      setStatus(`CSV保存完了: ${filename}`);
    } catch (error) {
      if (error && error.name === "AbortError") {
        setStatus("CSV保存をキャンセルしました。");
        return;
      }
      throw error;
    }
  }

  async function initFirebase() {
    if (!window.firebase || !window.APP_FIREBASE_CONFIG) {
      throw new Error("Firebase設定の読み込みに失敗しました。firebase/firebase-config.js を確認してください。");
    }

    if (!window.firebase.apps.length) {
      window.firebase.initializeApp(window.APP_FIREBASE_CONFIG);
    }

    const syncOptions = window.APP_FIREBASE_SYNC_OPTIONS || {};
    state.collection = syncOptions.collection || DEFAULT_COLLECTION;

    const auth = window.firebase.auth();
    if (syncOptions.useAnonymousAuth !== false && !auth.currentUser) {
      await auth.signInAnonymously();
    }

    state.db = window.firebase.firestore();
  }

  async function loadRows(options = {}) {
    const skipDuplicateCleanup = options.skipDuplicateCleanup === true;
    setError("");
    setStatus("読み込み中...");

    if (!state.db) {
      await initFirebase();
    }

    const [snap, settingsDocSnaps] = await Promise.all([
      state.db
        .collection(state.collection)
        .orderBy("updatedAt", "desc")
        .get(),
      loadSettingsDocsFromFirestore()
    ]);

    const driverReadingMap = collectDriverReadingMap([...snap.docs, ...settingsDocSnaps]);
    const inspectionDocs = snap.docs.filter((docSnap) => !isSettingsBackupDocId(docSnap.id));
    const extractedRows = inspectionDocs.map((docSnap) => extractRow(docSnap, driverReadingMap));

    if (!skipDuplicateCleanup) {
      try {
        const removedCount = await removeMonthlyDuplicateRows(extractedRows);
        if (removedCount > 0) {
          setStatus(`重複データ ${removedCount} 件を上書き整理しました。再読み込み中...`);
          await loadRows({ skipDuplicateCleanup: true });
          return;
        }
      } catch (error) {
        console.error(error);
        setError(`重複データ整理に失敗しました: ${error.message || error}`);
      }
    }

    state.rows = extractedRows;
    state.fieldOptions = collectFieldOptions(state.rows, collectMasterFieldOptions(settingsDocSnaps), driverReadingMap);
    refreshDateFilterSelectOptions();
    refreshFilterSelectOptions();
    state.expandedDocIds.clear();
    applyFilters();
  }

  function bindEvents() {
    [ui.fromYear, ui.toYear].forEach((node) => {
      node.addEventListener("change", () => {
        refreshDateFilterSelectOptions();
        applyFilters();
      });
    });

    [ui.fromMonth, ui.toMonth].forEach((node) => {
      node.addEventListener("change", applyFilters);
    });

    ui.report.addEventListener("input", applyFilters);

    [ui.vehicle, ui.driver].forEach((node) => {
      node.addEventListener("change", applyFilters);
    });

    syncMonthSelectDisabledStates();

    ui.sortButtons.forEach((button) => {
      button.addEventListener("click", () => {
        toggleSort(button.dataset.sortField);
      });
    });

    ui.reloadBtn.addEventListener("click", async () => {
      try {
        await loadRows();
      } catch (error) {
        console.error(error);
        setError(`再読込に失敗しました: ${error.message || error}`);
      }
    });

    ui.csvBtn.addEventListener("click", async () => {
      setError("");
      try {
        await downloadCsv();
      } catch (error) {
        console.error(error);
        setError(`CSV保存に失敗しました: ${error.message || error}`);
      }
    });

    if (ui.csvImportBtn && ui.csvImportInput) {
      ui.csvImportBtn.addEventListener("click", () => {
        setError("");
        ui.csvImportInput.value = "";
        ui.csvImportInput.click();
      });

      ui.csvImportInput.addEventListener("change", async (event) => {
        const file = event.target.files && event.target.files[0];
        if (!file) return;

        ui.csvImportBtn.disabled = true;
        try {
          await importCsvFile(file);
        } catch (error) {
          console.error(error);
          setError(`CSV取込に失敗しました: ${error.message || error}`);
        } finally {
          ui.csvImportBtn.disabled = false;
          ui.csvImportInput.value = "";
        }
      });
    }

    ui.printBtn.addEventListener("click", () => {
      setError("");
      if (!renderPrintSheets()) {
        setError("印刷できるデータがありません。");
        return;
      }
      window.print();
    });

    ui.editForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      try {
        await saveEditedRow();
      } catch (error) {
        console.error(error);
        const message = `更新に失敗しました: ${error.message || error}`;
        setEditError(message);
        setError(message);
      }
    });

    ui.cancelEditBtn.addEventListener("click", () => {
      closeEditModal();
    });

    ui.editModal.addEventListener("click", (event) => {
      if (event.target === ui.editModal) {
        closeEditModal();
      }
    });

    ui.editTruckType.addEventListener("change", () => {
      if (!state.editingCurrent) return;
      state.editingCurrent.truckType = normalizeText(ui.editTruckType.value);
      renderEditTireInputs();
    });

    ui.editTireBody.addEventListener("change", (event) => {
      const node = event.target.closest("[data-tire-id][data-tire-field]");
      if (node && state.editingCurrent) {
        const tireId = normalizeText(node.dataset.tireId);
        const tireField = normalizeText(node.dataset.tireField);
        if (!tireId || !tireField) return;

        const tires = asPlainObject(state.editingCurrent.tires);
        const tire = asPlainObject(tires[tireId]);
        tire[tireField] = canonicalizeTireValue(tireField, node.value);
        tires[tireId] = tire;
        state.editingCurrent.tires = tires;
        return;
      }

      const spareNode = event.target.closest("[data-spare-field]");
      if (!spareNode || !state.editingCurrent) return;
      const spareField = normalizeText(spareNode.dataset.spareField);
      if (!spareField) return;

      const spare = asPlainObject(state.editingCurrent.spare);
      if (spareField === "condition") {
        spare.condition = canonicalizeTireValue("groove", spareNode.value);
      } else if (spareField === "maker" || spareField === "type") {
        spare[spareField] = normalizeText(spareNode.value);
      }
      state.editingCurrent.spare = spare;
    });

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !ui.editModal.hidden) {
        closeEditModal();
      }
    });

    ui.resultBody.addEventListener("click", (event) => {
      const actionButton = event.target.closest("button[data-action]");
      if (actionButton) {
        const docId = actionButton.dataset.docId;
        if (!docId) return;

        const action = actionButton.dataset.action;
        if (action === "edit") {
          openEditModal(docId);
        } else if (action === "delete") {
          (async () => {
            try {
              await deleteRow(docId);
            } catch (error) {
              console.error(error);
              setError(`削除に失敗しました: ${error.message || error}`);
            }
          })();
        }
        return;
      }

      const tr = event.target.closest("tr.data-row");
      if (!tr) return;
      const docId = tr.dataset.docId;
      if (!docId) return;

      if (state.expandedDocIds.has(docId)) {
        state.expandedDocIds.delete(docId);
      } else {
        state.expandedDocIds.add(docId);
      }
      renderTable();
    });
  }

  async function init() {
    bindEvents();
    try {
      await loadRows();
    } catch (error) {
      console.error(error);
      setStatus("読み込み失敗");
      setError(`データ取得に失敗しました: ${error.message || error}`);
    }
  }

  init();
})();
