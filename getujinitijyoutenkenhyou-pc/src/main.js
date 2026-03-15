import { initializeApp } from "https://www.gstatic.com/firebasejs/12.2.1/firebase-app.js";
import { getAuth, signInAnonymously } from "https://www.gstatic.com/firebasejs/12.2.1/firebase-auth.js";
import {
  collection,
  doc,
  getDoc,
  getDocs,
  getFirestore,
  limit,
  query,
  serverTimestamp,
  setDoc,
  where
} from "https://www.gstatic.com/firebasejs/12.2.1/firebase-firestore.js";

const CHECK_STATES = ["", "レ", "☓", "▲"];
const HOLIDAY_MARK = "休";
const EXCEL_TEMPLATE_FILE_NAME = "月次日常点検 2026.xlsx";
const EXCEL_TEMPLATE_ASSET_FILE_NAME = "monthly-inspection-template.xlsx";
const EXCEL_TEMPLATE_API_PATH = "/api/excel-template";
const EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const EXCEL_TEMPLATE_SHEET_NAME = "日常点検記録表原本";
const EXCEL_MONTH_SHEET_NAMES = {
  1: "日常点検記録表1月",
  2: "日常点検記録表2月",
  3: "日常点検記録表3月"
};
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const EXCEL_SHEET_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const EXCEL_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const EXCEL_CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const EXCEL_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const EXCEL_DRAWING_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const EXCEL_DAY_COLUMNS = Array.from({ length: 31 }, (_, index) => columnNumberToLabel(index + 10));
const EXCEL_WEEKDAY_LABELS = ["日", "月", "火", "水", "木", "金", "土"];
const EXCEL_CHECK_START_ROW = 7;
const EXCEL_BOTTOM_STAMP_ROW = 29;
const EXCEL_IMAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const EXCEL_DRAWING_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const EXCEL_PNG_CONTENT_TYPE = "image/png";
const EXCEL_EMUS_PER_PIXEL = 9525;
const EXCEL_STAMP_IMAGE_SIZES = {
  large: { width: 54, height: 54 },
  small: { width: 22, height: 22 }
};
let jsZipModulePromise = null;

const firebaseConfig = {
  apiKey: "AIzaSyDwthcbGvnkb2Q-K7NjX8SMvVdGZCUUDeA",
  authDomain: "getujinitijyoutenkenhyou.firebaseapp.com",
  projectId: "getujinitijyoutenkenhyou",
  storageBucket: "getujinitijyoutenkenhyou.firebasestorage.app",
  messagingSenderId: "683991833697",
  appId: "1:683991833697:web:a7e0e3b3a85993e7729e20",
  measurementId: "G-TDM7LJ221S"
};

const referenceFirebaseConfig = {
  apiKey: "AIzaSyAlpiGkwyoEW8U8X7HpK4XiqfwW8e_YOdQ",
  authDomain: "getujityretenkenhyou.firebaseapp.com",
  projectId: "getujityretenkenhyou",
  storageBucket: "getujityretenkenhyou.firebasestorage.app",
  messagingSenderId: "818371379903",
  appId: "1:818371379903:web:421a1b390e41a48d2cfc0a",
  measurementId: "G-CPV1MW7ETR"
};

const FIRESTORE_COLLECTION = "monthlyInspectionEntries";
const VEHICLE_SETTINGS_DOC = {
  collection: "monthly_tire_autosave",
  id: "monthly_tire_company_settings_backup_vehicles_slot1"
};
const DRIVER_SETTINGS_DOC = {
  collection: "monthly_tire_autosave",
  id: "monthly_tire_company_settings_backup_drivers_slot1"
};
const sharedSettings = window.SharedAppSettings || null;
const CHECK_FIELD_ORDER = [
  "brake_pedal",
  "brake_fluid",
  "air_pressure",
  "exhaust_sound",
  "parking_brake",
  "tire_pressure",
  "tire_damage",
  "tire_tread",
  "wheel_nut",
  "battery_fluid",
  "coolant",
  "fan_belt",
  "engine_oil",
  "engine_start",
  "engine_response",
  "lights_status",
  "washer_fluid",
  "wiper_status",
  "air_tank_water",
  "documents",
  "emergency_tools",
  "report_changes"
];
const CHECK_FIELD_INDEX = Object.fromEntries(CHECK_FIELD_ORDER.map((fieldKey, index) => [fieldKey, index]));
const CSV_HEADER = ["recordType", "dayOrKey", "fieldKey", "value"];

const GROUPS = [
  {
    category: "1. ブレーキ",
    contents: ["踏みしろ、きき", "液量", "空気圧力の上り具合", "バルブからの排気音", "レバーの引きしろ"]
  },
  {
    category: "2. タイヤ",
    contents: ["空気圧", "亀裂・損傷・異常磨耗", "※溝の深さ", "ホイールナット・ボルト・スペア"]
  },
  { category: "3. バッテリー", contents: ["※液量"] },
  {
    category: "4. エンジン",
    contents: ["※冷却水の量", "※ファンベルトの張り具合・損傷", "※エンジンオイルの量", "※かかり具合、異音", "※低速、加速の状態"]
  },
  { category: "5. 燈火装置", contents: ["点灯・点滅具合、汚れ及び損傷"] },
  { category: "6. ワイパー", contents: ["※液量、噴射状態", "※ワイパー払拭状態"] },
  { category: "7. エアタンク", contents: ["エアタンクに凝水がない"] },
  {
    category: "8. その他",
    contents: ["検査証・保険証・記録簿の備付", "非常用信号具・工具類・停止表示板", "報告事項・変更事項"]
  }
];
const INSPECTION_ITEM_LABELS = GROUPS.flatMap((group) => group.contents);

const monthEl = document.getElementById("month");
const vehicleEl = document.getElementById("vehicle");
const driverEl = document.getElementById("driver");
const monthTextEl = document.getElementById("monthText");
const vehicleTextEl = document.getElementById("vehicleText");
const driverTextEl = document.getElementById("driverText");
const statusEl = document.getElementById("status");
const toolbarEl = document.getElementById("toolbar");
const inspectionTableEl = document.getElementById("inspectionTable");
const datesRowEl = document.getElementById("datesRow");
const daysRowEl = document.getElementById("daysRow");
const bodyEl = document.getElementById("inspectionBody");
const maintenanceFooterRowEl = document.getElementById("maintenanceFooterRow");
const titleHeadEl = document.getElementById("titleHead");
const operationHeadEl = document.getElementById("operationHead");
const maintenanceHeadEl = document.getElementById("maintenanceHead");
const driverHeadEl = document.getElementById("driverHead");
const exportExcelBtnEl = document.getElementById("exportExcelBtn");
const exportCsvBtnEl = document.getElementById("exportCsvBtn");
const importCsvBtnEl = document.getElementById("importCsvBtn");
const helpBtnEl = document.getElementById("helpBtn");
const csvImportInputEl = document.getElementById("csvImportInput");

const state = {
  checks: {},
  operationManager: "",
  maintenanceManager: "",
  maintenanceBottomByDay: {},
  maintenanceRecordsByDay: {},
  holidayDays: [],
  loadedDocId: null,
  vehicleOptions: [],
  driverOptions: [],
  driverStorageMap: {}
};

function columnNumberToLabel(columnNumber) {
  let value = columnNumber;
  let label = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function getReiwaYear(year) {
  return year >= 2019 ? year - 2018 : year;
}

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const referenceApp = initializeApp(referenceFirebaseConfig, "reference-app");
const referenceDb = getFirestore(referenceApp);
const referenceAuth = getAuth(referenceApp);

function normalizeOptionValue(value) {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeWhitespace(value) {
  return normalizeOptionValue(value).replace(/\s+/g, " ");
}

function stripDriverReading(value) {
  return normalizeWhitespace(value).replace(/\s*[（(].*?[）)]\s*/g, "").trim();
}

function normalizeDriverLookupKey(value) {
  return stripDriverReading(value)
    .normalize("NFKC")
    .replace(/\s+/g, "")
    .trim();
}

function rememberDriverStorageValue(value) {
  const rawValue = normalizeOptionValue(value);
  const normalizedValue = stripDriverReading(rawValue);
  if (!normalizedValue) {
    return;
  }

  const existingValue = state.driverStorageMap[normalizedValue];
  if (!existingValue || existingValue === normalizedValue || rawValue !== normalizedValue) {
    state.driverStorageMap[normalizedValue] = rawValue || normalizedValue;
  }
}

function getSelectedDriverStorageValue(value = driverEl.value) {
  const normalizedValue = stripDriverReading(value);
  if (!normalizedValue) {
    return "";
  }
  return state.driverStorageMap[normalizedValue] || normalizedValue;
}

function getDriverIdentity(value = driverEl.value) {
  const storageValue = getSelectedDriverStorageValue(value);
  const displayValue = stripDriverReading(value || storageValue);
  const aliases = [...new Set([storageValue, displayValue].map((item) => normalizeOptionValue(item)).filter(Boolean))];
  return {
    storageValue,
    displayValue,
    aliases,
    normalizedKey: normalizeDriverLookupKey(storageValue || displayValue)
  };
}

function sortOptions(values) {
  return [...values].sort((left, right) => left.localeCompare(right, "ja"));
}

function getDriverSortKey(value) {
  return stripDriverReading(value);
}

function sortDriverOptions(values) {
  return [...values].sort((left, right) => {
    const leftKey = getDriverSortKey(left);
    const rightKey = getDriverSortKey(right);
    return leftKey.localeCompare(rightKey, "ja");
  });
}

function getStringArray(source, fieldName = "values") {
  if (!source || typeof source !== "object" || !Array.isArray(source[fieldName])) {
    return [];
  }
  return source[fieldName].map((value) => normalizeOptionValue(value)).filter(Boolean);
}

function getLocalSharedOptions() {
  if (!sharedSettings || typeof sharedSettings.ensureState !== "function") {
    return {
      vehicles: [],
      rawDrivers: []
    };
  }

  const sharedState = sharedSettings.ensureState();
  return {
    vehicles: Array.isArray(sharedState.vehicles)
      ? sharedState.vehicles.map((value) => normalizeOptionValue(value)).filter(Boolean)
      : [],
    rawDrivers: Array.isArray(sharedState.drivers)
      ? sharedState.drivers.map((value) => normalizeOptionValue(value)).filter(Boolean)
      : []
  };
}

function mergeUniqueOptions() {
  const merged = [];
  const seen = new Set();

  Array.from(arguments).forEach((values) => {
    (values || []).forEach((value) => {
      const normalizedValue = normalizeOptionValue(value);
      if (!normalizedValue || seen.has(normalizedValue)) {
        return;
      }
      seen.add(normalizedValue);
      merged.push(normalizedValue);
    });
  });

  return merged;
}

function buildReferenceDocPath(referenceDoc) {
  return `${referenceDoc.collection}/${referenceDoc.id}`;
}

async function ensureReferenceAuth() {
  if (referenceAuth.currentUser) {
    return referenceAuth.currentUser;
  }
  const credential = await signInAnonymously(referenceAuth);
  return credential.user;
}

function setSelectOptions(selectEl, options, placeholder, selectedValue = "") {
  const normalizeSelectValue = (value) => (
    selectEl === driverEl ? stripDriverReading(value) : normalizeOptionValue(value)
  );
  const normalizedSelectedValue = normalizeSelectValue(selectedValue);
  const uniqueOptions = [...new Set(options.map((option) => normalizeSelectValue(option)).filter(Boolean))];

  if (normalizedSelectedValue && !uniqueOptions.includes(normalizedSelectedValue)) {
    uniqueOptions.unshift(normalizedSelectedValue);
  }

  selectEl.innerHTML = "";

  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = placeholder;
  selectEl.append(placeholderOption);

  uniqueOptions.forEach((optionValue) => {
    const optionEl = document.createElement("option");
    optionEl.value = optionValue;
    optionEl.textContent = optionValue;
    selectEl.append(optionEl);
  });

  selectEl.value = normalizedSelectedValue;
}

function ensureSelectValue(selectEl, value) {
  const normalizedValue = selectEl === driverEl
    ? stripDriverReading(value)
    : normalizeOptionValue(value);
  if (!normalizedValue) {
    selectEl.value = "";
    return;
  }

  const hasOption = Array.from(selectEl.options).some((option) => option.value === normalizedValue);
  if (!hasOption) {
    const optionEl = document.createElement("option");
    optionEl.value = normalizedValue;
    optionEl.textContent = normalizedValue;
    selectEl.append(optionEl);
  }

  selectEl.value = normalizedValue;
}

function escapeCsvValue(value) {
  const text = value == null ? "" : String(value);
  if (!/[",\r\n]/.test(text)) {
    return text;
  }
  return `"${text.replace(/"/g, "\"\"")}"`;
}

function serializeCsv(rows) {
  return rows.map((row) => row.map((value) => escapeCsvValue(value)).join(",")).join("\r\n");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];

    if (inQuotes) {
      if (char === "\"") {
        if (text[index + 1] === "\"") {
          value += "\"";
          index += 1;
        } else {
          inQuotes = false;
        }
      } else {
        value += char;
      }
      continue;
    }

    if (char === "\"") {
      inQuotes = true;
      continue;
    }

    if (char === ",") {
      row.push(value);
      value = "";
      continue;
    }

    if (char === "\r" || char === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      if (char === "\r" && text[index + 1] === "\n") {
        index += 1;
      }
      continue;
    }

    value += char;
  }

  if (inQuotes) {
    throw new Error("CSVの引用符が閉じられていません");
  }

  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }

  return rows.filter((csvRow) => csvRow.some((cell) => cell !== ""));
}

async function loadReferenceOptions() {
  const selectedVehicle = normalizeOptionValue(vehicleEl.value);
  const selectedDriver = normalizeOptionValue(driverEl.value);
  const localOptions = getLocalSharedOptions();
  const localDrivers = localOptions.rawDrivers.map((value) => {
    rememberDriverStorageValue(value);
    return stripDriverReading(value);
  });

  vehicleEl.disabled = true;
  driverEl.disabled = true;

  try {
    await ensureReferenceAuth();

    const [vehicleSnapshot, driverSnapshot] = await Promise.all([
      getDoc(doc(referenceDb, VEHICLE_SETTINGS_DOC.collection, VEHICLE_SETTINGS_DOC.id)),
      getDoc(doc(referenceDb, DRIVER_SETTINGS_DOC.collection, DRIVER_SETTINGS_DOC.id))
    ]);

    const vehicleDocExists = vehicleSnapshot.exists();
    const driverDocExists = driverSnapshot.exists();
    const vehicles = mergeUniqueOptions(
      localOptions.vehicles,
      vehicleDocExists ? getStringArray(vehicleSnapshot.data()) : []
    );
    const rawDrivers = mergeUniqueOptions(
      localOptions.rawDrivers,
      driverDocExists ? getStringArray(driverSnapshot.data()) : []
    );
    rawDrivers.forEach((value) => rememberDriverStorageValue(value));
    const drivers = mergeUniqueOptions(
      localDrivers,
      rawDrivers.map((value) => stripDriverReading(value))
    );

    state.vehicleOptions = sortOptions(vehicles);
    state.driverOptions = sortDriverOptions(drivers);

    setSelectOptions(vehicleEl, state.vehicleOptions, "車番を選択", selectedVehicle);
    setSelectOptions(driverEl, state.driverOptions, "運転者を選択", selectedDriver);

    vehicleEl.disabled = false;
    driverEl.disabled = false;
    syncHeaderInfo();

    if ((!vehicleDocExists || !driverDocExists) && !localOptions.vehicles.length && !localOptions.rawDrivers.length) {
      setStatus(
        `候補設定ドキュメント未検出: project=${referenceFirebaseConfig.projectId} vehicle=${buildReferenceDocPath(VEHICLE_SETTINGS_DOC)} exists=${vehicleDocExists} driver=${buildReferenceDocPath(DRIVER_SETTINGS_DOC)} exists=${driverDocExists}`,
        true
      );
      return;
    }

    if (!state.vehicleOptions.length && !state.driverOptions.length) {
      setStatus(
        `候補設定は取得できましたが values が空です: vehicleCount=${vehicles.length} driverCount=${drivers.length}`,
        true
      );
      return;
    }

    if (!vehicleDocExists || !driverDocExists) {
      setStatus(`ローカル設定を読み込みました: 車番 ${state.vehicleOptions.length}件 / 運転者 ${state.driverOptions.length}件`);
      return;
    }

    setStatus(`候補一覧を読み込みました: 車番 ${state.vehicleOptions.length}件 / 運転者 ${state.driverOptions.length}件`);
  } catch (error) {
    state.vehicleOptions = sortOptions(localOptions.vehicles);
    state.driverOptions = sortDriverOptions(localDrivers);
    setSelectOptions(vehicleEl, state.vehicleOptions, "車番を選択", selectedVehicle);
    setSelectOptions(driverEl, state.driverOptions, "運転者を選択", selectedDriver);
    vehicleEl.disabled = false;
    driverEl.disabled = false;
    syncHeaderInfo();

    if (state.vehicleOptions.length || state.driverOptions.length) {
      setStatus(`ローカル設定を読み込みました。クラウド候補の取得には失敗しました: ${error.message}`, true);
      return;
    }

    setStatus(`候補一覧の取得に失敗しました: ${error.message}`, true);
  }
}

function getSelectedYearMonth() {
  const [yearText, monthText] = monthEl.value.split("-");
  const year = Number(yearText) || 2026;
  const month = Number(monthText) || 1;
  return { year, month };
}

function getDaysInSelectedMonth() {
  const { year, month } = getSelectedYearMonth();
  return new Date(year, month, 0).getDate();
}

function checkKey(itemIndex, day) {
  return `${itemIndex}_${day}`;
}

function getInspectionItemCount() {
  return INSPECTION_ITEM_LABELS.length;
}

function normalizeMaintenanceRecordValue(value) {
  return typeof value === "string" ? value.trim() : "";
}

function sanitizeMaintenanceRecordsByDay(recordsByDay = {}) {
  return Object.fromEntries(
    Object.entries(recordsByDay).filter(([, value]) => normalizeMaintenanceRecordValue(value))
  );
}

function getMaintenanceRecordsByDayFromSource(source = {}) {
  return sanitizeMaintenanceRecordsByDay({
    ...(source.maintenanceRecordsByDay || {}),
    ...(source.maintenanceNotesByDay || {})
  });
}

function getTriangleItemsForDay(day) {
  return INSPECTION_ITEM_LABELS.filter((label, rowIndex) => state.checks[checkKey(rowIndex, day)] === "▲");
}

function syncMaintenanceRecordsByDay() {
  state.maintenanceRecordsByDay = Object.fromEntries(
    Object.entries(state.maintenanceRecordsByDay).filter(([dayText, value]) => {
      const day = Number(dayText);
      return day >= 1 && day <= getDaysInSelectedMonth() && getTriangleItemsForDay(day).length && normalizeMaintenanceRecordValue(value);
    })
  );
}

function promptMaintenanceRecordForDay(day) {
  const triangleItems = getTriangleItemsForDay(day);
  if (!triangleItems.length) {
    return false;
  }

  const currentValue = state.maintenanceRecordsByDay[String(day)] || "";
  const nextValue = window.prompt(
    [
      `${day}日の整備記録を入力してください。`,
      `点検内容: ${triangleItems.join("、")}`,
      "",
      "空欄で OK を押すと整備記録を削除します。"
    ].join("\n"),
    currentValue
  );

  if (nextValue === null) {
    return false;
  }

  const normalizedValue = normalizeMaintenanceRecordValue(nextValue);
  if (normalizedValue) {
    state.maintenanceRecordsByDay[String(day)] = normalizedValue;
  } else {
    delete state.maintenanceRecordsByDay[String(day)];
  }

  return true;
}

function getMaintenanceRecordEntries() {
  syncMaintenanceRecordsByDay();
  const entries = [];
  const daysInMonth = getDaysInSelectedMonth();

  for (let day = 1; day <= daysInMonth; day += 1) {
    const triangleItems = getTriangleItemsForDay(day);
    if (!triangleItems.length) {
      continue;
    }

    const savedRecord = normalizeMaintenanceRecordValue(state.maintenanceRecordsByDay[String(day)] || "");
    const inspectionText = triangleItems.join("、");
    entries.push({
      day,
      text: savedRecord ? `${inspectionText} ${savedRecord}` : inspectionText,
      hasSavedRecord: Boolean(savedRecord)
    });
  }

  return entries;
}

function renderMaintenanceRecordCell() {
  const recordCell = document.getElementById("maintenanceRecordCell");
  if (!recordCell) {
    return;
  }

  recordCell.replaceChildren();
  const entries = getMaintenanceRecordEntries();
  if (!entries.length) {
    return;
  }

  const list = document.createElement("div");
  list.className = "maintenance-record-list";

  entries.forEach(({ day, text, hasSavedRecord }) => {
    const entry = document.createElement("div");
    entry.className = "maintenance-record-entry";
    if (!hasSavedRecord) {
      entry.classList.add("is-fallback");
    }
    entry.textContent = `${day}日 ${text}`;
    entry.title = `${day}日の整備記録を入力・訂正`;
    entry.addEventListener("click", () => {
      const updated = promptMaintenanceRecordForDay(day);
      if (!updated) {
        return;
      }
      renderMaintenanceRecordCell();
      setStatus(`${day}日の整備記録を更新しました。保存すると Firebase に反映されます。`);
    });
    list.append(entry);
  });

  recordCell.append(list);
}

function isHolidayDay(day) {
  return state.holidayDays.includes(day);
}

function setHolidayHeaderState(day, isHoliday) {
  document.querySelectorAll(`[data-day="${day}"]`).forEach((cell) => {
    cell.classList.toggle("is-holiday", isHoliday);
  });
}

function setCheckCellState(cell, value, isHoliday) {
  if (!cell) {
    return;
  }
  cell.textContent = value;
  cell.classList.toggle("is-holiday", isHoliday);
}

function applyHolidayChecks(day) {
  for (let itemIndex = 0; itemIndex < getInspectionItemCount(); itemIndex += 1) {
    const key = checkKey(itemIndex, day);
    state.checks[key] = HOLIDAY_MARK;
    const cell = bodyEl.querySelector(`[data-check-key="${key}"]`);
    setCheckCellState(cell, "", true);
  }
}

function clearHolidayChecks(day) {
  for (let itemIndex = 0; itemIndex < getInspectionItemCount(); itemIndex += 1) {
    const key = checkKey(itemIndex, day);
    delete state.checks[key];
    const cell = bodyEl.querySelector(`[data-check-key="${key}"]`);
    setCheckCellState(cell, "", false);
  }
}

function inferHolidayDaysFromChecks(checks) {
  const daysInMonth = getDaysInSelectedMonth();
  const itemCount = getInspectionItemCount();
  const inferredDays = [];

  for (let day = 1; day <= daysInMonth; day += 1) {
    let hasHolidayMark = false;
    let allHoliday = true;

    for (let itemIndex = 0; itemIndex < itemCount; itemIndex += 1) {
      const value = checks[checkKey(itemIndex, day)];
      if (value !== HOLIDAY_MARK) {
        allHoliday = false;
        break;
      }
      hasHolidayMark = true;
    }

    if (allHoliday && hasHolidayMark) {
      inferredDays.push(day);
    }
  }

  return inferredDays;
}

function mergeHolidayDays(days, checks = state.checks) {
  return [...new Set([...(days || []), ...inferHolidayDaysFromChecks(checks)].map((day) => Number(day)))]
    .filter((day) => Number.isInteger(day) && day >= 1 && day <= getDaysInSelectedMonth())
    .sort((left, right) => left - right);
}

function buildHolidayPayload(days = state.holidayDays, checks = state.checks) {
  const normalizedDays = mergeHolidayDays(days, checks);
  const dayEntries = normalizedDays.map((day) => [String(day), true]);
  return {
    holidayDays: normalizedDays,
    holidays: normalizedDays.map((day) => String(day)),
    holidayFlagsByDay: Object.fromEntries(dayEntries),
    isHolidayByDay: Object.fromEntries(dayEntries)
  };
}

function extractHolidayDays(recordData = {}, checks = {}) {
  const collectedDays = [];

  if (Array.isArray(recordData.holidayDays)) {
    collectedDays.push(...recordData.holidayDays);
  }
  if (Array.isArray(recordData.holidays)) {
    collectedDays.push(...recordData.holidays);
  }

  [recordData.holidayFlagsByDay, recordData.isHolidayByDay].forEach((mapValue) => {
    if (!mapValue || typeof mapValue !== "object") {
      return;
    }
    Object.entries(mapValue).forEach(([dayText, enabled]) => {
      if (enabled) {
        collectedDays.push(dayText);
      }
    });
  });

  return mergeHolidayDays(collectedDays, checks);
}

function syncHolidayChecks() {
  state.holidayDays = mergeHolidayDays(state.holidayDays, state.checks);
  state.holidayDays.forEach((day) => applyHolidayChecks(day));
}

function markHolidayForDay(day) {
  if (isHolidayDay(day)) {
    if (!window.confirm(`${day}日の休日設定を解除しますか？`)) {
      return;
    }

    state.holidayDays = state.holidayDays.filter((holidayDay) => holidayDay !== day);
    clearHolidayChecks(day);
    setHolidayHeaderState(day, false);
    syncMaintenanceRecordsByDay();
    renderMaintenanceRecordCell();
    setStatus(`${day}日の休日設定を解除しました。保存すると反映されます。`);
    return;
  }

  if (!window.confirm(`${day}日を休日にしますか？`)) {
    return;
  }

  state.holidayDays = [...state.holidayDays, day].sort((left, right) => left - right);
  applyHolidayChecks(day);
  setHolidayHeaderState(day, true);
  syncMaintenanceRecordsByDay();
  renderMaintenanceRecordCell();
  setStatus(`${day}日を休日に設定しました。保存すると反映されます。`);
}

function rotateCheck(value) {
  const index = CHECK_STATES.indexOf(value);
  return CHECK_STATES[(index + 1) % CHECK_STATES.length];
}

function escapeXmlText(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildHankoSvgMarkup(name, size = "small") {
  const trimmedName = String(name || "").trim();
  if (!trimmedName) {
    return "";
  }

  const { width, height } = EXCEL_STAMP_IMAGE_SIZES[size] || EXCEL_STAMP_IMAGE_SIZES.small;
  const fontSize = size === "large" ? 22 : 14;
  const strokeWidth = size === "large" ? 4 : 3;
  const inset = strokeWidth;

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <circle cx="${width / 2}" cy="${height / 2}" r="${(Math.min(width, height) / 2) - inset}" fill="rgba(255,255,255,0.18)" stroke="#c61717" stroke-width="${strokeWidth}" />
  <text x="50%" y="54%" text-anchor="middle" dominant-baseline="middle" fill="#c61717" font-size="${fontSize}" font-weight="700" font-family="'Yu Gothic', 'Hiragino Kaku Gothic ProN', sans-serif" letter-spacing="1">${escapeXmlText(trimmedName)}</text>
</svg>`;
}

async function renderSvgToPngArrayBuffer(svgMarkup, width, height) {
  const svgBlob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
  const objectUrl = URL.createObjectURL(svgBlob);

  try {
    const image = await new Promise((resolve, reject) => {
      const nextImage = new Image();
      nextImage.onload = () => resolve(nextImage);
      nextImage.onerror = () => reject(new Error("印画像の描画に失敗しました"));
      nextImage.src = objectUrl;
    });

    const scale = 2;
    const canvas = document.createElement("canvas");
    canvas.width = width * scale;
    canvas.height = height * scale;
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("印画像の描画コンテキストを取得できませんでした");
    }

    context.scale(scale, scale);
    context.drawImage(image, 0, 0, width, height);

    const pngBlob = await new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) {
          resolve(blob);
          return;
        }
        reject(new Error("印画像の PNG 変換に失敗しました"));
      }, EXCEL_PNG_CONTENT_TYPE);
    });

    return await pngBlob.arrayBuffer();
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function createHanko(name, size = "small") {
  if (!name) return "";
  return `<div class="hanko hanko-${size}"><span>${name}</span></div>`;
}

function setStamp(target, value) {
  state[target] = value;
  const idMap = {
    operationManager: "operationManagerSlot",
    maintenanceManager: "maintenanceManagerSlot"
  };
  const slot = document.getElementById(idMap[target]);
  slot.innerHTML = createHanko(value, "large");
}

function setBottomStampByDay(day, value) {
  const dayKey = String(day);
  if (value) {
    state.maintenanceBottomByDay[dayKey] = value;
  } else {
    delete state.maintenanceBottomByDay[dayKey];
  }
  const cell = maintenanceFooterRowEl.querySelector(`[data-bottom-day="${day}"]`);
  if (cell) {
    cell.innerHTML = createHanko(value, "small");
  }
}

function toggleStamp(target, value) {
  const nextValue = state[target] === value ? "" : value;
  setStamp(target, nextValue);
}

function toggleBottomStampByDay(day, value) {
  const dayKey = String(day);
  const nextValue = state.maintenanceBottomByDay[dayKey] === value ? "" : value;
  setBottomStampByDay(day, nextValue);
}

function renderBottomStampRow() {
  maintenanceFooterRowEl.querySelectorAll(".bottom-day-cell").forEach((el) => el.remove());
  const maintenanceRecordFooterCell = document.getElementById("maintenanceRecordFooterCell");
  const daysInMonth = getDaysInSelectedMonth();
  for (let day = 1; day <= daysInMonth; day += 1) {
    const cell = document.createElement("td");
    cell.className = "bottom-day-cell";
    cell.dataset.bottomDay = String(day);
    cell.innerHTML = createHanko(state.maintenanceBottomByDay[String(day)] || "", "small");
    cell.addEventListener("click", () => {
      toggleBottomStampByDay(day, "若本");
    });
    maintenanceFooterRowEl.insertBefore(cell, maintenanceRecordFooterCell);
  }
}

function toWeekdayLabel(year, month, day) {
  const dayOfWeek = new Date(year, month - 1, day).getDay();
  const labels = ["日", "月", "火", "水", "木", "金", "土"];
  return labels[dayOfWeek];
}

function renderDays() {
  datesRowEl.innerHTML = '<th colspan="4" rowspan="2">点検個所</th><th colspan="4" rowspan="2" class="content-head"><div class="content-head-inner"><span class="content-title">点検内容</span><span class="day-mark-stack"><span class="day-mark-cell">日</span><span class="day-mark-cell">曜</span></span></div></th>';
  daysRowEl.innerHTML = "";

  const { year, month } = getSelectedYearMonth();
  const daysInMonth = getDaysInSelectedMonth();
  const managerSpan = 4;
  const titleSpan = Math.max(1, daysInMonth - managerSpan * 2);
  titleHeadEl.colSpan = titleSpan;
  driverHeadEl.colSpan = titleSpan;
  operationHeadEl.colSpan = managerSpan;
  maintenanceHeadEl.colSpan = managerSpan;
  document.getElementById("operationManagerSlot").colSpan = managerSpan;
  document.getElementById("maintenanceManagerSlot").colSpan = managerSpan;

  for (let day = 1; day <= daysInMonth; day += 1) {
    const dateTh = document.createElement("th");
    dateTh.className = "day holiday-trigger";
    dateTh.dataset.day = String(day);
    dateTh.textContent = String(day);
    dateTh.title = `${day}日を休日に設定`;
    dateTh.addEventListener("click", () => {
      markHolidayForDay(day);
    });
    datesRowEl.append(dateTh);

    const dowTh = document.createElement("th");
    dowTh.className = "day holiday-trigger";
    dowTh.dataset.day = String(day);
    dowTh.textContent = toWeekdayLabel(year, month, day);
    dowTh.title = `${day}日を休日に設定`;
    dowTh.addEventListener("click", () => {
      markHolidayForDay(day);
    });
    if (isHolidayDay(day)) {
      dateTh.classList.add("is-holiday");
      dowTh.classList.add("is-holiday");
    }
    daysRowEl.append(dowTh);
  }
}

function syncToolbarWidth() {
  const tableWidth = inspectionTableEl.offsetWidth;
  if (tableWidth > 0) {
    toolbarEl.style.width = `${tableWidth}px`;
  }
}

function printSheet() {
  syncHeaderInfo();
  window.print();
}

function showHelp() {
  window.alert(
    [
      "日付を押すと休日の設定になります。もう一度押すと解除できます。",
      "読込、保存はFirebaseのデータに読込、保存されます。",
      "印刷はA4横で印刷されます。"
    ].join("\n")
  );
}

function renderBody() {
  bodyEl.innerHTML = "";
  const daysInMonth = getDaysInSelectedMonth();
  const itemCount = getInspectionItemCount();
  let rowIndex = 0;
  GROUPS.forEach((group) => {
    group.contents.forEach((line, groupLineIndex) => {
      const tr = document.createElement("tr");

      if (groupLineIndex === 0) {
        const category = document.createElement("td");
        category.className = "category";
        category.colSpan = 4;
        category.rowSpan = group.contents.length;
        category.textContent = group.category;
        tr.append(category);
      }

      const content = document.createElement("td");
      content.className = "content";
      content.colSpan = 4;
      content.textContent = line;
      tr.append(content);

      for (let day = 1; day <= daysInMonth; day += 1) {
        const key = checkKey(rowIndex, day);
        const td = document.createElement("td");
        td.className = "check-cell";
        td.dataset.checkKey = key;
        const isHoliday = isHolidayDay(day);
        const displayValue = isHoliday ? "" : (state.checks[key] || "");
        setCheckCellState(td, displayValue, isHoliday);
        td.addEventListener("click", () => {
          if (isHolidayDay(day)) {
            return;
          }
          const next = rotateCheck(state.checks[key] || "");
          state.checks[key] = next;
          if (!next) {
            delete state.checks[key];
          }
          syncMaintenanceRecordsByDay();
          setCheckCellState(td, next, false);
          renderMaintenanceRecordCell();
        });
        tr.append(td);
      }

      if (rowIndex === 0) {
        const maintenanceRecordCell = document.createElement("td");
        maintenanceRecordCell.id = "maintenanceRecordCell";
        maintenanceRecordCell.className = "maintenance-record-cell";
        maintenanceRecordCell.rowSpan = itemCount;
        tr.append(maintenanceRecordCell);
      }

      bodyEl.append(tr);
      rowIndex += 1;
    });
  });

  renderMaintenanceRecordCell();
}

function syncHeaderInfo() {
  const [, month] = monthEl.value.split("-");
  monthTextEl.textContent = month ? String(Number(month)) : "-";
  vehicleTextEl.textContent = vehicleEl.value.trim() || "-";
  driverTextEl.textContent = stripDriverReading(driverEl.value) || "-";
}

function clearLoadedDocId() {
  state.loadedDocId = null;
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b00020" : "#1e2a35";
}

function buildRecordKey(month, vehicle, driver) {
  return `${month}__${vehicle}__${driver}`;
}

function resetRecordState() {
  state.checks = {};
  state.operationManager = "";
  state.maintenanceManager = "";
  state.maintenanceBottomByDay = {};
  state.maintenanceRecordsByDay = {};
  state.holidayDays = [];
  state.loadedDocId = null;
}

function sanitizeFileNamePart(value, fallback) {
  const normalizedValue = normalizeOptionValue(value);
  if (!normalizedValue) {
    return fallback;
  }
  return normalizedValue.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
}

function downloadBlob(blob, fileName) {
  const downloadUrl = URL.createObjectURL(blob);
  const linkEl = document.createElement("a");

  linkEl.href = downloadUrl;
  linkEl.download = fileName;
  document.body.append(linkEl);
  linkEl.click();
  linkEl.remove();
  URL.revokeObjectURL(downloadUrl);
}

function buildExcelFileName() {
  const month = monthEl.value || "month";
  const vehicle = sanitizeFileNamePart(vehicleEl.value, "vehicle");
  const driver = sanitizeFileNamePart(stripDriverReading(driverEl.value), "driver");
  return `月次日常点検_${month}_${vehicle}_${driver}.xlsx`;
}

async function getJsZipModule() {
  if (!jsZipModulePromise) {
    jsZipModulePromise = import("https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm")
      .then((module) => module.default);
  }
  return jsZipModulePromise;
}

function getExcelTemplateUrlCandidates() {
  return [...new Set([
    new URL(`./assets/${EXCEL_TEMPLATE_ASSET_FILE_NAME}`, import.meta.url).href,
    EXCEL_TEMPLATE_API_PATH,
    new URL(`../${EXCEL_TEMPLATE_FILE_NAME}`, import.meta.url).href
  ])];
}

async function fetchExcelTemplateArrayBuffer() {
  const failures = [];

  for (const templateUrl of getExcelTemplateUrlCandidates()) {
    try {
      const response = await fetch(templateUrl, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      return await response.arrayBuffer();
    } catch (error) {
      failures.push(`${templateUrl}: ${error.message}`);
    }
  }

  throw new Error(`Excelテンプレートを取得できませんでした: ${failures.join(" / ")}`);
}

function parseXmlDocument(xmlText) {
  const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
    throw new Error("ExcelテンプレートのXML解析に失敗しました");
  }
  return xmlDoc;
}

function serializeXmlDocument(xmlDoc) {
  const xmlText = new XMLSerializer().serializeToString(xmlDoc);
  return xmlText.startsWith("<?xml")
    ? xmlText
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xmlText}`;
}

function getZipDirectoryPath(filePath) {
  return filePath.slice(0, filePath.lastIndexOf("/"));
}

function getZipBaseName(filePath) {
  return filePath.slice(filePath.lastIndexOf("/") + 1);
}

function resolveZipPath(fromPath, targetPath) {
  const baseParts = getZipDirectoryPath(fromPath).split("/").filter(Boolean);
  const targetParts = targetPath.split("/").filter(Boolean);

  if (targetPath.startsWith("/")) {
    return targetParts.join("/");
  }

  const resolvedParts = [...baseParts];
  targetParts.forEach((part) => {
    if (!part || part === ".") {
      return;
    }
    if (part === "..") {
      resolvedParts.pop();
      return;
    }
    resolvedParts.push(part);
  });

  return resolvedParts.join("/");
}

function createRelationshipsDocument() {
  return parseXmlDocument(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"></Relationships>`
  );
}

function clearElementChildren(element) {
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }
}

function ensureContentTypeDefault(contentTypesDoc, extension, contentType) {
  const existing = Array.from(contentTypesDoc.getElementsByTagNameNS(EXCEL_CONTENT_TYPES_NAMESPACE, "Default"))
    .find((node) => (node.getAttribute("Extension") || "").toLowerCase() === extension.toLowerCase());
  if (existing) {
    existing.setAttribute("ContentType", contentType);
    return;
  }

  const root = contentTypesDoc.documentElement;
  const defaultNode = contentTypesDoc.createElementNS(EXCEL_CONTENT_TYPES_NAMESPACE, "Default");
  defaultNode.setAttribute("Extension", extension);
  defaultNode.setAttribute("ContentType", contentType);
  root.append(defaultNode);
}

function getNextWorkbookMediaIndex(workbook) {
  const mediaNames = Object.keys(workbook.files).filter((filePath) => /^xl\/media\/stamp-\d+\.png$/.test(filePath));
  if (!mediaNames.length) {
    return 1;
  }

  return Math.max(...mediaNames.map((filePath) => Number(filePath.match(/stamp-(\d+)\.png$/)?.[1] || 0))) + 1;
}

async function createExcelStampMediaStore(workbook, contentTypesDoc) {
  const mediaCache = new Map();
  let nextMediaIndex = getNextWorkbookMediaIndex(workbook);

  ensureContentTypeDefault(contentTypesDoc, "png", EXCEL_PNG_CONTENT_TYPE);

  return {
    async getStampImage(name, size) {
      const trimmedName = String(name || "").trim();
      if (!trimmedName) {
        return null;
      }

      const cacheKey = `${size}:${trimmedName}`;
      if (!mediaCache.has(cacheKey)) {
        mediaCache.set(cacheKey, (async () => {
          const dimensions = EXCEL_STAMP_IMAGE_SIZES[size] || EXCEL_STAMP_IMAGE_SIZES.small;
          const svgMarkup = buildHankoSvgMarkup(trimmedName, size);
          const imageBuffer = await renderSvgToPngArrayBuffer(svgMarkup, dimensions.width, dimensions.height);
          const mediaPath = `xl/media/stamp-${nextMediaIndex}.png`;
          nextMediaIndex += 1;
          workbook.file(mediaPath, imageBuffer);
          return {
            mediaPath,
            ...dimensions
          };
        })());
      }

      return mediaCache.get(cacheKey);
    }
  };
}

function columnLabelToNumber(columnLabel) {
  return columnLabel.split("").reduce((total, char) => (total * 26) + (char.charCodeAt(0) - 64), 0);
}

function parseCellReference(cellRef) {
  const match = /^([A-Z]+)(\d+)$/.exec(cellRef);
  if (!match) {
    throw new Error(`不正なセル参照です: ${cellRef}`);
  }

  const [, columnLabel, rowText] = match;
  return {
    columnLabel,
    columnNumber: columnLabelToNumber(columnLabel),
    rowNumber: Number(rowText)
  };
}

function getWorkbookSheetTarget(workbookDoc, workbookRelsDoc, sheetName) {
  const sheets = Array.from(workbookDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "sheet"));
  const targetSheet = sheets.find((sheet) => sheet.getAttribute("name") === sheetName);
  if (!targetSheet) {
    return null;
  }

  const relationshipId = targetSheet.getAttributeNS(EXCEL_RELATIONSHIP_NAMESPACE, "id") || targetSheet.getAttribute("r:id");
  const relationships = Array.from(
    workbookRelsDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")
  );
  const relationship = relationships.find((item) => item.getAttribute("Id") === relationshipId);
  if (!relationship) {
    return null;
  }

  const targetPath = relationship.getAttribute("Target") || "";
  return {
    sheet: targetSheet,
    index: sheets.indexOf(targetSheet),
    path: targetPath.startsWith("/")
      ? targetPath.replace(/^\//, "")
      : `xl/${targetPath.replace(/^xl\//, "")}`
  };
}

function resolveExcelSheetTarget(workbookDoc, workbookRelsDoc, month) {
  const preferredSheetName = EXCEL_MONTH_SHEET_NAMES[month] || EXCEL_TEMPLATE_SHEET_NAME;
  return getWorkbookSheetTarget(workbookDoc, workbookRelsDoc, preferredSheetName)
    || getWorkbookSheetTarget(workbookDoc, workbookRelsDoc, EXCEL_TEMPLATE_SHEET_NAME);
}

function getInspectionSheetTargets(workbookDoc, workbookRelsDoc) {
  const sheets = Array.from(workbookDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "sheet"));
  return sheets
    .map((sheet) => getWorkbookSheetTarget(workbookDoc, workbookRelsDoc, sheet.getAttribute("name")))
    .filter((target) => target && /^日常点検記録表/.test(target.sheet.getAttribute("name")));
}

function getElementChildren(parentNode, localName) {
  return Array.from(parentNode.childNodes).filter((child) => (
    child.nodeType === Node.ELEMENT_NODE && (!localName || child.localName === localName)
  ));
}

function buildStyleSignature(xfNode) {
  return Array.from(xfNode.attributes)
    .map((attribute) => [attribute.name, attribute.value])
    .filter(([name]) => name !== "fillId" && name !== "applyFill")
    .sort(([leftName], [rightName]) => leftName.localeCompare(rightName))
    .map(([name, value]) => `${name}=${value}`)
    .join("|");
}

function buildStyleFillVariantMap(stylesDoc) {
  const cellXfs = stylesDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "cellXfs")[0];
  if (!cellXfs) {
    return new Map();
  }

  const xfNodes = getElementChildren(cellXfs, "xf");
  const styleIndexByFillAndSignature = new Map();

  xfNodes.forEach((xfNode, styleIndex) => {
    const signature = buildStyleSignature(xfNode);
    const fillId = xfNode.getAttribute("fillId") || "0";
    styleIndexByFillAndSignature.set(`${fillId}::${signature}`, styleIndex);
  });

  const ensureStyleVariant = (xfNode, fillId) => {
    const signature = buildStyleSignature(xfNode);
    const variantKey = `${fillId}::${signature}`;
    const existingStyleIndex = styleIndexByFillAndSignature.get(variantKey);
    if (Number.isInteger(existingStyleIndex)) {
      return existingStyleIndex;
    }

    const clonedXfNode = xfNode.cloneNode(true);
    clonedXfNode.setAttribute("fillId", fillId);
    if (fillId === "0") {
      clonedXfNode.removeAttribute("applyFill");
    } else {
      clonedXfNode.setAttribute("applyFill", "1");
    }

    cellXfs.append(clonedXfNode);
    const styleIndex = getElementChildren(cellXfs, "xf").length - 1;
    cellXfs.setAttribute("count", String(styleIndex + 1));
    styleIndexByFillAndSignature.set(variantKey, styleIndex);
    return styleIndex;
  };

  const variantMap = new Map();
  xfNodes.forEach((xfNode, styleIndex) => {
    variantMap.set(styleIndex, {
      normal: ensureStyleVariant(xfNode, "0"),
      holiday: ensureStyleVariant(xfNode, "2"),
      inactive: ensureStyleVariant(xfNode, "3")
    });
  });

  return variantMap;
}

function getWorksheetRelationshipsPath(worksheetPath) {
  return `${getZipDirectoryPath(worksheetPath)}/_rels/${getZipBaseName(worksheetPath)}.rels`;
}

function getDrawingRelationshipsPath(drawingPath) {
  return `${getZipDirectoryPath(drawingPath)}/_rels/${getZipBaseName(drawingPath)}.rels`;
}

function getWorksheetDrawingPath(workbook, worksheetDoc, worksheetPath) {
  const drawingNode = worksheetDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "drawing")[0];
  if (!drawingNode) {
    return null;
  }

  const relationshipId = drawingNode.getAttributeNS(EXCEL_RELATIONSHIP_NAMESPACE, "id") || drawingNode.getAttribute("r:id");
  if (!relationshipId) {
    return null;
  }

  const relationshipsPath = getWorksheetRelationshipsPath(worksheetPath);
  const relationshipsFile = workbook.file(relationshipsPath);
  if (!relationshipsFile) {
    return null;
  }

  return relationshipsFile.async("string").then((xmlText) => {
    const relationshipsDoc = parseXmlDocument(xmlText);
    const relationship = Array.from(relationshipsDoc.getElementsByTagNameNS(PACKAGE_RELATIONSHIP_NAMESPACE, "Relationship"))
      .find((item) => item.getAttribute("Id") === relationshipId && item.getAttribute("Type") === EXCEL_DRAWING_RELATIONSHIP_TYPE);

    if (!relationship) {
      return null;
    }

    return resolveZipPath(worksheetPath, relationship.getAttribute("Target") || "");
  });
}

function appendTextElement(parentNode, namespace, localName, textContent) {
  const element = parentNode.ownerDocument.createElementNS(namespace, localName);
  element.textContent = String(textContent);
  parentNode.append(element);
  return element;
}

function createDrawingAnchor(worksheetDoc, placement, relationshipId, shapeId) {
  const root = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:oneCellAnchor");
  const fromNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:from");
  const extNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:ext");
  const picNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:pic");
  const nvPicPrNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:nvPicPr");
  const cNvPrNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:cNvPr");
  const cNvPicPrNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:cNvPicPr");
  const picLocksNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:picLocks");
  const blipFillNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:blipFill");
  const blipNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:blip");
  const stretchNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:stretch");
  const fillRectNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:fillRect");
  const spPrNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:spPr");
  const xfrmNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:xfrm");
  const offNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:off");
  const picExtNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:ext");
  const prstGeomNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:prstGeom");
  const avLstNode = worksheetDoc.createElementNS(EXCEL_DRAWING_MAIN_NAMESPACE, "a:avLst");
  const clientDataNode = worksheetDoc.createElementNS(EXCEL_DRAWING_NAMESPACE, "xdr:clientData");
  const { width, height, columnNumber, rowNumber, columnOffset = 0, rowOffset = 0 } = placement;

  appendTextElement(fromNode, EXCEL_DRAWING_NAMESPACE, "xdr:col", columnNumber - 1);
  appendTextElement(fromNode, EXCEL_DRAWING_NAMESPACE, "xdr:colOff", columnOffset);
  appendTextElement(fromNode, EXCEL_DRAWING_NAMESPACE, "xdr:row", rowNumber - 1);
  appendTextElement(fromNode, EXCEL_DRAWING_NAMESPACE, "xdr:rowOff", rowOffset);

  extNode.setAttribute("cx", String(width * EXCEL_EMUS_PER_PIXEL));
  extNode.setAttribute("cy", String(height * EXCEL_EMUS_PER_PIXEL));

  cNvPrNode.setAttribute("id", String(shapeId));
  cNvPrNode.setAttribute("name", `Stamp ${shapeId}`);
  cNvPrNode.setAttribute("descr", placement.name);

  picLocksNode.setAttribute("noChangeAspect", "1");
  cNvPicPrNode.append(picLocksNode);
  nvPicPrNode.append(cNvPrNode, cNvPicPrNode);

  blipNode.setAttributeNS(EXCEL_RELATIONSHIP_NAMESPACE, "r:embed", relationshipId);
  stretchNode.append(fillRectNode);
  blipFillNode.append(blipNode, stretchNode);

  offNode.setAttribute("x", "0");
  offNode.setAttribute("y", "0");
  picExtNode.setAttribute("cx", String(width * EXCEL_EMUS_PER_PIXEL));
  picExtNode.setAttribute("cy", String(height * EXCEL_EMUS_PER_PIXEL));
  xfrmNode.append(offNode, picExtNode);
  prstGeomNode.setAttribute("prst", "rect");
  prstGeomNode.append(avLstNode);
  spPrNode.append(xfrmNode, prstGeomNode);

  picNode.append(nvPicPrNode, blipFillNode, spPrNode);
  root.append(fromNode, extNode, picNode, clientDataNode);
  return root;
}

function getStampPlacements() {
  const placements = [];

  if (state.operationManager) {
    placements.push({
      name: state.operationManager,
      size: "large",
      cellRef: "AI2",
      columnOffset: 18000,
      rowOffset: 12000
    });
  }

  if (state.maintenanceManager) {
    placements.push({
      name: state.maintenanceManager,
      size: "large",
      cellRef: "AL2",
      columnOffset: 18000,
      rowOffset: 12000
    });
  }

  for (let day = 1; day <= EXCEL_DAY_COLUMNS.length; day += 1) {
    const stampName = state.maintenanceBottomByDay[String(day)];
    if (!stampName) {
      continue;
    }

    placements.push({
      name: stampName,
      size: "small",
      cellRef: `${EXCEL_DAY_COLUMNS[day - 1]}${EXCEL_BOTTOM_STAMP_ROW}`,
      columnOffset: 19050,
      rowOffset: 19050
    });
  }

  return placements.map((placement) => ({
    ...placement,
    ...parseCellReference(placement.cellRef),
    ...(EXCEL_STAMP_IMAGE_SIZES[placement.size] || EXCEL_STAMP_IMAGE_SIZES.small)
  }));
}

async function applyStampImagesToWorksheet(workbook, worksheetDoc, worksheetPath, placements, stampMediaStore) {
  if (!placements.length) {
    return;
  }

  const drawingPath = await getWorksheetDrawingPath(workbook, worksheetDoc, worksheetPath);
  if (!drawingPath) {
    throw new Error(`Excel シートの画像領域を特定できませんでした: ${worksheetPath}`);
  }

  const drawingFile = workbook.file(drawingPath);
  if (!drawingFile) {
    throw new Error(`Excel の drawing を開けません: ${drawingPath}`);
  }

  const drawingDoc = parseXmlDocument(await drawingFile.async("string"));
  const drawingRoot = drawingDoc.documentElement;
  clearElementChildren(drawingRoot);

  const drawingRelationshipsPath = getDrawingRelationshipsPath(drawingPath);
  const drawingRelationshipsDoc = createRelationshipsDocument();
  const drawingRelationshipsRoot = drawingRelationshipsDoc.documentElement;

  let nextRelationshipIndex = 1;
  let nextShapeId = 1;
  for (const placement of placements) {
    const image = await stampMediaStore.getStampImage(placement.name, placement.size);
    if (!image) {
      continue;
    }

    const relationshipId = `rId${nextRelationshipIndex}`;
    nextRelationshipIndex += 1;

    const relationshipNode = drawingRelationshipsDoc.createElementNS(PACKAGE_RELATIONSHIP_NAMESPACE, "Relationship");
    relationshipNode.setAttribute("Id", relationshipId);
    relationshipNode.setAttribute("Type", EXCEL_IMAGE_RELATIONSHIP_TYPE);
    relationshipNode.setAttribute("Target", `../media/${getZipBaseName(image.mediaPath)}`);
    drawingRelationshipsRoot.append(relationshipNode);

    drawingRoot.append(createDrawingAnchor(drawingDoc, { ...placement, ...image }, relationshipId, nextShapeId));
    nextShapeId += 1;
  }

  workbook.file(drawingPath, serializeXmlDocument(drawingDoc));
  workbook.file(drawingRelationshipsPath, serializeXmlDocument(drawingRelationshipsDoc));
}

function setWorkbookActiveSheet(workbookDoc, sheetIndex) {
  const workbookView = workbookDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "workbookView")[0];
  if (workbookView) {
    workbookView.setAttribute("activeTab", String(sheetIndex));
  }
}

function setWorksheetSelected(worksheetDoc, isSelected) {
  const sheetView = worksheetDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "sheetView")[0];
  if (!sheetView) {
    return;
  }

  if (isSelected) {
    sheetView.setAttribute("tabSelected", "1");
  } else {
    sheetView.removeAttribute("tabSelected");
  }
}

function getWorksheetRowsMap(worksheetDoc) {
  return new Map(
    Array.from(worksheetDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "row")).map((row) => [Number(row.getAttribute("r")), row])
  );
}

function ensureWorksheetRow(worksheetDoc, rowsMap, rowNumber) {
  let row = rowsMap.get(rowNumber);
  if (row) {
    return row;
  }

  const sheetData = worksheetDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "sheetData")[0];
  row = worksheetDoc.createElementNS(EXCEL_SHEET_NAMESPACE, "row");
  row.setAttribute("r", String(rowNumber));

  const insertBefore = Array.from(sheetData.childNodes)
    .find((candidate) => candidate.nodeType === Node.ELEMENT_NODE && Number(candidate.getAttribute("r")) > rowNumber);
  sheetData.insertBefore(row, insertBefore || null);
  rowsMap.set(rowNumber, row);
  return row;
}

function getRowCells(row) {
  return Array.from(row.childNodes).filter((child) => child.nodeType === Node.ELEMENT_NODE && child.localName === "c");
}

function ensureWorksheetCell(worksheetDoc, rowsMap, cellMap, cellRef) {
  let cell = cellMap.get(cellRef);
  if (cell) {
    return cell;
  }

  const { columnNumber, rowNumber } = parseCellReference(cellRef);
  const row = ensureWorksheetRow(worksheetDoc, rowsMap, rowNumber);
  const rowCells = getRowCells(row);
  const insertBefore = rowCells.find((candidate) => {
    const { columnNumber: candidateColumn } = parseCellReference(candidate.getAttribute("r"));
    return candidateColumn > columnNumber;
  });
  const styleSource = (
    rowCells.find((candidate) => {
      const { columnNumber: candidateColumn } = parseCellReference(candidate.getAttribute("r"));
      return candidateColumn < columnNumber;
    })
    || insertBefore
  );

  cell = worksheetDoc.createElementNS(EXCEL_SHEET_NAMESPACE, "c");
  cell.setAttribute("r", cellRef);
  if (styleSource?.hasAttribute("s")) {
    cell.setAttribute("s", styleSource.getAttribute("s"));
  }

  row.insertBefore(cell, insertBefore || null);
  cellMap.set(cellRef, cell);
  return cell;
}

function buildWorksheetContext(worksheetDoc) {
  const rowsMap = getWorksheetRowsMap(worksheetDoc);
  const cellMap = new Map(
    Array.from(worksheetDoc.getElementsByTagNameNS(EXCEL_SHEET_NAMESPACE, "c")).map((cell) => [cell.getAttribute("r"), cell])
  );
  return { worksheetDoc, rowsMap, cellMap };
}

function setWorksheetCellText(worksheetContext, cellRef, value) {
  const cell = ensureWorksheetCell(
    worksheetContext.worksheetDoc,
    worksheetContext.rowsMap,
    worksheetContext.cellMap,
    cellRef
  );

  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }

  if (!value) {
    cell.removeAttribute("t");
    return;
  }

  const inlineStringEl = cell.ownerDocument.createElementNS(EXCEL_SHEET_NAMESPACE, "is");
  const textEl = cell.ownerDocument.createElementNS(EXCEL_SHEET_NAMESPACE, "t");

  cell.setAttribute("t", "inlineStr");
  textEl.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
  textEl.textContent = value;
  inlineStringEl.append(textEl);
  cell.append(inlineStringEl);
}

function applyHolidayStylesToWorksheet(worksheetDoc, styleFillVariantMap) {
  const { cellMap } = buildWorksheetContext(worksheetDoc);
  const daysInMonth = getDaysInSelectedMonth();

  for (let day = 1; day <= EXCEL_DAY_COLUMNS.length; day += 1) {
    const columnLabel = EXCEL_DAY_COLUMNS[day - 1];
    const isCustomHoliday = day <= daysInMonth && isHolidayDay(day);
    const styleKind = day > daysInMonth
      ? "inactive"
      : (isCustomHoliday ? "holiday" : "normal");

    for (let rowNumber = 5; rowNumber <= 30; rowNumber += 1) {
      const cell = cellMap.get(`${columnLabel}${rowNumber}`);
      if (!cell || !cell.hasAttribute("s")) {
        continue;
      }

      const currentStyleIndex = Number(cell.getAttribute("s"));
      const variants = styleFillVariantMap.get(currentStyleIndex);
      const nextStyleIndex = variants?.[styleKind];

      if (Number.isInteger(nextStyleIndex)) {
        cell.setAttribute("s", String(nextStyleIndex));
      }
    }
  }
}

function populateExcelWorksheet(worksheetDoc, options = {}) {
  const { year, month } = getSelectedYearMonth();
  const daysInMonth = getDaysInSelectedMonth();
  const driverIdentity = getDriverIdentity();
  const worksheetContext = buildWorksheetContext(worksheetDoc);
  const { isTemplateLayout = false } = options;

  setWorksheetCellText(worksheetContext, "A3", `令和${getReiwaYear(year)}年${month}月`);

  if (isTemplateLayout) {
    setWorksheetCellText(worksheetContext, "F3", `車番：${vehicleEl.value.trim()}`);
    setWorksheetCellText(worksheetContext, "J3", `運転者名（点検者）：${driverIdentity.displayValue}`);
  } else {
    setWorksheetCellText(worksheetContext, "H3", vehicleEl.value.trim());
    setWorksheetCellText(worksheetContext, "P3", driverIdentity.displayValue);
  }

  setWorksheetCellText(worksheetContext, "AI2", "");
  setWorksheetCellText(worksheetContext, "AL2", "");

  for (let day = 1; day <= EXCEL_DAY_COLUMNS.length; day += 1) {
    const columnLabel = EXCEL_DAY_COLUMNS[day - 1];
    const dayLabel = day <= daysInMonth ? String(day) : "";
    const weekday = day <= daysInMonth
      ? EXCEL_WEEKDAY_LABELS[new Date(year, month - 1, day).getDay()]
      : "";
    setWorksheetCellText(worksheetContext, `${columnLabel}5`, dayLabel);
    setWorksheetCellText(worksheetContext, `${columnLabel}6`, weekday);
    setWorksheetCellText(worksheetContext, `${columnLabel}${EXCEL_BOTTOM_STAMP_ROW}`, "");

    for (let itemIndex = 0; itemIndex < CHECK_FIELD_ORDER.length; itemIndex += 1) {
      const rawValue = day <= daysInMonth
        ? (state.checks[checkKey(itemIndex, day)] || "")
        : "";
      const displayValue = rawValue === HOLIDAY_MARK ? "" : rawValue;
      setWorksheetCellText(worksheetContext, `${columnLabel}${EXCEL_CHECK_START_ROW + itemIndex}`, displayValue);
    }
  }
}

async function downloadExcel() {
  const vehicle = vehicleEl.value.trim();
  const driverIdentity = getDriverIdentity();
  if (!vehicle || !driverIdentity.storageValue) {
    setStatus("Excel保存前に車番・運転者を選択してください", true);
    return;
  }

  syncHolidayChecks();
  setStatus("Excelファイルを作成しています...");

  const [JSZip, templateBuffer] = await Promise.all([
    getJsZipModule(),
    fetchExcelTemplateArrayBuffer()
  ]);
  const workbook = await JSZip.loadAsync(templateBuffer);
  const contentTypesDoc = parseXmlDocument(await workbook.file("[Content_Types].xml").async("string"));
  const stylesDoc = parseXmlDocument(await workbook.file("xl/styles.xml").async("string"));
  const workbookDoc = parseXmlDocument(await workbook.file("xl/workbook.xml").async("string"));
  const workbookRelsDoc = parseXmlDocument(await workbook.file("xl/_rels/workbook.xml.rels").async("string"));
  const { month } = getSelectedYearMonth();
  const targetSheet = resolveExcelSheetTarget(workbookDoc, workbookRelsDoc, month);
  const templateSheet = getWorkbookSheetTarget(workbookDoc, workbookRelsDoc, EXCEL_TEMPLATE_SHEET_NAME);
  const inspectionSheets = getInspectionSheetTargets(workbookDoc, workbookRelsDoc);
  const styleFillVariantMap = buildStyleFillVariantMap(stylesDoc);
  const stampPlacements = getStampPlacements();
  const stampMediaStore = await createExcelStampMediaStore(workbook, contentTypesDoc);

  if (!targetSheet) {
    throw new Error("Excelテンプレート内の出力先シートが見つかりません");
  }

  if (!EXCEL_MONTH_SHEET_NAMES[month]) {
    targetSheet.sheet.setAttribute("name", `日常点検記録表${month}月`);
  }

  for (const sheetTarget of inspectionSheets) {
    const worksheetFile = workbook.file(sheetTarget.path);
    if (!worksheetFile) {
      throw new Error(`Excelテンプレートのワークシートを開けません: ${sheetTarget.path}`);
    }

    const worksheetDoc = parseXmlDocument(await worksheetFile.async("string"));
    populateExcelWorksheet(worksheetDoc, {
      isTemplateLayout: sheetTarget.path === templateSheet?.path
    });
    applyHolidayStylesToWorksheet(worksheetDoc, styleFillVariantMap);
    await applyStampImagesToWorksheet(workbook, worksheetDoc, sheetTarget.path, stampPlacements, stampMediaStore);
    setWorksheetSelected(worksheetDoc, sheetTarget.path === targetSheet.path);
    workbook.file(sheetTarget.path, serializeXmlDocument(worksheetDoc));
  }

  workbook.file("[Content_Types].xml", serializeXmlDocument(contentTypesDoc));
  workbook.file("xl/styles.xml", serializeXmlDocument(stylesDoc));
  setWorkbookActiveSheet(workbookDoc, targetSheet.index);
  workbook.file("xl/workbook.xml", serializeXmlDocument(workbookDoc));

  const excelBlob = await workbook.generateAsync({
    type: "blob",
    mimeType: EXCEL_MIME_TYPE
  });

  downloadBlob(excelBlob, buildExcelFileName());
  setStatus("Excelファイルを保存しました");
}

function buildCsvRows() {
  syncHolidayChecks();
  syncMaintenanceRecordsByDay();
  const driverIdentity = getDriverIdentity();

  const rows = [
    CSV_HEADER,
    ["meta", "month", "", monthEl.value],
    ["meta", "vehicle", "", vehicleEl.value.trim()],
    ["meta", "driver", "", driverIdentity.storageValue],
    ["meta", "driverDisplay", "", driverIdentity.displayValue],
    ["meta", "operationManager", "", state.operationManager],
    ["meta", "maintenanceManager", "", state.maintenanceManager]
  ];

  state.holidayDays
    .slice()
    .sort((left, right) => left - right)
    .forEach((day) => {
      rows.push(["holiday", String(day), "", "1"]);
    });

  const checksByDay = toFirestoreChecksByDay(state.checks);
  Object.entries(checksByDay)
    .sort(([leftDay], [rightDay]) => Number(leftDay) - Number(rightDay))
    .forEach(([day, valuesByField]) => {
      CHECK_FIELD_ORDER.forEach((fieldKey) => {
        const value = valuesByField[fieldKey];
        if (typeof value === "string" && value) {
          rows.push(["check", day, fieldKey, value]);
        }
      });
    });

  Object.entries(state.maintenanceBottomByDay)
    .sort(([leftDay], [rightDay]) => Number(leftDay) - Number(rightDay))
    .forEach(([day, value]) => {
      if (value) {
        rows.push(["bottomStamp", day, "", value]);
      }
    });

  Object.entries(sanitizeMaintenanceRecordsByDay(state.maintenanceRecordsByDay))
    .sort(([leftDay], [rightDay]) => Number(leftDay) - Number(rightDay))
    .forEach(([day, value]) => {
      const normalizedValue = normalizeMaintenanceRecordValue(value);
      if (normalizedValue) {
        rows.push(["maintenanceRecord", day, "", normalizedValue]);
      }
    });

  return rows;
}

function downloadCsv() {
  const csvText = serializeCsv(buildCsvRows());
  const blob = new Blob(["\uFEFF", csvText], { type: "text/csv;charset=utf-8;" });
  const month = monthEl.value || "month";
  const vehicle = sanitizeFileNamePart(vehicleEl.value, "vehicle");
  const driver = sanitizeFileNamePart(stripDriverReading(driverEl.value), "driver");

  downloadBlob(blob, `${month}_${vehicle}_${driver}_inspection.csv`);

  setStatus("CSVファイルを保存しました");
}

function parseImportedCsv(rows) {
  if (!rows.length) {
    throw new Error("CSVにデータがありません");
  }

  const [headerRow, ...dataRows] = rows;
  const normalizedHeader = headerRow.map((value) => normalizeOptionValue(value));
  if (normalizedHeader.join(",") !== CSV_HEADER.join(",")) {
    throw new Error("このCSVは月次日常点検アプリの形式ではありません");
  }

  const imported = {
    month: monthEl.value,
    vehicle: vehicleEl.value.trim(),
    driver: getDriverIdentity().storageValue,
    checks: {},
    operationManager: "",
    maintenanceManager: "",
    maintenanceBottomByDay: {},
    maintenanceRecordsByDay: {},
    holidayDays: []
  };

  dataRows.forEach((row) => {
    const [recordType = "", dayOrKey = "", fieldKey = "", value = ""] = row;

    if (recordType === "meta") {
      if (dayOrKey === "month" && /^\d{4}-\d{2}$/.test(value)) {
        imported.month = value;
      } else if (dayOrKey === "vehicle") {
        imported.vehicle = value;
      } else if (dayOrKey === "driver") {
        imported.driver = normalizeOptionValue(value);
      } else if (dayOrKey === "driverDisplay" && !imported.driver) {
        imported.driver = normalizeOptionValue(value);
      } else if (dayOrKey === "operationManager") {
        imported.operationManager = value;
      } else if (dayOrKey === "maintenanceManager") {
        imported.maintenanceManager = value;
      }
      return;
    }

    if (recordType === "holiday") {
      const day = Number(dayOrKey);
      if (Number.isInteger(day) && day >= 1) {
        imported.holidayDays.push(day);
      }
      return;
    }

    if (recordType === "bottomStamp") {
      const day = Number(dayOrKey);
      if (Number.isInteger(day) && day >= 1 && value) {
        imported.maintenanceBottomByDay[String(day)] = value;
      }
      return;
    }

    if (recordType === "maintenanceRecord" || recordType === "maintenanceNote") {
      const day = Number(dayOrKey);
      const normalizedValue = normalizeMaintenanceRecordValue(value);
      if (Number.isInteger(day) && day >= 1 && normalizedValue) {
        imported.maintenanceRecordsByDay[String(day)] = normalizedValue;
      }
      return;
    }

    if (recordType === "check") {
      const day = Number(dayOrKey);
      const rowIndex = CHECK_FIELD_INDEX[fieldKey];
      if (Number.isInteger(day) && day >= 1 && Number.isInteger(rowIndex) && value) {
        imported.checks[checkKey(rowIndex, day)] = value;
      }
    }
  });

  imported.holidayDays = mergeHolidayDays(imported.holidayDays, imported.checks);
  return imported;
}

function applyImportedRecord(imported) {
  resetRecordState();

  if (/^\d{4}-\d{2}$/.test(imported.month)) {
    monthEl.value = imported.month;
  }

  rememberDriverStorageValue(imported.driver);
  ensureSelectValue(vehicleEl, imported.vehicle);
  ensureSelectValue(driverEl, imported.driver);

  const daysInMonth = getDaysInSelectedMonth();
  const filteredChecks = {};
  Object.entries(imported.checks).forEach(([key, value]) => {
    const [, dayText] = key.split("_");
    const day = Number(dayText);
    if (day >= 1 && day <= daysInMonth) {
      filteredChecks[key] = value;
    }
  });

  state.checks = filteredChecks;
  state.operationManager = imported.operationManager || "";
  state.maintenanceManager = imported.maintenanceManager || "";
  state.maintenanceBottomByDay = Object.fromEntries(
    Object.entries(imported.maintenanceBottomByDay).filter(([dayText, value]) => {
      const day = Number(dayText);
      return day >= 1 && day <= daysInMonth && Boolean(value);
    })
  );
  state.maintenanceRecordsByDay = Object.fromEntries(
    Object.entries(sanitizeMaintenanceRecordsByDay(imported.maintenanceRecordsByDay)).filter(([dayText]) => {
      const day = Number(dayText);
      return day >= 1 && day <= daysInMonth;
    })
  );
  state.holidayDays = mergeHolidayDays(imported.holidayDays, state.checks).filter((day) => day >= 1 && day <= daysInMonth);

  syncHolidayChecks();
  syncMaintenanceRecordsByDay();
  syncHeaderInfo();
  renderDays();
  renderBody();
  renderBottomStampRow();
  setStamp("operationManager", state.operationManager);
  setStamp("maintenanceManager", state.maintenanceManager);
  syncToolbarWidth();
}

async function importCsvFile(file) {
  const text = (await file.text()).replace(/^\uFEFF/, "");
  const rows = parseCsv(text);
  const imported = parseImportedCsv(rows);
  applyImportedRecord(imported);
}

function getSaveLocationMessage(month, vehicle, driver) {
  return `保存先: Firestore / ${FIRESTORE_COLLECTION}\n一致キー: ${buildRecordKey(month, vehicle, driver)}`;
}

function toFirestoreChecksByDay(checks) {
  const checksByDay = {};
  Object.entries(checks).forEach(([cellKey, value]) => {
    if (!value) return;
    const [rowIndexText, dayText] = cellKey.split("_");
    const rowIndex = Number(rowIndexText);
    const day = String(Number(dayText));
    const fieldKey = CHECK_FIELD_ORDER[rowIndex];
    if (!fieldKey || !day) return;
    if (!checksByDay[day]) {
      checksByDay[day] = {};
    }
    checksByDay[day][fieldKey] = value;
  });
  return checksByDay;
}

function fromFirestoreChecksByDay(checksByDay = {}) {
  const checks = {};
  Object.entries(checksByDay).forEach(([dayText, valuesByField]) => {
    const day = Number(dayText);
    if (!day || typeof valuesByField !== "object" || valuesByField === null) return;
    CHECK_FIELD_ORDER.forEach((fieldKey, rowIndex) => {
      const value = valuesByField[fieldKey];
      if (typeof value === "string" && value) {
        checks[checkKey(rowIndex, day)] = value;
      }
    });
  });
  return checks;
}

async function findRecord(month, vehicle, driver) {
  const recordsRef = collection(db, FIRESTORE_COLLECTION);
  const recordQuery = query(
    recordsRef,
    where("month", "==", month),
    where("vehicle", "==", vehicle),
    where("driver", "==", driver),
    limit(1)
  );
  const snapshot = await getDocs(recordQuery);
  if (snapshot.empty) {
    const fallbackQuery = query(
      recordsRef,
      where("month", "==", month),
      where("vehicle", "==", vehicle),
      limit(50)
    );
    const fallbackSnapshot = await getDocs(fallbackQuery);
    if (fallbackSnapshot.empty) {
      return null;
    }

    const targetDriverKey = normalizeDriverLookupKey(driver);
    const matchedDoc = fallbackSnapshot.docs.find((recordDoc) => {
      const recordData = recordDoc.data();
      const candidateValues = [
        recordData.driver,
        recordData.driverRaw,
        recordData.driverDisplay,
        ...(Array.isArray(recordData.driverAliases) ? recordData.driverAliases : [])
      ];
      return candidateValues.some((candidate) => normalizeDriverLookupKey(candidate || "") === targetDriverKey);
    });

    if (!matchedDoc) {
      return null;
    }

    return {
      id: matchedDoc.id,
      data: matchedDoc.data()
    };
  }
  const recordDoc = snapshot.docs[0];
  return {
    id: recordDoc.id,
    data: recordDoc.data()
  };
}

async function loadRecord() {
  const month = monthEl.value;
  const vehicle = vehicleEl.value.trim();
  const driver = getDriverIdentity().storageValue;
  if (!vehicle || !driver) {
    setStatus("読込前に車番・運転者を入力してください", true);
    return;
  }

  const record = await findRecord(month, vehicle, driver);
  if (!record) {
    resetRecordState();
    setStamp("operationManager", "");
    setStamp("maintenanceManager", "");
    renderDays();
    renderBody();
    renderBottomStampRow();
    setStatus("Firestore に一致データがないため新規入力モードです。");
    return;
  }

  state.loadedDocId = record.id;
  state.checks = fromFirestoreChecksByDay(record.data.checksByDay);
  rememberDriverStorageValue(record.data.driver || "");
  setStamp("operationManager", record.data.operationManager || "");
  setStamp("maintenanceManager", record.data.maintenanceManager || "");
  state.maintenanceBottomByDay = record.data.maintenanceBottomByDay || {};
  state.maintenanceRecordsByDay = getMaintenanceRecordsByDayFromSource(record.data);
  state.holidayDays = extractHolidayDays(record.data, state.checks);
  syncHolidayChecks();
  syncMaintenanceRecordsByDay();
  renderDays();
  renderBody();
  renderBottomStampRow();
  setStatus("読込完了");
}

async function saveRecord() {
  const month = monthEl.value;
  const vehicle = vehicleEl.value.trim();
  const driverIdentity = getDriverIdentity();
  const driver = driverIdentity.storageValue;
  if (!vehicle || !driver) {
    setStatus("保存前に車番・運転者を入力してください", true);
    return;
  }

  const saveLocationMessage = getSaveLocationMessage(month, vehicle, driver);
  const accepted = window.confirm(`保存先を確認してください。\n\n${saveLocationMessage}\n\nこの場所に保存しますか？`);
  if (!accepted) {
    setStatus("保存をキャンセルしました");
    return;
  }

  const existingRecord = await findRecord(month, vehicle, driver);
  syncHolidayChecks();
  syncMaintenanceRecordsByDay();
  const holidayPayload = buildHolidayPayload(state.holidayDays, state.checks);
  const rawDocId = buildRecordKey(month, vehicle, driverIdentity.storageValue);
  const displayDocId = buildRecordKey(month, vehicle, driverIdentity.displayValue);
  const docIds = [...new Set([existingRecord?.id, state.loadedDocId, rawDocId, displayDocId].filter(Boolean))];
  const basePayload = {
    month,
    vehicle,
    driver,
    driverRaw: driverIdentity.storageValue,
    driverDisplay: driverIdentity.displayValue,
    driverAliases: driverIdentity.aliases,
    driverNormalized: driverIdentity.normalizedKey,
    checksByDay: toFirestoreChecksByDay(state.checks),
    operationManager: state.operationManager,
    maintenanceManager: state.maintenanceManager,
    maintenanceBottomByDay: state.maintenanceBottomByDay,
    maintenanceRecordsByDay: sanitizeMaintenanceRecordsByDay(state.maintenanceRecordsByDay),
    maintenanceNotesByDay: sanitizeMaintenanceRecordsByDay(state.maintenanceRecordsByDay),
    ...holidayPayload,
    updatedAt: serverTimestamp()
  };

  await Promise.all(docIds.map((docId) => {
    const payload = {
      ...basePayload,
      driver: docId === displayDocId ? driverIdentity.displayValue : driverIdentity.storageValue
    };
    return setDoc(doc(db, FIRESTORE_COLLECTION, docId), payload);
  }));
  state.loadedDocId = rawDocId;
  setStatus("保存完了");
}

monthEl.addEventListener("change", () => {
  clearLoadedDocId();
  syncHeaderInfo();
  renderDays();
  renderBody();
  renderBottomStampRow();
  syncToolbarWidth();
});
vehicleEl.addEventListener("change", () => {
  clearLoadedDocId();
  syncHeaderInfo();
});
driverEl.addEventListener("change", () => {
  clearLoadedDocId();
  syncHeaderInfo();
});
window.addEventListener("resize", syncToolbarWidth);

document.getElementById("loadBtn").addEventListener("click", () => {
  loadRecord().catch((err) => setStatus(`読込失敗: ${err.message}`, true));
});

document.getElementById("printBtn").addEventListener("click", () => {
  printSheet();
});

helpBtnEl.addEventListener("click", () => {
  showHelp();
});

document.getElementById("saveBtn").addEventListener("click", () => {
  saveRecord().catch((err) => setStatus(`保存失敗: ${err.message}`, true));
});

document.getElementById("operationManagerSlot").addEventListener("click", () => toggleStamp("operationManager", "岸田"));
document.getElementById("maintenanceManagerSlot").addEventListener("click", () => toggleStamp("maintenanceManager", "若本"));

exportExcelBtnEl.addEventListener("click", () => {
  downloadExcel().catch((error) => {
    setStatus(`Excel保存失敗: ${error.message}`, true);
  });
});

exportCsvBtnEl.addEventListener("click", () => {
  try {
    downloadCsv();
  } catch (error) {
    setStatus(`CSV保存失敗: ${error.message}`, true);
  }
});

importCsvBtnEl.addEventListener("click", () => {
  csvImportInputEl.value = "";
  csvImportInputEl.click();
});

csvImportInputEl.addEventListener("change", (event) => {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  importCsvFile(file).catch((error) => {
    setStatus(`CSV読込失敗: ${error.message}`, true);
  });
});

syncHeaderInfo();
renderDays();
renderBody();
renderBottomStampRow();
syncToolbarWidth();
loadReferenceOptions().catch((err) => setStatus(`候補一覧の取得に失敗しました: ${err.message}`, true));
