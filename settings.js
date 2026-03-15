(function () {
  "use strict";

  const sharedSettings = window.SharedAppSettings;
  const VEHICLE_BACKUP = Object.freeze({
    kind: "vehicles",
    docId: "monthly_tire_company_settings_backup_vehicles_slot1",
    label: "車両番号"
  });
  const DRIVER_BACKUP = Object.freeze({
    kind: "drivers",
    docId: "monthly_tire_company_settings_backup_drivers_slot1",
    label: "運転者名"
  });
  const DEFAULT_COLLECTION = "monthly_tire_autosave";

  const elements = {
    vehicleInput: document.getElementById("vehicleInput"),
    addVehicleButton: document.getElementById("addVehicleButton"),
    saveVehiclesButton: document.getElementById("saveVehiclesButton"),
    restoreVehiclesButton: document.getElementById("restoreVehiclesButton"),
    deleteVehiclesButton: document.getElementById("deleteVehiclesButton"),
    vehicleLocalStatus: document.getElementById("vehicleLocalStatus"),
    vehicleBackupStatus: document.getElementById("vehicleBackupStatus"),
    driverNameInput: document.getElementById("driverNameInput"),
    driverReadingInput: document.getElementById("driverReadingInput"),
    addDriverButton: document.getElementById("addDriverButton"),
    saveDriversButton: document.getElementById("saveDriversButton"),
    restoreDriversButton: document.getElementById("restoreDriversButton"),
    deleteDriversButton: document.getElementById("deleteDriversButton"),
    vehicleList: document.getElementById("vehicleList"),
    driverList: document.getElementById("driverList"),
    driverLocalStatus: document.getElementById("driverLocalStatus"),
    driverBackupStatus: document.getElementById("driverBackupStatus"),
    globalStatus: document.getElementById("globalStatus")
  };

  const state = {
    shared: sharedSettings.ensureState(),
    backupMeta: {
      vehicles: null,
      drivers: null
    },
    cloudReady: false,
    directoryEnabled: false,
    db: null,
    directoryDb: null,
    directoryError: "",
    working: false
  };

  bindEvents();
  render();
  void initializeCloud();

  function bindEvents() {
    elements.addVehicleButton.addEventListener("click", addVehicle);
    elements.vehicleInput.addEventListener("keydown", function (event) {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      addVehicle();
    });

    elements.addDriverButton.addEventListener("click", addDriver);
    elements.driverNameInput.addEventListener("keydown", function (event) {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      addDriver();
    });
    elements.driverReadingInput.addEventListener("keydown", function (event) {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      addDriver();
    });

    elements.saveVehiclesButton.addEventListener("click", function () {
      void saveBackup(VEHICLE_BACKUP);
    });
    elements.restoreVehiclesButton.addEventListener("click", function () {
      void restoreBackup(VEHICLE_BACKUP);
    });
    elements.deleteVehiclesButton.addEventListener("click", clearAllVehicles);

    elements.saveDriversButton.addEventListener("click", function () {
      void saveBackup(DRIVER_BACKUP);
    });
    elements.restoreDriversButton.addEventListener("click", function () {
      void restoreBackup(DRIVER_BACKUP);
    });
    elements.deleteDriversButton.addEventListener("click", clearAllDrivers);
  }

  function refreshSharedState() {
    state.shared = sharedSettings.ensureState();
  }

  function setGlobalStatus(message, isError) {
    elements.globalStatus.textContent = message || "";
    elements.globalStatus.style.color = isError ? "#b00020" : "";
  }

  function render() {
    refreshSharedState();

    elements.vehicleLocalStatus.textContent = "ローカル登録数: " + state.shared.vehicles.length + "件";
    elements.driverLocalStatus.textContent = "ローカル登録数: " + state.shared.drivers.length + "件";

    renderValueList(elements.vehicleList, state.shared.vehicles, removeVehicle);
    renderValueList(elements.driverList, state.shared.drivers, removeDriver);
    renderBackupStatus(VEHICLE_BACKUP, elements.vehicleBackupStatus);
    renderBackupStatus(DRIVER_BACKUP, elements.driverBackupStatus);
    renderButtons();
  }

  function renderValueList(container, values, onDelete) {
    container.innerHTML = "";

    if (!values.length) {
      const empty = document.createElement("div");
      empty.className = "empty-list";
      empty.textContent = "まだ登録されていません。";
      container.appendChild(empty);
      return;
    }

    values.forEach(function (value) {
      const item = document.createElement("div");
      item.className = "value-item";

      const label = document.createElement("span");
      label.className = "value-item-label";
      label.textContent = value;

      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "mini-button danger value-delete-button";
      deleteButton.textContent = "削除";
      deleteButton.addEventListener("click", function () {
        onDelete(value);
      });

      item.appendChild(label);
      item.appendChild(deleteButton);
      container.appendChild(item);
    });
  }

  function renderBackupStatus(definition, node) {
    if (!state.cloudReady) {
      node.textContent = "バックアップ: 利用できません";
      return;
    }

    const meta = state.backupMeta[definition.kind];
    if (!meta) {
      node.textContent = "バックアップ: 未保存";
      return;
    }

    const updatedAt = meta.serverUpdatedAt || meta.clientUpdatedAt || "";
    const updatedLabel = updatedAt ? formatDateTime(updatedAt) : "日時不明";
    node.textContent = "バックアップ: " + updatedLabel + " / " + Number(meta.valueCount || 0) + "件";
  }

  function renderButtons() {
    const disabledCloud = state.working || !state.cloudReady;
    elements.saveVehiclesButton.disabled = disabledCloud || !state.shared.vehicles.length;
    elements.saveDriversButton.disabled = disabledCloud || !state.shared.drivers.length;
    elements.restoreVehiclesButton.disabled = disabledCloud || !state.backupMeta.vehicles;
    elements.restoreDriversButton.disabled = disabledCloud || !state.backupMeta.drivers;
    elements.deleteVehiclesButton.disabled = state.working || !state.shared.vehicles.length;
    elements.deleteDriversButton.disabled = state.working || !state.shared.drivers.length;
  }

  function addVehicle() {
    const value = sharedSettings.normalizeText(elements.vehicleInput.value);
    if (!value) {
      setGlobalStatus("車両番号を入力してください。", true);
      return;
    }

    if (state.shared.vehicles.includes(value)) {
      setGlobalStatus("同じ車両番号はすでに登録されています。", true);
      return;
    }

    sharedSettings.addVehicle(value);
    elements.vehicleInput.value = "";
    render();
    setGlobalStatus("車両番号を追加しました。");
  }

  function addDriver() {
    const name = sharedSettings.normalizeText(elements.driverNameInput.value);
    const reading = sharedSettings.normalizeText(elements.driverReadingInput.value);
    if (!name) {
      setGlobalStatus("運転者名を入力してください。", true);
      return;
    }

    sharedSettings.addDriver(name, reading);
    elements.driverNameInput.value = "";
    elements.driverReadingInput.value = "";
    render();
    setGlobalStatus("運転者名を追加しました。");
  }

  function removeVehicle(value) {
    if (!window.confirm("車両番号「" + value + "」を削除しますか？")) {
      return;
    }

    sharedSettings.saveVehicles(state.shared.vehicles.filter(function (entry) {
      return entry !== value;
    }));
    render();
    setGlobalStatus("車両番号を削除しました。");
  }

  function removeDriver(value) {
    if (!window.confirm("運転者名「" + value + "」を削除しますか？")) {
      return;
    }

    sharedSettings.saveDrivers(state.shared.drivers.filter(function (entry) {
      return entry !== value;
    }));
    render();
    setGlobalStatus("運転者名を削除しました。");
  }

  function clearAllVehicles() {
    if (!state.shared.vehicles.length) {
      return;
    }

    if (!window.confirm("この端末内の車両番号を全件削除しますか？\nFirebase バックアップは削除しません。")) {
      return;
    }

    sharedSettings.saveVehicles([]);
    render();
    setGlobalStatus("この端末内の車両番号を全件削除しました。Firebase バックアップは残っています。");
  }

  function clearAllDrivers() {
    if (!state.shared.drivers.length) {
      return;
    }

    if (!window.confirm("この端末内の運転者名を全件削除しますか？\nFirebase バックアップは削除しません。")) {
      return;
    }

    sharedSettings.saveDrivers([]);
    render();
    setGlobalStatus("この端末内の運転者名を全件削除しました。Firebase バックアップは残っています。");
  }

  async function initializeCloud() {
    state.directoryEnabled = hasDirectorySyncTarget();

    try {
      state.db = await ensurePrimaryDb();
      state.cloudReady = true;
      state.directoryDb = null;

      if (state.directoryEnabled) {
        try {
          state.directoryDb = await ensureDirectoryDb();
          state.directoryError = "";
        } catch (error) {
          state.directoryDb = null;
          state.directoryError = formatErrorReason(error);
          console.warn("Failed to initialize employee directory cloud:", error);
        }
      }

      await refreshBackups();
    } catch (error) {
      state.cloudReady = false;
      state.db = null;
      state.directoryDb = null;
      state.directoryError = "";
      render();
      setGlobalStatus("Firebase に接続できないため、バックアップ機能は使えません。", true);
      console.warn("Failed to initialize settings cloud:", error);
    }
  }

  async function ensurePrimaryDb() {
    if (!window.firebase || !window.APP_FIREBASE_CONFIG) {
      throw new Error("firebase_config_missing");
    }

    return ensureDb(window.APP_FIREBASE_CONFIG, window.APP_FIREBASE_SYNC_OPTIONS || {}, null);
  }

  async function ensureDirectoryDb() {
    if (!hasDirectorySyncTarget()) {
      return null;
    }

    const syncOptions = window.APP_FIREBASE_DIRECTORY_SYNC_OPTIONS || {};
    return ensureDb(
      window.APP_FIREBASE_DIRECTORY_CONFIG,
      syncOptions,
      syncOptions.appName || "employee-directory"
    );
  }

  async function ensureDb(config, syncOptions, appName) {
    const app = getOrCreateFirebaseApp(config, appName);
    const auth = app.auth();
    if (syncOptions.useAnonymousAuth !== false && !auth.currentUser) {
      await auth.signInAnonymously();
    }

    return app.firestore();
  }

  function getCollectionName() {
    const syncOptions = window.APP_FIREBASE_SYNC_OPTIONS || {};
    return syncOptions.collection || DEFAULT_COLLECTION;
  }

  function getDirectoryCollectionName() {
    const syncOptions = window.APP_FIREBASE_DIRECTORY_SYNC_OPTIONS || {};
    return syncOptions.collection || DEFAULT_COLLECTION;
  }

  function hasDirectorySyncTarget() {
    const syncOptions = window.APP_FIREBASE_DIRECTORY_SYNC_OPTIONS || {};
    return Boolean(syncOptions.enabled && window.APP_FIREBASE_DIRECTORY_CONFIG);
  }

  function getOrCreateFirebaseApp(config, appName) {
    if (!appName) {
      if (!window.firebase.apps.length) {
        return window.firebase.initializeApp(config);
      }
      return window.firebase.app();
    }

    const existingApp = window.firebase.apps.find(function (app) {
      return app.name === appName;
    });
    if (existingApp) {
      return existingApp;
    }

    return window.firebase.initializeApp(config, appName);
  }

  function getDocId(definition, syncOptions) {
    const docIds = syncOptions.docIds || {};
    return docIds[definition.kind] || definition.docId;
  }

  function getBackupDocRef(definition) {
    return state.db.collection(getCollectionName()).doc(definition.docId);
  }

  function getDirectoryDocRef(definition) {
    if (!state.directoryDb) {
      return null;
    }

    const syncOptions = window.APP_FIREBASE_DIRECTORY_SYNC_OPTIONS || {};
    return state.directoryDb
      .collection(getDirectoryCollectionName())
      .doc(getDocId(definition, syncOptions));
  }

  async function ensureDirectoryReady() {
    if (!state.directoryEnabled) {
      return null;
    }

    if (state.directoryDb) {
      return state.directoryDb;
    }

    try {
      state.directoryDb = await ensureDirectoryDb();
      state.directoryError = "";
      return state.directoryDb;
    } catch (error) {
      state.directoryDb = null;
      state.directoryError = formatErrorReason(error);
      throw error;
    }
  }

  async function refreshBackups() {
    if (!state.cloudReady) {
      render();
      return;
    }

    const snapshots = await Promise.all([
      getBackupDocRef(VEHICLE_BACKUP).get(),
      getBackupDocRef(DRIVER_BACKUP).get()
    ]);

    state.backupMeta.vehicles = snapshots[0].exists ? buildBackupMeta(snapshots[0]) : null;
    state.backupMeta.drivers = snapshots[1].exists ? buildBackupMeta(snapshots[1]) : null;
    render();
  }

  function buildBackupMeta(snapshot) {
    const data = snapshot.data() || {};
    return {
      valueCount: Array.isArray(data.values) ? data.values.length : Number(data.valueCount || 0),
      clientUpdatedAt: data.clientUpdatedAt || "",
      serverUpdatedAt: formatFirestoreDate(data.updatedAt)
    };
  }

  async function saveBackup(definition) {
    const values = definition.kind === "vehicles" ? state.shared.vehicles : state.shared.drivers;
    if (!values.length) {
      setGlobalStatus(definition.label + "の登録データがありません。", true);
      return;
    }

    state.working = true;
    render();

    try {
      await getBackupDocRef(definition).set(buildBackupPayload(definition, values, "integrated-settings"));

      let directorySaveFailed = false;
      let directoryError = "";
      if (state.directoryEnabled) {
        try {
          await ensureDirectoryReady();
          await getDirectoryDocRef(definition).set(
            buildBackupPayload(definition, values, "integrated-settings-directory")
          );
          state.directoryError = "";
        } catch (error) {
          directorySaveFailed = true;
          directoryError = formatErrorReason(error);
          state.directoryError = directoryError;
          console.warn("Failed to save employee directory backup:", error);
        }
      }

      await refreshBackups();
      if (directorySaveFailed) {
        const detail = directoryError ? " (" + directoryError + ")" : "";
        setGlobalStatus(definition.label + "のバックアップを保存しました。社員名簿用 Firebase への追加保存は失敗しました。" + detail, true);
      } else if (state.directoryEnabled) {
        setGlobalStatus(definition.label + "のバックアップを保存しました。社員名簿用 Firebase にも保存しました。");
      } else {
        setGlobalStatus(definition.label + "のバックアップを保存しました。");
      }
    } catch (error) {
      setGlobalStatus(definition.label + "のバックアップ保存に失敗しました。", true);
      console.warn("Failed to save backup:", error);
    } finally {
      state.working = false;
      render();
    }
  }

  function buildBackupPayload(definition, values, source) {
    return {
      kind: definition.kind,
      slot: 1,
      values: values,
      valueCount: values.length,
      clientUpdatedAt: new Date().toISOString(),
      updatedAt: window.firebase.firestore.FieldValue.serverTimestamp(),
      source: source
    };
  }

  function formatErrorReason(error) {
    if (!error) {
      return "";
    }

    if (error.code) {
      return String(error.code);
    }

    if (error.message) {
      return String(error.message);
    }

    return "";
  }

  async function restoreBackup(definition) {
    state.working = true;
    render();

    try {
      const snapshot = await getBackupDocRef(definition).get();
      if (!snapshot.exists) {
        setGlobalStatus(definition.label + "のバックアップはありません。", true);
        return;
      }

      const data = snapshot.data() || {};
      const values = Array.isArray(data.values) ? data.values : [];

      if (definition.kind === "vehicles") {
        sharedSettings.saveVehicles(values);
      } else {
        sharedSettings.saveDrivers(values);
      }

      await refreshBackups();
      setGlobalStatus(definition.label + "をバックアップから復元しました。");
    } catch (error) {
      setGlobalStatus(definition.label + "の復元に失敗しました。", true);
      console.warn("Failed to restore backup:", error);
    } finally {
      state.working = false;
      render();
    }
  }

  function formatFirestoreDate(value) {
    if (!value) {
      return "";
    }
    if (typeof value.toDate === "function") {
      try {
        return value.toDate().toISOString();
      } catch {
        return "";
      }
    }
    return String(value);
  }

  function formatDateTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "日時不明";
    }

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    return year + "-" + month + "-" + day + " " + hours + ":" + minutes;
  }
})();
