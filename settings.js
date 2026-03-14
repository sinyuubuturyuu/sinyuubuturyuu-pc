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
    label: "乗務員"
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
    db: null,
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
    elements.deleteVehiclesButton.addEventListener("click", function () {
      void deleteBackup(VEHICLE_BACKUP);
    });

    elements.saveDriversButton.addEventListener("click", function () {
      void saveBackup(DRIVER_BACKUP);
    });
    elements.restoreDriversButton.addEventListener("click", function () {
      void restoreBackup(DRIVER_BACKUP);
    });
    elements.deleteDriversButton.addEventListener("click", function () {
      void deleteBackup(DRIVER_BACKUP);
    });
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

    renderValueList(elements.vehicleList, state.shared.vehicles);
    renderValueList(elements.driverList, state.shared.drivers);
    renderBackupStatus(VEHICLE_BACKUP, elements.vehicleBackupStatus);
    renderBackupStatus(DRIVER_BACKUP, elements.driverBackupStatus);
    renderButtons();
  }

  function renderValueList(container, values) {
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
      item.textContent = value;
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
    elements.deleteVehiclesButton.disabled = disabledCloud || !state.backupMeta.vehicles;
    elements.deleteDriversButton.disabled = disabledCloud || !state.backupMeta.drivers;
  }

  function addVehicle() {
    const value = sharedSettings.normalizeText(elements.vehicleInput.value);
    if (!value) {
      setGlobalStatus("車両番号を入力してください。", true);
      return;
    }

    if (state.shared.vehicles.includes(value)) {
      setGlobalStatus("同じ車両番号は登録済みです。", true);
      return;
    }

    sharedSettings.addVehicle(value);
    elements.vehicleInput.value = "";
    render();
    setGlobalStatus("車両番号を登録しました。");
  }

  function addDriver() {
    const name = sharedSettings.normalizeText(elements.driverNameInput.value);
    const reading = sharedSettings.normalizeText(elements.driverReadingInput.value);
    if (!name) {
      setGlobalStatus("乗務員名を入力してください。", true);
      return;
    }

    sharedSettings.addDriver(name, reading);
    elements.driverNameInput.value = "";
    elements.driverReadingInput.value = "";
    render();
    setGlobalStatus("乗務員を登録しました。");
  }

  async function initializeCloud() {
    try {
      state.db = await ensureDb();
      state.cloudReady = true;
      await refreshBackups();
    } catch (error) {
      state.cloudReady = false;
      state.db = null;
      render();
      setGlobalStatus("Firebase に接続できないため、バックアップ機能は使えません。", true);
      console.warn("Failed to initialize settings cloud:", error);
    }
  }

  async function ensureDb() {
    if (!window.firebase || !window.APP_FIREBASE_CONFIG) {
      throw new Error("firebase_config_missing");
    }

    if (!window.firebase.apps.length) {
      window.firebase.initializeApp(window.APP_FIREBASE_CONFIG);
    }

    const auth = window.firebase.auth();
    const syncOptions = window.APP_FIREBASE_SYNC_OPTIONS || {};
    if (syncOptions.useAnonymousAuth !== false && !auth.currentUser) {
      await auth.signInAnonymously();
    }

    return window.firebase.firestore();
  }

  function getCollectionName() {
    const syncOptions = window.APP_FIREBASE_SYNC_OPTIONS || {};
    return syncOptions.collection || DEFAULT_COLLECTION;
  }

  function getBackupDocRef(definition) {
    return state.db.collection(getCollectionName()).doc(definition.docId);
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
      await getBackupDocRef(definition).set({
        kind: definition.kind,
        slot: 1,
        values: values,
        valueCount: values.length,
        clientUpdatedAt: new Date().toISOString(),
        updatedAt: window.firebase.firestore.FieldValue.serverTimestamp(),
        source: "integrated-settings"
      });
      await refreshBackups();
      setGlobalStatus(definition.label + "のバックアップを保存しました。");
    } catch (error) {
      setGlobalStatus(definition.label + "のバックアップ保存に失敗しました。", true);
      console.warn("Failed to save backup:", error);
    } finally {
      state.working = false;
      render();
    }
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

  async function deleteBackup(definition) {
    if (!window.confirm(definition.label + "のバックアップを削除しますか？")) {
      return;
    }

    state.working = true;
    render();

    try {
      await getBackupDocRef(definition).delete();
      await refreshBackups();
      setGlobalStatus(definition.label + "のバックアップを削除しました。");
    } catch (error) {
      setGlobalStatus(definition.label + "のバックアップ削除に失敗しました。", true);
      console.warn("Failed to delete backup:", error);
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
