(function () {
  "use strict";

  const sharedSettings = window.SharedAppSettings || null;
  const referenceConfig = window.APP_FIREBASE_CONFIG || null;
  const referenceSyncOptions = window.APP_FIREBASE_SYNC_OPTIONS || {};
  const pointsConfig = window.DRIVER_POINTS_FIREBASE_CONFIG || null;
  const pointsSettings = window.DRIVER_POINTS_FIREBASE_SETTINGS || {};
  const SERVER_GET_OPTIONS = Object.freeze({
    source: "server"
  });
  const optionsDocRefs = Object.freeze({
    vehicles: {
      collection: referenceSyncOptions.collection || "monthly_tire_autosave",
      id: "monthly_tire_company_settings_backup_vehicles_slot1"
    },
    drivers: {
      collection: referenceSyncOptions.collection || "monthly_tire_autosave",
      id: "monthly_tire_company_settings_backup_drivers_slot1"
    }
  });

  const elements = {
    vehicleSelect: document.getElementById("vehicleSelect"),
    driverSelect: document.getElementById("driverSelect"),
    reloadButton: document.getElementById("reloadButton"),
    saveButton: document.getElementById("saveButton"),
    openLeaderboardButton: document.getElementById("openLeaderboardButton"),
    closeLeaderboardButton: document.getElementById("closeLeaderboardButton"),
    leaderboardOverlay: document.getElementById("leaderboardOverlay"),
    leaderboardStatus: document.getElementById("leaderboardStatus"),
    leaderboardBody: document.getElementById("leaderboardBody"),
    pointsValue: document.getElementById("pointsValue"),
    pointsInput: document.getElementById("pointsInput"),
    pointsMeta: document.getElementById("pointsMeta"),
    statusText: document.getElementById("statusText")
  };

  const state = {
    optionSourceReady: false,
    pointsDb: null,
    activeSchema: null,
    currentRecord: null,
    vehicleOptions: [],
    driverOptions: [],
    leaderboardRows: [],
    loadingPoints: false,
    savingPoints: false,
    loadingLeaderboard: false
  };

  bindEvents();
  void initialize();

  function bindEvents() {
    elements.vehicleSelect.addEventListener("change", function () {
      void loadPointsForCurrentSelection();
    });

    elements.driverSelect.addEventListener("change", function () {
      void loadPointsForCurrentSelection();
    });

    elements.reloadButton.addEventListener("click", function () {
      void loadPointsForCurrentSelection(true);
    });

    elements.saveButton.addEventListener("click", function () {
      void savePoints();
    });

    elements.openLeaderboardButton.addEventListener("click", function () {
      void openLeaderboard();
    });

    elements.closeLeaderboardButton.addEventListener("click", function () {
      closeLeaderboard();
    });

    elements.pointsInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        void savePoints();
      }
    });

    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && !elements.leaderboardOverlay.hidden) {
        closeLeaderboard();
      }
    });
  }

  async function initialize() {
    setStatus("候補を読み込んでいます...");
    syncButtons();

    try {
      await Promise.all([
        loadSelectableOptions(),
        initializePointsDb()
      ]);

      state.optionSourceReady = true;

      if (state.vehicleOptions.length && state.driverOptions.length) {
        elements.vehicleSelect.value = state.vehicleOptions[0].value;
        elements.driverSelect.value = state.driverOptions[0].value;
        await loadPointsForCurrentSelection();
      } else {
        setStatus("車番または乗務員の候補がまだありません。設定画面で登録してください。", true);
      }
    } catch (error) {
      console.warn("Failed to initialize driver points page:", error);
      setStatus("初期化に失敗しました: " + formatError(error), true);
    } finally {
      syncButtons();
    }
  }

  async function loadSelectableOptions() {
    const localOptions = getLocalOptions();
    let cloudVehicles = [];
    let cloudDrivers = [];

    if (referenceConfig && window.firebase) {
      try {
        const referenceDb = await ensureDb(referenceConfig, {
          useAnonymousAuth: referenceSyncOptions.useAnonymousAuth !== false
        }, "shared-settings-reference");
        const snapshots = await Promise.all([
          referenceDb.collection(optionsDocRefs.vehicles.collection).doc(optionsDocRefs.vehicles.id).get(),
          referenceDb.collection(optionsDocRefs.drivers.collection).doc(optionsDocRefs.drivers.id).get()
        ]);

        cloudVehicles = getStringArray(snapshots[0].exists ? snapshots[0].data() : null);
        cloudDrivers = getStringArray(snapshots[1].exists ? snapshots[1].data() : null);
      } catch (error) {
        console.warn("Failed to load shared options from Firebase:", error);
      }
    }

    state.vehicleOptions = buildVehicleOptions(localOptions.vehicles, cloudVehicles);
    state.driverOptions = buildDriverOptions(localOptions.drivers, cloudDrivers);
    renderOptions(elements.vehicleSelect, state.vehicleOptions, "車番を選択");
    renderOptions(elements.driverSelect, state.driverOptions, "乗務員を選択");
  }

  async function initializePointsDb() {
    if (!pointsConfig || !window.firebase) {
      throw new Error("driver_points_config_missing");
    }

    state.pointsDb = await ensureDb(pointsConfig, pointsSettings, pointsSettings.appName || "driver-points-app");
  }

  async function ensureDb(config, settings, appName) {
    const app = getOrCreateFirebaseApp(config, appName);
    const auth = app.auth();

    if (settings.useAnonymousAuth !== false && !auth.currentUser) {
      try {
        await auth.signInAnonymously();
      } catch (error) {
        console.warn("Anonymous auth was not available:", error);
      }
    }

    return app.firestore();
  }

  function getOrCreateFirebaseApp(config, appName) {
    const existingApp = window.firebase.apps.find(function (app) {
      return app.name === appName;
    });
    if (existingApp) {
      return existingApp;
    }
    return window.firebase.initializeApp(config, appName);
  }

  function getLocalOptions() {
    if (!sharedSettings || typeof sharedSettings.ensureState !== "function") {
      return {
        vehicles: [],
        drivers: []
      };
    }

    const sharedState = sharedSettings.ensureState();
    return {
      vehicles: Array.isArray(sharedState.vehicles) ? sharedState.vehicles : [],
      drivers: Array.isArray(sharedState.drivers) ? sharedState.drivers : []
    };
  }

  function buildVehicleOptions() {
    const unique = [];
    const seen = new Set();

    Array.from(arguments).forEach(function (values) {
      (values || []).forEach(function (value) {
        const normalizedValue = normalizeText(value);
        if (!normalizedValue || seen.has(normalizedValue)) {
          return;
        }
        seen.add(normalizedValue);
        unique.push({
          value: normalizedValue,
          label: normalizedValue,
          key: normalizedValue
        });
      });
    });

    return unique.sort(function (left, right) {
      return left.label.localeCompare(right.label, "ja", { numeric: true });
    });
  }

  function buildDriverOptions() {
    const mergedDrivers = [];

    Array.from(arguments).forEach(function (values) {
      (values || []).forEach(function (value) {
        const rawValue = normalizeText(value);
        if (rawValue) {
          mergedDrivers.push(rawValue);
        }
      });
    });

    const orderedDrivers = sharedSettings && typeof sharedSettings.normalizeDrivers === "function"
      ? sharedSettings.normalizeDrivers(mergedDrivers)
      : mergedDrivers;
    const options = [];
    const seen = new Set();

    orderedDrivers.forEach(function (value) {
      const rawValue = normalizeText(value);
      const label = normalizeDriverName(value);
      const key = normalizeDriverKey(value);

      if (!label || !key || seen.has(key)) {
        return;
      }

      seen.add(key);
      options.push({
        value: rawValue || label,
        label: label,
        key: key
      });
    });

    return options;
  }

  function renderOptions(select, options, placeholder) {
    select.innerHTML = "";

    const placeholderOption = document.createElement("option");
    placeholderOption.value = "";
    placeholderOption.textContent = placeholder;
    select.appendChild(placeholderOption);

    options.forEach(function (option) {
      const optionElement = document.createElement("option");
      optionElement.value = option.value;
      optionElement.textContent = option.label;
      select.appendChild(optionElement);
    });
  }

  async function loadPointsForCurrentSelection(forceReloadSchema) {
    const vehicle = normalizeText(elements.vehicleSelect.value);
    const driverOption = getSelectedDriverOption();

    state.currentRecord = null;
    elements.pointsMeta.textContent = "ポイントを検索しています...";
    elements.pointsValue.textContent = "--";

    if (!vehicle || !driverOption) {
      elements.pointsMeta.textContent = "車番と乗務員を選択してください。";
      elements.pointsInput.value = "";
      setStatus("車番と乗務員を選択してください。", true);
      syncButtons();
      return;
    }

    if (!state.pointsDb) {
      setStatus("ポイント用 Firebase に接続できていません。", true);
      syncButtons();
      return;
    }

    state.loadingPoints = true;
    syncButtons();

    try {
      const schema = await resolveSchema(forceReloadSchema === true);
      const record = await findPointRecord(schema, vehicle, driverOption);

      state.activeSchema = schema;
      state.currentRecord = record;

      if (!record) {
        elements.pointsValue.textContent = "0";
        elements.pointsInput.value = "";
        elements.pointsMeta.textContent = "該当データはまだありません。新規保存すると作成されます。";
        setStatus("該当するポイントデータは未登録です。");
        return;
      }

      const points = getRecordPoints(record, schema);
      elements.pointsValue.textContent = String(points);
      elements.pointsInput.value = "";
      elements.pointsMeta.textContent = "";
      setStatus("ポイントを読み込みました。");
    } catch (error) {
      console.warn("Failed to load points:", error);
      elements.pointsMeta.textContent = "ポイントの読み込みに失敗しました。";
      setStatus("ポイントの読み込みに失敗しました: " + formatError(error), true);
    } finally {
      state.loadingPoints = false;
      syncButtons();
    }
  }

  async function resolveSchema(forceReloadSchema) {
    if (!forceReloadSchema && state.activeSchema) {
      return state.activeSchema;
    }
    const collectionCandidates = buildCollectionCandidates();

    let firstUsableSchema = null;
    for (const collectionName of collectionCandidates) {
      const schema = await inspectCollection(collectionName);
      if (!schema) {
        continue;
      }

      if (!firstUsableSchema) {
        firstUsableSchema = schema;
      }

      if (schema.detectedFromDocuments) {
        state.activeSchema = schema;
        return schema;
      }
    }

    state.activeSchema = firstUsableSchema || buildFallbackSchema(pointsSettings.preferredCollection || "driverPoints");
    return state.activeSchema;
  }

  function buildCollectionCandidates() {
    const preferredCollection = normalizeText(pointsSettings.preferredCollection);
    if (preferredCollection) {
      return [preferredCollection];
    }

    const candidates = [];
    const seen = new Set();

    (pointsSettings.collectionCandidates || []).forEach(function (value) {
      const normalizedValue = normalizeText(value);
      if (!normalizedValue || seen.has(normalizedValue)) {
        return;
      }
      seen.add(normalizedValue);
      candidates.push(normalizedValue);
    });

    return candidates;
  }

  async function inspectCollection(collectionName) {
    if (!collectionName) {
      return null;
    }

    try {
      const snapshot = await getServerQuerySnapshot(
        state.pointsDb.collection(collectionName).limit(100)
      );
      if (snapshot.empty) {
        return buildFallbackSchema(collectionName);
      }

      const docs = snapshot.docs.map(function (docSnapshot) {
        return {
          id: docSnapshot.id,
          data: docSnapshot.data() || {}
        };
      });

      return inferSchema(collectionName, docs);
    } catch (error) {
      console.warn("Failed to inspect collection:", collectionName, error);
      return null;
    }
  }

  function inferSchema(collectionName, docs) {
    const sampleFields = new Set();

    docs.forEach(function (entry) {
      Object.keys(entry.data || {}).forEach(function (fieldName) {
        sampleFields.add(fieldName);
      });
    });

    const schema = {
      collectionName: collectionName,
      vehicleField: inferFieldName(sampleFields, pointsSettings.vehicleFieldCandidates, /vehicle|car|truck/i),
      driverField: inferFieldName(sampleFields, pointsSettings.driverFieldCandidates, /driver|name|staff|employee/i),
      pointsField: inferFieldName(sampleFields, pointsSettings.pointsFieldCandidates, /point|score/i),
      updatedAtField: inferFieldName(sampleFields, pointsSettings.updatedAtFieldCandidates, /updated|created|modified/i),
      docIdPatterns: Array.isArray(pointsSettings.docIdPatterns) ? pointsSettings.docIdPatterns.slice() : [],
      detectedFromDocuments: true
    };

    if (!schema.pointsField) {
      schema.pointsField = "points";
      schema.detectedFromDocuments = false;
    }
    if (!schema.vehicleField) {
      schema.vehicleField = "vehicle";
      schema.detectedFromDocuments = false;
    }
    if (!schema.driverField) {
      schema.driverField = "driver";
      schema.detectedFromDocuments = false;
    }
    if (!schema.updatedAtField) {
      schema.updatedAtField = "updatedAt";
    }

    return schema;
  }

  function inferFieldName(fieldNames, candidates, fallbackPattern) {
    for (const candidate of candidates || []) {
      if (fieldNames.has(candidate)) {
        return candidate;
      }
    }

    for (const fieldName of fieldNames) {
      if (fallbackPattern.test(fieldName)) {
        return fieldName;
      }
    }

    return "";
  }

  function buildFallbackSchema(collectionName) {
    return {
      collectionName: collectionName,
      vehicleField: "vehicleNumber",
      driverField: "driverKey",
      pointsField: "totalPoints",
      updatedAtField: "updatedAt",
      docIdPatterns: Array.isArray(pointsSettings.docIdPatterns) ? pointsSettings.docIdPatterns.slice() : [],
      detectedFromDocuments: false
    };
  }

  async function findPointRecord(schema, vehicle, driverOption) {
    let matchingRecord = null;
    const collectionRef = state.pointsDb.collection(schema.collectionName);
    const summaryKindValue = normalizeText(pointsSettings.summaryKindValue);

    if (summaryKindValue) {
      for (const vehicleField of uniqueFieldNames((pointsSettings.vehicleFieldCandidates || []).concat(schema.vehicleField))) {
        try {
          const summarySnapshot = await getServerQuerySnapshot(
            collectionRef
              .where("kind", "==", summaryKindValue)
              .where(vehicleField, "==", vehicle)
              .limit(100)
          );
          const summaryCandidates = summarySnapshot.docs.map(function (docSnapshot) {
            return {
              id: docSnapshot.id,
              ref: docSnapshot.ref,
              data: docSnapshot.data() || {}
            };
          }).filter(function (record) {
            return recordMatchesSelection(record, schema, vehicle, driverOption);
          });

          matchingRecord = pickBestPointRecord(summaryCandidates, schema);
          if (matchingRecord) {
            return matchingRecord;
          }
        } catch (error) {
          console.warn("Summary filtered query failed:", vehicleField, error);
        }
      }

      try {
        const summarySnapshot = await getServerQuerySnapshot(
          collectionRef
            .where("kind", "==", summaryKindValue)
            .limit(300)
        );
        const summaryCandidates = summarySnapshot.docs.map(function (docSnapshot) {
          return {
            id: docSnapshot.id,
            ref: docSnapshot.ref,
            data: docSnapshot.data() || {}
          };
        }).filter(function (record) {
          return recordMatchesSelection(record, schema, vehicle, driverOption);
        });

        matchingRecord = pickBestPointRecord(summaryCandidates, schema);
        if (matchingRecord) {
          return matchingRecord;
        }
      } catch (error) {
        console.warn("Summary collection scan failed:", error);
      }
    }

    for (const vehicleField of uniqueFieldNames((pointsSettings.vehicleFieldCandidates || []).concat(schema.vehicleField))) {
      try {
        const filteredSnapshot = await getServerQuerySnapshot(
          collectionRef
            .where(vehicleField, "==", vehicle)
            .limit(100)
        );
        const candidates = filteredSnapshot.docs.map(function (docSnapshot) {
          return {
            id: docSnapshot.id,
            ref: docSnapshot.ref,
            data: docSnapshot.data() || {}
          };
        }).filter(function (record) {
          return recordMatchesSelection(record, schema, vehicle, driverOption);
        });

        matchingRecord = pickBestPointRecord(candidates, schema);
        if (matchingRecord) {
          return matchingRecord;
        }
      } catch (error) {
        console.warn("Vehicle filtered query failed:", vehicleField, error);
      }
    }

    if (!matchingRecord) {
      const snapshot = await getServerQuerySnapshot(
        collectionRef.limit(300)
      );
      const candidates = snapshot.docs
        .map(function (docSnapshot) {
          return {
            id: docSnapshot.id,
            ref: docSnapshot.ref,
            data: docSnapshot.data() || {}
          };
        })
        .filter(function (record) {
          return recordMatchesSelection(record, schema, vehicle, driverOption);
        });
      matchingRecord = pickBestPointRecord(candidates, schema);
    }

    if (matchingRecord) {
      return matchingRecord;
    }

    for (const docId of buildCandidateDocIds(schema, vehicle, driverOption)) {
      const docSnapshot = await getServerDocumentSnapshot(
        state.pointsDb.collection(schema.collectionName).doc(docId)
      );
      if (docSnapshot.exists) {
        return {
          id: docSnapshot.id,
          ref: docSnapshot.ref,
          data: docSnapshot.data() || {}
        };
      }
    }

    return null;
  }

  function recordMatchesSelection(record, schema, vehicle, driverOption) {
    const vehicleKeys = collectNormalizedFieldValues(
      record.data,
      uniqueFieldNames([schema.vehicleField].concat(pointsSettings.vehicleFieldCandidates || [])),
      normalizeText
    );
    const driverKeys = collectNormalizedFieldValues(
      record.data,
      uniqueFieldNames(["driverRaw", "driverDisplay", "driverAliases", schema.driverField].concat(pointsSettings.driverFieldCandidates || [])),
      normalizeDriverKey
    );
    const selectedDriverKeys = [driverOption.value, driverOption.label]
      .map(function (value) {
        return normalizeDriverKey(value);
      })
      .filter(Boolean);

    return vehicleKeys.includes(vehicle) && driverKeys.some(function (value) {
      return selectedDriverKeys.includes(value);
    });
  }

  function pickBestPointRecord(records, schema) {
    if (!Array.isArray(records) || !records.length) {
      return null;
    }

    const candidates = records.filter(isSummaryPointRecord);
    if (!candidates.length) {
      return null;
    }

    candidates.sort(function (left, right) {
      return getRecordUpdatedAtTime(right.data, schema) - getRecordUpdatedAtTime(left.data, schema);
    });

    return candidates[0] || null;
  }

  function buildCandidateDocIds(schema, vehicle, driverOption) {
    const docIds = [];
    const seen = new Set();
    const rawDriver = normalizeText(driverOption.value);
    const displayDriver = normalizeText(driverOption.label);

    (schema.docIdPatterns || []).forEach(function (pattern) {
      [
        pattern.replace("{vehicle}", sanitizeDocIdPart(vehicle)).replace("{driver}", sanitizeDocIdPart(rawDriver)),
        pattern.replace("{vehicle}", sanitizeDocIdPart(vehicle)).replace("{driver}", sanitizeDocIdPart(displayDriver))
      ].forEach(function (docId) {
        if (!docId || seen.has(docId)) {
          return;
        }
        seen.add(docId);
        docIds.push(docId);
      });
    });

    return docIds;
  }

  async function savePoints() {
    const vehicle = normalizeText(elements.vehicleSelect.value);
    const driverOption = getSelectedDriverOption();
    const rawDeltaPoints = normalizeText(elements.pointsInput.value);
    const deltaPoints = Number(rawDeltaPoints);

    if (!vehicle || !driverOption) {
      setStatus("保存する前に車番と乗務員を選択してください。", true);
      return;
    }

    if (!rawDeltaPoints || !Number.isFinite(deltaPoints)) {
      setStatus("加点または減点するポイント数を入力してください。", true);
      return;
    }

    if (!state.pointsDb) {
      setStatus("ポイント用 Firebase に接続できていません。", true);
      return;
    }

    state.savingPoints = true;
    syncButtons();

    try {
      const schema = await resolveSchema(false);
      const existingRecord = await findPointRecord(schema, vehicle, driverOption);
      const collectionRef = state.pointsDb.collection(schema.collectionName);
      const currentPoints = existingRecord ? getRecordPoints(existingRecord, schema) : 0;
      const nextPoints = currentPoints + deltaPoints;
      const payload = {};
      const pointsFieldName = resolvePointsFieldName(existingRecord ? existingRecord.data : null, schema);

      payload[pointsFieldName] = nextPoints;
      payload[schema.updatedAtField] = window.firebase.firestore.FieldValue.serverTimestamp();

      if (existingRecord) {
        await collectionRef.doc(existingRecord.id).set(payload, { merge: true });
      } else {
        payload.kind = normalizeText(pointsSettings.summaryKindValue) || "driver_points_summary";
        payload.vehicleNumber = vehicle;
        payload.vehicleKey = vehicle;
        payload.driverName = driverOption.label;
        payload.driverKey = normalizeDriverKey(driverOption.label);
        payload[schema.vehicleField] = vehicle;
        payload[schema.driverField] = schema.driverField === "driverKey"
          ? normalizeDriverKey(driverOption.label)
          : driverOption.label;
        payload.driverRaw = driverOption.value;
        payload.createdAt = window.firebase.firestore.FieldValue.serverTimestamp();

        const newDocId = buildNewDocId(schema, vehicle, driverOption);
        await collectionRef.doc(newDocId).set(payload, { merge: true });
      }

      await loadPointsForCurrentSelection();
      setStatus("ポイントを更新しました: " + currentPoints + (deltaPoints >= 0 ? " + " : " - ") + Math.abs(deltaPoints) + " = " + nextPoints);
    } catch (error) {
      console.warn("Failed to save points:", error);
      setStatus("ポイントの保存に失敗しました: " + formatError(error), true);
    } finally {
      state.savingPoints = false;
      syncButtons();
    }
  }

  function buildNewDocId(schema, vehicle, driverOption) {
    const summaryPrefix = "driver_points_summary_";
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return summaryPrefix + window.crypto.randomUUID().replace(/-/g, "").slice(0, 8);
    }
    return summaryPrefix + Date.now().toString(16);
  }

  async function openLeaderboard() {
    elements.leaderboardOverlay.hidden = false;
    elements.leaderboardOverlay.setAttribute("aria-hidden", "false");
    document.body.classList.add("leaderboard-open");
    elements.leaderboardStatus.textContent = "一覧を読み込んでいます...";
    elements.leaderboardBody.innerHTML = '<tr><td colspan="3">読み込み中...</td></tr>';
    await loadLeaderboard();
  }

  function closeLeaderboard() {
    elements.leaderboardOverlay.hidden = true;
    elements.leaderboardOverlay.setAttribute("aria-hidden", "true");
    document.body.classList.remove("leaderboard-open");
  }

  async function loadLeaderboard() {
    if (!state.pointsDb) {
      elements.leaderboardStatus.textContent = "Firebase に接続できていません。";
      elements.leaderboardBody.innerHTML = '<tr><td colspan="3">読み込みに失敗しました。</td></tr>';
      return;
    }

    state.loadingLeaderboard = true;
    syncButtons();

    try {
      const schema = await resolveSchema(true);
      const summaryKindValue = normalizeText(pointsSettings.summaryKindValue);
      const collectionRef = state.pointsDb.collection(schema.collectionName);
      let records = [];

      if (summaryKindValue) {
        const snapshot = await getServerQuerySnapshot(
          collectionRef.where("kind", "==", summaryKindValue)
        );
        records = snapshot.docs.map(function (docSnapshot) {
          return {
            id: docSnapshot.id,
            ref: docSnapshot.ref,
            data: docSnapshot.data() || {}
          };
        });
      } else {
        const snapshot = await getServerQuerySnapshot(collectionRef);
        records = snapshot.docs.map(function (docSnapshot) {
          return {
            id: docSnapshot.id,
            ref: docSnapshot.ref,
            data: docSnapshot.data() || {}
          };
        }).filter(isSummaryPointRecord);
      }

      state.leaderboardRows = buildLeaderboardRows(records, schema);
      renderLeaderboard();
    } catch (error) {
      console.warn("Failed to load leaderboard:", error);
      elements.leaderboardStatus.textContent = "一覧の読み込みに失敗しました: " + formatError(error);
      elements.leaderboardBody.innerHTML = '<tr><td colspan="3">一覧の読み込みに失敗しました。</td></tr>';
    } finally {
      state.loadingLeaderboard = false;
      syncButtons();
    }
  }

  function buildLeaderboardRows(records, schema) {
    const latestByDriverVehicle = new Map();
    const driverMetaByKey = new Map();
    const driverOrderByKey = new Map();

    state.driverOptions.forEach(function (option, index) {
      if (!driverOrderByKey.has(option.key)) {
        driverOrderByKey.set(option.key, index);
      }
      if (!driverMetaByKey.has(option.key)) {
        driverMetaByKey.set(option.key, {
          key: option.key,
          name: option.label,
          order: index
        });
      }
    });

    (records || []).forEach(function (record) {
      if (!isSummaryPointRecord(record)) {
        return;
      }

      const driverKey = resolveRecordDriverKey(record.data);
      if (!driverKey) {
        return;
      }

      const vehicleKey = resolveRecordVehicleKey(record.data);
      const pairKey = vehicleKey + "::" + driverKey;
      const currentTime = getRecordUpdatedAtTime(record.data, schema);
      const existingRecord = latestByDriverVehicle.get(pairKey);

      if (!existingRecord || currentTime >= getRecordUpdatedAtTime(existingRecord.data, schema)) {
        latestByDriverVehicle.set(pairKey, record);
      }

      if (!driverMetaByKey.has(driverKey)) {
        driverMetaByKey.set(driverKey, {
          key: driverKey,
          name: resolveRecordDriverLabel(record.data),
          order: Number.MAX_SAFE_INTEGER
        });
      }
    });

    const totalsByDriver = new Map();

    latestByDriverVehicle.forEach(function (record) {
      const driverKey = resolveRecordDriverKey(record.data);
      if (!driverKey) {
        return;
      }

      const points = getRecordPoints(record, schema);
      const existing = totalsByDriver.get(driverKey);
      if (existing) {
        existing.points += points;
        return;
      }

      const meta = driverMetaByKey.get(driverKey) || {
        key: driverKey,
        name: resolveRecordDriverLabel(record.data),
        order: Number.MAX_SAFE_INTEGER
      };
      totalsByDriver.set(driverKey, {
        key: driverKey,
        name: meta.name,
        order: meta.order,
        points: points
      });
    });

    driverMetaByKey.forEach(function (meta, driverKey) {
      if (!totalsByDriver.has(driverKey)) {
        totalsByDriver.set(driverKey, {
          key: driverKey,
          name: meta.name,
          order: meta.order,
          points: 0
        });
      }
    });

    return Array.from(totalsByDriver.values()).sort(function (left, right) {
      if (right.points !== left.points) {
        return right.points - left.points;
      }
      if (left.order !== right.order) {
        return left.order - right.order;
      }
      return left.name.localeCompare(right.name, "ja", { numeric: true, sensitivity: "base" });
    });
  }

  function renderLeaderboard() {
    const rows = state.leaderboardRows || [];

    if (!rows.length) {
      elements.leaderboardStatus.textContent = "表示できる社員データがありません。";
      elements.leaderboardBody.innerHTML = '<tr><td colspan="3">表示できる社員データがありません。</td></tr>';
      return;
    }

    elements.leaderboardStatus.textContent = rows.length + "名のポイントを表示しています。";
    elements.leaderboardBody.innerHTML = rows.map(function (row, index) {
      return [
        "<tr>",
        '<td class="leaderboard-rank">' + String(index + 1) + "</td>",
        '<td class="leaderboard-name">' + escapeHtml(row.name) + "</td>",
        '<td class="leaderboard-points">' + String(row.points) + "</td>",
        "</tr>"
      ].join("");
    }).join("");
  }

  function resolveRecordDriverKey(source) {
    const values = collectNormalizedFieldValues(
      source,
      uniqueFieldNames(["driverKey", "driverRaw", "driverName", "driver", "driverDisplay", "driverAliases"]),
      normalizeDriverKey
    );
    return values[0] || "";
  }

  function resolveRecordVehicleKey(source) {
    const values = collectNormalizedFieldValues(
      source,
      uniqueFieldNames(["vehicleKey", "vehicleNumber", "vehicle", "vehicleNo"]),
      normalizeText
    );
    return values[0] || "";
  }

  function resolveRecordDriverLabel(source) {
    const safeSource = source && typeof source === "object" ? source : {};
    return normalizeText(safeSource.driverName)
      || normalizeDriverName(safeSource.driverRaw)
      || normalizeDriverName(safeSource.driver)
      || normalizeDriverName(safeSource.driverDisplay)
      || "名称未設定";
  }

  function getSelectedDriverOption() {
    const selectedValue = normalizeText(elements.driverSelect.value);
    if (!selectedValue) {
      return null;
    }

    return state.driverOptions.find(function (option) {
      return option.value === selectedValue;
    }) || null;
  }

  function getStringArray(source) {
    if (!source || !Array.isArray(source.values)) {
      return [];
    }
    return source.values.map(function (value) {
      return normalizeText(value);
    }).filter(Boolean);
  }

  function normalizeText(value) {
    return String(value == null ? "" : value).trim();
  }

  function normalizeDriverName(value) {
    if (sharedSettings && typeof sharedSettings.normalizeDriverName === "function") {
      return normalizeText(sharedSettings.normalizeDriverName(value));
    }
    return normalizeText(value);
  }

  function normalizeDriverKey(value) {
    return normalizeDriverName(value)
      .normalize("NFKC")
      .replace(/\s+/g, "")
      .toLowerCase();
  }

  function sanitizeDocIdPart(value) {
    return normalizeText(value).replace(/[\/\\?#\[\]]/g, "-");
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function isSummaryPointRecord(record) {
    const summaryKindValue = normalizeText(pointsSettings.summaryKindValue);
    const recordId = normalizeText(record && record.id);
    const recordKind = normalizeText(record && record.data ? record.data.kind : "");

    if (summaryKindValue && recordKind === summaryKindValue) {
      return true;
    }

    return recordId.startsWith("driver_points_summary_");
  }

  function resolvePointsFieldName(source, schema) {
    const fieldNames = uniqueFieldNames((pointsSettings.pointsFieldCandidates || []).concat(schema.pointsField));
    const safeSource = source && typeof source === "object" ? source : {};

    for (const fieldName of fieldNames) {
      const value = safeSource[fieldName];
      if (typeof value === "number") {
        return fieldName;
      }
      if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) {
        return fieldName;
      }
    }

    return fieldNames[0] || "totalPoints";
  }

  function getRecordPoints(record, schema) {
    const fieldName = resolvePointsFieldName(record && record.data ? record.data : null, schema);
    return getNumericValue(record && record.data ? record.data[fieldName] : undefined);
  }

  function uniqueFieldNames(values) {
    const unique = [];
    const seen = new Set();

    (values || []).forEach(function (value) {
      const fieldName = normalizeText(value);
      if (!fieldName || seen.has(fieldName)) {
        return;
      }
      seen.add(fieldName);
      unique.push(fieldName);
    });

    return unique;
  }

  function collectNormalizedFieldValues(source, fieldNames, normalizer) {
    const normalize = typeof normalizer === "function" ? normalizer : normalizeText;
    const values = [];
    const seen = new Set();
    const safeSource = source && typeof source === "object" ? source : {};

    fieldNames.forEach(function (fieldName) {
      const rawValue = safeSource[fieldName];
      const entries = Array.isArray(rawValue) ? rawValue : [rawValue];

      entries.forEach(function (entry) {
        const normalizedValue = normalize(entry);
        if (!normalizedValue || seen.has(normalizedValue)) {
          return;
        }
        seen.add(normalizedValue);
        values.push(normalizedValue);
      });
    });

    return values;
  }

  function getRecordUpdatedAtTime(source, schema) {
    const fieldNames = uniqueFieldNames([schema.updatedAtField].concat(pointsSettings.updatedAtFieldCandidates || []));
    const safeSource = source && typeof source === "object" ? source : {};

    for (const fieldName of fieldNames) {
      const value = safeSource[fieldName];
      const time = getTimeValue(value);
      if (time > 0) {
        return time;
      }
    }

    return 0;
  }

  function getNumericValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }
    const numericValue = Number(value);
    return Number.isFinite(numericValue) ? numericValue : 0;
  }

  function getTimeValue(value) {
    if (!value) {
      return 0;
    }
    if (typeof value.toDate === "function") {
      try {
        return value.toDate().getTime();
      } catch {
        return 0;
      }
    }
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? 0 : value.getTime();
    }
    const numericValue = Number(value);
    if (Number.isFinite(numericValue) && numericValue > 0) {
      return numericValue;
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
  }

  function formatDateTime(value) {
    if (!value) {
      return "";
    }

    let date = null;
    if (typeof value.toDate === "function") {
      try {
        date = value.toDate();
      } catch {
        date = null;
      }
    } else if (value instanceof Date) {
      date = value;
    } else {
      const parsed = new Date(value);
      if (!Number.isNaN(parsed.getTime())) {
        date = parsed;
      }
    }

    if (!date) {
      return "";
    }

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    return year + "-" + month + "-" + day + " " + hours + ":" + minutes;
  }

  function syncButtons() {
    const hasSelection = Boolean(normalizeText(elements.vehicleSelect.value) && getSelectedDriverOption());
    elements.reloadButton.disabled = !hasSelection || state.loadingPoints || !state.optionSourceReady;
    elements.saveButton.disabled = !hasSelection || state.loadingPoints || state.savingPoints || !state.pointsDb;
    elements.openLeaderboardButton.disabled = state.loadingLeaderboard || !state.pointsDb;
  }

  function setStatus(message, isError) {
    elements.statusText.textContent = message || "";
    elements.statusText.style.color = isError ? "#b00020" : "";
  }

  function formatError(error) {
    if (!error) {
      return "unknown_error";
    }
    if (error.code) {
      return String(error.code);
    }
    if (error.message) {
      return String(error.message);
    }
    return String(error);
  }

  function getServerQuerySnapshot(query) {
    return query.get(SERVER_GET_OPTIONS);
  }

  function getServerDocumentSnapshot(docRef) {
    return docRef.get(SERVER_GET_OPTIONS);
  }
})();
