// Firebase runtime config
// Fill production values and set enabled=true.
// Keep keys out of index.html by using this file.
(function () {
  "use strict";

  window.APP_FIREBASE_CONFIG = {
    apiKey: "AIzaSyAlpiGkwyoEW8U8X7HpK4XiqfwW8e_YOdQ",
    authDomain: "getujityretenkenhyou.firebaseapp.com",
    projectId: "getujityretenkenhyou",
    appId: "1:818371379903:web:421a1b390e41a48d2cfc0a",
    messagingSenderId: "818371379903",
    storageBucket: "getujityretenkenhyou.firebasestorage.app"
  };

  window.APP_FIREBASE_SYNC_OPTIONS = {
    enabled: true,
    // Firestore collection name
    collection: "monthly_tire_autosave",
    // Prefix for document id
    documentPrefix: "monthly_tire",
    // Company identifier for future access control
    companyCode: "company",
    // Use anonymous auth (no user login UI)
    useAnonymousAuth: true,
    // Retry flush interval (ms)
    autoFlushIntervalMs: 15000
  };

  // Optional additional sync target for future shared master data.
  // Fill the config after creating the employee directory project,
  // then set enabled=true to dual-write from settings.html.
  window.APP_FIREBASE_DIRECTORY_CONFIG = {
    apiKey: "AIzaSyBRfFhSznNAHdLPtgHqBTnLaRosAoFWuEA",
    authDomain: "syainmeibo-d78af.firebaseapp.com",
    projectId: "syainmeibo-d78af",
    storageBucket: "syainmeibo-d78af.firebasestorage.app",
    messagingSenderId: "379907949120",
    appId: "1:379907949120:web:35fa133fd5e954be5f15a5",
    measurementId: "G-7V1NZEGC18"
  };

  window.APP_FIREBASE_DIRECTORY_SYNC_OPTIONS = {
    enabled: true,
    appName: "employee-directory",
    collection: "monthly_tire_autosave",
    docIds: {
      vehicles: "monthly_tire_company_settings_backup_vehicles_slot1",
      drivers: "monthly_tire_company_settings_backup_drivers_slot1"
    },
    useAnonymousAuth: true
  };
})();
