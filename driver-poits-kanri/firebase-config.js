(function () {
  "use strict";

  window.DRIVER_POINTS_FIREBASE_CONFIG = {
    apiKey: "AIzaSyB-aw3aOIGPdyXAY6N-3a4W2Qx3hL9VLsI",
    authDomain: "driver-points.firebaseapp.com",
    projectId: "driver-points",
    storageBucket: "driver-points.firebasestorage.app",
    messagingSenderId: "273561411339",
    appId: "1:273561411339:web:5586a980cd0c6d79ec0c0e",
    measurementId: "G-9JYHCD7Q90"
  };

  window.DRIVER_POINTS_FIREBASE_SETTINGS = {
    appName: "driver-points-app",
    useAnonymousAuth: true,
    preferredCollection: "monthly_tire_autosave",
    collectionCandidates: [
      "monthly_tire_autosave",
      "driverPoints",
      "driver_points",
      "points",
      "driver-points",
      "pointRecords",
      "driverPointRecords"
    ],
    vehicleFieldCandidates: [
      "vehicleNumber",
      "vehicleKey",
      "vehicle",
      "carNumber",
      "truckNumber",
      "vehicleNo"
    ],
    driverFieldCandidates: [
      "driverKey",
      "driver",
      "driverName",
      "driverDisplay",
      "employeeName",
      "staffName",
      "name"
    ],
    pointsFieldCandidates: [
      "totalPoints",
      "dailyInspectionPoints",
      "points",
      "point",
      "grantedPoints",
      "currentPoints",
      "score"
    ],
    updatedAtFieldCandidates: [
      "updatedAt",
      "updated_at",
      "lastUpdatedAt",
      "modifiedAt",
      "createdAt"
    ],
    summaryKindValue: "driver_points_summary",
    docIdPatterns: [
      "{vehicle}__{driver}",
      "{vehicle}_{driver}",
      "{vehicle}-{driver}",
      "{driver}__{vehicle}",
      "{driver}_{vehicle}",
      "{driver}-{vehicle}"
    ]
  };
})();
