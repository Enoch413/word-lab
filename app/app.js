const STORAGE_KEY = "word-lab-desktop-state-v1";
const HANDLE_DB_NAME = "word-lab-directory-handles";
const HANDLE_STORE_NAME = "handles";
const SAVE_ROOT_HANDLE_KEY = "save-root";
const SNAPSHOT_FILE_NAME = "WORD_LAB_단어복원.json";
const MAX_PREVIEW_ITEMS = 5;
const LAB_STORAGE_NAME = "WORD LAB";
const LAST_USED_CLASS_KEY_PREFIX = `${LAB_STORAGE_NAME}:last-class`;
const FIREBASE_SDK_VERSION = "10.12.2";
const FIREBASE_CONFIG_GLOBAL_KEYS = [
  "__WORD_LAB_FIREBASE_CONFIG__",
  "__CODELAB_FIREBASE_CONFIG__",
  "__FIREBASE_CONFIG__",
  "firebaseConfig",
];
const CLASS_PRESET_CONFIG = [
  { id: "mon-wed", label: "월/수 반" },
  { id: "tue-thu", label: "화/목 반" },
];
const DEFAULT_SETTINGS = {
  quizMode: "word-to-meaning",
  questionCount: 50,
  sessionCount: 3,
};

const state = {
  activePresetId: CLASS_PRESET_CONFIG[0].id,
  presets: createDefaultPresets(),
  saveRootHandle: null,
  rememberedSaveRootHandle: null,
  saveRootName: "",
  toastTimer: null,
  isGenerating: false,
  currentUser: null,
  userProfile: null,
};

const nativeBridge = {
  ready: false,
  nextId: 1,
  pending: new Map(),
};

const elements = {};
let firebaseServicesPromise = null;
let authContextPromise = null;

document.addEventListener("DOMContentLoaded", () => {
  void initializeApp();
});

function createDefaultPresets() {
  return CLASS_PRESET_CONFIG.reduce((accumulator, preset) => {
    accumulator[preset.id] = {
      classes: [],
      layoutLocked: false,
    };
    return accumulator;
  }, {});
}

function hasNativeHost() {
  return Boolean(window.chrome?.webview);
}

function initializeNativeBridge() {
  if (!hasNativeHost() || nativeBridge.ready) {
    return;
  }

  window.chrome.webview.addEventListener("message", (event) => {
    const message = event.data || {};
    const pending = nativeBridge.pending.get(message.id);

    if (!pending) {
      return;
    }

    nativeBridge.pending.delete(message.id);

    if (message.ok) {
      pending.resolve(message.result ?? null);
      return;
    }

    pending.reject(new Error(message.error || "네이티브 브리지 호출에 실패했습니다."));
  });

  nativeBridge.ready = true;
}

function callNativeHost(method, payload = {}) {
  if (!hasNativeHost()) {
    return Promise.reject(new Error("네이티브 호스트를 사용할 수 없습니다."));
  }

  initializeNativeBridge();

  const id = nativeBridge.nextId;
  nativeBridge.nextId += 1;

  return new Promise((resolve, reject) => {
    nativeBridge.pending.set(id, { resolve, reject });
    window.chrome.webview.postMessage({ id, method, payload });
  });
}

function applyNativeSaveRootStatus(status) {
  const remembered = Boolean(status?.remembered);
  const connected = Boolean(status?.connected);
  state.rememberedSaveRootHandle = remembered ? { native: true } : null;
  state.saveRootHandle = connected ? { native: true } : null;
  state.saveRootName = status?.name ? String(status.name) : "";
}

// Auth/profile lookup is optional. If it fails, the app keeps the existing local fallback behavior.
async function applyAuthenticatedAdminDefaults() {
  const authContext = await loadCurrentUserProfile();
  state.currentUser = authContext?.user || null;
  state.userProfile = authContext?.profile || null;

  if (!shouldApplyAdminDefaults(state.currentUser, state.userProfile)) {
    return;
  }

  const selection = resolveAdminClassSelection(
    state.currentUser.uid,
    state.userProfile
  );
  if (!selection) {
    return;
  }

  applyAdminClassSelection(selection);
  rememberLastUsedClassForCurrentAdmin(selection.classReference);
}

function shouldApplyAdminDefaults(user, profile) {
  return Boolean(
    user?.uid &&
    String(profile?.role || "").trim().toLocaleLowerCase() === "admin"
  );
}

async function loadCurrentUserProfile() {
  if (!authContextPromise) {
    authContextPromise = (async () => {
      try {
        const injectedContext = getInjectedAuthContext();
        if (injectedContext) {
          return injectedContext;
        }

        const compatContext = await loadCompatFirebaseUserContext();
        if (compatContext) {
          return compatContext;
        }

        return loadModularFirebaseUserContext();
      } catch (error) {
        console.error(error);
        return null;
      }
    })();
  }

  return authContextPromise;
}

function getInjectedAuthContext() {
  const candidate =
    window.__WORD_LAB_AUTH_CONTEXT__ || window.__CODELAB_AUTH_CONTEXT__;
  const user = candidate?.currentUser || candidate?.user || null;

  if (!user?.uid) {
    return null;
  }

  return {
    user,
    profile: candidate?.userProfile || candidate?.profile || null,
  };
}

async function loadCompatFirebaseUserContext() {
  const firebase = window.firebase;
  if (
    !firebase ||
    typeof firebase.auth !== "function" ||
    typeof firebase.firestore !== "function"
  ) {
    return null;
  }

  const auth = firebase.auth();
  const user = await new Promise((resolve) => {
    const unsubscribe = auth.onAuthStateChanged(
      (nextUser) => {
        unsubscribe();
        resolve(nextUser || null);
      },
      () => {
        unsubscribe();
        resolve(null);
      }
    );
  });

  if (!user?.uid) {
    return null;
  }

  let profile = null;

  try {
    const snapshot = await firebase
      .firestore()
      .collection("users")
      .doc(user.uid)
      .get();
    profile = snapshot.exists ? snapshot.data() || null : null;
  } catch (error) {
    console.error(error);
  }

  return { user, profile };
}

async function loadModularFirebaseUserContext() {
  const firebaseServices = await loadFirebaseServices();
  if (!firebaseServices) {
    return null;
  }

  const user = await new Promise((resolve) => {
    const unsubscribe = firebaseServices.onAuthStateChanged(
      firebaseServices.auth,
      (nextUser) => {
        unsubscribe();
        resolve(nextUser || null);
      },
      () => {
        unsubscribe();
        resolve(null);
      }
    );
  });

  if (!user?.uid) {
    return null;
  }

  let profile = null;

  try {
    const snapshot = await firebaseServices.getDoc(
      firebaseServices.doc(firebaseServices.db, "users", user.uid)
    );
    profile = snapshot.exists() ? snapshot.data() || null : null;
  } catch (error) {
    console.error(error);
  }

  return { user, profile };
}

async function loadFirebaseServices() {
  if (!firebaseServicesPromise) {
    firebaseServicesPromise = (async () => {
      const injectedServices =
        window.__WORD_LAB_FIREBASE__ || window.__CODELAB_FIREBASE__;
      if (
        injectedServices?.auth &&
        injectedServices?.db &&
        typeof injectedServices.onAuthStateChanged === "function" &&
        typeof injectedServices.doc === "function" &&
        typeof injectedServices.getDoc === "function"
      ) {
        return injectedServices;
      }

      const firebaseConfig = getFirebaseConfig();
      if (!firebaseConfig) {
        return null;
      }

      const [appModule, authModule, firestoreModule] = await Promise.all([
        import(
          `https://www.gstatic.com/firebasejs/${FIREBASE_SDK_VERSION}/firebase-app.js`
        ),
        import(
          `https://www.gstatic.com/firebasejs/${FIREBASE_SDK_VERSION}/firebase-auth.js`
        ),
        import(
          `https://www.gstatic.com/firebasejs/${FIREBASE_SDK_VERSION}/firebase-firestore.js`
        ),
      ]);

      const firebaseApp = appModule.getApps().length
        ? appModule.getApp()
        : appModule.initializeApp(firebaseConfig);

      return {
        auth: authModule.getAuth(firebaseApp),
        db: firestoreModule.getFirestore(firebaseApp),
        onAuthStateChanged: authModule.onAuthStateChanged,
        doc: firestoreModule.doc,
        getDoc: firestoreModule.getDoc,
      };
    })().catch((error) => {
      console.error(error);
      return null;
    });
  }

  return firebaseServicesPromise;
}

function getFirebaseConfig() {
  for (const key of FIREBASE_CONFIG_GLOBAL_KEYS) {
    const normalized = normalizeFirebaseConfig(window[key]);
    if (normalized) {
      return normalized;
    }
  }

  for (const key of FIREBASE_CONFIG_GLOBAL_KEYS) {
    try {
      const rawValue = localStorage.getItem(key);
      const normalized = normalizeFirebaseConfig(rawValue);
      if (normalized) {
        return normalized;
      }
    } catch (error) {
      console.error(error);
    }
  }

  return null;
}

function normalizeFirebaseConfig(candidate) {
  if (!candidate) {
    return null;
  }

  let value = candidate;
  if (typeof value === "string") {
    try {
      value = JSON.parse(value);
    } catch (error) {
      return null;
    }
  }

  if (
    !value ||
    typeof value !== "object" ||
    typeof value.apiKey !== "string" ||
    typeof value.projectId !== "string"
  ) {
    return null;
  }

  return value;
}

function resolveAdminClassSelection(uid, profile) {
  const profileClassIds = normalizeProfileClassIds(profile?.classIds);
  if (!profileClassIds.length) {
    if (normalizeLookupValue(profile?.adminScope) === "all") {
      return null;
    }
    return null;
  }

  const assignments = profileClassIds
    .map((classId) => resolveClassAssignment(classId))
    .filter(Boolean)
    .filter(
      (assignment, index, items) =>
        items.findIndex(
          (candidate) =>
            candidate.presetId === assignment.presetId &&
            candidate.classReference === assignment.classReference
        ) === index
    );

  if (!assignments.length) {
    return null;
  }

  if (assignments.length === 1) {
    return assignments[0];
  }

  const lastUsedClass = normalizeLookupValue(getLastUsedClassForUser(uid));
  if (lastUsedClass) {
    const matchingAssignment = assignments.find((assignment) =>
      assignment.matchKeys.includes(lastUsedClass)
    );
    if (matchingAssignment) {
      return matchingAssignment;
    }
  }

  return assignments[0];
}

function normalizeProfileClassIds(classIds) {
  if (!Array.isArray(classIds)) {
    return [];
  }

  return classIds
    .map((item) => {
      if (typeof item === "string") {
        return item;
      }
      if (item && typeof item === "object") {
        return item.classId || item.id || item.className || item.name || "";
      }
      return "";
    })
    .map((value) => String(value || "").trim())
    .filter(Boolean);
}

function resolveClassAssignment(classId) {
  return findClassAssignment(classId) || findPresetAssignment(classId);
}

function findClassAssignment(classId) {
  const normalizedClassId = normalizeLookupValue(classId);
  if (!normalizedClassId) {
    return null;
  }

  for (const preset of CLASS_PRESET_CONFIG) {
    const presetState = state.presets[preset.id];
    const classes = Array.isArray(presetState?.classes) ? presetState.classes : [];
    const matchedClass = classes.find((classData) =>
      getClassMatchKeys(classData).includes(normalizedClassId)
    );

    if (matchedClass) {
      return createClassAssignment(preset, matchedClass, classId);
    }
  }

  return null;
}

function findPresetAssignment(classId) {
  const normalizedClassId = normalizeLookupValue(classId);
  if (!normalizedClassId) {
    return null;
  }

  const preset = CLASS_PRESET_CONFIG.find((item) =>
    [item.id, item.label]
      .map((value) => normalizeLookupValue(value))
      .includes(normalizedClassId)
  );
  if (!preset) {
    return null;
  }

  const firstClass = state.presets[preset.id]?.classes?.[0] || null;
  return createClassAssignment(preset, firstClass, classId);
}

function createClassAssignment(preset, classData, classId) {
  const classReference = getClassReference(classData, classId || preset.id);
  const matchKeys = new Set([
    normalizeLookupValue(classId),
    normalizeLookupValue(preset.id),
    normalizeLookupValue(preset.label),
    normalizeLookupValue(classReference),
  ]);

  getClassMatchKeys(classData).forEach((value) => {
    matchKeys.add(value);
  });

  return {
    presetId: preset.id,
    classReference,
    matchKeys: Array.from(matchKeys).filter(Boolean),
  };
}

function applyAdminClassSelection(selection) {
  if (!selection?.presetId) {
    return;
  }

  moveClassToFront(selection.presetId, selection.classReference);
  state.activePresetId = selection.presetId;
  setCheckedValue("class-preset", state.activePresetId);
}

function moveClassToFront(presetId, classReference) {
  const presetState = state.presets[presetId];
  if (!presetState?.classes?.length || !classReference) {
    return;
  }

  const normalizedReference = normalizeLookupValue(classReference);
  const classIndex = presetState.classes.findIndex((classData) =>
    getClassMatchKeys(classData).includes(normalizedReference)
  );

  if (classIndex <= 0) {
    return;
  }

  const [selectedClass] = presetState.classes.splice(classIndex, 1);
  presetState.classes.unshift(selectedClass);
}

function getLastUsedClassForUser(uid) {
  if (!uid) {
    return "";
  }

  try {
    return String(
      localStorage.getItem(getLastUsedClassStorageKey(uid)) || ""
    ).trim();
  } catch (error) {
    console.error(error);
    return "";
  }
}

function rememberLastUsedClassForCurrentAdmin(classReference = "") {
  if (!shouldApplyAdminDefaults(state.currentUser, state.userProfile)) {
    return;
  }

  const value = String(classReference || getActiveClassReference() || "").trim();
  if (!value) {
    return;
  }

  try {
    localStorage.setItem(
      getLastUsedClassStorageKey(state.currentUser.uid),
      value
    );
  } catch (error) {
    console.error(error);
  }
}

function getLastUsedClassStorageKey(uid) {
  return `${LAST_USED_CLASS_KEY_PREFIX}:${uid}`;
}

function getActiveClassReference() {
  const activePreset = getActivePresetState();
  return getClassReference(activePreset.classes?.[0], state.activePresetId);
}

function getCardClassReference(card) {
  if (!card) {
    return state.activePresetId;
  }

  return getClassReference(
    {
      classId: card.dataset.classId || "",
      className: card.querySelector(".class-name-input")?.value || "",
    },
    state.activePresetId
  );
}

function getClassReference(classData, fallbackValue = "") {
  return String(
    classData?.classId ||
      classData?.id ||
      classData?.className ||
      classData?.name ||
      fallbackValue ||
      ""
  ).trim();
}

function getClassMatchKeys(classData) {
  return [
    classData?.classId,
    classData?.id,
    classData?.className,
    classData?.name,
  ]
    .map((value) => normalizeLookupValue(value))
    .filter(Boolean);
}

function normalizeLookupValue(value) {
  return String(value || "")
    .trim()
    .toLocaleLowerCase()
    .replace(/[\s/_-]+/g, "");
}

async function initializeApp() {
  cacheElements();
  bindGlobalEvents();
  initializeNativeBridge();
  restoreState();
  await applyAuthenticatedAdminDefaults();

  renderActivePreset();
  syncPresetUi();
  syncSaveRootUi();
  await restoreSaveRootHandle();
  updateSummary();
  setRunStatus("아직 생성 전입니다.", "success");
}

function cacheElements() {
  elements.classList = document.getElementById("class-list");
  elements.classTemplate = document.getElementById("class-card-template");
  elements.addClassBtn = document.getElementById("add-class-btn");
  elements.seedDemoBtn = document.getElementById("seed-demo-btn");
  elements.toggleLayoutLockBtn = document.getElementById("toggle-layout-lock-btn");
  elements.importSnapshotBtn = document.getElementById("import-snapshot-btn");
  elements.snapshotFileInput = document.getElementById("snapshot-file-input");
  elements.pickFolderBtn = document.getElementById("pick-folder-btn");
  elements.generateBtn = document.getElementById("generate-btn");
  elements.questionCountInput = document.getElementById("question-count");
  elements.sessionCountInput = document.getElementById("session-count-input");
  elements.sessionCountDisplay = document.getElementById("session-count-display");
  elements.sessionCountDecreaseBtn = document.getElementById("session-count-decrease");
  elements.sessionCountIncreaseBtn = document.getElementById("session-count-increase");
  elements.saveRootStatus = document.getElementById("save-root-status");
  elements.metricClassCount = document.getElementById("metric-class-count");
  elements.metricWordCount = document.getElementById("metric-word-count");
  elements.metricPdfCount = document.getElementById("metric-pdf-count");
  elements.metricFolderName = document.getElementById("metric-folder-name");
  elements.runStatus = document.getElementById("run-status");
  elements.toast = document.getElementById("toast");
  elements.renderStage = document.getElementById("render-stage");
}

function bindGlobalEvents() {
  elements.addClassBtn.addEventListener("click", () => {
    if (isActivePresetLocked()) {
      showToast("반 구성 편집을 먼저 켜야 반을 추가할 수 있습니다.");
      return;
    }

    addClassCard();
    persistState();
    updateSummary();
  });

  elements.seedDemoBtn.addEventListener("click", seedDemoData);
  elements.toggleLayoutLockBtn.addEventListener("click", toggleActivePresetLock);
  elements.importSnapshotBtn.addEventListener("click", () => {
    elements.snapshotFileInput.click();
  });
  elements.snapshotFileInput.addEventListener("change", handleSnapshotImport);
  elements.pickFolderBtn.addEventListener("click", handlePickFolderClick);
  elements.generateBtn.addEventListener("click", handleGenerate);

  elements.questionCountInput.addEventListener("input", () => {
    persistState();
    updateSummary();
  });

  elements.sessionCountDecreaseBtn.addEventListener("click", () => {
    adjustSessionCount(-1);
  });

  elements.sessionCountIncreaseBtn.addEventListener("click", () => {
    adjustSessionCount(1);
  });

  document.querySelectorAll("input[name='quiz-mode']").forEach((radio) => {
    radio.addEventListener("change", () => {
      persistState();
      updateSummary();
    });
  });

  document.querySelectorAll("input[name='class-preset']").forEach((radio) => {
    radio.addEventListener("change", () => {
      switchActivePreset(radio.value);
    });
  });
}

function restoreState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    applySettings(DEFAULT_SETTINGS);
    state.presets = createDefaultPresets();
    state.activePresetId = CLASS_PRESET_CONFIG[0].id;
    return;
  }

  try {
    const saved = JSON.parse(raw);
    const normalized = normalizeSavedState(saved);

    applySettings(normalized.settings);
    state.presets = normalized.presets;
    state.activePresetId = normalized.activePresetId;
    setCheckedValue("class-preset", state.activePresetId);
  } catch (error) {
    console.error(error);
    applySettings(DEFAULT_SETTINGS);
    state.presets = createDefaultPresets();
    state.activePresetId = CLASS_PRESET_CONFIG[0].id;
    showToast("이전 저장 상태를 불러오지 못해 기본값으로 시작했습니다.");
  }
}

function applySettings(settings) {
  const merged = { ...DEFAULT_SETTINGS, ...settings };
  elements.questionCountInput.value = clampNumber(
    merged.questionCount,
    5,
    100,
    DEFAULT_SETTINGS.questionCount
  );
  setCheckedValue("quiz-mode", merged.quizMode);
  setSessionCount(merged.sessionCount);
}

function normalizeSavedState(saved) {
  const defaults = createDefaultPresets();
  const legacyClasses = Array.isArray(saved.classes) ? sanitizeSavedClasses(saved.classes) : [];
  const presetSource = saved.presets && typeof saved.presets === "object"
    ? saved.presets
    : {};

  const presets = CLASS_PRESET_CONFIG.reduce((accumulator, preset) => {
    const incomingPreset = presetSource[preset.id] || {};
    const classes = Array.isArray(incomingPreset.classes)
      ? sanitizeSavedClasses(incomingPreset.classes)
      : preset.id === CLASS_PRESET_CONFIG[0].id && legacyClasses.length
        ? legacyClasses
        : [];

    accumulator[preset.id] = {
      classes: classes.length ? classes : defaults[preset.id].classes,
      layoutLocked: typeof incomingPreset.layoutLocked === "boolean"
        ? incomingPreset.layoutLocked
        : classes.length > 0,
    };
    return accumulator;
  }, {});

  const activePresetId = CLASS_PRESET_CONFIG.some((preset) => preset.id === saved.activePresetId)
    ? saved.activePresetId
    : CLASS_PRESET_CONFIG[0].id;

  return {
    settings: saved.settings || DEFAULT_SETTINGS,
    presets,
    activePresetId,
  };
}

function sanitizeSavedClasses(classes) {
  return classes.map((classData) => ({
    classId: String(classData?.classId || classData?.id || ""),
    className: String(classData?.className || ""),
    rawText: String(classData?.rawText || ""),
  }));
}

function persistState() {
  captureActivePresetClassesFromDom();
  const payload = {
    settings: getSettings(),
    activePresetId: state.activePresetId,
    presets: CLASS_PRESET_CONFIG.reduce((accumulator, preset) => {
      const presetState = state.presets[preset.id] || { classes: [], layoutLocked: false };
      accumulator[preset.id] = {
        classes: sanitizeSavedClasses(presetState.classes || []),
        layoutLocked: Boolean(presetState.layoutLocked),
      };
      return accumulator;
    }, {}),
  };

  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function getSettings() {
  return {
    quizMode: getCheckedValue("quiz-mode", DEFAULT_SETTINGS.quizMode),
    questionCount: clampNumber(
      Number(elements.questionCountInput.value),
      5,
      100,
      DEFAULT_SETTINGS.questionCount
    ),
    sessionCount: getSessionCount(),
  };
}

function getActivePresetState() {
  if (!state.presets[state.activePresetId]) {
    state.presets[state.activePresetId] = { classes: [], layoutLocked: false };
  }
  return state.presets[state.activePresetId];
}

function isActivePresetLocked() {
  return Boolean(getActivePresetState().layoutLocked);
}

function captureActivePresetClassesFromDom() {
  if (!elements.classList) {
    return;
  }

  const activePreset = getActivePresetState();
  activePreset.classes = collectClassCards().map((card) => ({
    classId: String(card.dataset.classId || ""),
    className: card.querySelector(".class-name-input").value,
    rawText: card.querySelector(".paste-area").value,
  }));
}

function buildSnapshotPayload(dateFolderName = "") {
  captureActivePresetClassesFromDom();

  return {
    version: 1,
    appName: "WORD LAB",
    exportedAt: new Date().toISOString(),
    dateFolderName,
    activePresetId: state.activePresetId,
    settings: getSettings(),
    presets: CLASS_PRESET_CONFIG.reduce((accumulator, preset) => {
      const presetState = state.presets[preset.id] || { classes: [], layoutLocked: false };
      accumulator[preset.id] = {
        label: preset.label,
        layoutLocked: Boolean(presetState.layoutLocked),
        classes: sanitizeSavedClasses(presetState.classes || []),
      };
      return accumulator;
    }, {}),
  };
}

function normalizeSnapshotImport(snapshot) {
  if (!snapshot || typeof snapshot !== "object") {
    throw new Error("복원 파일 내용이 비어 있습니다.");
  }

  const normalized = normalizeSavedState({
    settings: snapshot.settings || getSettings(),
    presets: snapshot.presets || null,
    activePresetId: snapshot.activePresetId,
  });

  return {
    settings: normalized.settings,
    presets: CLASS_PRESET_CONFIG.reduce((accumulator, preset) => {
      const importedPreset = snapshot.presets?.[preset.id] || {};
      accumulator[preset.id] = {
        classes: sanitizeSavedClasses(normalized.presets[preset.id].classes || []),
        layoutLocked: typeof importedPreset.layoutLocked === "boolean"
          ? importedPreset.layoutLocked
          : normalized.presets[preset.id].layoutLocked,
      };
      return accumulator;
    }, {}),
    activePresetId: normalized.activePresetId,
  };
}

function renderActivePreset() {
  const activePreset = getActivePresetState();
  elements.classList.innerHTML = "";

  const classesToRender = activePreset.classes.length
    ? activePreset.classes
    : [{ className: "", rawText: "" }];

  classesToRender.forEach((classData) => addClassCard(classData));
  applyLayoutLockToCards();
}

async function handleSnapshotImport(event) {
  const file = event.target.files?.[0];
  elements.snapshotFileInput.value = "";

  if (!file) {
    return;
  }

  try {
    const raw = await file.text();
    const parsed = JSON.parse(raw);
    const normalized = normalizeSnapshotImport(parsed);

    state.presets = normalized.presets;
    state.activePresetId = normalized.activePresetId;
    applySettings(normalized.settings);
    setCheckedValue("class-preset", state.activePresetId);
    renderActivePreset();
    syncPresetUi();
    persistState();
    updateSummary();
    showToast("복원 파일을 불러왔습니다. 반별 단어가 다시 채워졌습니다.");
  } catch (error) {
    console.error(error);
    setRunStatus(`복원 파일을 읽지 못했습니다. ${formatErrorDetail(error)}`, "error");
    showToast(`복원 파일 오류: ${formatErrorDetail(error)}`);
  }
}

function switchActivePreset(presetId) {
  if (presetId === state.activePresetId) {
    return;
  }

  captureActivePresetClassesFromDom();
  state.activePresetId = presetId;
  renderActivePreset();
  syncPresetUi();
  persistState();
  rememberLastUsedClassForCurrentAdmin();
  updateSummary();
}

function toggleActivePresetLock() {
  const activePreset = getActivePresetState();
  const shouldLock = !activePreset.layoutLocked;

  if (shouldLock) {
    const meaningfulClassCount = collectClassModels().filter(
      (model) => model.className.trim() || model.parsed.entries.length
    ).length;

    if (!meaningfulClassCount) {
      showToast("반 구성을 먼저 입력한 뒤 잠글 수 있습니다.");
      return;
    }
  }

  activePreset.layoutLocked = shouldLock;
  applyLayoutLockToCards();
  syncPresetUi();
  persistState();
  updateSummary();
  showToast(
    shouldLock
      ? "반 구성 잠금이 켜졌습니다. 이제 단어만 수정하면 됩니다."
      : "반 구성 편집이 열렸습니다. 반 이름과 개수를 바꿀 수 있습니다."
  );
}

function syncPresetUi() {
  if (!elements.toggleLayoutLockBtn || !elements.addClassBtn) {
    return;
  }

  const activePreset = getActivePresetState();

  if (activePreset.layoutLocked) {
    elements.toggleLayoutLockBtn.textContent = "반 구성 편집";
    elements.addClassBtn.disabled = true;
  } else {
    elements.toggleLayoutLockBtn.textContent = "반 구성 잠금";
    elements.addClassBtn.disabled = false;
  }
}

function applyLayoutLockToCards() {
  const locked = isActivePresetLocked();

  collectClassCards().forEach((card) => {
    const nameInput = card.querySelector(".class-name-input");
    const removeBtn = card.querySelector(".remove-class-btn");

    nameInput.readOnly = locked;
    nameInput.classList.toggle("is-readonly", locked);
    card.classList.toggle("layout-locked", locked);
    removeBtn.disabled = locked;
  });
}

function addClassCard(initialData = {}) {
  const fragment = elements.classTemplate.content.cloneNode(true);
  const card = fragment.querySelector(".class-card");
  const nameInput = card.querySelector(".class-name-input");
  const textArea = card.querySelector(".paste-area");
  const title = card.querySelector("h4");
  const removeBtn = card.querySelector(".remove-class-btn");

  const initialName = initialData.className || "";
  card.dataset.classId = String(initialData.classId || initialData.id || "");
  nameInput.value = initialName;
  textArea.value = initialData.rawText || "";
  title.textContent = initialName || "새 반";

  const syncCard = () => {
    title.textContent = nameInput.value.trim() || "새 반";
    updateClassCard(card);
    persistState();
    rememberLastUsedClassForCurrentAdmin(getCardClassReference(card));
    updateSummary();
  };

  nameInput.addEventListener("input", syncCard);
  textArea.addEventListener("input", syncCard);

  removeBtn.addEventListener("click", () => {
    if (isActivePresetLocked()) {
      showToast("반 구성 편집을 먼저 켜야 반을 삭제할 수 있습니다.");
      return;
    }

    if (collectClassCards().length === 1) {
      showToast("반은 최소 1개는 남겨 두어야 합니다.");
      return;
    }

    card.remove();
    persistState();
    updateSummary();
  });

  elements.classList.appendChild(card);
  updateClassCard(card);
  applyLayoutLockToCards();
}

function updateClassCard(card) {
  const rawText = card.querySelector(".paste-area").value;
  const parsed = parseVocabulary(rawText);
  const validCount = card.querySelector(".valid-count");
  const mergedCount = card.querySelector(".merged-count");
  const invalidCount = card.querySelector(".invalid-count");
  const previewList = card.querySelector(".preview-list");
  const previewNote = card.querySelector(".preview-note");

  card._parsed = parsed;

  validCount.textContent = parsed.entries.length;
  mergedCount.textContent = parsed.mergedCount;
  invalidCount.textContent = parsed.invalidLines.length;

  previewList.innerHTML = parsed.entries
    .slice(0, MAX_PREVIEW_ITEMS)
    .map(
      (entry) => `
        <div class="preview-item">
          <strong>${escapeHtml(entry.word)}</strong>
          <span>${escapeHtml(entry.meaning)}</span>
        </div>
      `
    )
    .join("");

  if (!parsed.entries.length) {
    previewNote.textContent =
      "탭으로 구분된 단어 / 뜻 형식으로 붙여넣으면 여기서 바로 확인됩니다.";
  } else {
    const notes = [];
    if (parsed.invalidLines.length) {
      notes.push(`형식이 맞지 않아 제외된 줄 ${parsed.invalidLines.length}개`);
    }
    if (parsed.mergedCount) {
      notes.push(`중복 단어는 뜻을 합쳐서 ${parsed.mergedCount}개 정리`);
    }
    previewNote.textContent = notes.length
      ? notes.join(" · ")
      : `${parsed.entries.length}개 단어가 준비되었습니다.`;
  }
}

function parseVocabulary(rawText) {
  const entryMap = new Map();
  const invalidLines = [];
  let mergedCount = 0;
  const normalizedText = rawText.replace(/\u00a0/g, " ");
  const rows = normalizedText.includes("\t")
    ? parseTabularRows(normalizedText)
    : normalizedText.split(/\r?\n/).map((line) => [line]);

  rows.forEach((row, index) => {
    const parsedRow = row.length >= 2
      ? [row[0], buildMeaningFromParts(row)]
      : splitVocabularyLine(String(row[0] || "").trim());

    if (!parsedRow) {
      const rawRow = row.map((cell) => String(cell || "")).join("\t").trim();
      if (rawRow) {
        invalidLines.push({ lineNumber: index + 1, content: rawRow });
      }
      return;
    }

    let [word, meaning] = parsedRow;
    word = normalizeWordCell(stripWrappingQuotes(word));
    meaning = normalizeMeaningCell(stripWrappingQuotes(meaning));

    if (!word || !meaning) {
      invalidLines.push({
        lineNumber: index + 1,
        content: row.map((cell) => String(cell || "")).join("\t"),
      });
      return;
    }

    const key = word.toLocaleLowerCase();
    const existing = entryMap.get(key);

    if (existing) {
      if (!existing.meaningSet.has(meaning)) {
        existing.meaningSet.add(meaning);
      }
      mergedCount += 1;
      return;
    }

    entryMap.set(key, {
      id: `${key}-${index}`,
      word,
      meaningSet: new Set([meaning]),
    });
  });

  const entries = Array.from(entryMap.values()).map((entry) => ({
    id: entry.id,
    word: entry.word,
    meaning: Array.from(entry.meaningSet).join(" / "),
  }));

  return {
    entries,
    mergedCount,
    invalidLines,
  };
}

function parseTabularRows(rawText) {
  const text = rawText.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];

    if (char === "\"") {
      const nextChar = text[index + 1];

      if (inQuotes && nextChar === "\"") {
        cell += "\"";
        index += 1;
        continue;
      }

      inQuotes = !inQuotes;
      continue;
    }

    if (char === "\t" && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if (char === "\n" && !inQuotes) {
      row.push(cell);
      if (row.some((value) => String(value).trim())) {
        rows.push(row);
      }
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  row.push(cell);
  if (row.some((value) => String(value).trim())) {
    rows.push(row);
  }

  return rows;
}

function splitVocabularyLine(line) {
  if (line.includes("\t")) {
    const parts = line
      .split("\t")
      .map((part) => part.trim())
      .filter(Boolean);
    if (parts.length >= 2) {
      return [parts[0], buildMeaningFromParts(parts)];
    }
  }

  if (line.includes("|")) {
    const parts = line
      .split("|")
      .map((part) => part.trim())
      .filter(Boolean);
    if (parts.length >= 2) {
      return [parts[0], buildMeaningFromParts(parts)];
    }
  }

  if (line.includes(",")) {
    const parts = line
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean);
    if (parts.length >= 2) {
      return [parts[0], buildMeaningFromParts(parts, ", ")];
    }
  }

  const chunks = line
    .split(/\s{2,}/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (chunks.length >= 2) {
    return [chunks[0], buildMeaningFromParts(chunks)];
  }

  return null;
}

function buildMeaningFromParts(parts, separator = " / ") {
  const meanings = [];

  for (let index = 1; index < parts.length; index += 1) {
    const part = normalizeCell(stripWrappingQuotes(parts[index]));

    if (!part) {
      continue;
    }

    if (meanings.length && looksLikeExampleText(part)) {
      break;
    }

    meanings.push(part);
  }

  return normalizeMeaningCell(meanings.join(separator));
}

function normalizeCell(text) {
  return text.replace(/\s+/g, " ").trim();
}

function normalizeWordCell(text) {
  let value = normalizeCell(text);
  value = value.replace(/^[\[(][\u3131-\u318e\uac00-\ud7a3\s]+[\])]\s*/g, "");
  value = value.replace(/\s*[\[(][\u3131-\u318e\uac00-\ud7a3\s]+[\])]\s*$/g, "");
  return value.trim();
}

function normalizeMeaningCell(text) {
  let value = normalizeCell(text);

  if (!value) {
    return "";
  }

  const exampleSplitMatch = value.match(/^(.*?)(?:\s+)([A-Za-z].*)$/);
  if (exampleSplitMatch) {
    const prefix = normalizeCell(exampleSplitMatch[1]);
    const suffix = normalizeCell(exampleSplitMatch[2]);

    if (containsKorean(prefix) && looksLikeExampleText(suffix)) {
      value = prefix;
    }
  }

  return value.trim();
}

function looksLikeExampleText(text) {
  const value = normalizeCell(text);

  if (!value || !/[A-Za-z]/.test(value)) {
    return false;
  }

  if (/[{}]/.test(value)) {
    return true;
  }

  const englishWords = value.match(/[A-Za-z]+(?:'[A-Za-z]+)?/g) || [];
  if (englishWords.length < 3) {
    return false;
  }

  if (/[.!?]$/.test(value)) {
    return true;
  }

  return /^(?:I|We|You|He|She|They|It|This|That|These|Those|The|A|An|Please|My|Your|His|Her|Our|Their|Parents|Teachers|Students|People|Someone|Something)\b/.test(
    value
  );
}

function containsKorean(text) {
  return /[\u3131-\u318e\uac00-\ud7a3]/.test(String(text || ""));
}

function stripWrappingQuotes(text) {
  const value = String(text || "").trim();
  if (value.length >= 2 && value.startsWith("\"") && value.endsWith("\"")) {
    return value.slice(1, -1).trim();
  }
  return value;
}

function collectClassCards() {
  return Array.from(elements.classList.querySelectorAll(".class-card"));
}

function collectClassModels() {
  return collectClassCards().map((card) => ({
    className: card.querySelector(".class-name-input").value.trim(),
    parsed: card._parsed || parseVocabulary(card.querySelector(".paste-area").value),
  }));
}

function filterMeaningfulClassModels(classModels) {
  return classModels.filter(
    (model) => model.className.trim() || model.parsed.entries.length
  );
}

function filterPrintableClassModels(classModels) {
  return filterMeaningfulClassModels(classModels).filter(
    (model) => model.className.trim() && model.parsed.entries.length > 0
  );
}

function getEffectiveQuestionCount(model, settings) {
  return Math.max(
    0,
    Math.min(settings.questionCount, model.parsed.entries.length)
  );
}

function updateSummary() {
  const settings = getSettings();
  const classModels = filterMeaningfulClassModels(collectClassModels());
  const printableClassModels = filterPrintableClassModels(classModels);
  const totalWords = printableClassModels.reduce(
    (sum, model) => sum + model.parsed.entries.length,
    0
  );

  if (elements.metricClassCount) {
    elements.metricClassCount.textContent = String(printableClassModels.length);
  }

  if (elements.metricWordCount) {
    elements.metricWordCount.textContent = String(totalWords);
  }

  if (elements.metricPdfCount) {
    elements.metricPdfCount.textContent = String(
      printableClassModels.length * settings.sessionCount * 2
    );
  }

  setElementTextAndTitle(elements.metricFolderName, state.saveRootName || "미연결");

  if (!classModels.length) {
    setValidationNotes([]);
    return;
  }

  const messages = validateClassModels(classModels, settings, {
    includeSuccess: true,
  });
  setValidationNotes(messages);
}

function validateClassModels(classModels, settings, options = {}) {
  const meaningfulModels = filterMeaningfulClassModels(classModels);
  const messages = [];
  const names = new Map();

  if (!Number.isFinite(settings.questionCount) || settings.questionCount < 5) {
    messages.push({ type: "error", text: "문항 수는 5 이상이어야 합니다." });
  }

  meaningfulModels.forEach((model, index) => {
    const label = model.className || `${index + 1}번 반`;
    const folderKey = sanitizeFilePart(
      model.className || `class-${index + 1}`
    ).toLocaleLowerCase();

    if (!model.className.trim()) {
      messages.push({ type: "error", text: `${label}: 반 이름이 비어 있습니다.` });
      return;
    }

    if (!model.parsed.entries.length) {
      messages.push({
        type: "warning",
        text: `${label}: 유효한 단어가 없어 이번 생성에서는 건너뜁니다.`,
      });
      return;
    }

    if (names.has(folderKey)) {
      messages.push({
        type: "error",
        text: `${label}: 저장 폴더명이 겹치는 반이 있습니다.`,
      });
      return;
    }
    names.set(folderKey, true);

    const questionCount = getEffectiveQuestionCount(model, settings);

    if (questionCount < settings.questionCount) {
      messages.push({
        type: "warning",
        text: `${label}: 단어가 ${model.parsed.entries.length}개이므로 문항은 ${questionCount}개만 출력됩니다.`,
      });
    }

    const totalNeeded = questionCount * settings.sessionCount;
    if (model.parsed.entries.length < totalNeeded) {
      messages.push({
        type: "warning",
        text: `${label}: 차수 간 완전 무중복은 불가능합니다. ${model.parsed.entries.length}개 단어로 ${settings.sessionCount}차 x ${questionCount}문항을 만들면 총 ${totalNeeded}칸이 필요합니다.`,
      });
    } else if (options.includeSuccess && questionCount) {
      messages.push({
        type: "success",
        text: `${label}: ${settings.sessionCount}차까지 ${questionCount}문항 기준 완전 무중복 랜덤 구성이 가능합니다.`,
      });
    }
  });

  if (!meaningfulModels.length) {
    messages.push({ type: "error", text: "반 이름과 단어를 하나 이상 입력해 주세요." });
  }

  return messages;
}

function setValidationNotes(messages) {
  if (!elements.validationNotes) {
    return;
  }

  if (!messages.length) {
    elements.validationNotes.className = "status-panel info";
    elements.validationNotes.textContent =
      "반과 단어를 입력하면 여기서 검사가 표시됩니다.";
    return;
  }

  const hasError = messages.some((message) => message.type === "error");
  const hasWarning = messages.some((message) => message.type === "warning");

  elements.validationNotes.className = hasError
    ? "status-panel"
    : hasWarning
      ? "status-panel info"
      : "status-panel success";

  elements.validationNotes.innerHTML = messages
    .map((message) => `<div>${escapeHtml(prefixMessage(message))}</div>`)
    .join("");
}

function prefixMessage(message) {
  if (message.type === "error") {
    return `오류: ${message.text}`;
  }
  if (message.type === "warning") {
    return `주의: ${message.text}`;
  }
  return `확인: ${message.text}`;
}

async function handlePickFolderClick() {
  if (hasNativeHost()) {
    await chooseNativeSaveRoot();
    return;
  }

  if (!window.showDirectoryPicker) {
    showToast(
      "이 브라우저는 폴더 직접 저장을 지원하지 않습니다. Edge 또는 Chrome을 사용해 주세요."
    );
    return;
  }

  if (state.rememberedSaveRootHandle && !state.saveRootHandle) {
    const restored = await restoreRememberedSaveRoot(true);
    if (restored) {
      showToast("이전 저장 폴더 연결이 복원되었습니다.");
      return;
    }
  }

  await chooseSaveRoot();
}

async function chooseNativeSaveRoot() {
  try {
    const status = await callNativeHost("chooseSaveRoot");
    applyNativeSaveRootStatus(status);
    syncSaveRootUi();
    updateSummary();

    if (state.saveRootHandle) {
      showToast("저장 폴더 연결이 완료되었습니다.");
    }
  } catch (error) {
    console.error(error);
    showToast(`폴더 연결 중 오류가 발생했습니다. ${formatErrorDetail(error)}`);
  }
}

async function chooseSaveRoot() {
  if (!window.showDirectoryPicker) {
    showToast(
      "이 브라우저는 폴더 직접 저장을 지원하지 않습니다. Edge 또는 Chrome을 사용해 주세요."
    );
    return;
  }

  try {
    const handle = await window.showDirectoryPicker({
      mode: "readwrite",
      id: "word-lab-output",
      startIn: "desktop",
    });
    await activateSaveRootHandle(handle);
    updateSummary();
    showToast("저장 폴더 연결이 완료되었습니다.");
  } catch (error) {
    if (error) {
      console.error(error);
    }

    if (error && error.name === "AbortError") {
      showToast("Desktop 자체는 막힐 수 있습니다. 바탕화면 안의 일반 폴더를 새로 만들어 선택해 주세요.");
      return;
    }

    showToast("폴더 연결 중 오류가 발생했습니다.");
  }
}

async function activateSaveRootHandle(handle) {
  state.saveRootHandle = handle;
  state.rememberedSaveRootHandle = handle;
  state.saveRootName = handle.name || "";
  await persistSaveRootHandle(handle);
  syncSaveRootUi();
}

async function restoreSaveRootHandle() {
  if (hasNativeHost()) {
    try {
      const status = await callNativeHost("getSaveRootStatus");
      applyNativeSaveRootStatus(status);
    } catch (error) {
      console.error(error);
      state.saveRootHandle = null;
      state.rememberedSaveRootHandle = null;
      state.saveRootName = "";
    }

    syncSaveRootUi();
    return;
  }

  if (!("indexedDB" in window)) {
    return;
  }

  try {
    const handle = await getStoredSaveRootHandle();
    if (!handle) {
      return;
    }

    state.rememberedSaveRootHandle = handle;
    state.saveRootName = handle.name || "";

    const permission = await verifyDirectoryPermission(handle, false);
    if (permission === "granted") {
      state.saveRootHandle = handle;
    }
  } catch (error) {
    console.error(error);
    await clearStoredSaveRootHandle();
    state.saveRootHandle = null;
    state.rememberedSaveRootHandle = null;
    state.saveRootName = "";
  }

  syncSaveRootUi();
}

async function restoreRememberedSaveRoot(requestIfNeeded) {
  if (!state.rememberedSaveRootHandle) {
    return false;
  }

  try {
    const permission = await verifyDirectoryPermission(
      state.rememberedSaveRootHandle,
      requestIfNeeded
    );

    if (permission !== "granted") {
      syncSaveRootUi();
      return false;
    }

    state.saveRootHandle = state.rememberedSaveRootHandle;
    state.saveRootName = state.rememberedSaveRootHandle.name || state.saveRootName;
    syncSaveRootUi();
    updateSummary();
    return true;
  } catch (error) {
    console.error(error);
    await clearStoredSaveRootHandle();
    state.saveRootHandle = null;
    state.rememberedSaveRootHandle = null;
    state.saveRootName = "";
    syncSaveRootUi();
    updateSummary();
    return false;
  }
}

function syncSaveRootUi() {
  if (!elements.saveRootStatus || !elements.pickFolderBtn) {
    return;
  }

  const disconnectedLabel = "\uBBF8\uC5F0\uACB0";
  const folderLabel = state.saveRootName || disconnectedLabel;
  const compactFolderLabel = state.saveRootName
    ? getCompactSaveRootLabel(state.saveRootName)
    : disconnectedLabel;

  if (hasNativeHost()) {
    if (state.saveRootHandle) {
      setElementTextAndTitle(elements.saveRootStatus, compactFolderLabel, folderLabel);
      elements.pickFolderBtn.textContent = "\uC800\uC7A5 \uD3F4\uB354 \uBCC0\uACBD";
      return;
    }

    setElementTextAndTitle(elements.saveRootStatus, disconnectedLabel);
    elements.pickFolderBtn.textContent = "\uC800\uC7A5 \uD3F4\uB354 \uC5F0\uACB0";
    return;
  }

  if (state.saveRootHandle) {
    setElementTextAndTitle(elements.saveRootStatus, compactFolderLabel, folderLabel);
    elements.pickFolderBtn.textContent = "\uC800\uC7A5 \uD3F4\uB354 \uBCC0\uACBD";
    return;
  }

  if (state.rememberedSaveRootHandle && state.saveRootName) {
    setElementTextAndTitle(
      elements.saveRootStatus,
      compactFolderLabel,
      state.saveRootName
    );
    elements.pickFolderBtn.textContent = "\uC800\uC7A5 \uD3F4\uB354 \uB2E4\uC2DC \uC5F0\uACB0";
    return;
  }

  setElementTextAndTitle(elements.saveRootStatus, disconnectedLabel);
  elements.pickFolderBtn.textContent = "\uC800\uC7A5 \uD3F4\uB354 \uC5F0\uACB0";
}

async function verifyDirectoryPermission(handle, requestIfNeeded) {
  const options = { mode: "readwrite" };
  const currentPermission = await handle.queryPermission(options);

  if (currentPermission === "granted") {
    return "granted";
  }

  if (!requestIfNeeded) {
    return currentPermission;
  }

  return handle.requestPermission(options);
}

function openHandleDatabase() {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(HANDLE_DB_NAME, 1);

    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(HANDLE_STORE_NAME)) {
        request.result.createObjectStore(HANDLE_STORE_NAME);
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getStoredSaveRootHandle() {
  const database = await openHandleDatabase();

  try {
    return await new Promise((resolve, reject) => {
      const transaction = database.transaction(HANDLE_STORE_NAME, "readonly");
      const request = transaction.objectStore(HANDLE_STORE_NAME).get(SAVE_ROOT_HANDLE_KEY);

      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error);
    });
  } finally {
    database.close();
  }
}

async function persistSaveRootHandle(handle) {
  const database = await openHandleDatabase();

  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(HANDLE_STORE_NAME, "readwrite");
      const request = transaction.objectStore(HANDLE_STORE_NAME).put(handle, SAVE_ROOT_HANDLE_KEY);

      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || request.error);
      transaction.onabort = () => reject(transaction.error || request.error);
    });
  } finally {
    database.close();
  }
}

async function clearStoredSaveRootHandle() {
  if (!("indexedDB" in window)) {
    return;
  }

  const database = await openHandleDatabase();

  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(HANDLE_STORE_NAME, "readwrite");
      const request = transaction.objectStore(HANDLE_STORE_NAME).delete(SAVE_ROOT_HANDLE_KEY);

      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || request.error);
      transaction.onabort = () => reject(transaction.error || request.error);
    });
  } finally {
    database.close();
  }
}

async function handleGenerate() {
  if (state.isGenerating) {
    return;
  }

  const settings = getSettings();
  const allClassModels = collectClassModels();
  const classModels = filterMeaningfulClassModels(allClassModels);
  const printableClassModels = filterPrintableClassModels(classModels);
  const validationMessages = validateClassModels(classModels, settings);
  const blockingErrors = validationMessages.filter(
    (message) => message.type === "error"
  );

  setValidationNotes(validationMessages);

  if (blockingErrors.length) {
    setRunStatus("PDF 생성 실패", "error");
    showToast("입력 오류가 있어 생성할 수 없습니다.");
    return;
  }

  if (!printableClassModels.length) {
    setRunStatus("PDF 생성 실패", "error");
    showToast("생성할 단어가 있는 반이 없습니다.");
    return;
  }

  if (!(await ensureWritableSaveRootHandle())) {
    setRunStatus("PDF 생성 실패", "error");
    return;
  }

  setGenerating(true);
  const now = new Date();
  const dateFolderName = formatDateFolder(now);
  const dateLabel = formatDisplayDate(now);
  const warnings = validationMessages
    .filter((message) => message.type === "warning")
    .map((message) => message.text);

  try {
    const generationTarget = await createGenerationTarget(dateFolderName);
    let completedFiles = 0;
    const totalFiles = printableClassModels.length * settings.sessionCount * 2;
    let zipSaved = false;
    const zipEntries = [];

    for (const model of printableClassModels) {
      const safeClassName = sanitizeFilePart(model.className);
      const questionCount = getEffectiveQuestionCount(model, settings);
      const sessionBundles = buildSessionBundles(
        model.parsed.entries,
        settings.sessionCount,
        questionCount
      );

      if (!sessionBundles.fullyUnique) {
        warnings.push(`${model.className}: 차수 간 일부 단어 중복이 포함되었습니다.`);
      }

      for (let index = 0; index < sessionBundles.sessions.length; index += 1) {
        const sessionNumber = index + 1;
        setRunStatus("PDF 생성 중...", "running");
        const items = sessionBundles.sessions[index];

        const page = buildSheetPage({
          className: model.className,
          sessionNumber,
          items,
          quizMode: settings.quizMode,
          dateLabel,
        });

        const pdfBuffer = await renderPageToPdfBuffer(
          page,
          `입시코드학원 ${model.className} ${sessionNumber}차 단어시험지`
        );
        const fileName = `${safeClassName}_${sessionNumber}차.pdf`;
        await writeGeneratedPdfFile(generationTarget, safeClassName, fileName, pdfBuffer);
        zipEntries.push({
          path: `${dateFolderName}/${safeClassName}/${fileName}`,
          data: pdfBuffer,
        });
        completedFiles += 1;

        setRunStatus("PDF 생성 중...", "running");

        const answerPage = buildAnswerSheetPage({
          className: model.className,
          sessionNumber,
          items,
          quizMode: settings.quizMode,
          dateLabel,
        });

        const answerPdfBuffer = await renderPageToPdfBuffer(
          answerPage,
          `입시코드학원 ${model.className} ${sessionNumber}차 답지`
        );
        const answerFileName = `${safeClassName}_${sessionNumber}차_답지.pdf`;
        await writeGeneratedPdfFile(
          generationTarget,
          safeClassName,
          answerFileName,
          answerPdfBuffer
        );
        zipEntries.push({
          path: `${dateFolderName}/${safeClassName}/${answerFileName}`,
          data: answerPdfBuffer,
        });
        completedFiles += 1;
      }
    }

    try {
      const snapshotPayload = buildSnapshotPayload(dateFolderName);
      const snapshotText = JSON.stringify(snapshotPayload, null, 2);
      await writeGeneratedTextFile(generationTarget, SNAPSHOT_FILE_NAME, snapshotText);
      zipEntries.push({
        path: `${dateFolderName}/${SNAPSHOT_FILE_NAME}`,
        data: snapshotText,
      });
    } catch (snapshotError) {
      console.error(snapshotError);
      warnings.push(`복원 파일 저장 실패: ${formatErrorDetail(snapshotError)}`);
    }

    try {
      const zipBuffer = createZipArchive(zipEntries);
      await writeGeneratedZipFile(generationTarget, `${dateFolderName}.zip`, zipBuffer);
      zipSaved = true;
    } catch (zipError) {
      console.error(zipError);
      warnings.push(`ZIP 저장 실패: ${formatErrorDetail(zipError)}`);
    }

    setRunStatus("PDF 생성 완료", "success");
    showToast(
      zipSaved
        ? `PDF ${totalFiles}개와 ZIP 1개 생성이 끝났습니다.`
        : `PDF ${totalFiles}개 생성이 끝났습니다. ZIP 생성은 실패했습니다.`
    );
  } catch (error) {
    console.error(error);
    if (isSaveRootAccessError(error)) {
      state.saveRootHandle = null;
      syncSaveRootUi();
      updateSummary();
      setRunStatus("PDF 생성 실패", "error");
      showToast("저장 폴더 권한이 끊겨 PDF를 저장하지 못했습니다.");
      return;
    }

    setRunStatus("PDF 생성 실패", "error");
    showToast(`PDF 생성 중 오류: ${formatErrorDetail(error)}`);
  } finally {
    elements.renderStage.innerHTML = "";
    setGenerating(false);
  }
}

async function ensureWritableSaveRootHandle() {
  if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
    setRunStatus("PDF 라이브러리를 불러오지 못했습니다. 페이지를 새로고침해 주세요.", "error");
    showToast("PDF 라이브러리를 불러오지 못했습니다.");
    return false;
  }

  if (hasNativeHost()) {
    try {
      const status = await callNativeHost("getSaveRootStatus");
      applyNativeSaveRootStatus(status);
      syncSaveRootUi();
      updateSummary();

      if (!state.saveRootHandle) {
        await chooseNativeSaveRoot();
      }

      return Boolean(state.saveRootHandle);
    } catch (error) {
      console.error(error);
      setRunStatus(
        `저장 폴더 확인 중 오류가 발생했습니다. ${formatErrorDetail(error)}`,
        "error"
      );
      showToast("저장 폴더 확인 중 오류가 발생했습니다.");
      return false;
    }
  }

  if (!state.saveRootHandle && state.rememberedSaveRootHandle) {
    await restoreRememberedSaveRoot(true);
  }

  if (!state.saveRootHandle) {
    await chooseSaveRoot();
    if (!state.saveRootHandle) {
      return false;
    }
  }

  try {
    const permission = await verifyDirectoryPermission(state.saveRootHandle, true);
    if (permission !== "granted") {
      state.saveRootHandle = null;
      syncSaveRootUi();
      updateSummary();
      showToast("저장 폴더 권한이 없어 다시 연결이 필요합니다.");
      return false;
    }
  } catch (error) {
    console.error(error);
    state.saveRootHandle = null;
    syncSaveRootUi();
    updateSummary();
    showToast("저장 폴더 권한 확인 중 오류가 발생했습니다.");
    return false;
  }

  return true;
}

function isSaveRootAccessError(error) {
  const name = error?.name || "";
  return [
    "NotAllowedError",
    "SecurityError",
    "InvalidStateError",
    "NotFoundError",
    "NoModificationAllowedError",
  ].includes(name);
}

function formatErrorDetail(error) {
  const name = error?.name ? String(error.name) : "Error";
  const message = error?.message ? String(error.message).trim() : "";
  return message ? `${name}: ${message}` : name;
}

function buildSessionBundles(entries, sessionCount, questionCount) {
  if (entries.length >= sessionCount * questionCount) {
    const pool = shuffle(entries.map(cloneEntry));
    const sessions = [];

    for (let index = 0; index < sessionCount; index += 1) {
      const slice = pool.slice(index * questionCount, (index + 1) * questionCount);
      sessions.push(shuffle(slice.map(cloneEntry)));
    }

    return { sessions, fullyUnique: true };
  }

  const usage = new Map(entries.map((entry) => [entry.id, 0]));
  const signatures = new Set();
  const sessions = [];

  for (let index = 0; index < sessionCount; index += 1) {
    const selection = pickBalancedSession(
      entries,
      questionCount,
      usage,
      signatures
    );
    selection.forEach((entry) => {
      usage.set(entry.id, usage.get(entry.id) + 1);
    });
    sessions.push(selection.map(cloneEntry));
  }

  return { sessions, fullyUnique: false };
}

function pickBalancedSession(entries, questionCount, usage, signatures) {
  let best = null;

  for (let attempt = 0; attempt < 40; attempt += 1) {
    const ranked = shuffle(entries.map(cloneEntry))
      .map((entry) => ({
        entry,
        usage: usage.get(entry.id) || 0,
        tie: Math.random(),
      }))
      .sort((left, right) => {
        if (left.usage !== right.usage) {
          return left.usage - right.usage;
        }
        return left.tie - right.tie;
      });

    const selection = ranked.slice(0, questionCount).map((item) => item.entry);
    const signature = selection.map((entry) => entry.id).sort().join("|");
    const totalUsage = selection.reduce(
      (sum, entry) => sum + (usage.get(entry.id) || 0),
      0
    );
    const maxUsage = selection.reduce(
      (max, entry) => Math.max(max, usage.get(entry.id) || 0),
      0
    );
    const penalty = signatures.has(signature) ? 10000 : 0;
    const score = penalty + maxUsage * 100 + totalUsage;

    if (!best || score < best.score) {
      best = {
        selection: shuffle(selection.map(cloneEntry)),
        signature,
        score,
      };
    }

    if (!signatures.has(signature)) {
      break;
    }
  }

  signatures.add(best.signature);
  return best.selection;
}

function buildSheetPage({
  className,
  sessionNumber,
  items,
  quizMode,
  dateLabel,
}) {
  return buildPrintablePage({
    className,
    sessionNumber,
    items,
    quizMode,
    dateLabel,
    answerKey: false,
  });
}

function buildAnswerSheetPage({
  className,
  sessionNumber,
  items,
  quizMode,
  dateLabel,
}) {
  return buildPrintablePage({
    className,
    sessionNumber,
    items,
    quizMode,
    dateLabel,
    answerKey: true,
  });
}

function buildPrintablePage({
  className,
  sessionNumber,
  items,
  quizMode,
  dateLabel,
  answerKey,
}) {
  const columnCount = 4;
  const columns = splitIntoColumns(items, columnCount);
  let offset = 0;
  const columnMarkup = columns.map((column) => {
    const markup = renderSheetColumn(column, offset, quizMode, { answerKey });
    offset += column.length;
    return markup;
  });
  const title = answerKey ? "단어 답지" : "단어 시험지";
  const sideCardMarkup = answerKey
    ? `
        <div class="sheet-key-card">
          <span>답지</span>
          <strong>정답 확인용</strong>
        </div>
      `
    : `
        <div class="sheet-score-card">
          <span>점수</span>
          <div class="sheet-score-line"></div>
        </div>
      `;
  const metaCardMarkup = answerKey
    ? `
        <div class="sheet-info-card">
          <span>형태</span>
          <strong>정답지</strong>
        </div>
      `
    : `
        <div class="sheet-info-card">
          <span>이름</span>
          <div class="sheet-name-line"></div>
        </div>
      `;
  const page = document.createElement("section");
  page.className = "sheet-page";
  page.innerHTML = `
    <header class="sheet-header">
      <div class="sheet-header-copy">
        <div class="sheet-brand">입시코드학원</div>
        <div class="sheet-title">${title}</div>
      </div>
      <div class="sheet-header-side">
        <div class="sheet-badge-stack">
          <span class="sheet-badge sheet-class-badge">${escapeHtml(className)}</span>
          <span class="sheet-badge sheet-session-badge">${sessionNumber}차</span>
        </div>
        ${sideCardMarkup}
      </div>
    </header>

    <section class="sheet-info">
      <div class="sheet-info-card">
        <span>날짜</span>
        <strong>${escapeHtml(dateLabel)}</strong>
      </div>
      <div class="sheet-info-card">
        <span>출제</span>
        <strong>${escapeHtml(
          quizMode === "word-to-meaning" ? "영단어 -> 뜻" : "뜻 -> 영단어"
        )}</strong>
      </div>
      <div class="sheet-info-card">
        <span>문항 수</span>
        <strong>${items.length}문항</strong>
      </div>
      ${metaCardMarkup}
    </section>

    <section class="sheet-grid sheet-grid-4">
      ${columnMarkup.join("")}
    </section>

    <footer class="sheet-footer">
      <span>입시코드학원</span>
      <span>${escapeHtml(className)} · ${sessionNumber}차</span>
    </footer>
  `;

  elements.renderStage.innerHTML = "";
  elements.renderStage.appendChild(page);
  return page;
}

function renderSheetColumn(entries, baseIndex, quizMode, options = {}) {
  const { answerKey = false } = options;
  return `
    <div class="sheet-column">
      ${entries
        .map((entry, index) => {
          const { prompt, answer } = getPromptAndAnswer(entry, quizMode);
          const answerMarkup = answerKey
            ? `<div class="sheet-answer-text">${escapeHtml(answer)}</div>`
            : `<div class="sheet-answer-line"></div>`;
          return `
            <div class="sheet-row">
              <div class="sheet-row-top">
                <span class="sheet-num">${baseIndex + index + 1}.</span>
                <span class="sheet-prompt">${escapeHtml(prompt)}</span>
              </div>
              ${answerMarkup}
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function getPromptAndAnswer(entry, quizMode) {
  if (quizMode === "word-to-meaning") {
    return {
      prompt: entry.word,
      answer: entry.meaning,
    };
  }

  return {
    prompt: entry.meaning,
    answer: entry.word,
  };
}

function splitIntoColumns(items, columnCount) {
  if (columnCount === 1) {
    return [items];
  }

  const columns = Array.from({ length: columnCount }, () => []);
  const chunkSize = Math.ceil(items.length / columnCount);

  for (let index = 0; index < items.length; index += 1) {
    const columnIndex = Math.floor(index / chunkSize);
    columns[columnIndex].push(items[index]);
  }

  return columns.filter((column) => column.length);
}

async function createGenerationTarget(dateFolderName) {
  if (hasNativeHost()) {
    return {
      mode: "native",
      dateFolderName,
    };
  }

  return {
    mode: "browser",
    saveRootFolder: state.saveRootHandle,
    dateFolder: await ensureDirectory(state.saveRootHandle, dateFolderName),
    classFolders: new Map(),
  };
}

async function writeGeneratedPdfFile(target, classFolderName, fileName, pdfBuffer) {
  if (target.mode === "native") {
    const base64Data = await arrayBufferToBase64(pdfBuffer);
    await callNativeHost("savePdfFile", {
      dateFolderName: target.dateFolderName,
      classFolderName,
      fileName,
      base64Data,
    });
    return;
  }

  let classFolder = target.classFolders.get(classFolderName);
  if (!classFolder) {
    classFolder = await ensureDirectory(target.dateFolder, classFolderName);
    target.classFolders.set(classFolderName, classFolder);
  }

  await writePdfFile(classFolder, fileName, pdfBuffer);
}

async function writeGeneratedTextFile(target, fileName, text) {
  if (target.mode === "native") {
    await callNativeHost("saveTextFile", {
      dateFolderName: target.dateFolderName,
      fileName,
      text,
    });
    return;
  }

  await writeTextFile(target.dateFolder, fileName, text);
}

async function writeGeneratedZipFile(target, fileName, zipBuffer) {
  if (target.mode === "native") {
    const base64Data = await arrayBufferToBase64(zipBuffer, "application/zip");
    await callNativeHost("saveBinaryFile", {
      dateFolderName: target.dateFolderName,
      fileName,
      base64Data,
    });
    return;
  }

  await writeBinaryFile(target.saveRootFolder, fileName, zipBuffer);
}

function arrayBufferToBase64(buffer, mimeType = "application/octet-stream") {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || "");
      const [, base64 = ""] = dataUrl.split(",", 2);
      resolve(base64);
    };
    reader.onerror = () => reject(reader.error || new Error("파일 인코딩에 실패했습니다."));
    reader.readAsDataURL(new Blob([buffer], { type: mimeType }));
  });
}

async function renderPageToPdfBuffer(page, documentTitle) {
  await waitForLayout();

  const canvas = await window.html2canvas(page, {
    scale: 2,
    backgroundColor: "#f5f2eb",
    logging: false,
    useCORS: true,
  });

  const imageData = canvas.toDataURL("image/jpeg", 0.96);
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
    compress: true,
  });

  if (documentTitle) {
    pdf.setProperties({ title: documentTitle });
  }

  pdf.addImage(imageData, "JPEG", 0, 0, 210, 297, undefined, "FAST");
  const buffer = pdf.output("arraybuffer");

  canvas.width = 0;
  canvas.height = 0;

  return buffer;
}

async function writePdfFile(folderHandle, fileName, pdfBuffer) {
  await writeBinaryFile(folderHandle, fileName, pdfBuffer);
}

async function writeBinaryFile(folderHandle, fileName, binaryBuffer) {
  const fileHandle = await folderHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(binaryBuffer);
  await writable.close();
}

async function writeTextFile(folderHandle, fileName, text) {
  const fileHandle = await folderHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(text);
  await writable.close();
}

async function ensureDirectory(parentHandle, folderName) {
  return parentHandle.getDirectoryHandle(folderName, { create: true });
}

function createZipArchive(files) {
  const normalizedEntries = files.map((file) => {
    const path = normalizeZipPath(file.path);
    const nameBytes = new TextEncoder().encode(path);
    const dataBytes = toUint8Array(file.data);
    const crc32 = computeCrc32(dataBytes);
    const modifiedAt = file.modifiedAt instanceof Date ? file.modifiedAt : new Date();
    const { dosDate, dosTime } = getDosDateTime(modifiedAt);

    return {
      path,
      nameBytes,
      dataBytes,
      crc32,
      dosDate,
      dosTime,
    };
  });

  let localSize = 0;
  let centralSize = 0;
  const localRecords = [];
  const centralRecords = [];

  normalizedEntries.forEach((entry) => {
    const localOffset = localSize;
    const localRecord = new Uint8Array(30 + entry.nameBytes.length + entry.dataBytes.length);
    const localView = new DataView(localRecord.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, entry.dosTime, true);
    localView.setUint16(12, entry.dosDate, true);
    localView.setUint32(14, entry.crc32, true);
    localView.setUint32(18, entry.dataBytes.length, true);
    localView.setUint32(22, entry.dataBytes.length, true);
    localView.setUint16(26, entry.nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localRecord.set(entry.nameBytes, 30);
    localRecord.set(entry.dataBytes, 30 + entry.nameBytes.length);
    localRecords.push(localRecord);
    localSize += localRecord.length;

    const centralRecord = new Uint8Array(46 + entry.nameBytes.length);
    const centralView = new DataView(centralRecord.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, entry.dosTime, true);
    centralView.setUint16(14, entry.dosDate, true);
    centralView.setUint32(16, entry.crc32, true);
    centralView.setUint32(20, entry.dataBytes.length, true);
    centralView.setUint32(24, entry.dataBytes.length, true);
    centralView.setUint16(28, entry.nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, localOffset, true);
    centralRecord.set(entry.nameBytes, 46);
    centralRecords.push(centralRecord);
    centralSize += centralRecord.length;
  });

  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, normalizedEntries.length, true);
  endView.setUint16(10, normalizedEntries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, localSize, true);
  endView.setUint16(20, 0, true);

  const zipBuffer = new Uint8Array(localSize + centralSize + endRecord.length);
  let cursor = 0;

  localRecords.forEach((record) => {
    zipBuffer.set(record, cursor);
    cursor += record.length;
  });

  centralRecords.forEach((record) => {
    zipBuffer.set(record, cursor);
    cursor += record.length;
  });

  zipBuffer.set(endRecord, cursor);
  return zipBuffer.buffer;
}

function normalizeZipPath(path) {
  return String(path || "")
    .replace(/\\/g, "/")
    .split("/")
    .filter(Boolean)
    .join("/");
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) {
    return data;
  }

  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }

  if (typeof data === "string") {
    return new TextEncoder().encode(data);
  }

  throw new Error("ZIP 파일 데이터 형식을 처리할 수 없습니다.");
}

function computeCrc32(bytes) {
  const table = getCrc32Table();
  let crc = 0xffffffff;

  for (let index = 0; index < bytes.length; index += 1) {
    crc = table[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }

  return (crc ^ 0xffffffff) >>> 0;
}

let crc32Table = null;

function getCrc32Table() {
  if (crc32Table) {
    return crc32Table;
  }

  crc32Table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let shift = 0; shift < 8; shift += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    crc32Table[index] = value >>> 0;
  }

  return crc32Table;
}

function getDosDateTime(date) {
  const year = Math.max(1980, Math.min(2107, date.getFullYear()));
  const month = Math.max(1, date.getMonth() + 1);
  const day = Math.max(1, date.getDate());
  const hours = Math.max(0, Math.min(23, date.getHours()));
  const minutes = Math.max(0, Math.min(59, date.getMinutes()));
  const seconds = Math.max(0, Math.min(59, date.getSeconds()));

  return {
    dosDate: ((year - 1980) << 9) | (month << 5) | day,
    dosTime: (hours << 11) | (minutes << 5) | Math.floor(seconds / 2),
  };
}

function setGenerating(isGenerating) {
  state.isGenerating = isGenerating;
  elements.generateBtn.disabled = isGenerating;
  elements.pickFolderBtn.disabled = isGenerating;
  elements.addClassBtn.disabled = isGenerating;
  elements.seedDemoBtn.disabled = isGenerating;
}

function setRunStatus(message, type) {
  const typeClass =
    type === "success"
      ? "status-panel success"
      : type === "error"
        ? "status-panel"
        : "status-panel info";

  elements.runStatus.className = typeClass;
  elements.runStatus.textContent = message;
}

function seedDemoData() {
  const demoClasses = [
    {
      className: "2강서A",
      rawText: [
        "abandon\t버리다",
        "ability\t능력",
        "accurate\t정확한",
        "admire\t존경하다",
        "attempt\t시도하다",
        "benefit\t이익",
        "brief\t간단한",
        "challenge\t도전",
        "consider\t고려하다",
        "discover\t발견하다",
      ].join("\n"),
    },
    {
      className: "1강서B",
      rawText: [
        "ancient\t고대의",
        "brilliant\t훌륭한",
        "capture\t포획하다",
        "damage\t손상",
        "eager\t열망하는",
        "feature\t특징",
        "generous\t관대한",
        "harvest\t수확",
        "improve\t향상시키다",
        "journey\t여행",
      ].join("\n"),
    },
  ];

  const activePreset = getActivePresetState();
  activePreset.classes = demoClasses;
  activePreset.layoutLocked = false;
  renderActivePreset();
  syncPresetUi();
  applySettings({
    quizMode: "word-to-meaning",
    questionCount: 10,
    sessionCount: 2,
  });
  persistState();
  updateSummary();
  showToast("예시 데이터를 채웠습니다.");
}

function showToast(message) {
  window.clearTimeout(state.toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  state.toastTimer = window.setTimeout(() => {
    elements.toast.classList.remove("show");
  }, 2800);
}

function waitForLayout() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(async () => {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }
        resolve();
      });
    });
  });
}

function getCompactSaveRootLabel(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return "";
  }

  const segments = raw.split(/[\\/]+/).filter(Boolean);
  const primaryLabel = segments.length ? segments[segments.length - 1] : raw;

  return shortenMiddle(primaryLabel, 18);
}

function shortenMiddle(value, maxLength) {
  const text = String(value ?? "");
  if (text.length <= maxLength) {
    return text;
  }

  const safeMaxLength = Math.max(7, maxLength);
  const omission = "...";
  const visibleLength = safeMaxLength - omission.length;
  const frontLength = Math.ceil(visibleLength / 2);
  const backLength = Math.floor(visibleLength / 2);
  return `${text.slice(0, frontLength)}${omission}${text.slice(-backLength)}`;
}

function setElementTextAndTitle(element, text, titleText = text) {
  if (!element) {
    return;
  }

  const value = String(text ?? "");
  const titleValue = String(titleText ?? value);
  element.textContent = value;
  element.title = titleValue;
}

function adjustSessionCount(delta) {
  const currentValue = getSessionCount();
  const nextValue = setSessionCount(currentValue + delta);

  if (nextValue === currentValue) {
    return;
  }

  persistState();
  updateSummary();
}

function getSessionCount() {
  return clampNumber(
    Number(elements.sessionCountInput?.value),
    1,
    5,
    DEFAULT_SETTINGS.sessionCount
  );
}

function setSessionCount(value) {
  const nextValue = clampNumber(value, 1, 5, DEFAULT_SETTINGS.sessionCount);

  if (elements.sessionCountInput) {
    elements.sessionCountInput.value = String(nextValue);
  }

  if (elements.sessionCountDisplay) {
    elements.sessionCountDisplay.textContent = `${nextValue}차`;
  }

  if (elements.sessionCountDecreaseBtn) {
    const canDecrease = nextValue > 1;
    elements.sessionCountDecreaseBtn.disabled = !canDecrease;
    elements.sessionCountDecreaseBtn.setAttribute("aria-disabled", String(!canDecrease));
  }

  if (elements.sessionCountIncreaseBtn) {
    const canIncrease = nextValue < 5;
    elements.sessionCountIncreaseBtn.disabled = !canIncrease;
    elements.sessionCountIncreaseBtn.setAttribute("aria-disabled", String(!canIncrease));
  }

  return nextValue;
}

function setCheckedValue(name, value) {
  const target = document.querySelector(
    `input[name='${name}'][value='${CSS.escape(String(value))}']`
  );
  if (target) {
    target.checked = true;
  }
}

function getCheckedValue(name, fallback) {
  return document.querySelector(`input[name='${name}']:checked`)?.value ?? fallback;
}

function shuffle(items) {
  const array = [...items];
  for (let index = array.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [array[index], array[swapIndex]] = [array[swapIndex], array[index]];
  }
  return array;
}

function cloneEntry(entry) {
  return { ...entry };
}

function clampNumber(value, min, max, fallback) {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return Math.min(max, Math.max(min, Math.round(value)));
}

function formatDateFolder(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDisplayDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}.${month}.${day}`;
}

function sanitizeFilePart(value) {
  const sanitized = String(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");

  return sanitized || "class";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
