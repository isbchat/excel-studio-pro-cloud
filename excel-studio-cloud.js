(function () {
  const googleClientId = "1091629496669-5881hgv7d4t01eepcq60u8mmdkguc1cn.apps.googleusercontent.com";
  const firebaseConfig = {
  apiKey: "AIzaSyA-d_mtHglAge9ZggTWF3ABhUd_xn2PDKM",
  authDomain: "isb-messenger-75093.firebaseapp.com",
  databaseURL: "https://isb-messenger-75093-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "isb-messenger-75093",
  storageBucket: "isb-messenger-75093.firebasestorage.app",
  messagingSenderId: "970755431087",
  appId: "1:970755431087:web:09171e2ebca48e1b91125a"
};
  const firebaseReady = window.firebase && typeof firebase.auth === "function" && typeof firebase.database === "function";
  if (firebaseReady) firebase.initializeApp(firebaseConfig);
  const auth = firebaseReady ? firebase.auth() : null;
  const database = firebaseReady ? firebase.database() : null;
  let signedInUser = null;
  let workbookLoaded = false;
  let cloudWriteAllowed = false;
  let saveTimer = null;
  let googleInitialized = false;
  let anonymousMigrationPendingUid = null;
  const anonymousWorkbookKey = "excel_studio_pro_v3_guest";
  const skippedSignInKey = "excel_studio_pro_skip_signin_v1";

  if (auth) auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL).catch(() => {});

  function setCloudStatus(message, state) {
    const label = document.getElementById("cloudSyncStatus");
    const dot = document.getElementById("cloudStatusDot");
    if (label) label.textContent = message;
    const colors = { ok: "#22c55e", saving: "#f59e0b", error: "#ef4444", idle: "#94a3b8" };
    if (dot) dot.style.backgroundColor = colors[state] || colors.idle;
    const profileEmail = document.getElementById("profileEmailDisplay");
    if (profileEmail && signedInUser) profileEmail.textContent = `${signedInUser.email || "Google account"} · ${message}`;
  }

  function setGateMessage(message, retryVisible = false) {
    const box = document.getElementById("cloudAuthMessage");
    const retry = document.getElementById("cloudRetryButton");
    if (box) box.textContent = message;
    if (retry) retry.style.display = retryVisible ? "block" : "none";
  }

  function showGate(message) {
    const gate = document.getElementById("cloudAuthGate");
    if (gate) gate.style.display = "flex";
    if (message) setGateMessage(message);
  }

  function profileInitials(user) {
    const initials = (user.displayName || user.email || "G")
      .split(/[\s@._-]+/).filter(Boolean).slice(0, 2).map(part => part[0]).join("").toUpperCase();
    const node = document.getElementById("profileInitials");
    if (node) node.textContent = initials || "G";
  }

  function updateAccountUI(user) {
    const profile = document.getElementById("profileNameDisplay");
    const label = document.getElementById("googleProfileLabel");
    const button = document.getElementById("btnGoogleSignInAction");
    const email = document.getElementById("profileEmailDisplay");
    if (user) {
      const display = user.displayName || user.email || "Google account";
      if (profile) profile.textContent = display;
      if (label) label.textContent = display.split(" ")[0];
      if (button) button.textContent = "Sign out";
      if (email) email.textContent = `${user.email || "Google account"} · cloud sync चालू`;
      profileInitials(user);
      return;
    }

    if (profile) profile.textContent = "Guest mode";
    if (label) label.textContent = "Local";
    if (button) button.textContent = "Sign in to sync";
    if (email) email.textContent = "अभी काम इस डिवाइस पर सेव होगा; बाद में Google से sync करें।";
    const initials = document.getElementById("profileInitials");
    if (initials) initials.textContent = "G";
  }

  function workbookPath(uid) {
    return `excelStudioWorkbooks/${uid}/workbook`;
  }

  function localKey(uid) {
    return `excel_studio_pro_v3_${uid}`;
  }

  function makeWorkbookSnapshot() {
    return {
      schemaVersion: 1,
      updatedAt: Date.now(),
      sheets,
      activeSheetIdx,
      totalCols: TOTAL_COLS,
      totalRows: TOTAL_ROWS,
      colWidths,
      rowHeights
    };
  }

  function validWorkbook(data) {
    return !!data && Array.isArray(data.sheets) && data.sheets.length > 0 &&
      data.sheets.every(sheet => sheet && typeof sheet.name === "string" && sheet.cells && typeof sheet.cells === "object");
  }

  function readLocalWorkbook(uid) {
    try {
      const raw = localStorage.getItem(localKey(uid));
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return validWorkbook(parsed) ? parsed : null;
    } catch (error) {
      return null;
    }
  }

  function readAnonymousWorkbook() {
    try {
      const raw = localStorage.getItem(anonymousWorkbookKey);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return validWorkbook(parsed) ? parsed : null;
    } catch (error) {
      return null;
    }
  }

  function writeLocalWorkbook(snapshot, uid) {
    try {
      localStorage.setItem(localKey(uid), JSON.stringify(snapshot));
      return true;
    } catch (error) {
      showToast("Local backup storage भर गया है। Excel (.xlsx) export कर लें।");
      return false;
    }
  }

  function writeAnonymousWorkbook(snapshot) {
    try {
      localStorage.setItem(anonymousWorkbookKey, JSON.stringify(snapshot));
      return true;
    } catch (error) {
      showToast("Browser storage भर गई है। Workbook को Excel (.xlsx) में export करें।");
      return false;
    }
  }

  function applyWorkbook(snapshot) {
    sheets = snapshot.sheets;
    activeSheetIdx = Math.max(0, Math.min(Number(snapshot.activeSheetIdx) || 0, sheets.length - 1));
    TOTAL_COLS = Math.max(36, Math.min(1200, Number(snapshot.totalCols) || 36));
    TOTAL_ROWS = Math.max(80, Math.min(20000, Number(snapshot.totalRows) || 80));
    colWidths = snapshot.colWidths && typeof snapshot.colWidths === "object" ? snapshot.colWidths : {};
    rowHeights = snapshot.rowHeights && typeof snapshot.rowHeights === "object" ? snapshot.rowHeights : {};
    activeCell = { col: 0, row: 0 };
    selectionStart = { col: 0, row: 0 };
    selectionEnd = { col: 0, row: 0 };
    editingCoord = null;
  }

  function renderWorkbook() {
    buildGridDOM();
    renderSheetTabs();
    updateSelectionUI();
  }

  function hasSkippedSignIn() {
    try { return localStorage.getItem(skippedSignInKey) === "true"; }
    catch (error) { return false; }
  }

  function rememberSkippedSignIn() {
    try { localStorage.setItem(skippedSignInKey, "true"); }
    catch (error) {}
  }

  function clearSkippedSignIn() {
    try { localStorage.removeItem(skippedSignInKey); }
    catch (error) {}
  }

  function readLegacyWorkbook() {
    try {
      const raw = localStorage.getItem("excel_studio_pro_v3");
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      const legacy = Array.isArray(parsed) ? { sheets: parsed, activeSheetIdx: 0 } : parsed;
      return validWorkbook(legacy) ? legacy : null;
    } catch (error) {
      return null;
    }
  }

  function startLocalWorkbook(persistChoice = true) {
    if (signedInUser) return;
    if (persistChoice) rememberSkippedSignIn();
    clearTimeout(saveTimer);
    cloudWriteAllowed = false;

    let restoredLegacy = false;
    if (!workbookLoaded) {
      const saved = readAnonymousWorkbook();
      const legacy = saved ? null : readLegacyWorkbook();
      if (saved) applyWorkbook(saved);
      else if (legacy) {
        applyWorkbook(legacy);
        restoredLegacy = true;
      }
    }

    workbookLoaded = true;
    const gate = document.getElementById("cloudAuthGate");
    if (gate) gate.style.display = "none";
    updateAccountUI(null);
    renderWorkbook();
    if (restoredLegacy) writeAnonymousWorkbook(makeWorkbookSnapshot());
    setCloudStatus("Saved on this device", "idle");
    if (persistChoice && typeof showToast === "function") {
      showToast("अभी इस डिवाइस पर काम करें। बाद में Google से sign in करके cloud sync करें।");
    }
  }

  window.skipExcelStudioCloudSignIn = function () {
    startLocalWorkbook(true);
  };

  function chooseWorkbook(remote, local, anonymous) {
    if (remote && !validWorkbook(remote)) throw new Error("Cloud workbook format is not recognized.");
    if (local && !validWorkbook(local)) local = null;
    if (anonymous && !validWorkbook(anonymous)) anonymous = null;
    const candidates = [
      { data: remote, source: "cloud" },
      { data: local, source: "local" },
      { data: anonymous, source: "anonymous" }
    ].filter(candidate => candidate.data);
    candidates.sort((a, b) => Number(b.data.updatedAt || 0) - Number(a.data.updatedAt || 0));
    return candidates[0] || { data: null, source: "default" };
  }

  async function loadWorkbookForUser(user) {
    const sessionSnapshot = workbookLoaded ? makeWorkbookSnapshot() : null;
    signedInUser = user;
    workbookLoaded = false;
    cloudWriteAllowed = false;
    updateAccountUI(user);
    setCloudStatus("Loading workbook…", "saving");
    setGateMessage("Google account connected. आपका cloud workbook लोड हो रहा है…");

    const local = readLocalWorkbook(user.uid);
    const anonymous = readAnonymousWorkbook();
    try {
      const snapshot = await database.ref(workbookPath(user.uid)).once("value");
      const remote = snapshot.exists() ? snapshot.val() : null;
      const chosen = chooseWorkbook(remote, local, anonymous);

      if (!chosen.data) {
        const legacy = readLegacyWorkbook();
        if (legacy && confirm("इस browser में पुरानी workbook मिली। क्या उसे इस Google account में copy करें?")) {
          applyWorkbook(legacy);
          chosen.source = "legacy";
        }
      } else {
        applyWorkbook(chosen.data);
      }

      workbookLoaded = true;
      clearSkippedSignIn();
      const gate = document.getElementById("cloudAuthGate");
      if (gate) gate.style.display = "none";
      cloudWriteAllowed = false;
      renderWorkbook();
      cloudWriteAllowed = true;
      if (chosen.source !== "cloud") {
        writeLocalWorkbook(makeWorkbookSnapshot(), user.uid);
        anonymousMigrationPendingUid = chosen.source === "anonymous" ? user.uid : null;
        cloudAutoSaveLocal();
      } else {
        writeLocalWorkbook(makeWorkbookSnapshot(), user.uid);
        setCloudStatus("Saved to Google Cloud", "ok");
      }
    } catch (error) {
      const candidates = [
        { data: sessionSnapshot, source: "session" },
        { data: local, source: "local" },
        { data: anonymous, source: "anonymous" }
      ].filter(candidate => validWorkbook(candidate.data));
      candidates.sort((a, b) => Number(b.data.updatedAt || 0) - Number(a.data.updatedAt || 0));
      if (candidates[0]) applyWorkbook(candidates[0].data);
      workbookLoaded = true;
      cloudWriteAllowed = false;
      const gate = document.getElementById("cloudAuthGate");
      if (gate) gate.style.display = "none";
      renderWorkbook();
      writeLocalWorkbook(makeWorkbookSnapshot(), user.uid);
      setCloudStatus("Cloud unavailable; saved on this device", "error");
      showToast("Cloud workbook अभी नहीं खुली। काम इस device पर चलता रहेगा; sync बाद में फिर करें।");
    }
  }

  async function saveWorkbookToCloud(snapshot, uid) {
    if (!auth || !database || !cloudWriteAllowed || !auth.currentUser || auth.currentUser.uid !== uid) return;
    setCloudStatus("Saving to Google Cloud…", "saving");
    try {
      await database.ref(workbookPath(uid)).set(snapshot);
      if (auth.currentUser && auth.currentUser.uid === uid) {
        setCloudStatus("Saved to Google Cloud", "ok");
        if (anonymousMigrationPendingUid === uid) {
          const anonymous = readAnonymousWorkbook();
          if (!anonymous || Number(anonymous.updatedAt || 0) <= Number(snapshot.updatedAt || 0)) {
            try { localStorage.removeItem(anonymousWorkbookKey); }
            catch (error) {}
            anonymousMigrationPendingUid = null;
          }
        }
      }
    } catch (error) {
      setCloudStatus("Local copy saved; cloud save failed", "error");
      showToast("Cloud save नहीं हुआ। इस device की copy सुरक्षित है; बाद में Sync Now दबाएँ।");
    }
  }

  window.cloudAutoSaveLocal = function () {
    if (!workbookLoaded) return;
    const snapshot = makeWorkbookSnapshot();
    if (!signedInUser) {
      writeAnonymousWorkbook(snapshot);
      setCloudStatus("Saved on this device", "idle");
      return;
    }

    const userId = signedInUser.uid;
    writeLocalWorkbook(snapshot, userId);
    if (!cloudWriteAllowed) {
      setCloudStatus("Cloud unavailable; saved on this device", "error");
      return;
    }
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => saveWorkbookToCloud(snapshot, userId), 1000);
  };

  window.retryExcelStudioCloudLoad = function () {
    if (auth && auth.currentUser) return loadWorkbookForUser(auth.currentUser);
    showGate(auth ? "Google से sign in करें ताकि workbook cloud में sync हो।" : "Firebase अभी उपलब्ध नहीं है। अभी skip करके इस device पर काम कर सकते हैं।");
    if (auth) initGoogleSignIn();
  };

  window.logoutCloudUser = async function () {
    if (!auth || !auth.currentUser) {
      window.retryExcelStudioCloudLoad();
      return;
    }
    if (!confirm("Sign out करें? इस account की cloud workbook सुरक्षित रहेगी।")) return;
    clearTimeout(saveTimer);
    if (workbookLoaded) {
      const snapshot = makeWorkbookSnapshot();
      writeLocalWorkbook(snapshot, auth.currentUser.uid);
      writeAnonymousWorkbook(snapshot);
    }
    rememberSkippedSignIn();
    if (workbookLoaded && cloudWriteAllowed) await saveWorkbookToCloud(makeWorkbookSnapshot(), auth.currentUser.uid);
    await auth.signOut();
  };

  window.handleExcelStudioGoogleCredential = async function (response) {
    if (!response || !response.credential) {
      setGateMessage("Google sign-in credential नहीं मिला। फिर कोशिश करें।");
      return;
    }
    setGateMessage("Google account verify हो रहा है…");
    try {
      const credential = firebase.auth.GoogleAuthProvider.credential(response.credential);
      await auth.signInWithCredential(credential);
    } catch (error) {
      const messages = {
        "auth/operation-not-allowed": "Firebase Console में Google sign-in provider enable करें।",
        "auth/unauthorized-domain": "Firebase Authentication में isbchat.github.io को authorized domain जोड़ें।",
        "auth/invalid-credential": "Google OAuth client ID या Firebase Google provider configuration जाँचें।"
      };
      setGateMessage(messages[error.code] || `Google sign-in failed (${error.code || "error"}). Firebase setup जाँचें।`);
    }
  };

  function initGoogleSignIn(attempt = 0) {
    if (googleInitialized) return;
    if (!window.google || !google.accounts || !google.accounts.id) {
      if (attempt < 40) return setTimeout(() => initGoogleSignIn(attempt + 1), 150);
      setGateMessage("Google sign-in script load नहीं हुई। Internet जाँचें और Retry करें।", true);
      return;
    }
    google.accounts.id.initialize({
      client_id: googleClientId,
      callback: window.handleExcelStudioGoogleCredential,
      auto_select: false
    });
    google.accounts.id.renderButton(document.getElementById("cloudGoogleButtonContainer"), {
      theme: "filled_blue", size: "large", text: "continue_with", shape: "pill", width: 250
    });
    googleInitialized = true;
    setGateMessage("Continue with Google to load or create your cloud workbook.");
  }

  if (auth) {
    auth.onAuthStateChanged(user => {
      if (user) {
        clearSkippedSignIn();
        updateAccountUI(user);
        loadWorkbookForUser(user);
      } else {
        signedInUser = null;
        cloudWriteAllowed = false;
        updateAccountUI(null);
        setCloudStatus("Google sign-in required", "idle");
        if (hasSkippedSignIn()) {
          startLocalWorkbook(false);
        } else {
          workbookLoaded = false;
          showGate("Google से sign in करें या skip करके इस device पर काम करें।");
          initGoogleSignIn();
        }
      }
    });
  } else {
    document.addEventListener("DOMContentLoaded", () => {
      updateAccountUI(null);
      if (hasSkippedSignIn()) {
        startLocalWorkbook(false);
      } else {
        setCloudStatus("Google sign-in unavailable", "error");
        setGateMessage("Firebase/Google अभी नहीं खुला। Skip करके इस device पर काम कर सकते हैं।");
      }
    });
  }
})();

