(function () {
  const googleClientId = "1091629496669-5881hgv7d4t01eepcq60u8mmdkguc1cn.apps.googleusercontent.com";
  const firebaseConfig = {
  apiKey: "AIzaSyA-d_mtHg1Age9ZggTWF3ABhUd_xn2PDKM",
  authDomain: "isb-messenger-75093.firebaseapp.com",
  databaseURL: "https://isb-messenger-75093-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "isb-messenger-75093",
  storageBucket: "isb-messenger-75093.firebasestorage.app",
  messagingSenderId: "970755431087",
  appId: "1:970755431087:web:09171e2ebca48e1b91125a"
};
  const firebaseReady = window.firebase && typeof firebase.auth === "function" && typeof firebase.database === "function";
  if (!firebaseReady) {
    window.addEventListener("DOMContentLoaded", () => {
      const message = document.getElementById("cloudAuthMessage");
      if (message) message.textContent = "Firebase नहीं खुला। इंटरनेट कनेक्शन जाँचें और page फिर से खोलें।";
    });
    return;
  }

  firebase.initializeApp(firebaseConfig);
  const auth = firebase.auth();
  const database = firebase.database();
  let signedInUser = null;
  let workbookLoaded = false;
  let cloudWriteAllowed = false;
  let saveTimer = null;
  let googleInitialized = false;

  auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL).catch(() => {});

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
    const display = user.displayName || user.email || "Google account";
    const profile = document.getElementById("profileNameDisplay");
    const label = document.getElementById("googleProfileLabel");
    const button = document.getElementById("btnGoogleSignInAction");
    if (profile) profile.textContent = display;
    if (label) label.textContent = display.split(" ")[0];
    if (button) button.textContent = "Sign out";
    profileInitials(user);
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

  function writeLocalWorkbook(snapshot, uid) {
    try {
      localStorage.setItem(localKey(uid), JSON.stringify(snapshot));
      return true;
    } catch (error) {
      showToast("Local backup storage भर गया है। Excel (.xlsx) export कर लें।");
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

  function chooseWorkbook(remote, local) {
    if (remote && !validWorkbook(remote)) throw new Error("Cloud workbook format is not recognized.");
    if (local && !validWorkbook(local)) local = null;
    if (!remote) return { data: local, source: local ? "local" : "default" };
    if (!local) return { data: remote, source: "cloud" };
    return Number(local.updatedAt || 0) > Number(remote.updatedAt || 0)
      ? { data: local, source: "local" }
      : { data: remote, source: "cloud" };
  }

  async function loadWorkbookForUser(user) {
    signedInUser = user;
    workbookLoaded = false;
    cloudWriteAllowed = false;
    updateAccountUI(user);
    setCloudStatus("Loading workbook…", "saving");
    setGateMessage("Google account connected. आपका cloud workbook लोड हो रहा है…");

    const local = readLocalWorkbook(user.uid);
    try {
      const snapshot = await database.ref(workbookPath(user.uid)).once("value");
      const remote = snapshot.exists() ? snapshot.val() : null;
      const chosen = chooseWorkbook(remote, local);

      if (!chosen.data) {
        const legacyRaw = localStorage.getItem("excel_studio_pro_v3");
        if (legacyRaw) {
          try {
            const legacySheets = JSON.parse(legacyRaw);
            if (Array.isArray(legacySheets) && legacySheets.length &&
                confirm("इस browser में पुरानी local workbook मिली। क्या उसे इस Google account में copy करें?")) {
              applyWorkbook({ sheets: legacySheets, activeSheetIdx: 0 });
            }
          } catch (error) {}
        }
      } else {
        applyWorkbook(chosen.data);
      }

      workbookLoaded = true;
      document.getElementById("cloudAuthGate").style.display = "none";
      cloudWriteAllowed = false;
      renderWorkbook();
      cloudWriteAllowed = true;
      if (chosen.source === "local" || chosen.source === "default") {
        cloudAutoSaveLocal();
      } else {
        writeLocalWorkbook(makeWorkbookSnapshot(), user.uid);
        setCloudStatus("Saved to Google Cloud", "ok");
      }
    } catch (error) {
      const cached = readLocalWorkbook(user.uid);
      if (cached) {
        applyWorkbook(cached);
        workbookLoaded = true;
        cloudWriteAllowed = false;
        document.getElementById("cloudAuthGate").style.display = "none";
        renderWorkbook();
        setCloudStatus("Local copy loaded; cloud access failed", "error");
        showToast("Cloud workbook नहीं पढ़ सकी। Local copy सुरक्षित है; rules/network जाँचें।");
      } else {
        workbookLoaded = false;
        cloudWriteAllowed = false;
        showGate("Firebase workbook नहीं पढ़ सकी। Database rules और internet जाँचकर Retry करें।", true);
        setCloudStatus("Cloud unavailable", "error");
      }
    }
  }

  async function saveWorkbookToCloud(snapshot, uid) {
    if (!cloudWriteAllowed || !auth.currentUser || auth.currentUser.uid !== uid) return;
    setCloudStatus("Saving to Google Cloud…", "saving");
    try {
      await database.ref(workbookPath(uid)).set(snapshot);
      if (auth.currentUser && auth.currentUser.uid === uid) setCloudStatus("Saved to Google Cloud", "ok");
    } catch (error) {
      setCloudStatus("Local copy saved; cloud save failed", "error");
      showToast("Cloud save नहीं हुआ। Firebase rules और internet जाँचें।");
    }
  }

  window.cloudAutoSaveLocal = function () {
    if (!signedInUser || !workbookLoaded) return;
    const userId = signedInUser.uid;
    const snapshot = makeWorkbookSnapshot();
    writeLocalWorkbook(snapshot, userId);
    if (!cloudWriteAllowed) return;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => saveWorkbookToCloud(snapshot, userId), 1000);
  };

  window.retryExcelStudioCloudLoad = function () {
    if (auth.currentUser) return loadWorkbookForUser(auth.currentUser);
    showGate("पहले Google से sign in करें।");
    initGoogleSignIn();
  };

  window.logoutCloudUser = async function () {
    if (!auth.currentUser) return;
    if (!confirm("Sign out करें? इस account की cloud workbook सुरक्षित रहेगी।")) return;
    clearTimeout(saveTimer);
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

  auth.onAuthStateChanged(user => {
    if (user) {
      updateAccountUI(user);
      loadWorkbookForUser(user);
    } else {
      signedInUser = null;
      workbookLoaded = false;
      cloudWriteAllowed = false;
      document.getElementById("cloudAuthGate").style.display = "flex";
      const label = document.getElementById("googleProfileLabel");
      if (label) label.textContent = "Not signed in";
      setCloudStatus("Google sign-in required", "idle");
      setGateMessage("Google से sign in करें। आपकी workbook cloud से load होगी या नई workbook बनेगी।");
      initGoogleSignIn();
    }
  });
})();
