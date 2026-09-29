/*
 * Shared logic for the K1 keyboard-shortcut spike task panes.
 * Loaded by taskpane-a.html / taskpane-b.html / taskpane-c.html, each of which
 * defines a small `window.EMT_K1_CONFIG` object before this script runs, then
 * calls `window.emtK1Init()` once Office.js has loaded.
 *
 * config shape:
 * {
 *   name: string,                          // display name, e.g. "EMT K1 Named Keys"
 *   actions: [{ id: string, key: string }], // every action id + the key string bound to it in shortcuts-X.json
 *   inUseCheckKeys: [string],               // key strings to pass to Office.actions.areShortcutsInUse
 *   replaceCandidates: [{ action: string, key: string }] // a couple of candidates to try Office.actions.replaceShortcuts on
 * }
 */

(function () {
  "use strict";

  function nowStamp() {
    return new Date().toISOString().split("T")[1].replace("Z", "");
  }

  // Mirror everything to the local spike server (POST /log), which appends it to
  // %LOCALAPPDATA%\EmtSpike\k1-log.jsonl so results can be read without screenshots.
  function remote(kind, data) {
    try {
      const addin = (window.EMT_K1_CONFIG && window.EMT_K1_CONFIG.name) || "unknown";
      fetch("/log", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ts: new Date().toISOString(), addin, kind, data }),
      }).catch(function () { /* logging must never break the spike */ });
    } catch (e) { /* ignore */ }
  }

  function appendLog(message) {
    remote("log", message);
    const line = `${nowStamp()} ${message}`;
    console.log(line);
    const list = document.getElementById("log");
    if (list) {
      const li = document.createElement("li");
      li.textContent = line;
      list.prepend(li);
    }
  }

  function appendError(message) {
    remote("error", message);
    console.error(message);
    const box = document.getElementById("errors");
    if (box) {
      const p = document.createElement("p");
      p.className = "error-line";
      p.textContent = message;
      box.appendChild(p);
    }
  }

  function setText(id, text) {
    remote("field", { id: id, text: text });
    const el = document.getElementById(id);
    if (el) {
      el.textContent = text;
    }
  }

  function fired(actionId, keyLabel) {
    appendLog(`fired: ${actionId} (${keyLabel})`);
  }

  function renderActionsTable(config) {
    const tbody = document.getElementById("actions-tbody");
    if (!tbody) {
      return;
    }
    config.actions.forEach(({ id, key }) => {
      const tr = document.createElement("tr");
      const tdId = document.createElement("td");
      tdId.textContent = id;
      const tdKey = document.createElement("td");
      tdKey.textContent = key;
      tr.appendChild(tdId);
      tr.appendChild(tdKey);
      tbody.appendChild(tr);
    });
  }

  // Registers Office.actions.associate for every configured action. Kept at
  // top level (not inside Office.onReady) because SharedRuntime add-ins are
  // expected to associate actions as early as possible.
  function associateAll(config) {
    if (!window.Office || !Office.actions || typeof Office.actions.associate !== "function") {
      appendError("Office.actions.associate is not available - cannot register any shortcut handlers.");
      return;
    }
    config.actions.forEach(({ id, key }) => {
      try {
        Office.actions.associate(id, function (event) {
          fired(id, key);
          if (event && typeof event.completed === "function") {
            event.completed();
          }
        });
      } catch (err) {
        appendError(`Office.actions.associate threw for "${id}": ${err && err.message ? err.message : err}`);
      }
    });
    appendLog(`associated ${config.actions.length} action(s).`);
  }

  async function runDiagnostics(config) {
    setText("addin-name", config.name);

    try {
      const diag = Office.context.diagnostics;
      setText(
        "diagnostics",
        `host=${diag.host}, platform=${diag.platform}, version=${diag.version}`
      );
    } catch (err) {
      appendError(`Reading Office.context.diagnostics failed: ${err.message || err}`);
    }

    let sharedRuntimeOk = false;
    let keyboardShortcutsOk = false;
    try {
      sharedRuntimeOk = Office.context.requirements.isSetSupported("SharedRuntime", "1.1");
    } catch (err) {
      appendError(`isSetSupported(SharedRuntime,1.1) failed: ${err.message || err}`);
    }
    try {
      keyboardShortcutsOk = Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1");
    } catch (err) {
      appendError(`isSetSupported(KeyboardShortcuts,1.1) failed: ${err.message || err}`);
    }
    setText(
      "requirement-sets",
      `SharedRuntime 1.1 = ${sharedRuntimeOk}; KeyboardShortcuts 1.1 = ${keyboardShortcutsOk}`
    );

    if (!keyboardShortcutsOk) {
      appendLog("KeyboardShortcuts 1.1 not reported as supported - skipping getShortcuts/areShortcutsInUse/replaceShortcuts probes.");
      return;
    }

    // getShortcuts(): this is the key evidence of what Office actually parsed
    // out of shortcuts-X.json (vs. what we wrote in it).
    try {
      const result = await Office.actions.getShortcuts();
      setText("get-shortcuts", JSON.stringify(result, null, 2));
      appendLog(`getShortcuts() returned ${Array.isArray(result) ? result.length : "?"} entr(y/ies).`);
    } catch (err) {
      const msg = `getShortcuts() failed: ${err && err.message ? err.message : err}`;
      setText("get-shortcuts", msg);
      appendError(msg);
    }

    // areShortcutsInUse(): probe each key string ON ITS OWN, so one invalid string can't fail the whole call.
    const inUse = [];
    for (const k of config.inUseCheckKeys || []) {
      try {
        const r = await Office.actions.areShortcutsInUse([k]);
        inUse.push({ key: k, ok: true, result: r });
      } catch (err) {
        inUse.push({ key: k, ok: false, error: err && err.message ? err.message : String(err) });
      }
    }
    setText("in-use", JSON.stringify(inUse, null, 2));
    appendLog(`areShortcutsInUse(): ${inUse.filter((x) => x.ok).length}/${inUse.length} key strings accepted individually.`);

    // replaceShortcuts(): documented shape {actionId: key}. Each candidate re-keys a registered action;
    // after an accepted replace, read getShortcuts() back to see what Office stored.
    const replaceResults = [];
    for (const candidate of config.replaceCandidates || []) {
      const payload = {};
      payload[candidate.action] = candidate.key;
      try {
        await Office.actions.replaceShortcuts(payload);
        let stored = null;
        try { stored = await Office.actions.getShortcuts(); } catch (e) { stored = "getShortcuts failed: " + e.message; }
        replaceResults.push({ candidate, ok: true, stored });
        appendLog(`replaceShortcuts OK: ${candidate.action} -> ${candidate.key}; stored=${JSON.stringify(stored)}`);
      } catch (err) {
        replaceResults.push({ candidate, ok: false, error: err && err.message ? err.message : String(err) });
        appendLog(`replaceShortcuts REJECTED: ${candidate.action} -> ${candidate.key}: ${err.message || err}`);
      }
    }
    if (config.revertAfterReplace) {
      const revert = {};
      for (const c of config.replaceCandidates || []) revert[c.action] = null;
      try {
        await Office.actions.replaceShortcuts(revert);
        appendLog(`reverted to manifest defaults; now ${JSON.stringify(await Office.actions.getShortcuts())}`);
      } catch (err) {
        appendError(`revert failed: ${err.message || err}`);
      }
    }
    setText("replace-shortcuts", JSON.stringify(replaceResults, null, 2));
  }

  window.emtK1Init = function emtK1Init() {
    const config = window.EMT_K1_CONFIG;
    if (!config) {
      appendError("window.EMT_K1_CONFIG is missing - page cannot initialize.");
      return;
    }

    Office.onReady()
      .then(() => {
        appendLog(`Office.onReady resolved for ${config.name}.`);
        return runDiagnostics(config);
      })
      .catch((err) => {
        appendError(`Office.onReady rejected: ${err && err.message ? err.message : err}`);
      });

    // Associate immediately (do not wait on onReady) per SharedRuntime guidance.
    associateAll(config);
    renderActionsTable(config);
  };
})();
