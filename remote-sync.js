/**
 * v1.2 — Auto-refresh SRB + Rules CSV from SharePoint (Microsoft Graph) or
 * from two direct HTTPS URLs. Loaded after the main app script; expects
 * window.runOverlapAnalyze, acceptFile1, parseCSV, document.getElementById('dz2Text'), etc.
 */
(function () {
  'use strict';

  const STORAGE_KEY = 'overlapRemoteSync_v1.2';
  const GRAPHS = 'https://graph.microsoft.com/v1.0';
  const MSAL_CDN = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.27.0/lib/msal-browser.min.js';
  const SCOPES = ['User.Read', 'Sites.Read.All'];

  const DEFAULTS = {
    mode: 'graph',
    siteHostname: 'moodys.sharepoint.com',
    sitePath: '/sites/MA_Roadmap',
    libraryName: 'Territory Quota Management',
    folderPath:
      'Territory & Quota Planning 2027/External Pages in TQP/1. Overlap Report',
    srbFileName: '',
    rulesFileName: '',
    clientId: '',
    tenantId: '',
    srbUrl: '',
    rulesUrl: '',
    urlBearer: '',
    pollMinutes: 15,
    autoEnabled: false,
    lastSrbSig: '',
    lastRulesSig: '',
    lastPollOk: '',
    lastError: ''
  };

  let pollTimer = null;
  let msalReady = null;

  function $(id) { return document.getElementById(id); }

  function loadCfg() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return { ...DEFAULTS };
      return { ...DEFAULTS, ...JSON.parse(raw) };
    } catch (_) {
      return { ...DEFAULTS };
    }
  }

  function saveCfg(patch) {
    const cur = loadCfg();
    Object.assign(cur, patch);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cur));
    return cur;
  }

  function setStatus(msg) {
    const el = $('rsStatus');
    if (el) el.textContent = msg;
  }

  function graphEncodePath(path) {
    return path
      .split('/')
      .filter(function (s) { return s.length; })
      .map(function (seg) { return encodeURIComponent(seg); })
      .join('/');
  }

  async function graphFetch(token, url) {
    const r = await fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      cache: 'no-store'
    });
    if (!r.ok) {
      const t = await r.text();
      throw new Error('Graph ' + r.status + ': ' + (t.slice(0, 200) || r.statusText));
    }
    return r;
  }

  function getPcCtor() {
    if (typeof window.PublicClientApplication === 'function') return window.PublicClientApplication;
    if (window.msal && typeof window.msal.PublicClientApplication === 'function')
      return window.msal.PublicClientApplication;
    if (window.msalBrowser && typeof window.msalBrowser.PublicClientApplication === 'function')
      return window.msalBrowser.PublicClientApplication;
    return null;
  }

  async function ensureMsal() {
    if (getPcCtor()) return;
    if (!msalReady) {
      msalReady = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = MSAL_CDN;
        s.async = true;
        s.onload = function () { resolve(); };
        s.onerror = function () { reject(new Error('Failed to load MSAL from CDN')); };
        document.head.appendChild(s);
      });
    }
    await msalReady;
    if (!getPcCtor()) throw new Error('MSAL loaded but PublicClientApplication not found.');
  }

  function redirectUri() {
    const u = window.location.href.split('#')[0];
    return u.split('?')[0];
  }

  async function getMsalInstance(cfg) {
    await ensureMsal();
    const PCA = getPcCtor();
    return new PCA({
      auth: {
        clientId: cfg.clientId,
        authority: 'https://login.microsoftonline.com/' + cfg.tenantId,
        redirectUri: redirectUri()
      },
      cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
    });
  }

  async function acquireToken(cfg, msal) {
    const active = msal.getActiveAccount();
    const accounts = msal.getAllAccounts();
    const account = active || accounts[0];
    if (!account) throw new Error('Sign in first (Sign in button).');
    try {
      const silent = await msal.acquireTokenSilent({
        scopes: SCOPES,
        account: account
      });
      return silent.accessToken;
    } catch (_) {
      const popup = await msal.acquireTokenPopup({
        scopes: SCOPES,
        account: account
      });
      return popup.accessToken;
    }
  }

  async function graphSiteId(token, cfg) {
    const host = encodeURIComponent(cfg.siteHostname);
    const path = encodeURIComponent(cfg.sitePath.replace(/^\/+/, ''));
    const url = GRAPHS + '/sites/' + host + ':' + path;
    const r = await graphFetch(token, url);
    const j = await r.json();
    if (!j.id) throw new Error('Could not resolve SharePoint site.');
    return j.id;
  }

  async function graphDriveId(token, siteId, libraryName) {
    const url = GRAPHS + '/sites/' + siteId + '/drives';
    const r = await graphFetch(token, url);
    const j = await r.json();
    const drives = j.value || [];
    const match = drives.filter(function (d) {
      return d.name === libraryName || (d.name && d.name.toLowerCase() === libraryName.toLowerCase());
    });
    if (!match.length)
      throw new Error(
        'Document library "' +
          libraryName +
          '" not found. Available: ' +
          drives.map(function (d) { return d.name; }).join(', ')
      );
    return match[0].id;
  }

  async function graphFolderChildren(token, driveId, folderPath) {
    const enc = graphEncodePath(folderPath);
    const url = GRAPHS + '/drives/' + driveId + '/root:/' + enc + ':/children';
    const r = await graphFetch(token, url);
    const j = await r.json();
    return j.value || [];
  }

  function findItemByName(items, name) {
    const want = name.trim().toLowerCase();
    return items.filter(function (it) {
      return it.name && it.name.toLowerCase() === want;
    })[0];
  }

  async function graphDownloadItem(token, driveId, itemId, suggestedName) {
    const url = GRAPHS + '/drives/' + driveId + '/items/' + itemId + '/content';
    const r = await graphFetch(token, url);
    const blob = await r.blob();
    return new File([blob], suggestedName || 'download.csv', { type: 'text/csv' });
  }

  async function refreshFromGraph(cfg, silentToken) {
    if (!cfg.clientId || !cfg.tenantId) throw new Error('Enter Azure AD Application ID and Tenant ID.');
    if (!cfg.libraryName || !cfg.folderPath) throw new Error('Enter library name and folder path.');
    if (!cfg.srbFileName || !cfg.rulesFileName)
      throw new Error('Enter exact SRB and Rules CSV file names as stored in the folder.');

    const msal = await getMsalInstance(cfg);
    const token = silentToken || (await acquireToken(cfg, msal));

    const siteId = await graphSiteId(token, cfg);
    const driveId = await graphDriveId(token, siteId, cfg.libraryName);
    const items = await graphFolderChildren(token, driveId, cfg.folderPath);

    const srbItem = findItemByName(items, cfg.srbFileName);
    const rulesItem = findItemByName(items, cfg.rulesFileName);
    if (!srbItem) throw new Error('SRB file "' + cfg.srbFileName + '" not found in folder.');
    if (!rulesItem) throw new Error('Rules file "' + cfg.rulesFileName + '" not found in folder.');

    const sigSrb = srbItem.lastModifiedDateTime || srbItem.eTag || '';
    const sigRules = rulesItem.lastModifiedDateTime || rulesItem.eTag || '';

    const prev = loadCfg();
    const unchanged =
      prev.lastSrbSig === sigSrb &&
      prev.lastRulesSig === sigRules &&
      prev.lastSrbSig !== '';

    saveCfg({ lastSrbSig: sigSrb, lastRulesSig: sigRules, lastError: '' });

    if (unchanged) {
      setStatus(
        'Checked at ' +
          new Date().toLocaleString() +
          ' — no changes since last load (' +
          sigSrb.slice(0, 19) +
          ' / ' +
          sigRules.slice(0, 19) +
          ').'
      );
      saveCfg({ lastPollOk: new Date().toISOString() });
      return false;
    }

    const srbFile = await graphDownloadItem(token, driveId, srbItem.id, srbItem.name);
    const rulesFile = await graphDownloadItem(token, driveId, rulesItem.id, rulesItem.name);

    if (typeof window.acceptFile1 !== 'function') throw new Error('acceptFile1 missing.');
    window.acceptFile1(srbFile);

    const rulesText = await rulesFile.text();
    if (typeof window.parseCSV !== 'function') throw new Error('parseCSV missing.');
    if (typeof window.setOverlapRulesFile !== 'function') throw new Error('setOverlapRulesFile missing.');
    window.setOverlapRulesFile(window.parseCSV(rulesText), rulesItem.name);

    if (typeof window.runOverlapAnalyze === 'function') await window.runOverlapAnalyze();

    saveCfg({ lastPollOk: new Date().toISOString() });
    setStatus(
      'Reloaded at ' +
        new Date().toLocaleString() +
        ' from Graph.\nSRB mtime: ' +
        sigSrb +
        '\nRules mtime: ' +
        sigRules
    );
    return true;
  }

  async function headOrSig(url, bearer) {
    const headers = {};
    if (bearer && bearer.trim()) headers.Authorization = 'Bearer ' + bearer.trim();
    try {
      const h = await fetch(url, { method: 'HEAD', cache: 'no-store', headers: headers });
      if (h.ok) {
        return (
          h.headers.get('etag') ||
          h.headers.get('last-modified') ||
          ''
        );
      }
    } catch (_) { /* fall through */ }
    const r = await fetch(url, { method: 'GET', cache: 'no-store', headers: headers });
    if (!r.ok) throw new Error('GET ' + url + ' failed: ' + r.status);
    return (
      r.headers.get('etag') ||
      r.headers.get('last-modified') ||
      String(r.headers.get('content-length') || '')
    );
  }

  async function refreshFromUrls(cfg) {
    if (!cfg.srbUrl || !cfg.rulesUrl) throw new Error('Enter both HTTPS URLs.');
    const bearer = cfg.urlBearer || '';

    const sigSrb = await headOrSig(cfg.srbUrl, bearer);
    const sigRules = await headOrSig(cfg.rulesUrl, bearer);

    const prev = loadCfg();
    const unchanged =
      prev.lastSrbSig === sigSrb &&
      prev.lastRulesSig === sigRules &&
      prev.lastSrbSig !== '';

    saveCfg({ lastSrbSig: sigSrb, lastRulesSig: sigRules, lastError: '' });

    if (unchanged) {
      setStatus('Checked at ' + new Date().toLocaleString() + ' — URLs unchanged.');
      saveCfg({ lastPollOk: new Date().toISOString() });
      return false;
    }

    const headers = {};
    if (bearer.trim()) headers.Authorization = 'Bearer ' + bearer.trim();

    const [rs, rr] = await Promise.all([
      fetch(cfg.srbUrl, { cache: 'no-store', headers: headers }),
      fetch(cfg.rulesUrl, { cache: 'no-store', headers: headers })
    ]);
    if (!rs.ok) throw new Error('SRB URL failed: ' + rs.status);
    if (!rr.ok) throw new Error('Rules URL failed: ' + rr.status);

    const srbBlob = await rs.blob();
    const rulesText = await rr.text();

    const srbFile = new File([srbBlob], 'SRB_remote.csv', { type: 'text/csv' });
    window.acceptFile1(srbFile);
    if (typeof window.setOverlapRulesFile !== 'function') throw new Error('setOverlapRulesFile missing.');
    window.setOverlapRulesFile(window.parseCSV(rulesText), 'Rules_remote.csv');

    await window.runOverlapAnalyze();

    saveCfg({ lastPollOk: new Date().toISOString() });
    setStatus('Reloaded from URLs at ' + new Date().toLocaleString());
    return true;
  }

  async function checkNow() {
    const cfg = loadCfg();
    setStatus('Checking…');
    try {
      if (cfg.mode === 'url') await refreshFromUrls(cfg);
      else await refreshFromGraph(cfg);
    } catch (e) {
      saveCfg({ lastError: String(e.message || e) });
      setStatus('Error: ' + (e.message || e));
    }
  }

  function stopPolling() {
    if (pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }
  }

  function startPolling() {
    stopPolling();
    const cfg = loadCfg();
    if (!cfg.autoEnabled) return;
    var okGraph =
      cfg.mode !== 'url' &&
      cfg.clientId &&
      cfg.tenantId &&
      cfg.srbFileName &&
      cfg.rulesFileName &&
      cfg.libraryName &&
      cfg.folderPath;
    var okUrl = cfg.mode === 'url' && cfg.srbUrl && cfg.rulesUrl;
    if (!okGraph && !okUrl) return;
    const min = Math.max(2, Math.min(1440, Number(cfg.pollMinutes) || 15));
    const ms = min * 60 * 1000;
    pollTimer = setInterval(function () {
      checkNow();
    }, ms);
    setStatus('Automatic poll every ' + min + ' min (when tab is open).');
  }

  function fillForm() {
    const c = loadCfg();
    $('rsModeGraph').checked = c.mode !== 'url';
    $('rsModeUrl').checked = c.mode === 'url';
    $('rsClientId').value = c.clientId || '';
    $('rsTenantId').value = c.tenantId || '';
    $('rsLibrary').value = c.libraryName || '';
    $('rsFolder').value = c.folderPath || '';
    $('rsSrbName').value = c.srbFileName || '';
    $('rsRulesName').value = c.rulesFileName || '';
    $('rsSrbUrl').value = c.srbUrl || '';
    $('rsRulesUrl').value = c.rulesUrl || '';
    $('rsUrlBearer').value = c.urlBearer || '';
    $('rsPollMin').value = String(c.pollMinutes || 15);
    $('rsAutoEnabled').checked = !!c.autoEnabled;

    $('rsGraphPanel').style.display = c.mode === 'url' ? 'none' : '';
    $('rsUrlPanel').style.display = c.mode === 'url' ? '' : 'none';

    var st = '';
    if (c.lastError) st += 'Last error: ' + c.lastError + '\n';
    if (c.lastPollOk) st += 'Last successful poll: ' + c.lastPollOk + '\n';
    if (c.lastSrbSig) st += 'Last SRB sig: ' + c.lastSrbSig + '\n';
    if (c.lastRulesSig) st += 'Last Rules sig: ' + c.lastRulesSig + '\n';
    setStatus(st || 'Configure and Save, then Sign in (Graph) or Check now.');
  }

  function readForm() {
    const mode = $('rsModeUrl').checked ? 'url' : 'graph';
    return {
      mode: mode,
      clientId: $('rsClientId').value.trim(),
      tenantId: $('rsTenantId').value.trim(),
      libraryName: $('rsLibrary').value.trim(),
      folderPath: $('rsFolder').value.trim(),
      srbFileName: $('rsSrbName').value.trim(),
      rulesFileName: $('rsRulesName').value.trim(),
      srbUrl: $('rsSrbUrl').value.trim(),
      rulesUrl: $('rsRulesUrl').value.trim(),
      urlBearer: $('rsUrlBearer').value.trim(),
      pollMinutes: parseInt($('rsPollMin').value, 10) || 15,
      autoEnabled: $('rsAutoEnabled').checked,
      siteHostname: DEFAULTS.siteHostname,
      sitePath: DEFAULTS.sitePath
    };
  }

  async function signInGraph() {
    const cfg = readForm();
    saveCfg(cfg);
    if (!cfg.clientId || !cfg.tenantId) {
      setStatus('Enter Client ID and Tenant ID first.');
      return;
    }
    setStatus('Loading MSAL…');
    try {
      await ensureMsal();
      const msal = await getMsalInstance(cfg);
      try {
        await msal.loginPopup({ scopes: SCOPES });
      } catch (e) {
        if (e.errorCode !== 'user_cancelled') throw e;
        setStatus('Sign-in cancelled.');
        return;
      }
      const accounts = msal.getAllAccounts();
      if (accounts[0]) msal.setActiveAccount(accounts[0]);
      setStatus('Signed in. Click Check now to load files.');
    } catch (e) {
      setStatus('Sign-in error: ' + (e.message || e));
    }
  }

  function openModal() {
    fillForm();
    $('remoteSyncOverlay').classList.add('open');
  }

  function closeModal() {
    $('remoteSyncOverlay').classList.remove('open');
  }

  function init() {
    var btn = $('remoteSyncOpenBtn');
    if (btn) btn.addEventListener('click', openModal);
    $('remoteSyncClose').addEventListener('click', closeModal);
    $('remoteSyncOverlay').addEventListener('click', function (e) {
      if (e.target.id === 'remoteSyncOverlay') closeModal();
    });

    $('rsModeGraph').addEventListener('change', function () {
      $('rsGraphPanel').style.display = '';
      $('rsUrlPanel').style.display = 'none';
    });
    $('rsModeUrl').addEventListener('change', function () {
      $('rsGraphPanel').style.display = 'none';
      $('rsUrlPanel').style.display = '';
    });

    $('rsBtnSave').addEventListener('click', function () {
      saveCfg(readForm());
      fillForm();
      stopPolling();
      startPolling();
      setStatus('Settings saved.');
    });

    $('rsBtnSignIn').addEventListener('click', function () {
      signInGraph();
    });

    $('rsBtnCheckNow').addEventListener('click', function () {
      saveCfg(readForm());
      checkNow();
    });

    /* First-time defaults from repo path (no secrets). */
    var c = loadCfg();
    if (!c.folderPath) saveCfg({ folderPath: DEFAULTS.folderPath, libraryName: DEFAULTS.libraryName });

    stopPolling();
    startPolling();

    window.OverlapRemoteSync = {
      stop: function () {
        stopPolling();
        saveCfg({ autoEnabled: false });
        if ($('rsAutoEnabled')) $('rsAutoEnabled').checked = false;
      },
      openModal: openModal,
      checkNow: checkNow
    };
  }

  if (document.readyState === 'loading')
    document.addEventListener('DOMContentLoaded', init);
  else
    init();
})();
