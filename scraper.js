(async function () {
  if (window.__spExporterRunning) {
    alert('Export SharePointu už beží. Sleduj stav v pravom dolnom rohu stránky.');
    return;
  }
  window.__spExporterRunning = true;
  window.__spExporterAbort = false;

  const { spOptions } = await chrome.storage.local.get(['spOptions']);
  const opts = spOptions || { exportPages: true, exportDocs: true, exportLists: true, delayMs: 350 };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // Vytvorenie plávajúceho panelu priamo na stránke
  const hud = document.createElement('div');
  hud.id = 'sp-exporter-hud';
  hud.style.cssText = [
    'position:fixed', 'bottom:16px', 'right:16px', 'z-index:999999',
    'width:340px', 'background:#18181f', 'color:#f4f4f6',
    'border:1px solid #0078d4', 'border-radius:8px', 'padding:12px 14px',
    'font-family:Segoe UI,sans-serif', 'font-size:12px',
    'box-shadow:0 8px 24px rgba(0,0,0,0.45)'
  ].join(';');
  hud.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <strong style="color:#4cc2ff;font-size:13px;">SharePoint KB Exporter</strong>
      <button id="sp-stop-btn" style="background:#c42b1c;color:#fff;border:none;border-radius:4px;padding:3px 8px;cursor:pointer;font-size:11px;">Zastaviť</button>
    </div>
    <div id="sp-status-text" style="margin-bottom:6px;color:#e0e0e6;word-break:break-word;">Zisťujem štruktúru lokality...</div>
    <div id="sp-counters" style="color:#9a9aa8;font-size:11px;">Stránky: 0 | Dokumenty: 0 | Zoznamy: 0</div>
  `;
  document.body.appendChild(hud);

  document.getElementById('sp-stop-btn').onclick = () => {
    window.__spExporterAbort = true;
    setStatus('Zastavujem export...');
  };

  let countPages = 0;
  let countDocs = 0;
  let countLists = 0;

  function setStatus(text) {
    const el = document.getElementById('sp-status-text');
    if (el) el.textContent = text;
  }

  function updateCounters() {
    const el = document.getElementById('sp-counters');
    if (el) el.textContent = `Stránky: ${countPages} | Dokumenty: ${countDocs} | Zoznamy: ${countLists}`;
  }

  function sanitizeSegment(name) {
    return (name || 'polozka')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^\.+|\.+$/g, '')
      .slice(0, 65) || 'subor';
  }

  function sanitizeRelPath(relPath) {
    return relPath
      .split('/')
      .filter(Boolean)
      .map(sanitizeSegment)
      .join('/');
  }

  function textToDataUrl(text, mimeType) {
    const bytes = new TextEncoder().encode(text);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return `data:${mimeType};base64,${btoa(binary)}`;
  }

  function downloadUrl(url, filename) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({ action: 'DOWNLOAD_URL', url, filename }, resolve);
    });
  }

  function downloadData(dataUrl, filename) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({ action: 'DOWNLOAD_DATA', url: dataUrl, filename }, resolve);
    });
  }

  async function spFetchJson(url) {
    const res = await fetch(url, {
      credentials: 'include',
      headers: { Accept: 'application/json;odata=verbose' }
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }

  // 1. Automatické nájdenie koreňovej URL aktuálnej SharePoint lokality
  async function detectWebInfo() {
    const urlObj = new URL(window.location.href);
    let pathParts = urlObj.pathname.split('/').filter(Boolean);
    const specialIdx = pathParts.findIndex((p) =>
      ['_layouts', 'pages', 'sitepages', 'forms', 'lists'].includes(p.toLowerCase())
    );
    if (specialIdx !== -1) {
      pathParts = pathParts.slice(0, specialIdx);
    } else if (pathParts.length > 0 && pathParts[pathParts.length - 1].includes('.')) {
      pathParts.pop();
    }

    while (pathParts.length >= 0) {
      const candidate = urlObj.origin + (pathParts.length ? '/' + pathParts.join('/') : '');
      try {
        const data = await spFetchJson(`${candidate}/_api/web?$select=Url,Title,ServerRelativeUrl`);
        if (data && data.d && data.d.Url) {
          return {
            webUrl: data.d.Url.replace(/\/$/, ''),
            serverRelUrl: data.d.ServerRelativeUrl.replace(/\/$/, ''),
            title: data.d.Title || 'SharePoint_KB'
          };
        }
      } catch (e) {
        // skús o úroveň vyššie
      }
      if (pathParts.length === 0) break;
      pathParts.pop();
    }
    return null;
  }

  // Očistenie .aspx HTML stránky pre Copilot
  function cleanSharePointHtml(rawHtml, pageTitle, sourceUrl) {
    const doc = new DOMParser().parseFromString(rawHtml, 'text/html');
    const title = (doc.title || pageTitle || 'Stránka').trim();

    const selectorsToRemove = [
      'script', 'style', 'noscript', 'iframe',
      '#s4-ribbonrow', '#suiteBar', '#suiteBarDelta',
      '#sideNavBox', '#s4-titlerow', '#globalNavBox',
      '.ms-dialogHidden', '.ms-cui-topBar2', '#ms-designer-ribbon'
    ];
    selectorsToRemove.forEach((sel) => {
      doc.querySelectorAll(sel).forEach((el) => el.remove());
    });

    const mainNode =
      doc.querySelector('#DeltaPlaceHolderMain') ||
      doc.querySelector('#contentBox') ||
      doc.querySelector('.ms-rtestate-field') ||
      doc.body;

    const contentHtml = mainNode ? mainNode.innerHTML : '';

    return `<!DOCTYPE html>
<html lang="sk">
<head>
  <meta charset="UTF-8">
  <title>${title.replace(/</g, '&lt;')}</title>
</head>
<body>
  <h1>${title.replace(/</g, '&lt;')}</h1>
  <p><small>Zdroj: ${sourceUrl}</small></p>
  <hr>
  ${contentHtml}
</body>
</html>`;
  }

  function stripHtmlTags(val) {
    if (val === null || val === undefined) return '';
    const str = String(val);
    if (!str.includes('<')) return str.trim();
    const tmp = document.createElement('div');
    tmp.innerHTML = str
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n');
    return (tmp.textContent || tmp.innerText || '').replace(/\n{3,}/g, '\n\n').trim();
  }

  function buildCsv(items) {
    const skipKeys = new Set([
      '__metadata', 'FirstUniqueAncestorSecurableObject', 'RoleAssignments',
      'AttachmentFiles', 'ContentType', 'GetDlpPolicyTip', 'FieldValuesAsHtml',
      'FieldValuesAsText', 'FieldValuesForEdit', 'File', 'Folder',
      'ParentList', 'Properties', 'Versions', 'AuthorId', 'EditorId',
      'OData__UIVersionString', 'GUID', 'ComplianceAssetId'
    ]);

    const keysSet = new Set();
    items.forEach((it) => {
      Object.keys(it).forEach((k) => {
        if (!skipKeys.has(k) && typeof it[k] !== 'object') {
          keysSet.add(k);
        }
      });
    });

    const headers = Array.from(keysSet);
    const escapeCell = (v) => `"${stripHtmlTags(v).replace(/"/g, '""')}"`;

    const lines = [headers.map(escapeCell).join(',')];
    items.forEach((it) => {
      lines.push(headers.map((h) => escapeCell(it[h])).join(','));
    });

    return '\uFEFF' + lines.join('\r\n');
  }

  async function fetchAllPagedItems(initialUrl) {
    let all = [];
    let nextUrl = initialUrl;
    while (nextUrl && !window.__spExporterAbort) {
      const json = await spFetchJson(nextUrl);
      if (json && json.d && Array.isArray(json.d.results)) {
        all = all.concat(json.d.results);
      }
      nextUrl = json && json.d && json.d.__next ? json.d.__next : null;
    }
    return all;
  }

  try {
    const webInfo = await detectWebInfo();
    if (!webInfo) {
      setStatus('Chyba: Nepodarilo sa rozpoznať SharePoint REST API na tejto stránke.');
      window.__spExporterRunning = false;
      return;
    }

    const baseFolder = `SP_Export/${sanitizeSegment(webInfo.title)}`;
    setStatus(`Načítavam zoznam knižníc pre: ${webInfo.title}...`);

    const listsData = await spFetchJson(
      `${webInfo.webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    const lists = (listsData.d && listsData.d.results) || [];

    const docTemplates = new Set([101, 109, 119, 850]);
    const listTemplates = new Set([100, 103, 104, 105, 106, 107]);

    for (const list of lists) {
      if (window.__spExporterAbort) break;
      if (!list.ItemCount || list.ItemCount === 0) continue;

      const listTitleLower = (list.Title || '').toLowerCase();
      if (['form templates', 'style library', 'site assets', 'wfsvc'].includes(listTitleLower)) {
        continue;
      }

      // A) Knižnice dokumentov a webových stránok
      if (docTemplates.has(list.BaseTemplate)) {
        if (!opts.exportDocs && !opts.exportPages) continue;

        setStatus(`Čítam položky: ${list.Title} (${list.ItemCount})...`);
        const itemsUrl = `${webInfo.webUrl}/_api/web/lists(guid'${list.Id}')/items?$select=Id,Title,FileRef,FileLeafRef,FSObjType&$top=1000`;
        let items = [];
        try {
          items = await fetchAllPagedItems(itemsUrl);
        } catch (e) {
          continue;
        }

        for (const item of items) {
          if (window.__spExporterAbort) break;
          if (item.FSObjType !== 0 || !item.FileRef || !item.FileLeafRef) continue;

          const isAspx = item.FileLeafRef.toLowerCase().endsWith('.aspx');
          if (isAspx && !opts.exportPages) continue;
          if (!isAspx && !opts.exportDocs) continue;

          let relInsideSite = item.FileRef;
          if (webInfo.serverRelUrl && relInsideSite.toLowerCase().startsWith(webInfo.serverRelUrl.toLowerCase())) {
            relInsideSite = relInsideSite.slice(webInfo.serverRelUrl.length);
          }
          const cleanPath = sanitizeRelPath(relInsideSite);
          const fullFileUrl = window.location.origin + item.FileRef;

          setStatus(`Sťahujem: ${item.FileLeafRef}`);

          if (isAspx) {
            try {
              const res = await fetch(fullFileUrl, { credentials: 'include' });
              const rawHtml = await res.text();
              const cleanHtml = cleanSharePointHtml(rawHtml, item.Title || item.FileLeafRef, fullFileUrl);
              const dataUrl = textToDataUrl(cleanHtml, 'text/html;charset=utf-8');
              await downloadData(dataUrl, `${baseFolder}/${cleanPath}.html`);
              countPages++;
              updateCounters();
            } catch (e) {
              // preskoč nedostupnú stránku
            }
          } else {
            await downloadUrl(fullFileUrl, `${baseFolder}/${cleanPath}`);
            countDocs++;
            updateCounters();
          }

          await sleep(opts.delayMs);
        }
      }

      // B) SharePoint zoznamy (FAQ, KBContent, Predpisy...)
      if (listTemplates.has(list.BaseTemplate) && opts.exportLists) {
        setStatus(`Exportujem zoznam: ${list.Title} (${list.ItemCount})...`);
        try {
          const itemsUrl = `${webInfo.webUrl}/_api/web/lists(guid'${list.Id}')/items?$top=1000`;
          const items = await fetchAllPagedItems(itemsUrl);
          if (items.length > 0) {
            const csvText = buildCsv(items);
            const dataUrl = textToDataUrl(csvText, 'text/csv;charset=utf-8');
            const safeName = sanitizeSegment(list.Title);
            await downloadData(dataUrl, `${baseFolder}/Zoznamy/${safeName}.csv`);
            countLists++;
            updateCounters();
            await sleep(opts.delayMs);
          }
        } catch (e) {
          // preskoč zoznam bez prístupu
        }
      }
    }

    // Stiahnutie samotnej domovskej stránky (default.aspx), ak nie je v knižnici Pages
    if (opts.exportPages && !window.__spExporterAbort) {
      try {
        const homeUrl = `${webInfo.webUrl}/default.aspx`;
        const res = await fetch(homeUrl, { credentials: 'include' });
        if (res.ok) {
          const rawHtml = await res.text();
          const cleanHtml = cleanSharePointHtml(rawHtml, `${webInfo.title} - Domovská stránka`, homeUrl);
          const dataUrl = textToDataUrl(cleanHtml, 'text/html;charset=utf-8');
          await downloadData(dataUrl, `${baseFolder}/default.aspx.html`);
          countPages++;
          updateCounters();
        }
      } catch (e) {}
    }

    setStatus(window.__spExporterAbort ? 'Export bol zastavený.' : 'Hotovo! Všetko je v Stiahnuté súbory/SP_Export.');
  } catch (err) {
    setStatus(`Chyba: ${err.message}`);
  } finally {
    window.__spExporterRunning = false;
  }
})();
