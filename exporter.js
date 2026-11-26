(function () {
  if (window.__spSingleFileExporterLoaded) return;
  window.__spSingleFileExporterLoaded = true;

  let isExporting = false;
  let abortExport = false;
  let pauseExport = false;

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function sanitizeName(name, maxLen = 70) {
    return (name || 'subor')
      .replace(/[\\/:*?"<>|#%]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^\.+|\.+$/g, '')
      .slice(0, maxLen) || 'polozka';
  }

  function sanitizePath(relPath) {
    return relPath
      .split('/')
      .filter(Boolean)
      .map((seg) => sanitizeName(seg, 65))
      .join('/');
  }

  // Bezpečné uloženie Blobu bez posielania veľkých dát cez IPC
  async function saveBlobSafely(blob, filename) {
    const blobUrl = URL.createObjectURL(blob);
    try {
      const res = await new Promise((resolve) => {
        chrome.runtime.sendMessage(
          { action: 'DOWNLOAD_WAIT', url: blobUrl, filename },
          resolve
        );
      });

      // Fallback ak firemná politika blokuje blob URL v chrome.downloads
      if (!res || !res.ok) {
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = filename.split('/').pop();
        document.body.appendChild(a);
        a.click();
        a.remove();
        await sleep(400);
      }
    } finally {
      setTimeout(() => URL.revokeObjectURL(blobUrl), 2000);
    }
  }

  // Priame stiahnutie existujúceho súboru (.pptx, .docx, .pdf) s čakaním na dokončenie
  function downloadFileUrlSafely(fileUrl, filename) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage(
        { action: 'DOWNLOAD_WAIT', url: fileUrl, filename },
        resolve
      );
    });
  }

  function showToast(text, duration = 3500) {
    let toast = document.getElementById('sp-sf-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'sp-sf-toast';
      toast.style.cssText = [
        'position:fixed', 'bottom:18px', 'right:18px', 'z-index:2147483647',
        'background:#1b1b22', 'color:#fff', 'border:1px solid #0078d4',
        'border-radius:8px', 'padding:10px 14px', 'font-family:Segoe UI,sans-serif',
        'font-size:12px', 'box-shadow:0 6px 20px rgba(0,0,0,0.4)'
      ].join(';');
      document.body.appendChild(toast);
    }
    toast.textContent = text;
    toast.style.display = 'block';
    if (duration > 0) {
      clearTimeout(toast.__t);
      toast.__t = setTimeout(() => {
        toast.style.display = 'none';
      }, duration);
    }
  }

  // =========================================================================
  // 1. REŽIM SINGLEFILE MV3 (Uloženie aktuálne otvorenej stránky)
  // =========================================================================
  async function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  async function captureCurrentPageAsSingleFile(inlineImages = true) {
    showToast('SingleFile: Spracovávam aktuálnu stránku...', 0);

    const clone = document.documentElement.cloneNode(true);
    clone.querySelectorAll('script, noscript, #sp-sf-toast, #sp-batch-panel').forEach((el) => el.remove());

    // Vloženie CSS štýlov do dokumentu
    const styleChunks = [];
    for (const sheet of Array.from(document.styleSheets)) {
      try {
        if (sheet.cssRules) {
          const rules = Array.from(sheet.cssRules).map((r) => r.cssText).join('\n');
          styleChunks.push(rules);
        }
      } catch (e) {
        if (sheet.href) {
          try {
            const resp = await fetch(sheet.href, { credentials: 'include' });
            if (resp.ok) styleChunks.push(await resp.text());
          } catch (_) {}
        }
      }
    }

    clone.querySelectorAll('link[rel="stylesheet"]').forEach((l) => l.remove());
    const head = clone.querySelector('head') || clone;
    if (styleChunks.length > 0) {
      const styleEl = document.createElement('style');
      styleEl.textContent = styleChunks.join('\n');
      head.appendChild(styleEl);
    }

    // Prevod relatívnych odkazov na absolútne
    clone.querySelectorAll('[href]').forEach((el) => {
      const val = el.getAttribute('href');
      if (val && !val.startsWith('#') && !val.startsWith('javascript:') && !val.startsWith('mailto:')) {
        try {
          el.setAttribute('href', new URL(val, window.location.href).href);
        } catch (_) {}
      }
    });

    // Spracovanie obrázkov (iba do 500 KB, aby nezaťažili pamäť)
    const imgs = Array.from(clone.querySelectorAll('img'));
    for (const img of imgs) {
      const src = img.getAttribute('src');
      if (!src || src.startsWith('data:')) continue;
      let absUrl = src;
      try {
        absUrl = new URL(src, window.location.href).href;
        img.setAttribute('src', absUrl);
      } catch (_) {
        continue;
      }

      if (inlineImages && absUrl.startsWith(window.location.origin)) {
        try {
          const resp = await fetch(absUrl, { credentials: 'include' });
          if (resp.ok) {
            const imgBlob = await resp.blob();
            if (imgBlob.size <= 512 * 1024) {
              const dataUri = await blobToDataUrl(imgBlob);
              img.setAttribute('src', dataUri);
            }
          }
        } catch (_) {}
      }
    }

    const finalHtml = '<!DOCTYPE html>\n' + clone.outerHTML;
    const htmlBlob = new Blob([finalHtml], { type: 'text/html;charset=utf-8' });
    const title = sanitizeName(document.title || 'stranka', 80);
    await saveBlobSafely(htmlBlob, `SP_Export/SingleFile/${title}.html`);
    showToast(`Hotovo! Stránka uložená ako ${title}.html`);
  }

  // =========================================================================
  // 2. SHAREPOINT BATCH EXPORTER (S výberom knižníc a 1-by-1 frontou)
  // =========================================================================
  async function spFetchJson(url) {
    const res = await fetch(url, {
      credentials: 'include',
      headers: { Accept: 'application/json;odata=verbose' }
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }

  async function detectSharePointWeb() {
    const urlObj = new URL(window.location.href);
    let parts = urlObj.pathname.split('/').filter(Boolean);
    const specialIdx = parts.findIndex((p) =>
      ['_layouts', 'pages', 'sitepages', 'forms', 'lists'].includes(p.toLowerCase())
    );
    if (specialIdx !== -1) parts = parts.slice(0, specialIdx);
    else if (parts.length > 0 && parts[parts.length - 1].includes('.')) parts.pop();

    while (parts.length >= 0) {
      const candidate = urlObj.origin + (parts.length ? '/' + parts.join('/') : '');
      try {
        const data = await spFetchJson(`${candidate}/_api/web?$select=Url,Title,ServerRelativeUrl`);
        if (data?.d?.Url) {
          return {
            webUrl: data.d.Url.replace(/\/$/, ''),
            serverRelUrl: data.d.ServerRelativeUrl.replace(/\/$/, ''),
            title: data.d.Title || 'SharePoint_KB'
          };
        }
      } catch (_) {}
      if (parts.length === 0) break;
      parts.pop();
    }
    return null;
  }

  function cleanAspxForCopilot(rawHtml, pageTitle, sourceUrl) {
    const doc = new DOMParser().parseFromString(rawHtml, 'text/html');
    const title = (doc.title || pageTitle || 'Stránka').trim();

    [
      'script', 'style', 'noscript', 'iframe',
      '#s4-ribbonrow', '#suiteBar', '#suiteBarDelta',
      '#sideNavBox', '#s4-titlerow', '#globalNavBox',
      '.ms-dialogHidden', '.ms-cui-topBar2', '#ms-designer-ribbon'
    ].forEach((sel) => doc.querySelectorAll(sel).forEach((el) => el.remove()));

    const mainNode =
      doc.querySelector('#DeltaPlaceHolderMain') ||
      doc.querySelector('#contentBox') ||
      doc.querySelector('.ms-rtestate-field') ||
      doc.body;

    // Prevod relatívnych obrázkov a odkazov na absolútne
    if (mainNode) {
      mainNode.querySelectorAll('[src]').forEach((el) => {
        try {
          el.setAttribute('src', new URL(el.getAttribute('src'), sourceUrl).href);
        } catch (_) {}
      });
      mainNode.querySelectorAll('[href]').forEach((el) => {
        try {
          el.setAttribute('href', new URL(el.getAttribute('href'), sourceUrl).href);
        } catch (_) {}
      });
    }

    return `<!DOCTYPE html>
<html lang="sk">
<head><meta charset="UTF-8"><title>${title.replace(/</g, '&lt;')}</title></head>
<body>
<h1>${title.replace(/</g, '&lt;')}</h1>
<p><small>Zdroj: ${sourceUrl}</small></p>
<hr>
${mainNode ? mainNode.innerHTML : ''}
</body>
</html>`;
  }

  function stripHtml(val) {
    if (val === null || val === undefined) return '';
    const str = String(val);
    if (!str.includes('<')) return str.trim();
    const div = document.createElement('div');
    div.innerHTML = str.replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n');
    return (div.textContent || div.innerText || '').replace(/\n{3,}/g, '\n\n').trim();
  }

  function buildCsv(items) {
    const skip = new Set([
      '__metadata', 'FirstUniqueAncestorSecurableObject', 'RoleAssignments',
      'AttachmentFiles', 'ContentType', 'GetDlpPolicyTip', 'FieldValuesAsHtml',
      'FieldValuesAsText', 'FieldValuesForEdit', 'File', 'Folder',
      'ParentList', 'Properties', 'Versions', 'AuthorId', 'EditorId',
      'OData__UIVersionString', 'GUID', 'ComplianceAssetId'
    ]);
    const keys = new Set();
    items.forEach((it) => {
      Object.keys(it).forEach((k) => {
        if (!skip.has(k) && typeof it[k] !== 'object') keys.add(k);
      });
    });
    const headers = Array.from(keys);
    const esc = (v) => `"${stripHtml(v).replace(/"/g, '""')}"`;
    const rows = [headers.map(esc).join(',')];
    items.forEach((it) => rows.push(headers.map((h) => esc(it[h])).join(',')));
    return '\uFEFF' + rows.join('\r\n');
  }

  async function fetchPagedItems(url) {
    let all = [];
    let next = url;
    while (next && !abortExport) {
      const json = await spFetchJson(next);
      if (json?.d?.results) all = all.concat(json.d.results);
      next = json?.d?.__next || null;
    }
    return all;
  }

  async function openBatchPanel() {
    let panel = document.getElementById('sp-batch-panel');
    if (panel) {
      panel.style.display = 'block';
      return;
    }

    panel = document.createElement('div');
    panel.id = 'sp-batch-panel';
    panel.style.cssText = [
      'position:fixed', 'top:16px', 'right:16px', 'z-index:2147483647',
      'width:370px', 'max-height:88vh', 'overflow-y:auto',
      'background:#18181f', 'color:#f3f3f7', 'border:1px solid #0078d4',
      'border-radius:10px', 'padding:14px', 'font-family:Segoe UI,sans-serif',
      'font-size:12px', 'box-shadow:0 10px 30px rgba(0,0,0,0.55)'
    ].join(';');

    panel.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
        <strong style="color:#4cc2ff;font-size:14px;">SharePoint KB Exporter (1-by-1)</strong>
        <button id="sp-close-panel" style="background:none;border:none;color:#aaa;font-size:16px;cursor:pointer;">✕</button>
      </div>
      <div id="sp-panel-body">Načítavam štruktúru SharePointu...</div>
    `;
    document.body.appendChild(panel);
    document.getElementById('sp-close-panel').onclick = () => {
      panel.style.display = 'none';
    };

    const bodyEl = document.getElementById('sp-panel-body');
    const webInfo = await detectSharePointWeb();
    if (!webInfo) {
      bodyEl.innerHTML = '<div style="color:#ff7b72;">Nepodarilo sa rozpoznať SharePoint lokalitu na tejto URL.</div>';
      return;
    }

    const listsData = await spFetchJson(
      `${webInfo.webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate,ItemCount`
    );
    const rawLists = listsData?.d?.results || [];
    const docTemplates = new Set([101, 109, 119, 850]);
    const listTemplates = new Set([100, 103, 104, 105, 106, 107]);
    const ignoreTitles = new Set(['form templates', 'style library', 'site assets', 'wfsvc']);

    const validLists = rawLists.filter((l) => {
      if (!l.ItemCount || l.ItemCount <= 0) return false;
      if (ignoreTitles.has((l.Title || '').toLowerCase())) return false;
      return docTemplates.has(l.BaseTemplate) || listTemplates.has(l.BaseTemplate);
    });

    const listCheckboxesHtml = validLists
      .map((l, idx) => {
        const isPageLib = l.BaseTemplate === 119 || l.BaseTemplate === 850 || (l.Title || '').toLowerCase() === 'pages';
        const isDoc = docTemplates.has(l.BaseTemplate);
        const tag = isPageLib ? 'Stránky' : isDoc ? 'Dokumenty' : 'Zoznam/FAQ';
        const checked = isPageLib || !isDoc ? 'checked' : '';
        return `
          <label style="display:flex;align-items:center;justify-content:space-between;padding:5px 6px;background:#22222b;border-radius:5px;margin-bottom:5px;cursor:pointer;">
            <span style="display:flex;align-items:center;gap:7px;">
              <input type="checkbox" class="sp-lib-chk" data-idx="${idx}" ${checked}>
              <span>${l.Title}</span>
            </span>
            <span style="color:#9a9aa8;font-size:11px;">${tag} (${l.ItemCount})</span>
          </label>
        `;
      })
      .join('');

    bodyEl.innerHTML = `
      <div style="margin-bottom:8px;color:#b8b8c6;">Lokalita: <b>${webInfo.title}</b></div>
      <div style="display:flex;gap:8px;margin-bottom:8px;">
        <button id="sp-sel-all" style="background:#2b2b36;color:#ddd;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-size:11px;">Vybrať všetko</button>
        <button id="sp-sel-none" style="background:#2b2b36;color:#ddd;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-size:11px;">Zrušiť výber</button>
      </div>
      <div style="max-height:260px;overflow-y:auto;margin-bottom:10px;padding-right:2px;">
        ${listCheckboxesHtml}
      </div>
      <div style="margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;">
        <span>Bezpečná pauza po dokončení súboru:</span>
        <span><input id="sp-delay-input" type="number" value="800" min="200" max="10000" style="width:65px;background:#2b2b36;color:#fff;border:1px solid #444;border-radius:4px;padding:3px 6px;"> ms</span>
      </div>
      <div style="display:flex;gap:6px;margin-bottom:8px;">
        <button id="sp-run-btn" style="flex:1;background:#0078d4;color:#fff;border:none;border-radius:5px;padding:8px;font-weight:600;cursor:pointer;">Spustiť export vybraných</button>
        <button id="sp-pause-btn" style="display:none;background:#d97706;color:#fff;border:none;border-radius:5px;padding:8px 10px;cursor:pointer;">Pauza</button>
        <button id="sp-stop-btn" style="display:none;background:#c42b1c;color:#fff;border:none;border-radius:5px;padding:8px 10px;cursor:pointer;">Stop</button>
      </div>
      <div id="sp-batch-status" style="background:#121217;padding:8px;border-radius:5px;color:#d0d0dc;min-height:36px;line-height:1.4;">
        Pripravené. Odporúčanie: začni najprv so Stránkami a Zoznamami, potom veľké knižnice dokumentov.
      </div>
    `;

    document.getElementById('sp-sel-all').onclick = () => {
      document.querySelectorAll('.sp-lib-chk').forEach((c) => (c.checked = true));
    };
    document.getElementById('sp-sel-none').onclick = () => {
      document.querySelectorAll('.sp-lib-chk').forEach((c) => (c.checked = false));
    };

    const runBtn = document.getElementById('sp-run-btn');
    const pauseBtn = document.getElementById('sp-pause-btn');
    const stopBtn = document.getElementById('sp-stop-btn');
    const statusBox = document.getElementById('sp-batch-status');

    pauseBtn.onclick = () => {
      pauseExport = !pauseExport;
      pauseBtn.textContent = pauseExport ? 'Pokračovať' : 'Pauza';
      if (pauseExport) statusBox.textContent = 'Pozastavené používateľom...';
    };

    stopBtn.onclick = () => {
      abortExport = true;
      pauseExport = false;
      statusBox.textContent = 'Zastavujem po dokončení aktuálneho súboru...';
    };

    runBtn.onclick = async () => {
      if (isExporting) return;
      const selectedIndices = Array.from(document.querySelectorAll('.sp-lib-chk:checked')).map((c) =>
        parseInt(c.getAttribute('data-idx'), 10)
      );
      if (selectedIndices.length === 0) {
        statusBox.textContent = 'Vyber aspoň jednu knižnicu alebo zoznam.';
        return;
      }

      isExporting = true;
      abortExport = false;
      pauseExport = false;
      runBtn.style.display = 'none';
      pauseBtn.style.display = 'inline-block';
      stopBtn.style.display = 'inline-block';

      const delayMs = Math.max(200, parseInt(document.getElementById('sp-delay-input').value, 10) || 800);
      const storageData = await chrome.storage.local.get(['spDownloadedKeys']);
      const downloadedSet = new Set(storageData.spDownloadedKeys || []);

      const baseFolder = `SP_Export/${sanitizeName(webInfo.title, 50)}`;
      let savedCount = 0;
      let skippedCount = 0;

      const saveHistory = async (key) => {
        downloadedSet.add(key);
        await chrome.storage.local.set({ spDownloadedKeys: Array.from(downloadedSet) });
      };

      try {
        for (const idx of selectedIndices) {
          if (abortExport) break;
          const list = validLists[idx];
          statusBox.textContent = `Načítavam zoznam položiek: ${list.Title}...`;

          if (docTemplates.has(list.BaseTemplate)) {
            const itemsUrl = `${webInfo.webUrl}/_api/web/lists(guid'${list.Id}')/items?$select=Id,Title,FileRef,FileLeafRef,FSObjType&$top=1000`;
            const items = await fetchPagedItems(itemsUrl);

            for (const item of items) {
              while (pauseExport && !abortExport) await sleep(300);
              if (abortExport) break;
              if (item.FSObjType !== 0 || !item.FileRef || !item.FileLeafRef) continue;

              const itemKey = `file:${item.FileRef}`;
              if (downloadedSet.has(itemKey)) {
                skippedCount++;
                statusBox.textContent = `Uložené: ${savedCount} | Preskočené: ${skippedCount} (${item.FileLeafRef})`;
                continue;
              }

              let relPath = item.FileRef;
              if (webInfo.serverRelUrl && relPath.toLowerCase().startsWith(webInfo.serverRelUrl.toLowerCase())) {
                relPath = relPath.slice(webInfo.serverRelUrl.length);
              }
              const targetRel = sanitizePath(relPath);
              const fullUrl = window.location.origin + item.FileRef;
              const isAspx = item.FileLeafRef.toLowerCase().endsWith('.aspx');

              statusBox.textContent = `Sťahujem (1/1): ${item.FileLeafRef} | Uložené: ${savedCount}`;

              if (isAspx) {
                try {
                  const res = await fetch(fullUrl, { credentials: 'include' });
                  if (res.ok) {
                    const rawHtml = await res.text();
                    const cleanHtml = cleanAspxForCopilot(rawHtml, item.Title || item.FileLeafRef, fullUrl);
                    const blob = new Blob([cleanHtml], { type: 'text/html;charset=utf-8' });
                    await saveBlobSafely(blob, `${baseFolder}/${targetRel}.html`);
                    await saveHistory(itemKey);
                    savedCount++;
                  }
                } catch (_) {}
              } else {
                const dlRes = await downloadFileUrlSafely(fullUrl, `${baseFolder}/${targetRel}`);
                if (dlRes && dlRes.ok) {
                  await saveHistory(itemKey);
                  savedCount++;
                }
              }

              await sleep(delayMs);
            }
          } else if (listTemplates.has(list.BaseTemplate)) {
            while (pauseExport && !abortExport) await sleep(300);
            if (abortExport) break;

            const listKey = `list:${list.Id}`;
            if (downloadedSet.has(listKey)) {
              skippedCount++;
              continue;
            }

            statusBox.textContent = `Exportujem zoznam: ${list.Title}...`;
            const itemsUrl = `${webInfo.webUrl}/_api/web/lists(guid'${list.Id}')/items?$top=1000`;
            const items = await fetchPagedItems(itemsUrl);
            if (items.length > 0) {
              const csvText = buildCsv(items);
              const csvBlob = new Blob([csvText], { type: 'text/csv;charset=utf-8' });
              const safeTitle = sanitizeName(list.Title, 60);
              await saveBlobSafely(csvBlob, `${baseFolder}/Zoznamy/${safeTitle}.csv`);
              await saveHistory(listKey);
              savedCount++;
              await sleep(delayMs);
            }
          }
        }

        statusBox.textContent = abortExport
          ? `Zastavené. Uložené: ${savedCount}, preskočené: ${skippedCount}.`
          : `Dokončené! Nové súbory: ${savedCount}, už stiahnuté skôr: ${skippedCount}.`;
      } catch (err) {
        statusBox.textContent = `Chyba: ${err.message}`;
      } finally {
        isExporting = false;
        runBtn.style.display = 'inline-block';
        pauseBtn.style.display = 'none';
        stopBtn.style.display = 'none';
      }
    };
  }

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.command === 'SAVE_SINGLE_PAGE') {
      captureCurrentPageAsSingleFile(msg.inlineImages !== false);
    } else if (msg.command === 'OPEN_BATCH_PANEL') {
      openBatchPanel();
    }
  });
})();
