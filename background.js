chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'DOWNLOAD_URL' || msg.action === 'DOWNLOAD_DATA') {
    chrome.downloads.download(
      {
        url: msg.url,
        filename: msg.filename,
        conflictAction: 'uniquify',
        saveAs: false
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        } else {
          sendResponse({ ok: true, id: downloadId });
        }
      }
    );
    return true;
  }
});
