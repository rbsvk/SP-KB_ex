document.getElementById('startBtn').addEventListener('click', async () => {
  const options = {
    exportPages: document.getElementById('optPages').checked,
    exportDocs: document.getElementById('optDocs').checked,
    exportLists: document.getElementById('optLists').checked,
    delayMs: Math.max(100, parseInt(document.getElementById('optDelay').value, 10) || 350)
  };

  await chrome.storage.local.set({ spOptions: options });

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !tab.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ['scraper.js']
  });

  window.close();
});
