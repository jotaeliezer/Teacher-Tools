// background.js — BrightBridge service worker
// Opens the side panel when the extension icon is clicked on any Brightspace page.

chrome.action.onClicked.addListener((tab) => {
  chrome.sidePanel.open({ windowId: tab.windowId });
});

// Allow the side panel to open on all Brightspace pages automatically
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});
