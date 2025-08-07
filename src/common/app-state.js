/*
 * Application State Manager
 * Centralized state management for Outlook2OneNote add-in
 */

// Application state
let appState = {
  selectedNotebook: null,
  isAuthenticated: false,
  userSettings: {},
  sessionData: {}
};

// Notebook state management
export function getSelectedNotebook() {
  return appState.selectedNotebook;
}

export function setSelectedNotebook(notebook) {
  appState.selectedNotebook = notebook;
  console.log('Notebook selected:', notebook?.displayName);
}

export function clearSelectedNotebook() {
  appState.selectedNotebook = null;
  console.log('Notebook selection cleared');
}

// Authentication state
export function setAuthenticationStatus(isAuthenticated) {
  appState.isAuthenticated = isAuthenticated;
}

export function getAuthenticationStatus() {
  return appState.isAuthenticated;
}

// User settings (with Office.js persistence)
export function setUserSetting(key, value) {
  appState.userSettings[key] = value;
  Office.context.roamingSettings.set(key, value);
  Office.context.roamingSettings.saveAsync();
}

export function getUserSetting(key) {
  if (appState.userSettings[key] !== undefined) {
    return appState.userSettings[key];
  }
  return Office.context.roamingSettings.get(key);
}

// Session data (temporary, not persisted)
export function setSessionData(key, value) {
  appState.sessionData[key] = value;
}

export function getSessionData(key) {
  return appState.sessionData[key];
}

// Initialize state from persisted settings
export function initializeAppState() {
  // Load user settings from Office.js roaming settings
  const settings = Office.context.roamingSettings.get('userSettings');
  if (settings) {
    appState.userSettings = settings;
  }
}

// Debug helper (development only)
export function getAppState() {
  if (__DEV__) {
    return { ...appState };
  }
  return null;
}