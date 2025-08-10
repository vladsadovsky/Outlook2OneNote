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

// Notebook state management (with persistence)
export function getSelectedNotebook() {
  // First check in-memory state
  if (appState.selectedNotebook) {
    return appState.selectedNotebook;
  }
  
  // If not in memory, try to load from persistent storage
  const persistedNotebook = Office.context.roamingSettings.get('selectedNotebook');
  if (persistedNotebook) {
    appState.selectedNotebook = persistedNotebook;
    console.log('📖 Loaded notebook from persistent storage:', persistedNotebook?.displayName);
    return persistedNotebook;
  }
  
  return null;
}

export function setSelectedNotebook(notebook) {
  appState.selectedNotebook = notebook;
  
  // Persist to Office.js roaming settings
  Office.context.roamingSettings.set('selectedNotebook', notebook);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('💾 Notebook selection saved persistently:', notebook?.displayName);
    } else {
      console.warn('⚠️ Failed to save notebook selection:', result.error);
    }
  });
  
  console.log('📝 Notebook selected:', notebook?.displayName);
}

export function clearSelectedNotebook() {
  appState.selectedNotebook = null;
  
  // Clear from persistent storage
  Office.context.roamingSettings.remove('selectedNotebook');
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('🗑️ Notebook selection cleared from persistent storage');
    } else {
      console.warn('⚠️ Failed to clear notebook selection:', result.error);
    }
  });
  
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
  console.log('🔄 Initializing app state...');
  
  // Load user settings from Office.js roaming settings
  const settings = Office.context.roamingSettings.get('userSettings');
  if (settings) {
    appState.userSettings = settings;
    console.log('⚙️ Loaded user settings from storage');
  }
  
  // Load selected notebook from persistent storage
  const persistedNotebook = Office.context.roamingSettings.get('selectedNotebook');
  if (persistedNotebook) {
    appState.selectedNotebook = persistedNotebook;
    console.log('📖 Restored previously selected notebook:', persistedNotebook?.displayName);
  }
  
  console.log('✅ App state initialized');
}

// Debug helper (development only)
export function getAppState() {
  if (__DEV__) {
    return { ...appState };
  }
  return null;
}