export interface SettingsSchema {
  notebookId: string | null
  notebookDisplayName: string | null
  sectionNameFormat: string        // default: '{subject} ({date})'
  sortOrder: 'asc' | 'desc'       // default: 'asc'
  includeAttachments: boolean      // default: true
  preferredLink: 'web' | 'desktop' | 'both'  // default: 'both'
  showDebugPanel: boolean          // default: false
}
