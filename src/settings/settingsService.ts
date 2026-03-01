import type { SettingsSchema } from '@/types/settings'
import { debugLog, debugError } from '@/utils/logger'

export interface SettingsService {
  get(): SettingsSchema
  set(partial: Partial<SettingsSchema>): void
  save(): Promise<void>
}

export const DEFAULT_SETTINGS: SettingsSchema = {
  notebookId: null,
  notebookDisplayName: null,
  sectionNameFormat: '{subject} ({date})',
  sortOrder: 'asc',
  includeAttachments: true,
  preferredLink: 'both',
}

const SETTINGS_KEY = 'o2on_settings'

// Call after Office.onReady() — roamingSettings is not available before that.
export function createSettingsService(): SettingsService {
  const stored = Office.context.roamingSettings.get(SETTINGS_KEY) as Partial<SettingsSchema> | null
  let current: SettingsSchema = { ...DEFAULT_SETTINGS, ...(stored ?? {}) }

  debugLog('settings', 'Loaded settings', current)

  return {
    get(): SettingsSchema {
      return { ...current }
    },

    set(partial: Partial<SettingsSchema>): void {
      current = { ...current, ...partial }
      debugLog('settings', 'Settings updated', current)
    },

    save(): Promise<void> {
      Office.context.roamingSettings.set(SETTINGS_KEY, current)
      return new Promise((resolve, reject) => {
        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            debugLog('settings', 'Settings saved')
            resolve()
          } else {
            debugError('settings', 'Settings save failed', result.error)
            reject(new Error(result.error?.message ?? 'saveAsync failed'))
          }
        })
      })
    },
  }
}
